# -*- coding: utf-8 -*-
"""
사진대지 매크로 (한글 / HWP · HWPX)
================================================================
한글 문서의 표에 이미지를 순서대로 자동 삽입하는 프로그램.
표의 열 수와 캡션 행 유무를 스스로 감지한다.

필요 환경 : Windows + 한글(HWP) 설치
설치      : pip install -r requirements.txt
실행      : python hwp_photo_macro.py
"""
from __future__ import annotations

import ctypes
import functools
import os
import re
import shutil
import sys
import threading
import time
import tkinter as tk
import traceback
from ctypes import wintypes
from datetime import datetime
from tkinter import filedialog, messagebox, ttk, scrolledtext

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

# 한글 내부 단위: 1mm = 283.465 HWPUNIT
HWPUNIT_PER_MM = 283.465
# 화면 기준 96dpi: 1inch = 7200 HWPUNIT = 96px 이므로 1px = 75 HWPUNIT
HWPUNIT_PER_PX = 75
IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp")
MAX_CELL_WALK = 400


def app_dir() -> str:
    """exe 로 묶였을 때도 올바른 폴더를 돌려준다."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


LOG_FILE = os.path.join(app_dir(), "error_log.txt")


def natural_key(path: str):
    name = os.path.basename(path)
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def split_threshold(values):
    """1차원 값들을 두 무리로 갈랐을 때의 경계값. 무리가 하나뿐이면 None."""
    if len(values) < 2:
        return None
    lo, hi = min(values), max(values)
    if hi < lo * 1.6:
        return None
    c1, c2 = float(lo), float(hi)
    for _ in range(30):
        g1 = [v for v in values if abs(v - c1) <= abs(v - c2)]
        g2 = [v for v in values if abs(v - c1) > abs(v - c2)]
        if not g1 or not g2:
            break
        n1, n2 = sum(g1) / len(g1), sum(g2) / len(g2)
        if abs(n1 - c1) < 1 and abs(n2 - c2) < 1:
            break
        c1, c2 = n1, n2
    return (c1 + c2) / 2


# ==================================================================
#  한글 제어부
# ==================================================================
class HwpError(Exception):
    """사용자에게 그대로 보여줄 수 있는, 예상된 오류."""


class HwpController:
    def __init__(self):
        self.hwp = None

    @property
    def connected(self) -> bool:
        return self.hwp is not None

    def connect(self):
        """이미 실행 중인 한글에 먼저 붙어보고, 없으면 새로 띄운다."""
        try:
            import win32com.client as win32
        except ImportError:
            raise HwpError("pywin32 가 설치되어 있지 않습니다.\n  pip install pywin32")

        hwp, attached = None, False

        # 1) 이미 열려 있는 한글에 붙기
        try:
            hwp = win32.GetActiveObject("HWPFrame.HwpObject")
            attached = True
        except Exception:
            hwp = None

        # 2) 실패하면 새 인스턴스
        if hwp is None:
            try:
                hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            except Exception:
                try:
                    hwp = win32.Dispatch("HWPFrame.HwpObject")
                except Exception as e:
                    raise HwpError(
                        "한글에 연결하지 못했습니다.\n"
                        "한글이 설치되어 있는지 확인해 주세요.\n\n"
                        f"({e})")

        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathChecker")
        except Exception:
            pass
        try:
            hwp.XHwpWindows.Item(0).Visible = True
        except Exception:
            pass

        self.hwp = hwp
        return attached

    def _require(self):
        if self.hwp is None:
            raise HwpError("한글에 연결되어 있지 않습니다. [한글 연결]을 먼저 눌러주세요.")

    def open_document(self, path: str):
        """지정한 문서를 열어 작업 대상으로 삼는다."""
        self._require()
        last = None
        for args in ((path,), (path, "", "forceopen:true"), (path, "", "")):
            try:
                self.hwp.Open(*args)
                return
            except Exception as e:
                last = e
        raise HwpError(f"문서를 열지 못했습니다.\n{path}\n\n({last})")

    def doc_path(self) -> str:
        self._require()
        try:
            return self.hwp.Path or ""
        except Exception:
            return ""

    def doc_name(self) -> str:
        p = self.doc_path()
        return os.path.basename(p) if p else "(저장되지 않은 빈 문서)"

    # ---------------- 셀 정보 ----------------
    def cell_size(self):
        """현재 칸의 (안쪽 너비, 안쪽 높이) HWPUNIT. 표 밖이면 None."""
        self._require()
        try:
            cs = self.hwp.CellShape
        except Exception:
            return None
        if cs is None:
            return None

        def item(key, default=0):
            try:
                v = cs.Item(key)
                return default if v is None else int(v)
            except Exception:
                return default

        w, h = item("Width"), item("Height")
        if w <= 0 or h <= 0:
            return None
        w -= item("MarginLeft") + item("MarginRight")
        h -= item("MarginTop") + item("MarginBottom")
        return max(w, 1), max(h, 1)

    def in_table(self) -> bool:
        return self.cell_size() is not None

    def run(self, action: str) -> bool:
        self._require()
        try:
            return bool(self.hwp.Run(action))
        except Exception:
            return False

    # ---------------- 커서 상태 ----------------
    def save(self) -> bool:
        self._require()
        try:
            self.hwp.Save()
            return True
        except Exception:
            return False

    def make_backup(self) -> str:
        """현재 문서를 저장한 뒤 같은 폴더에 백업 사본을 만든다.

        한글이 COM 으로 이뤄진 편집을 되돌리기 목록에 쌓지 않는 경우가 있어,
        Undo 대신 이 사본이 실질적인 복구 수단이 된다."""
        path = self.doc_path()
        if not path:
            raise HwpError("문서가 아직 파일로 저장되지 않았습니다.\n"
                           "한글에서 먼저 저장한 뒤 다시 시도해 주세요.")
        self.save()
        stem, ext = os.path.splitext(path)
        backup = f"{stem}_backup_{datetime.now():%Y%m%d_%H%M%S}{ext}"
        shutil.copy2(path, backup)
        return backup

    def undo(self, times: int = 1) -> int:
        """한글의 되돌리기를 지정한 횟수만큼 실행한다."""
        self._require()
        done = 0
        for _ in range(max(times, 0)):
            before = self.get_pos()
            self.run("Undo")
            done += 1
            if before is not None and self.get_pos() == before and done > 1:
                pass  # 위치가 안 바뀌어도 되돌려진 경우가 있어 멈추지는 않는다
        return done

    def get_pos(self):
        """현재 캐럿 위치 (list, para, pos). 실패하면 None."""
        self._require()
        try:
            return self.hwp.GetPos()
        except Exception:
            return None

    def set_pos(self, pos) -> bool:
        if not pos:
            return False
        try:
            return bool(self.hwp.SetPos(*pos))
        except Exception:
            return False

    def escape_selection(self):
        """개체가 선택된 상태를 풀고 글자 편집 상태로 되돌린다.

        사진을 넣은 직후에는 그 그림이 선택된 상태라
        TableRightCell 같은 칸 이동 명령이 먹지 않는다."""
        try:
            self.hwp.Run("Cancel")
        except Exception:
            pass

    def move_cell(self, action: str) -> bool:
        """칸 이동. 실제로 커서가 움직였는지로 성공을 판정한다.

        hwp.Run() 은 동작이 성공해도 False 를 돌려주는 경우가 있어
        반환값을 믿으면 첫 칸에서 바로 멈춰버린다.
        GetPos() 는 칸마다 다른 list 번호를 주므로 이동 여부를 확실히 알 수 있다."""
        before = self.get_pos()
        self.run(action)
        after = self.get_pos()

        if before is not None and after is not None:
            if after != before:
                return True
            # 한 번 실패하면 개체 선택을 풀고 다시 시도
            self.escape_selection()
            self.run(action)
            return self.get_pos() != before

        # GetPos 를 못 쓰는 환경이면 칸 안에 있는지로 대신 판단
        return self.cell_size() is not None

    # ---------------- 표 훑기 ----------------
    def scan_cells(self):
        """현재 칸부터 표 끝까지 각 칸의 (너비, 높이)를 모으고 원래 칸으로 복귀."""
        self._require()
        sizes, steps = [], 0
        while len(sizes) < MAX_CELL_WALK:
            s = self.cell_size()
            if s is None:
                break
            sizes.append(s)
            if not self.move_cell("TableRightCell"):
                break
            steps += 1
        for _ in range(steps):
            self.move_cell("TableLeftCell")
        return sizes

    # ---------------- 삽입 ----------------
    def insert_text(self, text: str):
        self._require()
        if not text:
            return
        act = self.hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("Text", text)
        act.Execute(pset)

    @staticmethod
    def _image_ratio(path: str):
        if Image is None:
            return 4, 3
        try:
            with Image.open(path) as im:
                return im.size
        except Exception:
            return 4, 3

    def insert_picture_fit(self, path: str, margins_px=(0, 0, 0, 0), border_px=0.0,
                           post_adjust: bool = False):
        """현재 칸 크기에 맞춰 비율을 유지한 채 삽입. margins_px = (상, 하, 좌, 우) 픽셀

        post_adjust 를 켜면 삽입 후 ShapeObjDialog 로 크기와 테두리를 다시 지정한다.
        이 단계는 지정하지 않은 속성까지 기본값으로 함께 적용되므로,
        그림이 선택되지 않거나 배치가 이상해지면 꺼야 한다."""
        self._require()
        if not os.path.exists(path):
            raise HwpError(f"사진 파일을 찾을 수 없습니다.\n{path}")
        size = self.cell_size()
        if size is None:
            raise HwpError("커서가 표 안에 있지 않습니다.\n"
                           "한글 문서에서 사진을 넣을 칸을 클릭한 뒤 다시 시도해 주세요.")

        cw, ch = size
        top, bottom, left, right = (int(m * HWPUNIT_PER_PX) for m in margins_px)
        avail_w = max(cw - left - right, 1)
        avail_h = max(ch - top - bottom, 1)

        iw, ih = self._image_ratio(path)
        scale = min(avail_w / iw, avail_h / ih)
        w = max(int(iw * scale), 1)
        h = max(int(ih * scale), 1)

        hwp = self.hwp
        before = self.get_pos()          # 삽입 전 캐럿 위치를 기억해 둔다
        hwp.Run("ParagraphShapeAlignCenter")

        # sizeoption 3 = width/height 로 지정한 크기.
        # 2 는 "셀 크기에 맞추어" 라서 우리가 계산한 값을 무시하고 여백이 먹지 않는다.
        inserted = False
        for opt in (3, 2):
            try:
                hwp.InsertPicture(path, True, opt, False, False, 0, w, h)
                inserted = True
                break
            except Exception:
                continue
        if not inserted:
            hwp.InsertPicture(path, True)
        if post_adjust or border_px > 0:
            self._apply_shape(w, h, border_px)
        # [진단 1단계] Run("Cancel") 이 되돌리기 기록을 끊는지 확인하기 위해
        # 삽입 경로에서는 escape_selection() 을 호출하지 않는다.
        self.set_pos(before)
        return (cw, ch), (w, h)

    def _apply_shape(self, w: int, h: int, border_px: float):
        hwp = self.hwp
        try:
            hwp.FindCtrl()
            hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)
            so = hwp.HParameterSet.HShapeObject
            so.Width = w
            so.Height = h
            so.TreatAsChar = 1
            # 개체 보호가 켜진 채로 적용되면 그림을 클릭해도 선택되지 않는다
            for key in ("Lock", "Protect"):
                try:
                    setattr(so, key, 0)
                except Exception:
                    pass
            if border_px > 0:
                try:
                    so.LineShape.Type = 1
                    so.LineShape.Color = 0
                    so.LineShape.Width = max(int(border_px * HWPUNIT_PER_PX), 1)
                except Exception:
                    pass
            hwp.HAction.Execute("ShapeObjDialog", so.HSet)
        except Exception:
            pass
        finally:
            hwp.Run("Cancel")


# ==================================================================
#  전역 단축키
# ==================================================================
WM_HOTKEY = 0x0312
MOD_CONTROL, MOD_SHIFT = 0x0002, 0x0004


class GlobalHotkeys(threading.Thread):
    def __init__(self, bindings):
        super().__init__(daemon=True)
        self.bindings = bindings
        self._stop = threading.Event()

    def run(self):
        try:
            user32 = ctypes.windll.user32
        except Exception:
            return
        registered = []
        for hk_id, (mods, vk, _) in self.bindings.items():
            try:
                if user32.RegisterHotKey(None, hk_id, mods, vk):
                    registered.append(hk_id)
            except Exception:
                pass
        msg = wintypes.MSG()
        while not self._stop.is_set():
            try:
                if user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):
                    if msg.message == WM_HOTKEY:
                        b = self.bindings.get(msg.wParam)
                        if b:
                            b[2]()
            except Exception:
                pass
            time.sleep(0.03)
        for hk_id in registered:
            try:
                user32.UnregisterHotKey(None, hk_id)
            except Exception:
                pass

    def stop(self):
        self._stop.set()


# ==================================================================
#  오류를 눈에 보이게 만드는 장치
# ==================================================================
def guarded(func):
    """버튼 동작에서 터진 예외를 로그와 창으로 드러낸다.
    이게 없으면 windowed exe 에서는 아무 일도 안 일어난 것처럼 보인다."""
    @functools.wraps(func)
    def wrapper(self, *args, **kwargs):
        try:
            return func(self, *args, **kwargs)
        except HwpError as e:
            self.log(f"[안내] {e}")
            messagebox.showwarning("안내", str(e), parent=self)
        except Exception as e:
            tb = traceback.format_exc()
            self.log(f"[오류] {func.__name__}\n{tb}")
            messagebox.showerror(
                "오류",
                f"{type(e).__name__}: {e}\n\n"
                f"자세한 내용을 아래 파일에 기록했습니다.\n{LOG_FILE}",
                parent=self)
    return wrapper


# ==================================================================
#  GUI
# ==================================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("사진대지 매크로 (한글)")
        self.geometry("940x720")
        self.minsize(880, 640)

        self.ctrl = HwpController()
        self.photos = []
        self.cursor = 0
        self._thumb = None
        self._thr_cache = None   # 사진 칸 기준 높이. 매번 표를 훑지 않도록 재사용한다.

        self._build_ui()
        self._bind_keys()
        self._start_hotkeys()
        self._refresh()
        self.log("프로그램을 시작했습니다.")
        if Image is None:
            self.log("[경고] Pillow 를 불러오지 못해 미리보기가 꺼졌습니다.")

    # ---------------- 로그 ----------------
    def log(self, message: str):
        stamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {message}\n"
        try:
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", line)
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")
        except Exception:
            pass
        if message.startswith("[오류]"):
            try:
                with open(LOG_FILE, "a", encoding="utf-8") as f:
                    f.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} =====\n{message}\n")
            except Exception:
                pass

    # ---------------- 레이아웃 ----------------
    def _build_ui(self):
        root = ttk.Frame(self, padding=8)
        root.pack(fill="both", expand=True)

        bar = ttk.Frame(root)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(bar, text="한글 연결", command=self.on_connect, width=11).pack(side="left")
        ttk.Button(bar, text="문서 열기", command=self.on_open_doc, width=11).pack(side="left", padx=4)
        ttk.Button(bar, text="문서 다시 확인", command=self.on_refresh_doc, width=13).pack(side="left")
        self.lbl_status = ttk.Label(bar, text="연결 안 됨", foreground="#b00")
        self.lbl_status.pack(side="left", padx=10)

        body = ttk.Frame(root)
        body.pack(fill="both", expand=True)

        # ---- 왼쪽 ----
        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True)

        self.canvas = tk.Canvas(left, bg="#1a1a1a", height=230, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        list_area = ttk.Frame(left)
        list_area.pack(fill="both", expand=True, pady=(8, 0))

        self.tree = ttk.Treeview(list_area, columns=("no", "file", "caption"),
                                 show="headings", height=8)
        self.tree.heading("no", text="#")
        self.tree.heading("file", text="파일명")
        self.tree.heading("caption", text="캡션 (더블클릭 수정)")
        self.tree.column("no", width=40, anchor="center", stretch=False)
        self.tree.column("file", width=210)
        self.tree.column("caption", width=210)
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", self.on_edit_caption)

        sb = ttk.Scrollbar(list_area, orient="vertical", command=self.tree.yview)
        sb.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

        order = ttk.Frame(list_area)
        order.pack(side="left", fill="y", padx=4)
        for label, where in (("▲▲", "top"), ("▲", "up"), ("▼", "down"), ("▼▼", "bottom")):
            ttk.Button(order, text=label, width=4,
                       command=lambda w=where: self.move_item(w)).pack(pady=1)

        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=6)
        ttk.Button(btns, text="사진 추가", command=self.on_add).pack(side="left")
        ttk.Button(btns, text="폴더 추가", command=self.on_add_folder).pack(side="left", padx=4)
        ttk.Button(btns, text="이름순 정렬", command=self.on_sort_by_name).pack(side="left")
        ttk.Button(btns, text="선택 삭제", command=self.on_remove).pack(side="left", padx=4)
        ttk.Button(btns, text="모두 삭제", command=self.on_clear).pack(side="left")

        ttk.Label(left, text="진행 기록").pack(anchor="w", pady=(6, 0))
        self.txt_log = scrolledtext.ScrolledText(left, height=7, state="disabled",
                                                 font=("Consolas", 9), wrap="word")
        self.txt_log.pack(fill="x")

        # ---- 오른쪽 ----
        right = ttk.Frame(body, width=270)
        right.pack(side="left", fill="y", padx=(10, 0))
        right.pack_propagate(False)

        form = ttk.LabelFrame(right, text="표 구조", padding=8)
        form.pack(fill="x")
        self.var_mode = tk.StringVar(value="auto")
        ttk.Radiobutton(form, text="자동 감지 (권장)", value="auto",
                        variable=self.var_mode, command=self._toggle_mode)\
            .grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(form, text="큰 칸=사진, 작은 칸=캡션",
                  foreground="#666").grid(row=1, column=0, columnspan=2, sticky="w", padx=18)
        ttk.Label(form, text="사진 칸 최소 높이(mm)").grid(row=2, column=0, sticky="w",
                                                    padx=18, pady=(4, 0))
        self.var_min_h = tk.DoubleVar(value=0.0)
        self.var_min_h.trace_add("write", lambda *_: self._invalidate_plan())
        self.sp_min_h = ttk.Spinbox(form, from_=0, to=200, increment=5, width=6,
                                    textvariable=self.var_min_h)
        self.sp_min_h.grid(row=2, column=1, sticky="e", pady=(4, 0))
        ttk.Label(form, text="0 = 자동 계산", foreground="#666")\
            .grid(row=3, column=0, columnspan=2, sticky="w", padx=18)
        ttk.Radiobutton(form, text="모든 칸에 사진 넣기", value="all",
                        variable=self.var_mode, command=self._toggle_mode)\
            .grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        self.var_caption_on = tk.BooleanVar(value=True)
        ttk.Checkbutton(form, text="캡션 자동 입력", variable=self.var_caption_on)\
            .grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.var_auto_caption = tk.BooleanVar(value=False)
        ttk.Checkbutton(form, text="캡션 비면 파일명 사용", variable=self.var_auto_caption)\
            .grid(row=6, column=0, columnspan=2, sticky="w")

        opt = ttk.LabelFrame(right, text="여백 주기 (px)", padding=8)
        opt.pack(fill="x", pady=8)
        self.var_margin = {}
        for i, key in enumerate(("상", "하", "좌", "우")):
            ttk.Label(opt, text=key).grid(row=i, column=0, sticky="w", pady=1)
            v = tk.IntVar(value=0)
            self.var_margin[key] = v
            ttk.Spinbox(opt, from_=0, to=300, increment=1, width=8, textvariable=v)\
                .grid(row=i, column=1, sticky="e")

        bd = ttk.LabelFrame(right, text="테두리", padding=8)
        bd.pack(fill="x")
        self.var_border_on = tk.BooleanVar(value=False)
        ttk.Checkbutton(bd, text="사진 테두리 (px)", variable=self.var_border_on)\
            .grid(row=0, column=0, sticky="w")
        self.var_border_px = tk.IntVar(value=1)
        ttk.Spinbox(bd, from_=1, to=20, increment=1, width=6,
                    textvariable=self.var_border_px).grid(row=0, column=1, sticky="e")
        self.var_post_adjust = tk.BooleanVar(value=False)
        ttk.Checkbutton(bd, text="삽입 후 크기 다시 지정",
                        variable=self.var_post_adjust)\
            .grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Label(bd, text="그림이 선택되지 않으면 꺼두세요",
                  foreground="#666").grid(row=2, column=0, columnspan=2, sticky="w")

        act = ttk.Frame(right)
        act.pack(fill="x", pady=12)
        ttk.Button(act, text="한 장 넣기  (Ctrl+Q)", command=self.on_insert_one).pack(fill="x", pady=2)
        ttk.Button(act, text="한번에 넣기  (Ctrl+Shift+A)", command=self.on_insert_all).pack(fill="x", pady=2)
        ttk.Button(act, text="표 구조 확인", command=self.on_check_table).pack(fill="x", pady=2)
        ttk.Button(act, text="칸 이동 진단", command=self.on_diagnose_move).pack(fill="x", pady=2)
        ttk.Button(act, text="넣을 위치 처음으로", command=self.on_reset_cursor).pack(fill="x", pady=2)

        undo = ttk.LabelFrame(right, text="복구", padding=8)
        undo.pack(fill="x", pady=(8, 0))
        self.var_backup = tk.BooleanVar(value=True)
        ttk.Checkbutton(undo, text="넣기 전에 백업본 만들기",
                        variable=self.var_backup).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Button(undo, text="지금 백업본 만들기", command=self.on_backup)\
            .grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Label(undo, text="되돌리기 횟수").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.var_undo_n = tk.IntVar(value=10)
        ttk.Spinbox(undo, from_=1, to=200, increment=1, width=6,
                    textvariable=self.var_undo_n).grid(row=2, column=1, sticky="e", pady=(8, 0))
        ttk.Button(undo, text="되돌리기 시도", command=self.on_undo)\
            .grid(row=3, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        ttk.Label(undo, text="한글이 받아주지 않으면 동작하지 않습니다",
                  foreground="#666", wraplength=230).grid(row=4, column=0, columnspan=2, sticky="w")

        self.lbl_progress = ttk.Label(right, text="", foreground="#555", wraplength=250)
        self.lbl_progress.pack(fill="x")
        self._toggle_mode()

    def _toggle_mode(self):
        self.sp_min_h.configure(state="normal" if self.var_mode.get() == "auto" else "disabled")
        self._invalidate_plan()

    def _invalidate_plan(self):
        """표 기준값을 버린다. 다음 삽입 때 다시 계산한다."""
        self._thr_cache = None

    def _threshold(self) -> int:
        """사진 칸 기준 높이. 한 번 계산해두고 재사용한다.

        한 장 넣을 때마다 표 전체를 훑으면 칸이 많은 표에서 눈에 띄게 느리다."""
        if self._thr_cache is None:
            thr, _ = self._plan()
            self._thr_cache = thr
        return self._thr_cache

    def _bind_keys(self):
        self.bind("<Control-q>", lambda e: self.on_insert_one())
        self.bind("<Control-Shift-A>", lambda e: self.on_insert_all())
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _start_hotkeys(self):
        self.hotkeys = None
        if sys.platform != "win32":
            return
        try:
            self.hotkeys = GlobalHotkeys({
                1: (MOD_CONTROL, ord("Q"), lambda: self.after(0, self.on_insert_one)),
                2: (MOD_CONTROL | MOD_SHIFT, ord("A"), lambda: self.after(0, self.on_insert_all)),
            })
            self.hotkeys.start()
        except Exception:
            self.hotkeys = None

    # ---------------- 목록 ----------------
    def _refresh(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for i, p in enumerate(self.photos):
            mark = "✔" if i < self.cursor else str(i + 1)
            self.tree.insert("", "end", iid=str(i),
                             values=(mark, os.path.basename(p["path"]), p["caption"]))
        total = len(self.photos)
        self.lbl_progress.config(
            text=f"{self.cursor} / {total} 장 삽입됨" if total else "사진을 추가해 주세요.")

    def _selected_index(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    @guarded
    def on_add(self):
        chosen = filedialog.askopenfilenames(
            parent=self,
            title="사진 선택",
            filetypes=[("이미지 파일", ("*.jpg", "*.jpeg", "*.png", "*.bmp",
                                    "*.gif", "*.tif", "*.tiff", "*.webp")),
                       ("모든 파일", "*.*")])
        # Windows 에서 문자열 하나로 돌아오는 경우가 있어 splitlist 로 풀어준다
        paths = list(self.tk.splitlist(chosen)) if chosen else []
        self.log(f"사진 선택: {len(paths)}개")
        self._add_paths(paths)

    @guarded
    def on_add_folder(self):
        folder = filedialog.askdirectory(parent=self, title="사진 폴더 선택")
        if not folder:
            self.log("폴더 선택을 취소했습니다.")
            return
        names = [f for f in os.listdir(folder) if f.lower().endswith(IMAGE_EXTS)]
        paths = sorted((os.path.join(folder, f) for f in names), key=natural_key)
        self.log(f"폴더 선택: {folder} — 이미지 {len(paths)}개")
        if not paths:
            messagebox.showinfo("안내", "폴더 안에 이미지 파일이 없습니다.", parent=self)
            return
        self._add_paths(paths)

    def _add_paths(self, paths):
        existing = {p["path"] for p in self.photos}
        added = 0
        for p in paths:
            if p not in existing:
                self.photos.append({"path": p, "caption": ""})
                existing.add(p)
                added += 1
        self._refresh()
        self.log(f"목록에 {added}장 추가 (전체 {len(self.photos)}장)")
        if added and self.tree.get_children():
            self.tree.selection_set(str(len(self.photos) - added))

    @guarded
    def on_sort_by_name(self):
        if not self.photos:
            return
        self.photos.sort(key=lambda p: natural_key(p["path"]))
        self.cursor = 0
        self._refresh()
        self.log("이름순으로 정렬했습니다.")

    @guarded
    def on_remove(self):
        idx = self._selected_index()
        if idx is None:
            messagebox.showinfo("안내", "목록에서 삭제할 사진을 먼저 선택해 주세요.", parent=self)
            return
        self.photos.pop(idx)
        self.cursor = min(self.cursor, len(self.photos))
        self._refresh()

    @guarded
    def on_clear(self):
        if self.photos and messagebox.askyesno("확인", "목록을 모두 지울까요?", parent=self):
            self.photos.clear()
            self.cursor = 0
            self.canvas.delete("all")
            self._refresh()
            self.log("목록을 비웠습니다.")

    @guarded
    def move_item(self, where):
        i = self._selected_index()
        if i is None:
            return
        item = self.photos.pop(i)
        j = {"top": 0, "up": max(i - 1, 0),
             "down": min(i + 1, len(self.photos)), "bottom": len(self.photos)}[where]
        self.photos.insert(j, item)
        self._refresh()
        self.tree.selection_set(str(j))

    def on_select(self, _=None):
        i = self._selected_index()
        if i is not None:
            self._show_preview(self.photos[i]["path"])

    def _show_preview(self, path):
        self.canvas.delete("all")
        if Image is None:
            self.canvas.create_text(10, 10, anchor="nw", fill="#aaa",
                                    text="Pillow 미설치 — 미리보기 불가")
            return
        cw = self.canvas.winfo_width() or 400
        ch = self.canvas.winfo_height() or 230
        try:
            with Image.open(path) as im:
                im = im.copy()
                im.thumbnail((max(cw - 10, 50), max(ch - 10, 50)))
                self._thumb = ImageTk.PhotoImage(im)
            self.canvas.create_image(cw // 2, ch // 2, image=self._thumb)
        except Exception as e:
            self.canvas.create_text(10, 10, anchor="nw", fill="#f66", text=f"미리보기 실패: {e}")

    @guarded
    def on_edit_caption(self, event):
        row = self.tree.identify_row(event.y)
        if not row or self.tree.identify_column(event.x) != "#3":
            return
        i = int(row)
        box = self.tree.bbox(row, "caption")
        if not box:
            return
        x, y, w, h = box
        entry = ttk.Entry(self.tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, self.photos[i]["caption"])
        entry.focus_set()

        def commit(_=None):
            self.photos[i]["caption"] = entry.get()
            entry.destroy()
            self._refresh()
            self.tree.selection_set(row)

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", lambda e: entry.destroy())

    # ---------------- 한글 연동 ----------------
    def _update_status(self):
        if not self.ctrl.connected:
            self.lbl_status.config(text="연결 안 됨", foreground="#b00")
            return
        name = self.ctrl.doc_name()
        saved = bool(self.ctrl.doc_path())
        self.lbl_status.config(text=f"작업 대상: {name}",
                               foreground="#070" if saved else "#c60")

    @guarded
    def on_connect(self):
        attached = self.ctrl.connect()
        self._update_status()
        if attached:
            self.log(f"실행 중인 한글에 연결했습니다. 현재 문서: {self.ctrl.doc_name()}")
        else:
            self.log("한글을 새로 실행했습니다.")
        if not self.ctrl.doc_path():
            self.log("작업할 문서가 지정되지 않았습니다. [문서 열기]로 양식 파일을 골라주세요.")
            messagebox.showinfo(
                "문서를 골라주세요",
                "작업할 문서가 아직 지정되지 않았습니다.\n"
                "[문서 열기] 를 눌러 사진대지 양식 파일을 선택해 주세요.",
                parent=self)

    @guarded
    def on_open_doc(self):
        if not self.ctrl.connected:
            self.ctrl.connect()
            self.log("한글에 연결했습니다.")
        path = filedialog.askopenfilename(
            parent=self,
            title="작업할 한글 문서 선택",
            filetypes=[("한글 문서", ("*.hwp", "*.hwpx")), ("모든 파일", "*.*")])
        if not path:
            self.log("문서 선택을 취소했습니다.")
            return
        self.ctrl.open_document(path)
        self._invalidate_plan()
        self._update_status()
        self.log(f"문서를 열었습니다: {os.path.basename(path)}")

    @guarded
    def on_refresh_doc(self):
        """한글에서 다른 문서로 바꿔 작업할 때 현재 문서를 다시 읽는다."""
        if not self.ctrl.connected:
            messagebox.showinfo("안내", "먼저 [한글 연결]을 눌러주세요.", parent=self)
            return
        self._update_status()
        self.log(f"현재 문서: {self.ctrl.doc_name()}")

    @staticmethod
    def _mm(size):
        """(너비, 높이) HWPUNIT 을 읽기 쉬운 mm 문자열로."""
        w, h = size
        return f"{w / HWPUNIT_PER_MM:.0f}x{h / HWPUNIT_PER_MM:.0f}mm"

    def _insert_options(self):
        """삽입 옵션(여백/테두리)을 픽셀 단위로 돌려준다.

        주의: 이 메서드 이름을 _options 로 두면 안 된다.
        tkinter.Misc._options 를 가려버려서 파일 대화상자와 메시지 창이
        전부 TypeError 로 죽는다."""
        m = self.var_margin
        margins = (m["상"].get(), m["하"].get(), m["좌"].get(), m["우"].get())
        border = self.var_border_px.get() if self.var_border_on.get() else 0
        return margins, border

    def _plan(self):
        """표를 훑어 (사진 칸 기준 높이, 남은 사진 칸 개수)를 계산한다."""
        sizes = self.ctrl.scan_cells()
        if not sizes:
            raise HwpError("커서가 표 안에 있지 않습니다.\n"
                           "한글 문서에서 사진을 넣을 칸을 클릭한 뒤 다시 시도해 주세요.")
        if self.var_mode.get() == "all":
            thr = 0
        elif self.var_min_h.get() > 0:
            thr = int(self.var_min_h.get() * HWPUNIT_PER_MM)
        else:
            t = split_threshold([h for _, h in sizes])
            thr = int(t) if t else 0
        available = sum(1 for _, h in sizes if h >= thr)
        self.log(f"표 확인: 남은 칸 {len(sizes)}개, 사진 칸 {available}개, "
                 f"기준 {thr / HWPUNIT_PER_MM:.0f}mm")
        return thr, available

    def _caption_for(self, item):
        if not self.var_caption_on.get():
            return ""
        if item["caption"]:
            return item["caption"]
        if self.var_auto_caption.get():
            return os.path.splitext(os.path.basename(item["path"]))[0]
        return ""

    @guarded
    def on_check_table(self):
        if not self.ctrl.connected:
            messagebox.showinfo("안내", "먼저 [한글 연결]을 눌러주세요.", parent=self)
            return
        sizes = self.ctrl.scan_cells()
        thr, available = self._plan()
        self._thr_cache = thr
        heights = sorted({round(h / HWPUNIT_PER_MM) for _, h in sizes})
        remaining = len(self.photos) - self.cursor
        verdict = ""
        if remaining:
            if available >= remaining:
                verdict = f"\n\n넣을 사진 {remaining}장 → 칸이 충분합니다."
            else:
                verdict = (f"\n\n넣을 사진 {remaining}장 → 칸이 {remaining - available}개 부족합니다.")
        messagebox.showinfo(
            "표 구조 확인",
            f"커서 위치부터 남은 칸 : {len(sizes)}개\n"
            f"사진 칸으로 판단된 칸 : {available}개\n"
            f"기준 높이 : {thr / HWPUNIT_PER_MM:.0f}mm\n"
            f"칸 높이 종류 : {heights}mm"
            f"{verdict}\n\n"
            "결과가 맞지 않으면 [사진 칸 최소 높이]를 직접 지정해 보세요.",
            parent=self)

    @guarded
    def on_diagnose_move(self):
        """칸 이동이 왜 안 되는지 원인을 좁히기 위한 진단."""
        if not self.ctrl.connected:
            messagebox.showinfo("안내", "먼저 [한글 연결]을 눌러주세요.", parent=self)
            return
        lines = []
        start = self.ctrl.get_pos()
        lines.append(f"GetPos 사용 가능 : {'예' if start else '아니오'}")
        lines.append(f"시작 위치        : {start}")
        lines.append(f"표 안 여부       : {'예' if self.ctrl.in_table() else '아니오'}")

        raw = self.ctrl.run("TableRightCell")
        after = self.ctrl.get_pos()
        lines.append(f"Run 반환값       : {raw}")
        lines.append(f"이동 후 위치     : {after}")
        lines.append(f"실제로 움직였나  : {'예' if after != start else '아니오'}")

        if start:
            self.ctrl.set_pos(start)
        text = "\n".join(lines)
        self.log("칸 이동 진단\n" + text)
        messagebox.showinfo("칸 이동 진단", text, parent=self)

    @guarded
    def on_backup(self):
        if not self.ctrl.connected:
            messagebox.showinfo("안내", "먼저 [한글 연결]을 눌러주세요.", parent=self)
            return
        backup = self.ctrl.make_backup()
        self.log(f"백업본 생성: {os.path.basename(backup)}")
        messagebox.showinfo("백업 완료",
                            f"같은 폴더에 사본을 만들었습니다.\n\n{os.path.basename(backup)}",
                            parent=self)

    @guarded
    def on_undo(self):
        if not self.ctrl.connected:
            messagebox.showinfo("안내", "먼저 [한글 연결]을 눌러주세요.", parent=self)
            return
        n = self.var_undo_n.get()
        before = self.ctrl.get_pos()
        done = self.ctrl.undo(n)
        after = self.ctrl.get_pos()
        moved = before != after
        self.log(f"되돌리기 {done}회 시도 — 문서 상태 변화: {'있음' if moved else '없음'}")
        if not moved:
            messagebox.showinfo(
                "되돌리기가 동작하지 않습니다",
                "한글이 이 편집을 되돌리기 목록에 쌓지 않는 것으로 보입니다.\n\n"
                "[지금 백업본 만들기] 로 사본을 미리 만들어두고 작업하시거나,\n"
                "잘못된 경우 저장하지 않고 문서를 닫았다가 다시 여세요.",
                parent=self)

    @guarded
    def on_reset_cursor(self):
        self.cursor = 0
        self._invalidate_plan()
        self._refresh()
        self.log("넣을 위치를 목록 처음으로 되돌렸습니다.")

    def _check_saved(self) -> bool:
        """저장되지 않은 문서에서는 한글이 되돌리기를 받아주지 않는다."""
        if self.ctrl.doc_path():
            return True
        self.log("[경고] 문서가 파일로 저장되지 않았습니다.")
        return messagebox.askyesno(
            "저장되지 않은 문서입니다",
            "이 문서는 아직 파일로 저장되지 않았습니다.\n\n"
            "이 상태에서는 한글의 되돌리기(Ctrl+Z)가 동작하지 않고,\n"
            "백업본도 만들 수 없습니다.\n\n"
            "한글에서 먼저 저장하신 뒤 진행하시기를 권합니다.\n"
            "그래도 이대로 진행할까요?", parent=self)

    def _ready(self) -> bool:
        if not self.ctrl.connected:
            messagebox.showinfo("안내", "먼저 [한글 연결]을 눌러주세요.", parent=self)
            return False
        if not self._check_saved():
            return False
        if not self.photos:
            messagebox.showinfo("안내", "사진 목록이 비어 있습니다.", parent=self)
            return False
        if self.cursor >= len(self.photos):
            messagebox.showinfo("안내", "목록의 사진을 모두 넣었습니다.", parent=self)
            return False
        return True

    def _advance(self, thr: int, pending_caption: str) -> bool:
        """다음 사진 칸까지 이동하며 도중의 작은 칸에 캡션을 쓴다. 표 끝이면 False."""
        for step in range(MAX_CELL_WALK):
            if not self.ctrl.move_cell("TableRightCell"):
                self.log(f"이동 중단: 오른쪽 칸으로 갈 수 없습니다 ({step}칸 이동 후)")
                return False
            size = self.ctrl.cell_size()
            if size is None:
                self.log(f"이동 중단: 표 밖으로 나갔습니다 ({step + 1}칸 이동 후)")
                return False
            if size[1] >= thr:
                return True
            if pending_caption:
                self.ctrl.insert_text(pending_caption)
                pending_caption = ""
        return False

    @guarded
    def on_insert_one(self):
        if not self._ready():
            return
        margins, border = self._insert_options()
        thr = self._threshold()
        item = self.photos[self.cursor]
        cell, pic = self.ctrl.insert_picture_fit(
            item["path"], margins, border, post_adjust=self.var_post_adjust.get())
        self.cursor += 1
        self.log(f"삽입: {os.path.basename(item['path'])} — "
                 f"칸 {self._mm(cell)}, 사진 {self._mm(pic)}")
        self._advance(thr, self._caption_for(item))
        self._refresh()

    @guarded
    def on_insert_all(self):
        if not self._ready():
            return
        margins, border = self._insert_options()
        thr, available = self._plan()
        self._thr_cache = thr

        if self.var_backup.get():
            try:
                backup = self.ctrl.make_backup()
                self.log(f"백업본 생성: {os.path.basename(backup)}")
            except HwpError as e:
                if not messagebox.askyesno(
                        "백업을 만들지 못했습니다",
                        f"{e}\n\n백업 없이 그대로 진행할까요?", parent=self):
                    return
                self.log("백업 없이 진행합니다.")

        remaining = len(self.photos) - self.cursor
        if available < remaining:
            go = messagebox.askyesno(
                "칸이 부족합니다",
                f"이 표에 남은 사진 칸 : {available}개\n"
                f"넣어야 할 사진 : {remaining}장\n"
                f"→ {remaining - available}장이 들어가지 못합니다.\n\n"
                "칸이 찰 때까지만 채우고 멈춥니다.\n"
                "남은 사진은 목록에 남으니, 페이지를 추가한 뒤\n"
                "다음 표의 첫 칸을 클릭하고 다시 누르면 이어집니다.\n\n"
                "이대로 진행할까요?", parent=self)
            if not go:
                self.log("사용자가 삽입을 취소했습니다.")
                return

        inserted = 0
        while self.cursor < len(self.photos):
            item = self.photos[self.cursor]
            cell, pic = self.ctrl.insert_picture_fit(
                item["path"], margins, border, post_adjust=self.var_post_adjust.get())
            self.cursor += 1
            inserted += 1
            self.log(f"삽입 {inserted}: {os.path.basename(item['path'])} — "
                     f"칸 {self._mm(cell)}, 사진 {self._mm(pic)}")
            is_last = self.cursor >= len(self.photos)
            moved = self._advance(thr, self._caption_for(item))
            if is_last:
                break
            if not moved:
                messagebox.showinfo(
                    "표 끝에 도달",
                    f"{inserted}장을 넣고 표가 끝났습니다.\n"
                    f"다음 표의 첫 사진 칸을 클릭하고 다시 [한번에 넣기]를 누르세요.\n"
                    f"(남은 사진 {len(self.photos) - self.cursor}장)", parent=self)
                break
        self.log(f"완료: 이번에 {inserted}장 삽입")
        self._refresh()

    def on_close(self):
        if self.hotkeys:
            self.hotkeys.stop()
        self.destroy()


def main():
    app = App()

    def excepthook(exc_type, exc, tb):
        text = "".join(traceback.format_exception(exc_type, exc, tb))
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} =====\n{text}\n")
        except Exception:
            pass
        try:
            messagebox.showerror("예기치 못한 오류",
                                 f"{exc}\n\n기록: {LOG_FILE}", parent=app)
        except Exception:
            pass

    sys.excepthook = excepthook
    app.report_callback_exception = lambda *a: excepthook(*a)
    app.mainloop()


if __name__ == "__main__":
    main()