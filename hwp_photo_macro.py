# -*- coding: utf-8 -*-
"""
사진대지 매크로 (한글 / HWP · HWPX)
================================================================
한글에서 선택한 표 칸에 사진을 순서대로 넣는 프로그램.

사용 흐름
  1) 한글에서 사진대지 양식을 연다
  2) [한글 연결] 후, 작업 대상이 그 문서가 아니면 [문서 열기] 로 지정
  3) 사진을 목록에 담는다
  4) 한글에서 사진을 넣을 칸들을 드래그로 선택
  5) [감지된 칸에 넣기] — 목록 순서대로 삽입
     (칸 하나씩 넣으려면 칸을 클릭하고 [커서 칸에 한 장])

필요 환경 : Windows + 한글(HWP) 설치
설치      : pip install -r requirements.txt
실행      : python hwp_photo_macro.py
"""
from __future__ import annotations

import functools
import os
import re
import sys
import tkinter as tk
import traceback
from datetime import datetime
from tkinter import filedialog, messagebox, scrolledtext, ttk

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

# 한글 내부 단위
HWPUNIT_PER_MM = 283.465
HWPUNIT_PER_PX = 75          # 96dpi 기준: 1inch = 7200 HWPUNIT = 96px
IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp")
MAX_CELLS = 400              # 무한 루프 방지용 상한
FIELD_TAG = "__img_macro_target__"


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


LOG_FILE = os.path.join(app_dir(), "error_log.txt")


def natural_key(path: str):
    name = os.path.basename(path)
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


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

    def connect(self) -> bool:
        """실행 중인 한글에 먼저 붙어보고, 안 되면 새 인스턴스를 띄운다.

        GetActiveObject 가 실패하는 환경이 있어 폴백이 필요하다.
        다만 폴백으로 뜬 인스턴스는 사용자가 보고 있는 창과 다르므로,
        그 경우 [문서 열기] 로 작업할 문서를 지정해야 한다.
        돌려주는 값은 실행 중인 한글에 붙었는지 여부."""
        try:
            import win32com.client as win32
        except ImportError:
            raise HwpError("pywin32 가 설치되어 있지 않습니다.\n  pip install pywin32")

        hwp, attached = None, False
        try:
            hwp = win32.GetActiveObject("HWPFrame.HwpObject")
            attached = True
        except Exception:
            hwp = None

        if hwp is None:
            try:
                hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            except Exception:
                try:
                    hwp = win32.Dispatch("HWPFrame.HwpObject")
                except Exception as e:
                    raise HwpError("한글에 연결하지 못했습니다.\n"
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

    def open_document(self, path: str) -> bool:
        """작업할 문서를 연다. 이미 그 문서가 대상이면 건너뛴다.

        같은 파일을 중복해서 열면 '공유 위반' 이 뜨거나 편집이 제한된다."""
        self._require()
        current = self.doc_path()
        if current and os.path.normcase(os.path.abspath(current)) == \
                os.path.normcase(os.path.abspath(path)):
            return False
        last = None
        for args in ((path,), (path, "", "forceopen:true"), (path, "", "")):
            try:
                self.hwp.Open(*args)
                return True
            except Exception as e:
                last = e
        raise HwpError(f"문서를 열지 못했습니다.\n{path}\n\n({last})")

    def _require(self):
        if self.hwp is None:
            raise HwpError("한글에 연결되어 있지 않습니다. [한글 연결]을 먼저 눌러주세요.")

    # ---------------- 문서 ----------------
    def doc_path(self) -> str:
        self._require()
        try:
            return self.hwp.Path or ""
        except Exception:
            return ""

    def doc_name(self) -> str:
        p = self.doc_path()
        return os.path.basename(p) if p else "(저장되지 않은 문서)"

    def documents(self):
        self._require()
        names = []
        try:
            docs = self.hwp.XHwpDocuments
            for i in range(docs.Count):
                try:
                    p = docs.Item(i).FullName
                except Exception:
                    p = ""
                names.append(os.path.basename(p) if p else "(저장 안 된 문서)")
        except Exception as e:
            names.append(f"(목록을 읽지 못함: {e})")
        return names

    # ---------------- 커서 · 칸 ----------------
    def run(self, action: str) -> bool:
        self._require()
        try:
            return bool(self.hwp.Run(action))
        except Exception:
            return False

    def get_pos(self):
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
        """개체 선택 상태를 풀어 글자 편집 상태로 되돌린다."""
        try:
            self.hwp.Run("Cancel")
        except Exception:
            pass

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

    def goto_cell(self, list_id: int) -> bool:
        return self.set_pos((list_id, 0, 0)) and self.cell_size() is not None

    # ---------------- 선택 영역 읽기 ----------------
    def _set_field_name(self, name: str, option: int = 0) -> bool:
        """현재 위치(또는 셀 블록 전체)에 셀필드 이름을 붙인다.

        인자 순서는 (필드명, 안내문, 메모, 옵션) 이다.
        예전에 두 번째 자리에 옵션 숫자를 넣었더니 안내문 자리로 들어가
        선택 영역 전체가 아니라 커서가 있는 칸 하나에만 이름이 붙었다.
        실제 예제들은 이름 하나만 넘기므로 그 형태를 먼저 시도한다."""
        for args in ((name,), (name, "", "", option), (name, "", ""), (name, option)):
            try:
                self.hwp.SetCurFieldName(*args)
                return True
            except Exception:
                continue
        return False

    def _count_field(self, name: str) -> int:
        try:
            raw = self.hwp.GetFieldList(1, 0)
        except Exception:
            return 0
        if not raw:
            return 0
        return sum(1 for it in re.split(r"[\x02\r\n]+", str(raw))
                   if it and it.split("{{")[0] == name)

    def _goto_field(self, name: str, idx: int) -> bool:
        target = f"{name}{{{{{idx}}}}}"
        for args in ((target, True, True, False), (target,)):
            try:
                if self.hwp.MoveToField(*args):
                    return True
            except Exception:
                continue
        return False

    def _clear_field_name(self, name: str):
        for _ in range(MAX_CELLS):
            if not self._goto_field(name, 0):
                break
            if not self._set_field_name(""):
                break

    def selected_cells(self, name: str = FIELD_TAG, option: int = 0):
        """드래그로 선택한 칸들의 번호를 순서대로 돌려준다.

        셀 블록 상태에서 SetCurFieldName 을 실행하면 선택된 모든 칸에
        같은 이름이 붙는다. 그 이름을 인덱스로 순회해 칸 번호를 모으고
        이름은 지운다. 범위를 추측하지 않으므로 세로 선택이나
        떨어진 선택도 그대로 처리된다."""
        self._require()
        if not self._set_field_name(name, option):
            return []
        # 예제들이 이름을 붙인 직후 Cancel 로 셀 블록을 푼다. 그래야
        # 이어지는 MoveToField 가 제대로 동작한다.
        self.escape_selection()
        count = self._count_field(name)
        if count == 0:
            return []
        ids = []
        for i in range(count):
            if self._goto_field(name, i):
                p = self.get_pos()
                if p and self.cell_size() is not None:
                    ids.append(int(p[0]))
        self._clear_field_name(name)
        return ids

    def block_cell_addresses(self):
        """선택된 블록에 들어 있는 칸들의 (열, 행) 주소를 뽑는다.

        HWPML2X 로 블록만 내보내면 <CELL ColAddr="2" RowAddr="4" ...> 형태로
        각 칸의 주소가 들어 있다. 선택을 건드리지 않는다."""
        self._require()
        try:
            raw = self.hwp.GetTextFile("HWPML2X", "saveblock:true")
        except Exception as e:
            raise HwpError(f"블록을 읽지 못했습니다.\n({e})")
        if not raw:
            return []
        text = str(raw)
        addrs = []
        for m in re.finditer(r"<CELL\b[^>]*>", text):
            tag = m.group(0)
            col = re.search(r'ColAddr="(\d+)"', tag)
            row = re.search(r'RowAddr="(\d+)"', tag)
            if col and row:
                addrs.append((int(col.group(1)), int(row.group(1))))
        return addrs

    def cell_addr(self):
        """현재 커서가 있는 칸의 주소 문자열. 상태 표시줄의 (D5) 와 같은 값."""
        self._require()
        try:
            ind = self.hwp.KeyIndicator()
        except Exception:
            return None
        try:
            for part in ind:
                if isinstance(part, str):
                    m = re.search(r"\(([A-Z]+\d+)\)", part)
                    if m:
                        return m.group(1)
        except TypeError:
            pass
        return str(ind)

    def _probe_field_count(self):
        """이름을 붙인 직후 개수만 세어본다. 몇 칸에 붙었는지 바로 알 수 있다."""
        name = FIELD_TAG + "_probe"
        ok = self._set_field_name(name)
        n = self._count_field(name)
        self.escape_selection()
        self._clear_field_name(name)
        return f"SetCurFieldName 성공={ok}, 이름 붙은 칸={n}"

    def block_text_probe(self):
        """선택된 블록의 텍스트를 그대로 읽는다. 선택을 건드리지 않는다.

        칸이 비어 있어도 칸마다 구분자가 들어가므로,
        조각 개수로 한글이 몇 칸을 블록으로 보고 있는지 알 수 있다."""
        self._require()
        out = {}
        for fmt, opt in (("TEXT", "saveblock:true"), ("TEXT", "saveblock"),
                         ("HWPML2X", "saveblock:true")):
            key = f"GetTextFile({fmt},{opt})"
            try:
                raw = self.hwp.GetTextFile(fmt, opt)
            except Exception as e:
                out[key] = f"실패: {e}"
                continue
            if raw is None:
                out[key] = "None"
                continue
            text = str(raw)
            if fmt == "HWPML2X":
                out[key] = f"길이 {len(text)}, <CELL 개수 {text.count('<CELL ')}"
            else:
                segs = [x for x in re.split(r"\r\n|\n|\r", text)]
                out[key] = f"길이 {len(text)}, 조각 {len(segs)}개, 앞부분 {text[:40]!r}"
        return out

    def selection_mode(self):
        """0이 아니면 무언가 선택된 상태. 셀 블록도 여기에 잡힌다."""
        try:
            return int(self.hwp.SelectionMode)
        except Exception:
            return None

    def selection_probe(self, option: int = 0):
        """선택 영역을 읽는 방법들의 결과를 그대로 모은다 (진단용).

        앞쪽은 선택을 건드리지 않는 검사, 뒤쪽은 선택을 소모하는 검사다."""
        self._require()
        report = {}

        # (1) 선택을 건드리지 않는 검사
        report["SelectionMode"] = self.selection_mode()
        report["GetPos"] = self.get_pos()
        report["표 안 여부"] = self.in_table()
        try:
            report["GetSelectedPos"] = self.hwp.GetSelectedPos()
        except Exception as e:
            report["GetSelectedPos"] = f"실패: {e}"
        report["셀 주소(KeyIndicator)"] = self.cell_addr()
        try:
            addrs = self.block_cell_addresses()
            report["블록 칸 주소"] = f"{len(addrs)}개 {addrs[:20]}"
        except Exception as e:
            report["블록 칸 주소"] = f"실패: {e}"
        try:
            report.update(self.block_text_probe())
        except Exception as e:
            report["블록 텍스트"] = f"실패: {e}"

        # (2) 여기서부터는 선택이 사라진다
        try:
            report["필드 붙인 결과"] = self._probe_field_count()
        except Exception as e:
            report["필드 붙인 결과"] = f"실패: {e}"

        report["열린 문서"] = self.documents()
        return report

    # ---------------- 삽입 ----------------
    @staticmethod
    def _image_ratio(path: str):
        if Image is None:
            return 4, 3
        try:
            with Image.open(path) as im:
                return im.size
        except Exception:
            return 4, 3

    def insert_picture_fit(self, path: str, margins_px=(0, 0, 0, 0),
                           border_px=0, post_adjust: bool = False):
        """현재 칸 크기에 맞춰 비율을 유지한 채 삽입.
        margins_px = (상, 하, 좌, 우) 픽셀"""
        self._require()
        if not os.path.exists(path):
            raise HwpError(f"사진 파일을 찾을 수 없습니다.\n{path}")
        size = self.cell_size()
        if size is None:
            raise HwpError("커서가 표 안에 있지 않습니다.\n"
                           "한글에서 사진을 넣을 칸을 클릭한 뒤 다시 시도해 주세요.")

        cw, ch = size
        top, bottom, left, right = (int(m * HWPUNIT_PER_PX) for m in margins_px)
        avail_w = max(cw - left - right, 1)
        avail_h = max(ch - top - bottom, 1)

        iw, ih = self._image_ratio(path)
        scale = min(avail_w / iw, avail_h / ih)
        w = max(int(iw * scale), 1)
        h = max(int(ih * scale), 1)

        hwp = self.hwp
        hwp.Run("ParagraphShapeAlignCenter")

        # sizeoption 3 = width/height 로 지정한 크기.
        # 2 는 "셀 크기에 맞추어" 라서 계산한 값을 무시하고 여백이 먹지 않는다.
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

        # 그림이 선택된 채로 남으면 다음 동작이 막힌다
        self.escape_selection()
        return (cw, ch), (w, h)

    def _apply_shape(self, w: int, h: int, border_px: int):
        """삽입한 그림의 크기와 테두리를 다시 지정한다.

        지정하지 않은 속성까지 기본값으로 함께 적용되므로,
        그림이 선택되지 않거나 배치가 이상해지면 이 단계를 꺼야 한다."""
        hwp = self.hwp
        try:
            hwp.FindCtrl()
            hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)
            so = hwp.HParameterSet.HShapeObject
            so.Width = w
            so.Height = h
            so.TreatAsChar = 1
            for key in ("Lock", "Protect"):     # 개체 보호가 켜지면 선택이 안 된다
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
            self.log(f"[오류] {func.__name__}\n{traceback.format_exc()}")
            messagebox.showerror(
                "오류",
                f"{type(e).__name__}: {e}\n\n기록: {LOG_FILE}", parent=self)
    return wrapper


# ==================================================================
#  GUI
# ==================================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("사진대지 매크로 (한글)")
        self.geometry("900x680")
        self.minsize(860, 620)

        self.ctrl = HwpController()
        self.photos = []
        self.cursor = 0
        self._thumb = None

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._refresh()
        self.log("프로그램을 시작했습니다.")
        if Image is None:
            self.log("[경고] Pillow 를 불러오지 못해 미리보기가 꺼졌습니다.")

    # ---------------- 로그 ----------------
    def log(self, message: str):
        line = f"[{datetime.now():%H:%M:%S}] {message}\n"
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
        self.lbl_status = ttk.Label(bar, text="연결 안 됨", foreground="#b00")
        self.lbl_status.pack(side="left", padx=10)

        body = ttk.Frame(root)
        body.pack(fill="both", expand=True)

        # ---- 왼쪽: 미리보기 · 목록 · 기록 ----
        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True)

        self.canvas = tk.Canvas(left, bg="#1a1a1a", height=230, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        list_area = ttk.Frame(left)
        list_area.pack(fill="both", expand=True, pady=(8, 0))

        self.tree = ttk.Treeview(list_area, columns=("no", "file"),
                                 show="headings", height=8)
        self.tree.heading("no", text="#")
        self.tree.heading("file", text="파일명")
        self.tree.column("no", width=45, anchor="center", stretch=False)
        self.tree.column("file", width=380)
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

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

        # ---- 오른쪽: 삽입 · 옵션 ----
        right = ttk.Frame(body, width=265)
        right.pack(side="left", fill="y", padx=(10, 0))
        right.pack_propagate(False)

        act = ttk.LabelFrame(right, text="삽입", padding=8)
        act.pack(fill="x")
        ttk.Button(act, text="선택한 칸에 넣기",
                   command=self.on_insert_selection).pack(fill="x", pady=2)
        ttk.Label(act, text="한글에서 칸들을 드래그로 선택한 뒤 누르세요",
                  foreground="#666", wraplength=225).pack(anchor="w")
        ttk.Button(act, text="커서 칸에 한 장",
                   command=self.on_insert_one).pack(fill="x", pady=(8, 2))
        ttk.Button(act, text="넣을 위치 처음으로",
                   command=self.on_reset_cursor).pack(fill="x", pady=2)
        self.lbl_progress = ttk.Label(act, text="", foreground="#555", wraplength=225)
        self.lbl_progress.pack(fill="x", pady=(4, 0))

        opt = ttk.LabelFrame(right, text="여백 주기 (px)", padding=8)
        opt.pack(fill="x", pady=8)
        self.var_margin = {}
        for i, key in enumerate(("상", "하", "좌", "우")):
            ttk.Label(opt, text=key).grid(row=i, column=0, sticky="w", pady=1)
            v = tk.IntVar(value=0)
            self.var_margin[key] = v
            ttk.Spinbox(opt, from_=0, to=300, increment=1, width=8,
                        textvariable=v).grid(row=i, column=1, sticky="e")

        bd = ttk.LabelFrame(right, text="사진 속성", padding=8)
        bd.pack(fill="x")
        self.var_border_on = tk.BooleanVar(value=False)
        ttk.Checkbutton(bd, text="테두리 (px)", variable=self.var_border_on)\
            .grid(row=0, column=0, sticky="w")
        self.var_border_px = tk.IntVar(value=1)
        ttk.Spinbox(bd, from_=1, to=20, increment=1, width=6,
                    textvariable=self.var_border_px).grid(row=0, column=1, sticky="e")
        self.var_post_adjust = tk.BooleanVar(value=False)
        ttk.Checkbutton(bd, text="삽입 후 크기 다시 지정",
                        variable=self.var_post_adjust)\
            .grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Label(bd, text="그림이 선택되지 않으면 꺼두세요",
                  foreground="#666", wraplength=225)\
            .grid(row=2, column=0, columnspan=2, sticky="w")

        diag = ttk.LabelFrame(right, text="진단", padding=8)
        diag.pack(fill="x", pady=(8, 0))
        row = ttk.Frame(diag)
        row.pack(fill="x")
        ttk.Label(row, text="셀필드 option").pack(side="left")
        self.var_field_option = tk.IntVar(value=0)
        ttk.Spinbox(row, from_=0, to=7, increment=1, width=5,
                    textvariable=self.var_field_option).pack(side="right")
        ttk.Label(diag, text="보통 0 으로 둡니다. 칸이 1개로만 잡힐 때만 바꿔 시험",
                  foreground="#666", wraplength=225).pack(anchor="w", pady=(2, 4))
        ttk.Button(diag, text="선택 영역 확인",
                   command=self.on_check_selection).pack(fill="x", pady=2)
        ttk.Button(diag, text="커서 위치 확인",
                   command=self.on_check_cursor).pack(fill="x", pady=2)

    # ---------------- 선택 영역 감시 ----------------
    def _refresh(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for i, p in enumerate(self.photos):
            mark = "✔" if i < self.cursor else str(i + 1)
            self.tree.insert("", "end", iid=str(i),
                             values=(mark, os.path.basename(p)))
        total = len(self.photos)
        self.lbl_progress.config(
            text=f"{self.cursor} / {total} 장 삽입됨" if total else "사진을 추가해 주세요.")

    def _selected_index(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    @guarded
    def on_add(self):
        chosen = filedialog.askopenfilenames(
            parent=self, title="사진 선택",
            filetypes=[("이미지 파일", tuple("*" + e for e in IMAGE_EXTS)),
                       ("모든 파일", "*.*")])
        paths = list(self.tk.splitlist(chosen)) if chosen else []
        self.log(f"사진 선택: {len(paths)}개")
        self._add_paths(paths)          # 고른 순서 그대로 담는다

    @guarded
    def on_add_folder(self):
        folder = filedialog.askdirectory(parent=self, title="사진 폴더 선택")
        if not folder:
            self.log("폴더 선택을 취소했습니다.")
            return
        names = [f for f in os.listdir(folder) if f.lower().endswith(IMAGE_EXTS)]
        paths = sorted((os.path.join(folder, f) for f in names), key=natural_key)
        self.log(f"폴더 선택: 이미지 {len(paths)}개")
        if not paths:
            messagebox.showinfo("안내", "폴더 안에 이미지 파일이 없습니다.", parent=self)
            return
        self._add_paths(paths)

    def _add_paths(self, paths):
        existing = set(self.photos)
        added = 0
        for p in paths:
            if p not in existing:
                self.photos.append(p)
                existing.add(p)
                added += 1
        self._refresh()
        self.log(f"목록에 {added}장 추가 (전체 {len(self.photos)}장)")

    @guarded
    def on_sort_by_name(self):
        if not self.photos:
            return
        self.photos.sort(key=natural_key)
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
            self._show_preview(self.photos[i])

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

    # ---------------- 한글 연동 ----------------
    def _update_status(self):
        if not self.ctrl.connected:
            self.lbl_status.config(text="연결 안 됨", foreground="#b00")
            return
        saved = bool(self.ctrl.doc_path())
        self.lbl_status.config(text=f"작업 대상: {self.ctrl.doc_name()}",
                               foreground="#070" if saved else "#c60")

    @guarded
    def on_connect(self):
        attached = self.ctrl.connect()
        self._update_status()
        docs = self.ctrl.documents()
        self.log("실행 중인 한글에 연결했습니다." if attached
                 else "[주의] 실행 중인 한글에 붙지 못해 새 인스턴스를 띄웠습니다.")
        self.log(f"작업 대상: {self.ctrl.doc_name()} / 열린 문서 {len(docs)}개: {docs}")
        if not attached or not self.ctrl.doc_path():
            messagebox.showinfo(
                "작업할 문서를 지정해 주세요",
                "지금 연결된 한글은 사용자가 보고 있는 창과 다를 수 있습니다.\n\n"
                "[문서 열기] 로 사진을 넣을 문서를 직접 골라주세요.\n"
                "그 창에서 칸을 선택하셔야 프로그램이 알아볼 수 있습니다.",
                parent=self)

    @guarded
    def on_open_doc(self):
        """작업할 한글 문서를 직접 지정한다."""
        if not self.ctrl.connected:
            self.ctrl.connect()
        path = filedialog.askopenfilename(
            parent=self, title="작업할 한글 문서 선택",
            filetypes=[("한글 문서", ("*.hwp", "*.hwpx")), ("모든 파일", "*.*")])
        if not path:
            self.log("문서 선택을 취소했습니다.")
            return
        opened = self.ctrl.open_document(path)
        self._update_status()
        self.log(("문서를 열었습니다: " if opened else "이미 열려 있는 문서입니다: ")
                 + os.path.basename(path))
        docs = self.ctrl.documents()
        if len(docs) > 1:
            messagebox.showwarning(
                "문서가 여러 개 열려 있습니다",
                f"연결된 한글에 문서가 {len(docs)}개 열려 있습니다.\n"
                "어느 문서를 보는지 보장되지 않아 커서를 잘못 읽을 수 있습니다.\n\n"
                "작업할 문서 하나만 남기고 닫아주세요.", parent=self)

    @staticmethod
    def _mm(size):
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

    def _check_saved(self) -> bool:
        """저장되지 않은 문서에서는 한글이 되돌리기를 받아주지 않는다."""
        if self.ctrl.doc_path():
            return True
        self.log("[경고] 문서가 파일로 저장되지 않았습니다.")
        return messagebox.askyesno(
            "저장되지 않은 문서입니다",
            "이 문서는 아직 파일로 저장되지 않았습니다.\n\n"
            "이 상태에서는 한글의 되돌리기(Ctrl+Z)가 동작하지 않습니다.\n"
            "한글에서 먼저 저장하시기를 권합니다.\n\n"
            "그래도 이대로 진행할까요?", parent=self)

    def _ready(self) -> bool:
        if not self.ctrl.connected:
            messagebox.showinfo("안내", "먼저 [한글 연결]을 눌러주세요.", parent=self)
            return False
        if not self.photos:
            messagebox.showinfo("안내", "사진 목록이 비어 있습니다.", parent=self)
            return False
        if self.cursor >= len(self.photos):
            messagebox.showinfo("안내", "목록의 사진을 모두 넣었습니다.", parent=self)
            return False
        return self._check_saved()

    @guarded
    def on_reset_cursor(self):
        self.cursor = 0
        self._refresh()
        self.log("넣을 위치를 목록 처음으로 되돌렸습니다.")

    @guarded
    def on_insert_one(self):
        """한글에서 커서가 있는 칸에 한 장 넣는다."""
        if not self._ready():
            return
        margins, border = self._insert_options()
        path = self.photos[self.cursor]
        cell, pic = self.ctrl.insert_picture_fit(
            path, margins, border, post_adjust=self.var_post_adjust.get())
        self.cursor += 1
        self.log(f"삽입: {os.path.basename(path)} — 칸 {self._mm(cell)}, 사진 {self._mm(pic)}")
        self._refresh()

    @guarded
    def on_insert_selection(self):
        """한글에서 드래그로 선택한 칸들에 목록 순서대로 넣는다."""
        if not self._ready():
            return
        cells = self.ctrl.selected_cells(option=self.var_field_option.get())
        if not cells:
            messagebox.showinfo(
                "칸을 선택해 주세요",
                "선택된 칸을 찾지 못했습니다.\n\n"
                "한글에서 사진을 넣을 칸들을 드래그로 선택한 뒤\n"
                "다시 이 버튼을 눌러주세요.",
                parent=self)
            return

        remaining = len(self.photos) - self.cursor
        count = min(len(cells), remaining)
        self.log(f"선택한 칸 {len(cells)}개(option={self.var_field_option.get()}), "
                 f"넣을 사진 {remaining}장 → {count}장 삽입")

        if len(cells) == 1 and remaining > 1:
            if not messagebox.askyesno(
                    "칸이 1개로만 잡혔습니다",
                    "여러 칸을 선택하셨다면 [진단]의 셀필드 option 을\n"
                    "0 → 1 → 2 → 3 으로 바꿔가며 다시 드래그해 보세요.\n"
                    "드래그할 때마다 [감지된 칸] 숫자가 갱신됩니다.\n\n"
                    "이대로 1장만 넣을까요?", parent=self):
                self.log("사용자가 삽입을 취소했습니다.")
                return

        if len(cells) != remaining:
            if not messagebox.askyesno(
                    "개수가 다릅니다",
                    f"선택한 칸 : {len(cells)}개\n"
                    f"넣을 사진 : {remaining}장\n\n"
                    f"앞에서부터 {count}장만 넣습니다. 진행할까요?", parent=self):
                self.log("사용자가 삽입을 취소했습니다.")
                return

        margins, border = self._insert_options()
        inserted = 0
        for lid in cells[:count]:
            if not self.ctrl.goto_cell(lid):
                self.log(f"칸 {lid} 로 이동하지 못해 건너뜁니다.")
                continue
            path = self.photos[self.cursor]
            cell, pic = self.ctrl.insert_picture_fit(
                path, margins, border, post_adjust=self.var_post_adjust.get())
            self.cursor += 1
            inserted += 1
            self.log(f"삽입 {inserted}: {os.path.basename(path)} — "
                     f"칸 {self._mm(cell)}, 사진 {self._mm(pic)}")
        self.log(f"완료: 이번에 {inserted}장 삽입")
        self._refresh()

    # ---------------- 진단 ----------------
    @guarded
    def on_check_selection(self):
        if not self.ctrl.connected:
            messagebox.showinfo("안내", "먼저 [한글 연결]을 눌러주세요.", parent=self)
            return
        probe = self.ctrl.selection_probe(option=self.var_field_option.get())
        detail = "\n".join(f"{k} : {v}" for k, v in probe.items())
        self.log("선택 영역 확인\n" + detail)
        messagebox.showinfo("선택 영역 확인", detail, parent=self)

    @guarded
    def on_check_cursor(self):
        if not self.ctrl.connected:
            messagebox.showinfo("안내", "먼저 [한글 연결]을 눌러주세요.", parent=self)
            return
        pos = self.ctrl.get_pos()
        size = self.ctrl.cell_size()
        self.log(f"커서 확인: {pos}, 칸 {size}")
        messagebox.showinfo(
            "커서 위치 확인",
            f"커서 위치 : {pos}\n"
            f"칸 크기   : {size}\n"
            f"작업 대상 : {self.ctrl.doc_name()}\n"
            f"열린 문서 : {self.ctrl.documents()}\n\n"
            "한글에서 다른 칸을 클릭하고 다시 눌러보세요.\n"
            "위치가 바뀌지 않으면 프로그램이 보는 한글 창이\n"
            "지금 클릭하는 창과 다른 것입니다.", parent=self)



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
            messagebox.showerror("예기치 못한 오류", f"{exc}\n\n기록: {LOG_FILE}", parent=app)
        except Exception:
            pass

    sys.excepthook = excepthook
    app.report_callback_exception = lambda *a: excepthook(*a)
    app.mainloop()


if __name__ == "__main__":
    main()