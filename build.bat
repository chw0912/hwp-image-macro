@echo off
chcp 65001 > nul
echo ============================================
echo   사진대지 매크로 - 실행파일 빌드
echo ============================================
echo.

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

pyinstaller --onefile --windowed ^
  --name hwp_image_macro ^
  --hidden-import win32timezone ^
  --clean ^
  hwp_photo_macro.py

echo.
echo 빌드 완료: dist\hwp_image_macro.exe
pause