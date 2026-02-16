@echo off
chcp 65001 >nul
echo Сборка KP_Generator.exe ...
echo.

set ROOT=%~dp0
cd /d "%ROOT%"

pip install -r requirements.txt -q
pip install pyinstaller -q

pyinstaller --noconfirm --onefile --windowed --name "KP_Generator" ^
  --add-data "kp_generator/assets;kp_generator/assets" ^
  --hidden-import "pytesseract" ^
  --hidden-import "openpyxl" ^
  --hidden-import "docx" ^
  --hidden-import "PIL" ^
  --hidden-import "fitz" ^
  kp_generator/app.py

if %ERRORLEVEL% equ 0 (
    echo.
    echo Готово. EXE: dist\KP_Generator.exe
    echo Рядом с exe можно положить папку tesseract с tesseract.exe для OCR PDF.
) else (
    echo Ошибка сборки.
    exit /b 1
)
