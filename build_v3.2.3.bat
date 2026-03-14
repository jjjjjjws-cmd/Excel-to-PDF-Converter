@echo off
echo =====================================
echo Excel to PDF v3.2.3 - Build
echo =====================================
echo.

echo [1/4] Install packages...
python -m pip install selenium==4.15.0 openpyxl pyinstaller --break-system-packages
echo.

echo [2/4] Clean...
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
echo.

echo [3/4] Build...
python -m PyInstaller --onefile --noconsole --name Excel_to_PDF_v3.2.3 --hidden-import=selenium.webdriver.chrome.service --hidden-import=selenium.webdriver.common.service excel_to_pdf_v3.2.3.py
echo.

echo [4/4] Verify...
if exist "dist\Excel_to_PDF_v3.2.3.exe" (
    echo =====================================
    echo SUCCESS!
    echo File: dist\Excel_to_PDF_v3.2.3.exe
    echo =====================================
) else (
    echo =====================================
    echo ERROR!
    echo =====================================
)

pause
