@echo off
REM Builds a single-file Windows .exe with the Dhan icon (needs Python installed)
cd /d "%~dp0"
if not exist dhan.ico (
    echo ERROR: dhan.ico not found. Put dhan.ico in this same folder and run again.
    pause
    exit /b 1
)
if not exist build_venv (
    python -m venv build_venv
)
call build_venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install requests pyinstaller
pyinstaller --onefile --windowed --clean --icon dhan.ico --add-data "dhan.ico;." --name DhanOptionsTrader dhan_options_trader.py
echo.
echo Done. Your app is in the "dist" folder: dist\DhanOptionsTrader.exe
pause
