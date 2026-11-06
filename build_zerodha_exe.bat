@echo off
REM Builds a single-file Windows .exe with the Zerodha icon (needs Python installed)
cd /d "%~dp0"
if not exist zerodha.ico (
    echo ERROR: zerodha.ico not found. Put zerodha.ico in this same folder and run again.
    pause
    exit /b 1
)
if not exist build_venv (
    python -m venv build_venv
)
call build_venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install requests pyinstaller
pyinstaller --onefile --windowed --clean --icon zerodha.ico --add-data "zerodha.ico;." --name ZerodhaOptionsTrader zerodha_options_trader.py
echo.
echo Done. Your app is in the "dist" folder: dist\ZerodhaOptionsTrader.exe
pause
