@echo off
REM Builds a single-file Windows .exe with the Fyers icon (needs Python installed)
cd /d "%~dp0"
if not exist fyers.ico (
    echo ERROR: fyers.ico not found. Put fyers.ico in this same folder and run again.
    pause
    exit /b 1
)
if not exist build_venv (
    python -m venv build_venv
)
call build_venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install requests pyinstaller
pyinstaller --onefile --windowed --clean --icon fyers.ico --add-data "fyers.ico;." --name FyersOptionsTrader fyers_options_trader.py
echo.
echo Done. Your app is in the "dist" folder: dist\FyersOptionsTrader.exe
pause
