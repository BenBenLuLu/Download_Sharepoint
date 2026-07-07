@echo off
REM ── URL Batch Downloader ── Build script for Windows ──────────────────────
REM Run this file on your Windows machine to produce URL_Batch_Downloader.exe
REM Prerequisites: Python 3.10+, pip

echo [1/3] Installing / updating dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [2/3] Building executable...
pyinstaller url_downloader.spec --clean

echo.
echo [3/3] Done.
echo Output: dist\URL_Batch_Downloader.exe
pause
