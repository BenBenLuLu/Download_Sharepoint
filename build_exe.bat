@echo off
REM ── URL Batch Downloader ── Build script for Windows ──────────────────────
REM Double-click this file to produce dist\URL_Batch_Downloader.exe
REM Prerequisites: Python 3.10+ added to PATH

echo [1/3] Installing / updating dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: pip install failed. Make sure Python is in PATH.
    pause & exit /b 1
)

echo.
echo [2/3] Building executable (this may take 1-3 minutes)...
pyinstaller url_downloader.spec --clean --noconfirm
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    pause & exit /b 1
)

echo.
echo [3/3] Done!
echo Output: dist\URL_Batch_Downloader.exe
pause
