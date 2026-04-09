@echo off
echo ============================================
echo   ChromeOpener - EXE Build
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo         Download from https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Install PyInstaller
echo [1/3] Installing PyInstaller...
pip install pyinstaller>=6.0 --quiet
echo.

:: Build EXE
echo [2/3] Building ChromeOpener.exe...
pyinstaller chrome_opener.spec --distpath ./release --workpath ./build_temp --noconfirm
echo.

:: Clean up build artifacts
echo [3/3] Cleaning up...
rmdir /s /q build_temp 2>nul

echo.
echo ============================================
echo   Build complete!
echo.
echo   Output: release\ChromeOpener.exe
echo.
echo   Usage:
echo     ChromeOpener.exe             Open all URLs
echo     ChromeOpener.exe --add URL   Add a URL
echo     ChromeOpener.exe --remove 0  Remove URL at index
echo     ChromeOpener.exe --list      Show all URLs
echo     ChromeOpener.exe --edit      Edit config.json directly
echo     ChromeOpener.exe --clear     Remove all URLs
echo.
echo   config.json is saved next to the exe.
echo   You can also edit config.json directly with Notepad.
echo ============================================
pause
