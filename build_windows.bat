@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
    echo Python launcher not found. Install Python from python.org first.
    pause
    exit /b 1
)

py -3 -m pip install -r requirements.txt
if errorlevel 1 goto :error

py -3 -m PyInstaller --clean --noconfirm ge_tool.spec
if errorlevel 1 goto :error

echo Build completed. See the dist folder.
pause
exit /b 0

:error
echo Build failed.
pause
exit /b 1
