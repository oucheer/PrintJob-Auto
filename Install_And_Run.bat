@echo off
setlocal
chcp 65001 >nul 2>&1
cd /d "%~dp0"
title PrintCase Studio - One-click install and launch

echo ============================================================
echo  PrintCase Studio - One-click install and launch
echo  First run prepares Python, requirements and print component
echo ============================================================
echo.

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_windows.ps1" -Launch
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo [FAILED] Installation or startup did not complete.
    echo Please send logs\setup.log to product support.
    echo.
    pause
)

endlocal & exit /b %RC%
