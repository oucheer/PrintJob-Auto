@echo off
setlocal
chcp 65001 >nul 2>&1
cd /d "%~dp0"
title PrintCase Studio

rem Run the same self-healing installer/launcher. Safe to double-click repeatedly.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_windows.ps1" -Launch
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo.
  echo Application setup/startup failed. See logs\setup.log.
  pause
)
endlocal & exit /b %RC%
