@echo off
title PulmoDoc Suite Launcher
cd /d "%~dp0"

echo ====================================================================
echo                 PULMODOC SUITE — MEDICAL PG PLATFORM
echo ====================================================================
echo Starting local clinical education server...
echo.
echo Standalone Web Portal: %~dp0launch_app.html
echo.

where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    python pulmo_server.py
) else if exist "C:\Users\harshith k r\AppData\Local\Programs\Python\Python312\python.exe" (
    "C:\Users\harshith k r\AppData\Local\Programs\Python\Python312\python.exe" pulmo_server.py
) else (
    echo Python executable not found on standard paths.
    echo Opening standalone offline medical portal in browser...
    start "" "%~dp0launch_app.html"
)

pause
