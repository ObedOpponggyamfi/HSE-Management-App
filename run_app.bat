@echo off
setlocal enableextensions
title HSE Management App
cd /d "%~dp0"

REM ===== change this if you want a different port =====
set "PORT=5050"

echo ============================================================
echo   ASANKO GOLD MINE - HSE MANAGEMENT APP
echo ============================================================
echo.

REM ----- locate a Python interpreter -----
set "PYTHON="
where python >nul 2>nul && set "PYTHON=python"
if not defined PYTHON where py >nul 2>nul && set "PYTHON=py"
if not defined PYTHON (
    echo [ERROR] Python 3 was not found on your PATH.
    echo Install it from https://www.python.org/downloads/ and tick "Add Python to PATH".
    echo.
    pause
    exit /b 1
)
echo Using Python: %PYTHON%

REM ----- install dependencies only if something is missing -----
%PYTHON% -c "import flask, flask_sqlalchemy, flask_login, flask_wtf, pandas, openpyxl, numpy" >nul 2>nul
if errorlevel 1 (
    echo Installing required packages ^(first run only^)...
    %PYTHON% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Could not install dependencies. Check your internet connection.
        echo.
        pause
        exit /b 1
    )
)

REM ----- create sample data on first run -----
if not exist "data\incident_register.xlsx" (
    echo Generating sample data...
    %PYTHON% generate_dummy_data.py
)

REM ----- launch (app.py opens your browser automatically) -----
echo.
echo Starting the app on http://127.0.0.1:%PORT%/
echo Your browser will open shortly. Keep this window open while using the app.
echo Press Ctrl+C here to stop the app.
echo.
%PYTHON% app.py

echo.
echo The app has stopped.
pause
endlocal
