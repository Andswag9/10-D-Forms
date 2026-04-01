@echo off
title CMBS Investor Reporting Tool

:: ============================================================
:: CMBS REPORTING TOOL LAUNCHER
:: Place this .bat file anywhere (Desktop, shortcut on Desktop)
:: Point SCRIPT_DIR to wherever you saved cmbs_report.py
:: ============================================================

set SCRIPT_DIR=S:\Lenders\Z. CMBS Test\6. CMBS Report - 10D Forms\Python

:: --- Find Python ---
set PYTHON=python
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    set PYTHON=py
    where py >nul 2>&1
    if %ERRORLEVEL% NEQ 0 (
        echo ERROR: Python not found on this machine.
        echo Please install Python from https://python.org
        pause
        exit /b 1
    )
)

:: --- Check if packages are already installed ---
echo Checking dependencies...
%PYTHON% -c "import openpyxl, xlrd, xlwt, xlutils, dateutil, win32com.client" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Installing required packages (one-time setup^)...
    :: --no-warn-script-location suppresses the PATH warning that caused
    :: the false "Failed to install" error. We verify via import, not exit code.
    %PYTHON% -m pip install openpyxl "xlrd==1.2.0" xlwt xlutils python-dateutil pywin32 --quiet --no-warn-script-location >nul 2>&1

    :: Verify the import actually works after install
    %PYTHON% -c "import openpyxl, xlrd, xlwt, xlutils, dateutil, win32com.client" >nul 2>&1
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo ERROR: Package install failed. Try running this manually in a terminal:
        echo   python -m pip install openpyxl "xlrd==1.2.0" xlwt xlutils python-dateutil pywin32
        pause
        exit /b 1
    )
    echo Packages installed successfully.
)

:: --- Run the tool ---
echo.
echo Starting CMBS Reporting Tool...
echo.
cd /d "%SCRIPT_DIR%"
%PYTHON% cmbs_report.py

:: Keep window open if there was an error
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Script exited with an error (code %ERRORLEVEL%^).
    pause
)
