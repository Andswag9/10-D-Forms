@echo off
echo ============================================================
echo   10-D: Split signed PDF, then create Outlook drafts
echo ============================================================
echo.

cd /d "%~dp0..\Step 2. Split 10 - D Script\"
python split_email_10d_forms.py --email-after
set EX=%ERRORLEVEL%

if %EX% neq 0 (
    echo.
    echo Something went wrong. Check the messages above.
    pause
    exit /b %EX%
)

echo.
pause
