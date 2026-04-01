@echo off
echo ============================================================
echo   10-D Email Draft Creator (drafts only — run after split)
echo   For split + drafts in one step use: Split_and_Email_10D.bat
echo ============================================================
echo.

cd /d "%~dp0"
python create_10d_emails.py

REM --- Keep window open if there was an error ---
if errorlevel 1 (
    echo.
    echo Something went wrong. Check the error above.
    pause
)
