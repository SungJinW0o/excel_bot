@echo off
REM ===================================================
REM Excel Bot - Enhanced Safe Runner
REM Author: SungJinWoo
REM ===================================================

REM 1. Set script folder as working directory
cd /d "%~dp0"

REM 2. Create input/output folders if they don't exist
if not exist "input_data" mkdir "input_data"
if not exist "output_data" mkdir "output_data"

REM 3. Activate virtual environment
call .venv\Scripts\activate

REM 4. DRY_RUN toggle: true=test (no emails), false=real run
set /p DRY_RUN="DRY_RUN mode? Enter true or false (default: true): "
if "%DRY_RUN%"=="" set DRY_RUN=true
echo DRY_RUN=%DRY_RUN%

REM 5. Optional: SMTP/env settings (uncomment and edit if needed)
REM set SMTP_USER=your-email@gmail.com
REM set SMTP_PASS=your-app-password
REM set SMTP_HOST=smtp.gmail.com
REM set SMTP_PORT=587
REM set SMTP_SENDER=your-email@gmail.com

REM 6. Run the bot
echo.
echo Running Excel Bot...
excel-bot --dry-run %DRY_RUN%
set BOT_EXIT_CODE=%ERRORLEVEL%

REM 7. Open summary report if exists
if exist "output_data\summary_report.xlsx" (
    start "" "output_data\summary_report.xlsx"
)

REM 8. Display log summary if DRY_RUN
if "%DRY_RUN%"=="true" (
    echo.
    echo DRY_RUN: Logs/events captured in output_data.
    echo.
)

REM 9. Pause for user review
echo.
echo Excel Bot finished with exit code %BOT_EXIT_CODE%.
pause
exit /b %BOT_EXIT_CODE%
