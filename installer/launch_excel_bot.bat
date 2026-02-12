@echo off
setlocal EnableExtensions

cd /d "%~dp0"

if exist ".venv\Scripts\activate.bat" (
    call ".venv\Scripts\activate.bat"
)

set /p DRY_RUN="DRY_RUN mode? Enter true or false (default: true): "
if "%DRY_RUN%"=="" set DRY_RUN=true

python run_bot.py --dry-run %DRY_RUN%
set "BOT_EXIT_CODE=%ERRORLEVEL%"

echo.
echo Excel Bot finished with exit code %BOT_EXIT_CODE%.
pause
exit /b %BOT_EXIT_CODE%
