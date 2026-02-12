@echo off
setlocal EnableExtensions

cd /d "%~dp0"

if not exist "input_data" mkdir "input_data"
if not exist "output_data" mkdir "output_data"

set "PYTHON_EXE=python"
if exist ".venv\Scripts\python.exe" set "PYTHON_EXE=.venv\Scripts\python.exe"

echo ==================================================
echo                   Excel Bot
echo ==================================================
echo.
echo Select run mode:
echo   [1] Safe test run (DRY_RUN=true)  ^(Recommended^)
echo   [2] Live run (DRY_RUN=false)
echo.

set "MODE=1"
set /p MODE="Enter 1 or 2 (default: 1): "
if "%MODE%"=="" set "MODE=1"

if "%MODE%"=="2" (
    set "DRY_RUN=false"
) else (
    if not "%MODE%"=="1" (
        echo Invalid selection. Using safe test run.
    )
    set "DRY_RUN=true"
)

echo.
echo Running with DRY_RUN=%DRY_RUN%
echo Input folder : "%CD%\input_data"
echo Output folder: "%CD%\output_data"
echo.

"%PYTHON_EXE%" run_bot.py --dry-run %DRY_RUN%
set "BOT_EXIT_CODE=%ERRORLEVEL%"

echo.
if "%BOT_EXIT_CODE%"=="0" (
    echo Run completed successfully.
    echo Report: "%CD%\output_data\summary_report.xlsx"
    echo Logs  : "%CD%\logs\events.jsonl"
) else (
    echo Run failed with exit code %BOT_EXIT_CODE%.
)

echo.
pause
exit /b %BOT_EXIT_CODE%
