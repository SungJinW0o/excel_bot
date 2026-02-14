@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "PYTHON_EXE=python"
set "PYTHONW_EXE=pythonw"
set "EXCEL_BOT_UI_SCALE=1.0"
if exist ".venv\Scripts\python.exe" set "PYTHON_EXE=.venv\Scripts\python.exe"
if exist ".venv\Scripts\pythonw.exe" set "PYTHONW_EXE=.venv\Scripts\pythonw.exe"

if /I "%PYTHON_EXE%"=="python" (
    where python >nul 2>&1
    if errorlevel 1 (
        echo ERROR: Python was not found.
        echo Install Python 3.9+ or reinstall using setup.exe from GitHub Releases.
        pause
        exit /b 1
    )
)

"%PYTHON_EXE%" -c "import PySide6, excel_bot.gui" >nul 2>&1
if errorlevel 1 (
    powershell -NoProfile -Command "[void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [Windows.Forms.MessageBox]::Show('GUI dependencies are missing or broken. Reinstall with setup.exe or run: ""%PYTHON_EXE%"" -m pip install --upgrade PySide6','Excel Bot GUI', 'OK', 'Error')"
    exit /b 1
)

start "" "%PYTHONW_EXE%" run_bot_gui.py
exit /b 0
