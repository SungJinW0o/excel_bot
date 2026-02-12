@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "PYTHON_EXE=python"
if exist ".venv\Scripts\python.exe" set "PYTHON_EXE=.venv\Scripts\python.exe"

"%PYTHON_EXE%" run_bot_gui.py
exit /b %ERRORLEVEL%
