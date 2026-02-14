@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM ===========================================
REM EXCEL_BOT SETUP, TEST, AND PACKAGE WORKFLOW
REM ===========================================

set "BOT_FOLDER=%~dp0"
if "%BOT_FOLDER:~-1%"=="\" set "BOT_FOLDER=%BOT_FOLDER:~0,-1%"
set "ENTRYPOINT=run_bot.py"
set "VENV_PY=venv\Scripts\python.exe"
set "VENV_ACTIVATE=venv\Scripts\activate.bat"
set "RUN_PY=%VENV_PY%"

echo.
echo [1/12] Using bot folder: "%BOT_FOLDER%"
cd /d "%BOT_FOLDER%" || (
    echo ERROR: Failed to enter bot folder.
    exit /b 1
)

if not exist "%ENTRYPOINT%" (
    echo ERROR: Entrypoint "%ENTRYPOINT%" not found.
    exit /b 1
)

echo.
echo [2/12] Checking Python and pip...
python --version || goto :fail
python -m pip --version || goto :fail

echo.
echo [3/12] Creating virtual environment (venv) if missing...
if not exist "%VENV_PY%" (
    python -m venv venv || goto :fail
)

echo.
echo [4/12] Activating virtual environment...
if exist "%VENV_ACTIVATE%" (
    call "%VENV_ACTIVATE%" || goto :fail
) else (
    echo WARNING: activate.bat not found. Continuing with direct venv python: "%VENV_PY%"
)

if not exist "%VENV_PY%" (
    echo ERROR: venv python not found at "%VENV_PY%".
    goto :fail
)

echo.
echo [5/12] Upgrading pip...
python -m pip --python "%VENV_PY%" install --upgrade pip >nul 2>&1

echo.
echo [6/12] Installing dependencies...
if exist requirements.txt (
    echo requirements.txt found. Installing from file...
    python -m pip --python "%VENV_PY%" install -r requirements.txt
    if errorlevel 1 (
        echo WARNING: venv dependency install failed. Falling back to system Python.
        set "RUN_PY=python"
    )
) else (
    echo requirements.txt not found. Installing baseline bot packages...
    python -m pip --python "%VENV_PY%" install pandas openpyxl
    if errorlevel 1 (
        echo WARNING: venv dependency install failed. Falling back to system Python.
        set "RUN_PY=python"
    )
)

echo.
echo [7/12] Freezing installed packages to requirements.lock.txt...
if /I "%RUN_PY%"=="python" (
    python -m pip freeze > requirements.lock.txt || goto :fail
) else (
    python -m pip --python "%VENV_PY%" freeze > requirements.lock.txt || goto :fail
)

echo.
echo [7b/12] Verifying runtime dependencies...
"%RUN_PY%" -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo ERROR: Required runtime packages ^(pandas/openpyxl^) are unavailable.
    goto :fail
)

echo.
echo [8/12] Verifying config files...
if not exist config.json (
    echo WARNING: config.json not found in project root.
)
if not exist users.json (
    echo WARNING: users.json not found in project root.
)
echo Config file check completed.

echo.
echo [9/12] Running dry-run validation...
"%RUN_PY%" "%ENTRYPOINT%" --dry-run true --headless
if errorlevel 1 (
    echo ERROR: Dry-run failed. Fix errors before packaging.
    goto :fail
)
echo Dry-run successful.

echo.
echo [10/12] Optional EXE build...
set "BUILD_EXE=n"
set /p BUILD_EXE=Do you want to rebuild excel_bot.exe? (y/n): 
if /i "%BUILD_EXE%"=="y" (
    if /I "%RUN_PY%"=="python" (
        python -m pip install pyinstaller || goto :fail
        python -m PyInstaller --noconfirm --clean --onefile --console --name excel_bot ^
          --add-data "config.json;." ^
          --add-data "users.json;." ^
          --add-data "excel_bot\config.json;excel_bot" ^
          --add-data "excel_bot\users.json;excel_bot" ^
          "%ENTRYPOINT%" || goto :fail
    ) else (
        python -m pip --python "%VENV_PY%" install pyinstaller || goto :fail
        "%VENV_PY%" -m PyInstaller --noconfirm --clean --onefile --console --name excel_bot ^
          --add-data "config.json;." ^
          --add-data "users.json;." ^
          --add-data "excel_bot\config.json;excel_bot" ^
          --add-data "excel_bot\users.json;excel_bot" ^
          "%ENTRYPOINT%" || goto :fail
    )

    if exist "dist\excel_bot.exe" (
        echo EXE created: dist\excel_bot.exe
    ) else (
        echo ERROR: EXE build reported success but dist\excel_bot.exe was not found.
        goto :fail
    )
) else (
    echo Skipping EXE build.
)

echo.
echo [11/12] Deactivating virtual environment...
if exist "%VENV_ACTIVATE%" call deactivate >nul 2>&1

echo.
echo [12/12] Completed.
echo Setup, test, and optional packaging finished successfully.
pause
exit /b 0

:fail
set "EXIT_CODE=%ERRORLEVEL%"
if "%EXIT_CODE%"=="" set "EXIT_CODE=1"
echo.
echo Workflow failed with exit code %EXIT_CODE%.
if exist "%VENV_ACTIVATE%" call deactivate >nul 2>&1
pause
exit /b %EXIT_CODE%
