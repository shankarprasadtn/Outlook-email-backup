@echo off
title Outlook SharePoint Sync - Compile
echo ===================================================
echo Outlook and SharePoint Sync Utility Compiler
echo ===================================================
echo.

:: Check for Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python 3 is not installed or not added to your PATH.
    pause
    exit /b 1
)

:: Create Virtual Environment if it doesn't exist
if not exist .venv (
    echo Creating virtual environment (.venv)...
    python -m venv .venv
)

:: Activate Virtual Environment
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo [ERROR] Failed to activate virtual environment.
    pause
    exit /b 1
)

:: Install/Upgrade dependencies including pyinstaller
echo Verifying dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to verify dependencies.
    pause
    exit /b 1
)

:: Run PyInstaller to compile app.py to single exe
echo.
echo Compiling app.py into a single executable file...
echo This might take a minute...
echo.

:: We collect customtkinter files and build a single-file noconsole executable using the spec file
pyinstaller --clean OutlookSharePointSync.spec

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Compilation failed. See errors above.
    pause
    exit /b 1
)

echo.
echo ===================================================
echo [SUCCESS] Compilation completed successfully!
echo Executable is located in the "dist" directory:
echo dist\OutlookSharePointSync.exe
echo ===================================================
echo.
pause
deactivate
