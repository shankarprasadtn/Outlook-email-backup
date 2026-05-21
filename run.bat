@echo off
title Outlook SharePoint Sync - Run
echo ===================================================
echo Outlook and SharePoint Sync Utility Launcher
echo ===================================================
echo.

:: Check for Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python 3 is not installed or not added to your PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo and ensure "Add Python to PATH" is checked during installation.
    echo.
    pause
    exit /b 1
)

:: Create Virtual Environment if it doesn't exist
if not exist .venv (
    echo Creating virtual environment (.venv)...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
)

:: Activate Virtual Environment
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo [ERROR] Failed to activate virtual environment.
    pause
    exit /b 1
)

:: Install Dependencies
echo Installing dependencies from requirements.txt...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

:: Run the Application
echo Starting Outlook SharePoint Sync GUI...
python app.py
if %errorlevel% neq 0 (
    echo [ERROR] Application exited with code %errorlevel%.
    pause
)

deactivate
