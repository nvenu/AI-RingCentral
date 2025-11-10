@echo off
REM RingCentral Reports - Windows Installation Script
REM This script installs all required Python packages

echo ========================================
echo RingCentral Reports - Installation
echo ========================================
echo.

REM Check if Python is installed
echo Checking Python installation...
python --version
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)
echo.

REM Check pip
echo Checking pip...
pip --version
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: pip is not available
    echo.
    pause
    exit /b 1
)
echo.

REM Upgrade pip
echo Upgrading pip to latest version...
python -m pip install --upgrade pip
echo.

REM Install required packages
echo ========================================
echo Installing Required Packages
echo ========================================
echo.
echo This may take a few minutes...
echo.

pip install -r requirements.txt

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo SUCCESS: All packages installed!
    echo ========================================
    echo.
    
    REM Verify installations
    echo Verifying installations...
    python -c "import ringcentral; print('✓ ringcentral installed')"
    python -c "import openpyxl; print('✓ openpyxl installed')"
    python -c "import dotenv; print('✓ python-dotenv installed')"
    echo.
    
    REM Check for .env file
    if not exist ".env" (
        echo ========================================
        echo NEXT STEP: Configure Email Settings
        echo ========================================
        echo.
        echo Creating .env file from template...
        copy .env.example .env
        echo.
        echo Please edit .env file and add your EMAIL_PASSWORD
        echo.
        echo Opening .env file in Notepad...
        timeout /t 2 >nul
        notepad .env
    ) else (
        echo .env file already exists
    )
    
    echo.
    echo ========================================
    echo Installation Complete!
    echo ========================================
    echo.
    echo Next steps:
    echo 1. Make sure .env file has your EMAIL_PASSWORD
    echo 2. Test by running: run_reports.bat
    echo 3. Set up Task Scheduler for daily automation
    echo.
    echo See WINDOWS_INSTALLATION_GUIDE.md for details
    echo.
) else (
    echo.
    echo ========================================
    echo ERROR: Installation failed
    echo ========================================
    echo.
    echo Please check the error messages above
    echo.
    echo Common solutions:
    echo 1. Run this script as Administrator
    echo 2. Check your internet connection
    echo 3. Try: pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org -r requirements.txt
    echo.
)

pause
