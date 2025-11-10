@echo off
REM RingCentral Reports - Automated Runner
REM This batch file runs the report generation script

echo ========================================
echo RingCentral Reports - Starting...
echo ========================================
echo.
echo Current Date/Time: %DATE% %TIME%
echo.

REM Change to the script directory
cd /d "%~dp0"
echo Working Directory: %CD%
echo.

REM Check if Python is available
echo Checking Python installation...
python --version
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Python is not found or not in PATH
    echo Please install Python or add it to your PATH
    echo.
    pause
    exit /b 1
)
echo.

REM Check if the Python script exists
if not exist "generate_and_send_reports.py" (
    echo ERROR: generate_and_send_reports.py not found!
    echo Current directory: %CD%
    echo.
    pause
    exit /b 1
)

REM Check if .env file exists
if not exist ".env" (
    echo WARNING: .env file not found!
    echo Make sure EMAIL_PASSWORD is configured
    echo.
)

echo ========================================
echo Running Report Generation...
echo ========================================
echo.

REM Run the Python script
python generate_and_send_reports.py

REM Check if the script ran successfully
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo SUCCESS: Reports completed successfully!
    echo ========================================
    echo.
    echo Check your email for the reports.
) else (
    echo.
    echo ========================================
    echo ERROR: Reports failed with code %ERRORLEVEL%
    echo ========================================
    echo.
    echo Please check the error messages above.
)

echo.
echo Finished at: %DATE% %TIME%
echo.

REM Keep window open to see results
echo Press any key to close this window...
pause > nul

exit /b %ERRORLEVEL%
