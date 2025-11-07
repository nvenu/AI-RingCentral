@echo off
REM RingCentral Reports - Automated Runner
REM This batch file runs the report generation script

echo ========================================
echo RingCentral Reports - Starting...
echo ========================================
echo.

REM Change to the script directory
cd /d "%~dp0"

REM Run the Python script
python generate_and_send_reports.py

REM Check if the script ran successfully
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo Reports completed successfully!
    echo ========================================
) else (
    echo.
    echo ========================================
    echo ERROR: Reports failed with code %ERRORLEVEL%
    echo ========================================
)

REM Keep window open for 10 seconds to see results
timeout /t 10

exit
