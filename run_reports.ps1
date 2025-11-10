# RingCentral Reports - PowerShell Runner
# This PowerShell script runs the report generation

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "RingCentral Reports - Starting..." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Current Date/Time: $(Get-Date)" -ForegroundColor Yellow
Write-Host ""

# Change to script directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptPath
Write-Host "Working Directory: $scriptPath" -ForegroundColor Yellow
Write-Host ""

# Check Python installation
Write-Host "Checking Python installation..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Python is not found or not in PATH" -ForegroundColor Red
    Write-Host "Please install Python or add it to your PATH" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host ""

# Check if Python script exists
if (-not (Test-Path "generate_and_send_reports.py")) {
    Write-Host "ERROR: generate_and_send_reports.py not found!" -ForegroundColor Red
    Write-Host "Current directory: $scriptPath" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if .env file exists
if (-not (Test-Path ".env")) {
    Write-Host "WARNING: .env file not found!" -ForegroundColor Yellow
    Write-Host "Make sure EMAIL_PASSWORD is configured" -ForegroundColor Yellow
    Write-Host ""
}

# Check if required Python packages are installed
Write-Host "Checking Python packages..." -ForegroundColor Yellow
$packages = @("ringcentral", "openpyxl")
foreach ($package in $packages) {
    $installed = python -c "import $package" 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "WARNING: Python package '$package' may not be installed" -ForegroundColor Yellow
    } else {
        Write-Host "  ✓ $package installed" -ForegroundColor Green
    }
}
Write-Host ""

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Running Report Generation..." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Run the Python script
python generate_and_send_reports.py

# Check exit code
if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "SUCCESS: Reports completed successfully!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Check your email for the reports." -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "ERROR: Reports failed with code $LASTEXITCODE" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please check the error messages above." -ForegroundColor Red
}

Write-Host ""
Write-Host "Finished at: $(Get-Date)" -ForegroundColor Yellow
Write-Host ""

# Keep window open
Write-Host "Press any key to close this window..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

exit $LASTEXITCODE
