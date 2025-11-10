# Troubleshooting Guide - Windows Batch File Issues

## 🔍 Common Issues and Solutions

### Issue 1: "Python is not recognized as an internal or external command"

**Problem:** Python is not in your Windows PATH

**Solution A: Add Python to PATH**
1. Open Start Menu
2. Search for "Environment Variables"
3. Click "Edit the system environment variables"
4. Click "Environment Variables" button
5. Under "System variables", find "Path"
6. Click "Edit"
7. Click "New"
8. Add your Python installation path (e.g., `C:\Python39\` or `C:\Users\YourName\AppData\Local\Programs\Python\Python39\`)
9. Also add the Scripts folder (e.g., `C:\Python39\Scripts\`)
10. Click OK on all windows
11. **Restart Command Prompt** and try again

**Solution B: Use Full Python Path in Batch File**
1. Find where Python is installed:
   - Open Command Prompt
   - Type: `where python`
   - Copy the path (e.g., `C:\Python39\python.exe`)

2. Edit `run_reports.bat`
3. Replace `python` with the full path:
   ```batch
   "C:\Python39\python.exe" generate_and_send_reports.py
   ```

**Solution C: Use Python Launcher**
Replace `python` with `py`:
```batch
py generate_and_send_reports.py
```

---

### Issue 2: "generate_and_send_reports.py not found"

**Problem:** The batch file can't find the Python script

**Solution:**
1. Make sure you're running the batch file from the correct folder
2. Right-click `run_reports.bat` → Edit
3. Add this line after `cd /d "%~dp0"`:
   ```batch
   dir *.py
   ```
4. This will list all Python files to verify they're there

---

### Issue 3: "Module not found" errors

**Problem:** Required Python packages are not installed

**Solution:**
1. Open Command Prompt as Administrator
2. Navigate to your project folder:
   ```
   cd C:\path\to\your\ringcentral\folder
   ```
3. Install requirements:
   ```
   pip install -r requirements.txt
   ```
   Or install individually:
   ```
   pip install ringcentral openpyxl python-dotenv
   ```

---

### Issue 4: Email not sending

**Problem:** EMAIL_PASSWORD not set or incorrect

**Solution:**
1. Check if `.env` file exists in your project folder
2. Open `.env` file
3. Make sure it has:
   ```
   EMAIL_PASSWORD=your_actual_password_here
   ```
4. No quotes, no spaces around the `=`
5. Save the file

---

### Issue 5: Batch file closes immediately

**Problem:** Can't see error messages

**Solution:**
The new batch file now has `pause` at the end, but if it still closes:
1. Open Command Prompt manually
2. Navigate to your folder:
   ```
   cd C:\path\to\your\ringcentral\folder
   ```
3. Run the batch file:
   ```
   run_reports.bat
   ```
4. Now you can see all error messages

---

## 🔧 Alternative: Use PowerShell Script

If the batch file doesn't work, try the PowerShell script:

### Method 1: Run PowerShell Script Directly
1. Right-click `run_reports.ps1`
2. Select "Run with PowerShell"

### Method 2: If PowerShell Execution Policy Blocks It
1. Open PowerShell as Administrator
2. Run:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
3. Type `Y` and press Enter
4. Now try running `run_reports.ps1` again

### Method 3: Run from PowerShell Command Line
1. Open PowerShell
2. Navigate to your folder:
   ```powershell
   cd C:\path\to\your\ringcentral\folder
   ```
3. Run:
   ```powershell
   .\run_reports.ps1
   ```

---

## 🧪 Manual Testing Steps

### Step 1: Test Python
```cmd
python --version
```
Should show: `Python 3.x.x`

### Step 2: Test Python Script Directly
```cmd
cd C:\path\to\your\ringcentral\folder
python generate_and_send_reports.py
```

### Step 3: Check Dependencies
```cmd
pip list
```
Look for:
- ringcentral
- openpyxl
- python-dotenv

### Step 4: Test Each Script Individually
```cmd
python improved_call_logs.py
python analyze_fax_senders.py
python send_complete_reports.py
```

---

## 📋 Diagnostic Batch File

Create a file called `test_setup.bat` with this content:

```batch
@echo off
echo ========================================
echo System Diagnostic
echo ========================================
echo.

echo 1. Python Version:
python --version
echo.

echo 2. Python Location:
where python
echo.

echo 3. Current Directory:
cd
echo.

echo 4. Python Files in Directory:
dir *.py
echo.

echo 5. .env File Exists:
if exist ".env" (echo YES) else (echo NO)
echo.

echo 6. Installed Packages:
pip list | findstr "ringcentral openpyxl"
echo.

echo 7. Test Python Import:
python -c "import ringcentral; import openpyxl; print('All packages OK')"
echo.

pause
```

Run this to diagnose issues.

---

## 🎯 For Task Scheduler Issues

### Issue: Task shows "Running" but nothing happens

**Solution:**
1. In Task Scheduler, right-click your task → Properties
2. Actions tab → Edit the action
3. In "Start in (optional)", add the full folder path:
   ```
   C:\Users\YourName\Desktop\ringcentral
   ```
4. Make sure "Run with highest privileges" is checked

### Issue: Task runs but reports aren't generated

**Solution:**
1. Change the task action to run PowerShell instead:
   - Program/script: `powershell.exe`
   - Add arguments: `-ExecutionPolicy Bypass -File "C:\full\path\to\run_reports.ps1"`
   - Start in: `C:\full\path\to\ringcentral\folder`

### Issue: Need to see what's happening

**Solution - Add Logging:**
1. Edit `run_reports.bat`
2. Add at the end before `exit`:
   ```batch
   echo Log saved to: %CD%\report_log.txt
   python generate_and_send_reports.py > report_log.txt 2>&1
   ```
3. Check `report_log.txt` after the task runs

---

## 📞 Quick Help Checklist

Before asking for help, verify:
- [ ] Python is installed and in PATH
- [ ] All Python packages are installed (`pip install -r requirements.txt`)
- [ ] `.env` file exists with EMAIL_PASSWORD
- [ ] Batch file is in the same folder as Python scripts
- [ ] You can run `python generate_and_send_reports.py` manually
- [ ] Internet connection is working
- [ ] RingCentral JWT token is valid

---

## 💡 Best Practice: Use Full Paths

Edit `run_reports.bat` to use full paths:

```batch
@echo off
cd /d "C:\Users\YourName\Desktop\ringcentral"
"C:\Python39\python.exe" generate_and_send_reports.py
pause
```

Replace:
- `C:\Users\YourName\Desktop\ringcentral` with your actual folder path
- `C:\Python39\python.exe` with your actual Python path

---

## ✅ Success Indicators

When working correctly, you should see:
1. ✅ Python version displayed
2. ✅ "Step 1/3: Generating Call Productivity Report"
3. ✅ "Step 2/3: Generating Fax Sender Analysis"
4. ✅ "Step 3/3: Sending Comprehensive Email Report"
5. ✅ "Email sent successfully to:"
6. ✅ All three email addresses listed
7. ✅ "ALL REPORTS GENERATED AND SENT SUCCESSFULLY!"

If you see all of these, everything is working!
