# Windows Installation Guide

## 📋 Complete Setup Instructions for Windows

Follow these steps to set up the RingCentral Reports system on Windows.

---

## Step 1: Install Python

### Download Python
1. Go to: https://www.python.org/downloads/
2. Click **"Download Python 3.11.x"** (or latest version)
3. Run the installer

### Important During Installation:
✅ **CHECK THE BOX:** "Add Python to PATH" (VERY IMPORTANT!)
✅ Click "Install Now"

### Verify Installation:
1. Open **Command Prompt** (Press `Windows Key`, type `cmd`, press Enter)
2. Type: `python --version`
3. Should show: `Python 3.11.x` or similar
4. Type: `pip --version`
5. Should show: `pip 23.x.x` or similar

**If Python is not recognized:**
- You need to add Python to PATH manually (see TROUBLESHOOTING.md)
- Or reinstall Python and make sure to check "Add Python to PATH"

---

## Step 2: Copy Project Files to Windows

### Option A: Using USB Drive or Network Share
1. Copy the entire project folder to your Windows computer
2. Recommended location: `C:\RingCentral\` or `Desktop\RingCentral\`

### Option B: Using Git (if you have it)
```cmd
cd C:\
git clone <your-repository-url> RingCentral
cd RingCentral
```

### Files You Should Have:
```
RingCentral/
├── generate_and_send_reports.py
├── improved_call_logs.py
├── analyze_fax_senders.py
├── send_complete_reports.py
├── run_reports.bat
├── run_reports.ps1
├── requirements.txt
├── .env.example
├── README.md
└── exports/ (folder will be created automatically)
```

---

## Step 3: Install Required Python Packages

1. Open **Command Prompt as Administrator**
   - Press `Windows Key`
   - Type `cmd`
   - Right-click "Command Prompt"
   - Select "Run as administrator"

2. Navigate to your project folder:
   ```cmd
   cd C:\RingCentral
   ```
   (Replace with your actual path)

3. Install all required packages:
   ```cmd
   pip install -r requirements.txt
   ```

### If requirements.txt is missing, install manually:
```cmd
pip install ringcentral
pip install openpyxl
pip install python-dotenv
```

### Verify Installation:
```cmd
pip list
```

You should see:
- ringcentral
- openpyxl
- python-dotenv
- (and their dependencies)

---

## Step 4: Configure Email Settings

1. Copy `.env.example` to `.env`:
   ```cmd
   copy .env.example .env
   ```

2. Open `.env` file with Notepad:
   ```cmd
   notepad .env
   ```

3. Add your email password:
   ```
   EMAIL_PASSWORD=your_actual_password_here
   ```
   
   **Important:**
   - No quotes around the password
   - No spaces around the `=`
   - Use your Office365/Outlook email password
   - If you have 2FA enabled, you may need an app password

4. Save and close Notepad

---

## Step 5: Test the Setup

### Test 1: Run the Batch File
1. Navigate to your project folder in File Explorer
2. Double-click `run_reports.bat`
3. A command window should open and show progress
4. Wait for it to complete (takes 5-10 minutes)
5. Check if you received the email

### Test 2: Run from Command Prompt (Better for seeing errors)
```cmd
cd C:\RingCentral
python generate_and_send_reports.py
```

### What You Should See:
```
============================================================
📊 RINGCENTRAL REPORTS - MASTER GENERATOR
============================================================
🕒 Started at: 2025-11-07 16:00:00

============================================================
🚀 Step 1/3: Generating Call Productivity Report
============================================================
📡 Connecting to RingCentral...
✅ Authentication successful
...
✅ ALL REPORTS GENERATED AND SENT SUCCESSFULLY!
```

---

## Step 6: Set Up Automatic Daily Execution

### Option A: Windows Task Scheduler (Recommended)

Follow the detailed guide in `WINDOWS_SCHEDULER_SETUP.md`

**Quick Steps:**
1. Open Task Scheduler
2. Create Basic Task
3. Name: "RingCentral Daily Reports"
4. Trigger: Daily at 4:00 PM
5. Action: Start a program
6. Program: Browse to `run_reports.bat`
7. Finish and test

### Option B: PowerShell Scheduled Job

```powershell
# Open PowerShell as Administrator
$trigger = New-JobTrigger -Daily -At "4:00 PM"
$action = New-ScheduledTaskAction -Execute "C:\RingCentral\run_reports.bat"
Register-ScheduledTask -TaskName "RingCentral Reports" -Trigger $trigger -Action $action
```

---

## 📦 Complete Package List

Here's what gets installed with `pip install -r requirements.txt`:

### Main Packages:
- **ringcentral** - RingCentral API SDK
- **openpyxl** - Excel file generation
- **python-dotenv** - Environment variable management

### Dependencies (installed automatically):
- requests
- urllib3
- certifi
- et-xmlfile
- pyjwt
- pubnub
- pycryptodomex

---

## 🔧 System Requirements

### Minimum Requirements:
- **OS:** Windows 10 or Windows 11
- **Python:** 3.8 or higher (3.11 recommended)
- **RAM:** 2 GB minimum
- **Disk Space:** 500 MB free space
- **Internet:** Stable connection required

### Recommended:
- **OS:** Windows 11
- **Python:** 3.11 or higher
- **RAM:** 4 GB or more
- **Disk Space:** 1 GB free space

---

## 🌐 Network Requirements

The script needs to connect to:
- **RingCentral API:** platform.ringcentral.com (HTTPS, port 443)
- **Email Server:** smtp.office365.com (SMTP, port 587)

**Firewall Settings:**
- Allow outbound connections on ports 443 and 587
- If behind corporate firewall, may need IT approval

---

## 📝 Configuration Files

### .env File
```
EMAIL_PASSWORD=your_password_here
```

### Email Recipients (in send_complete_reports.py)
```python
receiver_emails = [
    "dogden@HomeCareForYou.com",
    "DrBrar@HomeCareForYou.com",
    "nvenu@solifetec.com"
]
```

---

## ✅ Installation Checklist

Before running the reports, verify:

- [ ] Python 3.8+ installed
- [ ] Python added to PATH
- [ ] All project files copied to Windows
- [ ] `pip install -r requirements.txt` completed successfully
- [ ] `.env` file created with EMAIL_PASSWORD
- [ ] Test run completed successfully (`run_reports.bat`)
- [ ] Email received by all recipients
- [ ] Task Scheduler configured (if automating)

---

## 🆘 Common Installation Issues

### Issue: "pip is not recognized"

**Solution:**
```cmd
python -m pip install -r requirements.txt
```

### Issue: "Permission denied" when installing packages

**Solution:**
Run Command Prompt as Administrator

### Issue: SSL Certificate errors

**Solution:**
```cmd
pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org -r requirements.txt
```

### Issue: Old Python version

**Solution:**
1. Uninstall old Python from Control Panel
2. Download and install latest Python from python.org
3. Make sure to check "Add Python to PATH"

---

## 🔄 Updating the Scripts

If you need to update the scripts later:

1. Copy new script files to your Windows folder
2. Overwrite existing files
3. No need to reinstall Python packages unless requirements.txt changed
4. Test with `run_reports.bat`

---

## 📞 Getting Help

If you encounter issues:

1. Check `TROUBLESHOOTING.md` for solutions
2. Run the diagnostic batch file (see TROUBLESHOOTING.md)
3. Check Windows Event Viewer for Task Scheduler errors
4. Verify internet connection
5. Check RingCentral JWT token is still valid

---

## 🎯 Quick Start Commands

```cmd
# Navigate to project folder
cd C:\RingCentral

# Install packages
pip install -r requirements.txt

# Create .env file
copy .env.example .env
notepad .env

# Test run
python generate_and_send_reports.py

# Or use batch file
run_reports.bat
```

---

## 📊 What Happens When It Runs

1. **Connects to RingCentral** - Authenticates with JWT token
2. **Fetches call logs** - Gets all calls for previous day
3. **Fetches fax logs** - Gets all faxes for previous day
4. **Generates Excel reports** - Creates two detailed reports
5. **Sends email** - Emails reports to 3 recipients
6. **Total time:** 5-10 minutes

---

## 🔐 Security Notes

- Keep your `.env` file secure (contains email password)
- Don't share your JWT token
- `.env` file should not be committed to Git (already in .gitignore)
- Email password should be an app-specific password if using 2FA

---

## ✅ Installation Complete!

Once all steps are completed:
- Reports will run automatically at 4 PM daily
- You'll receive email with Excel reports
- Reports are also saved in `exports/` folder
- No manual intervention needed

**Next Steps:**
1. Monitor first few automated runs
2. Check email delivery
3. Verify report accuracy
4. Adjust schedule if needed (see WINDOWS_SCHEDULER_SETUP.md)
