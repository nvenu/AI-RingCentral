# Quick Start Guide - Windows

## 🚀 5-Minute Setup

### Step 1: Install Python (5 minutes)
1. Download: https://www.python.org/downloads/
2. Run installer
3. ✅ **CHECK:** "Add Python to PATH"
4. Click "Install Now"

### Step 2: Install Packages (2 minutes)
1. Double-click: **`install_windows.bat`**
2. Wait for installation to complete
3. Edit `.env` file when it opens
4. Add your email password

### Step 3: Test (10 minutes)
1. Double-click: **`run_reports.bat`**
2. Wait for completion
3. Check your email

### Step 4: Schedule (5 minutes)
1. Open Task Scheduler
2. Create Basic Task
3. Daily at 4:00 PM
4. Run: `run_reports.bat`

---

## 📦 What to Install on Windows

### Required Software:
1. **Python 3.8+** (from python.org)
   - Make sure "Add Python to PATH" is checked!

### Required Python Packages:
Run `install_windows.bat` or manually:
```cmd
pip install ringcentral
pip install openpyxl
pip install python-dotenv
```

### Configuration:
1. Create `.env` file
2. Add: `EMAIL_PASSWORD=your_password`

---

## 📁 Files You Need

### Essential Files:
- ✅ `generate_and_send_reports.py` - Main script
- ✅ `improved_call_logs.py` - Call report generator
- ✅ `analyze_fax_senders.py` - Fax analysis
- ✅ `send_complete_reports.py` - Email sender
- ✅ `requirements.txt` - Package list
- ✅ `.env` - Email password (you create this)

### Helper Files:
- 📝 `run_reports.bat` - Easy runner
- 📝 `install_windows.bat` - Easy installer
- 📝 `WINDOWS_INSTALLATION_GUIDE.md` - Full guide

---

## ⚡ Quick Commands

### Install Everything:
```cmd
install_windows.bat
```

### Run Reports:
```cmd
run_reports.bat
```

### Manual Run:
```cmd
cd C:\path\to\RingCentral
python generate_and_send_reports.py
```

### Check Installation:
```cmd
python --version
pip list
```

---

## ✅ Checklist

- [ ] Python installed
- [ ] Python in PATH (type `python --version` in cmd)
- [ ] Packages installed (`install_windows.bat`)
- [ ] `.env` file created with EMAIL_PASSWORD
- [ ] Test run successful (`run_reports.bat`)
- [ ] Email received
- [ ] Task Scheduler configured

---

## 🎯 What Gets Installed

### Python Packages:
1. **ringcentral** - Connects to RingCentral API
2. **openpyxl** - Creates Excel reports
3. **python-dotenv** - Reads .env file
4. **requests** - HTTP requests (dependency)
5. **python-dateutil** - Date handling (dependency)

### Total Size: ~50 MB

---

## 📧 Email Configuration

Edit `send_complete_reports.py` to change recipients:

```python
receiver_emails = [
    "dogden@HomeCareForYou.com",
    "DrBrar@HomeCareForYou.com",
    "nvenu@solifetec.com"
]
```

---

## 🔧 Troubleshooting

### Python not found?
- Reinstall Python with "Add to PATH" checked
- Or add manually: Control Panel → System → Environment Variables

### Packages won't install?
- Run Command Prompt as Administrator
- Or try: `python -m pip install -r requirements.txt`

### Email not sending?
- Check `.env` file has EMAIL_PASSWORD
- No quotes, no spaces: `EMAIL_PASSWORD=yourpassword`

### Reports not generating?
- Check internet connection
- Verify RingCentral JWT token is valid
- Check Windows Firewall isn't blocking

---

## 📊 What You'll Get

### Daily at 4 PM:
1. **Email** sent to 3 recipients
2. **2 Excel files** attached:
   - Call productivity report
   - Fax sender analysis
3. **HTML report** in email body

### Reports Include:
- Total calls made/received
- Talk time by employee
- Fax activity by sender
- Top performers
- Complete call/fax logs

---

## 🎉 Done!

Once installed:
- Runs automatically every day at 4 PM
- No manual work needed
- Reports emailed automatically
- Files saved in `exports/` folder

**Need help?** See `WINDOWS_INSTALLATION_GUIDE.md` or `TROUBLESHOOTING.md`
