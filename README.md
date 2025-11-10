# RingCentral Reports - Automated Call & Fax Analytics

Automated system that generates daily call productivity and fax sender analysis reports from RingCentral, then emails them to management.

## 📊 What This Does

- Fetches all call and fax data from RingCentral API
- Generates detailed Excel reports showing:
  - Call productivity by employee (calls made/received, talk time, success rates)
  - Fax activity by sender (who sent each fax through the main fax line)
- Emails comprehensive reports to management daily
- Runs automatically at 4 PM every day (via Windows Task Scheduler)

## 🚀 Quick Start for Windows

### 1. Install Python
Download from: https://www.python.org/downloads/
- ✅ **CHECK:** "Add Python to PATH" during installation

### 2. Install Packages
Double-click: **`install_windows.bat`**

### 3. Configure Email
Edit `.env` file and add:
```
EMAIL_PASSWORD=your_email_password
```

### 4. Test
Double-click: **`run_reports.bat`**

### 5. Schedule
Follow: **`WINDOWS_SCHEDULER_SETUP.md`**

---

## 📁 Project Structure

### Main Scripts:
- **`generate_and_send_reports.py`** - Master script (runs everything)
- **`improved_call_logs.py`** - Generates call productivity report
- **`analyze_fax_senders.py`** - Analyzes fax senders
- **`send_complete_reports.py`** - Sends email with reports

### Windows Automation:
- **`run_reports.bat`** - Easy runner for Windows
- **`run_reports.ps1`** - PowerShell alternative
- **`install_windows.bat`** - Automated installer

### Documentation:
- **`QUICK_START_WINDOWS.md`** - 5-minute setup guide
- **`WINDOWS_INSTALLATION_GUIDE.md`** - Complete installation guide
- **`WINDOWS_SCHEDULER_SETUP.md`** - Task Scheduler setup
- **`TROUBLESHOOTING.md`** - Common issues and solutions

### Configuration:
- **`requirements.txt`** - Python package dependencies
- **`.env`** - Email password (you create this)
- **`.env.example`** - Template for .env file

### Output:
- **`exports/`** - Generated Excel reports saved here

---

## 📧 Email Recipients

Reports are sent to:
- dogden@HomeCareForYou.com
- DrBrar@HomeCareForYou.com
- nvenu@solifetec.com

To change recipients, edit `send_complete_reports.py` (line 265-270)

---

## 📦 Requirements

- Python 3.8 or higher
- Windows 10/11
- Internet connection
- RingCentral account with API access

---

## 🔧 Installation

See **`WINDOWS_INSTALLATION_GUIDE.md`** for complete instructions.

Quick install:
```cmd
# Run the installer
install_windows.bat

# Or manually
pip install -r requirements.txt
```

---

## 🎯 Usage

### Manual Run:
```cmd
run_reports.bat
```

### Automated (Daily at 4 PM):
Set up Windows Task Scheduler - see `WINDOWS_SCHEDULER_SETUP.md`

---

## 📊 Generated Reports

### 1. Call Productivity Report
- Employee name and extension
- Calls made/received
- Talk time and average duration
- Success rates
- Missed calls
- Fax sent/received counts

### 2. Fax Sender Analysis
- Who sent each fax (with extension)
- Fax counts by employee
- Complete fax log with timestamps
- External fax senders

---

## 🔐 Security

- Email password stored in `.env` file (not committed to Git)
- RingCentral JWT token embedded in scripts
- All API calls use HTTPS

---

## 🆘 Support

- **Quick issues:** See `TROUBLESHOOTING.md`
- **Installation help:** See `WINDOWS_INSTALLATION_GUIDE.md`
- **Scheduler help:** See `WINDOWS_SCHEDULER_SETUP.md`

---

## ✅ Features

- ✅ Automatic daily report generation
- ✅ Detailed call analytics by employee
- ✅ Fax sender tracking (solves "who sent this fax?" problem)
- ✅ Beautiful HTML email reports
- ✅ Excel attachments for detailed analysis
- ✅ Rate limit handling (no API errors)
- ✅ Error recovery and retry logic
- ✅ Windows Task Scheduler integration

---

## 📝 License

Internal use only - HomeCareForYou
