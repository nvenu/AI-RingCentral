# RingCentral Reports - Automated Call & Fax Analytics

Automated system that generates daily call productivity and fax sender analysis reports from RingCentral, then emails them to management.

## 📊 What This Does

- Fetches all call and fax data from RingCentral API
- Generates detailed Excel reports showing:
  - Call productivity by employee (calls made/received, talk time, success rates)
  - Fax activity by sender (who sent each fax through the main fax line)
- Emails comprehensive reports to management daily
- Runs automatically at 4:00 PM every day

## 🚀 Quick Start - AWS Deployment

**Target Server:** AWS Ubuntu (34.218.221.142)

### One-Command Deployment

```bash
./deploy_to_aws.sh /path/to/your-aws-key.pem
```

See **`START_HERE_AWS.md`** for step-by-step instructions.

### What You Need

1. **AWS SSH Key (.pem file)** - Your private key for server access
2. **Email Password** - For nvenu@solifetec.com

### Quick Steps

1. Run deployment script (uploads files and installs everything)
2. Connect to server and configure email password
3. Test the script
4. Set up cron for 4:00 PM daily automation

**Total time:** ~15 minutes

---

## 📁 Project Structure

### Main Scripts:
- **`generate_and_send_reports.py`** - Master script (runs everything)
- **`improved_call_logs.py`** - Generates call productivity report
- **`analyze_fax_senders.py`** - Analyzes fax senders
- **`send_complete_reports.py`** - Sends email with reports
- **`generate_specific_date_report.py`** - Generate reports for specific date

### Deployment Scripts:
- **`deploy_to_aws.sh`** - One-command AWS deployment
- **`deploy.sh`** - General Linux server deployment
- **`test_deployment.sh`** - Verify deployment

### Documentation:
- **`START_HERE_AWS.md`** - Quick AWS deployment guide
- **`AWS_DEPLOYMENT_INSTRUCTIONS.md`** - Complete AWS deployment instructions
- **`EMAIL_CONFIGURATION.md`** - Email setup instructions
- **`FINAL_EMAIL_FORMAT.md`** - Email template documentation
- **`SUCCESSFUL_FAX_FILTERING.md`** - Fax filtering documentation

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

See **`START_HERE_AWS.md`** for AWS deployment.

For other Linux servers:
```bash
./deploy.sh
```

---

## 🎯 Usage

### Automated (Daily at 4:00 PM on AWS):
Set up via cron after deployment - see `START_HERE_AWS.md`

### Manual Run:
```bash
# On server
source venv/bin/activate
python generate_and_send_reports.py
```

### Generate Report for Specific Date:
```bash
python generate_specific_date_report.py
```

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

- **AWS Deployment:** See `START_HERE_AWS.md` or `AWS_DEPLOYMENT_INSTRUCTIONS.md`
- **Email Issues:** See `EMAIL_CONFIGURATION.md`
- **Test Deployment:** Run `./test_deployment.sh` on server

---

## ✅ Features

- ✅ Automatic daily report generation at 4:00 PM
- ✅ Detailed call analytics by employee
- ✅ Fax sender tracking (solves "who sent this fax?" problem)
- ✅ Beautiful HTML email reports
- ✅ Excel attachments for detailed analysis
- ✅ Rate limit handling (no API errors)
- ✅ Error recovery and retry logic
- ✅ AWS server deployment with cron automation

---

## 📝 License

Internal use only - HomeCareForYou
