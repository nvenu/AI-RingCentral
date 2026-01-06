# RingCentral Reports - Automated Call & Fax Analytics

Automated system that generates daily call productivity and fax sender analysis reports from RingCentral, then emails them to management.

## 📊 What This Does

- Fetches all call and fax data from RingCentral API
- Generates detailed Excel reports showing:
  - Call productivity by employee (calls made/received, talk time, success rates)
  - Fax activity by sender (who sent each fax through the main fax line)
- Emails comprehensive reports to management daily
- Runs automatically at 4:00 PM IST via GitHub Actions

## 🚀 Setup (GitHub Actions)

### 1. Add Repository Secrets

Go to your GitHub repository → Settings → Secrets and variables → Actions → New repository secret

Add these secrets:

| Secret Name | Value |
|-------------|-------|
| `EMAIL_PASSWORD` | Your email password for nvenu@solifetec.com |
| `SENDER_EMAIL` | nvenu@solifetec.com |
| `RECIPIENT_EMAILS` | dogden@HomeCareForYou.com,DrBrar@HomeCareForYou.com,nvenu@solifetec.com |

### 2. Push to GitHub

```bash
git add .
git commit -m "Add GitHub Actions workflow"
git push
```

### 3. Done!

Reports will automatically run daily at 4:00 PM PST.

## 🎯 Usage

### Automatic (Daily at 4:00 PM IST)
GitHub Actions runs automatically - no action needed.

### Manual Trigger
1. Go to repository → Actions tab
2. Select "Daily RingCentral Reports"
3. Click "Run workflow"

### Local Testing
```bash
pip install -r requirements.txt
cp .env.example .env
# Edit .env and add EMAIL_PASSWORD
python generate_and_send_reports.py
```

## 📧 Email Recipients

Reports are sent to:
- dogden@HomeCareForYou.com
- DrBrar@HomeCareForYou.com
- nvenu@solifetec.com

To change recipients, update the `RECIPIENT_EMAILS` secret in GitHub.

## ✅ Features

- ✅ Automatic daily report generation at 4:00 PM IST
- ✅ GitHub Actions - no server needed
- ✅ Detailed call analytics by employee
- ✅ Fax sender tracking
- ✅ Beautiful HTML email reports
- ✅ Excel attachments
- ✅ Reports saved as artifacts for 30 days

## 🔧 Changing the Schedule

Edit `.github/workflows/daily-reports.yml`:

```yaml
schedule:
  - cron: '30 10 * * *'  # 4:00 PM IST = 10:30 UTC
```

## 📝 License

Internal use only - HomeCareForYou
