# Email Configuration Guide

## 📧 Email Behavior

The system sends **ONE comprehensive email** per day with both reports attached.

### Email Details:
- **Subject:** `📊 RingCentral Complete Activity Report - YYYY-MM-DD`
- **Attachments:** 
  1. Call Productivity Report (Excel)
  2. Fax Sender Analysis (Excel)
- **Body:** Beautiful HTML report with summary statistics

---

## 🔧 How to Configure Recipients

### Location: `.env` file

Edit the `.env` file in your project folder:

```properties
# Recipient emails (comma-separated, no spaces after commas)
RECIPIENT_EMAILS=email1@domain.com,email2@domain.com,email3@domain.com
```

### Examples:

**Send to 3 people:**
```
RECIPIENT_EMAILS=dogden@HomeCareForYou.com,DrBrar@HomeCareForYou.com,nvenu@solifetec.com
```

**Send to 2 people (remove Dr. Brar):**
```
RECIPIENT_EMAILS=dogden@HomeCareForYou.com,nvenu@solifetec.com
```

**Send to 4 people (add Della):**
```
RECIPIENT_EMAILS=dogden@HomeCareForYou.com,DrBrar@HomeCareForYou.com,della@HomeCareForYou.com,nvenu@solifetec.com
```

**Send to only 1 person:**
```
RECIPIENT_EMAILS=nvenu@solifetec.com
```

---

## ⚠️ Important Rules:

1. **No spaces** after commas
2. **No quotes** around email addresses
3. Use **comma** to separate multiple emails
4. Changes take effect immediately (no restart needed)

---

## 📝 Complete `.env` Configuration:

```properties
# Email password for sending reports
EMAIL_PASSWORD=your_actual_password

# Sender email address
SENDER_EMAIL=nvenu@solifetec.com

# Recipient emails (who receives the reports)
RECIPIENT_EMAILS=dogden@HomeCareForYou.com,DrBrar@HomeCareForYou.com,nvenu@solifetec.com
```

---

## 🔍 What Was Fixed:

### Before (Problem):
- ❌ 2 emails sent per day:
  1. "RingCentral Call Logs - 2025-11-09"
  2. "RingCentral Complete Activity Report - 2025-11-09"

### After (Fixed):
- ✅ 1 email sent per day:
  - "📊 RingCentral Complete Activity Report - 2025-11-09"
  - Contains both reports as attachments
  - Includes comprehensive HTML summary

---

## 📅 When Emails Are Sent:

- **Time:** 4:00 PM daily (via Windows Task Scheduler)
- **Data:** Previous day's call and fax records
- **Example:** Email sent on Nov 10 at 4 PM contains Nov 9 data

---

## 🛠️ How to Edit Recipients:

### On Windows:
1. Navigate to your RingCentral project folder
2. Right-click `.env` → Open with Notepad
3. Edit the `RECIPIENT_EMAILS` line
4. Save and close

### On Mac:
```bash
nano .env
# or
open .env
```

---

## ✅ Testing:

To test email configuration:
```cmd
# Windows
run_reports.bat

# Or manually
python generate_and_send_reports.py
```

Check that:
- ✅ Only 1 email is sent
- ✅ Subject is "RingCentral Complete Activity Report"
- ✅ All configured recipients receive it
- ✅ Both Excel files are attached

---

## 📞 Support:

If you need to change recipients, just edit the `.env` file - no code changes needed!
