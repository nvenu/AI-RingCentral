# Windows Task Scheduler Setup Guide

## 📋 Automated Daily Reports at 4 PM

This guide will help you set up the RingCentral reports to run automatically every day at 4 PM on Windows.

---

## 🚀 Quick Setup Steps

### Step 1: Verify the Batch File

1. Double-click `run_reports.bat` to test it manually
2. Make sure it runs successfully and generates reports
3. Check that the email is sent to all recipients

### Step 2: Open Windows Task Scheduler

**Method 1: Search**
- Press `Windows Key`
- Type "Task Scheduler"
- Click on "Task Scheduler" app

**Method 2: Run Command**
- Press `Windows Key + R`
- Type: `taskschd.msc`
- Press Enter

### Step 3: Create a New Task

1. In Task Scheduler, click **"Create Basic Task..."** in the right panel
   
2. **Name and Description:**
   - Name: `RingCentral Daily Reports`
   - Description: `Generates and emails RingCentral call and fax reports daily at 4 PM`
   - Click **Next**

3. **Trigger (When):**
   - Select: **Daily**
   - Click **Next**

4. **Daily Settings:**
   - Start date: Today's date
   - Start time: **4:00:00 PM** (16:00:00)
   - Recur every: **1 days**
   - Click **Next**

5. **Action:**
   - Select: **Start a program**
   - Click **Next**

6. **Start a Program:**
   - Program/script: Click **Browse** and select `run_reports.bat`
   - Or paste the full path, for example:
     ```
     C:\Users\YourUsername\Desktop\ringcentral\run_reports.bat
     ```
   - Start in (optional): Enter the folder path:
     ```
     C:\Users\YourUsername\Desktop\ringcentral
     ```
   - Click **Next**

7. **Summary:**
   - Review your settings
   - Check the box: **"Open the Properties dialog for this task when I click Finish"**
   - Click **Finish**

### Step 4: Configure Advanced Settings

When the Properties dialog opens:

1. **General Tab:**
   - ✅ Check: "Run whether user is logged on or not"
   - ✅ Check: "Run with highest privileges"
   - Configure for: **Windows 10** (or your Windows version)

2. **Triggers Tab:**
   - Double-click the trigger to edit
   - ✅ Check: "Enabled"
   - Advanced settings:
     - ✅ Check: "Stop task if it runs longer than: 2 hours"
   - Click **OK**

3. **Actions Tab:**
   - Verify your batch file path is correct

4. **Conditions Tab:**
   - ⬜ Uncheck: "Start the task only if the computer is on AC power"
   - ✅ Check: "Wake the computer to run this task" (if you want it to run even if PC is sleeping)

5. **Settings Tab:**
   - ✅ Check: "Allow task to be run on demand"
   - ✅ Check: "Run task as soon as possible after a scheduled start is missed"
   - ✅ Check: "If the task fails, restart every: 10 minutes"
   - Set: "Attempt to restart up to: 3 times"
   - If the running task does not end when requested: **Stop the existing instance**

6. Click **OK** to save

7. **Enter your Windows password** when prompted

---

## ✅ Testing the Scheduled Task

### Test 1: Run Manually
1. In Task Scheduler, find your task: `RingCentral Daily Reports`
2. Right-click → **Run**
3. Check if reports are generated and email is sent

### Test 2: Check Task History
1. Right-click your task → **Properties**
2. Go to **History** tab
3. Verify the task ran successfully (Result: 0x0)

### Test 3: Verify Output
1. Check the `exports` folder for new reports
2. Check your email inbox for the report email

---

## 🔧 Troubleshooting

### Issue: Task doesn't run at scheduled time

**Solution 1: Check Task Status**
- Open Task Scheduler
- Find your task
- Status should be "Ready" (not "Disabled")
- Last Run Result should be "The operation completed successfully (0x0)"

**Solution 2: Verify Python Path**
- Open Command Prompt
- Type: `python --version`
- If it doesn't work, you need to add Python to PATH
- Or modify `run_reports.bat` to use full Python path:
  ```batch
  "C:\Python39\python.exe" generate_and_send_reports.py
  ```

**Solution 3: Check Permissions**
- Right-click the task → Properties
- General tab → Make sure "Run with highest privileges" is checked
- Make sure the user account has permission to run the script

### Issue: Email not sending

**Solution:**
- Check that `.env` file has correct `EMAIL_PASSWORD`
- Test manually by running `run_reports.bat`
- Check Windows Firewall isn't blocking SMTP (port 587)

### Issue: Reports not generating

**Solution:**
- Check that all Python dependencies are installed:
  ```
  pip install -r requirements.txt
  ```
- Verify RingCentral JWT token is valid
- Check internet connection

---

## 📊 What Happens Daily at 4 PM

1. **4:00 PM** - Task Scheduler triggers the batch file
2. **4:00-4:05 PM** - Script generates call productivity report
3. **4:05-4:10 PM** - Script generates fax sender analysis
4. **4:10 PM** - Email sent to:
   - dogden@HomeCareForYou.com
   - DrBrar@HomeCareForYou.com
   - nvenu@solifetec.com

Total time: ~10 minutes

---

## 📝 Viewing Task Logs

To see when the task ran and if it was successful:

1. Open Task Scheduler
2. Find your task: `RingCentral Daily Reports`
3. Click on **History** tab (bottom panel)
4. Look for:
   - **Task Started** (Event ID: 100)
   - **Task Completed** (Event ID: 102)
   - **Action Completed** (Event ID: 201)

---

## 🔄 Modifying the Schedule

To change the time or frequency:

1. Open Task Scheduler
2. Find your task: `RingCentral Daily Reports`
3. Right-click → **Properties**
4. Go to **Triggers** tab
5. Double-click the trigger
6. Modify the time or recurrence
7. Click **OK** → **OK**

---

## 🛑 Disabling or Deleting the Task

**To Disable (pause):**
1. Right-click the task
2. Select **Disable**

**To Delete:**
1. Right-click the task
2. Select **Delete**
3. Confirm

---

## 💡 Alternative: Run at Different Times

If you want multiple reports per day:

1. Create multiple triggers in the same task:
   - Right-click task → Properties
   - Triggers tab → Click **New**
   - Add another time (e.g., 9 AM, 4 PM)

2. Or create separate tasks:
   - Morning Report: 9:00 AM
   - Afternoon Report: 4:00 PM

---

## 📞 Support

If you encounter issues:
1. Check the batch file runs manually first
2. Review Task Scheduler History logs
3. Check Windows Event Viewer for errors
4. Verify all file paths are correct

---

## ✅ Checklist

Before finalizing:
- [ ] Batch file runs successfully when double-clicked
- [ ] Python and all dependencies are installed
- [ ] `.env` file has correct EMAIL_PASSWORD
- [ ] Task is created in Task Scheduler
- [ ] Task is set to run at 4:00 PM daily
- [ ] Task is enabled (not disabled)
- [ ] Test run completed successfully
- [ ] Email received by all recipients

---

**Setup Complete! Your reports will now run automatically every day at 4 PM.**
