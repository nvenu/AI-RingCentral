#!/usr/bin/env python3
"""
Generate Reports for Specific Date
Allows you to generate reports for any specific date
"""

import subprocess
import sys
import os
from datetime import datetime

# ============================================
# CONFIGURE THE DATE HERE
# ============================================
REPORT_DATE = "2025-11-20"  # Format: YYYY-MM-DD

def update_date_in_scripts(target_date):
    """Temporarily update the date in scripts"""
    print(f"📅 Configuring scripts for date: {target_date}")
    
    # Parse the date
    date_obj = datetime.strptime(target_date, "%Y-%m-%d")
    date_from = date_obj.strftime("%Y-%m-%dT00:00:00.000Z")
    date_to = date_obj.strftime("%Y-%m-%dT23:59:59.999Z")
    
    return date_from, date_to, target_date

def run_script_with_date(script_name, description, date_from, date_to, date_str):
    """Run a Python script with specific date parameters"""
    print(f"\n{'='*60}")
    print(f"🚀 {description}")
    print('='*60)
    
    # Set environment variables for the date
    env = os.environ.copy()
    env['REPORT_DATE_FROM'] = date_from
    env['REPORT_DATE_TO'] = date_to
    env['REPORT_DATE_STR'] = date_str
    
    try:
        # Run with date parameters
        result = subprocess.run(
            [sys.executable, script_name, date_from, date_to, date_str],
            capture_output=False,
            text=True,
            check=True,
            env=env
        )
        print(f"✅ {description} completed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} failed with error code {e.returncode}")
        return False

def main():
    print("\n" + "="*60)
    print("📊 RINGCENTRAL REPORTS - SPECIFIC DATE GENERATOR")
    print("="*60)
    print(f"🕒 Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📅 Target Date: {REPORT_DATE}")
    
    # Get date parameters
    date_from, date_to, date_str = update_date_in_scripts(REPORT_DATE)
    
    print(f"\n📋 Date Range:")
    print(f"   From: {date_from}")
    print(f"   To:   {date_to}")
    
    # Note about successful faxes only
    print(f"\n✅ IMPORTANT: Only counting SUCCESSFUL faxes")
    print(f"   • Sent faxes with status: 'Sent'")
    print(f"   • Received faxes with status: 'Received'")
    print(f"   • Unsuccessful attempts (Busy, No Answer, Failed) are excluded")
    
    # Step 1: Generate call productivity report
    success1 = run_script_with_date(
        'improved_call_logs.py',
        f'Step 1/3: Generating Call Productivity Report for {REPORT_DATE}',
        date_from, date_to, date_str
    )
    
    if not success1:
        print("\n❌ Failed to generate call report. Stopping.")
        sys.exit(1)
    
    # Step 2: Generate fax analysis report
    success2 = run_script_with_date(
        'analyze_fax_senders.py',
        f'Step 2/3: Generating Fax Sender Analysis for {REPORT_DATE}',
        date_from, date_to, date_str
    )
    
    if not success2:
        print("\n❌ Failed to generate fax report. Stopping.")
        sys.exit(1)
    
    # Step 3: Send comprehensive email
    success3 = run_script_with_date(
        'send_complete_reports.py',
        f'Step 3/3: Sending Comprehensive Email Report for {REPORT_DATE}',
        date_from, date_to, date_str
    )
    
    if not success3:
        print("\n❌ Failed to send email.")
        sys.exit(1)
    
    # Summary
    print("\n" + "="*60)
    print("✅ ALL REPORTS GENERATED AND SENT SUCCESSFULLY!")
    print("="*60)
    print(f"\n📋 Generated Reports for: {REPORT_DATE}")
    print(f"   1. Call Productivity Report (Excel)")
    print(f"   2. Fax Sender Analysis (Excel) - Successful faxes only")
    print(f"   3. Comprehensive HTML Email")
    
    print(f"\n📧 Email Recipients:")
    print(f"   • Configured in .env file (RECIPIENT_EMAILS)")
    
    print(f"\n📁 Files saved in: exports/")
    print(f"   • {date_str}-nightangle-calls.xlsx")
    print(f"   • fax_analysis_{date_str}.xlsx")
    
    print(f"\n🕒 Completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60 + "\n")

if __name__ == "__main__":
    # Check if date was provided as command line argument
    if len(sys.argv) > 1:
        REPORT_DATE = sys.argv[1]
        print(f"📅 Using date from command line: {REPORT_DATE}")
    
    main()
