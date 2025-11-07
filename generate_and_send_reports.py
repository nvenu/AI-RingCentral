#!/usr/bin/env python3
"""
Master Script - Generate and Send RingCentral Reports
Runs all report generation and sends comprehensive email
"""

import subprocess
import sys
from datetime import datetime

def run_script(script_name, description):
    """Run a Python script and handle errors"""
    print(f"\n{'='*60}")
    print(f"🚀 {description}")
    print('='*60)
    
    try:
        result = subprocess.run(
            [sys.executable, script_name],
            capture_output=False,
            text=True,
            check=True
        )
        print(f"✅ {description} completed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} failed with error code {e.returncode}")
        return False
    except FileNotFoundError:
        print(f"❌ Script {script_name} not found")
        return False

def main():
    print("\n" + "="*60)
    print("📊 RINGCENTRAL REPORTS - MASTER GENERATOR")
    print("="*60)
    print(f"🕒 Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Step 1: Generate call productivity report
    success1 = run_script(
        'improved_call_logs.py',
        'Step 1/3: Generating Call Productivity Report'
    )
    
    if not success1:
        print("\n❌ Failed to generate call report. Stopping.")
        sys.exit(1)
    
    # Step 2: Generate fax analysis report
    success2 = run_script(
        'analyze_fax_senders.py',
        'Step 2/3: Generating Fax Sender Analysis'
    )
    
    if not success2:
        print("\n❌ Failed to generate fax report. Stopping.")
        sys.exit(1)
    
    # Step 3: Send comprehensive email
    success3 = run_script(
        'send_complete_reports.py',
        'Step 3/3: Sending Comprehensive Email Report'
    )
    
    if not success3:
        print("\n❌ Failed to send email.")
        sys.exit(1)
    
    # Summary
    print("\n" + "="*60)
    print("✅ ALL REPORTS GENERATED AND SENT SUCCESSFULLY!")
    print("="*60)
    print(f"\n📋 Generated Reports:")
    print(f"   1. Call Productivity Report (Excel)")
    print(f"   2. Fax Sender Analysis (Excel)")
    print(f"   3. Comprehensive HTML Email")
    
    print(f"\n📧 Email Recipients:")
    print(f"   • dogden@HomeCareForYou.com")
    print(f"   • DrBrar@HomeCareForYou.com")
    print(f"   • nvenu@solifetec.com")
    
    print(f"\n🕒 Completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60 + "\n")

if __name__ == "__main__":
    main()
