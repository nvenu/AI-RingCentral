#!/usr/bin/env python3
"""
Complete RingCentral Analytics Runner
Runs extension-based analytics and generates HTML dashboard in one command
"""

import subprocess
import sys
import os
import glob
from datetime import datetime

def run_command(command, description):
    """Run a command and handle errors"""
    print(f"🔄 {description}...")
    try:
        result = subprocess.run([sys.executable] + command.split()[1:], 
                              capture_output=True, text=True, check=True)
        print(f"✅ {description} completed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Error in {description}:")
        print(e.stderr)
        return False

def clean_old_files():
    """Clean up old CSV and HTML files from root directory"""
    print("🧹 Cleaning up old files...")
    
    # Files to clean up from root directory
    cleanup_patterns = [
        'extension_based_analytics_*.csv',
        'extension_analytics_dashboard_*.html',
        'simple_contact_analytics_*.csv',
        'call_analytics_dashboard_*.html',
        'contact_call_analytics_*.csv',
        'call_logs.csv'
    ]
    
    import glob
    cleaned_count = 0
    
    for pattern in cleanup_patterns:
        files = glob.glob(pattern)
        for file in files:
            try:
                os.remove(file)
                cleaned_count += 1
                print(f"   🗑️  Removed: {file}")
            except:
                pass
    
    if cleaned_count > 0:
        print(f"✅ Cleaned up {cleaned_count} old files")
    else:
        print("✅ No old files to clean up")

def main():
    print("🚀 Starting Complete RingCentral Call Analytics...")
    print("=" * 60)
    
    # Step 0: Clean up old files
    clean_old_files()
    
    # Step 1: Run extension-based analytics
    if not run_command("python extension_based_analytics.py", "Extension-based analytics"):
        print("💡 Make sure your RingCentral credentials are valid")
        return
    
    # Step 2: Generate HTML dashboard
    if not run_command("python extension_dashboard_generator.py", "HTML dashboard generation"):
        print("💡 Check if the CSV file was generated properly")
        return
    
    # Step 3: Find and display results
    print("\n🎉 Complete analytics finished successfully!")
    print("=" * 60)
    
    # Find generated files in exports folder
    exports_dir = "exports"
    if os.path.exists(exports_dir):
        csv_files = [f for f in os.listdir(exports_dir) if f.startswith('extension_based_analytics_') and f.endswith('.csv')]
        html_files = [f for f in os.listdir(exports_dir) if f.startswith('extension_analytics_dashboard_') and f.endswith('.html')]
        
        if csv_files:
            latest_csv = sorted(csv_files)[-1]
            print(f"📊 CSV Analytics: exports/{latest_csv}")
        
        if html_files:
            latest_html = sorted(html_files)[-1]
            html_path = os.path.join(exports_dir, latest_html)
            print(f"🌐 HTML Dashboard: {html_path}")
            
            # Try to open the dashboard
            try:
                import webbrowser
                webbrowser.open(f'file://{os.path.abspath(html_path)}')
                print("🌐 Opening dashboard in your browser...")
            except:
                print("💡 Manually open the HTML file to view the dashboard")
    
    print("\n📋 What you got:")
    print("• Extension-based call analytics (CSV)")
    print("• Interactive HTML dashboard with charts")
    print("• Internal vs external contact separation")
    print("• Searchable contact tables")
    print("• All files organized in 'exports' folder")

if __name__ == "__main__":
    main()