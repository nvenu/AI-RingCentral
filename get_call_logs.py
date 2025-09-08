#!/usr/bin/env python3
"""
Clean RingCentral Call Logs Fetcher
Simple script to get call logs and save to CSV
"""

import csv
import sys
import os
import json
import smtplib
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from collections import defaultdict

try:
    from ringcentral import SDK
except ImportError:
    print("❌ Error: RingCentral SDK not installed")
    print("Run: pip install ringcentral")
    sys.exit(1)

# Load environment variables from .env file if it exists
def load_env():
    """Load environment variables from .env file"""
    env_file = '.env'
    if os.path.exists(env_file):
        with open(env_file, 'r') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    os.environ[key.strip()] = value.strip()

# Load environment variables
load_env()

# Your RingCentral credentials
CLIENT_ID = "c97CkXtDJ4wcupAJ8ieF67"
CLIENT_SECRET = "5r8CRbMwHUjarHHjytVUSE7HawPIWiW3zeDyMhntFmCs"
SERVER_URL = "https://platform.ringcentral.com"
JWT_TOKEN = "eyJraWQiOiI4NzYyZjU5OGQwNTk0NGRiODZiZjVjYTk3ODA0NzYwOCIsInR5cCI6IkpXVCIsImFsZyI6IlJTMjU2In0.eyJhdWQiOiJodHRwczovL3BsYXRmb3JtLnJpbmdjZW50cmFsLmNvbS9yZXN0YXBpL29hdXRoL3Rva2VuIiwic3ViIjoiNjM2NzI0NzAwMzEiLCJpc3MiOiJodHRwczovL3BsYXRmb3JtLnJpbmdjZW50cmFsLmNvbSIsImV4cCI6MTgxOTg0MzE5OSwiaWF0IjoxNzU2ODk0MjU0LCJqdGkiOiJGMzhkNTd5dVJIcWZFVU5uX0t0M2NnIn0.WGDnyfTBEJ05PkwpHNCRYj-0ZkbLOkS5lvgzdPeiFCACSNSUAhzPKuYYZ_zUAuqBUDatw59soFVGeid6pRZDp8etXADwvHSwe7wai_w0zvTA1i_rGJGP3pwXEmdagMMudIQx-aWCnjazbjRWNXb_fWi5vpH4sf1jxUJQAERxF4iSmr0S__imKaV0D_26-mdlw021lb1Px3NiSh0YXo2ogHSW5jQHZ-HILYi-HaCDskTDr_34ccbe1tOb9bgnixmz6SflGMjupsfiBTzxvw1Ii-quVxyzWeiVYaiIG6vtJGKsrN8ocdXWtGzNGpT57MSqNnLMUBEcBqY4amvn8UXb6g"

def safe_get_attr(obj, attr, default=''):
    """Safely get attribute from JsonObject or dict"""
    try:
        if hasattr(obj, attr):
            return getattr(obj, attr, default)
        elif hasattr(obj, 'get'):
            return obj.get(attr, default)
        elif hasattr(obj, '__getitem__'):
            return obj[attr] if attr in obj else default
        else:
            return default
    except:
        return default

def load_cached_contacts():
    """Load contacts from local cache files"""
    extensions = {}
    contacts = {}
    
    # Load extensions cache
    ext_file = 'contacts/extensions.json'
    if os.path.exists(ext_file):
        try:
            with open(ext_file, 'r') as f:
                extensions = json.load(f)
            print(f"📋 Loaded {len(extensions)} cached extensions")
        except:
            print("⚠️  Could not load extensions cache")
    
    # Load contacts cache
    contacts_file = 'contacts/contacts.json'
    if os.path.exists(contacts_file):
        try:
            with open(contacts_file, 'r') as f:
                contacts = json.load(f)
            print(f"📋 Loaded {len(contacts)} cached contacts")
        except:
            print("⚠️  Could not load contacts cache")
    
    return extensions, contacts

def save_cached_contacts(extensions, contacts):
    """Save contacts to local cache files"""
    os.makedirs('contacts', exist_ok=True)
    
    # Save extensions
    try:
        with open('contacts/extensions.json', 'w') as f:
            json.dump(extensions, f, indent=2)
        print(f"💾 Saved {len(extensions)} extensions to cache")
    except Exception as e:
        print(f"⚠️  Could not save extensions cache: {e}")
    
    # Save contacts
    try:
        with open('contacts/contacts.json', 'w') as f:
            json.dump(contacts, f, indent=2)
        print(f"💾 Saved {len(contacts)} contacts to cache")
    except Exception as e:
        print(f"⚠️  Could not save contacts cache: {e}")

def fetch_and_cache_extensions(platform):
    """Fetch all extensions and cache them locally"""
    print("🔄 Fetching extensions from RingCentral...")
    extensions = {}
    
    try:
        page = 1
        per_page = 1000
        
        while True:
            response = platform.get("/restapi/v1.0/account/~/extension", {
                "perPage": per_page,
                "page": page
            })
            
            response_data = response.json()
            ext_records = safe_get_attr(response_data, 'records', [])
            
            if not ext_records:
                break
            
            for ext in ext_records:
                ext_id = str(safe_get_attr(ext, 'id', ''))
                if ext_id:
                    # Get contact info
                    contact = safe_get_attr(ext, 'contact', {})
                    first_name = safe_get_attr(contact, 'firstName', '')
                    last_name = safe_get_attr(contact, 'lastName', '')
                    name = (first_name + ' ' + last_name).strip()
                    
                    if not name:
                        name = safe_get_attr(ext, 'name', '')
                    if not name:
                        name = safe_get_attr(contact, 'displayName', '')
                    
                    extensions[ext_id] = {
                        'name': name if name else f"Extension {ext_id}",
                        'extensionNumber': safe_get_attr(ext, 'extensionNumber', ''),
                        'type': safe_get_attr(ext, 'type', ''),
                        'status': safe_get_attr(ext, 'status', '')
                    }
            
            # Check pagination
            paging = safe_get_attr(response_data, 'paging', {})
            if len(ext_records) < per_page or not safe_get_attr(paging, 'hasNextPage', False):
                break
            
            page += 1
        
        print(f"✅ Fetched {len(extensions)} extensions")
        return extensions
        
    except Exception as e:
        print(f"⚠️  Could not fetch extensions: {e}")
        return {}

def fetch_and_cache_contacts(platform):
    """Fetch all contacts and cache them locally"""
    print("🔄 Fetching contacts from RingCentral...")
    contacts = {}
    
    try:
        page = 1
        per_page = 1000
        
        while True:
            response = platform.get("/restapi/v1.0/account/~/directory/contacts", {
                "perPage": per_page,
                "page": page
            })
            
            response_data = response.json()
            contact_records = safe_get_attr(response_data, 'records', [])
            
            if not contact_records:
                break
            
            for contact in contact_records:
                # Get name
                first_name = safe_get_attr(contact, 'firstName', '')
                last_name = safe_get_attr(contact, 'lastName', '')
                name = (first_name + ' ' + last_name).strip()
                
                if not name:
                    name = safe_get_attr(contact, 'name', '')
                if not name:
                    name = safe_get_attr(contact, 'displayName', '')
                
                # Get phone numbers
                phone_numbers = safe_get_attr(contact, 'phoneNumbers', [])
                for phone_entry in phone_numbers:
                    phone_number = safe_get_attr(phone_entry, 'phoneNumber', '')
                    if phone_number and name:
                        contacts[phone_number] = name
            
            # Check pagination
            paging = safe_get_attr(response_data, 'paging', {})
            if len(contact_records) < per_page or not safe_get_attr(paging, 'hasNextPage', False):
                break
            
            page += 1
        
        print(f"✅ Fetched {len(contacts)} contact phone numbers")
        return contacts
        
    except Exception as e:
        print(f"⚠️  Could not fetch contacts: {e}")
        return {}

def get_contact_name(phone_number, extension_id, extensions_cache, contacts_cache):
    """Get contact name from cached data"""
    # First try extension lookup
    if extension_id and str(extension_id) in extensions_cache:
        return extensions_cache[str(extension_id)]['name']
    
    # Then try phone number lookup
    if phone_number and phone_number in contacts_cache:
        return contacts_cache[phone_number]
    
    # Return phone number if no name found
    return phone_number or f"Ext {extension_id}" if extension_id else "Unknown"

def send_email_with_attachment(csv_filename, date_str, total_records, total_users):
    """Send email with CSV attachment"""
    print("📧 Sending email with CSV attachment...")
    
    # Email configuration
    sender_email = "nvenu@solifetec.com"
    receiver_email = "nvenu@solifetec.com"
    password = os.getenv('EMAIL_PASSWORD')
    
    if not password:
        print("❌ Error: EMAIL_PASSWORD environment variable not set")
        print("💡 Please set EMAIL_PASSWORD in your environment")
        return False
    
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = f"RingCentral Call Logs - {date_str}"
        
        # Email body
        body = f"""
Hello,

Please find attached the RingCentral call productivity report for {date_str}.

Summary:
• Total Call Records: {total_records:,}
• Active Users/Extensions: {total_users}
• Date Range: {date_str}
• Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

The productivity report includes:
• Inbound Calls - Calls received by each user
• Outbound Calls - Calls made by each user  
• Fax Activity - Faxes sent and received
• Call Duration - Total and average talk time
• Success Rate - Connected vs missed calls
• Extension Numbers - For easy identification

This consolidated view makes it easy to assess each user's phone productivity and activity levels.

Best regards,
RingCentral Analytics System
        """.strip()
        
        msg.attach(MIMEText(body, 'plain'))
        
        # Attach CSV file
        with open(csv_filename, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename= {csv_filename}'
        )
        msg.attach(part)
        
        # Send email using Office365 SMTP
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(sender_email, password)
        text = msg.as_string()
        server.sendmail(sender_email, receiver_email, text)
        server.quit()
        
        print(f"✅ Email sent successfully to {receiver_email}")
        return True
        
    except Exception as e:
        print(f"❌ Failed to send email: {str(e)}")
        print("💡 Make sure EMAIL_PASSWORD is set and Office365 credentials are correct")
        return False

def main():
    print("🔄 Starting RingCentral Call Logs Fetch...")
    
    try:
        # Initialize SDK
        print("📡 Connecting to RingCentral...")
        rcsdk = SDK(CLIENT_ID, CLIENT_SECRET, SERVER_URL)
        platform = rcsdk.platform()
        
        # Login with JWT
        print("🔐 Authenticating with JWT token...")
        platform.login(jwt=JWT_TOKEN)
        print("✅ Authentication successful")
        
        # Load cached contacts or fetch fresh data
        extensions_cache, contacts_cache = load_cached_contacts()
        
        # Check if we need to refresh cache (daily refresh)
        cache_date_file = 'contacts/last_updated.txt'
        need_refresh = True
        
        if os.path.exists(cache_date_file):
            try:
                with open(cache_date_file, 'r') as f:
                    last_updated = f.read().strip()
                if last_updated == datetime.now().strftime('%Y-%m-%d'):
                    need_refresh = False
                    print("📋 Using today's cached contact data")
            except:
                pass
        
        if need_refresh or not extensions_cache or not contacts_cache:
            print("🔄 Refreshing contact cache...")
            extensions_cache = fetch_and_cache_extensions(platform)
            contacts_cache = fetch_and_cache_contacts(platform)
            save_cached_contacts(extensions_cache, contacts_cache)
            
            # Update cache date
            try:
                with open(cache_date_file, 'w') as f:
                    f.write(datetime.now().strftime('%Y-%m-%d'))
            except:
                pass
        
        # Calculate yesterday's date
        yesterday = datetime.now() - timedelta(days=3)
        date_from = yesterday.strftime("%Y-%m-%dT00:00:00.000Z")
        date_to = yesterday.strftime("%Y-%m-%dT23:59:59.999Z")
        date_str = yesterday.strftime("%Y-%m-%d")
        
        print(f"📅 Fetching call logs for: {date_str}")
        
        # Fetch ALL call logs using pagination
        print("📞 Fetching ALL call logs...")
        all_records = []
        page = 1
        per_page = 1000  # Maximum allowed by API
        
        while True:
            print(f"📄 Fetching page {page}...")
            response = platform.get("/restapi/v1.0/account/~/call-log", {
                "dateFrom": date_from,
                "dateTo": date_to,
                "view": "Detailed",
                "perPage": per_page,
                "page": page
            })
            
            # Get response data using safe access
            response_data = response.json()
            records = safe_get_attr(response_data, 'records', [])
            
            if not records:
                break  # No more records
            
            all_records.extend(records)
            print(f"📊 Page {page}: {len(records)} records (Total so far: {len(all_records)})")
            
            # Check if we have more pages
            paging = safe_get_attr(response_data, 'paging', {})
            has_next = safe_get_attr(paging, 'hasNextPage', False)
            
            if len(records) < per_page or not has_next:
                break  # Last page
            
            page += 1
        
        records = all_records
        print(f"🎉 Found {len(records)} total call records across all pages")
        
        if not records:
            print("⚠️  No call records found for the specified date range")
            return
        
        # Process records and group by extension/contact
        print("💾 Processing records and grouping by extension...")
        grouped_records = defaultdict(list)
        
        for i, record in enumerate(records, 1):
            print(f"🔍 Processing record {i}/{len(records)}...", end='\r')
            
            # Extract phone numbers and extension IDs safely
            from_number = None
            from_extension_id = None
            from_data = safe_get_attr(record, 'from')
            if from_data:
                from_number = safe_get_attr(from_data, 'phoneNumber')
                from_extension_id = safe_get_attr(from_data, 'extensionId')
            
            to_number = None
            to_extension_id = None
            to_data = safe_get_attr(record, 'to')
            if to_data:
                try:
                    # Handle both array and single object cases
                    if hasattr(to_data, '__len__') and len(to_data) > 0:  # Array-like
                        first_to = to_data[0]
                        to_number = safe_get_attr(first_to, 'phoneNumber')
                        to_extension_id = safe_get_attr(first_to, 'extensionId')
                    else:  # Single object
                        to_number = safe_get_attr(to_data, 'phoneNumber')
                        to_extension_id = safe_get_attr(to_data, 'extensionId')
                except:
                    to_number = None
                    to_extension_id = None
            
            # Get contact names
            from_name = get_contact_name(from_number, from_extension_id, extensions_cache, contacts_cache)
            to_name = get_contact_name(to_number, to_extension_id, extensions_cache, contacts_cache)
            
            # Determine the internal user (extension holder) for grouping
            internal_user = "External/Unknown"
            call_direction_for_user = safe_get_attr(record, 'direction', '')
            
            # For inbound calls: the internal user is the receiver (to)
            if call_direction_for_user == 'Inbound':
                if to_extension_id and str(to_extension_id) in extensions_cache:
                    internal_user = f"Ext {to_extension_id} - {extensions_cache[str(to_extension_id)]['name']}"
                    user_activity = "Received Call"
                else:
                    internal_user = f"External - {to_name}"
                    user_activity = "External Inbound"
            
            # For outbound calls: the internal user is the caller (from)
            elif call_direction_for_user == 'Outbound':
                if from_extension_id and str(from_extension_id) in extensions_cache:
                    internal_user = f"Ext {from_extension_id} - {extensions_cache[str(from_extension_id)]['name']}"
                    user_activity = "Made Call"
                else:
                    internal_user = f"External - {from_name}"
                    user_activity = "External Outbound"
            
            group_key = internal_user
            
            # Create record with names and user activity
            processed_record = {
                "internal_user": internal_user,
                "user_activity": user_activity,
                "id": safe_get_attr(record, 'id'),
                "direction": safe_get_attr(record, 'direction'),
                "from": from_number or '',
                "from_name": from_name,
                "to": to_number or '',
                "to_name": to_name,
                "external_party": to_name if call_direction_for_user == 'Outbound' else from_name,
                "startTime": safe_get_attr(record, 'startTime'),
                "duration": safe_get_attr(record, 'duration'),
                "result": safe_get_attr(record, 'result'),
                "type": safe_get_attr(record, 'type')
            }
            
            grouped_records[group_key].append(processed_record)
        
        # Ensure exports folder exists
        os.makedirs('exports', exist_ok=True)
        
        # Generate consolidated productivity summary
        csv_filename = f"exports/call_productivity_{date_str}.csv"
        
        with open(csv_filename, "w", newline="", encoding="utf-8") as f:
            fieldnames = [
                "internal_user", "extension_number",
                "inbound_calls", "outbound_calls", 
                "fax_received_count", "fax_received_contacts", 
                "fax_sent_count", "fax_sent_contacts",
                "total_calls", "total_duration_minutes", "avg_call_duration",
                "successful_calls", "missed_calls", "success_rate"
            ]
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            
            for user_name in sorted(grouped_records.keys()):
                user_records = grouped_records[user_name]
                
                # Get extension number
                ext_number = ""
                for record in user_records:
                    if "Ext " in user_name:
                        ext_number = user_name.split(" - ")[0].replace("Ext ", "")
                        break
                
                # Count by direction and type
                inbound_voice = len([r for r in user_records if r['direction'] == 'Inbound' and r['type'] == 'Voice'])
                outbound_voice = len([r for r in user_records if r['direction'] == 'Outbound' and r['type'] == 'Voice'])
                
                # Get fax details with contacts
                fax_received_records = [r for r in user_records if r['direction'] == 'Inbound' and r['type'] == 'Fax']
                fax_sent_records = [r for r in user_records if r['direction'] == 'Outbound' and r['type'] == 'Fax']
                
                fax_received_count = len(fax_received_records)
                fax_sent_count = len(fax_sent_records)
                
                # Get unique contacts for received faxes
                fax_received_contacts = []
                for record in fax_received_records:
                    contact = record['from_name'] if record['from_name'] != record['from'] else record['from']
                    if contact and contact not in fax_received_contacts:
                        fax_received_contacts.append(contact)
                
                # Get unique contacts for sent faxes
                fax_sent_contacts = []
                for record in fax_sent_records:
                    contact = record['to_name'] if record['to_name'] != record['to'] else record['to']
                    if contact and contact not in fax_sent_contacts:
                        fax_sent_contacts.append(contact)
                
                total_calls = len(user_records)
                
                # Calculate duration metrics (convert seconds to minutes)
                durations = []
                for r in user_records:
                    try:
                        duration = int(r['duration']) if r['duration'] else 0
                        durations.append(duration)
                    except:
                        durations.append(0)
                
                total_duration_seconds = sum(durations)
                total_duration_minutes = round(total_duration_seconds / 60, 2)
                avg_duration_minutes = round((total_duration_seconds / len(durations)) / 60, 2) if durations else 0
                
                # Count successful vs missed calls
                successful_calls = len([r for r in user_records if r['result'] in ['Call connected', 'Accepted', 'Received']])
                missed_calls = len([r for r in user_records if r['result'] in ['Missed', 'No Answer', 'Busy', 'Receive Error', 'Fax Receipt Error']])
                success_rate = round((successful_calls / total_calls * 100), 1) if total_calls > 0 else 0
                
                writer.writerow({
                    "internal_user": user_name,
                    "extension_number": ext_number,
                    "inbound_calls": inbound_voice,
                    "outbound_calls": outbound_voice,
                    "fax_received_count": fax_received_count,
                    "fax_received_contacts": "; ".join(fax_received_contacts) if fax_received_contacts else "",
                    "fax_sent_count": fax_sent_count,
                    "fax_sent_contacts": "; ".join(fax_sent_contacts) if fax_sent_contacts else "",
                    "total_calls": total_calls,
                    "total_duration_minutes": total_duration_minutes,
                    "avg_call_duration": avg_duration_minutes,
                    "successful_calls": successful_calls,
                    "missed_calls": missed_calls,
                    "success_rate": f"{success_rate}%"
                })
        
        print(f"\n✅ Call productivity report saved to {csv_filename}")
        print(f"📈 Total records exported: {len(records)}")
        print(f"👥 Activity tracked for {len(grouped_records)} users/extensions")
        print("📝 CSV shows: inbound/outbound calls, fax activity, duration, success rate")
        print(f"📁 File saved in exports folder")
        
        # Send email with single attachment
        email_sent = send_email_with_attachment(csv_filename, date_str, len(records), len(grouped_records))
        
        if email_sent:
            print("🎉 Process completed successfully - CSV generated and emailed!")
        else:
            print("⚠️  CSV generated but email failed - check EMAIL_PASSWORD environment variable")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        print("💡 Make sure your JWT token is valid and has ReadCallLog permission")
        sys.exit(1)

if __name__ == "__main__":
    main()