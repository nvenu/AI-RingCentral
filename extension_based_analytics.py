#!/usr/bin/env python3
"""
Extension-Based RingCentral Call Analytics
Groups calls by extension number for cleaner contact analytics
"""

import csv
import sys
import os
from collections import defaultdict
from datetime import datetime

try:
    from ringcentral import SDK
except ImportError:
    print("❌ Error: RingCentral SDK not installed")
    print("Run: pip install ringcentral")
    sys.exit(1)

# Your RingCentral credentials
CLIENT_ID = "c97CkXtDJ4wcupAJ8ieF67"
CLIENT_SECRET = "5r8CRbMwHUjarHHjytVUSE7HawPIWiW3zeDyMhntFmCs"
SERVER_URL = "https://platform.ringcentral.com"
JWT_TOKEN = "eyJraWQiOiI4NzYyZjU5OGQwNTk0NGRiODZiZjVjYTk3ODA0NzYwOCIsInR5cCI6IkpXVCIsImFsZyI6IlJTMjU2In0.eyJhdWQiOiJodHRwczovL3BsYXRmb3JtLnJpbmdjZW50cmFsLmNvbS9yZXN0YXBpL29hdXRoL3Rva2VuIiwic3ViIjoiNjM2NzI0NzAwMzEiLCJpc3MiOiJodHRwczovL3BsYXRmb3JtLnJpbmdjZW50cmFsLmNvbSIsImV4cCI6MTgxOTg0MzE5OSwiaWF0IjoxNzU2ODk0MjU0LCJqdGkiOiJGMzhkNTd5dVJIcWZFVU5uX0t0M2NnIn0.WGDnyfTBEJ05PkwpHNCRYj-0ZkbLOkS5lvgzdPeiFCACSNSUAhzPKuYYZ_zUAuqBUDatw59soFVGeid6pRZDp8etXADwvHSwe7wai_w0zvTA1i_rGJGP3pwXEmdagMMudIQx-aWCnjazbjRWNXb_fWi5vpH4sf1jxUJQAERxF4iSmr0S__imKaV0D_26-mdlw021lb1Px3NiSh0YXo2ogHSW5jQHZ-HILYi-HaCDskTDr_34ccbe1tOb9bgnixmz6SflGMjupsfiBTzxvw1Ii-quVxyzWeiVYaiIG6vtJGKsrN8ocdXWtGzNGpT57MSqNnLMUBEcBqY4amvn8UXb6g"

def safe_get_attr(obj, attr, default=None):
    """Safely get attribute from object, handling both dict and object formats"""
    try:
        if hasattr(obj, attr):
            return getattr(obj, attr, default)
        elif isinstance(obj, dict):
            return obj.get(attr, default)
        elif hasattr(obj, '__dict__'):
            return obj.__dict__.get(attr, default)
        else:
            return default
    except:
        return default

def fetch_extension_directory(platform):
    """Fetch company directory and build extension-based mapping"""
    print("📋 Fetching company directory for extension mapping...")
    
    extension_to_contact = {}
    
    try:
        # Fetch directory contacts
        response = platform.get("/restapi/v1.0/account/~/directory/contacts")
        contacts_data = response.json()
        
        # Handle different response formats
        if hasattr(contacts_data, 'records'):
            contacts = contacts_data.records
        else:
            contacts = []
        
        print(f"👥 Found {len(contacts)} contacts in directory")
        
        for contact in contacts:
            # Extract contact info safely
            first_name = safe_get_attr(contact, 'firstName', "")
            last_name = safe_get_attr(contact, 'lastName', "")
            ext_num = safe_get_attr(contact, 'extensionNumber')
            ext_id = safe_get_attr(contact, 'extensionId')
            email = safe_get_attr(contact, 'email')
            phone_objects = safe_get_attr(contact, 'phoneNumbers', [])
            
            # Build contact name
            name = f"{first_name} {last_name}".strip()
            if not name:
                name = f"Extension {ext_num or 'Unknown'}"
            
            # Extract primary phone number
            primary_phone = None
            if phone_objects:
                for phone_obj in phone_objects:
                    phone_num = safe_get_attr(phone_obj, 'phoneNumber')
                    if phone_num:
                        primary_phone = phone_num
                        break
            
            # Map extension number to contact info
            if ext_num:
                extension_to_contact[str(ext_num)] = {
                    "name": name,
                    "extensionId": ext_id,
                    "extensionNumber": ext_num,
                    "email": email,
                    "primaryPhone": primary_phone
                }
        
        print(f"📞 Mapped {len(extension_to_contact)} extensions to contacts")
        return extension_to_contact
        
    except Exception as e:
        print(f"⚠️  Warning: Could not fetch contacts directory: {str(e)}")
        return {}

def process_extension_based_calls(platform, extension_mapping):
    """Process call logs grouped by extension numbers"""
    print("📞 Fetching and processing call logs by extension...")
    
    # Initialize aggregation structure - keyed by extension number
    extension_stats = defaultdict(lambda: {
        "contact_name": None,
        "extension_number": None,
        "primary_phone": None,
        "email": None,
        "calls_received": 0,
        "calls_made": 0,
        "fax_received": 0,
        "fax_sent": 0,
        "total_inbound_duration": 0,
        "total_outbound_duration": 0,
        "total_calls": 0,
        "avg_call_duration": 0,
        "external_calls": 0,  # Calls to/from external numbers
        "internal_calls": 0   # Calls to/from other extensions
    })
    
    # Also track external numbers that don't have extensions
    external_stats = defaultdict(lambda: {
        "contact_name": None,
        "phone_number": None,
        "calls_received": 0,
        "calls_made": 0,
        "fax_received": 0,
        "fax_sent": 0,
        "total_inbound_duration": 0,
        "total_outbound_duration": 0,
        "total_calls": 0,
        "avg_call_duration": 0
    })
    
    try:
        # Fetch call logs
        params = {
            "dateFrom": "2025-08-28T00:00:00.000Z",
            "dateTo": "2025-09-04T23:59:59.999Z",
            "view": "Detailed",
            "perPage": 1000
        }
        
        response = platform.get("/restapi/v1.0/account/~/call-log", params)
        response_data = response.json()
        
        # Handle different response formats
        if hasattr(response_data, 'records'):
            records = response_data.records
        else:
            records = []
        
        total_records = len(records)
        print(f"📊 Processing {total_records} call records...")
        
        for record in records:
            # Extract record info safely
            direction = safe_get_attr(record, 'direction', "")
            call_type = safe_get_attr(record, 'type', "")
            duration = safe_get_attr(record, 'duration', 0)
            
            # Extract extension and phone information
            contact_ext = None
            contact_phone = None
            is_internal = False
            
            if direction == "Inbound":
                # For inbound calls, track the caller
                from_data = safe_get_attr(record, 'from_')
                if from_data:
                    contact_ext = safe_get_attr(from_data, 'extensionNumber')
                    contact_phone = safe_get_attr(from_data, 'phoneNumber')
                    
                    # Check if it's an internal call (has extension)
                    if contact_ext:
                        is_internal = True
                
            elif direction == "Outbound":
                # For outbound calls, track who we called
                to_data = safe_get_attr(record, 'to')
                if to_data:
                    try:
                        # Handle array-like objects
                        if hasattr(to_data, '__getitem__'):
                            to_item = to_data[0]
                        else:
                            to_item = to_data
                        
                        contact_ext = safe_get_attr(to_item, 'extensionNumber')
                        contact_phone = safe_get_attr(to_item, 'phoneNumber')
                        
                        # Check if it's an internal call (has extension)
                        if contact_ext:
                            is_internal = True
                    except:
                        continue
            
            # Skip if no contact information found
            if not contact_ext and not contact_phone:
                continue
            
            # Process based on whether it's internal (extension-based) or external
            if is_internal and contact_ext:
                # Internal call - group by extension
                ext_key = str(contact_ext)
                stats = extension_stats[ext_key]
                
                # Get contact info from directory
                if ext_key in extension_mapping:
                    contact_info = extension_mapping[ext_key]
                    stats["contact_name"] = contact_info["name"]
                    stats["extension_number"] = contact_info["extensionNumber"]
                    stats["primary_phone"] = contact_info["primaryPhone"]
                    stats["email"] = contact_info["email"]
                else:
                    stats["contact_name"] = f"Extension {contact_ext}"
                    stats["extension_number"] = contact_ext
                
                # Count call types
                if call_type == "Fax":
                    if direction == "Inbound":
                        stats["fax_received"] += 1
                    elif direction == "Outbound":
                        stats["fax_sent"] += 1
                else:
                    # Regular voice calls
                    if direction == "Inbound":
                        stats["calls_received"] += 1
                        stats["total_inbound_duration"] += duration
                    elif direction == "Outbound":
                        stats["calls_made"] += 1
                        stats["total_outbound_duration"] += duration
                
                stats["internal_calls"] += 1
                stats["total_calls"] += 1
                
            else:
                # External call - group by phone number
                if not contact_phone:
                    continue
                    
                phone_key = contact_phone
                stats = external_stats[phone_key]
                
                stats["contact_name"] = contact_phone
                stats["phone_number"] = contact_phone
                
                # Count call types
                if call_type == "Fax":
                    if direction == "Inbound":
                        stats["fax_received"] += 1
                    elif direction == "Outbound":
                        stats["fax_sent"] += 1
                else:
                    # Regular voice calls
                    if direction == "Inbound":
                        stats["calls_received"] += 1
                        stats["total_inbound_duration"] += duration
                    elif direction == "Outbound":
                        stats["calls_made"] += 1
                        stats["total_outbound_duration"] += duration
                
                stats["total_calls"] += 1
        
        print(f"✅ Processed {total_records} total call records")
        
        # Calculate averages for extension-based stats
        for stats in extension_stats.values():
            total_duration = stats["total_inbound_duration"] + stats["total_outbound_duration"]
            if stats["total_calls"] > 0:
                stats["avg_call_duration"] = round(total_duration / stats["total_calls"], 2)
        
        # Calculate averages for external stats
        for stats in external_stats.values():
            total_duration = stats["total_inbound_duration"] + stats["total_outbound_duration"]
            if stats["total_calls"] > 0:
                stats["avg_call_duration"] = round(total_duration / stats["total_calls"], 2)
        
        return extension_stats, external_stats
        
    except Exception as e:
        print(f"❌ Error processing call logs: {str(e)}")
        return {}, {}

def save_extension_analytics(extension_stats, external_stats, filename="extension_based_analytics.csv"):
    """Save extension-based analytics to CSV in organized folder structure"""
    
    # Create exports directory
    exports_dir = "exports"
    os.makedirs(exports_dir, exist_ok=True)
    
    # Update filename with full path
    full_path = os.path.join(exports_dir, filename)
    print(f"💾 Saving extension-based analytics to {full_path}...")
    
    if not extension_stats and not external_stats:
        print("⚠️  No data to save")
        return
    
    # Combine and sort all contacts by total calls
    all_contacts = []
    
    # Add extension-based contacts
    for stats in extension_stats.values():
        # Add missing fields for consistency
        stats["contact_type"] = "Internal Extension"
        stats["external_calls"] = stats["total_calls"] - stats.get("internal_calls", 0)
        all_contacts.append(stats)
    
    # Add external contacts
    for stats in external_stats.values():
        # Create a clean copy with consistent fields
        clean_stats = {
            "contact_name": stats["contact_name"],
            "contact_type": "External Number",
            "extension_number": None,
            "primary_phone": stats.get("phone_number"),
            "email": None,
            "calls_received": stats["calls_received"],
            "calls_made": stats["calls_made"],
            "total_calls": stats["total_calls"],
            "fax_received": stats["fax_received"],
            "fax_sent": stats["fax_sent"],
            "internal_calls": 0,
            "external_calls": stats["total_calls"],
            "total_inbound_duration": stats["total_inbound_duration"],
            "total_outbound_duration": stats["total_outbound_duration"],
            "avg_call_duration": stats["avg_call_duration"]
        }
        all_contacts.append(clean_stats)
    
    # Sort by total calls (descending)
    sorted_contacts = sorted(all_contacts, key=lambda x: x["total_calls"], reverse=True)
    
    fieldnames = [
        "contact_name", "contact_type", "extension_number", "primary_phone", "email",
        "calls_received", "calls_made", "total_calls",
        "fax_received", "fax_sent",
        "internal_calls", "external_calls",
        "total_inbound_duration", "total_outbound_duration", 
        "avg_call_duration"
    ]
    
    with open(full_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        
        for stats in sorted_contacts:
            writer.writerow(stats)
    
    print(f"✅ Analytics saved to {full_path}")
    print(f"📈 Total contacts analyzed: {len(sorted_contacts)}")
    
    return full_path

def print_extension_summary(extension_stats, external_stats):
    """Print a summary of the extension-based analytics"""
    if not extension_stats and not external_stats:
        return
    
    print("\n📊 EXTENSION-BASED CALL ANALYTICS SUMMARY")
    print("=" * 60)
    
    # Internal extension stats
    internal_contacts = len(extension_stats)
    internal_calls = sum(s["total_calls"] for s in extension_stats.values())
    internal_inbound = sum(s["calls_received"] for s in extension_stats.values())
    internal_outbound = sum(s["calls_made"] for s in extension_stats.values())
    
    # External number stats
    external_contacts = len(external_stats)
    external_calls = sum(s["total_calls"] for s in external_stats.values())
    external_inbound = sum(s["calls_received"] for s in external_stats.values())
    external_outbound = sum(s["calls_made"] for s in external_stats.values())
    
    print(f"🏢 Internal Extensions: {internal_contacts} contacts, {internal_calls} calls")
    print(f"   📥 Inbound: {internal_inbound} | 📤 Outbound: {internal_outbound}")
    print(f"🌐 External Numbers: {external_contacts} contacts, {external_calls} calls")
    print(f"   📥 Inbound: {external_inbound} | 📤 Outbound: {external_outbound}")
    print(f"📞 Total: {internal_calls + external_calls} calls across {internal_contacts + external_contacts} contacts")
    
    # Top 5 most active internal extensions
    if extension_stats:
        top_extensions = sorted(
            extension_stats.values(), 
            key=lambda x: x["total_calls"], 
            reverse=True
        )[:5]
        
        print(f"\n🏆 TOP 5 MOST ACTIVE EXTENSIONS:")
        for i, contact in enumerate(top_extensions, 1):
            name = contact["contact_name"]
            ext = contact["extension_number"]
            calls = contact["total_calls"]
            print(f"{i}. {name} (Ext {ext}): {calls} calls")
    
    # Top 5 most active external numbers
    if external_stats:
        top_external = sorted(
            external_stats.values(), 
            key=lambda x: x["total_calls"], 
            reverse=True
        )[:5]
        
        print(f"\n🌐 TOP 5 MOST ACTIVE EXTERNAL NUMBERS:")
        for i, contact in enumerate(top_external, 1):
            phone = contact["phone_number"]
            calls = contact["total_calls"]
            print(f"{i}. {phone}: {calls} calls")

def main():
    print("🚀 Starting Extension-Based RingCentral Call Analytics...")
    
    try:
        # Initialize SDK
        print("📡 Connecting to RingCentral...")
        rcsdk = SDK(CLIENT_ID, CLIENT_SECRET, SERVER_URL)
        platform = rcsdk.platform()
        
        # Login with JWT
        print("🔐 Authenticating with JWT token...")
        platform.login(jwt=JWT_TOKEN)
        print("✅ Authentication successful")
        
        # Fetch extension directory
        extension_mapping = fetch_extension_directory(platform)
        
        # Process call logs with extension-based grouping
        extension_stats, external_stats = process_extension_based_calls(platform, extension_mapping)
        
        # Ensure exports directory exists
        os.makedirs("exports", exist_ok=True)
        
        # Generate timestamp for filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"extension_based_analytics_{timestamp}.csv"
        
        # Save results
        csv_path = save_extension_analytics(extension_stats, external_stats, filename)
        
        # Print summary
        print_extension_summary(extension_stats, external_stats)
        
        print(f"\n🎉 Extension-based analysis complete! Check {csv_path} for detailed results.")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        print("💡 Make sure your JWT token is valid and has ReadCallLog + ReadAccounts permissions")
        sys.exit(1)

if __name__ == "__main__":
    main()