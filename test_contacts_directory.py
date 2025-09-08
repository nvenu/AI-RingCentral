#!/usr/bin/env python3
"""
Test RingCentral Contacts Directory API
Simple script to test fetching company directory contacts
"""

import sys
import json

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

def main():
    print("🔄 Testing RingCentral Contacts Directory API...")
    
    try:
        # Initialize SDK
        print("📡 Connecting to RingCentral...")
        rcsdk = SDK(CLIENT_ID, CLIENT_SECRET, SERVER_URL)
        platform = rcsdk.platform()
        
        # Login with JWT
        print("🔐 Authenticating with JWT token...")
        platform.login(jwt=JWT_TOKEN)
        print("✅ Authentication successful")
        
        # Fetch directory contacts
        print("📋 Fetching company directory contacts...")
        response = platform.get("/restapi/v1.0/account/~/directory/contacts")
        contacts_data = response.json()
        
        # Handle different response formats
        if hasattr(contacts_data, 'records'):
            contacts = contacts_data.records
        elif isinstance(contacts_data, dict):
            contacts = contacts_data.get("records", [])
        else:
            contacts = []
        print(f"👥 Found {len(contacts)} contacts in directory")
        
        if contacts:
            print("\n📝 Sample contact structure:")
            # Convert to dict if needed for JSON serialization
            sample_contact = contacts[0]
            if hasattr(sample_contact, '__dict__'):
                sample_dict = {k: v for k, v in sample_contact.__dict__.items() if not k.startswith('_')}
                print(json.dumps(sample_dict, indent=2, default=str))
            else:
                print(json.dumps(sample_contact, indent=2, default=str))
            
            print(f"\n📊 Contact Summary:")
            print("-" * 40)
            
            for i, contact in enumerate(contacts[:10]):  # Show first 10
                # Handle both dict and object formats
                if hasattr(contact, 'firstName'):
                    first_name = getattr(contact, 'firstName', "")
                    last_name = getattr(contact, 'lastName', "")
                    ext_num = getattr(contact, 'extensionNumber', "N/A")
                    phone_numbers = getattr(contact, 'phoneNumbers', [])
                    email = getattr(contact, 'email', "N/A")
                else:
                    first_name = contact.get("firstName", "")
                    last_name = contact.get("lastName", "")
                    ext_num = contact.get("extensionNumber", "N/A")
                    phone_numbers = contact.get("phoneNumbers", [])
                    email = contact.get("email", "N/A")
                
                name = f"{first_name} {last_name}".strip()
                
                # Extract phone numbers
                phone_list = []
                if phone_numbers:
                    for p in phone_numbers:
                        if hasattr(p, 'phoneNumber'):
                            phone_list.append(getattr(p, 'phoneNumber', ""))
                        else:
                            phone_list.append(p.get("phoneNumber", ""))
                
                print(f"{i+1}. {name or 'No Name'}")
                print(f"   Extension: {ext_num}")
                print(f"   Phones: {', '.join(phone_list) if phone_list else 'None'}")
                print(f"   Email: {email}")
                print()
            
            if len(contacts) > 10:
                print(f"... and {len(contacts) - 10} more contacts")
        else:
            print("⚠️  No contacts found in directory")
        
        print("✅ Directory test completed successfully")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        print("💡 Make sure your JWT token is valid and has ReadAccounts permission")
        sys.exit(1)

if __name__ == "__main__":
    main()