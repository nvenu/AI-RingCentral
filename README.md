# RingCentral Call Analytics System

A comprehensive call analytics system that fetches RingCentral call logs, enriches them with company directory contacts, and generates interactive HTML dashboards.

## 🚀 Quick Start

### 🎯 Generate Improved Call Productivity Reports
The main script with all the fixes and improvements:

```bash
python run_improved_reports.py
```

This creates improved CSV reports with:
- **✅ Extension numbers properly populated** from API data
- **✅ Internal users correctly mapped** with directory enrichment  
- **✅ Faxes properly separated**: Fax Sent vs Fax Received
- **✅ Voice calls and faxes analyzed separately**
- **✅ String formatting** (no float/scientific notation)
- **✅ Comprehensive validation** against RingCentral dashboard

### Aggregated Analytics Dashboard
Run the complete analytics system:

```bash
python run_complete_analytics.py
```

This will:
- Fetch call logs from RingCentral
- **Smart extension extraction** from phone numbers (reduces "Unknown" entries)
- Group calls by extension numbers for clean contact analytics
- **Separate voice calls and faxes** for better insights
- Generate CSV reports in the `exports/` folder
- Create interactive HTML dashboards
- Open the dashboard in your browser automatically

## 📁 Project Structure

```
├── exports/                           # All generated files go here
│   ├── improved_call_productivity_*.csv  # Improved aggregated reports
│   └── lubna_validation_*.csv            # Lubna validation reports
├── improved_call_logs.py             # 🎯 Main script: Improved call productivity generator
├── run_improved_reports.py           # 🚀 Wrapper: Run improved reports with validation
├── validate_lubna_data.py            # 🔍 Validation: Lubna-specific data validation
├── README.md                         # 📖 Documentation
├── .env                              # 🔐 Configuration (your credentials)
└── .env.example                      # 📋 Configuration template
└── requirements.txt                  # Python dependencies
```

## 🛠️ Setup

1. **Install dependencies:**
   ```bash
   pip install ringcentral
   ```

2. **Update credentials** in the scripts:
   - Edit `CLIENT_ID`, `CLIENT_SECRET`, and `JWT_TOKEN` in the Python files
   - Your current credentials are already configured

3. **Test connection:**
   ```bash
   python test_contacts_directory.py
   ```

## 📊 Available Scripts

| Script | Purpose | Output |
|--------|---------|--------|
| `run_complete_analytics.py` | **Main script** - runs everything | CSV + HTML in exports/ |
| `extension_based_analytics.py` | Generate extension-based analytics | CSV in exports/ |
| `extension_dashboard_generator.py` | Create HTML dashboard | HTML in exports/ |
| `test_contacts_directory.py` | Test RingCentral connection | Terminal output |
| `get_call_logs.py` | Basic call logs (legacy) | CSV in exports/ |

## 🎯 Key Features

### Extension-Based Analytics
- **Clean Contact Grouping**: Calls grouped by extension numbers
- **Internal vs External**: Separate tracking for company staff vs external contacts
- **Full Contact Details**: Names, extensions, emails, phone numbers from company directory

### Interactive Dashboard
- **Tabbed Interface**: Switch between Internal Extensions, External Numbers, and All Contacts
- **Visual Charts**: Bar charts and doughnut charts for call volume analysis
- **Search Functionality**: Search across names, extensions, or phone numbers
- **Responsive Design**: Works on desktop and mobile devices

### Organized Output
- All files saved to `exports/` folder
- Timestamped filenames for version tracking
- Automatic cleanup of old files from root directory

## 📈 Analytics Insights

The system provides:

**Internal Extensions (Company Staff):**
- Call volume per employee
- Total call duration per person
- Inbound vs outbound call patterns
- Employee contact details (name, extension, email, direct phone)

**External Communications:**
- Most active external phone numbers
- Customer/vendor communication patterns
- Fax transmission tracking
- Call duration analysis

## 🔧 Configuration

### Date Range
Edit the date range in `extension_based_analytics.py`:
```python
params = {
    "dateFrom": "2025-08-28T00:00:00.000Z",
    "dateTo": "2025-09-04T23:59:59.999Z",
    # ... other params
}
```

### Credentials
Update your RingCentral credentials in the scripts:
```python
CLIENT_ID = "your_client_id"
CLIENT_SECRET = "your_client_secret"
JWT_TOKEN = "your_jwt_token"
```

## 📊 Detailed Call Log Reports (NEW!)

Generate call-by-call detailed reports perfect for validation and compliance:

### Report Format
| Date/Time | Extension Number | Direction | Call Type | Action | From Phone | To Phone | Duration (sec) |
|-----------|------------------|-----------|-----------|---------|------------|----------|----------------|
| 2025-09-04 10:32 | 102 | Outbound | Voice | Phone Call | 3175551000 | 3175552000 | 180 |
| 2025-09-04 11:12 | 108 | Inbound | Fax | Fax Received | 3175553000 | 3175551080 | 0 |
| 2025-09-04 13:45 | 104 | Outbound | Fax | Fax Sent | 3175551040 | 3175554000 | 0 |

### Key Features
- **Extension Numbers**: Clearly mapped from extensionId fields
- **Fax Separation**: Distinct "Fax Sent" vs "Fax Received" actions
- **Validation Ready**: Compare directly with RingCentral dashboard
- **Complete Details**: Every call/fax with full information

### Validation Process
1. Run `python validate_user_data.py`
2. Enter specific extension ID (e.g., Lubna's extension)
3. Compare counts with RingCentral Dashboard → Analytics → Performance Reports
4. Verify data accuracy and completeness

## 🎯 Smart Extension Extraction

The system automatically extracts extension numbers from phone numbers, significantly reducing "External - Unknown" entries:

### Supported Formats
- `+1234567890x101` → Extension 101
- `+1234567890ext102` → Extension 102  
- `+1234567890EXT103` → Extension 103 (case-insensitive)
- `+1234567890(104)` → Extension 104
- `105` → Extension 105 (standalone 3-4 digits)

### Benefits
- **Fewer "Unknown" entries**: Properly categorizes calls with hidden extensions
- **Better attribution**: Faxes and calls mapped to correct users
- **Improved analytics**: More accurate productivity insights

## 📋 Sample Output

### CSV Analytics
- Contact names with extension numbers (including extracted ones)
- **Voice calls section**: received/made/total voice calls only
- **Fax section**: received/sent faxes with contact details
- Call durations and averages (voice calls only)
- Success rates (voice calls only)
- Internal vs external call classification

### HTML Dashboard
- Interactive charts showing top contacts by call volume
- Searchable tables with all contact details
- Responsive design with tabbed interface
- Automatic browser opening

## 🚀 Usage Examples

### 🚀 Generate Improved Reports
```bash
python run_improved_reports.py
```

### 🔍 Validate Lubna's Data
```bash
python validate_lubna_data.py
```

### 📊 Generate Reports Only (without validation)
```bash
python improved_call_logs.py
```
```

## 📊 Output Files

All generated files are saved in the `exports/` folder with timestamps:

- `extension_based_analytics_YYYYMMDD_HHMMSS.csv`
- `extension_analytics_dashboard_YYYYMMDD_HHMMSS.html`

## 🔍 Troubleshooting

**Authentication Issues:**
- Run `python test_contacts_directory.py` to verify credentials
- Check JWT token expiration
- Ensure ReadCallLog and ReadAccounts permissions

**No Data Issues:**
- Verify date range in the script
- Check if there are calls in the specified period
- Ensure company directory has contacts

**File Organization:**
- All outputs automatically go to `exports/` folder
- Old files in root directory are cleaned up automatically
- Use `run_complete_analytics.py` for best experience

## 📞 Support

The system is designed for RingCentral call analytics with company directory integration. It provides clean, organized reporting for business call analysis and contact management insights.