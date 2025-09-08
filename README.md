# RingCentral Call Analytics System

A comprehensive call analytics system that fetches RingCentral call logs, enriches them with company directory contacts, and generates interactive HTML dashboards.

## 🚀 Quick Start

Run the complete analytics system with one command:

```bash
python run_complete_analytics.py
```

This will:
- Fetch call logs from RingCentral
- Group calls by extension numbers for clean contact analytics
- Generate CSV reports in the `exports/` folder
- Create interactive HTML dashboards
- Open the dashboard in your browser automatically

## 📁 Project Structure

```
├── exports/                           # All generated files go here
│   ├── extension_based_analytics_*.csv    # CSV analytics reports
│   └── extension_analytics_dashboard_*.html # Interactive HTML dashboards
├── extension_based_analytics.py       # Main analytics script (extension-based)
├── extension_dashboard_generator.py   # HTML dashboard generator
├── run_complete_analytics.py         # Complete automation script
├── test_contacts_directory.py        # Test RingCentral connection
├── get_call_logs.py                  # Basic call logs fetcher
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

## 📋 Sample Output

### CSV Analytics
- Contact names with extension numbers
- Call counts (received/made/total)
- Fax counts (received/sent)
- Call durations and averages
- Internal vs external call classification

### HTML Dashboard
- Interactive charts showing top contacts by call volume
- Searchable tables with all contact details
- Responsive design with tabbed interface
- Automatic browser opening

## 🚀 Usage Examples

### Run Complete Analysis
```bash
python run_complete_analytics.py
```

### Generate Only CSV
```bash
python extension_based_analytics.py
```

### Generate Only Dashboard (from existing CSV)
```bash
python extension_dashboard_generator.py
```

### Test Connection
```bash
python test_contacts_directory.py
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