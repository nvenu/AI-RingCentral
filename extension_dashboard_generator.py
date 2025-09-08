#!/usr/bin/env python3
"""
Extension-Based HTML Dashboard Generator for RingCentral Call Analytics
Creates an interactive HTML dashboard from the extension-based CSV analytics data
"""

import csv
import json
from datetime import datetime
import os

def read_extension_analytics_csv(filename):
    """Read the extension-based analytics CSV file and return data"""
    contacts = []
    
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                # Convert numeric fields
                for field in ['calls_received', 'calls_made', 'total_calls', 'fax_received', 'fax_sent', 
                             'internal_calls', 'external_calls', 'total_inbound_duration', 
                             'total_outbound_duration', 'avg_call_duration']:
                    try:
                        row[field] = float(row[field]) if row[field] else 0
                    except:
                        row[field] = 0
                contacts.append(row)
        
        return contacts
    except Exception as e:
        print(f"Error reading CSV: {e}")
        return []

def generate_extension_html_dashboard(contacts, output_filename="extension_analytics_dashboard.html"):
    """Generate HTML dashboard with extension-based analytics"""
    
    if not contacts:
        print("No data to generate dashboard")
        return
    
    # Separate internal and external contacts
    internal_contacts = [c for c in contacts if c['contact_type'] == 'Internal Extension']
    external_contacts = [c for c in contacts if c['contact_type'] == 'External Number']
    
    # Calculate summary statistics
    total_contacts = len(contacts)
    total_internal = len(internal_contacts)
    total_external = len(external_contacts)
    
    total_calls = sum(c['total_calls'] for c in contacts)
    total_inbound = sum(c['calls_received'] for c in contacts)
    total_outbound = sum(c['calls_made'] for c in contacts)
    total_fax = sum(c['fax_received'] + c['fax_sent'] for c in contacts)
    total_duration = sum(c['total_inbound_duration'] + c['total_outbound_duration'] for c in contacts)
    
    internal_calls = sum(c['total_calls'] for c in internal_contacts)
    external_calls = sum(c['total_calls'] for c in external_contacts)
    
    # Get top contacts for charts
    top_internal = sorted(internal_contacts, key=lambda x: x['total_calls'], reverse=True)[:10]
    top_external = sorted(external_contacts, key=lambda x: x['total_calls'], reverse=True)[:10]
    top_duration = sorted(contacts, key=lambda x: x['total_inbound_duration'] + x['total_outbound_duration'], reverse=True)[:10]
    
    # Prepare chart data
    internal_labels = [c['contact_name'] for c in top_internal]
    internal_calls_data = [c['total_calls'] for c in top_internal]
    internal_extensions = [f"Ext {c['extension_number']}" for c in top_internal]
    
    external_labels = [c['contact_name'][:15] + '...' if len(c['contact_name']) > 15 else c['contact_name'] for c in top_external]
    external_calls_data = [c['total_calls'] for c in top_external]
    
    duration_labels = [c['contact_name'][:20] + '...' if len(c['contact_name']) > 20 else c['contact_name'] for c in top_duration]
    duration_data = [c['total_inbound_duration'] + c['total_outbound_duration'] for c in top_duration]
    
    # Generate HTML
    html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Extension-Based Call Analytics Dashboard</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }}
        
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #ff6b6b 0%, #ee5a24 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }}
        
        .header p {{
            font-size: 1.1em;
            opacity: 0.9;
        }}
        
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #f8f9fa;
        }}
        
        .stat-card {{
            background: white;
            padding: 25px;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
            transition: transform 0.3s ease;
        }}
        
        .stat-card:hover {{
            transform: translateY(-5px);
        }}
        
        .stat-card.internal {{
            border-left: 4px solid #3498db;
        }}
        
        .stat-card.external {{
            border-left: 4px solid #e74c3c;
        }}
        
        .stat-number {{
            font-size: 2.2em;
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 10px;
        }}
        
        .stat-label {{
            color: #7f8c8d;
            font-size: 0.85em;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .charts-section {{
            padding: 30px;
        }}
        
        .charts-grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 30px;
            margin-bottom: 30px;
        }}
        
        .chart-container {{
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
        }}
        
        .chart-title {{
            font-size: 1.2em;
            color: #2c3e50;
            margin-bottom: 20px;
            text-align: center;
            font-weight: 600;
        }}
        
        .section-tabs {{
            display: flex;
            background: #34495e;
            margin: 0 30px;
            border-radius: 10px 10px 0 0;
        }}
        
        .tab {{
            flex: 1;
            padding: 15px;
            text-align: center;
            color: white;
            cursor: pointer;
            transition: background 0.3s;
        }}
        
        .tab.active {{
            background: #2c3e50;
        }}
        
        .tab:hover {{
            background: #2c3e50;
        }}
        
        .table-section {{
            padding: 0 30px 30px;
        }}
        
        .table-container {{
            background: white;
            border-radius: 0 0 10px 10px;
            overflow: hidden;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
        }}
        
        .tab-content {{
            display: none;
        }}
        
        .tab-content.active {{
            display: block;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        
        th, td {{
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #ecf0f1;
        }}
        
        th {{
            background: #f8f9fa;
            font-weight: 600;
            color: #2c3e50;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        
        tr:hover {{
            background: #f8f9fa;
        }}
        
        .contact-name {{
            font-weight: 600;
            color: #2c3e50;
        }}
        
        .extension {{
            background: #3498db;
            color: white;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 0.8em;
            font-weight: bold;
        }}
        
        .phone-number {{
            color: #e74c3c;
            font-family: monospace;
        }}
        
        .duration {{
            color: #27ae60;
            font-weight: 600;
        }}
        
        .search-box {{
            margin: 20px;
            padding: 12px;
            width: calc(100% - 40px);
            border: 2px solid #ecf0f1;
            border-radius: 8px;
            font-size: 1em;
        }}
        
        .search-box:focus {{
            outline: none;
            border-color: #3498db;
        }}
        
        @media (max-width: 768px) {{
            .charts-grid {{
                grid-template-columns: 1fr;
            }}
            
            .stats-grid {{
                grid-template-columns: repeat(2, 1fr);
            }}
            
            .header h1 {{
                font-size: 2em;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📞 Extension-Based Call Analytics</h1>
            <p>Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-number">{total_contacts:,}</div>
                <div class="stat-label">Total Contacts</div>
            </div>
            <div class="stat-card internal">
                <div class="stat-number">{total_internal:,}</div>
                <div class="stat-label">Internal Extensions</div>
            </div>
            <div class="stat-card external">
                <div class="stat-number">{total_external:,}</div>
                <div class="stat-label">External Numbers</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{int(total_calls):,}</div>
                <div class="stat-label">Total Calls</div>
            </div>
            <div class="stat-card internal">
                <div class="stat-number">{int(internal_calls):,}</div>
                <div class="stat-label">Internal Calls</div>
            </div>
            <div class="stat-card external">
                <div class="stat-number">{int(external_calls):,}</div>
                <div class="stat-label">External Calls</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{int(total_fax):,}</div>
                <div class="stat-label">Total Faxes</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{int(total_duration/60):,}m</div>
                <div class="stat-label">Total Duration</div>
            </div>
        </div>
        
        <div class="charts-section">
            <div class="charts-grid">
                <div class="chart-container">
                    <div class="chart-title">🏢 Top Internal Extensions</div>
                    <canvas id="internalChart"></canvas>
                </div>
                <div class="chart-container">
                    <div class="chart-title">🌐 Top External Numbers</div>
                    <canvas id="externalChart"></canvas>
                </div>
            </div>
        </div>
        
        <div class="section-tabs">
            <div class="tab active" onclick="showTab('internal')">🏢 Internal Extensions ({total_internal})</div>
            <div class="tab" onclick="showTab('external')">🌐 External Numbers ({total_external})</div>
            <div class="tab" onclick="showTab('all')">📋 All Contacts ({total_contacts})</div>
        </div>
        
        <div class="table-section">
            <div class="table-container">
                <input type="text" class="search-box" id="searchBox" placeholder="🔍 Search contacts, extensions, or phone numbers...">
                
                <div id="internal-content" class="tab-content active">
                    <table>
                        <thead>
                            <tr>
                                <th>Contact Name</th>
                                <th>Extension</th>
                                <th>Phone</th>
                                <th>Email</th>
                                <th>Received</th>
                                <th>Made</th>
                                <th>Total</th>
                                <th>Duration</th>
                            </tr>
                        </thead>
                        <tbody>
"""
    
    # Add internal extension rows
    for contact in internal_contacts:
        total_duration_min = int((contact['total_inbound_duration'] + contact['total_outbound_duration']) / 60)
        
        html_content += f"""
                            <tr>
                                <td class="contact-name">{contact['contact_name']}</td>
                                <td><span class="extension">{contact['extension_number']}</span></td>
                                <td class="phone-number">{contact['primary_phone'] or ''}</td>
                                <td>{contact['email'] or ''}</td>
                                <td>{int(contact['calls_received'])}</td>
                                <td>{int(contact['calls_made'])}</td>
                                <td><strong>{int(contact['total_calls'])}</strong></td>
                                <td class="duration">{total_duration_min}m</td>
                            </tr>
"""
    
    html_content += """
                        </tbody>
                    </table>
                </div>
                
                <div id="external-content" class="tab-content">
                    <table>
                        <thead>
                            <tr>
                                <th>Phone Number</th>
                                <th>Received</th>
                                <th>Made</th>
                                <th>Total</th>
                                <th>Fax R/S</th>
                                <th>Duration</th>
                                <th>Avg Call</th>
                            </tr>
                        </thead>
                        <tbody>
"""
    
    # Add external number rows
    for contact in external_contacts:
        total_duration_min = int((contact['total_inbound_duration'] + contact['total_outbound_duration']) / 60)
        avg_duration_sec = int(contact['avg_call_duration'])
        
        html_content += f"""
                            <tr>
                                <td class="phone-number">{contact['contact_name']}</td>
                                <td>{int(contact['calls_received'])}</td>
                                <td>{int(contact['calls_made'])}</td>
                                <td><strong>{int(contact['total_calls'])}</strong></td>
                                <td>{int(contact['fax_received'])}/{int(contact['fax_sent'])}</td>
                                <td class="duration">{total_duration_min}m</td>
                                <td>{avg_duration_sec}s</td>
                            </tr>
"""
    
    html_content += """
                        </tbody>
                    </table>
                </div>
                
                <div id="all-content" class="tab-content">
                    <table>
                        <thead>
                            <tr>
                                <th>Contact</th>
                                <th>Type</th>
                                <th>Extension/Phone</th>
                                <th>Received</th>
                                <th>Made</th>
                                <th>Total</th>
                                <th>Duration</th>
                            </tr>
                        </thead>
                        <tbody>
"""
    
    # Add all contacts
    for contact in contacts:
        total_duration_min = int((contact['total_inbound_duration'] + contact['total_outbound_duration']) / 60)
        
        if contact['contact_type'] == 'Internal Extension':
            identifier = f"<span class='extension'>{contact['extension_number']}</span>"
        else:
            identifier = f"<span class='phone-number'>{contact['contact_name']}</span>"
        
        type_badge = "🏢 Internal" if contact['contact_type'] == 'Internal Extension' else "🌐 External"
        
        html_content += f"""
                            <tr>
                                <td class="contact-name">{contact['contact_name']}</td>
                                <td>{type_badge}</td>
                                <td>{identifier}</td>
                                <td>{int(contact['calls_received'])}</td>
                                <td>{int(contact['calls_made'])}</td>
                                <td><strong>{int(contact['total_calls'])}</strong></td>
                                <td class="duration">{total_duration_min}m</td>
                            </tr>
"""
    
    html_content += f"""
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    
    <script>
        // Chart.js configuration
        Chart.defaults.font.family = "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
        
        // Internal Extensions Chart
        const internalCtx = document.getElementById('internalChart').getContext('2d');
        new Chart(internalCtx, {{
            type: 'bar',
            data: {{
                labels: {json.dumps(internal_labels)},
                datasets: [{{
                    label: 'Total Calls',
                    data: {json.dumps(internal_calls_data)},
                    backgroundColor: '#3498db',
                    borderRadius: 4
                }}]
            }},
            options: {{
                responsive: true,
                plugins: {{
                    legend: {{
                        display: false
                    }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true
                    }}
                }}
            }}
        }});
        
        // External Numbers Chart
        const externalCtx = document.getElementById('externalChart').getContext('2d');
        new Chart(externalCtx, {{
            type: 'bar',
            data: {{
                labels: {json.dumps(external_labels)},
                datasets: [{{
                    label: 'Total Calls',
                    data: {json.dumps(external_calls_data)},
                    backgroundColor: '#e74c3c',
                    borderRadius: 4
                }}]
            }},
            options: {{
                responsive: true,
                plugins: {{
                    legend: {{
                        display: false
                    }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true
                    }}
                }}
            }}
        }});
        
        // Tab functionality
        function showTab(tabName) {{
            // Hide all tab contents
            document.querySelectorAll('.tab-content').forEach(content => {{
                content.classList.remove('active');
            }});
            
            // Remove active class from all tabs
            document.querySelectorAll('.tab').forEach(tab => {{
                tab.classList.remove('active');
            }});
            
            // Show selected tab content
            document.getElementById(tabName + '-content').classList.add('active');
            
            // Add active class to clicked tab
            event.target.classList.add('active');
        }}
        
        // Search functionality
        document.getElementById('searchBox').addEventListener('keyup', function() {{
            const searchTerm = this.value.toLowerCase();
            const activeContent = document.querySelector('.tab-content.active');
            const table = activeContent.querySelector('table');
            const rows = table.getElementsByTagName('tr');
            
            for (let i = 1; i < rows.length; i++) {{
                const row = rows[i];
                const text = row.textContent.toLowerCase();
                
                if (text.includes(searchTerm)) {{
                    row.style.display = '';
                }} else {{
                    row.style.display = 'none';
                }}
            }}
        }});
    </script>
</body>
</html>
"""
    
    # Write HTML file
    with open(output_filename, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"✅ Extension-based HTML dashboard generated: {output_filename}")
    return output_filename

def main():
    # Find the most recent extension analytics CSV file in organized folder
    exports_dir = "exports"
    
    if not os.path.exists(exports_dir):
        print("❌ No exports folder found. Run the extension analytics script first.")
        return
    
    csv_files = [f for f in os.listdir(exports_dir) if f.startswith('extension_based_analytics_') and f.endswith('.csv')]
    
    if not csv_files:
        print("❌ No extension analytics CSV files found. Run the extension analytics script first.")
        return
    
    # Use the most recent file
    latest_csv = sorted(csv_files)[-1]
    csv_path = os.path.join(exports_dir, latest_csv)
    print(f"📊 Using extension analytics file: {csv_path}")
    
    # Read data and generate dashboard
    contacts = read_extension_analytics_csv(csv_path)
    
    if contacts:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        html_filename = f"extension_analytics_dashboard_{timestamp}.html"
        html_path = os.path.join(exports_dir, html_filename)
        
        dashboard_file = generate_extension_html_dashboard(contacts, html_path)
        
        print(f"🎉 Extension-based dashboard created successfully!")
        print(f"📁 Open {dashboard_file} in your web browser to view the analytics")
        
        # Try to open in browser (optional)
        try:
            import webbrowser
            webbrowser.open(f'file://{os.path.abspath(dashboard_file)}')
            print("🌐 Opening dashboard in your default browser...")
        except:
            print("💡 Manually open the HTML file in your browser to view the dashboard")
    else:
        print("❌ No data found in CSV file")

if __name__ == "__main__":
    main()