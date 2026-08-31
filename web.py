import pandas as pd
from datetime import datetime
import re

# Load your Excel file
excel_file = "Stcapmark.xlsx"  # Ensure this file exists
df = pd.read_excel(excel_file)

# Get column A (first column) and all other columns which have at least one non-null (filled) value
column_a = df.iloc[:, 0]  # First column (A)

# Select all other columns that have at least one non-null value
other_cols = []
for col in df.columns[1:]:
    try:
        if df[col].notna().any():
            other_cols.append(col)
    except Exception:
        # If column operations fail for any reason, still try to include it
        other_cols.append(col)

# If no other_cols found, fallback to last 6 columns as earlier behavior (safe fallback)
if len(other_cols) == 0:
    other_cols = df.columns[-6:].tolist()

selected_columns_df = df.loc[:, other_cols]

# Combine them into a new dataframe
filtered_df = pd.concat([column_a.reset_index(drop=True), selected_columns_df.reset_index(drop=True)], axis=1)

# Ensure headers are clean strings (convert datetimes -> YYYY-MM-DD if needed)
new_cols = []
for col in filtered_df.columns:
    if isinstance(col, (pd.Timestamp, datetime)):
        new_cols.append(pd.to_datetime(col).strftime("%Y-%m-%d"))
    else:
        new_cols.append(str(col))
filtered_df.columns = new_cols

# Convert dataframe to HTML
html_table = filtered_df.to_html(index=False, border=0, table_id="marketTable", classes="dataframe")

# Full HTML with Tabs + Chart.js
html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Intraday Market Cap Dashboard</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script src="https://unpkg.com/lightweight-charts@4.1.1/dist/lightweight-charts.standalone.production.js"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&family=Roboto:wght@300;400;500&display=swap" rel="stylesheet">
  <style>
    :root {{
      --primary: #1e293b;
      --secondary: #334155;
      --accent: #6366f1;
      --success: #10b981;
      --danger: #ef4444;
      --warning: #f59e0b;
      --light: #f1f5f9;
      --dark: #0f172a;
      --text: #1e293b;
      --text-light: #64748b;
      --bg-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      --card-shadow: 0 10px 30px rgba(0,0,0,0.08);
      --hover-shadow: 0 20px 40px rgba(0,0,0,0.12);
    }}
    
    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}
    
    body {{
      font-family: 'Roboto', sans-serif;
      background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
      background-attachment: fixed;
      color: var(--text);
      line-height: 1.7;
      margin: 0;
      padding: 0;
      min-height: 100vh;
    }}
    
    .container {{
      max-width: 1440px;
      margin: 0 auto;
      padding: 28px;
    }}
    
    header {{
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 32px 28px;
      border-radius: 20px;
      margin-bottom: 32px;
      box-shadow: 0 15px 45px rgba(102, 126, 234, 0.35);
      position: relative;
      overflow: hidden;
    }}
    
    header::before {{
      content: '';
      position: absolute;
      top: -50%;
      right: -10%;
      width: 500px;
      height: 500px;
      background: rgba(255, 255, 255, 0.08);
      border-radius: 50%;
    }}
    
    header::after {{
      content: '';
      position: absolute;
      bottom: -40%;
      left: -8%;
      width: 350px;
      height: 350px;
      background: rgba(255, 255, 255, 0.05);
      border-radius: 50%;
    }}
    
    .menu-toggle {{
      background: rgba(255, 255, 255, 0.2);
      backdrop-filter: blur(10px);
      border: 2px solid rgba(255, 255, 255, 0.35);
      color: white;
      width: 48px;
      height: 48px;
      border-radius: 12px;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 20px;
      transition: all 0.3s;
    }}
    
    .menu-toggle:hover {{
      background: rgba(255, 255, 255, 0.35);
      border-color: rgba(255, 255, 255, 0.7);
      transform: scale(1.05);
    }}
    
    .sidebar {{
      position: fixed;
      top: 0;
      left: -320px;
      width: 300px;
      height: 100vh;
      background: rgba(255, 255, 255, 0.95);
      backdrop-filter: blur(20px);
      box-shadow: 4px 0 30px rgba(0,0,0,0.2);
      z-index: 2000;
      transition: left 0.3s ease-in-out;
      display: flex;
      flex-direction: column;
    }}
    
    .sidebar.open {{
      left: 0;
    }}
    
    .sidebar-header {{
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 24px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }}
    
    .sidebar-header h3 {{
      margin: 0;
      font-family: 'Montserrat', sans-serif;
      font-size: 1.3rem;
    }}
    
    .sidebar-close {{
      background: rgba(255, 255, 255, 0.2);
      border: none;
      color: white;
      font-size: 20px;
      cursor: pointer;
      width: 36px;
      height: 36px;
      display: flex;
      align-items: center;
      justify-content: center;
      border-radius: 8px;
      transition: all 0.3s;
    }}
    
    .sidebar-close:hover {{
      background: rgba(255, 255, 255, 0.3);
      transform: rotate(90deg);
    }}
    
    .sidebar-content {{
      padding: 24px;
      overflow-y: auto;
      flex: 1;
    }}
    
    .sidebar-link {{
      display: flex;
      align-items: center;
      gap: 14px;
      width: 100%;
      padding: 16px 18px;
      margin-bottom: 10px;
      background: white;
      color: var(--primary);
      border: 1px solid #e2e8f0;
      border-radius: 12px;
      cursor: pointer;
      font-size: 15px;
      font-weight: 500;
      transition: all 0.3s;
      text-align: left;
      box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }}
    
    .sidebar-link:hover {{
      background: var(--accent);
      color: white;
      border-color: var(--accent);
      transform: translateX(6px);
      box-shadow: 0 6px 20px rgba(99, 102, 241, 0.35);
    }}
    
    .sidebar-link.active {{
      background: linear-gradient(135deg, #10b981 0%, #059669 100%);
      color: white;
      border-color: transparent;
      box-shadow: 0 6px 20px rgba(16, 185, 129, 0.35);
    }}
    
    .sidebar-overlay {{
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(0, 0, 0, 0.5);
      backdrop-filter: blur(4px);
      z-index: 1999;
      opacity: 0;
      visibility: hidden;
      transition: opacity 0.3s, visibility 0.3s;
    }}
    
    .sidebar-overlay.open {{
      opacity: 1;
      visibility: visible;
    }}
    
    h1, h2, h3, h4 {{
      font-family: 'Montserrat', sans-serif;
      font-weight: 700;
    }}
    
    h1 {{
      font-size: 2.5rem;
      margin-bottom: 10px;
    }}
    
    .subtitle {{
      font-size: 1.1rem;
      opacity: 0.9;
      margin-bottom: 20px;
    }}
    
    .dashboard-stats {{
      display: flex;
      flex-wrap: wrap;
      gap: 20px;
      margin-bottom: 32px;
    }}
    
    .stat-card {{
      background: white;
      border-radius: 16px;
      padding: 24px;
      flex: 1;
      min-width: 220px;
      box-shadow: 0 10px 35px rgba(0,0,0,0.08);
      transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
      position: relative;
      overflow: hidden;
      border: 1px solid rgba(0,0,0,0.04);
    }}
    
    .stat-card:hover {{
      transform: translateY(-8px);
      box-shadow: 0 20px 50px rgba(0,0,0,0.15);
    }}
    
    .stat-card::before {{
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      width: 5px;
      height: 100%;
      background: linear-gradient(180deg, #667eea, #764ba2);
      border-radius: 16px 0 0 16px;
    }}
    
    .stat-card:nth-child(1)::before {{ background: linear-gradient(180deg, #667eea, #764ba2); }}
    .stat-card:nth-child(2)::before {{ background: linear-gradient(180deg, #f093fb, #f5576c); }}
    .stat-card:nth-child(3)::before {{ background: linear-gradient(180deg, #4facfe, #00f2fe); }}
    .stat-card:nth-child(4)::before {{ background: linear-gradient(180deg, #43e97b, #38f9d7); }}
    
    .stat-value {{
      font-size: 2.4rem;
      font-weight: 800;
      background: linear-gradient(135deg, var(--primary) 0%, var(--accent) 100%);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
      margin: 10px 0 6px;
      line-height: 1.2;
    }}
    
    .stat-label {{
      color: var(--text-light);
      font-size: 0.85rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 1px;
    }}
    
    .stat-desc {{
      color: var(--text-light);
      font-size: 0.8rem;
      opacity: 0.8;
      margin-top: 4px;
    }}
    
    /* Tabs */
    .tab-container {{
      background: white;
      border-radius: 16px;
      overflow: hidden;
      box-shadow: 0 10px 40px rgba(0,0,0,0.08);
      margin-bottom: 32px;
      border: 1px solid rgba(0,0,0,0.04);
    }}
    
    .tab {{
      display: flex;
      background: linear-gradient(90deg, #1e293b 0%, #334155 100%);
      overflow-x: auto;
    }}
    
    .tab button {{
      background: transparent;
      color: rgba(255, 255, 255, 0.85);
      border: none;
      outline: none;
      cursor: pointer;
      padding: 18px 28px;
      font-family: 'Montserrat', sans-serif;
      font-weight: 600;
      font-size: 16px;
      transition: all 0.3s;
      white-space: nowrap;
      position: relative;
    }}
    
    .tab button:hover {{
      background: rgba(255, 255, 255, 0.12);
      color: white;
    }}
    
    .tab button.active {{
      background: white;
      color: var(--primary);
      position: relative;
    }}
    
    .tab button.active::after {{
      content: '';
      position: absolute;
      bottom: 0;
      left: 0;
      width: 100%;
      height: 4px;
      background: linear-gradient(90deg, #667eea, #764ba2);
    }}
    
    .tabcontent {{
      display: none;
      padding: 28px;
      animation: fadeIn 0.4s ease;
    }}
    
    /* Data View will be shown via menu/tab selection */
    
    .controls {{
      margin-bottom: 28px;
      padding: 24px;
      background: white;
      border-radius: 16px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.06);
      border: 1px solid rgba(0,0,0,0.04);
    }}
    
    .controls h3 {{
      margin-bottom: 18px;
      color: var(--secondary);
      font-family: 'Montserrat', sans-serif;
    }}
    
    .filter-buttons {{
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      margin-bottom: 18px;
    }}

    .search-container {{
      display: flex;
      flex-wrap: nowrap;
      gap: 15px;
      align-items: center;
      margin-top: 18px;
      padding: 20px;
      background: white;
      border-radius: 14px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.06);
      overflow-x: auto;
      border: 1px solid rgba(0,0,0,0.04);
    }}

    .search-container label {{
      font-weight: 700;
      color: var(--primary);
      margin-right: 8px;
      font-size: 0.9rem;
    }}

    .search-container input, .search-container select {{
      padding: 12px 16px;
      border: 2px solid #e2e8f0;
      border-radius: 10px;
      font-size: 14px;
      transition: all 0.3s;
      background: #f8fafc;
    }}

    .search-container input:focus, .search-container select:focus {{
      border-color: var(--accent);
      outline: none;
      box-shadow: 0 0 0 4px rgba(99, 102, 241, 0.15);
      background: white;
    }}

    .search-container button {{
      padding: 12px 20px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      border: none;
      border-radius: 10px;
      cursor: pointer;
      transition: all 0.3s;
      font-weight: 600;
      display: flex;
      align-items: center;
      gap: 8px;
      box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
    }}

    .search-container button:hover {{
      transform: translateY(-2px);
      box-shadow: 0 8px 25px rgba(102, 126, 234, 0.4);
    }}

    .search-container button:active {{
      transform: translateY(0);
    }}
    
    button {{
      padding: 12px 20px;
      font-size: 14px;
      font-weight: 600;
      background: white;
      color: var(--primary);
      border: 2px solid #e2e8f0;
      border-radius: 10px;
      cursor: pointer;
      transition: all 0.3s;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }}
    
    button:hover {{
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      border-color: transparent;
      transform: translateY(-2px);
      box-shadow: 0 8px 25px rgba(102, 126, 234, 0.35);
    }}
    
    button.active {{
      background: linear-gradient(135deg, #10b981 0%, #059669 100%);
      color: white;
      border-color: transparent;
      box-shadow: 0 4px 15px rgba(16, 185, 129, 0.35);
    }}
    
    .download-btn {{
      background: linear-gradient(135deg, #10b981 0%, #059669 100%);
      color: white;
      border: none;
      padding: 14px 24px;
      margin-bottom: 28px;
      font-weight: 700;
      box-shadow: 0 6px 20px rgba(16, 185, 129, 0.35);
    }}
    
    .download-btn:hover {{
      background: linear-gradient(135deg, #059669 0%, #10b981 100%);
      transform: translateY(-2px);
      box-shadow: 0 10px 30px rgba(16, 185, 129, 0.45);
    }}
    
    /* Table Styling */
    .table-container {{
      overflow-x: auto;
      border-radius: 16px;
      box-shadow: 0 10px 40px rgba(0,0,0,0.08);
      margin-top: 24px;
      position: relative;
      border: 1px solid rgba(0,0,0,0.04);
    }}
    
    .table-scroll-area {{
      max-height: 460px;
      overflow-y: auto;
    }}
    
    .scroll-btn {{
      position: absolute;
      right: 12px;
      width: 38px;
      height: 38px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      border: none;
      border-radius: 50%;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 14px;
      z-index: 10;
      box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
      transition: all 0.3s;
    }}
    
    .scroll-btn:hover {{
      transform: scale(1.1);
      box-shadow: 0 6px 25px rgba(102, 126, 234, 0.5);
    }}
    
    .scroll-up {{
      top: 65px;
    }}
    
    .scroll-down {{
      bottom: 20px;
    }}
    
    table {{
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      font-size: 0.92rem;
    }}
    
    th {{
      background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
      color: white;
      padding: 16px 14px;
      text-align: left;
      font-weight: 700;
      position: sticky;
      top: 0;
      font-family: 'Montserrat', sans-serif;
      font-size: 0.82rem;
      text-transform: uppercase;
      letter-spacing: 0.6px;
      border-bottom: 2px solid rgba(255,255,255,0.1);
    }}
    
    th:first-child {{
      border-top-left-radius: 16px;
    }}
    
    th:last-child {{
      border-top-right-radius: 16px;
    }}
    
    #marketTable th:first-child {{
      position: sticky;
      top: 0;
      left: 0;
      z-index: 3;
    }}

    #marketTable td:first-child {{
      position: sticky;
      left: 0;
      z-index: 1;
      background: white;
      font-weight: 600;
    }}
    
    td {{
      padding: 14px 12px;
      border-bottom: 1px solid #f1f5f9;
    }}
    
    tr {{
      transition: all 0.2s;
    }}
    
    tr:nth-child(even) {{
      background: #f8fafc;
    }}
    
    tr:hover {{
      background: linear-gradient(90deg, #f0f7ff 0%, #f8fafc 100%);
      transform: scale(1.005);
      box-shadow: 0 2px 12px rgba(0,0,0,0.04);
    }}
    
    .increase-light {{ background-color: rgba(16, 185, 129, 0.12) !important; }}
    .increase-dark {{ background-color: rgba(16, 185, 129, 0.25) !important; }}
    .decrease-light {{ background-color: rgba(239, 68, 68, 0.12) !important; }}
    .decrease-dark {{ background-color: rgba(239, 68, 68, 0.25) !important; }}
    
    /* Form Elements */
    select, input {{
      padding: 12px 16px;
      border: 2px solid #e2e8f0;
      border-radius: 10px;
      font-family: 'Roboto', sans-serif;
      font-size: 14px;
      margin: 10px 0;
      width: 100%;
      max-width: 300px;
      transition: all 0.3s;
      background: #f8fafc;
    }}
    
    select:focus, input:focus {{
      outline: none;
      border-color: var(--accent);
      box-shadow: 0 0 0 4px rgba(99, 102, 241, 0.15);
      background: white;
    }}
    
    /* Chart Containers */
    .chart-container {{
      background: white;
      padding: 28px;
      border-radius: 16px;
      box-shadow: 0 10px 40px rgba(0,0,0,0.08);
      margin: 24px 0;
      border: 1px solid rgba(0,0,0,0.04);
    }}
    
    .info-panel {{
      margin-top: 24px;
      padding: 24px;
      background: white;
      border-radius: 16px;
      box-shadow: 0 10px 40px rgba(0,0,0,0.08);
      border: 1px solid rgba(0,0,0,0.04);
    }}
    
    .info-item {{
      margin-bottom: 18px;
      padding: 18px;
      background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
      border-radius: 12px;
      border-left: 5px solid var(--accent);
      box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }}
    
    .rank-history {{
      margin-top: 12px;
      font-size: 0.9em;
      color: var(--text-light);
    }}
    
    .rank-entry {{
      display: inline-block;
      margin-right: 16px;
      margin-bottom: 10px;
      padding: 8px 14px;
      background: white;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
      border: 1px solid #f1f5f9;
    }}
    
    .rank-label {{
      font-weight: 700;
      color: var(--primary);
    }}
    
    .rank-tooltip {{
      position: absolute;
      background: rgba(30, 41, 59, 0.95);
      backdrop-filter: blur(10px);
      color: white;
      padding: 14px 18px;
      border-radius: 10px;
      font-size: 13px;
      pointer-events: none;
      z-index: 1000;
      max-width: 320px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.3);
      border: 1px solid rgba(255,255,255,0.1);
    }}
    
    /* Responsive Design */
    @media (max-width: 768px) {{
      .dashboard-stats {{
        flex-direction: column;
      }}
      
      .tab {{
        flex-wrap: nowrap;
        overflow-x: auto;
      }}
      
      .tab button {{
        padding: 14px 20px;
        font-size: 14px;
      }}
      
      .filter-buttons {{
        flex-direction: column;
        align-items: stretch;
      }}
      
      button {{
        justify-content: center;
      }}
      
      .table-container {{
        overflow-x: auto;
        position: relative;
        border-radius: 12px;
      }}
      
      th, td {{
        padding: 10px 12px;
        font-size: 0.85rem;
      }}
      
      .stat-card {{
        min-width: 100%;
      }}
      
      .search-container {{
        flex-wrap: wrap;
      }}
    }}
    
    /* Animation */
    @keyframes fadeIn {{
      from {{ opacity: 0; transform: translateY(15px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}
    
    @keyframes slideInLeft {{
      from {{ opacity: 0; transform: translateX(-30px); }}
      to {{ opacity: 1; transform: translateX(0); }}
    }}
    
    .tabcontent {{
      animation: fadeIn 0.4s ease;
    }}

    #dataView {{
      animation: none !important;
    }}

    #dataView .table-container {{
      overflow-x: hidden !important;
    }}
    
    .sidebar-link {{
      animation: slideInLeft 0.3s ease forwards;
    }}
    
    .sidebar-link:nth-child(1) {{ animation-delay: 0.05s; }}
    .sidebar-link:nth-child(2) {{ animation-delay: 0.1s; }}
    .sidebar-link:nth-child(3) {{ animation-delay: 0.15s; }}
    .sidebar-link:nth-child(4) {{ animation-delay: 0.2s; }}
    .sidebar-link:nth-child(5) {{ animation-delay: 0.25s; }}
    .sidebar-link:nth-child(6) {{ animation-delay: 0.3s; }}
    
    /* Loading indicator */
    .loading {{
      display: inline-block;
      width: 20px;
      height: 20px;
      border: 3px solid rgba(255,255,255,.3);
      border-radius: 50%;
      border-top-color: #fff;
      animation: spin 1s ease-in-out infinite;
    }}
    
    @keyframes spin {{
      to {{ transform: rotate(360deg); }}
    }}
  </style>
</head>
<body>

<div class="container">
  <div class="sidebar" id="sidebar">
    <div class="sidebar-header">
      <h3>Menu</h3>
      <button class="sidebar-close" onclick="toggleMenu()">
        <i class="fas fa-times"></i>
      </button>
    </div>
    <div class="sidebar-content">
      <button class="sidebar-link" onclick="location.reload()" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border-color: transparent; box-shadow: 0 4px 15px rgba(102, 126, 234, 0.35);">
        <i class="fas fa-home"></i> Home
      </button>

      <button class="sidebar-link" onclick="openTabFromMenu(event, 'dataView')">
        <i class="fas fa-table"></i> Data View
      </button>
      <button class="sidebar-link" onclick="openTabFromMenu(event, 'stockAnalysis')">
        <i class="fas fa-chart-line"></i> Stock Analysis
      </button>
      <button class="sidebar-link" onclick="openTabFromMenu(event, 'marketOverview')">
        <i class="fas fa-globe"></i> Market Overview
      </button>
      <button class="sidebar-link" onclick="openTabFromMenu(event, 'topPerformers')">
        <i class="fas fa-trophy"></i> Top Performers
      </button>
      <button class="sidebar-link" onclick="openTabFromMenu(event, 'rankChanges')">
        <i class="fas fa-arrow-up"></i> Rank Up
      </button>
      <button class="sidebar-link" onclick="openTabFromMenu(event, 'historyRank')">
        <i class="fas fa-history"></i> History Rank
      </button>
      <button class="sidebar-link" onclick="toggleIndicatorMenu(event)" id="indicatorBtn">
        <i class="fas fa-lightbulb"></i> Indicator
        <i class="fas fa-chevron-down" id="indicatorArrow" style="margin-left: auto; font-size: 12px; transition: transform 0.3s;"></i>
      </button>
      <div id="indicatorSubmenu" style="display: none; margin-left: 15px; margin-bottom: 10px;">
        <button class="sidebar-link" onclick="showIndicatorOption('candle')" style="padding: 10px 15px; font-size: 13px; margin-bottom: 5px;">
          <i class="fas fa-chart-bar"></i> CANDLE COMBINATION
        </button>
        <button class="sidebar-link" onclick="showIndicatorOption('intraday')" style="padding: 10px 15px; font-size: 13px; margin-bottom: 5px;">
          <i class="fas fa-bolt"></i> INTRADAY
        </button>
        <button class="sidebar-link" onclick="showIndicatorOption('options')" style="padding: 10px 15px; font-size: 13px; margin-bottom: 5px;">
          <i class="fas fa-layer-group"></i> OPTIONS TRADE
        </button>
      </div>
    </div>
  </div>
  <div class="sidebar-overlay" id="sidebarOverlay" onclick="closeMenuOnOverlay()"></div>
  <header>
    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px;">
      <div>
        <h1>Market Data Analysis Dashboard</h1>
        <p class="subtitle"></p>
      </div>
      <button class="menu-toggle" id="menuToggle" onclick="toggleMenu()" style="margin-right: 12px;">
        <i class="fas fa-bars"></i>
      </button>
      <button class="menu-toggle" id="homeToggle" onclick="location.reload()" title="Home">
        <i class="fas fa-home"></i>
      </button>
    </div>
  </header>
  
  <div class="dashboard-stats">
    <div class="stat-card">
      <div style="font-size: 1.8rem; margin-bottom: 8px; opacity: 0.9;"><i class="fas fa-chart-bar"></i></div>
      <div class="stat-label">TOTAL STOCKS</div>
      <div class="stat-value" id="totalStocks">0</div>
      <div class="stat-desc">Across all exchanges</div>
    </div>
    <div class="stat-card">
      <div style="font-size: 1.8rem; margin-bottom: 8px; opacity: 0.9;"><i class="fas fa-coins"></i></div>
      <div class="stat-label">AVG MARKET CAP</div>
      <div class="stat-value" id="avgMarketCap">0</div>
      <div class="stat-desc">Latest trading session</div>
    </div>
    <div class="stat-card">
      <div style="font-size: 1.8rem; margin-bottom: 8px; opacity: 0.9;"><i class="fas fa-trophy"></i></div>
      <div class="stat-label">TOP PERFORMER</div>
      <div class="stat-value" id="topPerformer">-</div>
      <div class="stat-desc">Highest delivery %</div>
    </div>
    <div class="stat-card">
      <div style="font-size: 1.8rem; margin-bottom: 8px; opacity: 0.9;"><i class="fas fa-clock"></i></div>
      <div class="stat-label">LAST UPDATED</div>
      <div class="stat-value" id="lastUpdated">{datetime.now().strftime('%Y-%m-%d')}</div>
      <div class="stat-desc">Data freshness</div>
    </div>
  </div>

  <div class="tab-container">
    <div class="tab">
    </div>

    <!-- Data View Tab -->
    <div id="dataView" class="tabcontent">
      <button class="download-btn" onclick="downloadExcel()">
        <i class="fas fa-download"></i> Download Market Cap Data
      </button>

      <div class="controls">
        <h3><i class="fas fa-sliders-h"></i> Details of Stocks</h3>
        <div class="filter-buttons">
          <button onclick="toggleFilter('marketcap')" id="marketcapBtn">
            <i class="fas fa-money-bill-wave"></i> Market Cap
          </button>
          <button onclick="toggleFilter('delivery')" id="deliveryBtn">
            <i class="fas fa-truck-loading"></i> Delivery Qty
          </button>
          <button onclick="toggleFilter('rank')" id="rankBtn">
            <i class="fas fa-trophy"></i> Rank
          </button>
          <button onclick="arrangeRank()" id="rankArrangeBtn">
            <i class="fas fa-sort-amount-down"></i> Rank Arrange
          </button>
          <button onclick="filterIntradayPicks()" id="intradayBtn">
            <i class="fas fa-bolt"></i> Intraday Stocks
          </button>
          <button onclick="showAll()" id="showAllBtn">
            <i class="fas fa-eye"></i> Show All
          </button>
        </div>

        <div class="search-container">
          <input type="text" id="searchInput" placeholder="Search stock name..." list="stockSuggestions">
          <datalist id="stockSuggestions"></datalist>
          <button onclick="searchStock()"><i class="fas fa-search"></i> Search</button>
          <button onclick="clearSearch()"><i class="fas fa-times"></i> Clear</button>
          <label for="startDate">Start Date:</label>
          <input type="date" id="startDate">
          <label for="endDate">End Date:</label>
          <input type="date" id="endDate">
          <button onclick="filterByDateRange()"><i class="fas fa-calendar-alt"></i> Filter Dates</button>
          <button onclick="clearDateFilters()"><i class="fas fa-calendar-times"></i> Clear Dates</button>
        </div>
      </div>

      <div class="table-container">
        <div class="table-scroll-area" id="tableScrollArea">
        {html_table}
        </div>
        <button class="scroll-btn scroll-up" onclick="scrollTable(-1)" title="Scroll Up">
          <i class="fas fa-chevron-up"></i>
        </button>
        <button class="scroll-btn scroll-down" onclick="scrollTable(1)" title="Scroll Down">
          <i class="fas fa-chevron-down"></i>
        </button>
      </div>
    </div>

    <!-- Stock Analysis Tab -->
    <div id="stockAnalysis" class="tabcontent">
      <h3><i class="fas fa-chart-line"></i> Individual Stock Analysis</h3>
      <button onclick="openTab(null, lastStockSource || 'historyRank')" style="margin-bottom: 18px; padding: 10px 20px; border: none; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border-radius: 10px; cursor: pointer; font-weight: 600; box-shadow: 0 4px 15px rgba(102, 126, 234, 0.35); transition: all 0.3s; display: inline-flex; align-items: center; gap: 8px;">
        <i class="fas fa-arrow-left"></i> Back
      </button>
      
      <div class="controls" style="margin-bottom: 20px; padding: 20px; background: white; border-radius: 14px; box-shadow: 0 4px 20px rgba(0,0,0,0.06); border: 1px solid rgba(0,0,0,0.04);">
        <label for="stockSelect" style="font-weight: 700; color: var(--primary); margin-right: 12px; font-size: 0.95rem;">Select Stock:</label>
        <select id="stockSelect" onchange="const d=getStockAnalysisDuration();updateChart(d?d.start:'',d?d.end:'');" style="padding: 12px 16px; border: 2px solid #e2e8f0; border-radius: 10px; font-size: 14px; background: #f8fafc; transition: all 0.3s; min-width: 200px;"></select>
        <button onclick="openTradingView()" style="margin-left: 15px; padding: 12px 20px; background: linear-gradient(135deg, #2962FF 0%, #1E53E5 100%); color: white; border: none; border-radius: 10px; cursor: pointer; font-weight: 600; font-size: 14px; display: inline-flex; align-items: center; gap: 8px; box-shadow: 0 4px 15px rgba(41, 98, 255, 0.3); transition: all 0.3s;">
          <i class="fas fa-chart-line"></i> View on TradingView
        </button>
        <label for="stockAnalysisStartDate" style="font-weight: 700; color: var(--primary); margin-right: 8px; font-size: 0.95rem; margin-left: 20px;">Start Date:</label>
        <input type="date" id="stockAnalysisStartDate" onchange="filterStockAnalysisByDuration()" style="padding: 8px 12px; border: 2px solid #e2e8f0; border-radius: 8px; font-size: 13px; background: #f8fafc; transition: all 0.3s; width: 140px;">
        <label for="stockAnalysisEndDate" style="font-weight: 700; color: var(--primary); margin-right: 8px; font-size: 0.95rem; margin-left: 12px;">End Date:</label>
        <input type="date" id="stockAnalysisEndDate" onchange="filterStockAnalysisByDuration()" style="padding: 8px 12px; border: 2px solid #e2e8f0; border-radius: 8px; font-size: 13px; background: #f8fafc; transition: all 0.3s; width: 140px;">
        <button onclick="clearStockAnalysisDuration()" id="clearStockAnalysisDurationBtn" style="margin-left: 12px; width: 42px; height: 42px; padding: 0; display: inline-flex; align-items: center; justify-content: center; border: none; border-radius: 10px; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: white; cursor: pointer; box-shadow: 0 4px 12px rgba(239, 68, 68, 0.3); transition: all 0.3s;">
          <i class="fas fa-times"></i>
        </button>
      </div>
      
      <div class="chart-container">
        <div style="position: relative; height:400px; margin-top: 20px;">
          <canvas id="stockChart"></canvas>
        </div>
      </div>
      <div class="info-panel" id="marketCapExtremes" style="margin-top: 20px;">
        <h4><i class="fas fa-chart-bar"></i> Market Cap Extremes</h4>
        <div id="highestMarketCap" style="margin-top: 10px; padding: 12px; background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%); border-radius: 8px; border-left: 4px solid #28a745;">
          <strong>Highest Market Cap:</strong> <span id="highestMCValue">-</span> on <span id="highestMCDate">-</span>
        </div>
        <div id="lowestMarketCap" style="margin-top: 10px; padding: 12px; background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%); border-radius: 8px; border-left: 4px solid #dc3545;">
          <strong>Lowest Market Cap:</strong> <span id="lowestMCValue">-</span> on <span id="lowestMCDate">-</span>
        </div>
      </div>
    </div>

    <!-- Market Overview Tab -->
    <div id="marketOverview" class="tabcontent">
      <h3><i class="fas fa-globe"></i> Market Overview</h3>
      
      <button onclick="showTop50MarketCap()" id="top50Btn">
        <i class="fas fa-chart-bar"></i> Show Top 50 Market Cap
      </button>
      
      <div class="chart-container">
        <div style="position: relative; height:500px;">
          <canvas id="marketOverviewChart"></canvas>
        </div>
      </div>
      
      <div class="info-panel" id="marketOverviewInfo">
        <p>Click "Show Top 50 Market Cap" to display the histogram and stock information.</p>
      </div>
    </div>

    <!-- Top Performers Tab -->
    <div id="topPerformers" class="tabcontent">
      <h3><i class="fas fa-trophy"></i> Top 10 Performing Stocks</h3>

      <div class="tab-container" style="margin-bottom: 20px; background: white; border-radius: 14px; box-shadow: 0 4px 20px rgba(0,0,0,0.06); border: 1px solid rgba(0,0,0,0.04); overflow: hidden;">
        <div style="display: flex; background: #f1f5f9;">
          <button onclick="switchTopPerformersTab('regular')" id="topPerformersRegularTab" style="flex: 1; padding: 14px 20px; border: none; background: white; color: var(--primary); font-weight: 700; cursor: pointer; font-size: 14px; border-bottom: 3px solid var(--accent); transition: all 0.3s;">
            <i class="fas fa-table"></i> Regular
          </button>
          <button onclick="switchTopPerformersTab('duration')" id="topPerformersDurationTab" style="flex: 1; padding: 14px 20px; border: none; background: #f1f5f9; color: var(--text-light); font-weight: 600; cursor: pointer; font-size: 14px; border-bottom: 3px solid transparent; transition: all 0.3s;">
            <i class="fas fa-calendar-week"></i> Duration
          </button>
        </div>

        <div style="padding: 20px;">
          <div id="topPerformersRegularControl">
            <div class="info-panel" id="topPerformersList">
              <p>Click "Show Top Performers" to display the list of top 10 stocks by delivery percentage.</p>
            </div>

            <button onclick="showTopPerformers()" id="showTopPerformersBtn" style="margin-top: 15px;">
              <i class="fas fa-list"></i> Show Top Performers
            </button>
          </div>

          <div id="topPerformersDurationControl" style="display: none;">
            <label for="topPerformersStartDate" style="font-weight: 700; color: var(--primary); margin-right: 8px; font-size: 0.95rem;">Start Date:</label>
            <input type="date" id="topPerformersStartDate" onchange="filterTopPerformersByDuration()" style="padding: 12px 16px; border: 2px solid #e2e8f0; border-radius: 10px; font-size: 14px; background: #f8fafc; transition: all 0.3s;">
            <label for="topPerformersEndDate" style="font-weight: 700; color: var(--primary); margin-right: 8px; font-size: 0.95rem; margin-left: 12px;">End Date:</label>
            <input type="date" id="topPerformersEndDate" onchange="filterTopPerformersByDuration()" style="padding: 12px 16px; border: 2px solid #e2e8f0; border-radius: 10px; font-size: 14px; background: #f8fafc; transition: all 0.3s;">
            <button onclick="clearTopPerformersDuration()" id="clearTopPerformersDurationBtn" style="margin-left: 12px; width: 42px; height: 42px; padding: 0; display: inline-flex; align-items: center; justify-content: center; border: none; border-radius: 10px; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: white; cursor: pointer; box-shadow: 0 4px 12px rgba(239, 68, 68, 0.3); transition: all 0.3s;">
              <i class="fas fa-times"></i>
            </button>
          </div>
        </div>
      </div>

      <div class="chart-container" id="topPerformersChartContainer" style="margin-top: 24px; display: none;">
        <h4 style="margin-bottom: 16px; color: var(--secondary); font-family: 'Montserrat', sans-serif;">🏆 Top 10 by Delivery Qty &gt; 70, Dates Count</h4>
        <div style="position: relative; height: 350px;">
          <canvas id="topPerformersChart"></canvas>
        </div>
      </div>
    </div>

    <!-- Rank Changes Tab -->
    <div id="rankChanges" class="tabcontent">
      <h3><i class="fas fa-arrow-up"></i> Top 10 Rank Improvers</h3>

      <div class="info-panel" id="rankChangesList">
        <p>Click "Show Rank Changes" to display stocks with the biggest rank improvements (yesterday vs today).</p>
      </div>

      <button onclick="showRankChanges()" id="showRankChangesBtn">
        <i class="fas fa-list"></i> Show Rank Changes
      </button>
    </div>

    <!-- History Rank Tab -->
    <div id="historyRank" class="tabcontent">
      <h3><i class="fas fa-history"></i> Top 10 Rank Improvers by Date</h3>

      <div class="tab-container" style="margin-bottom: 20px; background: white; border-radius: 14px; box-shadow: 0 4px 20px rgba(0,0,0,0.06); border: 1px solid rgba(0,0,0,0.04); overflow: hidden;">
        <div style="display: flex; background: #f1f5f9;">
          <button onclick="switchHistoryRankTab('single')" id="historyRankSingleTab" style="flex: 1; padding: 14px 20px; border: none; background: white; color: var(--primary); font-weight: 700; cursor: pointer; font-size: 14px; border-bottom: 3px solid var(--accent); transition: all 0.3s;">
            <i class="fas fa-calendar-day"></i> Select Date
          </button>
          <button onclick="switchHistoryRankTab('duration')" id="historyRankDurationTab" style="flex: 1; padding: 14px 20px; border: none; background: #f1f5f9; color: var(--text-light); font-weight: 600; cursor: pointer; font-size: 14px; border-bottom: 3px solid transparent; transition: all 0.3s;">
            <i class="fas fa-calendar-week"></i> Duration Date
          </button>
        </div>

        <div style="padding: 20px;">
          <div id="historyRankSingleControl">
            <label for="historyRankDatePicker" style="font-weight: 700; color: var(--primary); margin-right: 12px; font-size: 0.95rem;">Select Date:</label>
            <input type="date" id="historyRankDatePicker" onchange="showHistoryRankByDate()" style="padding: 12px 16px; border: 2px solid #e2e8f0; border-radius: 10px; font-size: 14px; background: #f8fafc; transition: all 0.3s;">
            <button onclick="clearHistoryRankDate()" id="clearHistoryRankDateBtn" style="margin-left: 12px; width: 42px; height: 42px; padding: 0; display: inline-flex; align-items: center; justify-content: center; border: none; border-radius: 10px; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: white; cursor: pointer; box-shadow: 0 4px 12px rgba(239, 68, 68, 0.3); transition: all 0.3s;">
              <i class="fas fa-times"></i>
            </button>
          </div>

          <div id="historyRankDurationControl" style="display: none;">
            <label for="historyRankStartDate" style="font-weight: 700; color: var(--primary); margin-right: 8px; font-size: 0.95rem;">Start Date:</label>
            <input type="date" id="historyRankStartDate" onchange="filterHistoryRankByDuration()" style="padding: 12px 16px; border: 2px solid #e2e8f0; border-radius: 10px; font-size: 14px; background: #f8fafc; transition: all 0.3s;">
            <label for="historyRankEndDate" style="font-weight: 700; color: var(--primary); margin-right: 8px; font-size: 0.95rem; margin-left: 12px;">End Date:</label>
            <input type="date" id="historyRankEndDate" onchange="filterHistoryRankByDuration()" style="padding: 12px 16px; border: 2px solid #e2e8f0; border-radius: 10px; font-size: 14px; background: #f8fafc; transition: all 0.3s;">
            <button onclick="clearHistoryRankDuration()" id="clearHistoryRankDurationBtn" style="margin-left: 12px; width: 42px; height: 42px; padding: 0; display: inline-flex; align-items: center; justify-content: center; border: none; border-radius: 10px; background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: white; cursor: pointer; box-shadow: 0 4px 12px rgba(239, 68, 68, 0.3); transition: all 0.3s;">
              <i class="fas fa-times"></i>
            </button>
          </div>
        </div>
      </div>

      <div class="info-panel" id="historyRankList">
        <p>Click "Show History Rank" to display top 10 rank improvers for each date transition.</p>
      </div>

      <button onclick="showHistoryRank()" id="showHistoryRankBtn" style="padding: 14px 28px; font-size: 15px; font-weight: 700; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 12px; cursor: pointer; box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4); transition: all 0.3s; display: inline-flex; align-items: center; gap: 10px;">
        <i class="fas fa-history"></i> Show History Rank
      </button>
    </div>
  </div>
</div>

<script>
// ---- TAB SWITCHING ----
function openTab(evt, tabName) {{
  const tabcontent = document.getElementsByClassName("tabcontent");
  for (let i = 0; i < tabcontent.length; i++) {{
    tabcontent[i].style.display = "none";
  }}
  const tablinks = document.getElementsByClassName("tablinks");
  for (let i = 0; i < tablinks.length; i++) {{
    tablinks[i].classList.remove("active");
  }}
  document.getElementById(tabName).style.display = "block";
  if (evt && evt.currentTarget) {{
    evt.currentTarget.classList.add("active");
  }}
  updateSidebarActive(tabName);
  toggleMenu(false);
  const dashboardStats = document.querySelector('.dashboard-stats');
  if (dashboardStats) dashboardStats.style.display = 'none';
}}

function openTabFromMenu(evt, tabName) {{
  const tabcontent = document.getElementsByClassName("tabcontent");
  for (let i = 0; i < tabcontent.length; i++) {{
    tabcontent[i].style.display = "none";
  }}
  const tablinks = document.getElementsByClassName("tablinks");
  for (let i = 0; i < tablinks.length; i++) {{
    tablinks[i].classList.remove("active");
  }}
  document.getElementById(tabName).style.display = "block";
  if (evt && evt.currentTarget) {{
    evt.currentTarget.classList.add("active");
  }}
  updateSidebarActive(tabName);
  toggleMenu(false);
  const dashboardStats = document.querySelector('.dashboard-stats');
  if (dashboardStats) dashboardStats.style.display = 'none';
}}

function updateSidebarActive(tabName) {{
  const sidebarLinks = document.querySelectorAll('.sidebar-link');
  sidebarLinks.forEach(link => {{
    link.classList.remove('active');
    if (link.getAttribute('onclick').includes("'" + tabName + "'")) {{
      link.classList.add('active');
    }}
  }});
}}

function toggleMenu(forceState) {{
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('sidebarOverlay');
  const isOpen = sidebar.classList.contains('open');
  const shouldOpen = typeof forceState === 'boolean' ? forceState : !isOpen;

  if (shouldOpen) {{
    sidebar.classList.add('open');
    overlay.classList.add('open');
  }} else {{
    sidebar.classList.remove('open');
    overlay.classList.remove('open');
  }}
}}

function closeMenuOnOverlay() {{
  toggleMenu(false);
}}

// ---- DASHBOARD STATS ----
function updateDashboardStats() {{
  const table = document.getElementById('marketTable');
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  
  // Update total stocks
  document.getElementById('totalStocks').textContent = rows.length;
  
  // Find the latest market cap column
  const headers = Array.from(table.querySelectorAll('th'));
  let marketCapColIndex = -1;
  
  for (let i = headers.length - 1; i >= 1; i--) {{
    const headerText = headers[i].textContent.trim().toLowerCase();
    if (headerText.includes('market')) {{
      marketCapColIndex = i;
      break;
    }}
  }}
  
  // Calculate average market cap
  if (marketCapColIndex !== -1) {{
    let total = 0;
    let count = 0;
    
    rows.forEach(row => {{
      const cell = row.cells[marketCapColIndex];
      if (cell) {{
        const value = parseFloat(cell.textContent.trim().replace(/,/g, ''));
        if (!isNaN(value)) {{
          total += value;
          count++;
        }}
      }}
    }});
    
    if (count > 0) {{
      const avg = total / count;
      document.getElementById('avgMarketCap').textContent = avg.toLocaleString(undefined, {{
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
      }});
    }}
  }}
  
  // Find top performer (highest delivery %)
  let deliveryColIndex = -1;
  for (let i = headers.length - 1; i >= 1; i--) {{
    const headerText = headers[i].textContent.trim().toLowerCase();
    if (headerText.includes('delivery')) {{
      deliveryColIndex = i;
      break;
    }}
  }}
  
  if (deliveryColIndex !== -1) {{
    let topPerformer = '';
    let topValue = 0;
    
    rows.forEach(row => {{
      const cell = row.cells[deliveryColIndex];
      if (cell) {{
        const value = parseFloat(cell.textContent.trim().replace(/,/g, '').replace('%', ''));
        if (!isNaN(value) && value > topValue) {{
          topValue = value;
          topPerformer = row.cells[0].textContent.trim();
        }}
      }}
    }});
    
    if (topPerformer) {{
      document.getElementById('topPerformer').textContent = topPerformer;
    }}
  }}
}}

// ---- TABLE FILTERS ----
const originalTable = document.getElementById('marketTable').cloneNode(true);
let activeFilters = new Set();

function toggleFilter(type) {{
  const btn = document.getElementById(type + 'Btn');
  if (activeFilters.has(type)) {{
    activeFilters.delete(type);
    btn.classList.remove('active');
  }} else {{
    activeFilters.add(type);
    btn.classList.add('active');
  }}
  applyFilters();
}}

function showAll() {{
  activeFilters.clear();
  document.querySelectorAll('.filter-buttons button').forEach(btn => btn.classList.remove('active'));
  document.getElementById('showAllBtn').classList.add('active');
  const parent = document.getElementById('marketTable').parentNode;
  parent.removeChild(document.getElementById('marketTable'));
  const clone = originalTable.cloneNode(true);
  parent.appendChild(clone);
  applyColorFormatting();
  updateDashboardStats();
}}

function applyFilters() {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  const lcHeaders = headers.map(h => h.textContent.trim().toLowerCase());

  const columnsToShow = new Set();
  columnsToShow.add(0);

  if (activeFilters.size === 0) {{
    for (let i = 1; i < headers.length; i++) columnsToShow.add(i);
  }} else {{
    for (let i = 1; i < headers.length; i++) {{
      const h = lcHeaders[i];
      activeFilters.forEach(f => {{
        if (f === 'marketcap' && h.includes('market')) columnsToShow.add(i);
        if (f === 'delivery' && h.includes('delivery')) columnsToShow.add(i);
        if (f === 'rank' && h.includes('rank')) columnsToShow.add(i);
      }});
    }}
  }}

  headers.forEach((h, idx) => {{
    h.style.display = columnsToShow.has(idx) ? '' : 'none';
  }});
  rows.forEach(row => {{
    const cells = Array.from(row.querySelectorAll('td'));
    cells.forEach((cell, idx) => {{
      cell.style.display = columnsToShow.has(idx) ? '' : 'none';
    }});
  }});

  document.getElementById('showAllBtn').classList.remove('active');
  applyColorFormatting();
}}

function applyColorFormatting() {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  const lcHeaders = headers.map(h => h.textContent.trim().toLowerCase());
  
  // Find indices for market cap, delivery, and rank columns
  const marketCapIndices = [];
  const deliveryIndices = [];
  const rankIndices = [];
  
  for (let i = 1; i < headers.length; i++) {{
    const h = lcHeaders[i];
    if (h.includes('market')) marketCapIndices.push(i);
    else if (h.includes('delivery')) deliveryIndices.push(i);
    else if (h.includes('rank')) rankIndices.push(i);
  }}
  
  // Apply color formatting for each column type
  rows.forEach(row => {{
    const cells = row.querySelectorAll('td');
    
    // Color formatting for Market Cap columns
    marketCapIndices.forEach((colIndex, idx) => {{
      if (idx > 0) {{ // Skip first column as there's no previous to compare
        const prevColIndex = marketCapIndices[idx - 1];
        const prevTxt = cells[prevColIndex].textContent.trim().replace(/,/g,'').replace('%','');
        const curTxt = cells[colIndex].textContent.trim().replace(/,/g,'').replace('%','');
        const prev = parseFloat(prevTxt);
        const cur = parseFloat(curTxt);
        
        cells[colIndex].classList.remove('increase-light','increase-dark','decrease-light','decrease-dark');
        
        if (!isNaN(prev) && !isNaN(cur) && prev !== 0) {{
          const change = ((cur - prev) / Math.abs(prev)) * 100;
          if (change > 0) {{
            if (change >= 90) cells[colIndex].classList.add('increase-dark');
            else if (change >= 10) cells[colIndex].classList.add('increase-light');
          }} else if (change < 0) {{
            if (change <= -90) cells[colIndex].classList.add('decrease-dark');
            else if (change <= -10) cells[colIndex].classList.add('decrease-light');
          }}
        }}
      }}
    }});
    
    // Color formatting for Delivery Qty columns
    deliveryIndices.forEach((colIndex, idx) => {{
      if (idx > 0) {{ // Skip first column as there's no previous to compare
        const prevColIndex = deliveryIndices[idx - 1];
        const prevTxt = cells[prevColIndex].textContent.trim().replace(/,/g,'').replace('%','');
        const curTxt = cells[colIndex].textContent.trim().replace(/,/g,'').replace('%','');
        const prev = parseFloat(prevTxt);
        const cur = parseFloat(curTxt);
        
        cells[colIndex].classList.remove('increase-light','increase-dark','decrease-light','decrease-dark');
        
        if (!isNaN(prev) && !isNaN(cur) && prev !== 0) {{
          const change = ((cur - prev) / Math.abs(prev)) * 100;
          if (change > 0) {{
            if (change >= 90) cells[colIndex].classList.add('increase-dark');
            else if (change >= 10) cells[colIndex].classList.add('increase-light');
          }} else if (change < 0) {{
            if (change <= -90) cells[colIndex].classList.add('decrease-dark');
            else if (change <= -10) cells[colIndex].classList.add('decrease-light');
          }}
        }}
      }}
    }});
    
    // Color formatting for Rank columns
    rankIndices.forEach((colIndex, idx) => {{
      if (idx > 0) {{ // Skip first column as there's no previous to compare
        const prevColIndex = rankIndices[idx - 1];
        const prevTxt = cells[prevColIndex].textContent.trim().replace(/,/g,'').replace(/[^0-9]/g, '');
        const curTxt = cells[colIndex].textContent.trim().replace(/,/g,'').replace(/[^0-9]/g, '');
        const prev = parseFloat(prevTxt);
        const cur = parseFloat(curTxt);
        
        cells[colIndex].classList.remove('increase-light','increase-dark','decrease-light','decrease-dark');
        
        if (!isNaN(prev) && !isNaN(cur) && prev !== 0) {{
          // For rank, lower is better, so we reverse the logic
          if (cur < prev) {{ // Improved rank (lower number)
            cells[colIndex].classList.add('increase-dark');
          }} else if (cur > prev) {{ // Worsened rank (higher number)
            cells[colIndex].classList.add('decrease-dark');
          }}
        }}
      }}
    }});
  }});
}}

// ---- SEARCH FUNCTIONALITY ----
function populateStockSuggestions() {{
  const table = document.getElementById('marketTable');
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  const datalist = document.getElementById('stockSuggestions');
  datalist.innerHTML = '';
  const stockNames = new Set();
  rows.forEach(row => {{
    const stock = row.cells[0].textContent.trim();
    if (stock) stockNames.add(stock);
  }});
  stockNames.forEach(name => {{
    const option = document.createElement('option');
    option.value = name;
    datalist.appendChild(option);
  }});
}}

function searchStock() {{
  const searchTerm = document.getElementById('searchInput').value.trim().toLowerCase();
  const table = document.getElementById('marketTable');
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  rows.forEach(row => {{
    const stockName = row.cells[0].textContent.trim().toLowerCase();
    if (stockName.includes(searchTerm)) {{
      row.style.display = '';
    }} else {{
      row.style.display = 'none';
    }}
  }});
  applyColorFormatting();
  updateDashboardStats();
}}

function clearSearch() {{
  document.getElementById('searchInput').value = '';
  const table = document.getElementById('marketTable');
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  rows.forEach(row => row.style.display = '');
  applyColorFormatting();
  updateDashboardStats();
}}

function setDatePickerRanges() {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const datePattern = /^\\d{{4}}-\\d{{2}}-\\d{{2}}/;
  const dates = [];

  headers.forEach((header, index) => {{
    if (index === 0) return;
    const text = header.textContent.trim();
    const match = text.match(datePattern);
    if (match) {{
      dates.push(match[0]);
    }}
  }});

  if (dates.length > 0) {{
    dates.sort();
    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');
    startDateInput.min = dates[0];
    startDateInput.max = dates[dates.length - 1];
    endDateInput.min = dates[0];
    endDateInput.max = dates[dates.length - 1];
    startDateInput.value = dates[0];
    endDateInput.value = dates[dates.length - 1];

    // Dynamic min/max updates
    startDateInput.addEventListener('change', function() {{
      if (this.value) {{
        endDateInput.min = this.value;
      }}
    }});
    endDateInput.addEventListener('change', function() {{
      if (this.value) {{
        startDateInput.max = this.value;
      }}
    }});
  }}
}}

function clearDateFilters() {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const datePattern = /^\\d{{4}}-\\d{{2}}-\\d{{2}}/;
  const dates = [];

  headers.forEach((header, index) => {{
    if (index === 0) return;
    const text = header.textContent.trim();
    const match = text.match(datePattern);
    if (match) {{
      dates.push(match[0]);
    }}
  }});

  if (dates.length > 0) {{
    dates.sort();
    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');
    startDateInput.value = '';
    endDateInput.value = '';
    startDateInput.min = dates[0];
    startDateInput.max = dates[dates.length - 1];
    endDateInput.min = dates[0];
    endDateInput.max = dates[dates.length - 1];
  }}

  filterByDateRange();
}}

function filterByDateRange() {{
  const startDateInput = document.getElementById('startDate').value;
  const endDateInput = document.getElementById('endDate').value;
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  const datePattern = /^\\d{{4}}-\\d{{2}}-\\d{{2}}/;

  const startDate = startDateInput ? new Date(startDateInput) : null;
  const endDate = endDateInput ? new Date(endDateInput) : null;

  headers.forEach((header, index) => {{
    if (index === 0) {{
      header.style.display = '';
      return;
    }}
    const text = header.textContent.trim();
    const match = text.match(datePattern);
    if (!match) {{
      header.style.display = 'none';
      return;
    }}
    const colDate = new Date(match[0]);
    let show = true;
    if (startDate && colDate < startDate) show = false;
    if (endDate && colDate > endDate) show = false;
    header.style.display = show ? '' : 'none';
  }});

  rows.forEach(row => {{
    const cells = Array.from(row.querySelectorAll('td'));
    cells.forEach((cell, index) => {{
      if (index === 0) {{
        cell.style.display = '';
        return;
      }}
      const header = headers[index];
      cell.style.display = header.style.display;
    }});
  }});

  applyColorFormatting();
  updateDashboardStats();
}}

function arrangeRank() {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const tbody = table.querySelector('tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));

  const rankIndices = [];
  headers.forEach((h, i) => {{
    if (h.textContent.toLowerCase().includes('rank')) rankIndices.push(i);
  }});
  if (rankIndices.length === 0) {{
    alert('No rank column found!');
    return;
  }}
  const rankIndex = rankIndices[rankIndices.length - 1];
  rows.sort((a, b) => {{
    const aCell = a.cells[rankIndex]?.textContent.trim() || '';
    const bCell = b.cells[rankIndex]?.textContent.trim() || '';
    const aBlank = (aCell === '' || aCell === '-' || aCell.toLowerCase() === 'nan');
    const bBlank = (bCell === '' || bCell === '-' || bCell.toLowerCase() === 'nan');
    if (aBlank && bBlank) return 0;
    if (aBlank) return 1;
    if (bBlank) return -1;
    const aVal = parseInt(aCell.replace(/[^0-9]/g,''), 10);
    const bVal = parseInt(bCell.replace(/[^0-9]/g,''), 10);
    return (isNaN(aVal) ? 1 : aVal) - (isNaN(bVal) ? 1 : bVal);
  }});
  while (tbody.firstChild) tbody.removeChild(tbody.firstChild);
  rows.forEach(r => tbody.appendChild(r));
  applyColorFormatting();
}}

function filterIntradayPicks() {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  const deliveryIndices = [];
  headers.forEach((h, i) => {{
    if (h.textContent.toLowerCase().includes('delivery')) deliveryIndices.push(i);
  }});
  if (deliveryIndices.length === 0) {{
    alert('No delivery columns found!');
    return;
  }}
  const targetIndex = deliveryIndices[deliveryIndices.length - 1];
  rows.forEach(row => {{
    const cell = row.cells[targetIndex];
    if (!cell) {{
      row.style.display = 'none';
      return;
    }}
    const raw = cell.textContent.trim().replace(/,/g,'').replace('%','');
    const val = parseFloat(raw);
    row.style.display = (!isNaN(val) && val > 70) ? '' : 'none';
  }});
  activeFilters.clear();
  applyColorFormatting();
}}

function downloadExcel() {{
  const table = document.getElementById("marketTable");
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(table);
  XLSX.utils.book_append_sheet(wb, ws, "IntradayData");
  XLSX.writeFile(wb, "IntradayMarketData.xlsx");
}}

function scrollTable(direction) {{
  const area = document.getElementById('tableScrollArea');
  if (!area) return;
  const rowHeight = 45;
  area.scrollBy({{ top: direction * rowHeight, behavior: 'smooth' }});
}}

function openTradingView() {{
  const stock = document.getElementById('stockSelect').value;
  if (!stock) {{
    alert('Please select a stock first.');
    return;
  }}
  // Open TradingView with the selected stock symbol (NSE exchange)
  const tradingViewUrl = 'https://www.tradingview.com/chart/?symbol=NSE:' + encodeURIComponent(stock);
  window.open(tradingViewUrl, '_blank');
}}

function toggleIndicatorMenu(event) {{
  const submenu = document.getElementById('indicatorSubmenu');
  const arrow = document.getElementById('indicatorArrow');
  if (submenu.style.display === 'none') {{
    submenu.style.display = 'block';
    arrow.style.transform = 'rotate(180deg)';
  }} else {{
    submenu.style.display = 'none';
    arrow.style.transform = 'rotate(0deg)';
  }}
}}

function showIndicatorOption(option) {{
  const messages = {{
    'candle': 'CANDLE COMBINATION indicator selected',
    'intraday': 'INTRADAY indicator selected',
    'options': 'OPTIONS TRADE indicator selected'
  }};
  alert(messages[option]);
}}

function updateMarketCapExtremes(stock, dates, marketCapData) {{
  if (!dates.length || !marketCapData.length) {{
    document.getElementById('highestMCValue').textContent = '-';
    document.getElementById('highestMCDate').textContent = '-';
    document.getElementById('lowestMCValue').textContent = '-';
    document.getElementById('lowestMCDate').textContent = '-';
    return;
  }}

  let highestVal = -Infinity;
  let lowestVal = Infinity;
  let highestDate = '';
  let lowestDate = '';

  for (let i = 0; i < marketCapData.length; i++) {{
    const val = marketCapData[i];
    const date = dates[i];
    if (val > highestVal) {{
      highestVal = val;
      highestDate = date;
    }}
    if (val < lowestVal && val > 0) {{
      lowestVal = val;
      lowestDate = date;
    }}
  }}

  document.getElementById('highestMCValue').textContent = highestVal > 0 ? highestVal.toLocaleString() + ' Cr' : '-';
  document.getElementById('highestMCDate').textContent = highestDate || '-';
  document.getElementById('lowestMCValue').textContent = lowestVal < Infinity ? lowestVal.toLocaleString() + ' Cr' : '-';
  document.getElementById('lowestMCDate').textContent = lowestDate || '-';
}}

// ---- STOCK ANALYSIS (Chart.js) ----
let stockChart;
function loadStockList() {{
  const table = document.getElementById('marketTable');
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  const select = document.getElementById('stockSelect');
  select.innerHTML = "";
  rows.forEach(row => {{
    const stock = row.cells[0].textContent.trim();
    if (stock) {{
      const option = document.createElement('option');
      option.value = stock;
      option.textContent = stock;
      select.appendChild(option);
    }}
  }});
}}

let lastStockSource = 'historyRank';

function showStockDetails(stockName, sourceTab) {{
  if (sourceTab) {{
    lastStockSource = sourceTab;
  }}
  const select = document.getElementById('stockSelect');
  if (select) {{
    select.value = stockName;
    const d = getStockAnalysisDuration();
    updateChart(d ? d.start : '', d ? d.end : '');
  }}
  const tabcontent = document.getElementsByClassName("tabcontent");
  for (let i = 0; i < tabcontent.length; i++) {{
    tabcontent[i].style.display = "none";
  }}
  const tablinks = document.getElementsByClassName("tablinks");
  for (let i = 0; i < tablinks.length; i++) {{
    tablinks[i].classList.remove("active");
  }}
  document.getElementById('stockAnalysis').style.display = "block";
  updateSidebarActive('stockAnalysis');
  toggleMenu(false);
  const dashboardStats = document.querySelector('.dashboard-stats');
  if (dashboardStats) dashboardStats.style.display = 'none';
}}

function updateChart(startDate, endDate) {{
  const stock = document.getElementById('stockSelect').value;
  const table = document.getElementById('marketTable');
  const headerNodes = Array.from(table.querySelectorAll('th'));
  const headers = headerNodes.map(h => h.textContent.trim());
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  let row;
  rows.forEach(r => {{
    if (r.cells[0].textContent.trim() === stock) row = r;
  }});
  if (!row) return;

  // Pattern to extract date at start like 2025-06-18
  const datePattern = /^\\d{{4}}-\\d{{2}}-\\d{{2}}/;

  // Build ordered list of unique dates found in headers (excluding the first header which is symbol name)
  const dates = [];
  const mapping = {{}}; // mapping[date] = {{ marketIdx: idx or null, deliveryIdx: idx or null, rankIdx: idx or null }}

  for (let i = 1; i < headers.length; i++) {{
    const h = headers[i];
    const m = h.match(datePattern);
    if (!m) continue;
    const d = m[0];
    if (!dates.includes(d)) dates.push(d);
    if (!mapping[d]) mapping[d] = {{ marketIdx: null, deliveryIdx: null, rankIdx: null }};
    const lower = h.toLowerCase();
    if (lower.includes('market')) {{
      mapping[d].marketIdx = i; // column index in table (th/td)
    }} else if (lower.includes('delivery')) {{
      mapping[d].deliveryIdx = i;
    }} else if (lower.includes('rank')) {{
      mapping[d].rankIdx = i;
    }} else {{
      // if header does not explicitly include 'market' or 'delivery', try to guess by keyword
      // default: do nothing
    }}
  }}

  // Filter dates by range if provided
  const filteredDates = dates.filter(d => {{
    if (startDate && d < startDate) return false;
    if (endDate && d > endDate) return false;
    return true;
  }});

  // Build arrays of data in the same order as filtered dates
  const marketCapData = [];
  const deliveryData = [];
  const rankData = [];
  const labels = filteredDates.slice(); // copy

  filteredDates.forEach(d => {{
    const map = mapping[d] || {{marketIdx:null, deliveryIdx:null, rankIdx:null}};
    let marketVal = 0;
    let deliveryVal = 0;
    let rankVal = 0;
    if (map.marketIdx !== null) {{
      const txt = (row.cells[map.marketIdx] && row.cells[map.marketIdx].textContent) ? row.cells[map.marketIdx].textContent.trim().replace(/,/g,'').replace('%','') : '';
      marketVal = parseFloat(txt) || 0;
    }}
    if (map.deliveryIdx !== null) {{
      const txt = (row.cells[map.deliveryIdx] && row.cells[map.deliveryIdx].textContent) ? row.cells[map.deliveryIdx].textContent.trim().replace(/,/g,'').replace('%','') : '';
      deliveryVal = parseFloat(txt) || 0;
    }}
    if (map.rankIdx !== null) {{
      const txt = (row.cells[map.rankIdx] && row.cells[map.rankIdx].textContent) ? row.cells[map.rankIdx].textContent.trim().replace(/,/g,'') : '';
      rankVal = parseFloat(txt) || 0;
    }}
    marketCapData.push(marketVal);
    deliveryData.push(deliveryVal);
    rankData.push(rankVal);
  }});

  // Find highest and lowest market cap dates
  updateMarketCapExtremes(stock, filteredDates, marketCapData);

  // Calculate market cap changes for coloring
  const marketCapColors = [];
  for (let i = 0; i < marketCapData.length; i++) {{
    if (i === 0) {{
      // First data point - no previous value to compare with
      marketCapColors.push('rgba(54, 162, 235, 0.6)'); // Default blue
    }} else {{
      const current = marketCapData[i];
      const previous = marketCapData[i-1];
      
      if (current > previous) {{
        marketCapColors.push('rgba(75, 192, 75, 0.6)'); // Green for increase
      }} else if (current < previous) {{
        marketCapColors.push('rgba(255, 99, 132, 0.6)'); // Red for decrease
      }} else {{
        marketCapColors.push('rgba(54, 162, 235, 0.6)'); // Blue for no change
      }}
    }}
  }}

  // Create Chart: Market Cap => bar, Delivery Qty => line (same date)
  if (stockChart) stockChart.destroy();
  const ctx = document.getElementById('stockChart').getContext('2d');

  stockChart = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: labels, // only dates displayed on x-axis
      datasets: [
        {{
          label: 'Market Cap',
          type: 'bar',
          data: marketCapData,
          backgroundColor: marketCapColors,
          borderColor: marketCapColors.map(color => color.replace('0.6', '1')),
          borderWidth: 1,
          yAxisID: 'y1'
        }},
        {{
          label: 'Delivery Qty',
          type: 'line',
          data: deliveryData,
          borderColor: 'rgba(153, 102, 255, 0.9)',
          backgroundColor: 'rgba(153, 102, 255, 0.3)',
          fill: false,
          tension: 0.3,
          yAxisID: 'y'
        }},
        {{
          label: 'Threshold (70)',
          type: 'line',
          data: Array(labels.length).fill(70),
          borderColor: 'rgba(239, 68, 68, 0.8)',
          borderWidth: 2,
          borderDash: [6, 4],
          pointRadius: 0,
          fill: false,
          yAxisID: 'y'
        }},
        {{
          label: 'Rank',
          type: 'line',
          data: rankData,
          borderColor: 'rgba(255, 159, 64, 0.9)',
          backgroundColor: 'rgba(255, 159, 64, 0.3)',
          fill: false,
          tension: 0.3,
          yAxisID: 'y2',
          pointRadius: 4,
          pointHoverRadius: 6
        }}
      ]
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      interaction: {{
        mode: 'index',
        intersect: false
      }},
      plugins: {{
        legend: {{
          position: 'top',
          labels: {{
            filter: function(legendItem, chartData) {{
              return chartData.datasets[legendItem.datasetIndex].label !== 'Threshold (70)';
            }}
          }}
        }},
        title: {{
          display: true,
          text: 'Market Cap (Bar) vs Delivery Qty (Line) vs Rank (Line) for ' + stock
        }}
      }},
      scales: {{
        x: {{
          title: {{
            display: true,
            text: 'Date'
          }}
        }},
        // Primary y for Delivery Qty (visible)
        y: {{
          beginAtZero: true,
          title: {{
            display: true,
            text: 'Delivery Qty'
          }},
          position: 'left',
          ticks: {{
            callback: function(value, index, ticks) {{
              return value;
            }}
          }}
        }},
        // Secondary y for Market Cap (hidden), used to plot bars without scaling the visible axis
        y1: {{
          beginAtZero: true,
          display: false,
          position: 'right'
        }},
        // Tertiary y for Rank (visible)
        y2: {{
          beginAtZero: true,
          reverse: true,
          title: {{
            display: true,
            text: 'Rank'
          }},
          position: 'right',
          grid: {{
            drawOnChartArea: false
          }},
          ticks: {{
            callback: function(value, index, ticks) {{
              return value;
            }}
          }}
        }}
      }}
    }}
  }});
}}

function getStockAnalysisDuration() {{
  const startInput = document.getElementById('stockAnalysisStartDate');
  const endInput = document.getElementById('stockAnalysisEndDate');
  const start = startInput ? startInput.value : '';
  const end = endInput ? endInput.value : '';
  if (start && end && start > end) {{
    alert('Start date cannot be after end date.');
    return null;
  }}
  return {{ start, end }};
}}

function filterStockAnalysisByDuration() {{
  const duration = getStockAnalysisDuration();
  if (!duration) return;
  updateChart(duration.start, duration.end);
}}

function clearStockAnalysisDuration() {{
  const startInput = document.getElementById('stockAnalysisStartDate');
  const endInput = document.getElementById('stockAnalysisEndDate');
  if (startInput) startInput.value = '';
  if (endInput) endInput.value = '';
  updateChart('', '');
}}

// ---- MARKET OVERVIEW ----
let marketOverviewChart;
let rankTooltip = null;

function showTop50MarketCap() {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  
  // Find the latest market cap column
  let marketCapColIndex = -1;
  let marketCapDate = '';
  const datePattern = /^\\d{{4}}-\\d{{2}}-\\d{{2}}/;
  
  // Look for market cap columns from right to left (latest first)
  for (let i = headers.length - 1; i >= 1; i--) {{
    const headerText = headers[i].textContent.trim().toLowerCase();
    if (headerText.includes('market')) {{
      marketCapColIndex = i;
      const dateMatch = headers[i].textContent.trim().match(datePattern);
      marketCapDate = dateMatch ? dateMatch[0] : 'Latest';
      break;
    }}
  }}
  
  if (marketCapColIndex === -1) {{
    alert('No market cap columns found!');
    return;
  }}
  
  // Find the previous market cap column for comparison
  let prevMarketCapColIndex = -1;
  for (let i = marketCapColIndex - 1; i >= 1; i--) {{
    const headerText = headers[i].textContent.trim().toLowerCase();
    if (headerText.includes('market')) {{
      prevMarketCapColIndex = i;
      break;
    }}
  }}
  
  // Find all rank columns and get the last 5
  const allRankColumns = [];
  for (let i = 1; i < headers.length; i++) {{
    const headerText = headers[i].textContent.trim().toLowerCase();
    if (headerText.includes('rank')) {{
      const dateMatch = headers[i].textContent.trim().match(datePattern);
      const date = dateMatch ? dateMatch[0] : 'Unknown';
      allRankColumns.push({{ index: i, date: date }});
    }}
  }}
  
  // Sort rank columns by date (newest first) and take the last 5
  allRankColumns.sort((a, b) => new Date(b.date) - new Date(a.date));
  const last5RankColumns = allRankColumns.slice(0, 5);
  
  // Collect stock data with market cap values
  const stockData = [];
  rows.forEach(row => {{
    const stockName = row.cells[0].textContent.trim();
    const marketCapText = row.cells[marketCapColIndex]?.textContent.trim().replace(/,/g, '') || '';
    const marketCap = parseFloat(marketCapText);
    
    // Get rank values for the last 5 days
    const rankHistory = [];
    last5RankColumns.forEach(rankCol => {{
      const rankValue = row.cells[rankCol.index]?.textContent.trim() || '';
      rankHistory.push({{ date: rankCol.date, rank: rankValue }});
    }});
    
    if (!isNaN(marketCap) && marketCap > 0) {{
      let prevMarketCap = null;
      if (prevMarketCapColIndex !== -1) {{
        const prevMarketCapText = row.cells[prevMarketCapColIndex]?.textContent.trim().replace(/,/g, '') || '';
        prevMarketCap = parseFloat(prevMarketCapText);
      }}
      
      stockData.push({{
        name: stockName,
        marketCap: marketCap,
        prevMarketCap: prevMarketCap,
        rankHistory: rankHistory
      }});
    }}
  }});
  
  // Sort by market cap (descending) and take top 50
  stockData.sort((a, b) => b.marketCap - a.marketCap);
  const top50 = stockData.slice(0, 50);
  
  // Prepare data for chart
  const stockNames = top50.map(item => item.name);
  const marketCaps = top50.map(item => item.marketCap);
  
  // Determine colors based on market cap change
  const colors = top50.map(item => {{
    if (item.prevMarketCap === null || isNaN(item.prevMarketCap)) {{
      return 'rgba(54, 162, 235, 0.6)'; // Blue for no previous data
    }} else if (item.marketCap > item.prevMarketCap) {{
      return 'rgba(75, 192, 75, 0.6)'; // Green for increase
    }} else if (item.marketCap < item.prevMarketCap) {{
      return 'rgba(255, 99, 132, 0.6)'; // Red for decrease
    }} else {{
      return 'rgba(54, 162, 235, 0.6)'; // Blue for no change
    }}
  }});
  
  // Create the chart
  if (marketOverviewChart) marketOverviewChart.destroy();
  const ctx = document.getElementById('marketOverviewChart').getContext('2d');
  
  marketOverviewChart = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: stockNames,
      datasets: [{{
        label: 'Market Capitalization',
        data: marketCaps,
        backgroundColor: colors,
        borderColor: colors.map(color => color.replace('0.6', '1')),
        borderWidth: 1
      }}]
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{ position: 'top' }},
        title: {{
          display: true,
          text: 'Top 50 Stocks by Market Capitalization (' + marketCapDate + ')'
        }},
        tooltip: {{
          enabled: false,
          external: function(context) {{
            // Tooltip handling for rank history
            const tooltipEl = document.getElementById('rank-tooltip');
            
            // Create tooltip if it doesn't exist
            if (!tooltipEl) {{
              const newTooltip = document.createElement('div');
              newTooltip.id = 'rank-tooltip';
              newTooltip.className = 'rank-tooltip';
              document.body.appendChild(newTooltip);
            }}
            
            const tooltip = document.getElementById('rank-tooltip');
            
            // Hide if no tooltip
            if (context.tooltip.opacity === 0) {{
              tooltip.style.opacity = '0';
              return;
            }}
            
            // Set Text
            if (context.tooltip.dataPoints && context.tooltip.dataPoints.length > 0) {{
              const dataIndex = context.tooltip.dataPoints[0].dataIndex;
              const stock = top50[dataIndex];
              
              let html = '<div><strong>' + stock.name + '</strong></div>';
              html += '<div>Market Cap: ' + stock.marketCap.toLocaleString() + '</div>';
              html += '<div style="margin-top: 5px;"><strong>Rank History:</strong></div>';
              
              stock.rankHistory.forEach(rank => {{
                html += '<div>' + rank.date + ': ' + (rank.rank || 'N/A') + '</div>';
              }});
              
              tooltip.innerHTML = html;
            }}
            
            // Position tooltip
            const position = context.chart.canvas.getBoundingClientRect();
            tooltip.style.opacity = '1';
            tooltip.style.left = position.left + window.pageXOffset + context.tooltip.caretX + 'px';
            tooltip.style.top = position.top + window.pageYOffset + context.tooltip.caretY + 'px';
          }}
        }}
      }},
      scales: {{
        x: {{
          title: {{
            display: true,
            text: 'Stock Name'
          }},
          ticks: {{
            maxRotation: 45,
            minRotation: 45
          }}
        }},
        y: {{
          beginAtZero: true,
          title: {{
            display: true,
            text: 'Market Capitalization'
          }}
        }}
      }}
    }}
  }});
  
  // Update information panel with rank details
  const infoPanel = document.getElementById('marketOverviewInfo');
  infoPanel.innerHTML = '<h4>Rank Details (Top 50 Stocks)</h4>';
  infoPanel.innerHTML += '<p>Hover over bars to see rank history for the last 5 days</p>';
  
  // Add rank summary by date
  const rankSummary = {{}};
  last5RankColumns.forEach(rankCol => {{
    rankSummary[rankCol.date] = {{}};
  }});
  
  top50.forEach(stock => {{
    stock.rankHistory.forEach(rank => {{
      if (rank.rank && rankSummary[rank.date]) {{
        if (!rankSummary[rank.date][rank.rank]) {{
          rankSummary[rank.date][rank.rank] = [];
        }}
        rankSummary[rank.date][rank.rank].push(stock.name);
      }}
    }});
  }});
  
  // Display rank summary for each date
  for (const [date, ranks] of Object.entries(rankSummary)) {{
    const dateDiv = document.createElement('div');
    dateDiv.className = 'info-item';
    dateDiv.innerHTML = '<strong>' + date + ' Rank Summary:</strong>';
    
    const rankList = document.createElement('div');
    rankList.className = 'rank-history';
    
    // Sort ranks numerically
    const sortedRanks = Object.keys(ranks).sort((a, b) => parseInt(a) - parseInt(b));
    
    for (const rank of sortedRanks) {{
      const stocks = ranks[rank];
      const rankEntry = document.createElement('div');
      rankEntry.className = 'rank-entry';
      rankEntry.innerHTML = '<span class="rank-label">Rank ' + rank + ':</span> ' + stocks.join(', ');
      rankList.appendChild(rankEntry);
    }}
    
    dateDiv.appendChild(rankList);
    infoPanel.appendChild(dateDiv);
  }}
}}

function showTopPerformers() {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const rows = Array.from(table.querySelectorAll('tbody tr'));

  // Find the latest delivery column
  let deliveryColIndex = -1;
  for (let i = headers.length - 1; i >= 1; i--) {{
    const headerText = headers[i].textContent.trim().toLowerCase();
    if (headerText.includes('delivery')) {{
      deliveryColIndex = i;
      break;
    }}
  }}

  if (deliveryColIndex === -1) {{
    alert('No delivery columns found!');
    return;
  }}

  // Collect stock data with delivery values
  const stockData = [];
  rows.forEach(row => {{
    const stockName = row.cells[0].textContent.trim();
    const deliveryText = row.cells[deliveryColIndex]?.textContent.trim().replace(/,/g, '').replace('%', '') || '';
    const delivery = parseFloat(deliveryText);
    if (!isNaN(delivery) && delivery > 0) {{
      stockData.push({{ name: stockName, delivery: delivery }});
    }}
  }});

  // Sort by delivery descending and take top 10
  stockData.sort((a, b) => b.delivery - a.delivery);
  const top10 = stockData.slice(0, 10);

  // Display in info panel
  const infoPanel = document.getElementById('topPerformersList');
  infoPanel.innerHTML = '<h4>Top 10 Stocks by Delivery Percentage</h4>';
  const list = document.createElement('ol');
  top10.forEach((stock, index) => {{
    const item = document.createElement('li');
    const nameSpan = document.createElement('span');
    nameSpan.textContent = stock.name;
    nameSpan.style.fontWeight = '700';
    nameSpan.style.color = 'var(--primary)';
    nameSpan.style.cursor = 'pointer';
    nameSpan.style.textDecoration = 'underline';
    nameSpan.onclick = function() {{ showStockDetails(stock.name, 'topPerformers'); }};
    const deliveryText = document.createTextNode(' - ' + stock.delivery.toFixed(2) + '%');
    item.appendChild(nameSpan);
    item.appendChild(deliveryText);
    list.appendChild(item);
  }});
  infoPanel.appendChild(list);
}}

function showTopPerformersDuration(startDate, endDate) {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const rows = Array.from(table.querySelectorAll('tbody tr'));

  const allDeliveryIndices = [];
  const deliveryDateMap = {{}};
  headers.forEach((h, i) => {{
    if (i > 0 && h.textContent.trim().toLowerCase().includes('delivery')) {{
      allDeliveryIndices.push(i);
      const headerText = h.textContent.trim();
      const dateMatch = headerText.match(/(\\d{{4}}-\\d{{2}}-\\d{{2}})/);
      if (dateMatch) {{
        deliveryDateMap[i] = dateMatch[1];
      }}
    }}
  }});

  const stockDeliveryCounts = [];
  rows.forEach(row => {{
    const stockName = row.cells[0].textContent.trim();
    let count = 0;
    allDeliveryIndices.forEach(colIndex => {{
      const colDate = deliveryDateMap[colIndex];
      if (startDate && colDate && colDate < startDate) return;
      if (endDate && colDate && colDate > endDate) return;
      const txt = row.cells[colIndex]?.textContent.trim().replace(/,/g, '').replace('%', '') || '';
      const val = parseFloat(txt);
      if (!isNaN(val) && val > 70) count++;
    }});
    if (count > 0) {{
      stockDeliveryCounts.push({{ name: stockName, count: count }});
    }}
  }});

  stockDeliveryCounts.sort((a, b) => b.count - a.count);
  const top10counts = stockDeliveryCounts.slice(0, 10);

  const chartContainer = document.getElementById('topPerformersChartContainer');
  if (top10counts.length > 0) {{
    chartContainer.style.display = 'block';
    if (topPerformersChart) topPerformersChart.destroy();
    const ctx = document.getElementById('topPerformersChart').getContext('2d');
    topPerformersChart = new Chart(ctx, {{
      type: 'bar',
      data: {{
        labels: top10counts.map(item => item.name),
        datasets: [{{
          label: 'Dates with Delivery > 70%',
          data: top10counts.map(item => item.count),
          backgroundColor: [
            'rgba(102, 126, 234, 0.8)',
            'rgba(240, 147, 251, 0.8)',
            'rgba(79, 172, 254, 0.8)',
            'rgba(67, 233, 123, 0.8)',
            'rgba(250, 112, 154, 0.8)',
            'rgba(255, 159, 64, 0.8)',
            'rgba(54, 162, 235, 0.8)',
            'rgba(255, 99, 132, 0.8)',
            'rgba(153, 102, 255, 0.8)',
            'rgba(255, 206, 86, 0.8)'
          ],
          borderColor: [
            'rgba(102, 126, 234, 1)',
            'rgba(240, 147, 251, 1)',
            'rgba(79, 172, 254, 1)',
            'rgba(67, 233, 123, 1)',
            'rgba(250, 112, 154, 1)',
            'rgba(255, 159, 64, 1)',
            'rgba(54, 162, 235, 1)',
            'rgba(255, 99, 132, 1)',
            'rgba(153, 102, 255, 1)',
            'rgba(255, 206, 86, 1)'
          ],
          borderWidth: 2,
          borderRadius: 8
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        onClick: (event, elements, chart) => {{
          if (elements.length > 0) {{
            const index = elements[0].index;
            const label = chart.data.labels[index];
            if (label) {{
              showStockDetails(label, 'topPerformers');
            }}
          }}
        }},
        onHover: (event, chartElement, chart) => {{
          event.native.target.style.cursor = chartElement.length ? 'pointer' : 'default';
        }},
        plugins: {{
          legend: {{ display: false }},
          title: {{
            display: true,
            text: startDate && endDate
              ? '🏆 Top 10 by Delivery Qty > 70, Dates Count (' + startDate + ' to ' + endDate + ')'
              : '🏆 Top 10 by Delivery Qty > 70, Dates Count',
            font: {{ size: 16, weight: 'bold' }},
            color: '#1e293b'
          }}
        }},
        scales: {{
          y: {{
            beginAtZero: true,
            ticks: {{
              stepSize: 1,
              color: '#64748b'
            }},
            grid: {{
              color: '#f1f5f9'
            }}
          }},
          x: {{
            ticks: {{
              color: '#64748b',
              font: {{ weight: '600' }},
              maxRotation: 45,
              minRotation: 45
            }},
            grid: {{
              display: false
            }}
          }}
        }}
      }}
    }});
  }} else {{
    chartContainer.style.display = 'none';
  }}
}}

function switchTopPerformersTab(mode) {{
  const regularControl = document.getElementById('topPerformersRegularControl');
  const durationControl = document.getElementById('topPerformersDurationControl');
  const regularTab = document.getElementById('topPerformersRegularTab');
  const durationTab = document.getElementById('topPerformersDurationTab');
  const chartContainer = document.getElementById('topPerformersChartContainer');

  if (!regularControl || !durationControl || !regularTab || !durationTab) return;

  if (mode === 'regular') {{
    regularControl.style.display = 'block';
    durationControl.style.display = 'none';
    if (chartContainer) chartContainer.style.display = 'none';
    regularTab.style.background = 'white';
    regularTab.style.color = 'var(--primary)';
    regularTab.style.borderBottom = '3px solid var(--accent)';
    durationTab.style.background = '#f1f5f9';
    durationTab.style.color = 'var(--text-light)';
    durationTab.style.borderBottom = '3px solid transparent';
  }} else {{
    regularControl.style.display = 'none';
    durationControl.style.display = 'block';
    regularTab.style.background = '#f1f5f9';
    regularTab.style.color = 'var(--text-light)';
    regularTab.style.borderBottom = '3px solid transparent';
    durationTab.style.background = 'white';
    durationTab.style.color = 'var(--primary)';
    durationTab.style.borderBottom = '3px solid var(--accent)';
  }}
}}

function getTopPerformersDuration() {{
  const startInput = document.getElementById('topPerformersStartDate');
  const endInput = document.getElementById('topPerformersEndDate');
  const start = startInput ? startInput.value : '';
  const end = endInput ? endInput.value : '';
  if (start && end && start > end) {{
    alert('Start date cannot be after end date.');
    return null;
  }}
  return {{ start, end }};
}}

function filterTopPerformersByDuration() {{
  const duration = getTopPerformersDuration();
  if (!duration) return;
  showTopPerformersDuration(duration.start, duration.end);
}}

function clearTopPerformersDuration() {{
  const startInput = document.getElementById('topPerformersStartDate');
  const endInput = document.getElementById('topPerformersEndDate');
  if (startInput) startInput.value = '';
  if (endInput) endInput.value = '';
  showTopPerformersDuration('', '');
}}

function showRankChanges() {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const rows = Array.from(table.querySelectorAll('tbody tr'));

  // Find all rank columns
  const rankColumns = [];
  headers.forEach((header, index) => {{
    if (index > 0 && header.textContent.trim().toLowerCase().includes('rank')) {{
      rankColumns.push(index);
    }}
  }});

  if (rankColumns.length < 2) {{
    alert('Need at least 2 rank columns for comparison!');
    return;
  }}

  // Take the last two (most recent)
  const latestRankIndex = rankColumns[rankColumns.length - 1];
  const previousRankIndex = rankColumns[rankColumns.length - 2];

  // Collect stock data with rank changes
  const stockData = [];
  rows.forEach(row => {{
    const stockName = row.cells[0].textContent.trim();
    const latestRankText = row.cells[latestRankIndex]?.textContent.trim().replace(/[^0-9]/g, '') || '';
    const previousRankText = row.cells[previousRankIndex]?.textContent.trim().replace(/[^0-9]/g, '') || '';
    const latestRank = parseInt(latestRankText, 10);
    const previousRank = parseInt(previousRankText, 10);

    if (!isNaN(latestRank) && !isNaN(previousRank)) {{
      const change = previousRank - latestRank; // Positive means improvement
      if (change >= 10) {{
        stockData.push({{ name: stockName, previousRank: previousRank, latestRank: latestRank, change: change }});
      }}
    }}
  }});

  // Sort by largest improvement (highest change) and take top 10
  stockData.sort((a, b) => b.change - a.change);
  const topImprovers = stockData.slice(0, 10);

  // Display in info panel
  const infoPanel = document.getElementById('rankChangesList');
  infoPanel.innerHTML = '<h4>Top 10 Rank Improvers (Yesterday vs Today)</h4>';
  const list = document.createElement('ol');
  topImprovers.forEach((stock, index) => {{
    const item = document.createElement('li');
    item.innerHTML = `<strong>${{stock.name}}</strong> - Yesterday Rank: ${{Math.floor(stock.previousRank / 10)}}, Today Rank: ${{Math.floor(stock.latestRank / 10)}}`;
    list.appendChild(item);
  }});
  infoPanel.appendChild(list);
}}

let historyRankData = null;

function showHistoryRank() {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const rows = Array.from(table.querySelectorAll('tbody tr'));
  const datePattern = /^\\d{{4}}-\\d{{2}}-\\d{{2}}/;

  // Collect all rank columns with their dates
  const rankColumns = [];
  headers.forEach((header, index) => {{
    if (index === 0) return;
    const text = header.textContent.trim();
    const match = text.match(datePattern);
    if (match && text.toLowerCase().includes('rank')) {{
      rankColumns.push({{ index: index, date: match[0] }});
    }}
  }});

  if (rankColumns.length < 2) {{
    alert('Need at least 2 rank date columns for history!');
    return;
  }}

  // Sort rank columns by date
  rankColumns.sort((a, b) => new Date(a.date) - new Date(b.date));

  // Pre-compute top 10 improvers for each date transition
  const dailyTop10Map = {{}};
  const dailyTop10Stocks = {{}};
  for (let i = 1; i < rankColumns.length; i++) {{
    const prevCol = rankColumns[i - 1];
    const currCol = rankColumns[i];
    const stockData = [];
    rows.forEach(row => {{
      const stockName = row.cells[0].textContent.trim();
      const prevRankText = row.cells[prevCol.index]?.textContent.trim().replace(/[^0-9]/g, '') || '';
      const currRankText = row.cells[currCol.index]?.textContent.trim().replace(/[^0-9]/g, '') || '';
      const prevRank = parseInt(prevRankText, 10);
      const currRank = parseInt(currRankText, 10);

      if (!isNaN(prevRank) && !isNaN(currRank) && prevRank > 0 && currRank > 0) {{
        const improvement = prevRank - currRank;
        if (improvement > 0) {{
          stockData.push({{
            name: stockName,
            prevRank: prevRank,
            currRank: currRank,
            improvement: improvement
          }});
        }}
      }}
    }});

    stockData.sort((a, b) => b.improvement - a.improvement);
    const top10 = stockData.slice(0, 10);
    dailyTop10Map[currCol.date] = top10;
    top10.forEach(s => {{
      if (!dailyTop10Stocks[s.name]) dailyTop10Stocks[s.name] = [];
      dailyTop10Stocks[s.name].push(currCol.date);
    }});
  }}

  historyRankData = {{
    dailyTop10Map: dailyTop10Map,
    dailyTop10Stocks: dailyTop10Stocks,
    rankColumns: rankColumns
  }};

  // Populate date picker range
  const datePicker = document.getElementById('historyRankDatePicker');
  datePicker.value = '';
  if (rankColumns.length >= 2) {{
    const minDate = rankColumns[1].date;
    const maxDate = rankColumns[rankColumns.length - 1].date;
    datePicker.min = minDate;
    datePicker.max = maxDate;
  }}

  const duration = getHistoryRankDuration();
  if (duration) {{
    renderHistoryRankAllDates(duration.start, duration.end);
  }} else {{
    renderHistoryRankAllDates();
  }}
}}

function renderHistoryRankAllDates(startDate, endDate) {{
  if (!historyRankData) return;
  const {{ dailyTop10Map, dailyTop10Stocks, rankColumns }} = historyRankData;

  const infoPanel = document.getElementById('historyRankList');
  infoPanel.innerHTML = '<h4>Top 10 Rank Improvers by Date Transition</h4>';

  const counts = historyRankData.dailyTop10Stocks || {{}};
  const top5 = Object.entries(counts)
    .map(([name, dates]) => ({{
      name,
      count: dates.filter(d => {{
        if (startDate && d < startDate) return false;
        if (endDate && d > endDate) return false;
        return true;
      }}).length
    }}))
    .sort((a, b) => b.count - a.count)
    .slice(0, 5);

  if (top5.length > 0) {{
    const top5Section = document.createElement('div');
    top5Section.className = 'top5-section';
    top5Section.style.marginBottom = '24px';
    top5Section.style.padding = '22px';
    top5Section.style.background = 'linear-gradient(135deg, #e8f5e9 0%, #f1f8f2 100%)';
    top5Section.style.borderRadius = '14px';
    top5Section.style.borderLeft = '5px solid #10b981';
    top5Section.style.boxShadow = '0 4px 20px rgba(16, 185, 129, 0.12)';

    const top5Title = document.createElement('h5');
    top5Title.textContent = '🏆 Top 5 by Count of Dates';
    top5Title.style.marginBottom = '14px';
    top5Title.style.color = '#059669';
    top5Title.style.fontFamily = "'Montserrat', sans-serif";
    top5Title.style.fontWeight = '700';
    top5Title.style.fontSize = '1.1rem';
    top5Section.appendChild(top5Title);

    top5.forEach(item => {{
      const btn = document.createElement('button');
      btn.textContent = item.name + ' (' + item.count + ')';
      btn.style.margin = '6px 12px 6px 0';
      btn.style.padding = '10px 18px';
      btn.style.border = 'none';
      btn.style.borderRadius = '8px';
      btn.style.background = 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)';
      btn.style.color = 'white';
      btn.style.cursor = 'pointer';
      btn.style.fontSize = '0.88rem';
      btn.style.fontWeight = '600';
      btn.style.boxShadow = '0 4px 12px rgba(102, 126, 234, 0.3)';
      btn.style.transition = 'all 0.3s';
      btn.onmouseenter = function() {{ this.style.transform = 'translateY(-2px)'; this.style.boxShadow = '0 8px 25px rgba(102, 126, 234, 0.4)'; }};
      btn.onmouseleave = function() {{ this.style.transform = 'translateY(0)'; this.style.boxShadow = '0 4px 12px rgba(102, 126, 234, 0.3)'; }};
      btn.onclick = function() {{ showStockDetails(item.name, 'historyRank'); }};
      top5Section.appendChild(btn);
    }});

    infoPanel.appendChild(top5Section);
  }}

  for (let i = rankColumns.length - 1; i >= 1; i--) {{
    const prevCol = rankColumns[i - 1];
    const currCol = rankColumns[i];
    const dateLabel = currCol.date;
    const top10 = dailyTop10Map[currCol.date] || [];

    const section = document.createElement('div');
    section.style.marginBottom = '20px';
    section.style.padding = '18px';
    section.style.background = 'linear-gradient(135deg, #f8fafc 0%, #ffffff 100%)';
    section.style.borderRadius = '12px';
    section.style.borderLeft = '5px solid var(--accent)';

    const title = document.createElement('h5');
    title.textContent = dateLabel;
    title.style.marginBottom = '10px';
    title.style.color = 'var(--secondary)';
    section.appendChild(title);

    if (top10.length === 0) {{
      const noData = document.createElement('p');
      noData.textContent = 'No rank improvements found for this period.';
      noData.style.color = 'var(--text-light)';
      section.appendChild(noData);
    }} else {{
      const tableEl = document.createElement('table');
      tableEl.style.width = '100%';
      tableEl.style.borderCollapse = 'separate';
      tableEl.style.borderSpacing = '0';
      tableEl.style.fontSize = '0.92rem';
      tableEl.style.borderRadius = '12px';
      tableEl.style.overflow = 'hidden';
      tableEl.style.boxShadow = '0 4px 20px rgba(0,0,0,0.06)';

      const thead = document.createElement('thead');
      const headerRow = document.createElement('tr');
      ['Rank', 'Stock', 'Also In Top 10 Dates', 'Count of Dates', 'Show Details', 'Previous Rank', 'Current Rank', 'Improvement'].forEach(text => {{
        const th = document.createElement('th');
        th.textContent = text;
        th.style.padding = '14px 16px';
        th.style.background = 'linear-gradient(135deg, #1e293b 0%, #334155 100%)';
        th.style.color = 'white';
        th.style.textAlign = 'left';
        th.style.fontFamily = "'Montserrat', sans-serif";
        th.style.fontWeight = '700';
        th.style.fontSize = '0.82rem';
        th.style.textTransform = 'uppercase';
        th.style.letterSpacing = '0.6px';
        th.style.borderBottom = '2px solid rgba(255,255,255,0.1)';
        headerRow.appendChild(th);
      }});
      thead.appendChild(headerRow);
      tableEl.appendChild(thead);

      const tbody = document.createElement('tbody');
      top10.forEach((stock, idx) => {{
        const tr = document.createElement('tr');
        tr.style.transition = 'all 0.2s';
        if (idx % 2 === 1) tr.style.background = '#f8fafc';

        const prevDates = (dailyTop10Stocks[stock.name] || [])
          .filter(d => d !== currCol.date)
          .filter(d => {{
            if (startDate && d < startDate) return false;
            if (endDate && d > endDate) return false;
            return true;
          }})
          .sort((a, b) => new Date(b) - new Date(a));

        const tdRank = document.createElement('td');
        tdRank.textContent = idx + 1;
        tdRank.style.padding = '14px 16px';
        tdRank.style.fontWeight = '700';
        tdRank.style.color = 'var(--primary)';
        tdRank.style.textAlign = 'center';

        const tdName = document.createElement('td');
        tdName.textContent = stock.name;
        tdName.style.padding = '14px 16px';
        tdName.style.fontWeight = '600';

        const tdDates = document.createElement('td');
        tdDates.textContent = prevDates.length > 0 ? prevDates.join(', ') : '-';
        tdDates.style.padding = '14px 16px';
        tdDates.style.color = '#e65100';
        tdDates.style.fontWeight = '500';
        tdDates.style.fontSize = '0.85rem';

        const tdCount = document.createElement('td');
        tdCount.textContent = prevDates.length;
        tdCount.style.padding = '14px 16px';
        tdCount.style.fontWeight = '700';
        tdCount.style.textAlign = 'center';
        tdCount.style.color = 'var(--primary)';

        const tdAction = document.createElement('td');
        tdAction.style.padding = '14px 16px';
        const showBtn = document.createElement('button');
        showBtn.textContent = 'Show Details';
        showBtn.style.padding = '8px 16px';
        showBtn.style.border = 'none';
        showBtn.style.borderRadius = '8px';
        showBtn.style.background = 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)';
        showBtn.style.color = 'white';
        showBtn.style.cursor = 'pointer';
        showBtn.style.fontSize = '0.82rem';
        showBtn.style.fontWeight = '600';
        showBtn.style.boxShadow = '0 4px 12px rgba(102, 126, 234, 0.3)';
        showBtn.style.transition = 'all 0.3s';
        showBtn.onmouseenter = function() {{ this.style.transform = 'translateY(-1px)'; this.style.boxShadow = '0 6px 20px rgba(102, 126, 234, 0.4)'; }};
        showBtn.onmouseleave = function() {{ this.style.transform = 'translateY(0)'; this.style.boxShadow = '0 4px 12px rgba(102, 126, 234, 0.3)'; }};
        showBtn.onclick = function() {{ showStockDetails(stock.name, 'historyRank'); }};
        tdAction.appendChild(showBtn);

        const tdPrev = document.createElement('td');
        tdPrev.textContent = stock.prevRank / 10;
        tdPrev.style.padding = '14px 16px';
        tdPrev.style.textAlign = 'center';

        const tdCurr = document.createElement('td');
        tdCurr.textContent = stock.currRank / 10;
        tdCurr.style.padding = '14px 16px';
        tdCurr.style.color = 'var(--success)';
        tdCurr.style.fontWeight = '700';
        tdCurr.style.textAlign = 'center';

        const tdImprove = document.createElement('td');
        tdImprove.textContent = '+' + (stock.improvement / 10);
        tdImprove.style.padding = '14px 16px';
        tdImprove.style.color = 'var(--success)';
        tdImprove.style.fontWeight = '700';
        tdImprove.style.textAlign = 'center';

        tr.appendChild(tdRank);
        tr.appendChild(tdName);
        tr.appendChild(tdDates);
        tr.appendChild(tdCount);
        tr.appendChild(tdAction);
        tr.appendChild(tdPrev);
        tr.appendChild(tdCurr);
        tr.appendChild(tdImprove);
        tbody.appendChild(tr);
      }});
      tableEl.appendChild(tbody);
      section.appendChild(tableEl);
    }}

    infoPanel.appendChild(section);
  }}
}}

function switchHistoryRankTab(mode) {{
  const singleControl = document.getElementById('historyRankSingleControl');
  const durationControl = document.getElementById('historyRankDurationControl');
  const singleTab = document.getElementById('historyRankSingleTab');
  const durationTab = document.getElementById('historyRankDurationTab');

  if (!singleControl || !durationControl || !singleTab || !durationTab) return;

  if (mode === 'single') {{
    singleControl.style.display = 'block';
    durationControl.style.display = 'none';
    singleTab.style.background = 'white';
    singleTab.style.color = 'var(--primary)';
    singleTab.style.borderBottom = '3px solid var(--accent)';
    durationTab.style.background = '#f1f5f9';
    durationTab.style.color = 'var(--text-light)';
    durationTab.style.borderBottom = '3px solid transparent';
  }} else {{
    singleControl.style.display = 'none';
    durationControl.style.display = 'block';
    singleTab.style.background = '#f1f5f9';
    singleTab.style.color = 'var(--text-light)';
    singleTab.style.borderBottom = '3px solid transparent';
    durationTab.style.background = 'white';
    durationTab.style.color = 'var(--primary)';
    durationTab.style.borderBottom = '3px solid var(--accent)';
  }}
}}

function showHistoryRankByDate() {{
  const datePicker = document.getElementById('historyRankDatePicker');
  const selectedDate = datePicker.value;
  if (!selectedDate) {{
    const duration = getHistoryRankDuration();
    if (duration) {{
      renderHistoryRankAllDates(duration.start, duration.end);
    }} else {{
      renderHistoryRankAllDates();
    }}
    return;
  }}
  if (!historyRankData) {{
    const duration = getHistoryRankDuration();
    if (duration) {{
      renderHistoryRankAllDates(duration.start, duration.end);
    }} else {{
      renderHistoryRankAllDates();
    }}
    return;
  }}
  const validDates = historyRankData.rankColumns.slice(1).map(r => r.date);
  if (!validDates.includes(selectedDate)) {{
    alert('Selected date is not available in rank history.');
    return;
  }}
  const duration = getHistoryRankDuration();
  if (duration) {{
    renderHistoryRankForDate(selectedDate, duration.start, duration.end);
  }} else {{
    renderHistoryRankForDate(selectedDate);
  }}
}}

function clearHistoryRankDate() {{
  const datePicker = document.getElementById('historyRankDatePicker');
  datePicker.value = '';
  const duration = getHistoryRankDuration();
  if (duration) {{
    renderHistoryRankAllDates(duration.start, duration.end);
  }} else {{
    renderHistoryRankAllDates();
  }}
}}

function getHistoryRankDuration() {{
  const startInput = document.getElementById('historyRankStartDate');
  const endInput = document.getElementById('historyRankEndDate');
  const start = startInput ? startInput.value : '';
  const end = endInput ? endInput.value : '';
  if (start && end && start > end) {{
    alert('Start date cannot be after end date.');
    return null;
  }}
  return {{ start, end }};
}}

function filterHistoryRankByDuration() {{
  const duration = getHistoryRankDuration();
  if (!duration) return;
  const datePicker = document.getElementById('historyRankDatePicker');
  const selectedDate = datePicker.value;
  if (selectedDate) {{
    renderHistoryRankForDate(selectedDate, duration.start, duration.end);
  }} else {{
    renderHistoryRankAllDates(duration.start, duration.end);
  }}
}}

function clearHistoryRankDuration() {{
  const startInput = document.getElementById('historyRankStartDate');
  const endInput = document.getElementById('historyRankEndDate');
  if (startInput) startInput.value = '';
  if (endInput) endInput.value = '';
  const datePicker = document.getElementById('historyRankDatePicker');
  const selectedDate = datePicker.value;
  if (selectedDate) {{
    renderHistoryRankForDate(selectedDate);
  }} else {{
    renderHistoryRankAllDates();
  }}
}}

function renderHistoryRankForDate(selectedDate, startDate, endDate) {{
  if (!historyRankData) return;
  const {{ dailyTop10Map, dailyTop10Stocks, rankColumns }} = historyRankData;
  const top10 = dailyTop10Map[selectedDate] || [];
  const currCol = rankColumns.find(r => r.date === selectedDate);

  const infoPanel = document.getElementById('historyRankList');
  infoPanel.innerHTML = '<h4>Top 10 Rank Improvers - ' + selectedDate + '</h4>';

  const counts = historyRankData.dailyTop10Stocks || {{}};
  const top5 = Object.entries(counts)
    .map(([name, dates]) => ({{
      name,
      count: dates.filter(d => {{
        if (startDate && d < startDate) return false;
        if (endDate && d > endDate) return false;
        return true;
      }}).length
    }}))
    .sort((a, b) => b.count - a.count)
    .slice(0, 5);

  if (top5.length > 0) {{
    const top5Section = document.createElement('div');
    top5Section.className = 'top5-section';
    top5Section.style.marginBottom = '24px';
    top5Section.style.padding = '22px';
    top5Section.style.background = 'linear-gradient(135deg, #e8f5e9 0%, #f1f8f2 100%)';
    top5Section.style.borderRadius = '14px';
    top5Section.style.borderLeft = '5px solid #10b981';
    top5Section.style.boxShadow = '0 4px 20px rgba(16, 185, 129, 0.12)';

    const top5Title = document.createElement('h5');
    top5Title.textContent = '🏆 Top 5 by Count of Dates';
    top5Title.style.marginBottom = '14px';
    top5Title.style.color = '#059669';
    top5Title.style.fontFamily = "'Montserrat', sans-serif";
    top5Title.style.fontWeight = '700';
    top5Title.style.fontSize = '1.1rem';
    top5Section.appendChild(top5Title);

    top5.forEach(item => {{
      const btn = document.createElement('button');
      btn.textContent = item.name + ' (' + item.count + ')';
      btn.style.margin = '6px 12px 6px 0';
      btn.style.padding = '10px 18px';
      btn.style.border = 'none';
      btn.style.borderRadius = '8px';
      btn.style.background = 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)';
      btn.style.color = 'white';
      btn.style.cursor = 'pointer';
      btn.style.fontSize = '0.88rem';
      btn.style.fontWeight = '600';
      btn.style.boxShadow = '0 4px 12px rgba(102, 126, 234, 0.3)';
      btn.style.transition = 'all 0.3s';
      btn.onmouseenter = function() {{ this.style.transform = 'translateY(-2px)'; this.style.boxShadow = '0 8px 25px rgba(102, 126, 234, 0.4)'; }};
      btn.onmouseleave = function() {{ this.style.transform = 'translateY(0)'; this.style.boxShadow = '0 4px 12px rgba(102, 126, 234, 0.3)'; }};
      btn.onclick = function() {{ showStockDetails(item.name, 'historyRank'); }};
      top5Section.appendChild(btn);
    }});

    infoPanel.appendChild(top5Section);
  }}

  if (top10.length === 0) {{
    const noData = document.createElement('p');
    noData.textContent = 'No rank improvements found for this date.';
    noData.style.color = 'var(--text-light)';
    infoPanel.appendChild(noData);
    return;
  }}

  const tableEl = document.createElement('table');
  tableEl.style.width = '100%';
  tableEl.style.borderCollapse = 'separate';
  tableEl.style.borderSpacing = '0';
  tableEl.style.fontSize = '0.92rem';
  tableEl.style.borderRadius = '12px';
  tableEl.style.overflow = 'hidden';
  tableEl.style.boxShadow = '0 4px 20px rgba(0,0,0,0.06)';

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  ['Rank', 'Stock', 'Also In Top 10 Dates', 'Count of Dates', 'Show Details', 'Previous Rank', 'Current Rank', 'Improvement'].forEach(text => {{
    const th = document.createElement('th');
    th.textContent = text;
    th.style.padding = '14px 16px';
    th.style.background = 'linear-gradient(135deg, #1e293b 0%, #334155 100%)';
    th.style.color = 'white';
    th.style.textAlign = 'left';
    th.style.fontFamily = "'Montserrat', sans-serif";
    th.style.fontWeight = '700';
    th.style.fontSize = '0.82rem';
    th.style.textTransform = 'uppercase';
    th.style.letterSpacing = '0.6px';
    th.style.borderBottom = '2px solid rgba(255,255,255,0.1)';
    headerRow.appendChild(th);
  }});
  thead.appendChild(headerRow);
  tableEl.appendChild(thead);

  const tbody = document.createElement('tbody');
  top10.forEach((stock, idx) => {{
    const tr = document.createElement('tr');
    tr.style.transition = 'all 0.2s';
    if (idx % 2 === 1) tr.style.background = '#f8fafc';

    const prevDates = (historyRankData.dailyTop10Stocks[stock.name] || [])
      .filter(d => d !== selectedDate)
      .filter(d => {{
        if (startDate && d < startDate) return false;
        if (endDate && d > endDate) return false;
        return true;
      }})
      .sort((a, b) => new Date(b) - new Date(a));

    const tdRank = document.createElement('td');
    tdRank.textContent = idx + 1;
    tdRank.style.padding = '14px 16px';
    tdRank.style.fontWeight = '700';
    tdRank.style.color = 'var(--primary)';
    tdRank.style.textAlign = 'center';

    const tdName = document.createElement('td');
    tdName.textContent = stock.name;
    tdName.style.padding = '14px 16px';
    tdName.style.fontWeight = '600';

    const tdDates = document.createElement('td');
    tdDates.textContent = prevDates.length > 0 ? prevDates.join(', ') : '-';
    tdDates.style.padding = '14px 16px';
    tdDates.style.color = '#e65100';
    tdDates.style.fontWeight = '500';
    tdDates.style.fontSize = '0.85rem';

    const tdCount = document.createElement('td');
    tdCount.textContent = prevDates.length;
    tdCount.style.padding = '14px 16px';
    tdCount.style.fontWeight = '700';
    tdCount.style.textAlign = 'center';
    tdCount.style.color = 'var(--primary)';

    const tdAction = document.createElement('td');
    tdAction.style.padding = '14px 16px';
    const showBtn = document.createElement('button');
    showBtn.textContent = 'Show Details';
    showBtn.style.padding = '8px 16px';
    showBtn.style.border = 'none';
    showBtn.style.borderRadius = '8px';
    showBtn.style.background = 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)';
    showBtn.style.color = 'white';
    showBtn.style.cursor = 'pointer';
    showBtn.style.fontSize = '0.82rem';
    showBtn.style.fontWeight = '600';
    showBtn.style.boxShadow = '0 4px 12px rgba(102, 126, 234, 0.3)';
    showBtn.style.transition = 'all 0.3s';
    showBtn.onmouseenter = function() {{ this.style.transform = 'translateY(-1px)'; this.style.boxShadow = '0 6px 20px rgba(102, 126, 234, 0.4)'; }};
    showBtn.onmouseleave = function() {{ this.style.transform = 'translateY(0)'; this.style.boxShadow = '0 4px 12px rgba(102, 126, 234, 0.3)'; }};
    showBtn.onclick = function() {{ showStockDetails(stock.name, 'historyRank'); }};
    tdAction.appendChild(showBtn);

    const tdPrev = document.createElement('td');
    tdPrev.textContent = stock.prevRank / 10;
    tdPrev.style.padding = '14px 16px';
    tdPrev.style.textAlign = 'center';

    const tdCurr = document.createElement('td');
    tdCurr.textContent = stock.currRank / 10;
    tdCurr.style.padding = '14px 16px';
    tdCurr.style.color = 'var(--success)';
    tdCurr.style.fontWeight = '700';
    tdCurr.style.textAlign = 'center';

    const tdImprove = document.createElement('td');
    tdImprove.textContent = '+' + (stock.improvement / 10);
    tdImprove.style.padding = '14px 16px';
    tdImprove.style.color = 'var(--success)';
    tdImprove.style.fontWeight = '700';
    tdImprove.style.textAlign = 'center';

    tr.appendChild(tdRank);
    tr.appendChild(tdName);
    tr.appendChild(tdDates);
    tr.appendChild(tdCount);
    tr.appendChild(tdAction);
    tr.appendChild(tdPrev);
    tr.appendChild(tdCurr);
    tr.appendChild(tdImprove);
    tbody.appendChild(tr);
  }});
  tableEl.appendChild(tbody);
  infoPanel.appendChild(tableEl);
}}

let performanceChart;
let topPerformersChart;

function showDeliveryChart() {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const rows = Array.from(table.querySelectorAll('tbody tr'));

  // Find the latest delivery column
  let deliveryColIndex = -1;
  for (let i = headers.length - 1; i >= 1; i--) {{
    const headerText = headers[i].textContent.trim().toLowerCase();
    if (headerText.includes('delivery')) {{
      deliveryColIndex = i;
      break;
    }}
  }}

  if (deliveryColIndex === -1) {{
    alert('No delivery columns found!');
    return;
  }}

  // Collect stock data with delivery values
  const stockData = [];
  rows.forEach(row => {{
    const stockName = row.cells[0].textContent.trim();
    const deliveryText = row.cells[deliveryColIndex]?.textContent.trim().replace(/,/g, '').replace('%', '') || '';
    const delivery = parseFloat(deliveryText);
    if (!isNaN(delivery) && delivery > 0) {{
      stockData.push({{ name: stockName, delivery: delivery }});
    }}
  }});

  // Sort by delivery descending and take top 10
  stockData.sort((a, b) => b.delivery - a.delivery);
  const top10 = stockData.slice(0, 10);

  // Prepare data for chart
  const labels = top10.map(item => item.name);
  const data = top10.map(item => item.delivery);

  // Create pie chart
  if (performanceChart) performanceChart.destroy();
  const ctx = document.getElementById('performanceChart').getContext('2d');

  performanceChart = new Chart(ctx, {{
    type: 'pie',
    data: {{
      labels: labels,
      datasets: [{{
        data: data,
        backgroundColor: [
          '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF',
          '#FF9F40', '#FF6384', '#C9CBCF', '#4BC0C0', '#FF6384'
        ],
        borderWidth: 1
      }}]
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{
          position: 'right'
        }},
        title: {{
          display: true,
          text: 'Top 10 Stocks by Delivery Percentage'
        }}
      }}
    }}
  }});
}}

function showMarketCapTrend() {{
  const table = document.getElementById('marketTable');
  const headers = Array.from(table.querySelectorAll('th'));
  const rows = Array.from(table.querySelectorAll('tbody tr'));

  // Find all market cap columns
  const marketCapColumns = [];
  headers.forEach((header, index) => {{
    if (index > 0 && header.textContent.trim().toLowerCase().includes('market')) {{
      marketCapColumns.push(index);
    }}
  }});

  if (marketCapColumns.length < 2) {{
    alert('Need at least 2 market cap columns for trend!');
    return;
  }}

  // Get dates from headers
  const dates = [];
  marketCapColumns.forEach(colIndex => {{
    const headerText = headers[colIndex].textContent.trim();
    const dateMatch = headerText.match(/^\\d{{4}}-\\d{{2}}-\\d{{2}}/);
    if (dateMatch) {{
      dates.push(dateMatch[0]);
    }}
  }});

  // Get top stock by latest market cap
  const latestMarketCapIndex = marketCapColumns[marketCapColumns.length - 1];
  const stockData = [];
  rows.forEach(row => {{
    const stockName = row.cells[0].textContent.trim();
    const marketCapText = row.cells[latestMarketCapIndex]?.textContent.trim().replace(/,/g, '') || '';
    const marketCap = parseFloat(marketCapText);
    if (!isNaN(marketCap)) {{
      stockData.push({{ name: stockName, marketCap: marketCap }});
    }}
  }});

  stockData.sort((a, b) => b.marketCap - a.marketCap);
  const topStock = stockData[0];

  // Prepare data for TradingView-style chart
  const data = marketCapColumns.map((colIndex, i) => {{
    const cell = rows.find(row => row.cells[0].textContent.trim() === topStock.name)?.cells[colIndex];
    const value = cell ? parseFloat(cell.textContent.trim().replace(/,/g, '')) : null;
    return {{
      time: dates[i],
      value: value
    }};
  }}).filter(item => item.value !== null);

  // Clear previous chart
  const chartContainer = document.getElementById('performanceChart');
  chartContainer.innerHTML = '';

  // Create TradingView-style chart
  const chart = LightweightCharts.createChart(chartContainer, {{
    width: chartContainer.clientWidth,
    height: 400,
    layout: {{
      backgroundColor: '#ffffff',
      textColor: '#333',
    }},
    grid: {{
      vertLines: {{
        color: '#e1ecf2',
      }},
      horzLines: {{
        color: '#e1ecf2',
      }},
    }},
    crosshair: {{
      mode: LightweightCharts.CrosshairMode.Normal,
    }},
    rightPriceScale: {{
      borderColor: '#cccccc',
    }},
    timeScale: {{
      borderColor: '#cccccc',
      timeVisible: true,
      secondsVisible: false,
    }},
  }});

  const areaSeries = chart.addAreaSeries({{
    topColor: 'rgba(33, 150, 243, 0.56)',
    bottomColor: 'rgba(33, 150, 243, 0.04)',
    lineColor: 'rgba(33, 150, 243, 1)',
    lineWidth: 2,
  }});

  areaSeries.setData(data);

  // Add title
  const titleElement = document.createElement('div');
  titleElement.innerHTML = `<h4 style="text-align: center; margin-bottom: 10px;">${{topStock.name}} Market Cap Trend</h4>`;
  chartContainer.insertBefore(titleElement, chartContainer.firstChild);

  // Resize chart on window resize
  window.addEventListener('resize', () => {{
    chart.applyOptions({{ width: chartContainer.clientWidth }});
  }});
}}

document.addEventListener('DOMContentLoaded', function() {{
  applyColorFormatting();
  loadStockList();
  updateChart('', '');
  updateDashboardStats();
  populateStockSuggestions();
  setDatePickerRanges();
  switchHistoryRankTab('single');
  switchTopPerformersTab('regular');

  document.getElementById('searchInput').addEventListener('keypress', function(e) {{
    if (e.key === 'Enter') searchStock();
  }});

  // Create tooltip element
  if (!document.getElementById('rank-tooltip')) {{
    const tooltip = document.createElement('div');
    tooltip.id = 'rank-tooltip';
    tooltip.className = 'rank-tooltip';
    document.body.appendChild(tooltip);
  }}
}});
</script>

</body>
</html>
"1"
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("Enhanced HTML file 'index.html' created with modern design!")
