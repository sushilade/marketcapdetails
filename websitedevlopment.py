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
      --primary: #2c3e50;
      --secondary: #34495e;
      --accent: #3498db;
      --success: #2ecc71;
      --danger: #e74c3c;
      --warning: #f39c12;
      --light: #ecf0f1;
      --dark: #2c3e50;
      --text: #2c3e50;
      --text-light: #7f8c8d;
      --bg-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      --card-shadow: 0 10px 30px rgba(0,0,0,0.1);
      --hover-shadow: 0 15px 35px rgba(0,0,0,0.15);
    }}
    
    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}
    
    body {{
      font-family: 'Roboto', sans-serif;
      background: #f5f7fa;
      color: var(--text);
      line-height: 1.6;
      margin: 0;
      padding: 0;
    }}
    
    .container {{
      max-width: 1400px;
      margin: 0 auto;
      padding: 20px;
    }}
    
    header {{
      background: var(--bg-gradient);
      color: white;
      padding: 25px 20px;
      border-radius: 10px;
      margin-bottom: 25px;
      box-shadow: var(--card-shadow);
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
      gap: 15px;
      margin-bottom: 25px;
    }}
    
    .stat-card {{
      background: white;
      border-radius: 10px;
      padding: 20px;
      flex: 1;
      min-width: 200px;
      box-shadow: var(--card-shadow);
      transition: transform 0.3s, box-shadow 0.3s;
    }}
    
    .stat-card:hover {{
      transform: translateY(-5px);
      box-shadow: var(--hover-shadow);
    }}
    
    .stat-value {{
      font-size: 2rem;
      font-weight: 700;
      color: var(--accent);
      margin: 10px 0;
    }}
    
    .stat-label {{
      color: var(--text-light);
      font-size: 0.9rem;
    }}
    
    /* Tabs */
    .tab-container {{
      background: white;
      border-radius: 10px;
      overflow: hidden;
      box-shadow: var(--card-shadow);
      margin-bottom: 30px;
    }}
    
    .tab {{
      display: flex;
      background: var(--primary);
      overflow-x: auto;
    }}
    
    .tab button {{
      background: transparent;
      color: rgba(255, 255, 255, 0.8);
      border: none;
      outline: none;
      cursor: pointer;
      padding: 15px 25px;
      font-family: 'Montserrat', sans-serif;
      font-weight: 600;
      font-size: 16px;
      transition: all 0.3s;
      white-space: nowrap;
    }}
    
    .tab button:hover {{
      background: rgba(255, 255, 255, 0.1);
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
      background: var(--accent);
    }}
    
    .tabcontent {{
      display: none;
      padding: 25px;
    }}
    
    #dataView {{
      display: block;
    }}
    
    .controls {{
      margin-bottom: 25px;
      padding: 20px;
      background: var(--light);
      border-radius: 10px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }}
    
    .controls h3 {{
      margin-bottom: 15px;
      color: var(--secondary);
    }}
    
    .filter-buttons {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 15px;
    }}

    .search-container {{
      display: flex;
      flex-wrap: wrap;
      gap: 15px;
      align-items: center;
      margin-top: 15px;
      padding: 15px;
      background: rgba(255, 255, 255, 0.8);
      border-radius: 10px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }}

    .search-container label {{
      font-weight: 600;
      color: var(--primary);
      margin-right: 5px;
    }}

    .search-container input, .search-container select {{
      padding: 10px 12px;
      border: 1px solid #ddd;
      border-radius: 6px;
      font-size: 14px;
      transition: border-color 0.3s;
    }}

    .search-container input:focus, .search-container select:focus {{
      border-color: var(--accent);
      outline: none;
      box-shadow: 0 0 0 2px rgba(52, 152, 219, 0.2);
    }}

    .search-container button {{
      padding: 10px 16px;
      background: var(--accent);
      color: white;
      border: none;
      border-radius: 6px;
      cursor: pointer;
      transition: all 0.3s;
      font-weight: 500;
      display: flex;
      align-items: center;
      gap: 5px;
    }}

    .search-container button:hover {{
      background: #2980b9;
      transform: translateY(-1px);
      box-shadow: 0 4px 12px rgba(52, 152, 219, 0.3);
    }}

    .search-container button:active {{
      transform: translateY(0);
    }}
    
    button {{
      padding: 10px 18px;
      font-size: 14px;
      font-weight: 500;
      background: white;
      color: var(--primary);
      border: 1px solid #ddd;
      border-radius: 6px;
      cursor: pointer;
      transition: all 0.3s;
      display: flex;
      align-items: center;
      gap: 8px;
    }}
    
    button:hover {{
      background: var(--accent);
      color: white;
      border-color: var(--accent);
      transform: translateY(-2px);
      box-shadow: 0 5px 15px rgba(52, 152, 219, 0.3);
    }}
    
    button.active {{
      background: var(--success);
      color: white;
      border-color: var(--success);
    }}
    
    .download-btn {{
      background: var(--success);
      color: white;
      border: none;
      padding: 12px 20px;
      margin-bottom: 25px;
    }}
    
    .download-btn:hover {{
      background: #27ae60;
      box-shadow: 0 5px 15px rgba(46, 204, 113, 0.4);
    }}
    
    /* Table Styling */
    .table-container {{
      overflow-x: auto;
      border-radius: 10px;
      box-shadow: var(--card-shadow);
      margin-top: 20px;
    }}
    
    table {{
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      font-size: 0.9rem;
    }}
    
    th {{
      background: var(--primary);
      color: white;
      padding: 15px 12px;
      text-align: left;
      font-weight: 600;
      position: sticky;
      top: 0;
      font-family: 'Montserrat', sans-serif;
    }}
    
    th:first-child {{
      border-top-left-radius: 10px;
    }}
    
    th:last-child {{
      border-top-right-radius: 10px;
    }}
    
    td {{
      padding: 12px;
      border-bottom: 1px solid #eee;
    }}
    
    tr {{
      transition: background 0.2s;
    }}
    
    tr:nth-child(even) {{
      background: #f9fafb;
    }}
    
    tr:hover {{
      background: #f0f7ff;
    }}
    
    .increase-light {{ background-color: rgba(46, 204, 113, 0.15) !important; }}
    .increase-dark {{ background-color: rgba(46, 204, 113, 0.3) !important; }}
    .decrease-light {{ background-color: rgba(231, 76, 60, 0.15) !important; }}
    .decrease-dark {{ background-color: rgba(231, 76, 60, 0.3) !important; }}
    
    /* Form Elements */
    select, input {{
      padding: 10px 15px;
      border: 1px solid #ddd;
      border-radius: 6px;
      font-family: 'Roboto', sans-serif;
      font-size: 14px;
      margin: 10px 0;
      width: 100%;
      max-width: 300px;
    }}
    
    select:focus, input:focus {{
      outline: none;
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.2);
    }}
    
    /* Chart Containers */
    .chart-container {{
      background: white;
      padding: 20px;
      border-radius: 10px;
      box-shadow: var(--card-shadow);
      margin: 20px 0;
    }}
    
    .info-panel {{
      margin-top: 20px;
      padding: 20px;
      background: white;
      border-radius: 10px;
      box-shadow: var(--card-shadow);
    }}
    
    .info-item {{
      margin-bottom: 15px;
      padding: 15px;
      background: var(--light);
      border-radius: 8px;
      border-left: 4px solid var(--accent);
    }}
    
    .rank-history {{
      margin-top: 10px;
      font-size: 0.9em;
      color: var(--text-light);
    }}
    
    .rank-entry {{
      display: inline-block;
      margin-right: 15px;
      margin-bottom: 8px;
      padding: 5px 10px;
      background: white;
      border-radius: 4px;
      box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }}
    
    .rank-label {{
      font-weight: 600;
      color: var(--primary);
    }}
    
    .rank-tooltip {{
      position: absolute;
      background: rgba(44, 62, 80, 0.95);
      color: white;
      padding: 10px 15px;
      border-radius: 6px;
      font-size: 13px;
      pointer-events: none;
      z-index: 1000;
      max-width: 300px;
      box-shadow: 0 5px 15px rgba(0,0,0,0.2);
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
        padding: 12px 18px;
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
      }}
      
      th, td {{
        padding: 8px 10px;
        font-size: 0.85rem;
      }}
    }}
    
    /* Animation */
    @keyframes fadeIn {{
      from {{ opacity: 0; transform: translateY(10px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}
    
    .tabcontent {{
      animation: fadeIn 0.3s ease;
    }}
    
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
  <header>
    <h1>Market Data Analysis Dashboard</h1>
    <p class="subtitle"></p>
  </header>
  
  <div class="dashboard-stats">
    <div class="stat-card">
      <div class="stat-label">TOTAL STOCKS</div>
      <div class="stat-value" id="totalStocks">0</div>
      <div class="stat-desc">Across all exchanges</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">AVG MARKET CAP</div>
      <div class="stat-value" id="avgMarketCap">0</div>
      <div class="stat-desc">Latest trading session</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">TOP PERFORMER</div>
      <div class="stat-value" id="topPerformer">-</div>
      <div class="stat-desc">Highest delivery %</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">LAST UPDATED</div>
      <div class="stat-value" id="lastUpdated">{datetime.now().strftime('%Y-%m-%d')}</div>
      <div class="stat-desc">Data freshness</div>
    </div>
  </div>

  <div class="tab-container">
    <div class="tab">
      <button class="tablinks active" onclick="openTab(event, 'dataView')">
        <i class="fas fa-table"></i> Data View
      </button>
      <button class="tablinks" onclick="openTab(event, 'stockAnalysis')">
        <i class="fas fa-chart-line"></i> Stock Analysis
      </button>
      <button class="tablinks" onclick="openTab(event, 'marketOverview')">
        <i class="fas fa-globe"></i> Market Overview
      </button>
      <button class="tablinks" onclick="openTab(event, 'topPerformers')">
        <i class="fas fa-trophy"></i> Top Performers
      </button>
      <button class="tablinks" onclick="openTab(event, 'nonPerformers')">
        <i class="fas fa-thumbs-down"></i> Non Performers
      </button>
      <button class="tablinks" onclick="openTab(event, 'rankChanges')">
        <i class="fas fa-arrow-up"></i> Rank Up
      </button>
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
        {html_table}
      </div>
    </div>

    <!-- Stock Analysis Tab -->
    <div id="stockAnalysis" class="tabcontent">
      <h3><i class="fas fa-chart-line"></i> Individual Stock Analysis</h3>
      
      <div class="chart-container">
        <label for="stockSelect">Select Stock:</label>
        <select id="stockSelect" onchange="updateChart()"></select>
        
        <div style="position: relative; height:400px; margin-top: 20px;">
          <canvas id="stockChart"></canvas>
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

      <div class="info-panel" id="topPerformersList">
        <p>Click "Show Top Performers" to display the list of top 10 stocks by delivery percentage.</p>
      </div>

      <button onclick="showTopPerformers()" id="showTopPerformersBtn">
        <i class="fas fa-list"></i> Show Top Performers
      </button>
    </div>

    <!-- Non Performers Tab -->
    <div id="nonPerformers" class="tabcontent">
      <h3><i class="fas fa-thumbs-down"></i> Top 10 Non-Performing Stocks</h3>

      <div class="info-panel" id="nonPerformersList">
        <p>Click "Show Non Performers" to display the list of top 10 stocks with lowest delivery percentage.</p>
      </div>

      <button onclick="showNonPerformers()" id="showNonPerformersBtn">
        <i class="fas fa-list"></i> Show Non Performers
      </button>
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
  evt.currentTarget.classList.add("active");
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
    const aBlank = (aCell === '' || aCell === '-');
    const bBlank = (bCell === '' || bCell === '-');
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

function updateChart() {{
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
  const mapping = {{}}; // mapping[date] = {{ marketIdx: idx or null, deliveryIdx: idx or null }}

  for (let i = 1; i < headers.length; i++) {{
    const h = headers[i];
    const m = h.match(datePattern);
    if (!m) continue;
    const d = m[0];
    if (!dates.includes(d)) dates.push(d);
    if (!mapping[d]) mapping[d] = {{ marketIdx: null, deliveryIdx: null }};
    const lower = h.toLowerCase();
    if (lower.includes('market')) {{
      mapping[d].marketIdx = i; // column index in table (th/td)
    }} else if (lower.includes('delivery')) {{
      mapping[d].deliveryIdx = i;
    }} else if (lower.includes('rank')) {{
      // intentionally ignore rank for plotting
    }} else {{
      // if header does not explicitly include 'market' or 'delivery', try to guess by keyword
      // default: do nothing
    }}
  }}

  // Build arrays of data in the same order as dates
  const marketCapData = [];
  const deliveryData = [];
  const labels = dates.slice(); // copy

  dates.forEach(d => {{
    const map = mapping[d] || {{marketIdx:null, deliveryIdx:null}};
    let marketVal = 0;
    let deliveryVal = 0;
    if (map.marketIdx !== null) {{
      const txt = (row.cells[map.marketIdx] && row.cells[map.marketIdx].textContent) ? row.cells[map.marketIdx].textContent.trim().replace(/,/g,'').replace('%','') : '';
      marketVal = parseFloat(txt) || 0;
    }}
    if (map.deliveryIdx !== null) {{
      const txt = (row.cells[map.deliveryIdx] && row.cells[map.deliveryIdx].textContent) ? row.cells[map.deliveryIdx].textContent.trim().replace(/,/g,'').replace('%','') : '';
      deliveryVal = parseFloat(txt) || 0;
    }}
    marketCapData.push(marketVal);
    deliveryData.push(deliveryVal);
  }});

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
        legend: {{ position: 'top' }},
        title: {{
          display: true,
          text: 'Market Cap (Bar) vs Delivery Qty (Line) for ' + stock
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
        }}
      }}
    }}
  }});
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
    item.innerHTML = `<strong>${{stock.name}}</strong> - ${{stock.delivery.toFixed(2)}}%`;
    list.appendChild(item);
  }});
  infoPanel.appendChild(list);
}}

function showNonPerformers() {{
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
    if (!isNaN(delivery)) {{
      stockData.push({{ name: stockName, delivery: delivery }});
    }}
  }});

  // Sort by delivery ascending (lowest first) and take top 10
  stockData.sort((a, b) => a.delivery - b.delivery);
  const bottom10 = stockData.slice(0, 10);

  // Display in info panel
  const infoPanel = document.getElementById('nonPerformersList');
  infoPanel.innerHTML = '<h4>Top 10 Non-Performing Stocks by Delivery Percentage</h4>';
  const list = document.createElement('ol');
  bottom10.forEach((stock, index) => {{
    const item = document.createElement('li');
    item.innerHTML = `<strong>${{stock.name}}</strong> - ${{stock.delivery.toFixed(2)}}%`;
    list.appendChild(item);
  }});
  infoPanel.appendChild(list);
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

let performanceChart;
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
  updateChart();
  updateDashboardStats();
  populateStockSuggestions();
  setDatePickerRanges();

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