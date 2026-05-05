# Digital Enablement Team — Dashboard Deployment Guide

## Project Structure

```
Dashboard Looker/
├── server.js              # Express backend — reads Excel, serves API
├── package.json
├── .env                   # (create this) EXCEL_PATH configuration
├── public/
│   ├── index.html         # Looker Studio layout
│   ├── styles.css         # Cyan/teal theme styling
│   └── app.js             # Client-side logic (filtering, search, sorting)
└── ../pov/
    └── _Bandwidth Tracker.xlsx  # Data source
```

## Run Locally

```bash
# 1. Install dependencies
npm install

# 2. Start the server
npm run dev         # or: node server.js

# 3. Open in Chrome
#    http://localhost:3000
```

## Configure Data Source

### Option A: Local Excel File
Create a `.env` file (or set environment variable):
```env
EXCEL_PATH=C:/Users/KIIT/Downloads/pov/_Bandwidth Tracker.xlsx
```
By default it looks for `../pov/_Bandwidth Tracker.xlsx` relative to the project folder.

### Option B: Google Sheets (future)
Use the "Data Source" modal in the UI to paste a published Google Sheets CSV export link:
```
https://docs.google.com/spreadsheets/d/SHEET_ID/export?format=csv
```

### Option C: Excel Online / SharePoint (future)
Paste a SharePoint URL to the Excel file in the data source modal.

## Deploy to Azure Web App

### Prerequisites
- Azure CLI installed (`az login`)
- Node.js 18+

### Steps

```bash
# 1. Create Azure resource group
az group create --name rg-dashboard --location centralindia

# 2. Create App Service plan
az appservice plan create --name plan-dashboard --resource-group rg-dashboard --sku B1 --is-linux

# 3. Create Web App
az webapp create --name det-dashboard --resource-group rg-dashboard --plan plan-dashboard --runtime "NODE|18-lts"

# 4. Set the Excel path as an App Setting
az webapp config appsettings set --name det-dashboard --resource-group rg-dashboard --settings EXCEL_PATH=/home/site/wwwroot/data/_Bandwidth_Tracker.xlsx

# 5. Deploy code (from project root)
az webapp up --name det-dashboard --resource-group rg-dashboard
```

### Upload Excel to Azure
Upload the Excel file to the app's file system:
```bash
az webapp deploy --name det-dashboard --resource-group rg-dashboard --src-path "./../pov/_Bandwidth Tracker.xlsx" --target-path "/home/site/wwwroot/data/_Bandwidth_Tracker.xlsx" --type static
```

## Sheet Requirements

The Excel file must have these sheets:

### Bandwidth Tracker
| Column | Description |
|---|---|
| Date | Date in DD/MM/YYYY or Excel serial format |
| Project | Project name |
| Project Details | (resolved via Drop Down if empty) |
| Project Manager | (resolved via Drop Down if empty) |
| Name | Team member name |
| Role | Job role |
| Work Item | Type of work |
| Description | Task description (used for Yesterday/Today) |
| Time | Time spent |
| Leave Status | "Full Day", "Half Day", or empty |
| Free Bandwidth | Availability text |

### QBR Date & Blockers
| Column | Description |
|---|---|
| Date | Date |
| Project | Project name |
| Project Details | (resolved via Drop Down if empty) |
| Project Manager | (resolved via Drop Down if empty) |
| MBR/QBR Date | QBR date |
| Blockers / Dependencies | Blocker text |

### Drop Down
| Column | Description |
|---|---|
| Member | Team member name |
| Project Name | Project name (VLOOKUP key) |
| (Col C) | Project details text |
| Work Item | Work item type |
| Resource Type | Role type |
| Project Manager | Manager name |

## Features

- ✅ Dynamic column detection by header name (no hardcoded indexes)
- ✅ Auto-refresh every 5 minutes
- ✅ Project filtering with ONLY button
- ✅ Date range filtering
- ✅ Search across all data
- ✅ Sort by Team Details column
- ✅ Loading spinner
- ✅ Error handling
- ✅ Responsive layout
- ✅ Optimized for 1000+ rows
