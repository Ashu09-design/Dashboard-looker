# Digital Enablement Team - Daily Standup Dashboard (Dashboard Looker)

A lightweight, highly responsive, and data-driven dashboard application designed for the Digital Enablement Team. This tool visualizes team standups, bandwidth, leave statuses, and project blockers by pulling data directly from a centralized Excel file, Google Sheets, or SharePoint Excel link.

---

## Features

### **Data Integration & Live Updates**
- **Flexible Data Sources:** By default, it reads from a local Excel file (`_Bandwidth Tracker.xlsx`).
- **Cloud Connect:** Users can dynamically connect a **Google Sheet** (via a published export link) or a **SharePoint/OneDrive Excel** link directly from the UI using the "Data Source" modal.
- **Auto-Refresh:** The dashboard automatically polls the data source every 5 minutes to ensure the team is always looking at the latest standup data.

### **Python-Powered Excel Parsing (New)**
- **Smart Header Detection:** Scans up to 30 rows deep to find the real header row, handling title rows, blank rows, and merged cells that break traditional parsers.
- **Fuzzy Column Matching:** Uses string similarity matching with a configurable alias dictionary (`python/column_aliases.json`) to handle column name variations (e.g., "PM Name" vs "Project Manager" vs "PM").
- **Full .xlsb Support:** Native binary Excel (.xlsb) parsing via pandas/pyxlsb for the Leave Tracker.
- **Validation Diagnostics:** Every parse run outputs a diagnostic report showing which headers were found, which are missing, and where the header row was detected. Available via `/api/exec/data-health`.
- **Automatic Fallback:** If Python is not available, the system falls back to the original Node.js parsers seamlessly.

### **Standup & Delivery Tracking**
- **Smart "Yesterday & Today" Panels:** Automatically extracts and categorizes tasks into "What was done yesterday?" and "What will be done today?".
- **Deliverables Parsing:** A smart filter omits non-deliverable operational tasks (like daily syncs, calls, or meetings) from the main standup panels, focusing only on real work items.
- **Task Merging:** Groups multiple tasks for the same person on the same project into a cleanly comma-separated single-line summary for readability.

### **Bandwidth & Team Availability**
- **Visual Bandwidth Overview:** Displays clear percentage-based visual bars representing the total capacity of each team member.
- **Leave Status & Availability:** Automatically categorized tags indicate if a member is "Available", "Partial", on a "Half Day", or "On Leave" (Full Day).
- **Team Detail Table:** Sortable and searchable table outlining team roles, leave status, and exact bandwidth.

### **Executive Dashboard (Phase 2)**
- **Manager-only access** via JWT authentication.
- **6 independent data sources:** FTR Tracker, Team Details, SOW & PO, Governance, Leave Tracker, KPI.
- **Interactive charts** (Chart.js): FTE trends, KPI scorecards, SOW status.
- **Clickable KPI cards** with expandable detail panels.
- **Project health matrix** computed across all data sources.
- **File upload & cloud connect** for each data source independently.

### **Filtering & Search**
- **Global Search:** Find any team member, project name, or specific task instantly using the search bar.
- **Project Selection:** Filter the entire dashboard by selecting or deselecting specific ongoing projects from the right-hand sidebar.
- **Date Range Picker:** Drill down data to specific dates or sprints using the calendar filters.

### **Presentation Mode**
- Enables an immersive fullscreen, clean UI ideal for sharing on a TV monitor or during a Zoom/Teams screen-share session.

---

## Architecture & Stack

```
server.js (Express)
├── middleware/auth.js              # JWT authentication
├── services/execDataService.js     # Data caching, Python bridge, Node fallback
├── python/
│   ├── normalizer.py               # Smart Excel parser (header detection, fuzzy matching)
│   ├── column_aliases.json         # Column name alias mappings
│   └── requirements.txt            # Python dependencies
├── parsers/                        # Original Node.js parsers (used as fallback)
│   ├── teamParser.js
│   ├── ftrParser.js
│   ├── kpiParser.js
│   ├── leaveParser.js
│   ├── sowParser.js
│   └── governanceParser.js
├── data/                           # Python normalizer output (JSON, gitignored)
└── public/
    ├── index.html + app.js         # Phase 1 (Standup Dashboard)
    ├── executive.html + executive.js # Phase 2 (Executive Dashboard)
    └── login.html                  # Authentication
```

### **Data Flow (Executive Dashboard)**
```
Excel File (.xlsx/.xlsb) or Cloud Source
        │
        ▼
Python Normalizer (normalizer.py)
  - Smart header row detection (scans up to 30 rows)
  - Fuzzy column name matching via aliases
  - Data type inference (dates, numbers, currencies)
  - Full .xlsb binary format support
  - Outputs clean JSON + diagnostics
        │
        ▼ (falls back to Node.js parsers if Python unavailable)
        │
execDataService.js (in-memory cache)
        │
        ▼
REST API → Frontend (Chart.js, Vanilla JS)
```

### **Backend (`server.js`)**
- Node.js & Express server.
- **Phase 1 Endpoints:**
  - `/api/bandwidth`: Maps and cleans bandwidth, task, and leave status data.
  - `/api/qbr`: Extracts MBR/QBR dates and blockers.
  - `/api/dropdown`: Exposes lookup data (lookup tables for Managers and Project descriptions).
  - `/api/connect-source` & `/api/source-status`: Manages live cloud data source connections.
- **Phase 2 Endpoints (JWT-protected):**
  - `/api/exec/summary`: KPI summary (team size, projects, SOW, metrics).
  - `/api/exec/team-capacity`: Team details and leave impact.
  - `/api/exec/kpi-scorecards`: KPI metrics with target vs actual.
  - `/api/exec/sow-financial`: SOW & PO status and financial data.
  - `/api/exec/governance-risks`: Risks, audits, FTE trends, QBR schedule.
  - `/api/exec/leave-impact`: Leave tracker by person/month.
  - `/api/exec/ftr-metrics`: FTR QA metrics and pass rates.
  - `/api/exec/project-health`: Cross-file project health status.
  - `/api/exec/data-health`: Parsing diagnostics and validation report.
  - `/api/exec/upload-sources`: Upload Excel files.
  - `/api/exec/refresh`: Manual data refresh.

### **Frontend (`public/`)**
- Vanilla JavaScript (`app.js`), HTML5 (`index.html`), Vanilla CSS (`styles.css`).
- Clean, dependency-free frontend to ensure maximum performance and minimal footprint.
- Uses modern CSS Flexbox/Grid and variables for theming.

---

## Setup & Installation

### 1. Prerequisites
- **Node.js** (v14 or higher)
- **Python 3.8+** (recommended for improved Excel parsing)

### 2. Install Dependencies

**Node.js dependencies:**
```bash
npm install
```

**Python dependencies (recommended):**
```bash
pip install -r python/requirements.txt
```

This installs `openpyxl`, `pandas`, and `pyxlsb` for robust Excel parsing. If Python is not installed, the app will automatically fall back to the original Node.js parsers.

### 3. Data File Placement
Ensure you have the master bandwidth tracker Excel file located at the required path. By default, the server expects it at:
`../pov/_Bandwidth Tracker.xlsx`
*(Note: You can override this using the `EXCEL_PATH` environment variable).*

### 4. Environment Variables
Create a `.env` file in the project root:
```env
PORT=3000
EXCEL_PATH=../pov/_Bandwidth Tracker.xlsx
GOOGLE_SHEET_URL=                         # Optional: auto-connect Google Sheet
JWT_SECRET=your-secret-key
MANAGER_USER=admin
MANAGER_PASS_HASH=                        # bcrypt hash of manager password

# Executive data source paths (Phase 2)
EXEC_FTR_PATH=/path/to/ftr.xlsx
EXEC_TEAM_PATH=/path/to/team.xlsx
EXEC_SOW_PATH=/path/to/sow.xlsx
EXEC_GOVERNANCE_PATH=/path/to/governance.xlsx
EXEC_LEAVE_PATH=/path/to/leave.xlsb
EXEC_KPI_PATH=/path/to/kpi.xlsx
```

### 5. Start the Application
```bash
npm start
```

The server will initialize on port `3000`.
**Open your browser to: `http://localhost:3000`**

---

## Python Normalizer Details

The Python normalizer (`python/normalizer.py`) replaces the original Node.js Excel parsers with a more robust solution:

### Why Python?
| Problem | Node.js (xlsx library) | Python (openpyxl + pandas) |
|---------|----------------------|---------------------------|
| Header not in row 0-5 | Fails silently | Scans up to 30 rows deep |
| Column name variations | Exact match only | Fuzzy matching + alias dictionary |
| .xlsb binary format | Limited support | Native via pyxlsb |
| Empty/merged cells | Returns `__EMPTY` keys | Handles gracefully |
| No error feedback | Returns null | Full diagnostic report |

### Column Aliases
Edit `python/column_aliases.json` to add new column name variations. For example, if someone renames "Project Manager" to "PM Lead":

```json
{
  "sow": {
    "pm": ["PM Name", "PM", "Project Manager", "Manager", "Lead", "PM Lead"]
  }
}
```

### CLI Usage
Run the normalizer directly for testing/debugging:
```bash
# Parse a single source
python python/normalizer.py --source team --input path/to/team.xlsx --output data/team.json

# Parse all sources at once
python python/normalizer.py --source all --config sources.json --outdir data/
```

### Data Health Endpoint
After login, call `GET /api/exec/data-health` to see parsing diagnostics:
```json
{
  "team": {
    "name": "Team Details",
    "success": true,
    "counts": { "activeMembers": 42 },
    "diagnostics": {
      "sheets_found": ["Team details 2026", "Exit resources", "..."],
      "sections": {
        "current_team": { "header_row": 0, "columns_found": ["name", "role", "..."], "columns_missing": [] }
      }
    }
  }
}
```

---

## Mock Data Generator

If your Excel file is empty and you want to visualize the dashboard immediately, use the included Python script to fill it with reliable test data.

1. Ensure `EXCEL_PATH` inside `fill_recent_data.py` points to your active Excel file.
2. Run the script:
   ```bash
   python fill_recent_data.py
   ```
3. The script will automatically read the `Drop Down` lookup sheet, grab team members and projects, and append about 20 random tasks (both deliverables and ops) with realistic time tracking to the currently active dates (Yesterday and Today).
4. Refresh the web dashboard (or wait for the auto-refresh) to see the new data populated!
