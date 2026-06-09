require('dotenv').config();
const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const os = require('os');
const bcrypt = require('bcryptjs');
const multer = require('multer');
const { verifyToken, signToken } = require('./middleware/auth');
const execData = require('./services/execDataService');

const app = express();
const PORT = process.env.PORT || 3000;

// ── Multer config for Excel uploads ───────────────────────────────
const uploadDir = path.join(os.tmpdir(), 'dashboard-uploads');
if (!fs.existsSync(uploadDir)) {
  try {
    fs.mkdirSync(uploadDir, { recursive: true });
  } catch (e) {
    console.warn('Could not create uploadDir:', e.message);
  }
}
const upload = multer({
  dest: uploadDir,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (['.xlsx', '.xlsb', '.xls'].includes(ext)) cb(null, true);
    else cb(new Error('Only Excel files (.xlsx, .xlsb, .xls) are allowed'));
  },
});

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Data source state ─────────────────────────────────────────────────
let connectedSource = { type: null, url: null };

// ── Path to the Excel file ────────────────────────────────────────────
const EXCEL_PATH =
  process.env.EXCEL_PATH ||
  path.join(__dirname, '..', 'pov', '_Bandwidth Tracker.xlsx');

const CACHED_SHEET_PATH = path.join(os.tmpdir(), '_cached_sheet.xlsx');

// ══════════════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════════════

/** Extract Google Sheets spreadsheet ID from various URL formats */
function extractSheetId(url) {
  const m = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

/** Download a Google Sheet as XLSX and save locally */
async function downloadGoogleSheet(url) {
  const sheetId = extractSheetId(url);
  if (!sheetId) throw new Error('Invalid Google Sheets URL — could not extract sheet ID');

  const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
  console.log(`Fetching Google Sheet: ${exportUrl}`);

  const resp = await fetch(exportUrl);
  if (!resp.ok) throw new Error(`Google Sheets download failed: ${resp.status} ${resp.statusText}`);

  const buffer = Buffer.from(await resp.arrayBuffer());
  fs.writeFileSync(CACHED_SHEET_PATH, buffer);
  console.log(`Google Sheet downloaded & cached (${(buffer.length / 1024).toFixed(1)} KB)`);
  return XLSX.read(buffer, { type: 'buffer' });
}

/** Load workbook — from Google Sheets if connected, else local file */
async function loadWorkbook() {
  // If a Google Sheet is connected, download fresh copy
  if (connectedSource.type === 'google' && connectedSource.url) {
    return await downloadGoogleSheet(connectedSource.url);
  }

  // Fall back to local Excel file
  if (!fs.existsSync(EXCEL_PATH)) {
    console.warn(`Excel file not found at: ${EXCEL_PATH} — returning null`);
    return null;
  }
  return XLSX.readFile(EXCEL_PATH);
}

/** Synchronous fallback for startup / simple cases */
function loadWorkbookSync() {
  if (fs.existsSync(CACHED_SHEET_PATH) && connectedSource.type === 'google') {
    return XLSX.readFile(CACHED_SHEET_PATH);
  }
  if (!fs.existsSync(EXCEL_PATH)) {
    throw new Error(`Excel file not found at: ${EXCEL_PATH}`);
  }
  return XLSX.readFile(EXCEL_PATH);
}

// Auto-connect Google Sheet from env var
if (process.env.GOOGLE_SHEET_URL) {
  connectedSource = { type: 'google', url: process.env.GOOGLE_SHEET_URL };
  console.log('Auto-connected Google Sheet from env:', process.env.GOOGLE_SHEET_URL);
}

function cleanString(val) {
  if (val == null) return '';
  return val.toString().replace(/\u200b/g, '').trim();
}

function formatDate(val) {
  if (!val) return '';
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  const str = val.toString().trim();
  // YYYY-MM-DD
  const m = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  // DD/MM/YYYY
  const m2 = str.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m2) return `${m2[3]}-${m2[2]}-${m2[1]}`;
  return str;
}

/** Build lookup maps from the Drop Down sheet */
function buildDropdownLookup(wb) {
  const ws = wb.Sheets['Drop Down'];
  if (!ws) return { detailsByProject: {}, managerByProject: {} };

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  // Detect columns by header name
  const hdr = rows[0] || [];
  const colMap = {};
  hdr.forEach((h, i) => { if (h) colMap[cleanString(h).toLowerCase()] = i; });

  const projCol = colMap['project name'] ?? 1;
  const detailIdx = 2; // Column C has project details text
  const pmCol = colMap['project manager'] ?? 5;

  const detailsByProject = {};
  const managerByProject = {};

  for (let i = 1; i < rows.length; i++) {
    const projectName = cleanString(rows[i][projCol]);
    if (!projectName) continue;
    const detail = cleanString(rows[i][detailIdx]);
    const manager = cleanString(rows[i][pmCol]);
    if (detail) detailsByProject[projectName] = detail;
    if (manager) managerByProject[projectName] = manager;
  }
  return { detailsByProject, managerByProject };
}

/** Detect column indexes by header names (no hardcoded indexes) */
function detectColumns(headerRow) {
  const map = {};
  (headerRow || []).forEach((h, i) => {
    const key = cleanString(h).toLowerCase();
    if (key) map[key] = i;
  });
  return map;
}

// ══════════════════════════════════════════════════════════════════════
//  API: Bandwidth Tracker
// ══════════════════════════════════════════════════════════════════════

app.get('/api/bandwidth', async (req, res) => {
  try {
    const wb = await loadWorkbook();
    if (!wb) return res.json([]);
    const { detailsByProject, managerByProject } = buildDropdownLookup(wb);
    const ws = wb.Sheets['Bandwidth Tracker'];
    if (!ws) return res.status(404).json({ error: 'Sheet "Bandwidth Tracker" not found' });

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const cols = detectColumns(rows[0]);

    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const date = formatDate(r[cols['date'] ?? 0]);
      const project = cleanString(r[cols['project'] ?? 1]);
      if (!date && !project) continue;

      // Resolve Project Details & PM via VLOOKUP if empty
      let projectDetails = cleanString(r[cols['project details'] ?? 2]);
      let projectManager = cleanString(r[cols['project manager'] ?? 3]);
      if (!projectDetails || projectDetails.startsWith('='))
        projectDetails = detailsByProject[project] || '';
      if (!projectManager || projectManager.startsWith('='))
        projectManager = managerByProject[project] || '';

      data.push({
        date,
        project,
        projectDetails,
        projectManager,
        name: cleanString(r[cols['name'] ?? 4]),
        role: cleanString(r[cols['role'] ?? 5]),
        workItem: cleanString(r[cols['work item'] ?? 6]),
        description: cleanString(r[cols['description'] ?? 7]),
        time: r[cols['time'] ?? 8] || '',
        leaveStatus: cleanString(r[cols['leave status'] ?? 9]),
        freeBandwidth: cleanString(r[cols['free bandwidth'] ?? 10]),
      });
    }
    res.json(data);
  } catch (err) {
    console.error('Bandwidth API error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ══════════════════════════════════════════════════════════════════════
//  API: QBR Date & Blockers
// ══════════════════════════════════════════════════════════════════════

app.get('/api/qbr', async (req, res) => {
  try {
    const wb = await loadWorkbook();
    if (!wb) return res.json([]);
    const { detailsByProject, managerByProject } = buildDropdownLookup(wb);
    const ws = wb.Sheets['QBR Date & Blockers'];
    if (!ws) return res.status(404).json({ error: 'Sheet "QBR Date & Blockers" not found' });

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const cols = detectColumns(rows[0]);

    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const date = formatDate(r[cols['date'] ?? 0]);
      const project = cleanString(r[cols['project'] ?? 1]);
      if (!date && !project) continue;

      let projectDetails = cleanString(r[cols['project details'] ?? 2]);
      let projectManager = cleanString(r[cols['project manager'] ?? 3]);
      if (!projectDetails || projectDetails.startsWith('='))
        projectDetails = detailsByProject[project] || '';
      if (!projectManager || projectManager.startsWith('='))
        projectManager = managerByProject[project] || '';

      data.push({
        date,
        project,
        projectDetails,
        projectManager,
        mbrQbrDate: formatDate(r[cols['mbr/qbr date'] ?? 4]),
        blockers: cleanString(r[cols['blockers / dependencies'] ?? 5]),
      });
    }
    res.json(data);
  } catch (err) {
    console.error('QBR API error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ══════════════════════════════════════════════════════════════════════
//  API: Drop Down (lookup data)
// ══════════════════════════════════════════════════════════════════════

app.get('/api/dropdown', async (req, res) => {
  try {
    const wb = await loadWorkbook();
    if (!wb) return res.json([]);
    const ws = wb.Sheets['Drop Down'];
    if (!ws) return res.status(404).json({ error: 'Sheet "Drop Down" not found' });

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const cols = detectColumns(rows[0]);

    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      data.push({
        member: cleanString(r[cols['member'] ?? 0]),
        projectName: cleanString(r[cols['project name'] ?? 1]),
        projectDetails: cleanString(r[2]), // col C
        workItem: cleanString(r[cols['work item'] ?? 3]),
        resourceType: cleanString(r[cols['resource type'] ?? 4]),
        projectManager: cleanString(r[cols['project manager'] ?? 5]),
      });
    }
    res.json(data);
  } catch (err) {
    console.error('Dropdown API error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ══════════════════════════════════════════════════════════════════════
//  API: Connect external data source
// ══════════════════════════════════════════════════════════════════════

app.post('/api/connect-source', (req, res) => {
  try {
    const { type, url } = req.body;
    if (!type || !url) {
      return res.status(400).json({ success: false, error: 'Type and URL are required.' });
    }
    connectedSource = { type, url };
    console.log(`Data source connected: [${type}] ${url}`);
    res.json({ success: true, message: `${type} source connected successfully.` });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/source-status', (req, res) => {
  res.json({
    connected: connectedSource.type !== null,
    type: connectedSource.type,
    url: connectedSource.url,
  });
});

// ══════════════════════════════════════════════════════════════════
//  PHASE 2: AUTHENTICATION
// ══════════════════════════════════════════════════════════════════

// Generate a default hash if MANAGER_PASS_HASH is not set
const DEFAULT_PASS_HASH = process.env.MANAGER_PASS_HASH || bcrypt.hashSync('manager123', 10);

app.post('/api/auth/login', async (req, res) => {
  try {
    const { username, password } = req.body;
    const validUser = process.env.MANAGER_USER || 'admin';

    if (username !== validUser) {
      return res.status(401).json({ error: 'Invalid credentials.' });
    }

    const match = await bcrypt.compare(password, DEFAULT_PASS_HASH);
    if (!match) {
      return res.status(401).json({ error: 'Invalid credentials.' });
    }

    const token = signToken({ username, role: 'manager' });
    res.json({ success: true, token, expiresIn: '30m' });
  } catch (err) {
    console.error('Login error:', err.message);
    res.status(500).json({ error: 'Login failed.' });
  }
});

app.get('/api/auth/verify', verifyToken, (req, res) => {
  res.json({ valid: true, user: req.user });
});

// ══════════════════════════════════════════════════════════════════
//  PHASE 2: EXECUTIVE APIs (all JWT-protected)
// ══════════════════════════════════════════════════════════════════

// Middleware to ensure data is loaded in serverless environments
app.use('/api/exec/*', async (req, res, next) => {
  try {
    await execData.ensureLoaded();
    next();
  } catch (err) {
    console.error('Failed to ensure data is loaded:', err.message);
    next();
  }
});

app.get('/api/exec/summary', verifyToken, (req, res) => {
  try {
    res.json(execData.computeSummary());
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/team-capacity', verifyToken, (req, res) => {
  try {
    const team = execData.getData('team');
    const leave = execData.getData('leave');
    res.json({
      team: team || {},
      leave: leave || {},
      // Expose all parsed sub-data from team Excel tabs
      certifications: team?.certifications || {},
      skillMatrix: team?.skillMatrix || [],
      expertise: team?.expertise || [],
      exitResources: team?.exitResources || [],
      roleDistribution: team?.roleDistribution || {},
      designationDistribution: team?.designationDistribution || {},
      historicalTeamSize: team?.historicalTeamSize || {},
      activeMembers: team?.activeMembers || [],
      totalHeadcount: team?.totalHeadcount || 0,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/kpi-scorecards', verifyToken, (req, res) => {
  try {
    const kpi = execData.getData('kpi');
    if (!kpi) return res.json({ kpis: [], metrics: [], summary: {}, highlights: [], lowlights: [] });
    // Add metricName alias for frontend compatibility
    const kpis = (kpi.kpis || []).map(k => ({ ...k, metricName: k.metric || k.name || '' }));
    const metrics = (kpi.metrics || []).map(m => ({ ...m, metricName: m.metric || m.name || '' }));
    res.json({ ...kpi, kpis, metrics });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/sow-financial', verifyToken, (req, res) => {
  try {
    const sow = execData.getData('sow');
    if (!sow) return res.json({ projects: [], summary: {} });
    // Compute projectsByStatus for the SOW doughnut chart
    const projectsByStatus = {};
    (sow.projects || []).forEach(p => {
      const status = p.projectStatus || 'Unknown';
      projectsByStatus[status] = (projectsByStatus[status] || 0) + 1;
    });
    const summary = { ...sow.summary, projectsByStatus };
    res.json({ ...sow, summary });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/governance-risks', verifyToken, (req, res) => {
  try {
    const gov = execData.getData('governance');
    if (!gov) return res.json({ highlights: [], lowlights: [], risks: [], audits: [], fte: [], qbr: [], fteTrend: [], qbrSchedule: [], summary: {} });
    // Map parsed data to frontend-expected shapes
    const fteTrend = (gov.fte || []).map(entry => {
      const months = Object.entries(entry.monthlyFTE || {});
      return months.map(([month, totalFTE]) => ({ account: entry.account, month, totalFTE }));
    }).flat();
    // Aggregate FTE by month for the trend chart
    const fteByMonth = {};
    fteTrend.forEach(t => {
      fteByMonth[t.month] = (fteByMonth[t.month] || 0) + t.totalFTE;
    });
    const fteTrendAgg = Object.entries(fteByMonth).map(([month, totalFTE]) => ({ month, totalFTE }));
    res.json({
      ...gov,
      fteTrend: fteTrendAgg,
      qbrSchedule: gov.qbr || [],
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/leave-impact', verifyToken, (req, res) => {
  try {
    const leave = execData.getData('leave');
    res.json(leave || { currentMonth: {}, previousMonth: {} });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/ftr-metrics', verifyToken, (req, res) => {
  try {
    const ftr = execData.getData('ftr');
    res.json(ftr || { accounts: [], metrics: [], summary: {} });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/project-health', verifyToken, (req, res) => {
  try {
    const health = execData.computeProjectHealth();
    res.json(health);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/poc', verifyToken, (req, res) => {
  try {
    const poc = execData.getData('poc');
    res.json(poc || { pocs: [], aiUsecases: [], summary: {} });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── ALL SHEETS — Parent/Child tab data ────────────────────

/**
 * Smart header-row detection: scan first 10 rows and pick the one that
 * looks most like a header (many non-empty string cells, common keywords).
 * Penalises rows with data-like values (Yes/No, single-word repeats).
 */
function findBestHeaderRow(rows, maxScan = 10) {
  // Only include words that are unambiguously header labels
  const HEADER_WORDS = ['name', 'date', 'day', 'project', 'account', 'status',
    'resources', 'role', 'designation', 'month', 'week', 'client',
    'email', 'phone', 'description', 'kpi', 'metric', 'target', 'actual',
    'highlight', 'lowlight', 'risk', 'audit', 'sow', 'formula',
    'comments', 'impact', 'owner', 'certification', 'expertise',
    'consent', 'privacy', 'cookie', 'emp id', 'verified'];

  // Common data values that should NOT count as headers
  const DATA_VALUES = ['yes', 'no', 'na', 'n/a', 'none', 'beginner',
    'intermediate', 'advanced', 'expert', 'lead', 'no knowledge',
    'received', 'pending', 'active', 'inactive', 'ongoing', 'closed'];

  let bestIdx = 0;
  let bestScore = -1;

  const limit = Math.min(maxScan, rows.length);
  for (let i = 0; i < limit; i++) {
    const row = rows[i];
    if (!row || !Array.isArray(row)) continue;

    let nonEmpty = 0;
    let stringCells = 0;
    let keywordHits = 0;
    let dataValuePenalty = 0;
    let multiWordBonus = 0;

    for (const cell of row) {
      const s = String(cell || '').trim();
      if (!s) continue;
      nonEmpty++;
      if (typeof cell === 'string') {
        stringCells++;
        // Multi-word cells with spaces are likely header labels
        if (s.includes(' ') && s.length > 3) multiWordBonus++;
      }
      const lower = s.toLowerCase();
      if (HEADER_WORDS.some(kw => lower.includes(kw))) {
        keywordHits++;
      }
      // Penalise data-like values
      if (DATA_VALUES.includes(lower) || (!isNaN(Number(s)) && s.length <= 3)) {
        dataValuePenalty++;
      }
    }

    // Score: keyword hits weighted highest, multi-word labels, then strings
    // Subtract penalty for data-like values
    const score = keywordHits * 15 + multiWordBonus * 8 + stringCells * 2 + nonEmpty - dataValuePenalty * 6;
    // Require at least 2 non-empty cells to qualify
    if (nonEmpty >= 2 && score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }
  return bestIdx;
}

/**
 * Convert Excel serial date numbers to readable YYYY-MM-DD strings.
 * Serial numbers in the typical date range (30000-60000) are converted.
 */
function formatCellForDisplay(val) {
  if (val == null || val === '') return '';
  if (typeof val === 'number' && val > 25569 && val < 60000) {
    // Likely an Excel serial date
    try {
      const d = new Date((val - 25569) * 86400000);
      if (!isNaN(d.getTime())) {
        return d.toISOString().slice(0, 10);
      }
    } catch (_) { /* fall through */ }
  }
  return String(val);
}

app.get('/api/exec/all-sheets', verifyToken, (req, res) => {
  try {
    const XLSX = require('xlsx');
    const FRIENDLY = {
      ftr: { name: 'FTR Tracker', icon: '📊' },
      team: { name: 'Team Details', icon: '👥' },
      sow: { name: 'SOW & PO Tracker', icon: '💰' },
      governance: { name: 'Governance', icon: '🏛️' },
      leave: { name: 'Leave Tracker', icon: '🏖️' },
      kpi: { name: 'KPI Metrics', icon: '🎯' },
      poc: { name: 'POV / POC Tracker', icon: '🧪' },
    };

    const sources = [];

    for (const [key, src] of Object.entries(execData.DATA_SOURCES)) {
      const filePath = src.path;
      if (!filePath || !fs.existsSync(filePath)) continue;
      const file = path.basename(filePath);
      const meta = FRIENDLY[key] || { name: key, icon: '📄' };
      try {
        const wb = XLSX.readFile(filePath, { type: 'file' });
        const sheets = [];
        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          // Read all rows as arrays (header: 1) for smart detection
          const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          if (allRows.length === 0) {
            sheets.push({ name: sheetName, headers: [], rows: [], totalRows: 0 });
            continue;
          }

          // Smart header detection
          const headerIdx = findBestHeaderRow(allRows);
          const rawHeaders = allRows[headerIdx] || [];

          // Build clean, unique header names
          const seenHeaders = {};
          const headers = rawHeaders.map((h, colIdx) => {
            let name = String(h || '').trim();
            if (!name) name = `Col_${colIdx + 1}`;
            // De-duplicate header names
            if (seenHeaders[name]) {
              seenHeaders[name]++;
              name = `${name} (${seenHeaders[name]})`;
            } else {
              seenHeaders[name] = 1;
            }
            return name;
          }).filter((_, colIdx) => {
            // Only include columns where the header cell has content
            const h = String(rawHeaders[colIdx] || '').trim();
            return h !== '';
          });

          // Map header names to their column indices
          const headerColMap = [];
          rawHeaders.forEach((h, colIdx) => {
            if (String(h || '').trim() !== '') {
              headerColMap.push(colIdx);
            }
          });

          // Build rows from headerIdx+1 onward
          const dataRows = allRows.slice(headerIdx + 1);
          const totalDataRows = dataRows.length;
          const limitedRows = dataRows.slice(0, 500).map(row => {
            const clean = {};
            headers.forEach((hName, i) => {
              const colIdx = headerColMap[i];
              const val = colIdx < row.length ? row[colIdx] : '';
              clean[hName] = formatCellForDisplay(val);
            });
            return clean;
          });

          sheets.push({
            name: sheetName,
            headers,
            rows: limitedRows,
            totalRows: totalDataRows,
            headerRow: headerIdx,  // expose for debugging
          });
        }
        sources.push({ key, file, name: meta.name, icon: meta.icon, sheets });
      } catch (err) {
        sources.push({ key, file, name: meta.name, icon: meta.icon, sheets: [], error: err.message });
      }
    }

    // Sort by predefined order
    const order = ['ftr', 'team', 'sow', 'governance', 'leave', 'kpi', 'poc'];
    sources.sort((a, b) => (order.indexOf(a.key) - order.indexOf(b.key)));
    res.json({ sources });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Upload Excel files for executive dashboard ────────────────────
const execUploadFields = upload.fields([
  { name: 'ftr', maxCount: 1 },
  { name: 'team', maxCount: 1 },
  { name: 'sow', maxCount: 1 },
  { name: 'governance', maxCount: 1 },
  { name: 'leave', maxCount: 1 },
  { name: 'kpi', maxCount: 1 },
  { name: 'poc', maxCount: 1 },
]);

app.post('/api/exec/upload-sources', verifyToken, (req, res) => {
  execUploadFields(req, res, async (err) => {
    if (err) return res.status(400).json({ error: err.message });
    try {
      const uploaded = [];
      for (const [key, files] of Object.entries(req.files || {})) {
        if (files && files[0]) {
          const ext = path.extname(files[0].originalname).toLowerCase();
          const newPath = path.join(uploadDir, `${key}${ext}`);
          fs.renameSync(files[0].path, newPath);
          execData.setUploadedFile(key, newPath);
          uploaded.push(key);
        }
      }
      await execData.loadAll();
      res.json({ success: true, uploaded, message: `${uploaded.length} file(s) uploaded & data refreshed.` });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
});

// ── Connect SharePoint/cloud URLs ─────────────────────────────────
app.post('/api/exec/connect-sources', verifyToken, async (req, res) => {
  try {
    const { sources } = req.body; // { ftr: { url, type }, team: { url, type }, ... }
    if (!sources || typeof sources !== 'object') {
      return res.status(400).json({ error: 'Sources object is required.' });
    }
    const connected = [];
    for (const [key, cfg] of Object.entries(sources)) {
      if (cfg.url) {
        execData.setRemoteSource(key, cfg.url, cfg.type || 'sharepoint');
        connected.push(key);
      }
    }
    await execData.loadAll();
    res.json({ success: true, connected, message: `${connected.length} source(s) connected & data refreshed.` });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/source-status', verifyToken, (req, res) => {
  res.json(execData.getSourceStatus());
});

// ── Executive data manual refresh ─────────────────────────────────
app.post('/api/exec/refresh', verifyToken, async (req, res) => {
  try {
    await execData.loadAll();
    res.json({ success: true, message: 'Data refreshed.', lastRefresh: new Date().toISOString() });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Data health / parsing diagnostics ────────────────────────────────
app.get('/api/exec/data-health', verifyToken, (req, res) => {
  try {
    res.json(execData.getDiagnostics());
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Helper to safely format Markdown-like output from LLM to HTML
function convertMarkdownToHtml(text) {
  if (!text) return '';
  let html = text;
  // Strip code block wrappers
  html = html.replace(/```html?/gi, '').replace(/```/g, '');
  // Bold
  html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/__(.*?)__/g, '<strong>$1</strong>');
  // Italic
  html = html.replace(/\*(.*?)\*/g, '<em>$1</em>');
  html = html.replace(/_(.*?)_/g, '<em>$1</em>');
  return html.trim();
}

let cachedWorkingGeminiModel = null;

async function processQuestionWithAI(question, data, healthData, userKeys = {}) {
  const q = question.toLowerCase().trim();
  const geminiKey = userKeys.gemini || process.env.GEMINI_API_KEY;
  const openAiKey = userKeys.openai || process.env.OPENAI_API_KEY;

  // Try Google Gemini API if key is present
  if (geminiKey) {
    try {
      const payload = {
        contents: [
          {
            parts: [
              {
                text: `You are the AI Assistant for the Digital Enablement Manager Dashboard at Indegene. Answer the user's question strictly from the live dashboard JSON below — do NOT invent data, do NOT use outside knowledge for facts. If the answer is not in the data, say "I don't have that information in the dashboard yet."

LIVE DASHBOARD JSON (the single source of truth):
${JSON.stringify({
  lastRefresh: data.lastRefresh || null,
  team: {
    totalHeadcount: data.team?.totalHeadcount || (data.team?.activeMembers?.length || 0),
    activeMembers: data.team?.activeMembers || [],
    roleDistribution: data.team?.roleDistribution || {},
    designationDistribution: data.team?.designationDistribution || {},
    certifications: data.team?.certifications || {},
    skillMatrix: (data.team?.skillMatrix || []).slice(0, 50),
    exitResources: data.team?.exitResources || [],
  },
  sow: {
    projects: data.sow?.projects || [],
    summary: data.sow?.summary || {},
  },
  projectHealth: healthData || [],
  kpi: {
    metrics: data.kpi?.metrics || data.kpi?.kpis || [],
    summary: data.kpi?.summary || {},
    highlights: data.kpi?.highlights || [],
    lowlights: data.kpi?.lowlights || [],
  },
  governance: {
    risks: data.governance?.risks || [],
    highlights: data.governance?.highlights || [],
    lowlights: data.governance?.lowlights || [],
    audits: data.governance?.audits || [],
    qbr: data.governance?.qbr || [],
    summary: data.governance?.summary || {},
  },
  leave: {
    currentMonth: data.leave?.currentMonth || {},
    months: Object.keys(data.leave?.months || {}),
  },
  ftr: {
    accounts: data.ftr?.accounts || [],
    qaMetrics: data.ftr?.qaMetrics || [],
    summary: data.ftr?.summary || {},
  },
  poc: {
    pocs: data.poc?.pocs || [],
    aiUsecases: data.poc?.aiUsecases || [],
    summary: data.poc?.summary || {},
  }
}, null, 2)}

User's Question: "${question}"

RESPONSE RULES:
1. Be direct and concise. Answer ONLY what was asked.
2. NAME MATCHING IS FUZZY: when the user names a person/project, match on first name OR partial last name (e.g., "Bhavani" → "Bhavani V" or "Bhavani Venkatesh" — treat them as the SAME person). Same for projects.
3. For a single person: render a compact <ul> with Role, Designation, Manager (RM), Office, Email, Emp ID. Never a table for one person.
4. For reportee lookups ("reportees of X", "X ke under kaun hai"): list ALL members whose manager field contains X's first name (case-insensitive substring match), not exact equality. Show name + designation.
5. For comparisons or multi-row data: use a compact HTML <table> with minimal columns. Add inline style="border-collapse:collapse;width:100%" and <th style="text-align:left;padding:4px 8px;border-bottom:1px solid rgba(255,255,255,0.2)">.
6. KPI numbers: target/actual are decimals (0.95 = 95%). Always render as percentages.
7. FTR passRate is already a percentage number — append "%".
8. Risks: filter status === "Ongoing" when user asks for active/open risks.
9. Project health (Red/Amber/Green) comes from projectHealth array — use this when asked about project status.
10. Use HTML inline tags only: <strong>, <em>, <br>, <ul>, <li>, <table>, <tr>, <th>, <td>, <span>. NO markdown (**, *, \`\`\`). NO outer <html>/<body> wrappers.
11. Match the user's language style (English/Hindi/Hinglish), keep tone professional.
12. End with: <br><br><span style="font-size:0.8em;opacity:0.7">⚡ Powered by Gemini AI</span>`
              }
            ]
          }
        ]
      };
      
      const modelsToTry = cachedWorkingGeminiModel 
        ? [cachedWorkingGeminiModel]
        : [
            'gemini-2.5-flash',
            'gemini-2.0-flash-exp',
            'gemini-1.5-flash',
            'gemini-1.5-flash-latest',
            'gemini-pro',
            'gemini-1.5-pro'
          ];

      let lastErrText = '';
      for (const modelName of modelsToTry) {
        try {
          const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${geminiKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
          });

          if (res.ok) {
            const json = await res.json();
            const text = json.candidates?.[0]?.content?.parts?.[0]?.text;
            if (text) {
              cachedWorkingGeminiModel = modelName; // Cache the successful model
              return convertMarkdownToHtml(text);
            }
          } else {
            lastErrText = await res.text();
            console.warn(`[Gemini API] Model ${modelName} failed. Status: ${res.status}. Error: ${lastErrText}`);
          }
        } catch (e) {
          console.warn(`[Gemini API] Exception trying ${modelName}:`, e.message);
        }
      }

      // If we used a cached model and it failed, clear cache and retry the search
      if (cachedWorkingGeminiModel) {
        console.warn(`[Gemini API] Cached model ${cachedWorkingGeminiModel} failed. Invalidating cache and retrying all models...`);
        cachedWorkingGeminiModel = null;
        return await processQuestionWithAI(question, data, healthData, userKeys);
      }

      // If all models failed, do a diagnostic log of available models
      console.error(`[Gemini API] All models failed. Last error: ${lastErrText}`);
      try {
        const diagRes = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${geminiKey}`);
        if (diagRes.ok) {
          const diagJson = await diagRes.json();
          const available = (diagJson.models || []).map(m => m.name);
          console.log('[Gemini API] Available models for this key:', available);
        }
      } catch (diagErr) {
        console.error('[Gemini API] Could not fetch available models list:', diagErr.message);
      }
    } catch (e) {
      console.error('[Gemini API] Outermost exception:', e);
    }
  }

  // Try OpenAI API if key is present
  if (openAiKey) {
    try {
      const res = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${openAiKey}`
        },
        body: JSON.stringify({
          model: 'gpt-4o-mini',
          messages: [
            {
              role: 'system',
              content: `You are an AI Chatbot for the Digital Enablement Manager Dashboard. Answer queries based on the provided JSON data. Use HTML tags (<strong>, <em>, <br>, <ul>, <li>, <table>) for formatting. No markdown code blocks. Mention you are powered by OpenAI GPT.`
            },
            {
              role: 'user',
              content: `Data:\n${JSON.stringify({
                teamSize: data.team?.activeMembers?.length || 0,
                activeMembers: data.team?.activeMembers || [],
                projects: data.sow?.projects || [],
                health: healthData || [],
                kpiMetrics: data.kpi?.metrics || data.kpi?.kpis || [],
                governance: {
                  risks: data.governance?.risks || [],
                  highlights: data.governance?.highlights || [],
                  audits: data.governance?.audits || []
                },
                leave: data.leave?.currentMonth || {},
                ftr: data.ftr?.qaMetrics || [],
                poc: data.poc?.pocs || []
              })}\n\nQuestion: ${question}`
            }
          ]
        })
      });

      if (res.ok) {
        const json = await res.json();
        const text = json.choices?.[0]?.message?.content;
        if (text) {
          return convertMarkdownToHtml(text);
        }
      } else {
        const errText = await res.text();
        console.error('OpenAI API Error details:', errText);
      }
    } catch (e) {
      console.error('OpenAI API Exception:', e);
    }
  }

  // Fallback: Rule-based regex/keyword matcher
  const answer = processQuestion(q, data, healthData);
  return answer + `<br><br><span style="font-size: 0.85em; opacity: 0.65; display: block; border-top: 1px dashed rgba(255,255,255,0.2); padding-top: 5px; margin-top: 10px;">🤖 <em>Note: Running in Rule-Based mode. Click the ⚙️ settings icon in the chat header to add your Google Gemini API Key and unlock the full power of Gemini AI!</em></span>`;
}

// ── Chatbot API — answers questions from dashboard data ─────────────
app.post('/api/exec/chat', verifyToken, async (req, res) => {
  try {
    const { question } = req.body;
    if (!question || typeof question !== 'string') {
      return res.json({ answer: 'Please ask a valid question.' });
    }
    const q = question.toLowerCase().trim();
    const data = execData.getAllCached();
    const healthData = execData.computeProjectHealth();

    // Retrieve custom API keys passed by client
    const userKeys = {
      gemini: req.headers['x-gemini-api-key'] || '',
      openai: req.headers['x-openai-api-key'] || ''
    };

    const answer = await processQuestionWithAI(question, data, healthData, userKeys);
    res.json({ answer });
  } catch (err) {
    res.json({ answer: '❌ Sorry, I encountered an error processing your question: ' + err.message });
  }
});

function processQuestion(q, data, healthData) {
  const team = data.team || {};
  const sow = data.sow || {};
  const kpi = data.kpi || {};
  const gov = data.governance || {};
  const leave = data.leave || {};
  const ftr = data.ftr || {};
  const health = healthData || [];
  const poc = data.poc || {};
  const members = team.activeMembers || [];
  const risks = gov.risks || [];
  const highlights = gov.highlights || [];
  const lowlights = gov.lowlights || [];
  const projects = sow.projects || [];
  const metrics = kpi.metrics || kpi.kpis || [];
  const leaveData = leave.currentMonth || {};
  const pocs = poc.pocs || [];

  // ── 1. Reportees / under a manager (MUST be before generic person lookup) ──
  const isReporteeQ = matchWords(q, ['reportee', 'reportees', 'reporting', 'repotee', 'repotees'])
    || q.includes('under ') || q.includes('ke niche') || q.includes('ke under');
  if (isReporteeQ) {
    const mgr = findPersonInQuery(q, members);
    if (mgr) {
      const reportees = findReportees(mgr, members);
      if (reportees.length === 0) return `No reportees found under <strong>${mgr.name}</strong>.`;
      let html = `👥 <strong>${mgr.name}'s Reportees (${reportees.length})</strong><br><br>`;
      reportees.forEach(r => { html += `• ${r.name} — ${r.designation || r.role || '—'}<br>`; });
      return html;
    }
    if (matchWords(q, ['reportee', 'reportees', 'reporting'])) {
      return '🤔 Could not identify the manager name. Try: <em>"reportees of [full name]"</em>';
    }
  }

  // ── 2. Specific person lookup ──
  if (matchWords(q, ['who is', 'kaun hai', 'tell me about', 'details of', 'find', 'kaun h'])) {
    const person = findPersonInQuery(q, members);
    if (person) {
      let html = `👤 <strong>${person.name}</strong><br>`;
      html += `Role: ${person.role || '—'}<br>`;
      html += `Designation: ${person.designation || '—'}<br>`;
      html += `Manager (RM): ${person.manager || '—'}<br>`;
      html += `Office: ${person.location || '—'}<br>`;
      html += `Email: ${person.email || '—'}<br>`;
      html += `Emp ID: ${person.empId || '—'}<br>`;
      html += `DOJ: ${person.doj || '—'}`;
      return html;
    }
  }

  // ── 3. Team count / team info ──
  if (matchWords(q, ['team size', 'team member', 'headcount', 'total member', 'kitne log', 'kitne member', 'how many people', 'how many member', 'team count', 'total team'])) {
    const total = members.length;
    const offices = {};
    members.forEach(m => { if (m.location) offices[m.location] = (offices[m.location] || 0) + 1; });
    const officeStr = Object.entries(offices).map(([k, v]) => `${k}: ${v}`).join(', ');
    return `👥 <strong>Team Size: <span class="chat-stat">${total}</span></strong><br><br>` +
      `<strong>Office-wise:</strong> ${officeStr || 'N/A'}<br><br>` +
      `<strong>Designations:</strong> ${[...new Set(members.map(m => m.designation).filter(Boolean))].slice(0, 8).join(', ')}`;
  }

  // ── 4. Leave info ──
  if (matchWords(q, ['leave', 'chutti', 'absent', 'leave today', 'aaj kaun', 'on leave', 'kon leave'])) {
    const onToday = leaveData.onLeaveToday || [];
    const totalLeaves = leaveData.totalLeaves || 0;
    const byPerson = leaveData.byPerson || {};
    let html = `🏖️ <strong>Leave Summary — This Month</strong><br>`;
    html += `Total leaves: <span class="chat-stat">${totalLeaves}</span><br><br>`;
    if (onToday.length > 0) {
      html += `<strong>On Leave Today (${onToday.length}):</strong><br>`;
      onToday.forEach(p => { html += `• ${p.name} (${p.type || 'Leave'})<br>`; });
    } else {
      html += `✅ <em>Nobody on leave today!</em><br>`;
    }
    const topLeave = Object.entries(byPerson).sort((a, b) => b[1] - a[1]).slice(0, 5);
    if (topLeave.length > 0) {
      html += `<br><strong>Top Leave Takers:</strong><br>`;
      topLeave.forEach(([name, days]) => { html += `• ${name}: ${days} day(s)<br>`; });
    }
    return html;
  }

  // ── 5. Risks ──
  if (matchWords(q, ['risk', 'risks', 'issue', 'problem', 'challenge', 'blocker', 'dikkat'])) {
    const active = risks.filter(r => r.status?.toLowerCase() === 'ongoing' || !r.status);
    if (!active.length) return '✅ <strong>No active risks found!</strong> Everything is looking good.';
    let html = `⚠️ <strong>Active Risks: <span class="chat-stat">${active.length}</span></strong><br><br>`;
    active.slice(0, 6).forEach(r => {
      html += `<strong>${r.project || 'Unknown'}</strong> — ${r.pm || ''}<br>`;
      html += `${r.risk || r.description || 'No details'}<br>`;
      if (r.impact) html += `Impact: <em>${r.impact}</em><br>`;
      if (r.mitigation) html += `Mitigation: <em>${r.mitigation}</em><br>`;
      html += `<br>`;
    });
    if (active.length > 6) html += `<em>...and ${active.length - 6} more risks</em>`;
    return html;
  }

  // ── 6. KPI / performance ──
  if (matchWords(q, ['kpi', 'performance', 'sla', 'utilization', 'quality', 'csat', 'training', 'metric'])) {
    if (!metrics.length) return '📊 No KPI data available.';
    let html = `📊 <strong>KPI Performance Summary</strong><br><br>`;
    const grouped = {};
    metrics.forEach(m => {
      const proj = m.project || m.account || 'General';
      if (!grouped[proj]) grouped[proj] = [];
      grouped[proj].push(m);
    });
    const projKeys = Object.keys(grouped).slice(0, 4);
    projKeys.forEach(proj => {
      html += `<strong>${proj}:</strong><br>`;
      grouped[proj].slice(0, 6).forEach(m => {
        const target = m.target != null ? (m.target * 100).toFixed(0) + '%' : '—';
        const actual = m.actual != null ? (m.actual * 100).toFixed(0) + '%' : '—';
        const status = m.actual >= m.target ? '✅' : '⚠️';
        html += `${status} ${m.metricName || m.kpiName || 'Metric'}: Target ${target}, Actual ${actual}<br>`;
      });
      html += `<br>`;
    });
    if (projKeys.length < Object.keys(grouped).length) html += `<em>...${Object.keys(grouped).length - projKeys.length} more projects</em>`;
    return html;
  }

  // ── 7. SOW / PO / contract ──
  if (matchWords(q, ['sow', 'purchase order', 'statement of work', 'financial', 'billing', 'contract'])) {
    if (!projects.length) return '📋 No SOW data available.';
    let html = `📋 <strong>SOW & PO Summary</strong><br>`;
    html += `Total Projects: <span class="chat-stat">${projects.length}</span><br><br>`;
    const statusCount = {};
    projects.forEach(p => {
      const s = p.sowStatus || p.status || 'Unknown';
      statusCount[s] = (statusCount[s] || 0) + 1;
    });
    html += `<strong>Status Breakdown:</strong><br>`;
    Object.entries(statusCount).forEach(([s, c]) => { html += `• ${s}: ${c}<br>`; });
    const projName = findProjectInQuery(q, projects);
    if (projName) {
      const p = projects.find(pr => pr.projectName?.toLowerCase() === projName.toLowerCase());
      if (p) {
        html += `<br><strong>📂 ${p.projectName}:</strong><br>`;
        html += `Client: ${p.client || '—'}<br>SOW: ${p.sowStatus || '—'}<br>PO: ${p.poStatus || '—'}<br>`;
        if (p.sowValue) html += `Value: ${p.sowValue}<br>`;
        if (p.startDate) html += `Start: ${p.startDate}<br>`;
        if (p.endDate) html += `End: ${p.endDate}<br>`;
      }
    }
    return html;
  }

  // ── 8. FTR / QA metrics ──
  if (matchWords(q, ['ftr', 'first time right', 'qa metric', 'pass rate'])) {
    const qaMetrics = ftr.qaMetrics || ftr.metrics || [];
    if (!qaMetrics.length) return '✅ No FTR data available.';
    let html = `✅ <strong>FTR / QA Summary</strong><br>`;
    html += `Total QA metrics: <span class="chat-stat">${qaMetrics.length}</span><br><br>`;
    qaMetrics.slice(0, 8).forEach(m => {
      const rate = m.passRate != null ? `${Number(m.passRate).toFixed(1)}%` : (m.ftrRate || '—');
      const label = m.project || m.client || m.account || 'Unknown';
      const when = m.qaDate || m.month || '';
      html += `• ${label}${when ? ' — ' + when : ''}: <strong>${rate}</strong><br>`;
    });
    return html;
  }

  // ── 9. Project health / status ──
  if (matchWords(q, ['project', 'health', 'active project', 'kitne project', 'project list', 'projects'])) {
    if (!health.length && !projects.length) return '🏥 No project data available.';
    const src = health.length ? health : projects;
    let html = `🏥 <strong>Project Health</strong><br>`;
    html += `Total: <span class="chat-stat">${src.length}</span><br><br>`;
    const hCount = { Green: 0, Amber: 0, Red: 0 };
    src.forEach(p => { const h = p.health || 'Unknown'; if (hCount[h] !== undefined) hCount[h]++; });
    html += `🟢 Green: ${hCount.Green} &nbsp; 🟡 Amber: ${hCount.Amber} &nbsp; 🔴 Red: ${hCount.Red}<br><br>`;
    const redProjects = src.filter(p => p.health === 'Red');
    if (redProjects.length > 0) {
      html += `<strong>⚠️ Red Status Projects:</strong><br>`;
      redProjects.forEach(p => {
        html += `• <strong>${p.projectName || p.project}</strong> — ${p.reasons || p.pm || 'No details'}<br>`;
      });
    }
    return html;
  }

  // ── 10. Highlights / lowlights ──
  if (matchWords(q, ['highlight', 'highlights', 'achievement', 'achievements', 'good news', 'lowlight', 'lowlights', 'concern', 'concerns'])) {
    const hl = highlights.filter(h => h.highlight);
    const ll = lowlights.filter(h => h.lowlight);
    let html = `🌟 <strong>Recent Highlights (${hl.length})</strong><br>`;
    if (!hl.length) html += `<em>None recorded.</em><br>`;
    hl.slice(0, 5).forEach(h => { html += `• <strong>${h.project}</strong> (${h.month || ''}): ${h.highlight}<br>`; });
    html += `<br>⬇️ <strong>Recent Lowlights (${ll.length})</strong><br>`;
    if (!ll.length) html += `<em>None recorded.</em><br>`;
    ll.slice(0, 5).forEach(h => { html += `• <strong>${h.project}</strong> (${h.month || ''}): ${h.lowlight}<br>`; });
    return html;
  }

  // ── 11. POC / POV ──
  if (matchWords(q, ['poc', 'pov', 'proof of concept', 'proof of value'])) {
    if (!pocs.length) return '🧪 No POC/POV data available.';
    let html = `🧪 <strong>POC/POV Summary</strong><br>Total: <span class="chat-stat">${pocs.length}</span><br><br>`;
    pocs.slice(0, 5).forEach(p => {
      html += `• <strong>${p.title || p.project || 'Untitled'}</strong> — ${p.status || 'Unknown'} (SPOC: ${p.spoc || '—'})<br>`;
    });
    return html;
  }

  // ── 12. Summary / overview ──
  if (matchWords(q, ['summary', 'overview', 'dashboard', 'overall', 'sabkuch', 'sab batao', 'everything'])) {
    let html = `📊 <strong>Dashboard Summary</strong><br><br>`;
    html += `👥 Team Members: <span class="chat-stat">${members.length}</span><br>`;
    html += `📂 Projects: <span class="chat-stat">${projects.length || health.length || 0}</span><br>`;
    html += `⚠️ Active Risks: <span class="chat-stat">${risks.filter(r => r.status?.toLowerCase() === 'ongoing').length}</span><br>`;
    html += `🏖️ On Leave Today: <span class="chat-stat">${(leaveData.onLeaveToday || []).length}</span><br>`;
    html += `📊 KPI Metrics: <span class="chat-stat">${metrics.length}</span><br>`;
    html += `📋 SOW/PO Count: <span class="chat-stat">${projects.length}</span><br>`;
    html += `🧪 POC/POV: <span class="chat-stat">${pocs.length}</span><br>`;
    html += `🌟 Highlights: <span class="chat-stat">${highlights.filter(h => h.highlight).length}</span><br>`;
    html += `⬇️ Lowlights: <span class="chat-stat">${lowlights.filter(h => h.lowlight).length}</span>`;
    return html;
  }

  // ── 13. Fallback: try to find a person by name anywhere in the query ──
  const personFallback = findPersonInQuery(q, members);
  if (personFallback) {
    let html = `👤 <strong>${personFallback.name}</strong><br>`;
    html += `Role: ${personFallback.role || '—'}<br>`;
    html += `Designation: ${personFallback.designation || '—'}<br>`;
    html += `Manager (RM): ${personFallback.manager || '—'}<br>`;
    html += `Office: ${personFallback.location || '—'}<br>`;
    html += `Email: ${personFallback.email || '—'}`;
    return html;
  }

  // ── Help / unknown ──
  return `🤔 I'm not sure about that. Try asking:<br><br>` +
    `• <strong>"team size"</strong> — how many members<br>` +
    `• <strong>"who is on leave"</strong> — leave status<br>` +
    `• <strong>"show risks"</strong> — active risks<br>` +
    `• <strong>"kpi summary"</strong> — performance metrics<br>` +
    `• <strong>"sow status"</strong> — contracts<br>` +
    `• <strong>"project health"</strong> — project overview<br>` +
    `• <strong>"summary"</strong> — full overview<br>` +
    `• <strong>"who is [name]"</strong> — person details<br>` +
    `• <strong>"reportees of [name]"</strong> — team under a manager`;
}

/** Word-boundary aware matching — prevents 'po' matching inside 'reportee' */
function matchWords(q, phrases) {
  for (const phrase of phrases) {
    // For multi-word phrases, check direct inclusion
    if (phrase.includes(' ')) {
      if (q.includes(phrase)) return true;
      continue;
    }
    // For single words, use word boundary regex
    const regex = new RegExp('\\b' + phrase.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b', 'i');
    if (regex.test(q)) return true;
  }
  return false;
}

/** Reserved tokens that should never be treated as a person's name */
const RESERVED_TOKENS = new Set([
  'team', 'kpi', 'sow', 'poc', 'pov', 'ftr', 'qbr', 'leave', 'risk', 'risks',
  'project', 'projects', 'health', 'metric', 'metrics', 'status', 'summary',
  'overview', 'dashboard', 'client', 'clients', 'manager', 'reportee', 'highlight',
  'lowlight', 'who', 'what', 'how', 'when', 'where', 'why', 'show', 'list', 'give',
  'tell', 'about', 'find', 'details', 'total', 'count', 'active', 'ongoing',
]);

function findPersonInQuery(q, members) {
  const qLower = (q || '').toLowerCase();
  // 1. Full-name substring match (most accurate)
  let best = null;
  let bestLen = 0;
  for (const m of members) {
    if (!m.name) continue;
    const nm = m.name.toLowerCase();
    if (qLower.includes(nm) && nm.length > bestLen) { best = m; bestLen = nm.length; }
  }
  if (best) return best;
  // 2. First-name word-boundary match (≥4 chars, not a reserved token)
  for (const m of members) {
    const firstName = (m.name || '').split(' ')[0].toLowerCase();
    if (firstName.length < 4 || RESERVED_TOKENS.has(firstName)) continue;
    const re = new RegExp('\\b' + firstName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b', 'i');
    if (re.test(qLower)) return m;
  }
  return null;
}

/** Find reportees by fuzzy matching the manager's first name against each member's manager field */
function findReportees(mgr, members) {
  if (!mgr || !mgr.name) return [];
  const mgrFirst = mgr.name.split(' ')[0].toLowerCase();
  const mgrFull = mgr.name.toLowerCase();
  return members.filter(m => {
    const mgrField = (m.manager || '').toLowerCase().trim();
    if (!mgrField) return false;
    return mgrField === mgrFull
      || mgrField.includes(mgrFull)
      || mgrFull.includes(mgrField)
      || mgrField.split(/\s+/)[0] === mgrFirst;
  });
}

function findProjectInQuery(q, projects) {
  for (const p of projects) {
    const name = (p.projectName || '').toLowerCase();
    if (name && name.length > 3 && q.includes(name)) return p.projectName;
  }
  return null;
}

// ── Fallback → SPA ───────────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ── Export for Vercel serverless ──────────────────────────────────────
module.exports = app;

// ── Only listen when running locally (not on Vercel) ─────────────────
if (!process.env.VERCEL) {
  app.listen(PORT, async () => {
    console.log(`Dashboard server running at http://localhost:${PORT}`);
    console.log(`Excel path: ${EXCEL_PATH}`);
    // Load executive data sources on startup
    try {
      await execData.loadAll();
      console.log('Executive data sources loaded on startup.');
    } catch (e) {
      console.warn('Could not load exec data on startup:', e.message);
    }

    // ── Auto-refresh connected sources every 2 minutes ──
    // This ensures that when Google Sheets / SharePoint Excel is updated,
    // the changes flow into the dashboard automatically.
    const AUTO_REFRESH_MS = 2 * 60 * 1000; // 2 minutes
    setInterval(async () => {
      try {
        const hasRemote = Object.values(execData.DATA_SOURCES).some(
          s => s.type !== 'local' && s.url
        );
        if (hasRemote) {
          console.log('[Auto-Refresh] Re-downloading connected sources...');
          await execData.loadAll();
          console.log('[Auto-Refresh] Data refreshed at', new Date().toISOString());
        }
      } catch (e) {
        console.warn('[Auto-Refresh] Error:', e.message);
      }
    }, AUTO_REFRESH_MS);
  });
}

