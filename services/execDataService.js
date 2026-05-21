const path = require('path');
const fs = require('fs');
const os = require('os');
const { execSync } = require('child_process');

// ══════════════════════════════════════════════════════════════════
//  DATA SOURCE CONFIGURATION
// ══════════════════════════════════════════════════════════════════

const UPLOADS_DIR = path.join(__dirname, '..', 'uploads');
const WRITABLE_UPLOADS_DIR = path.join(os.tmpdir(), 'dashboard-uploads');

// Ensure writable uploads directory exists
if (!fs.existsSync(WRITABLE_UPLOADS_DIR)) {
    try {
        fs.mkdirSync(WRITABLE_UPLOADS_DIR, { recursive: true });
    } catch (e) {
        console.warn(`[ExecData] Could not create writable uploads dir:`, e.message);
    }
}

const DATA_SOURCES = {
    ftr: { path: process.env.EXEC_FTR_PATH || '', type: 'local', url: '' },
    team: { path: process.env.EXEC_TEAM_PATH || '', type: 'local', url: '' },
    sow: { path: process.env.EXEC_SOW_PATH || '', type: 'local', url: '' },
    governance: { path: process.env.EXEC_GOVERNANCE_PATH || '', type: 'local', url: '' },
    leave: { path: process.env.EXEC_LEAVE_PATH || '', type: 'local', url: '' },
    kpi: { path: process.env.EXEC_KPI_PATH || '', type: 'local', url: '' },
    poc: { path: process.env.EXEC_POC_PATH || '', type: 'local', url: '' },
};

// Auto-detect previously uploaded files in the writable uploads directory, or fallback to bundled uploads
for (const key of Object.keys(DATA_SOURCES)) {
    if (!DATA_SOURCES[key].path) {
        let match = null;
        if (fs.existsSync(WRITABLE_UPLOADS_DIR)) {
            const files = fs.readdirSync(WRITABLE_UPLOADS_DIR);
            match = files.find(f => f.toLowerCase().startsWith(key + '.'));
            if (match) {
                DATA_SOURCES[key].path = path.join(WRITABLE_UPLOADS_DIR, match);
                console.log(`[ExecData] Auto-detected writable upload: ${key} → ${match}`);
                continue;
            }
        }
        if (fs.existsSync(UPLOADS_DIR)) {
            const files = fs.readdirSync(UPLOADS_DIR);
            match = files.find(f => f.toLowerCase().startsWith(key + '.'));
            if (match) {
                DATA_SOURCES[key].path = path.join(UPLOADS_DIR, match);
                console.log(`[ExecData] Auto-detected bundled fallback: ${key} → ${match}`);
            }
        }
    }
}

const FRIENDLY_NAMES = {
    ftr: 'FTR Tracker',
    team: 'Team Details',
    sow: 'SOW & PO Tracker',
    governance: 'Governance',
    leave: 'Leave Tracker',
    kpi: 'KPI',
    poc: 'POV / POC Tracker',
};

// ══════════════════════════════════════════════════════════════════
//  IN-MEMORY CACHE
// ══════════════════════════════════════════════════════════════════

const cache = {
    ftr: null,
    team: null,
    sow: null,
    governance: null,
    leave: null,
    kpi: null,
    poc: null,
    lastRefresh: null,
};

// Store diagnostics from the Python normalizer
const diagnosticsCache = {};

// ══════════════════════════════════════════════════════════════════
//  PATHS
// ══════════════════════════════════════════════════════════════════

const PYTHON_SCRIPT = path.join(__dirname, '..', 'python', 'normalizer.py');
const DATA_DIR = path.join(os.tmpdir(), 'dashboard-data');

// Ensure data directory exists
if (!fs.existsSync(DATA_DIR)) {
    try {
        fs.mkdirSync(DATA_DIR, { recursive: true });
    } catch (e) {
        console.warn(`[ExecData] Could not create data dir:`, e.message);
    }
}


// ══════════════════════════════════════════════════════════════════
//  PYTHON NORMALIZER BRIDGE
// ══════════════════════════════════════════════════════════════════

/**
 * Detect Python command available on this system.
 */
let _pythonCmd = null;
let _pythonChecked = false;
function getPythonCmd() {
    if (process.env.VERCEL) {
        throw new Error('Python not found (disabled in Vercel environment).');
    }
    if (_pythonChecked) {
        if (_pythonCmd) return _pythonCmd;
        throw new Error('Python not found (cached check result).');
    }
    _pythonChecked = true;
    for (const cmd of ['python', 'python3', 'py']) {
        try {
            execSync(`${cmd} --version`, { stdio: 'pipe', timeout: 1000 });
            _pythonCmd = cmd;
            return cmd;
        } catch (_) { /* try next */ }
    }
    throw new Error('Python not found. Install Python 3.8+ and ensure it is on PATH.');
}

/**
 * Run the Python normalizer for a single data source.
 * Returns parsed data or null on failure.
 */
function pythonParse(key) {
    const src = DATA_SOURCES[key];
    const filePath = src.path;
    if (!filePath || !fs.existsSync(filePath)) {
        console.warn(`[ExecData] File not found for ${FRIENDLY_NAMES[key]}: ${filePath}`);
        return null;
    }

    const outputPath = path.join(DATA_DIR, `${key}.json`);
    const pythonCmd = getPythonCmd();

    try {
        // Call Python normalizer
        const cmd = `${pythonCmd} "${PYTHON_SCRIPT}" --source ${key} --input "${filePath}" --output "${outputPath}"`;
        const stdout = execSync(cmd, {
            timeout: 180000, // 180 seconds max per file
            encoding: 'utf-8',
            stdio: ['pipe', 'pipe', 'pipe'],
        });

        // Parse the summary printed to stdout
        let summary = {};
        try { summary = JSON.parse(stdout.trim()); } catch (_) { /* ignore */ }

        if (!summary.success) {
            console.error(`[ExecData] Python normalizer FAILED for ${FRIENDLY_NAMES[key]}: ${summary.error || 'unknown error'}`);
            diagnosticsCache[key] = { success: false, error: summary.error || 'unknown error' };
            return null;
        }

        // Read the full output JSON
        if (!fs.existsSync(outputPath)) {
            console.error(`[ExecData] Output file not created for ${FRIENDLY_NAMES[key]}`);
            return null;
        }

        const output = JSON.parse(fs.readFileSync(outputPath, 'utf-8'));

        // Store diagnostics
        diagnosticsCache[key] = {
            success: true,
            counts: summary.counts || {},
            diagnostics: output.diagnostics || {},
            timestamp: output.timestamp,
        };

        console.log(`[ExecData] Parsed ${FRIENDLY_NAMES[key]} via Python ✓`, summary.counts ? JSON.stringify(summary.counts) : '');
        return output.data;

    } catch (err) {
        const stderr = err.stderr ? err.stderr.toString().slice(0, 500) : '';
        console.error(`[ExecData] Error running Python normalizer for ${FRIENDLY_NAMES[key]}:`, err.message);
        if (stderr) console.error(`[ExecData] Python stderr: ${stderr}`);
        diagnosticsCache[key] = { success: false, error: err.message, stderr };
        return null;
    }
}

/**
 * Fallback: use old Node.js parsers if Python is not available.
 */
function nodeFallbackParse(key) {
    const src = DATA_SOURCES[key];
    const filePath = src.path;
    if (!filePath || !fs.existsSync(filePath)) {
        console.warn(`[ExecData] File not found for ${FRIENDLY_NAMES[key]}: ${filePath}`);
        return null;
    }

    try {
        let parserFn;
        switch (key) {
            case 'ftr': parserFn = require('../parsers/ftrParser').parseFTR; break;
            case 'team': parserFn = require('../parsers/teamParser').parseTeamDetails; break;
            case 'sow': parserFn = require('../parsers/sowParser').parseSOW; break;
            case 'governance': parserFn = require('../parsers/governanceParser').parseGovernance; break;
            case 'leave': parserFn = require('../parsers/leaveParser').parseLeaveTracker; break;
            case 'kpi': parserFn = require('../parsers/kpiParser').parseKPI; break;
            case 'poc': parserFn = require('../parsers/pocParser').parsePoC; break;
            default: return null;
        }
        const result = parserFn(filePath);
        console.log(`[ExecData] Parsed ${FRIENDLY_NAMES[key]} via Node fallback ✓`);
        diagnosticsCache[key] = { success: true, method: 'node-fallback' };
        return result;
    } catch (err) {
        console.error(`[ExecData] Node fallback also failed for ${FRIENDLY_NAMES[key]}:`, err.message);
        diagnosticsCache[key] = { success: false, error: err.message, method: 'node-fallback' };
        return null;
    }
}

/**
 * Parse a source — try Python first, fall back to Node.
 */
function safeParse(key) {
    // Try Python normalizer first
    let result = null;
    let usedPython = false;
    try {
        getPythonCmd(); // will throw if Python not found
        result = pythonParse(key);
        usedPython = true;
    } catch (err) {
        if (err.message.includes('Python not found')) {
            console.warn(`[ExecData] Python not available, using Node.js fallback for all sources`);
        }
    }

    // If Python failed or not available, use Node fallback
    if (result === null && !usedPython) {
        result = nodeFallbackParse(key);
    } else if (result === null && usedPython) {
        // Python was available but parsing failed — try Node fallback too
        console.warn(`[ExecData] Python failed for ${FRIENDLY_NAMES[key]}, trying Node fallback...`);
        result = nodeFallbackParse(key);
    }

    return result;
}

// ══════════════════════════════════════════════════════════════════
//  REMOTE DOWNLOAD
// ══════════════════════════════════════════════════════════════════

async function downloadRemoteFile(key) {
    const src = DATA_SOURCES[key];
    if (src.type !== 'sharepoint' && src.type !== 'google') return;
    if (!src.url) return;

    const uploadDir = WRITABLE_UPLOADS_DIR;
    if (!fs.existsSync(uploadDir)) {
        try {
            fs.mkdirSync(uploadDir, { recursive: true });
        } catch (e) {
            console.warn(`[ExecData] Could not create remote download uploads dir:`, e.message);
        }
    }

    let downloadUrl = src.url;
    if (src.type === 'google') {
        const m = src.url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
        if (m) downloadUrl = `https://docs.google.com/spreadsheets/d/${m[1]}/export?format=xlsx`;
    }

    try {
        const resp = await fetch(downloadUrl);
        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        const buffer = Buffer.from(await resp.arrayBuffer());
        const ext = key === 'leave' ? '.xlsb' : '.xlsx';
        const localPath = path.join(uploadDir, `${key}${ext}`);
        fs.writeFileSync(localPath, buffer);
        src.path = localPath;
        console.log(`[ExecData] Downloaded ${FRIENDLY_NAMES[key]} from ${src.type} (${(buffer.length / 1024).toFixed(1)} KB)`);
    } catch (err) {
        console.error(`[ExecData] Download failed for ${FRIENDLY_NAMES[key]}:`, err.message);
    }
}

// ══════════════════════════════════════════════════════════════════
//  PUBLIC API
// ══════════════════════════════════════════════════════════════════

async function loadAll() {
    // Download remote sources first
    for (const key of Object.keys(DATA_SOURCES)) {
        if (DATA_SOURCES[key].type !== 'local') {
            await downloadRemoteFile(key);
        }
    }

    // Parse all sources (Python first, Node fallback)
    for (const key of Object.keys(DATA_SOURCES)) {
        cache[key] = safeParse(key);
    }
    cache.lastRefresh = new Date().toISOString();

    console.log(`[ExecData] All sources loaded at ${cache.lastRefresh}`);
}

async function ensureLoaded() {
    if (!cache.lastRefresh) {
        console.log('[ExecData] Cache empty, loading all sources lazily...');
        await loadAll();
    }
}

function getData(key) {
    return cache[key];
}

function getAllCached() {
    return cache;
}

function getSourceStatus() {
    const status = {};
    for (const [key, src] of Object.entries(DATA_SOURCES)) {
        status[key] = {
            name: FRIENDLY_NAMES[key],
            type: src.type,
            path: src.path,
            url: src.url || '',
            loaded: cache[key] !== null,
        };
    }
    status.lastRefresh = cache.lastRefresh;
    return status;
}

/**
 * Get diagnostics from the last parse run.
 */
function getDiagnostics() {
    const result = {};
    for (const [key, diag] of Object.entries(diagnosticsCache)) {
        result[key] = {
            name: FRIENDLY_NAMES[key],
            ...diag,
        };
    }
    result.lastRefresh = cache.lastRefresh;
    return result;
}

/**
 * Update a data source to use an uploaded file.
 */
function setUploadedFile(key, filePath) {
    if (DATA_SOURCES[key]) {
        DATA_SOURCES[key].type = 'local';
        DATA_SOURCES[key].path = filePath;
        DATA_SOURCES[key].url = '';
    }
}

/**
 * Update a data source to use a SharePoint/cloud URL.
 */
function setRemoteSource(key, url, type = 'sharepoint') {
    if (DATA_SOURCES[key]) {
        DATA_SOURCES[key].type = type;
        DATA_SOURCES[key].url = url;
    }
}

/**
 * Compute cross-file project health status.
 */
function computeProjectHealth() {
    const sowData = cache.sow;
    const govData = cache.governance;
    const kpiData = cache.kpi;

    if (!sowData || !Array.isArray(sowData.projects)) return [];

    return sowData.projects.map(p => {
        let health = 'Green';
        let reasons = [];

        const sowStatus = (p.sowStatus || '').toLowerCase();
        const poStatus = (p.poStatus || '').toLowerCase();
        const projectStatus = (p.projectStatus || '').toLowerCase();
        const projectName = (p.projectName || '').toLowerCase();

        // Check SOW/PO status
        if (sowStatus !== 'received') {
            health = 'Amber';
            reasons.push('SOW pending');
        }
        if (poStatus !== 'received' && poStatus !== 'ytr') {
            health = 'Amber';
            reasons.push('PO pending');
        }

        // Check risks
        if (govData && Array.isArray(govData.risks)) {
            const projectRisks = govData.risks.filter(r =>
                (r.project || '').toLowerCase() === projectName &&
                (r.status || '').toLowerCase() === 'ongoing'
            );
            if (projectRisks.length > 0) {
                health = projectRisks.some(r => (r.impact || '').toLowerCase().includes('loss')) ? 'Red' : 'Amber';
                reasons.push(`${projectRisks.length} active risk(s)`);
            }
        }

        // Check project status
        if (projectStatus === 'yet to start') {
            health = 'Amber';
            reasons.push('Not yet started');
        }

        return {
            projectName: p.projectName,
            client: p.client,
            pm: p.pm,
            status: p.projectStatus,
            sowStatus: p.sowStatus,
            poStatus: p.poStatus,
            health,
            reasons: reasons.join('; '),
        };
    });
}

/**
 * Build the top-level executive summary.
 */
function computeSummary() {
    const team = cache.team;
    const sow = cache.sow;
    const kpi = cache.kpi;
    const ftr = cache.ftr;
    const leave = cache.leave;
    const gov = cache.governance;

    return {
        teamSize: team?.totalHeadcount || 0,
        activeProjects: sow?.summary?.activeProjects || 0,
        totalSOWValue: sow?.summary?.totalSOWValue || 0,
        kpiMetRate: kpi?.summary?.metRate || 0,
        ftrAvgRating: ftr?.summary?.avgRating || 0,
        onLeaveToday: leave?.currentMonth?.onLeaveToday?.length || 0,
        activeRisks: gov?.summary?.activeRisks || 0,
        lastRefresh: cache.lastRefresh,
    };
}

module.exports = {
    loadAll,
    ensureLoaded,
    getData,
    getAllCached,
    getSourceStatus,
    getDiagnostics,
    setUploadedFile,
    setRemoteSource,
    computeProjectHealth,
    computeSummary,
    DATA_SOURCES,
    FRIENDLY_NAMES,
};
