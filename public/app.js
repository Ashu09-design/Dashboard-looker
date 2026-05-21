/* ═══════════════════════════════════════════════════════════════
   Digital Enablement Team Dashboard — app.js
   ═══════════════════════════════════════════════════════════════ */

let bandwidthData = [];
let qbrData = [];
let dropdownData = [];
let sortAsc = true;
let presentMode = false;

const REFRESH_MS = 5 * 60 * 1000;

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

document.addEventListener('DOMContentLoaded', () => {
    fetchAll();
    bindUI();
    setInterval(fetchAll, REFRESH_MS);
});

// ══════════════════════════════════════════════════════════════════
//  DATA FETCHING
// ══════════════════════════════════════════════════════════════════

async function fetchAll() {
    try {
        showLoading(true);
        const [bw, qbr, dd] = await Promise.all([
            fetch('/api/bandwidth').then(r => r.json()),
            fetch('/api/qbr').then(r => r.json()),
            fetch('/api/dropdown').then(r => r.json()),
        ]);
        bandwidthData = bw;
        qbrData = qbr;
        dropdownData = dd;

        populateProjectFilter();
        applyFilters();
        updateTimestamp();
        showLoading(false);
    } catch (err) {
        console.error('Fetch error:', err);
        showLoading(false);
    }
}

function showLoading(show) {
    const el = $('#loadingOverlay');
    if (show) el.classList.remove('hidden');
    else el.classList.add('hidden');
}

function updateTimestamp() {
    const el = $('#lastUpdated');
    if (el) el.textContent = 'Updated: ' + new Date().toLocaleTimeString();
}

// ══════════════════════════════════════════════════════════════════
//  UI BINDINGS
// ══════════════════════════════════════════════════════════════════

function bindUI() {
    $('#dateRangeBtn').addEventListener('click', () => {
        $('#dateRangeDropdown').classList.toggle('hidden');
    });
    $('#applyDateRange').addEventListener('click', () => {
        const from = $('#filterDateFrom').value;
        const to = $('#filterDateTo').value;
        if (from && to) {
            $('#dateRangeLabel').textContent = fmtDisplay(from) + ' – ' + fmtDisplay(to);
        }
        $('#dateRangeDropdown').classList.add('hidden');
        applyFilters();
    });

    $('#globalSearch').addEventListener('input', applyFilters);

    $('#sortName').addEventListener('click', () => {
        sortAsc = !sortAsc;
        $('#sortName').querySelector('.sort-arrow').textContent = sortAsc ? '▼' : '▲';
        applyFilters();
    });

    // Refresh
    $('#btnRefresh').addEventListener('click', () => {
        const btn = $('#btnRefresh');
        btn.classList.add('refreshing');
        fetchAll().then(() => setTimeout(() => btn.classList.remove('refreshing'), 800));
    });

    // Present mode
    $('#btnPresent').addEventListener('click', () => togglePresentMode(true));
    $('#exitPresent').addEventListener('click', () => togglePresentMode(false));

    // Data Source modal
    $('#btnDataSource').addEventListener('click', () => {
        $('#dataSourceModal').classList.remove('hidden');
        loadSourceStatus();
    });
    $('#closeModal').addEventListener('click', () => $('#dataSourceModal').classList.add('hidden'));
    $('#dataSourceModal').addEventListener('click', (e) => {
        if (e.target === $('#dataSourceModal')) $('#dataSourceModal').classList.add('hidden');
    });

    // Select All / Deselect All
    $('#selectAll').addEventListener('click', () => {
        $('#projectCheckboxes').querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
        applyFilters();
    });
    $('#deselectAll').addEventListener('click', () => {
        $('#projectCheckboxes').querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
        applyFilters();
    });

    // Keyboard
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && presentMode) togglePresentMode(false);
    });
}

// ══════════════════════════════════════════════════════════════════
//  PRESENT MODE
// ══════════════════════════════════════════════════════════════════

function togglePresentMode(on) {
    presentMode = on;
    document.body.classList.toggle('present-mode', on);
    $('#presentBar').classList.toggle('hidden', !on);
    if (on) {
        document.documentElement.requestFullscreen?.().catch(() => { });
    } else {
        if (document.fullscreenElement) document.exitFullscreen?.();
    }
}

// ══════════════════════════════════════════════════════════════════
//  DATA SOURCE MODAL
// ══════════════════════════════════════════════════════════════════

async function loadSourceStatus() {
    try {
        const res = await fetch('/api/source-status');
        const data = await res.json();
        const el = $('#sourceStatus');
        if (data.connected) {
            el.className = 'source-status success';
            el.textContent = `✓ Connected to ${data.type}: ${data.url}`;
            el.classList.remove('hidden');
        } else {
            el.classList.add('hidden');
        }
    } catch { /* ignore */ }
}

window.connectSource = async function (type) {
    let url = '';
    if (type === 'google') url = $('#googleSheetUrl').value.trim();
    if (type === 'sharepoint') url = $('#sharepointUrl').value.trim();

    if (!url) {
        showSourceStatus('Please enter a valid URL.', 'error');
        return;
    }

    try {
        const res = await fetch('/api/connect-source', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ type, url }),
        });
        const data = await res.json();
        if (data.success) {
            showSourceStatus('✓ ' + data.message, 'success');
            setTimeout(() => { fetchAll(); $('#dataSourceModal').classList.add('hidden'); }, 1200);
        } else {
            showSourceStatus('✗ ' + (data.error || 'Connection failed'), 'error');
        }
    } catch (err) {
        showSourceStatus('✗ Network error: ' + err.message, 'error');
    }
}

function showSourceStatus(msg, type) {
    const el = $('#sourceStatus');
    el.className = 'source-status ' + type;
    el.textContent = msg;
    el.classList.remove('hidden');
}

// ══════════════════════════════════════════════════════════════════
//  PROJECT FILTER
// ══════════════════════════════════════════════════════════════════

function populateProjectFilter() {
    const projects = [...new Set(bandwidthData.map(d => d.project).filter(Boolean))].sort();
    const container = $('#projectCheckboxes');

    const prevChecked = new Set(
        [...container.querySelectorAll('input[type="checkbox"]:checked')].map(i => i.value)
    );
    const isFirst = container.children.length === 0;

    container.innerHTML = projects.map(p => {
        const checked = isFirst || prevChecked.has(p) ? 'checked' : '';
        return `
      <div class="project-row">
        <label><input type="checkbox" value="${esc(p)}" ${checked} />${esc(p)}</label>
        <button class="only-btn" data-project="${esc(p)}">ONLY</button>
      </div>`;
    }).join('');

    $('#projectCount').textContent = `(${projects.length})`;

    container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
        cb.addEventListener('change', applyFilters);
    });
    container.querySelectorAll('.only-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const proj = e.target.dataset.project;
            container.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = (cb.value === proj));
            applyFilters();
        });
    });

    // Auto-detect date range
    if (bandwidthData.length > 0) {
        const dates = [...new Set(bandwidthData.map(d => d.date).filter(Boolean))].sort();
        if (dates.length > 0 && !$('#filterDateFrom').value && !$('#filterDateTo').value) {
            $('#dateRangeLabel').textContent = fmtDisplay(dates[0]) + ' – ' + fmtDisplay(dates[dates.length - 1]);
        }
    }
}

// ══════════════════════════════════════════════════════════════════
//  FILTER + RENDER
// ══════════════════════════════════════════════════════════════════

function getCheckedProjects() {
    return new Set(
        [...$('#projectCheckboxes').querySelectorAll('input[type="checkbox"]:checked')].map(i => i.value)
    );
}

function applyFilters() {
    const from = $('#filterDateFrom').value;
    const to = $('#filterDateTo').value;
    const search = ($('#globalSearch').value || '').toLowerCase();
    const projects = getCheckedProjects();

    const filtered = bandwidthData.filter(d => {
        if (projects.size && !projects.has(d.project)) return false;
        if (from && d.date < from) return false;
        if (to && d.date > to) return false;
        if (search && !matchSearch(d, search)) return false;
        return true;
    });

    const filteredQBR = qbrData.filter(d => {
        if (projects.size && !projects.has(d.project)) return false;
        if (from && d.date < from) return false;
        if (to && d.date > to) return false;
        return true;
    });

    renderTeamTable(filtered);
    renderYesterdayToday(filtered);
    renderSidebarInfo(filtered, filteredQBR);
    renderBandwidthOverview(filtered);
}

function matchSearch(row, term) {
    return Object.values(row).some(v => v != null && v.toString().toLowerCase().includes(term));
}

// ══════════════════════════════════════════════════════════════════
//  TEAM TABLE
// ══════════════════════════════════════════════════════════════════

function renderTeamTable(data) {
    const body = $('#teamTableBody');
    const seen = new Map();
    data.forEach(d => { if (d.name && !seen.has(d.name)) seen.set(d.name, d); });

    if (seen.size === 0) {
        body.innerHTML = '<div class="empty-state">No data for current filters</div>';
        return;
    }

    let members = [...seen.values()];
    members.sort((a, b) => {
        const cmp = (a.name || '').localeCompare(b.name || '');
        return sortAsc ? cmp : -cmp;
    });

    body.innerHTML = members.map(d => `
    <div class="table-row">
      <span>${esc(d.name)}</span>
      <span>${esc(d.role)}</span>
      <span>${esc(d.freeBandwidth) || '—'}</span>
      <span>${leaveBadge(d.leaveStatus)}</span>
    </div>
  `).join('');
}

function leaveBadge(status) {
    if (!status) return '<span class="badge badge-avail">Available</span>';
    const s = status.toLowerCase();
    if (s.includes('full')) return `<span class="badge badge-full">${esc(status)}</span>`;
    if (s.includes('half')) return `<span class="badge badge-half">${esc(status)}</span>`;
    return `<span class="badge badge-avail">${esc(status)}</span>`;
}

// ══════════════════════════════════════════════════════════════════
//  YESTERDAY / TODAY
// ══════════════════════════════════════════════════════════════════

function renderYesterdayToday(data) {
    const dates = [...new Set(data.map(d => d.date).filter(Boolean))].sort().reverse();
    renderPanel('yesterdayList', data.filter(d => d.date === (dates[1] || '')));
    renderPanel('todayList', data.filter(d => d.date === (dates[0] || '')));
}

function renderPanel(id, items) {
    const el = document.getElementById(id);

    // 1. Filter: only deliverables (exclude calls, operations)
    const NON_DELIVERABLE = /\b(call|calls|meeting|meetings|operation|operations|operational|standup|sync)\b/i;
    const deliverables = items.filter(d => {
        const wi = (d.workItem || '').toLowerCase();
        const desc = (d.description || '').toLowerCase();
        // Exclude if workItem is purely a call/operation
        if (wi && NON_DELIVERABLE.test(wi)) return false;
        return true;
    });

    if (!deliverables.length) {
        el.innerHTML = '<div class="panel-item" style="color:#999">No entries</div>';
        return;
    }

    // 2. Group by person + project → merge descriptions with comma
    const grouped = new Map();
    deliverables.forEach(d => {
        const key = `${d.name}|||${d.project}`;
        if (!grouped.has(key)) {
            grouped.set(key, { name: d.name, project: d.project, descs: [] });
        }
        const desc = (d.description || d.workItem || '').trim();
        if (desc) grouped.get(key).descs.push(desc);
    });

    // 3. Render one line per person-project, all descriptions comma-joined
    el.innerHTML = [...grouped.values()].map(g => {
        const merged = g.descs.join(', ') || 'N/A';
        return `<div class="panel-item"><strong>${esc(g.name)}:</strong> ${esc(merged)}</div>`;
    }).join('');
}

// ══════════════════════════════════════════════════════════════════
//  LEFT SIDEBAR INFO
// ══════════════════════════════════════════════════════════════════

function renderSidebarInfo(bwData, qbrFiltered) {
    const projects = [...new Set(bwData.map(d => d.project).filter(Boolean))];
    let detailText = '—', pmText = '—';

    if (projects.length === 1) {
        const proj = projects[0];
        const ddMatch = dropdownData.find(d => d.projectName === proj);
        if (ddMatch && ddMatch.projectDetails) detailText = ddMatch.projectDetails;
        const bwMatch = bwData.find(d => d.project === proj && d.projectManager);
        if (bwMatch) pmText = bwMatch.projectManager;
        if (pmText === '—' && ddMatch && ddMatch.projectManager) pmText = ddMatch.projectManager;
    } else if (projects.length > 1) {
        detailText = `${projects.length} projects selected`;
        const pms = [...new Set(bwData.map(d => d.projectManager).filter(Boolean))];
        pmText = pms.length === 1 ? pms[0] : `${pms.length} managers`;
    }

    $('#projectDetailsDisplay').textContent = detailText;
    $('#pmDisplay').textContent = pmText;

    const blockerEntries = qbrFiltered.filter(d => d.blockers).map(d => d.blockers);
    $('#blockersDisplay').textContent = blockerEntries.length ? blockerEntries.join(', ') : 'No data';

    const qbrDates = [...new Set(qbrFiltered.map(d => d.mbrQbrDate).filter(Boolean))];
    $('#qbrDateDisplay').textContent = qbrDates.length ? fmtDisplay(qbrDates.sort().reverse()[0]) : '—';
}

// ══════════════════════════════════════════════════════════════════
//  FREE BANDWIDTH OVERVIEW
// ══════════════════════════════════════════════════════════════════

function renderBandwidthOverview(data) {
    const body = $('#bwOverviewBody');

    // Build per-person overview: aggregate across all their entries
    const personMap = new Map();

    data.forEach(d => {
        if (!d.name) return;
        if (!personMap.has(d.name)) {
            personMap.set(d.name, {
                name: d.name,
                role: d.role || '',
                entries: [],
                leaveStatuses: [],
                freeBandwidths: [],
            });
        }
        const p = personMap.get(d.name);
        p.entries.push(d);
        if (d.leaveStatus) p.leaveStatuses.push(d.leaveStatus);
        if (d.freeBandwidth) p.freeBandwidths.push(d.freeBandwidth);
    });

    if (personMap.size === 0) {
        body.innerHTML = '<div class="empty-state">No data for current filters</div>';
        return;
    }

    const persons = [...personMap.values()].sort((a, b) => a.name.localeCompare(b.name));

    // Get total dates in range for percentage calculation
    const allDates = [...new Set(data.map(d => d.date).filter(Boolean))];
    const totalDays = Math.max(allDates.length, 1);

    body.innerHTML = persons.map(p => {
        // Calculate bandwidth fill percentage
        const totalEntries = p.entries.length;
        const fillPercent = Math.min(Math.round((totalEntries / totalDays) * 100), 100);

        // Determine status: full day leave, half day, partial, or available
        const latestLeave = p.leaveStatuses.length > 0 ? p.leaveStatuses[p.leaveStatuses.length - 1] : '';
        const statusInfo = getStatusInfo(latestLeave, fillPercent);

        return `
      <div class="bw-person-row">
        <div class="bw-person-name" title="${esc(p.name)}">${esc(p.name)}</div>
        <div class="bw-bar-wrap">
          <div class="bw-bar-bg">
            <div class="bw-bar-fill ${statusInfo.fillClass}" style="width:${fillPercent}%">
              ${fillPercent > 15 ? `<span class="bw-bar-label">${fillPercent}%</span>` : ''}
            </div>
          </div>
        </div>
        <div class="bw-status-text ${statusInfo.textClass}">${statusInfo.label}</div>
      </div>`;
    }).join('');

    // Update subtitle
    const sub = $('#bwOverviewSubtitle');
    if (sub) {
        const onLeave = persons.filter(p => {
            const ls = p.leaveStatuses.length > 0 ? p.leaveStatuses[p.leaveStatuses.length - 1] : '';
            return ls.toLowerCase().includes('full');
        }).length;
        const partial = persons.filter(p => {
            const ls = p.leaveStatuses.length > 0 ? p.leaveStatuses[p.leaveStatuses.length - 1] : '';
            return ls.toLowerCase().includes('half');
        }).length;
        const avail = persons.length - onLeave - partial;
        sub.textContent = `${persons.length} members  •  ${avail} available  •  ${onLeave} on leave  •  ${partial} half day`;
    }
}

function getStatusInfo(leaveStatus, fillPercent) {
    const ls = (leaveStatus || '').toLowerCase();

    if (ls.includes('full')) {
        return { fillClass: 'fill-full', textClass: 'status-full', label: 'On Leave' };
    }
    if (ls.includes('half')) {
        return { fillClass: 'fill-half', textClass: 'status-half', label: 'Half Day' };
    }
    if (fillPercent < 60) {
        return { fillClass: 'fill-partial', textClass: 'status-partial', label: 'Partial' };
    }
    return { fillClass: 'fill-avail', textClass: 'status-avail', label: 'Available' };
}

// ══════════════════════════════════════════════════════════════════
//  UTILITIES
// ══════════════════════════════════════════════════════════════════

function esc(str) {
    if (str == null) return '';
    const d = document.createElement('div');
    d.textContent = str.toString();
    return d.innerHTML;
}

function fmtDisplay(dateStr) {
    if (!dateStr) return '';
    try {
        const d = new Date(dateStr + 'T00:00:00');
        if (isNaN(d.getTime())) return dateStr;
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        return `${months[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
    } catch { return dateStr; }
}

/* ═══════════════════════════════════════════════════════════════
   PHASE 2: EXECUTIVE DASHBOARD — INTEGRATED
   ═══════════════════════════════════════════════════════════════ */

const EXEC_TOKEN_KEY = 'exec_token';
let execChartFTE = null, execChartKPI = null, execChartSOW = null;

// ── Init exec UI bindings ───────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    // Check if already logged in
    const savedToken = sessionStorage.getItem(EXEC_TOKEN_KEY);
    if (savedToken) {
        fetch('/api/auth/verify', { headers: { Authorization: 'Bearer ' + savedToken } })
            .then(r => r.json())
            .then(d => { if (d.valid) showExecDashboard(); })
            .catch(() => { });
    }

    // Login modal
    const btnLogin = document.getElementById('btnManagerLogin');
    const btnLogout = document.getElementById('btnExecLogout');
    const modal = document.getElementById('loginModal');
    const btnSubmit = document.getElementById('loginSubmit');
    const btnClose = document.getElementById('loginClose');

    if (btnLogin) btnLogin.addEventListener('click', () => { modal.style.display = 'flex'; document.getElementById('loginUser').focus(); });
    if (btnClose) btnClose.addEventListener('click', () => { modal.style.display = 'none'; });
    if (modal) modal.addEventListener('click', e => { if (e.target === modal) modal.style.display = 'none'; });
    if (btnSubmit) btnSubmit.addEventListener('click', execDoLogin);
    const passInput = document.getElementById('loginPass');
    if (passInput) passInput.addEventListener('keydown', e => { if (e.key === 'Enter') execDoLogin(); });
    if (btnLogout) btnLogout.addEventListener('click', execLogout);

    // Exec buttons
    const btnRefresh = document.getElementById('btnExecRefresh');
    if (btnRefresh) btnRefresh.addEventListener('click', fetchExecData);
    const btnPrint = document.getElementById('btnExecPrint');
    if (btnPrint) btnPrint.addEventListener('click', () => window.print());
    const btnSources = document.getElementById('btnExecSources');
    if (btnSources) btnSources.addEventListener('click', () => { document.getElementById('execSourceModal').style.display = 'flex'; });
    const btnCloseSrc = document.getElementById('closeExecSrcModal');
    if (btnCloseSrc) btnCloseSrc.addEventListener('click', () => { document.getElementById('execSourceModal').style.display = 'none'; });
    const srcModal = document.getElementById('execSourceModal');
    if (srcModal) srcModal.addEventListener('click', e => { if (e.target === srcModal) srcModal.style.display = 'none'; });

    // Tabs
    document.querySelectorAll('.exec-tab').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.exec-tab').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.exec-tab-content').forEach(c => c.classList.remove('active'));
            btn.classList.add('active');
            const tabId = btn.dataset.tab === 'upload' ? 'tabUpload' : 'tabUrls';
            document.getElementById(tabId).classList.add('active');
        });
    });

    // Upload + Connect
    const btnUpload = document.getElementById('btnUploadAll');
    if (btnUpload) btnUpload.addEventListener('click', execUploadFiles);
    const btnConnect = document.getElementById('btnConnectUrls');
    if (btnConnect) btnConnect.addEventListener('click', execConnectUrls);
});

// ── Login ────────────────────────────────────────────────────────
async function execDoLogin() {
    const username = document.getElementById('loginUser').value.trim();
    const password = document.getElementById('loginPass').value;
    const errEl = document.getElementById('loginError');
    errEl.style.display = 'none';

    if (!username || !password) { errEl.textContent = 'Enter both User ID and Password'; errEl.style.display = 'block'; return; }

    try {
        const res = await fetch('/api/auth/login', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ username, password }),
        });
        const data = await res.json();
        if (res.ok && data.token) {
            sessionStorage.setItem(EXEC_TOKEN_KEY, data.token);
            document.getElementById('loginModal').style.display = 'none';
            showExecDashboard();
        } else {
            errEl.textContent = data.error || 'Login failed';
            errEl.style.display = 'block';
        }
    } catch (e) {
        errEl.textContent = 'Network error';
        errEl.style.display = 'block';
    }
}

function execLogout() {
    sessionStorage.removeItem(EXEC_TOKEN_KEY);
    document.getElementById('execSection').style.display = 'none';
    document.getElementById('btnManagerLogin').style.display = '';
    document.getElementById('btnExecLogout').style.display = 'none';
    // Show Phase 1 back
    document.getElementById('dashboard').style.display = '';
    document.getElementById('main-footer').style.display = '';
    document.querySelectorAll('.phase1-only').forEach(el => el.style.display = '');
    document.querySelector('#main-header h1').textContent = 'Digital Enablement Team - Daily Standup Updates';
    if (window.ManagerDash) window.ManagerDash.close();
}

function showExecDashboard() {
    // Hide Phase 1 completely
    document.getElementById('dashboard').style.display = 'none';
    document.getElementById('main-footer').style.display = 'none';
    document.querySelectorAll('.phase1-only').forEach(el => el.style.display = 'none');
    // Show the merged manager dashboard
    document.getElementById('execSection').style.display = 'block';
    document.getElementById('btnManagerLogin').style.display = 'none';
    document.getElementById('btnExecLogout').style.display = '';
    document.querySelector('#main-header h1').textContent = 'Executive Overview';
    if (window.ManagerDash) window.ManagerDash.open();
}

// ── Fetch all executive data ─────────────────────────────────────
async function fetchExecData() {
    const token = sessionStorage.getItem(EXEC_TOKEN_KEY);
    if (!token) return;
    const h = { Authorization: 'Bearer ' + token };
    try {
        const [summary, kpi, sow, gov, leave, ftr, health, team] = await Promise.all([
            fetch('/api/exec/summary', { headers: h }).then(r => r.json()),
            fetch('/api/exec/kpi-scorecards', { headers: h }).then(r => r.json()),
            fetch('/api/exec/sow-financial', { headers: h }).then(r => r.json()),
            fetch('/api/exec/governance-risks', { headers: h }).then(r => r.json()),
            fetch('/api/exec/leave-impact', { headers: h }).then(r => r.json()),
            fetch('/api/exec/ftr-metrics', { headers: h }).then(r => r.json()),
            fetch('/api/exec/project-health', { headers: h }).then(r => r.json()),
            fetch('/api/exec/team-capacity', { headers: h }).then(r => r.json()),
        ]);
        renderExecKPIs(summary, gov);
        renderExecFTEChart(gov);
        renderExecKPIChart(kpi);
        renderExecSOWChart(sow);
        renderExecRolesChart(team);
        renderExecQAChart(ftr);
        renderExecLeaveChart(leave);
        renderExecRiskPanel(gov);
        renderExecLeavePanel(leave);
        renderExecHealthTable(health);
        renderExecHL(gov);
        if(typeof cacheExecData==='function'){cacheExecData('team',team);cacheExecData('kpi',kpi);cacheExecData('sow',sow);cacheExecData('gov',gov);cacheExecData('leave',leave);cacheExecData('ftr',ftr);cacheExecData('health',health);}
        const ts = document.getElementById('execLastRefresh');
        if (ts) ts.textContent = new Date().toLocaleTimeString();
        // Also fetch all-sheets for Excel Data Browser
        try {
            const allSheets = await fetch('/api/exec/all-sheets', { headers: h }).then(r => r.json());
            renderExcelBrowser(allSheets);
        } catch (e) { console.warn('Excel browser fetch error:', e); }
    } catch (err) {
        console.error('Exec fetch error:', err);
    }
}

// ── KPI Cards ────────────────────────────────────────────────────
function renderExecKPIs(d, gov) {
    if (!d) return;
    const el = id => document.getElementById(id);
    el('valTeamSize').textContent = d.teamSize || 0;
    el('valActiveProjects').textContent = d.activeProjects || 0;
    el('valSOWValue').textContent = d.totalSOWValue ? fmtCurrency(d.totalSOWValue) : '—';
    el('valKPIMetRate').textContent = d.kpiMetRate ? d.kpiMetRate + '%' : '—';
    el('valFTRRating').textContent = d.ftrAvgRating ? d.ftrAvgRating + '%' : '—';
    el('valOnLeave').textContent = d.onLeaveToday || 0;
    el('valActiveRisks').textContent = d.activeRisks || 0;
    const sowRcv = gov?.highlights?.filter(h => h.highlight)?.length || '—';
    el('valSOWReceived').textContent = sowRcv;
}
function fmtCurrency(v) { return v >= 1e6 ? '$' + (v / 1e6).toFixed(1) + 'M' : v >= 1e3 ? '$' + (v / 1e3).toFixed(0) + 'K' : '$' + v; }

// ── Charts ──────────────────────────────────────────────────────
function renderExecFTEChart(gov) {
    const ctx = document.getElementById('chartFTETrend');
    if (!ctx) return;
    const trend = gov?.fteTrend || [];
    if (execChartFTE) execChartFTE.destroy();
    execChartFTE = new Chart(ctx, { type: 'line', data: { labels: trend.map(t => t.month), datasets: [{ label: 'Total FTE', data: trend.map(t => t.totalFTE), borderColor: '#0891b2', backgroundColor: 'rgba(8,145,178,.1)', fill: true, tension: .4, pointRadius: 4, pointBackgroundColor: '#0891b2' }] }, options: { responsive: true, plugins: { legend: { labels: { font: { size: 11 } } } }, scales: { y: { beginAtZero: true } } } });
}

function renderExecKPIChart(kpiData) {
    const ctx = document.getElementById('chartKPI');
    if (!ctx) return;
    const kpis = (kpiData?.kpis || []).slice(0, 10);
    if (execChartKPI) execChartKPI.destroy();
    execChartKPI = new Chart(ctx, { type: 'bar', data: { labels: kpis.map(k => k.metricName?.slice(0, 22) || ''), datasets: [{ label: 'Target %', data: kpis.map(k => k.target * 100), backgroundColor: 'rgba(168,85,247,.6)', borderRadius: 4 }, { label: 'Actual %', data: kpis.map(k => k.actual * 100), backgroundColor: 'rgba(8,145,178,.7)', borderRadius: 4 }] }, options: { indexAxis: 'y', responsive: true, scales: { x: { max: 120 } } } });
}

function renderExecSOWChart(sowData) {
    const ctx = document.getElementById('chartSOW');
    if (!ctx) return;
    const sm = sowData?.summary?.projectsByStatus || {};
    const labels = Object.keys(sm), values = Object.values(sm);
    const colors = ['#22c55e', '#f59e0b', '#0891b2', '#a855f7', '#ef4444', '#ec4899'];
    if (execChartSOW) execChartSOW.destroy();
    execChartSOW = new Chart(ctx, { type: 'doughnut', data: { labels, datasets: [{ data: values, backgroundColor: colors.slice(0, labels.length), borderWidth: 0 }] }, options: { responsive: true, plugins: { legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 10 } } } } });
}

// ── Risk Panel ──────────────────────────────────────────────────
function renderExecRiskPanel(gov) {
    const risks = (gov?.risks || []).filter(r => r.status?.toLowerCase() === 'ongoing');
    const body = document.getElementById('riskPanelBody');
    const badge = document.getElementById('riskCount');
    if (badge) badge.textContent = risks.length;
    if (!risks.length) { body.innerHTML = '<div style="color:#999">No active risks ✅</div>'; return; }
    body.innerHTML = risks.map(r => {
        const ic = r.impact?.toLowerCase().includes('loss') ? 'exec-impact-high' : r.impact?.toLowerCase().includes('delay') ? 'exec-impact-med' : 'exec-impact-low';
        return `<div class="exec-risk-item"><div class="exec-risk-project">${esc(r.project)} <span style="color:#999;font-weight:400">— ${esc(r.pm)}</span></div><div class="exec-risk-text">${esc(r.risk)}</div><span class="exec-risk-impact ${ic}">${esc(r.impact || 'Unknown')}</span></div>`;
    }).join('');
}

// ── Leave Panel ─────────────────────────────────────────────────
function renderExecLeavePanel(leave) {
    const body = document.getElementById('leavePanelBody');
    const cm = leave?.currentMonth || {};
    const byP = cm.byPerson || {};
    const onT = cm.onLeaveToday || [];
    let html = '';
    if (onT.length) {
        html += `<div style="margin-bottom:8px;color:#f59e0b;font-weight:600">📌 On Leave Today (${onT.length}):</div>`;
        html += onT.map(p => `<div class="exec-leave-item"><span class="exec-leave-name">${esc(p.name)}</span><span style="color:#999">${esc(p.type)}</span></div>`).join('');
        html += '<hr style="border-color:#eee;margin:10px 0">';
    }
    html += `<div style="margin-bottom:6px;font-weight:600;color:#333">Month Total: ${cm.totalLeaves || 0} leaves</div>`;
    const sorted = Object.entries(byP).filter(([, c]) => c > 0).sort((a, b) => b[1] - a[1]).slice(0, 15);
    if (!sorted.length) html += '<div style="color:#999">No leave data.</div>';
    else html += sorted.map(([n, c]) => `<div class="exec-leave-item"><span class="exec-leave-name">${esc(n)}</span><span class="exec-leave-count">${c} day(s)</span></div>`).join('');
    body.innerHTML = html;
}

// ── Health Table ────────────────────────────────────────────────
function renderExecHealthTable(health) {
    const tbody = document.getElementById('healthTableBody');
    if (!health?.length) { tbody.innerHTML = '<tr><td colspan="8" style="color:#999;text-align:center;padding:16px">No project data</td></tr>'; return; }
    tbody.innerHTML = health.map(p => {
        const cls = p.health === 'Green' ? 'exec-hb-green' : p.health === 'Red' ? 'exec-hb-red' : 'exec-hb-amber';
        return `<tr><td><strong>${esc(p.projectName)}</strong></td><td>${esc(p.client)}</td><td>${esc(p.pm)}</td><td>${esc(p.status)}</td><td>${esc(p.sowStatus)}</td><td>${esc(p.poStatus)}</td><td><span class="exec-hb ${cls}">${p.health}</span></td><td style="font-size:.72rem;color:#999">${esc(p.reasons)}</td></tr>`;
    }).join('');
}

// ── Highlights / Lowlights ──────────────────────────────────────
function renderExecHL(gov) {
    const hl = gov?.highlights || [];
    const hlEl = document.getElementById('highlightsList');
    const hItems = hl.filter(h => h.highlight);
    hlEl.innerHTML = hItems.length ? hItems.slice(0, 10).map(h => `<div class="exec-hl-item"><div class="exec-hl-proj">${esc(h.project)} — ${esc(h.month)}</div><div class="exec-hl-text">${esc(h.highlight)}</div></div>`).join('') : '<div style="color:#999">No highlights.</div>';
    const llEl = document.getElementById('lowlightsList');
    const lItems = hl.filter(h => h.lowlight);
    llEl.innerHTML = lItems.length ? lItems.slice(0, 10).map(h => `<div class="exec-hl-item"><div class="exec-hl-proj">${esc(h.project)} — ${esc(h.month)}</div><div class="exec-hl-text">${esc(h.lowlight)}</div></div>`).join('') : '<div style="color:#999">No lowlights.</div>';
}

// ── Data Source Upload ──────────────────────────────────────────
async function execUploadFiles() {
    const form = document.getElementById('uploadForm');
    const fd = new FormData(form);
    const st = document.getElementById('execSrcStatus');
    let hasFile = false;
    for (const [, f] of fd.entries()) { if (f && f.size > 0) { hasFile = true; break; } }
    if (!hasFile) { st.className = 'exec-src-status error'; st.textContent = 'Select at least one file.'; st.style.display = 'block'; return; }
    st.className = 'exec-src-status'; st.textContent = 'Uploading…'; st.style.display = 'block';
    try {
        const res = await fetch('/api/exec/upload-sources', { method: 'POST', headers: { Authorization: 'Bearer ' + sessionStorage.getItem(EXEC_TOKEN_KEY) }, body: fd });
        const data = await res.json();
        if (data.success) { st.className = 'exec-src-status success'; st.textContent = '✓ ' + data.message; setTimeout(() => { fetchExecData(); document.getElementById('execSourceModal').style.display = 'none'; }, 1200); }
        else { st.className = 'exec-src-status error'; st.textContent = '✗ ' + (data.error || 'Failed'); }
    } catch (e) { st.className = 'exec-src-status error'; st.textContent = '✗ ' + e.message; }
}

async function execConnectUrls() {
    const sources = {};
    [['ftr', 'urlFtr'], ['team', 'urlTeam'], ['sow', 'urlSow'], ['governance', 'urlGovernance'], ['leave', 'urlLeave'], ['kpi', 'urlKpi']].forEach(([k, id]) => {
        const v = document.getElementById(id).value.trim();
        if (v) sources[k] = { url: v, type: 'sharepoint' };
    });
    const st = document.getElementById('execSrcStatus');
    if (!Object.keys(sources).length) { st.className = 'exec-src-status error'; st.textContent = 'Enter at least one URL.'; st.style.display = 'block'; return; }
    st.className = 'exec-src-status'; st.textContent = 'Connecting…'; st.style.display = 'block';
    try {
        const res = await fetch('/api/exec/connect-sources', { method: 'POST', headers: { Authorization: 'Bearer ' + sessionStorage.getItem(EXEC_TOKEN_KEY), 'Content-Type': 'application/json' }, body: JSON.stringify({ sources }) });
        const data = await res.json();
        if (data.success) { st.className = 'exec-src-status success'; st.textContent = '✓ ' + data.message; setTimeout(() => { fetchExecData(); document.getElementById('execSourceModal').style.display = 'none'; }, 1200); }
        else { st.className = 'exec-src-status error'; st.textContent = '✗ ' + (data.error || 'Failed'); }
    } catch (e) { st.className = 'exec-src-status error'; st.textContent = '✗ ' + e.message; }
}

// ── NEW: Role Distribution Chart (Team Details) ─────────────────
let execChartRoles = null;
function renderExecRolesChart(teamData) {
    const ctx = document.getElementById('chartRoles');
    if (!ctx) return;
    const roles = (teamData?.team?.roleDistribution || teamData?.roleDistribution) || {};
    const labels = Object.keys(roles), values = Object.values(roles);
    const colors = ['#06b6d4', '#a855f7', '#22c55e', '#f59e0b', '#ec4899', '#3b82f6', '#ef4444', '#14b8a6'];
    if (execChartRoles) execChartRoles.destroy();
    execChartRoles = new Chart(ctx, { type: 'pie', data: { labels, datasets: [{ data: values, backgroundColor: colors.slice(0, labels.length), borderWidth: 2, borderColor: '#fff' }] }, options: { responsive: true, plugins: { legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 8 } } } } });
}

// ── NEW: QA Rating by Project (FTR Tracker) ─────────────────────
let execChartQA = null;
function renderExecQAChart(ftrData) {
    const ctx = document.getElementById('chartQA');
    if (!ctx) return;
    const accounts = ftrData?.accounts || [];
    const labels = accounts.slice(0, 8).map(a => (a.account || '').slice(0, 18));
    const values = accounts.slice(0, 8).map(a => a.avgRating || 0);
    if (execChartQA) execChartQA.destroy();
    execChartQA = new Chart(ctx, { type: 'bar', data: { labels, datasets: [{ label: 'Avg QA %', data: values, backgroundColor: values.map(v => v >= 90 ? '#22c55e' : v >= 75 ? '#f59e0b' : '#ef4444'), borderRadius: 6 }] }, options: { responsive: true, indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { max: 100 } } } });
}

// ── NEW: Top Leave Takers Chart (Leave Tracker) ─────────────────
let execChartLeave = null;
function renderExecLeaveChart(leaveData) {
    const ctx = document.getElementById('chartLeave');
    if (!ctx) return;
    const byP = leaveData?.currentMonth?.byPerson || {};
    const sorted = Object.entries(byP).filter(([, c]) => c > 0).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const labels = sorted.map(([n]) => n.split(' ').slice(0, 2).join(' '));
    const values = sorted.map(([, c]) => c);
    if (execChartLeave) execChartLeave.destroy();
    execChartLeave = new Chart(ctx, { type: 'bar', data: { labels, datasets: [{ label: 'Leave Days', data: values, backgroundColor: '#f59e0b', borderRadius: 6 }] }, options: { responsive: true, plugins: { legend: { display: false } } } });
}

// ── Upload Card File Labels ─────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('.exec-upload-card input[type="file"]').forEach(inp => {
        inp.addEventListener('change', () => {
            const lbl = inp.closest('.exec-upload-card').querySelector('.exec-uc-file');
            const card = inp.closest('.exec-upload-card');
            if (inp.files.length) { lbl.textContent = inp.files[0].name; card.classList.add('has-file'); }
            else { lbl.textContent = 'No file'; card.classList.remove('has-file'); }
        });
    });
});

//  Tracker Links Edit Logic 
(function initTrackers() {
    const STORAGE_KEY = 'exec_tracker_links';
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}'); } catch {}

    function loadTrackers() {
        document.querySelectorAll('.exec-tracker-card').forEach(card => {
            const key = card.dataset.key;
            const info = saved[key];
            if (info && info.url) {
                card.href = info.url;
                card.querySelector('.exec-tc-name').textContent = info.name || key;
                try { card.querySelector('.exec-tc-url').textContent = new URL(info.url).hostname; } catch { card.querySelector('.exec-tc-url').textContent = info.url.slice(0,30); }
                card.classList.add('active');
            }
        });
    }

    let editKey = null;
    const modal = document.getElementById('trackerEditModal');
    const nameInp = document.getElementById('trackerEditName');
    const urlInp = document.getElementById('trackerEditUrl');
    const titleEl = document.getElementById('trackerEditTitle');

    document.querySelectorAll('.exec-tracker-card').forEach(card => {
        card.addEventListener('click', e => {
            const info = saved[card.dataset.key];
            if (!info || !info.url) { e.preventDefault(); openEditModal(card.dataset.key); }
        });
        card.addEventListener('contextmenu', e => { e.preventDefault(); openEditModal(card.dataset.key); });
    });

    const addBtn = document.getElementById('btnAddTracker');
    if (addBtn) addBtn.addEventListener('click', () => {
        const firstEmpty = document.querySelector('.exec-tracker-card:not(.active)');
        openEditModal(firstEmpty ? firstEmpty.dataset.key : 'ftr');
    });

    function openEditModal(key) {
        editKey = key;
        const card = document.querySelector('.exec-tracker-card[data-key="' + key + '"]');
        if (titleEl) titleEl.textContent = 'Edit: ' + (card ? card.querySelector('.exec-tc-name').textContent : key);
        if (nameInp) nameInp.value = (saved[key] && saved[key].name) || (card ? card.querySelector('.exec-tc-name').textContent : '');
        if (urlInp) urlInp.value = (saved[key] && saved[key].url) || '';
        if (modal) { modal.style.display = 'flex'; nameInp.focus(); }
    }

    const saveBtn = document.getElementById('trackerEditSave');
    if (saveBtn) saveBtn.addEventListener('click', () => {
        if (!editKey) return;
        const name = nameInp.value.trim();
        const url = urlInp.value.trim();
        if (!url) { urlInp.style.border = '2px solid #ef4444'; return; }
        saved[editKey] = { name: name || editKey, url: url };
        localStorage.setItem(STORAGE_KEY, JSON.stringify(saved));
        loadTrackers();
        modal.style.display = 'none';
    });

    const cancelBtn = document.getElementById('trackerEditCancel');
    if (cancelBtn) cancelBtn.addEventListener('click', () => modal.style.display = 'none');
    if (modal) modal.addEventListener('click', e => { if (e.target === modal) modal.style.display = 'none'; });

    document.addEventListener('DOMContentLoaded', loadTrackers);
})();

//  KPI Card Click  Detail Panel 
var _execCache = {};
let _kpiDetailChart = null;

// Store fetched data globally for KPI clicks
function cacheExecData(key, data) { _execCache[key] = data; }

document.querySelectorAll('.exec-kpi[data-kpi]').forEach(card => {
    card.addEventListener('click', () => {
        const key = card.dataset.kpi;
        const panel = document.getElementById('kpiDetailPanel');
        const isOpen = panel.style.display === 'block' && card.classList.contains('kpi-active');
        document.querySelectorAll('.exec-kpi').forEach(c => c.classList.remove('kpi-active'));
        if (isOpen) { panel.style.display = 'none'; return; }
        card.classList.add('kpi-active');
        panel.style.display = 'block';
        renderKPIDetail(key);
    });
});

document.getElementById('kpiDetailClose')?.addEventListener('click', () => {
    document.getElementById('kpiDetailPanel').style.display = 'none';
    document.querySelectorAll('.exec-kpi').forEach(c => c.classList.remove('kpi-active'));
});

function renderKPIDetail(key) {
    const title = document.getElementById('kpiDetailTitle');
    const tableDiv = document.getElementById('kpiDetailTable');
    const ctx = document.getElementById('kpiDetailChart');
    if (_kpiDetailChart) { _kpiDetailChart.destroy(); _kpiDetailChart = null; }

    const configs = {
        team: () => {
            title.textContent = ' Team Breakdown';
            const d = _execCache.team || {};
            const roles = d.roleDistribution || {};
            const labels = Object.keys(roles), vals = Object.values(roles);
            _kpiDetailChart = new Chart(ctx, { type:'doughnut', data:{ labels, datasets:[{ data:vals, backgroundColor:['#06b6d4','#a855f7','#22c55e','#f59e0b','#ec4899','#3b82f6','#ef4444','#14b8a6'] }] }, options:{ responsive:true, plugins:{ legend:{ position:'bottom' } } } });
            const members = (d?.team?.activeMembers || d?.activeMembers) || [];
            tableDiv.innerHTML = '<table><thead><tr><th>Name</th><th>Role</th><th>Location</th></tr></thead><tbody>' + members.slice(0,15).map(m => '<tr><td>'+m.name+'</td><td>'+m.role+'</td><td>'+(m.location||'-')+'</td></tr>').join('') + '</tbody></table>';
        },
        projects: () => {
            title.textContent = ' Active Projects';
            const d = _execCache.health || [];
            const statuses = {}; d.forEach(p => { statuses[p.health||'Unknown'] = (statuses[p.health||'Unknown']||0)+1; });
            _kpiDetailChart = new Chart(ctx, { type:'doughnut', data:{ labels:Object.keys(statuses), datasets:[{ data:Object.values(statuses), backgroundColor:['#22c55e','#f59e0b','#ef4444','#94a3b8'] }] }, options:{ responsive:true, plugins:{ legend:{ position:'bottom' } } } });
            tableDiv.innerHTML = '<table><thead><tr><th>Project</th><th>Client</th><th>PM</th><th>Health</th></tr></thead><tbody>' + d.map(p => '<tr><td>'+p.project+'</td><td>'+(p.client||'-')+'</td><td>'+(p.pm||'-')+'</td><td>'+(p.health||'-')+'</td></tr>').join('') + '</tbody></table>';
        },
        sow: () => {
            title.textContent = ' SOW & PO Financial Details';
            const d = _execCache.sow || {};
            const projects = d.projects || [];
            const labels = projects.slice(0,10).map(p => (p.project||'').slice(0,20));
            const vals = projects.slice(0,10).map(p => p.sowValue||0);
            _kpiDetailChart = new Chart(ctx, { type:'bar', data:{ labels, datasets:[{ label:'SOW Value', data:vals, backgroundColor:'#22c55e', borderRadius:6 }] }, options:{ responsive:true, plugins:{ legend:{ display:false } } } });
            tableDiv.innerHTML = '<table><thead><tr><th>Project</th><th>SOW</th><th>PO</th><th>Status</th></tr></thead><tbody>' + projects.map(p => '<tr><td>'+p.project+'</td><td>$'+(p.sowValue||0).toLocaleString()+'</td><td>'+(p.poNumber||'-')+'</td><td>'+(p.status||'-')+'</td></tr>').join('') + '</tbody></table>';
        },
        kpi: () => {
            title.textContent = ' KPI Scorecards � All Metrics';
            const d = _execCache.kpi || {};
            const metrics = d.metrics || d.kpis || [];
            const labels = metrics.slice(0,10).map(m => (m.metric||m.name||'').slice(0,20));
            const targets = metrics.slice(0,10).map(m => m.target||0);
            const actuals = metrics.slice(0,10).map(m => m.actual||0);
            _kpiDetailChart = new Chart(ctx, { type:'bar', data:{ labels, datasets:[{ label:'Target', data:targets, backgroundColor:'rgba(59,130,246,.5)', borderRadius:4 },{ label:'Actual', data:actuals, backgroundColor:'#06b6d4', borderRadius:4 }] }, options:{ responsive:true, plugins:{ legend:{ position:'bottom' } } } });
            tableDiv.innerHTML = '<table><thead><tr><th>Metric</th><th>Target</th><th>Actual</th><th>Status</th></tr></thead><tbody>' + metrics.map(m => '<tr><td>'+(m.metric||m.name)+'</td><td>'+(m.target||'-')+'</td><td>'+(m.actual||'-')+'</td><td>'+(m.status||'-')+'</td></tr>').join('') + '</tbody></table>';
        },
        ftr: () => {
            title.textContent = ' FTR / QA Metrics';
            const d = _execCache.ftr || {};
            const accounts = d.accounts || [];
            const labels = accounts.map(a => (a.account||'').slice(0,18));
            const vals = accounts.map(a => a.avgRating||0);
            _kpiDetailChart = new Chart(ctx, { type:'bar', data:{ labels, datasets:[{ label:'QA Rating %', data:vals, backgroundColor:vals.map(v=>v>=90?'#22c55e':v>=75?'#f59e0b':'#ef4444'), borderRadius:6 }] }, options:{ responsive:true, indexAxis:'y', plugins:{ legend:{ display:false } }, scales:{ x:{ max:100 } } } });
            tableDiv.innerHTML = '<table><thead><tr><th>Account</th><th>Avg Rating</th><th>FTR Pass</th></tr></thead><tbody>' + accounts.map(a => '<tr><td>'+a.account+'</td><td>'+(a.avgRating||0)+'%</td><td>'+(a.ftrPass||'-')+'</td></tr>').join('') + '</tbody></table>';
        },
        leave: () => {
            title.textContent = ' Leave Detail � This Month';
            const d = _execCache.leave || {};
            const byP = (d.currentMonth && d.currentMonth.byPerson) || {};
            const sorted = Object.entries(byP).filter(function(e){return e[1]>0}).sort(function(a,b){return b[1]-a[1]}).slice(0,12);
            const labels = sorted.map(function(e){return e[0].split(' ').slice(0,2).join(' ')});
            const vals = sorted.map(function(e){return e[1]});
            _kpiDetailChart = new Chart(ctx, { type:'bar', data:{ labels:labels, datasets:[{ label:'Days', data:vals, backgroundColor:'#f59e0b', borderRadius:6 }] }, options:{ responsive:true, plugins:{ legend:{ display:false } } } });
            tableDiv.innerHTML = '<table><thead><tr><th>Person</th><th>Days</th></tr></thead><tbody>' + sorted.map(function(e){return '<tr><td>'+e[0]+'</td><td>'+e[1]+'</td></tr>'}).join('') + '</tbody></table>';
        },
        risks: () => {
            title.textContent = ' Active Risk Details';
            const d = _execCache.gov || {};
            const risks = d.risks || [];
            tableDiv.innerHTML = '<table><thead><tr><th>Risk</th><th>Owner</th><th>Impact</th><th>Status</th></tr></thead><tbody>' + risks.map(r => '<tr><td>'+(r.description||r.risk||'-')+'</td><td>'+(r.owner||'-')+'</td><td>'+(r.impact||'-')+'</td><td>'+(r.status||'-')+'</td></tr>').join('') + '</tbody></table>';
            const severity = {}; risks.forEach(r => { severity[r.impact||'Unknown'] = (severity[r.impact||'Unknown']||0)+1; });
            _kpiDetailChart = new Chart(ctx, { type:'pie', data:{ labels:Object.keys(severity), datasets:[{ data:Object.values(severity), backgroundColor:['#ef4444','#f59e0b','#22c55e','#94a3b8'] }] }, options:{ responsive:true, plugins:{ legend:{ position:'bottom' } } } });
        },
        highlights: () => {
            title.textContent = ' Highlights & Lowlights';
            const d = _execCache.gov || {};
            const hl = d.highlights || []; const ll = d.lowlights || [];
            tableDiv.innerHTML = '<h4 style="margin:0 0 6px;font-size:.82rem"> Highlights</h4><ul style="margin:0 0 12px;padding-left:18px;font-size:.78rem">' + hl.map(h => '<li>'+(h.highlight||h)+'</li>').join('') + '</ul><h4 style="margin:0 0 6px;font-size:.82rem"> Lowlights</h4><ul style="margin:0;padding-left:18px;font-size:.78rem">' + ll.map(l => '<li>'+(l.lowlight||l)+'</li>').join('') + '</ul>';
            _kpiDetailChart = new Chart(ctx, { type:'doughnut', data:{ labels:['Highlights','Lowlights'], datasets:[{ data:[hl.length, ll.length], backgroundColor:['#22c55e','#ef4444'] }] }, options:{ responsive:true, plugins:{ legend:{ position:'bottom' } } } });
        }
    };

    if (configs[key]) configs[key]();
}

// ══════════════════════════════════════════════════════════════════
//  EXCEL DATA BROWSER — Tab-Based Raw Data View
// ══════════════════════════════════════════════════════════════════

let _allSheetsData = null;

function renderExcelBrowser(data) {
    _allSheetsData = data;
    const nav = document.getElementById('excelNav');
    const status = document.getElementById('excelBrowserStatus');
    if (!nav) return;

    const sources = data?.sources || [];
    if (!sources.length) {
        nav.innerHTML = '<div style="padding:16px;color:#999;font-size:.8rem">No Excel files found</div>';
        if (status) status.textContent = 'No data';
        return;
    }

    let totalTabs = 0;
    sources.forEach(s => totalTabs += (s.sheets || []).length);
    if (status) status.textContent = `${sources.length} files · ${totalTabs} tabs`;

    let html = '';
    sources.forEach((src, si) => {
        const sheets = src.sheets || [];
        html += `<div class="excel-nav-parent" data-source="${si}">
            <span>${src.icon || '📄'}</span>
            <span>${esc(src.name)}</span>
            <span class="nav-badge">${sheets.length}</span>
            <span class="nav-chevron">▶</span>
        </div>
        <div class="excel-nav-children">`;
        sheets.forEach((sh, ti) => {
            html += `<div class="excel-nav-child" data-source="${si}" data-tab="${ti}">
                <span>📋</span>
                <span>${esc(sh.name)}</span>
                <span class="nav-rows">${sh.totalRows || sh.rows?.length || 0}r</span>
            </div>`;
        });
        html += '</div>';
    });
    nav.innerHTML = html;

    // Parent toggle
    nav.querySelectorAll('.excel-nav-parent').forEach(p => {
        p.addEventListener('click', () => {
            p.classList.toggle('expanded');
        });
    });

    // Child click -> render sheet
    nav.querySelectorAll('.excel-nav-child').forEach(c => {
        c.addEventListener('click', () => {
            nav.querySelectorAll('.excel-nav-child').forEach(x => x.classList.remove('active'));
            c.classList.add('active');
            const si = parseInt(c.dataset.source);
            const ti = parseInt(c.dataset.tab);
            const src = sources[si];
            const sh = src.sheets[ti];
            renderExcelSheet(src, sh);
        });
    });

    // Auto-expand first source
    const firstParent = nav.querySelector('.excel-nav-parent');
    if (firstParent) firstParent.classList.add('expanded');
}

function renderExcelSheet(source, sheet) {
    const area = document.getElementById('excelTableArea');
    if (!area || !sheet) return;

    const headers = sheet.headers || [];
    const rows = sheet.rows || [];

    let html = `<div class="excel-table-header">
        <div class="excel-table-breadcrumb">
            <span>${source.icon || '📄'} ${esc(source.name)}</span>
            <span class="sep">›</span>
            <span>${esc(sheet.name)}</span>
        </div>
        <div class="excel-table-meta">
            <span class="excel-meta-badge">${sheet.totalRows || rows.length} rows</span>
            <span class="excel-meta-badge">${headers.length} columns</span>
            <input type="text" class="excel-search" placeholder="🔍 Filter rows..." id="excelSheetFilter">
        </div>
    </div>`;

    if (rows.length === 0) {
        html += '<div class="excel-table-placeholder">This sheet has no data rows</div>';
        area.innerHTML = html;
        return;
    }

    html += '<table class="excel-data-table"><thead><tr><th>#</th>';
    headers.forEach(h => {
        html += `<th>${esc(h)}</th>`;
    });
    html += '</tr></thead><tbody id="excelTbody">';

    rows.forEach((row, i) => {
        html += `<tr><td class="row-num">${i + 1}</td>`;
        headers.forEach(h => {
            html += `<td title="${esc(String(row[h] || ''))}">${esc(String(row[h] || ''))}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';
    if (sheet.totalRows > rows.length) {
        html += `<div class="excel-truncated">Showing ${rows.length} of ${sheet.totalRows} rows</div>`;
    }

    area.innerHTML = html;

    // Filter logic
    const filterInput = document.getElementById('excelSheetFilter');
    if (filterInput) {
        filterInput.addEventListener('input', () => {
            const term = filterInput.value.toLowerCase();
            const tbody = document.getElementById('excelTbody');
            if (!tbody) return;
            Array.from(tbody.rows).forEach(tr => {
                const text = tr.textContent.toLowerCase();
                tr.style.display = text.includes(term) ? '' : 'none';
            });
        });
    }
}
