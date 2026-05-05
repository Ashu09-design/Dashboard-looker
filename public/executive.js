/* ═══════════════════════════════════════════════════════════════
   Executive Dashboard — Client Logic (executive.js)
   Redesigned: Parent/Child Excel Tab Navigation
   ═══════════════════════════════════════════════════════════════ */

const TOKEN_KEY = 'exec_token';
const REFRESH_MS = 5 * 60 * 1000;
const $ = sel => document.querySelector(sel);
const $$ = sel => document.querySelectorAll(sel);

let chartFTE = null;
let chartKPI = null;
let chartSOW = null;
let allSheetsData = null; // cached sheet data

// ══════════════════════════════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
    checkAuth();
    bindUI();
    fetchAllExecData();
    fetchAllSheets();
    setInterval(fetchAllExecData, REFRESH_MS);
});

// ══════════════════════════════════════════════════════════════════
//  AUTH
// ══════════════════════════════════════════════════════════════════

function getToken() { return sessionStorage.getItem(TOKEN_KEY); }

async function checkAuth() {
    const token = getToken();
    if (!token) { window.location.href = '/login.html'; return; }
    try {
        const res = await fetch('/api/auth/verify', {
            headers: { Authorization: 'Bearer ' + token },
        });
        if (!res.ok) throw new Error('Unauthorized');
    } catch {
        sessionStorage.removeItem(TOKEN_KEY);
        window.location.href = '/login.html';
    }
}

function authHeaders() {
    return { Authorization: 'Bearer ' + getToken(), 'Content-Type': 'application/json' };
}

function logout() {
    sessionStorage.removeItem(TOKEN_KEY);
    window.location.href = '/login.html';
}

// ══════════════════════════════════════════════════════════════════
//  DATA FETCHING
// ══════════════════════════════════════════════════════════════════

async function fetchAllExecData() {
    showLoading(true);
    try {
        const [summary, kpi, sow, gov, leave, ftr, health] = await Promise.all([
            apiFetch('/api/exec/summary'),
            apiFetch('/api/exec/kpi-scorecards'),
            apiFetch('/api/exec/sow-financial'),
            apiFetch('/api/exec/governance-risks'),
            apiFetch('/api/exec/leave-impact'),
            apiFetch('/api/exec/ftr-metrics'),
            apiFetch('/api/exec/project-health'),
        ]);

        renderKPICards(summary);
        renderFTETrendChart(gov);
        renderKPIChart(kpi);
        renderSOWChart(sow);
        renderRiskPanel(gov);
        renderLeavePanel(leave);
        renderProjectHealthTable(health);
        renderHighlightsLowlights(gov);
        updateTimestamp();
    } catch (err) {
        console.error('Executive fetch error:', err);
        if (err.message && err.message.includes('401')) logout();
    }
    showLoading(false);
}

async function fetchAllSheets() {
    try {
        const data = await apiFetch('/api/exec/all-sheets');
        allSheetsData = data;
        buildSidebar(data.sources || []);
    } catch (err) {
        console.error('All-sheets fetch error:', err);
    }
}

async function apiFetch(url) {
    const res = await fetch(url, { headers: { Authorization: 'Bearer ' + getToken() } });
    if (res.status === 401) throw new Error('401');
    return res.json();
}

function showLoading(show) {
    const el = $('#loadingOverlay');
    if (show) el.classList.remove('hidden');
    else el.classList.add('hidden');
}

function updateTimestamp() {
    const el = $('#execLastUpdated');
    if (el) el.textContent = 'Updated: ' + new Date().toLocaleTimeString();
}

// ══════════════════════════════════════════════════════════════════
//  SIDEBAR — Build Parent/Child Navigation
// ══════════════════════════════════════════════════════════════════

function buildSidebar(sources) {
    const nav = $('#sidebarNav');
    // Keep the overview link, remove old dynamic items
    const existing = nav.querySelectorAll('.nav-parent-group');
    existing.forEach(el => el.remove());

    sources.forEach(source => {
        const group = document.createElement('div');
        group.className = 'nav-parent-group';

        // Parent item
        const parent = document.createElement('div');
        parent.className = 'nav-item nav-parent';
        parent.innerHTML = `
            <span class="nav-icon">${source.icon}</span>
            <span class="nav-label">${esc(source.name)}</span>
            <span class="nav-badge">${source.sheets.length}</span>
            <span class="nav-chevron">▸</span>
        `;
        parent.addEventListener('click', () => {
            group.classList.toggle('expanded');
        });

        // Child items
        const children = document.createElement('div');
        children.className = 'nav-children';
        source.sheets.forEach((sheet, idx) => {
            const child = document.createElement('div');
            child.className = 'nav-item nav-child';
            child.innerHTML = `
                <span class="nav-dot"></span>
                <span class="nav-label">${esc(sheet.name)}</span>
                <span class="nav-row-count">${sheet.totalRows}r</span>
            `;
            child.addEventListener('click', (e) => {
                e.stopPropagation();
                showSheet(source, sheet);
                // Active state
                $$('.nav-item').forEach(n => n.classList.remove('active'));
                child.classList.add('active');
            });
            children.appendChild(child);
        });

        group.appendChild(parent);
        group.appendChild(children);
        nav.appendChild(group);
    });
}

// ══════════════════════════════════════════════════════════════════
//  VIEW SWITCHING
// ══════════════════════════════════════════════════════════════════

function showOverview() {
    $$('.content-view').forEach(v => v.classList.remove('active'));
    $('#viewOverview').classList.add('active');
    $$('.nav-item').forEach(n => n.classList.remove('active'));
    $('.nav-overview').classList.add('active');
}

function showSheet(source, sheet) {
    $$('.content-view').forEach(v => v.classList.remove('active'));
    $('#viewSheet').classList.add('active');

    // Update breadcrumb
    $('#sheetParentName').textContent = `${source.icon} ${source.name}`;
    $('#sheetChildName').textContent = sheet.name;
    $('#sheetRowCount').textContent = `${sheet.totalRows} rows`;
    $('#sheetColCount').textContent = `${sheet.headers.length} columns`;

    // Render the table
    renderSheetTable(sheet);

    // Bind search
    const searchEl = $('#sheetSearch');
    searchEl.value = '';
    searchEl.oninput = () => {
        const q = searchEl.value.toLowerCase();
        const rows = $$('#sheetTableWrap tbody tr');
        rows.forEach(row => {
            row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
        });
    };
}

function renderSheetTable(sheet) {
    const wrap = $('#sheetTableWrap');
    if (!sheet.headers.length || !sheet.rows.length) {
        wrap.innerHTML = '<div class="sheet-empty">No data in this sheet.</div>';
        return;
    }

    const headers = sheet.headers;
    let html = '<table class="sheet-table"><thead><tr>';
    html += '<th class="row-num">#</th>';
    headers.forEach(h => {
        html += `<th>${esc(h)}</th>`;
    });
    html += '</tr></thead><tbody>';

    sheet.rows.forEach((row, i) => {
        html += '<tr>';
        html += `<td class="row-num">${i + 1}</td>`;
        headers.forEach(h => {
            const val = row[h] || '';
            html += `<td>${esc(val)}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';
    if (sheet.totalRows > 500) {
        html += `<div class="sheet-truncated">Showing 500 of ${sheet.totalRows} rows</div>`;
    }
    wrap.innerHTML = html;
}

// ══════════════════════════════════════════════════════════════════
//  UI BINDINGS
// ══════════════════════════════════════════════════════════════════

function bindUI() {
    $('#btnExecRefresh').addEventListener('click', () => {
        const btn = $('#btnExecRefresh');
        btn.classList.add('refreshing');
        Promise.all([fetchAllExecData(), fetchAllSheets()])
            .then(() => setTimeout(() => btn.classList.remove('refreshing'), 800));
    });

    $('#btnLogout').addEventListener('click', logout);
    $('#btnExecPresent').addEventListener('click', () => togglePresent(true));
    $('#exitExecPresent').addEventListener('click', () => togglePresent(false));
    document.addEventListener('keydown', e => { if (e.key === 'Escape') togglePresent(false); });
    $('#btnExecPrint').addEventListener('click', () => window.print());
    $('#btnExecSources').addEventListener('click', () => $('#execSourceModal').classList.remove('hidden'));
    $('#closeExecModal').addEventListener('click', () => $('#execSourceModal').classList.add('hidden'));
    $('#execSourceModal').addEventListener('click', e => {
        if (e.target === $('#execSourceModal')) $('#execSourceModal').classList.add('hidden');
    });

    // Modal tabs
    $$('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            $$('.tab-btn').forEach(b => b.classList.remove('active'));
            $$('.tab-content').forEach(c => c.classList.remove('active'));
            btn.classList.add('active');
            $(`#tab${capitalize(btn.dataset.tab)}`).classList.add('active');
        });
    });

    // Upload + Connect
    $('#btnUploadAll').addEventListener('click', uploadFiles);
    $('#btnConnectUrls').addEventListener('click', connectUrls);

    // Overview nav
    $('.nav-overview').addEventListener('click', showOverview);
}

function togglePresent(on) {
    document.body.classList.toggle('exec-present-mode', on);
    $('#execPresentBar').classList.toggle('hidden', !on);
    if (on) document.documentElement.requestFullscreen?.().catch(() => { });
    else if (document.fullscreenElement) document.exitFullscreen?.();
}

function capitalize(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

// ══════════════════════════════════════════════════════════════════
//  KPI CARDS
// ══════════════════════════════════════════════════════════════════

function renderKPICards(data) {
    if (!data) return;
    $('#valTeamSize').textContent = data.teamSize || 0;
    $('#valActiveProjects').textContent = data.activeProjects || 0;
    $('#valSOWValue').textContent = data.totalSOWValue ? formatCurrency(data.totalSOWValue) : '—';
    $('#valKPIMetRate').textContent = data.kpiMetRate ? data.kpiMetRate + '%' : '—';
    $('#valFTRRating').textContent = data.ftrAvgRating ? data.ftrAvgRating + '%' : '—';
    $('#valOnLeave').textContent = data.onLeaveToday || 0;
}

function formatCurrency(val) {
    if (val >= 1000000) return '$' + (val / 1000000).toFixed(1) + 'M';
    if (val >= 1000) return '$' + (val / 1000).toFixed(0) + 'K';
    return '$' + val.toFixed(0);
}

// ══════════════════════════════════════════════════════════════════
//  CHARTS
// ══════════════════════════════════════════════════════════════════

function renderFTETrendChart(govData) {
    const ctx = $('#chartFTETrend');
    if (!ctx) return;
    const trend = govData?.fteTrend || [];
    const labels = trend.map(t => t.month);
    const values = trend.map(t => t.totalFTE);

    if (chartFTE) chartFTE.destroy();
    chartFTE = new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: [{
                label: 'Total FTE',
                data: values,
                borderColor: '#06b6d4',
                backgroundColor: 'rgba(6,182,212,.1)',
                fill: true,
                tension: .4,
                pointRadius: 4,
                pointBackgroundColor: '#06b6d4',
            }],
        },
        options: {
            responsive: true,
            plugins: { legend: { labels: { color: '#94a3b8', font: { size: 11 } } } },
            scales: {
                x: { ticks: { color: '#64748b' }, grid: { color: 'rgba(51,65,85,.3)' } },
                y: { ticks: { color: '#64748b' }, grid: { color: 'rgba(51,65,85,.3)' }, beginAtZero: true },
            },
        },
    });
}

function renderKPIChart(kpiData) {
    const ctx = $('#chartKPI');
    if (!ctx) return;
    const kpis = (kpiData?.kpis || []).slice(0, 10);
    const labels = kpis.map(k => truncate(k.metricName || k.metric || k.name || '', 25));
    const targets = kpis.map(k => (k.target * 100));
    const actuals = kpis.map(k => (k.actual * 100));

    if (chartKPI) chartKPI.destroy();
    chartKPI = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                { label: 'Target %', data: targets, backgroundColor: 'rgba(168,85,247,.6)', borderRadius: 4 },
                { label: 'Actual %', data: actuals, backgroundColor: 'rgba(6,182,212,.7)', borderRadius: 4 },
            ],
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            plugins: { legend: { labels: { color: '#94a3b8', font: { size: 11 } } } },
            scales: {
                x: { ticks: { color: '#64748b' }, grid: { color: 'rgba(51,65,85,.3)' }, max: 120 },
                y: { ticks: { color: '#94a3b8', font: { size: 10 } }, grid: { display: false } },
            },
        },
    });
}

function renderSOWChart(sowData) {
    const ctx = $('#chartSOW');
    if (!ctx) return;
    const summary = sowData?.summary || {};
    const statusMap = summary.projectsByStatus || {};
    const labels = Object.keys(statusMap);
    const values = Object.values(statusMap);
    const colors = ['#22c55e', '#f59e0b', '#06b6d4', '#a855f7', '#ef4444', '#ec4899'];

    if (chartSOW) chartSOW.destroy();
    chartSOW = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels,
            datasets: [{ data: values, backgroundColor: colors.slice(0, labels.length), borderWidth: 0 }],
        },
        options: {
            responsive: true,
            plugins: { legend: { position: 'bottom', labels: { color: '#94a3b8', font: { size: 10 }, padding: 12 } } },
        },
    });
}

// ══════════════════════════════════════════════════════════════════
//  RISK PANEL
// ══════════════════════════════════════════════════════════════════

function renderRiskPanel(govData) {
    const risks = govData?.risks || [];
    const body = $('#riskPanelBody');
    const count = $('#riskCount');
    const activeRisks = risks.filter(r => r.status && r.status.toLowerCase() === 'ongoing');
    count.textContent = activeRisks.length;

    if (activeRisks.length === 0) {
        body.innerHTML = '<div style="color:#64748b">No active governance risks. ✅</div>';
        return;
    }

    body.innerHTML = activeRisks.map(r => {
        const impactClass = (r.impact || '').toLowerCase().includes('loss') ? 'impact-high'
            : (r.impact || '').toLowerCase().includes('delay') ? 'impact-medium' : 'impact-low';
        return `
      <div class="risk-item">
        <div class="risk-project">${esc(r.project)} <span style="color:#64748b;font-weight:400">— ${esc(r.pm)}</span></div>
        <div class="risk-text">${esc(r.description || r.risk || '')}</div>
        <span class="risk-impact ${impactClass}">${esc(r.impact || 'Unknown')}</span>
      </div>`;
    }).join('');
}

// ══════════════════════════════════════════════════════════════════
//  LEAVE PANEL
// ══════════════════════════════════════════════════════════════════

function renderLeavePanel(leaveData) {
    const body = $('#leavePanelBody');
    const cm = leaveData?.currentMonth || {};
    const byPerson = cm.byPerson || {};
    const onToday = cm.onLeaveToday || [];

    let html = '';
    if (onToday.length > 0) {
        html += `<div style="margin-bottom:10px;color:#f59e0b;font-weight:600">📌 On Leave Today (${onToday.length}):</div>`;
        html += onToday.map(p => `<div class="leave-item"><span class="leave-name">${esc(p.name)}</span><span style="color:#94a3b8">${esc(p.type)}</span></div>`).join('');
        html += '<hr style="border-color:#334155;margin:12px 0">';
    }

    html += `<div style="margin-bottom:8px;font-weight:600;color:#f1f5f9">Month Summary (Total: ${cm.totalLeaves || 0} leaves)</div>`;

    const sorted = Object.entries(byPerson)
        .filter(([_, count]) => count > 0)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15);

    if (sorted.length === 0) {
        html += '<div style="color:#64748b">No leave data for current month.</div>';
    } else {
        html += sorted.map(([name, count]) => `
      <div class="leave-item">
        <span class="leave-name">${esc(name)}</span>
        <span class="leave-count">${count} day(s)</span>
      </div>`).join('');
    }

    body.innerHTML = html;
}

// ══════════════════════════════════════════════════════════════════
//  PROJECT HEALTH TABLE
// ══════════════════════════════════════════════════════════════════

function renderProjectHealthTable(healthData) {
    const tbody = $('#healthTableBody');
    if (!healthData || healthData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" style="color:#64748b;text-align:center;padding:20px">No project data</td></tr>';
        return;
    }

    tbody.innerHTML = healthData.map(p => {
        const cls = p.health === 'Green' ? 'health-green'
            : p.health === 'Red' ? 'health-red' : 'health-amber';
        return `
      <tr>
        <td><strong>${esc(p.projectName)}</strong></td>
        <td>${esc(p.client)}</td>
        <td>${esc(p.pm)}</td>
        <td>${esc(p.status)}</td>
        <td>${esc(p.sowStatus)}</td>
        <td>${esc(p.poStatus)}</td>
        <td><span class="health-badge ${cls}">${p.health}</span></td>
        <td style="font-size:.75rem;color:#94a3b8">${esc(p.reasons)}</td>
      </tr>`;
    }).join('');
}

// ══════════════════════════════════════════════════════════════════
//  HIGHLIGHTS & LOWLIGHTS
// ══════════════════════════════════════════════════════════════════

function renderHighlightsLowlights(govData) {
    const highlights = govData?.highlights || [];
    const lowlights = govData?.lowlights || [];

    const hlEl = $('#highlightsList');
    const hlItems = highlights.filter(h => h.highlight);
    if (hlItems.length === 0) {
        hlEl.innerHTML = '<div style="color:#64748b">No highlights reported.</div>';
    } else {
        hlEl.innerHTML = hlItems.slice(0, 10).map(h => `
      <div class="hl-item">
        <div class="hl-project">${esc(h.project)} — ${esc(h.month)}</div>
        <div class="hl-text">${esc(h.highlight)}</div>
      </div>`).join('');
    }

    const llEl = $('#lowlightsList');
    const llItems = lowlights.length > 0 ? lowlights : highlights.filter(h => h.lowlight);
    if (llItems.length === 0) {
        llEl.innerHTML = '<div style="color:#64748b">No lowlights reported.</div>';
    } else {
        llEl.innerHTML = llItems.slice(0, 10).map(h => `
      <div class="hl-item">
        <div class="hl-project">${esc(h.project)} — ${esc(h.month)}</div>
        <div class="hl-text">${esc(h.lowlight)}</div>
      </div>`).join('');
    }
}

// ══════════════════════════════════════════════════════════════════
//  DATA SOURCE MANAGEMENT
// ══════════════════════════════════════════════════════════════════

async function uploadFiles() {
    const form = $('#uploadForm');
    const formData = new FormData(form);
    const statusEl = $('#execSourceStatus');

    let hasFile = false;
    for (const [, file] of formData.entries()) {
        if (file && file.size > 0) { hasFile = true; break; }
    }
    if (!hasFile) {
        statusEl.className = 'source-status-box error';
        statusEl.textContent = 'Please select at least one file to upload.';
        statusEl.classList.remove('hidden');
        return;
    }

    statusEl.className = 'source-status-box';
    statusEl.textContent = 'Uploading…';
    statusEl.classList.remove('hidden');

    try {
        const res = await fetch('/api/exec/upload-sources', {
            method: 'POST',
            headers: { Authorization: 'Bearer ' + getToken() },
            body: formData,
        });
        const data = await res.json();
        if (data.success) {
            statusEl.className = 'source-status-box success';
            statusEl.textContent = '✓ ' + data.message;
            setTimeout(() => {
                fetchAllExecData();
                fetchAllSheets();
                $('#execSourceModal').classList.add('hidden');
            }, 1200);
        } else {
            statusEl.className = 'source-status-box error';
            statusEl.textContent = '✗ ' + (data.error || 'Upload failed');
        }
    } catch (err) {
        statusEl.className = 'source-status-box error';
        statusEl.textContent = '✗ Network error: ' + err.message;
    }
}

async function connectUrls() {
    const sources = {};
    const keys = ['ftr', 'team', 'sow', 'governance', 'leave', 'kpi'];
    const ids = ['urlFtr', 'urlTeam', 'urlSow', 'urlGovernance', 'urlLeave', 'urlKpi'];

    keys.forEach((key, i) => {
        const url = $(`#${ids[i]}`).value.trim();
        if (url) sources[key] = { url, type: 'sharepoint' };
    });

    const statusEl = $('#execSourceStatus');
    if (Object.keys(sources).length === 0) {
        statusEl.className = 'source-status-box error';
        statusEl.textContent = 'Please enter at least one URL.';
        statusEl.classList.remove('hidden');
        return;
    }

    statusEl.className = 'source-status-box';
    statusEl.textContent = 'Connecting…';
    statusEl.classList.remove('hidden');

    try {
        const res = await fetch('/api/exec/connect-sources', {
            method: 'POST',
            headers: authHeaders(),
            body: JSON.stringify({ sources }),
        });
        const data = await res.json();
        if (data.success) {
            statusEl.className = 'source-status-box success';
            statusEl.textContent = '✓ ' + data.message;
            setTimeout(() => {
                fetchAllExecData();
                fetchAllSheets();
                $('#execSourceModal').classList.add('hidden');
            }, 1200);
        } else {
            statusEl.className = 'source-status-box error';
            statusEl.textContent = '✗ ' + (data.error || 'Connection failed');
        }
    } catch (err) {
        statusEl.className = 'source-status-box error';
        statusEl.textContent = '✗ Network error: ' + err.message;
    }
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

function truncate(str, len) {
    if (!str) return '';
    return str.length > len ? str.slice(0, len) + '…' : str;
}
