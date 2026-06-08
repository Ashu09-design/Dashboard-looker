/* ═══════════════════════════════════════════════════════════════
   Manager Dashboard — Tabbed Curated Views (executive.js)
   Embedded inside the Daily Standup page (index.html).
   Wrapped in an IIFE so nothing collides with app.js globals.
   Exposes: window.ManagerDash = { open, close }
   ═══════════════════════════════════════════════════════════════ */

(function () {
'use strict';

const TOKEN_KEY = 'exec_token';
const REFRESH_MS = 30 * 1000; // 30s auto-refresh for live data from Google Sheets / SharePoint
const $ = sel => document.querySelector(sel);
const $$ = sel => document.querySelectorAll(sel);

// ── Module state ──────────────────────────────────────────────────
const state = {
    data: {},
    activeTab: 'overview',
    filters: {},
    charts: {},
    bound: false,
    intervalId: null,
};

const TABS = [
    { id: 'overview', label: 'Overview', icon: '🏠' },
    { id: 'team', label: 'Team Details', icon: '👥' },
    { id: 'projects', label: 'Projects', icon: '📂' },
    { id: 'sow', label: 'SOW & PO', icon: '💰' },
    { id: 'kpis', label: 'KPIs', icon: '🎯' },
    { id: 'ftr', label: 'FTR', icon: '✅' },
    { id: 'leave', label: 'Leave', icon: '🏖️' },
    { id: 'risks', label: 'Active Risks', icon: '⚠️' },
    { id: 'poc', label: 'POV / POC', icon: '🧪' },
];

// ══════════════════════════════════════════════════════════════════
//  PUBLIC API
// ══════════════════════════════════════════════════════════════════

function open() {
    if (!state.bound) { bindUI(); state.bound = true; }
    renderTabNav();
    fetchAllExecData(true); // First load animates
    checkSourceConnection();
    if (!state.intervalId) state.intervalId = setInterval(() => fetchAllExecData(false), REFRESH_MS);
    if (!state.sourceCheckId) state.sourceCheckId = setInterval(checkSourceConnection, 60000);
    if (!state.chatbotInitialized) { initChatbot(); state.chatbotInitialized = true; } else { showChatbot(); }
}

function close() {
    if (state.intervalId) { clearInterval(state.intervalId); state.intervalId = null; }
    if (state.sourceCheckId) { clearInterval(state.sourceCheckId); state.sourceCheckId = null; }
    destroyCharts();
    hideChatbot();
}

window.ManagerDash = { open, close };

// ══════════════════════════════════════════════════════════════════
//  AUTH HELPERS
// ══════════════════════════════════════════════════════════════════

function getToken() { return sessionStorage.getItem(TOKEN_KEY); }

function authHeaders() {
    return { Authorization: 'Bearer ' + getToken(), 'Content-Type': 'application/json' };
}

function onUnauthorized() {
    close();
    document.getElementById('btnExecLogout')?.click();
}

// ══════════════════════════════════════════════════════════════════
//  DATA FETCHING
// ══════════════════════════════════════════════════════════════════

async function apiFetch(url) {
    const res = await fetch(url, { headers: { Authorization: 'Bearer ' + getToken() } });
    if (res.status === 401) throw new Error('401');
    return res.json();
}

async function fetchAllExecData(triggerAnimation = false) {
    showLoading(true);
    try {
        const [summary, kpi, sow, gov, leave, ftr, health, team, poc] = await Promise.all([
            apiFetch('/api/exec/summary'),
            apiFetch('/api/exec/kpi-scorecards'),
            apiFetch('/api/exec/sow-financial'),
            apiFetch('/api/exec/governance-risks'),
            apiFetch('/api/exec/leave-impact'),
            apiFetch('/api/exec/ftr-metrics'),
            apiFetch('/api/exec/project-health'),
            apiFetch('/api/exec/team-capacity'),
            apiFetch('/api/exec/poc'),
        ]);
        state.data = { summary, kpi, sow, gov, leave, ftr, health, team, poc };
        updateTimestamp();
        renderTab(triggerAnimation);
    } catch (err) {
        console.error('Manager dashboard fetch error:', err);
        if (err.message && err.message.includes('401')) { onUnauthorized(); return; }
    }
    showLoading(false);
}

/** Check if a data source (Google Sheet / SharePoint) is connected */
async function checkSourceConnection() {
    try {
        const status = await apiFetch('/api/exec/source-status');
        state.sourceConnected = status || {};
        renderConnectionBadge();
    } catch (_) { /* ignore */ }
}

function renderConnectionBadge() {
    let badge = document.getElementById('mdLiveBadge');
    if (!badge) {
        badge = document.createElement('span');
        badge.id = 'mdLiveBadge';
        badge.className = 'live-badge';
        const toolbar = document.querySelector('.md-toolbar-right');
        if (toolbar) toolbar.prepend(badge);
    }
    const s = state.sourceConnected || {};
    const anyConnected = Object.values(s).some(v => v && typeof v === 'object' && v.loaded);
    if (anyConnected) {
        const connectedCount = Object.values(s).filter(v => v && typeof v === 'object' && v.loaded).length;
        badge.innerHTML = `<span class="live-dot"></span> Live · ${connectedCount} source(s)`;
        badge.classList.add('connected');
        badge.classList.remove('disconnected');
    } else {
        badge.innerHTML = `<span class="live-dot off"></span> No live source`;
        badge.classList.add('disconnected');
        badge.classList.remove('connected');
    }
}

function showLoading(show) {
    const el = $('#mdLoading');
    if (!el) return;
    el.classList.toggle('hidden', !show);
}

function updateTimestamp() {
    const el = $('#mdUpdated');
    if (el) el.textContent = 'Updated: ' + new Date().toLocaleTimeString();
}

// ══════════════════════════════════════════════════════════════════
//  TAB NAVIGATION
// ══════════════════════════════════════════════════════════════════

function renderTabNav() {
    const nav = $('#tabNav');
    if (!nav) return;
    nav.innerHTML = TABS.map(t =>
        `<button class="tab-link ${t.id === state.activeTab ? 'active' : ''}" data-tab="${t.id}">
            <span class="tab-ico">${t.icon}</span>${esc(t.label)}
        </button>`).join('');
    nav.querySelectorAll('.tab-link').forEach(btn => {
        btn.addEventListener('click', () => {
            state.activeTab = btn.dataset.tab;
            renderTabNav();
            renderTab(true); // Animate tab change
            $('#execSection')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
        });
    });
}

// ══════════════════════════════════════════════════════════════════
//  FILTER BAR
// ══════════════════════════════════════════════════════════════════

const TAB_FILTERS = {
    overview: ['rm', 'client', 'month', 'project'],
    team: ['rm', 'client', 'project'],
    projects: ['client', 'month'],
    sow: ['client', 'project', 'startDate', 'endDate'],
    kpis: ['client', 'project', 'kpiMetric'],
    ftr: ['client', 'month', 'project'],
    leave: ['rm', 'client', 'project'],
    risks: ['client', 'project', 'month'],
    poc: ['spoc'],
};

function getClientProjects(clientName) {
    if (!clientName) return new Set();
    const sowProjects = state.data.sow?.projects || [];
    const matched = sowProjects.filter(p => p.client && p.projectName &&
        p.client.toLowerCase() === clientName.toLowerCase());
    const projectNames = new Set(matched.map(p => p.projectName.toLowerCase()));
    
    const kpiMetrics = state.data.kpi?.metrics || [];
    kpiMetrics.forEach(m => {
        if (m.account && m.account.toLowerCase() === clientName.toLowerCase() && m.project) {
            projectNames.add(m.project.toLowerCase());
        }
    });
    return projectNames;
}


function uniqueSorted(arr) {
    return [...new Set(arr.filter(v => v != null && String(v).trim() !== ''))].sort();
}

function filterOptions(key) {
    const d = state.data;
    switch (key) {
        case 'office':
            return uniqueSorted((d.team?.activeMembers || []).map(m => m.location));
        case 'role':
            return uniqueSorted((d.team?.activeMembers || []).map(m => m.role));
        case 'designation':
            return uniqueSorted((d.team?.activeMembers || []).map(m => m.designation));
        case 'rm': {
            const rmMap = getRMAliasMap();
            return [...new Set(Object.values(rmMap))].sort();
        }
        case 'project':
            return uniqueSorted([
                ...(d.health || []).map(p => p.projectName),
                ...(d.sow?.projects || []).map(p => p.projectName),
                ...(d.kpi?.scorecardProjects || []).map(p => p.project),
            ]);
        case 'client':
            return uniqueSorted([
                ...(d.sow?.projects || []).map(p => p.client),
                ...(d.kpi?.metrics || []).map(m => m.account),
                ...(d.ftr?.accounts || []).map(a => a.account),
            ]);
        case 'month':
            return uniqueSorted([
                ...(d.kpi?.metrics || []).map(m => m.month),
                ...(d.kpi?.scorecardProjects || []).map(p => p.month),
                ...(d.gov?.risks || []).map(r => r.month),
                ...(d.ftr?.qaMetrics || []).map(q => monthLabel(q.qaDate)),
            ]);
        case 'kpiMetric':
            return uniqueSorted(['Timeliness (SLA%)', 'Utilization (Effort %)', 'Quality (FTR %)', 'CSAT (Frequency)', 'CSAT (Rating)', 'Client Training %']);
        case 'pocStatus':
            return uniqueSorted((d.poc?.pocs || []).map(p => p.status));
        case 'manager':
            return uniqueSorted((d.poc?.pocs || []).map(p => p.manager));
        case 'spoc':
            return uniqueSorted((d.poc?.pocs || []).map(p => p.spoc).filter(Boolean));
        default:
            return [];
    }
}

// ── Smart name deduplication for RM dropdown ─────────────────────
// Names like "Bhavani V", "Bhavani Venkatesh", "Bhavani venkatesh"
// are recognized as the same person → longest version kept.

function deduplicateNames(names) {
    const trimmed = names.filter(n => n && n.trim()).map(n => n.trim());
    if (!trimmed.length) return {};

    // 1) Group exact case-insensitive duplicates, prefer proper-cased
    const byLower = {};
    trimmed.forEach(n => {
        const key = n.toLowerCase();
        if (!byLower[key] || (n[0] === n[0].toUpperCase() && byLower[key][0] !== byLower[key][0].toUpperCase())) {
            byLower[key] = n;
        }
    });

    // 2) Sort longest first so most-complete name becomes canonical
    const unique = Object.values(byLower);
    const sorted = [...unique].sort((a, b) => b.length - a.length);

    const aliasMap = {};   // lowercase variant → canonical display name
    const consumed = new Set();

    for (const name of sorted) {
        const lower = name.toLowerCase();
        if (consumed.has(lower)) continue;

        const words = lower.split(/\s+/);
        const firstName = words[0];

        // This name is canonical (longest not-yet-consumed)
        aliasMap[lower] = name;
        consumed.add(lower);

        // Absorb shorter names that share the first name and whose every
        // word is a prefix-of or prefixed-by a word in the longer name.
        for (const other of sorted) {
            const ol = other.toLowerCase();
            if (consumed.has(ol)) continue;
            const ow = ol.split(/\s+/);
            if (ow[0] !== firstName) continue;

            const isSubset = ow.every(o =>
                words.some(w => w.startsWith(o) || o.startsWith(w)));
            if (isSubset) {
                aliasMap[ol] = name;
                consumed.add(ol);
            }
        }
    }
    return aliasMap;
}

let _rmAliasMap = null;
let _rmAliasVersion = null;

function getRMAliasMap() {
    const members = state.data.team?.activeMembers || [];
    const version = members.length + '|' + (members[0]?.name || '');
    if (_rmAliasMap && _rmAliasVersion === version) return _rmAliasMap;
    const allMgrs = members.map(m => m.manager).filter(Boolean);
    _rmAliasMap = deduplicateNames(allMgrs);
    _rmAliasVersion = version;
    return _rmAliasMap;
}

/** Map any manager name variant to its canonical form */
function canonicalRM(name) {
    if (!name) return '';
    const map = getRMAliasMap();
    return map[name.toLowerCase().trim()] || name;
}

/** True if memberManager (any variant) matches the selected RM canonical name */
function rmMatches(memberManager, selectedRM) {
    if (!selectedRM) return true;
    if (!memberManager) return false;
    return canonicalRM(memberManager) === selectedRM;
}

const FILTER_LABELS = {
    office: '🏢 Office', role: '👤 Role', designation: '👤 Designation', rm: '👔 RM',
    project: '📂 Project', client: '🏷️ Client',
    month: '📅 Month', kpiMetric: '📊 KPI Metric',
    pocStatus: '🚦 Status', manager: '👔 Manager', spoc: '👤 SPOC',
    startDate: '📅 Start Date', endDate: '📅 End Date',
};

function renderFilterBar() {
    const bar = $('#filterBar');
    if (!bar) return;
    const keys = TAB_FILTERS[state.activeTab] || [];
    if (keys.length === 0) { bar.innerHTML = ''; bar.classList.add('hidden'); return; }
    bar.classList.remove('hidden');

    state.filters[state.activeTab] = state.filters[state.activeTab] || {};
    const cur = state.filters[state.activeTab];

    let html = '<span class="filter-bar-label">🔎 Filters:</span>';
    keys.forEach(key => {
        if (key === 'startDate' || key === 'endDate') {
            html += `<label class="filter-ctrl">${FILTER_LABELS[key] || key}
                <input type="date" data-filter="${key}" value="${cur[key] || ''}" /></label>`;
        } else {
            const opts = filterOptions(key);
            html += `<label class="filter-ctrl">${FILTER_LABELS[key] || key}
                <select data-filter="${key}">
                    <option value="">All</option>
                    ${opts.map(o => `<option value="${esc(o)}" ${cur[key] === o ? 'selected' : ''}>${esc(o)}</option>`).join('')}
                </select></label>`;
        }
    });
    html += '<button class="filter-clear" id="btnClearFilters">Clear</button>';
    bar.innerHTML = html;

    bar.querySelectorAll('select[data-filter]').forEach(sel => {
        sel.addEventListener('change', () => {
            cur[sel.dataset.filter] = sel.value;
            renderTab();
        });
    });
    bar.querySelectorAll('input[data-filter]').forEach(inp => {
        inp.addEventListener('change', () => {
            cur[inp.dataset.filter] = inp.value;
            renderTab();
        });
    });
    $('#btnClearFilters').addEventListener('click', () => {
        state.filters[state.activeTab] = {};
        renderTab();
    });
}

function getFilter(key) {
    return (state.filters[state.activeTab] || {})[key] || '';
}

// ══════════════════════════════════════════════════════════════════
//  TAB DISPATCH
// ══════════════════════════════════════════════════════════════════

function renderTab(triggerAnimation = false) {
    renderFilterBar();
    destroyCharts();
    const el = $('#tabContent');
    if (!el) return;
    
    if (triggerAnimation) {
        // Reset and trigger smooth fade-in-slide animation on tab change
        el.classList.remove('fade-in-slide');
        void el.offsetWidth; // Force reflow
    }
    
    el.innerHTML = '';
    const fn = {
        overview: tabOverview, team: tabTeam, projects: tabProjects, sow: tabSow,
        kpis: tabKpis, ftr: tabFtr, leave: tabLeave, risks: tabRisks, poc: tabPoc,
    }[state.activeTab];
    if (fn) fn(el);
    
    if (triggerAnimation) {
        el.classList.add('fade-in-slide');
    }
}

function destroyCharts() {
    Object.values(state.charts).forEach(c => { try { c.destroy(); } catch (_) {} });
    state.charts = {};
}

// ══════════════════════════════════════════════════════════════════
//  DOM / RENDER HELPERS
// ══════════════════════════════════════════════════════════════════

function section(title, sub) {
    return `<div class="sec-head"><h2>${title}</h2>${sub ? `<span class="sec-sub">${esc(sub)}</span>` : ''}</div>`;
}

function statCard(icon, value, label, cls) {
    return `<div class="kpi-card ${cls || ''}">
        <div class="kpi-icon">${icon}</div>
        <div class="kpi-content">
            <div class="kpi-value">${value}</div>
            <div class="kpi-label">${esc(label)}</div>
        </div></div>`;
}

function dataTable(headers, rows, opts) {
    opts = opts || {};
    if (!rows.length) return `<div class="empty-note">${esc(opts.empty || 'No data available.')}</div>`;
    let h = `<div class="tbl-wrap"><table class="m-table"><thead><tr>`;
    headers.forEach(hd => { h += `<th>${esc(hd)}</th>`; });
    h += `</tr></thead><tbody>`;
    rows.forEach(r => {
        h += '<tr>';
        r.forEach(c => { h += `<td>${c == null ? '' : c}</td>`; });
        h += '</tr>';
    });
    h += `</tbody></table></div>`;
    if (opts.note) h += `<div class="tbl-note">${esc(opts.note)}</div>`;
    return h;
}

function pivotTable(rowKeys, colKeys, cellFn, opts) {
    opts = opts || {};
    if (!rowKeys.length || !colKeys.length) {
        return `<div class="empty-note">${esc(opts.empty || 'Not enough data for this view.')}</div>`;
    }
    let h = `<div class="tbl-wrap"><table class="m-table pivot"><thead><tr><th>${esc(opts.corner || '')}</th>`;
    colKeys.forEach(c => { h += `<th>${esc(c)}</th>`; });
    h += `</tr></thead><tbody>`;
    rowKeys.forEach(rk => {
        h += `<tr><td class="pivot-row-h">${esc(rk)}</td>`;
        colKeys.forEach(ck => {
            const v = cellFn(rk, ck);
            h += `<td>${v == null || v === '' ? '<span class="muted">—</span>' : v}</td>`;
        });
        h += '</tr>';
    });
    h += `</tbody></table></div>`;
    if (opts.note) h += `<div class="tbl-note">${esc(opts.note)}</div>`;
    return h;
}

function badge(text, cls) { return `<span class="badge ${cls || ''}">${esc(text)}</span>`; }

function chartCanvas(id) { return `<canvas id="${id}"></canvas>`; }

function makeChart(id, config) {
    const ctx = document.getElementById(id);
    if (!ctx || typeof Chart === 'undefined') return;
    state.charts[id] = new Chart(ctx, config);
}

const CHART_AXES = {
    x: { ticks: { color: '#64748b', font: { size: 10 } }, grid: { color: 'rgba(51,65,85,.25)' } },
    y: { ticks: { color: '#64748b', font: { size: 10 } }, grid: { color: 'rgba(51,65,85,.25)' }, beginAtZero: true },
};
const CHART_LEGEND = { labels: { color: '#94a3b8', font: { size: 11 } } };

// ══════════════════════════════════════════════════════════════════
//  TAB: OVERVIEW
// ══════════════════════════════════════════════════════════════════

function getRmPmAndMemberNames(rmName) {
    const pmNames = new Set();
    if (!rmName) return pmNames;
    pmNames.add(rmName.toLowerCase());
    const members = (state.data.team?.activeMembers || []).filter(m => rmMatches(m.manager, rmName));
    members.forEach(m => pmNames.add(m.name.toLowerCase()));
    return pmNames;
}

function getRmProjects(rmName) {
    if (!rmName) return new Set();
    const pmNames = getRmPmAndMemberNames(rmName);
    const sowProjects = state.data.sow?.projects || [];
    const matched = sowProjects.filter(p => p.pm && p.projectName && pmNames.has(p.pm.toLowerCase()));
    return new Set(matched.map(p => p.projectName.toLowerCase()));
}

// PMs running the given client's projects.
function pmsForClient(client) {
    if (!client) return [];
    const clientProjs = getClientProjects(client);
    const sowProjects = state.data.sow?.projects || [];
    return uniqueSorted(sowProjects
        .filter(p => p.projectName && clientProjs.has(p.projectName.toLowerCase()))
        .map(p => p.pm));
}

// PMs running a specific project.
function pmsForProject(projectName) {
    if (!projectName) return [];
    const sowProjects = state.data.sow?.projects || [];
    return uniqueSorted(sowProjects
        .filter(p => p.projectName === projectName)
        .map(p => p.pm));
}

// Members under the given PMs — the PM themselves or anyone reporting to
// them. Returns [] when pmNames is empty so an unresolved filter shows
// nothing rather than the whole team.
function membersForPMs(members, pmNames) {
    const pms = pmNames.filter(Boolean).map(p => p.toLowerCase().trim());
    if (!pms.length) return [];
    return members.filter(m => {
        const nm = (m.name || '').toLowerCase().trim();
        const mgr = (m.manager || '').toLowerCase().trim();
        return pms.some(pm =>
            nm === pm || mgr === pm ||
            (nm && (nm.includes(pm) || pm.includes(nm))));
    });
}

function tabOverview(el) {
    const rm = getFilter('rm'), month = getFilter('month'), proj = getFilter('project'), client = getFilter('client');

    // 1. Filter Team Size
    let members = state.data.team?.activeMembers || [];
    if (rm) members = members.filter(m => rmMatches(m.manager, rm));
    if (client) members = membersForPMs(members, pmsForClient(client));
    if (proj) members = membersForPMs(members, pmsForProject(proj));
    const teamSize = members.length;

    // 2. Filter Active Projects
    let healthRows = state.data.health || [];
    if (proj) healthRows = healthRows.filter(p => p.projectName === proj);
    if (client) healthRows = healthRows.filter(p => p.client === client);
    if (rm) {
        const rmProjs = getRmProjects(rm);
        healthRows = healthRows.filter(p => p.projectName && rmProjs.has(p.projectName.toLowerCase()));
    }
    const activeProjects = healthRows.length;

    // 3. Filter KPI Met Rate
    let kpiMetrics = state.data.kpi?.metrics || [];
    if (proj) kpiMetrics = kpiMetrics.filter(m => (m.project || '').toLowerCase().includes(proj.toLowerCase()));
    if (month) kpiMetrics = kpiMetrics.filter(m => m.month === month);
    if (client) kpiMetrics = kpiMetrics.filter(m => m.account === client);
    if (rm) {
        const rmProjs = getRmProjects(rm);
        kpiMetrics = kpiMetrics.filter(m => m.project && rmProjs.has(m.project.toLowerCase()));
    }
    const metCount = kpiMetrics.filter(m => m.status === 'Met').length;
    const kpiMetRate = kpiMetrics.length ? Math.round((metCount / kpiMetrics.length) * 100) : null;

    // 4. Filter Avg FTR Rating
    let qa = state.data.ftr?.qaMetrics || [];
    if (proj) qa = qa.filter(q => (q.project || '').toLowerCase().includes(proj.toLowerCase()));
    if (month) qa = qa.filter(q => monthLabel(q.qaDate) === month);
    if (client) qa = qa.filter(q => {
        const lp = (q.project || '').toLowerCase();
        return lp.includes(client.toLowerCase()) || client.toLowerCase().includes(lp.split(' ')[0]);
    });
    if (rm) {
        const rmProjs = getRmProjects(rm);
        qa = qa.filter(q => q.project && rmProjs.has(q.project.toLowerCase()));
    }
    const ftrAvgRating = qa.length ? Math.round(qa.reduce((s, q) => s + (q.passRate || 0), 0) / qa.length) : null;

    // 5. Filter On Leave Today
    let onTodayLeaves = state.data.leave?.currentMonth?.onLeaveToday || [];
    if (rm || proj || client) {
        const teamNames = new Set(members.map(m => m.name.toLowerCase()));
        onTodayLeaves = onTodayLeaves.filter(l => teamNames.has(l.name.toLowerCase()));
    }
    const onLeaveToday = onTodayLeaves.length;

    // 6. Filter Active Risks
    let risks = (state.data.gov?.risks || []).filter(r => (r.status || '').toLowerCase() === 'ongoing');
    if (proj) risks = risks.filter(r => r.project === proj);
    if (month) risks = risks.filter(r => r.month === month);
    if (client) {
        const clientProjs = getClientProjects(client);
        risks = risks.filter(r => r.project && clientProjs.has(r.project.toLowerCase()));
    }
    if (rm) {
        const rmProjs = getRmProjects(rm);
        risks = risks.filter(r => r.project && rmProjs.has(r.project.toLowerCase()));
    }
    const activeRisks = risks.length;

    // 7. PoC Use Cases
    let pocs = state.data.poc?.pocs || [];
    if (rm) pocs = pocs.filter(p => p.manager === rm);
    if (client) {
        const pmSet = new Set(pmsForClient(client).map(n => n.toLowerCase()));
        pocs = pocs.filter(p => p.manager && pmSet.has(p.manager.toLowerCase()));
    }
    if (proj) {
        const pmSet = new Set(pmsForProject(proj).map(n => n.toLowerCase()));
        pocs = pocs.filter(p => p.manager && pmSet.has(p.manager.toLowerCase()));
    }
    const pocTotal = pocs.length;

    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('👥', teamSize, 'Team Size')}
            ${statCard('📂', activeProjects, 'Active Projects')}
            ${statCard('🎯', kpiMetRate != null ? kpiMetRate + '%' : '—', 'KPI Met Rate')}
            ${statCard('✅', ftrAvgRating != null ? ftrAvgRating + '%' : '—', 'Avg FTR Rating')}
            ${statCard('🏖️', onLeaveToday, 'On Leave Today', onLeaveToday > 0 ? 'kpi-card-alert' : '')}
            ${statCard('⚠️', activeRisks, 'Active Risks', activeRisks > 0 ? 'kpi-card-alert' : '')}
            ${statCard('🧪', pocTotal, 'PoC Use Cases')}
        </div>
        <div class="grid-2">
            <div class="m-card">${section('🏥 Project Health Snapshot')}<div id="ovHealth"></div></div>
            <div class="m-card">${section('⚠️ Active Risk Alerts')}<div id="ovRisks"></div></div>
        </div>
        <div class="grid-2">
            <div class="m-card">${section('🌟 Recent Highlights')}<div id="ovHL"></div></div>
            <div class="m-card">${section('⬇️ Recent Lowlights')}<div id="ovLL"></div></div>
        </div>`;

    const pf = { project: proj, month: month, rm: rm, client: client };
    widgetProjectHealth($('#ovHealth'), pf, 6);
    widgetRiskList($('#ovRisks'), pf, 6);
    widgetHighlights($('#ovHL'), 6, pf);
    widgetLowlights($('#ovLL'), 6, pf);
}

// ══════════════════════════════════════════════════════════════════
//  TAB: TEAM DETAILS
// ══════════════════════════════════════════════════════════════════

function tabTeam(el) {
    const t = state.data.team || {};
    let members = t.activeMembers || [];
    const proj = getFilter('project'), rm = getFilter('rm'), client = getFilter('client');
    if (rm) members = members.filter(m => rmMatches(m.manager, rm));
    if (client) members = membersForPMs(members, pmsForClient(client));
    if (proj) members = membersForPMs(members, pmsForProject(proj));

    const desigDist = {};
    members.forEach(m => { desigDist[m.designation || 'Unknown'] = (desigDist[m.designation || 'Unknown'] || 0) + 1; });
    const officeDist = {};
    members.forEach(m => {
        const o = m.location || 'Unmapped';
        officeDist[o] = (officeDist[o] || 0) + 1;
    });

    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('👥', members.length, 'Members Shown')}
            ${statCard('🏢', Object.keys(officeDist).length, 'Offices')}
            ${statCard('👤', Object.keys(desigDist).length, 'Distinct Designations')}
        </div>
        <div class="grid-2">
            <div class="m-card">${section('🏢 Members by Office')}${chartCanvas('chTeamOffice')}</div>
            <div class="m-card">${section('👤 Members by Designation')}${chartCanvas('chTeamRole')}</div>
        </div>
        <div class="m-card">
            ${section('👥 Team Roster', 'Office mapping per member')}
            <div id="teamRoster"></div>
        </div>`;

    $('#teamRoster').innerHTML = dataTable(
        ['Name', 'Role', 'Designation', 'Manager (RM)', 'Emp ID', 'Office / Location', 'Email', 'Date of Joining'],
        members.map(m => [
            `<strong>${esc(m.name)}</strong>`, esc(m.role), esc(m.designation), esc(m.manager || ''),
            esc(m.empId), badge(m.location || 'Unmapped', m.location ? 'b-blue' : 'b-grey'),
            esc(m.email), esc(m.doj),
        ]),
        { empty: 'No team members match the filters.' }
    );

    barChart('chTeamOffice', Object.keys(officeDist), Object.values(officeDist), 'Members', '#06b6d4');
    barChart('chTeamRole', Object.keys(desigDist), Object.values(desigDist), 'Members', '#a855f7');
}

// ══════════════════════════════════════════════════════════════════
//  TAB: PROJECTS
// ══════════════════════════════════════════════════════════════════

function tabProjects(el) {
    const monthF = getFilter('month'), clientF = getFilter('client');
    el.innerHTML = `
        <div class="m-card">${section('🏥 Project Health Status', 'Overall health of all projects')}<div id="pjHealth"></div></div>`;

    const pf = {};
    if (monthF) pf.month = monthF;
    if (clientF) pf.client = clientF;

    widgetProjectHealth($('#pjHealth'), pf);
}

// ══════════════════════════════════════════════════════════════════
//  TAB: SOW & PO
// ══════════════════════════════════════════════════════════════════

function tabSow(el) {
    let projects = state.data.sow?.projects || [];
    const proj = getFilter('project'), startD = getFilter('startDate'), endD = getFilter('endDate');
    const client = getFilter('client');
    if (client) projects = projects.filter(p => p.client === client);
    if (proj) projects = projects.filter(p => p.projectName === proj);
    if (startD) projects = projects.filter(p => p.startDate >= startD);
    if (endD) projects = projects.filter(p => p.endDate <= endD);

    const statusCount = {};
    projects.forEach(p => {
        const k = `SOW ${p.sowStatus || '—'}`;
        statusCount[k] = (statusCount[k] || 0) + 1;
    });

    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('📂', projects.length, 'Projects')}
            ${statCard('📄', projects.filter(p => /received/i.test(p.sowStatus)).length, 'SOW Received')}
            ${statCard('🧾', projects.filter(p => /received/i.test(p.poStatus)).length, 'PO Received')}
        </div>
        <div class="m-card">
            ${section('💰 SOW & PO Financial Details')}
            <div id="sowFin"></div>
        </div>
        <div class="grid-2">
            <div class="m-card">${section('📋 SOW / PO Status')}${chartCanvas('chSowStatus')}</div>
            <div class="m-card">${section('🧾 Invoice Status', 'Derived from PO status')}<div id="sowInvoice"></div></div>
        </div>`;

    $('#sowFin').innerHTML = dataTable(
        ['Project', 'Client', 'PM', 'SOW Value', 'DE Revenue Share', 'Start', 'End', 'SOW Status', 'PO Status', 'Project Status'],
        projects.map(p => [
            `<strong>${esc(p.projectName)}</strong>`, esc(p.client), esc(p.pm),
            p.sowValue ? '$' + Number(p.sowValue).toLocaleString() : '<span class="muted">—</span>',
            p.deRevenueShare ? '$' + Number(p.deRevenueShare).toLocaleString() : '<span class="muted">—</span>',
            esc(p.startDate), esc(p.endDate),
            statusBadge(p.sowStatus), statusBadge(p.poStatus), statusBadge(p.projectStatus),
        ]),
        { empty: 'No SOW projects match the filter.' }
    );

    $('#sowInvoice').innerHTML = dataTable(
        ['Project', 'Client', 'PO Status', 'Invoice Status'],
        projects.map(p => {
            const inv = /received/i.test(p.poStatus) ? badge('Invoice Raised', 'b-green')
                : /ytr/i.test(p.poStatus) ? badge('Yet To Raise', 'b-amber')
                : badge('Pending', 'b-red');
            return [esc(p.projectName), esc(p.client), statusBadge(p.poStatus), inv];
        }),
        { empty: 'No data.', note: 'Invoice status is derived from PO status (source file has no explicit invoice column).' }
    );

    doughnutChart('chSowStatus', Object.keys(statusCount), Object.values(statusCount));
}

// ══════════════════════════════════════════════════════════════════
//  TAB: KPIs
// ══════════════════════════════════════════════════════════════════

function tabKpis(el) {
    const kpi = state.data.kpi || {};
    const proj = getFilter('project'), kpiMetric = getFilter('kpiMetric'), client = getFilter('client');

    let metrics = kpi.metrics || [];
    if (client) metrics = metrics.filter(m => m.account === client);
    if (proj) metrics = metrics.filter(m => (m.project || '').toLowerCase().includes(proj.toLowerCase()));
    if (kpiMetric) {
        const search = kpiMetric.split('(')[0].trim().toLowerCase();
        metrics = metrics.filter(m => (m.name || m.metric || '').toLowerCase().includes(search));
    }

    const totalMetrics = metrics.length;
    const metCount = metrics.filter(m => {
        const s = (m.status || '').toLowerCase();
        if (s.includes('not') || s.includes('miss')) return false;
        return s === 'met' || s.includes('achieved') || s.includes('green') ||
            (m.actual >= m.target && m.target > 0);
    }).length;
    const metRate = totalMetrics > 0 ? Math.round((metCount / totalMetrics) * 100) : 0;

    let scRows = kpi.scorecardProjects || [];
    if (client) scRows = scRows.filter(r => r.client === client);
    if (proj) scRows = scRows.filter(r => r.project === proj);

    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('🎯', totalMetrics > 0 ? metRate + '%' : '—', 'KPI Met Rate')}
            ${statCard('📊', totalMetrics || 0, 'Total Metrics')}
            ${statCard('✅', metCount || 0, 'Targets Met')}
            ${statCard('📂', scRows.length, 'Scorecard Entries')}
        </div>
        <div class="m-card">${section('📊 Project-wise Monthly KPI Breakdown', 'Timeliness, Utilization, Quality, CSAT, Client Training')}<div id="kpiBreakdown"></div></div>
        <div class="m-card">${section('🎯 KPI Scorecard', 'Target vs Actual')}<div id="kpiScore"></div></div>
        <div class="m-card">${section('🏷️ KPI — Client-wise & Monthly View', 'Avg score across all projects')}<div id="kpiPivot"></div></div>
        <div class="grid-2">
            <div class="m-card">${section('📈 FTE Utilization — Client-wise & Monthly')}<div id="kpiUtil"></div></div>
            <div class="m-card">${section('⚠️ Project Risks')}<div id="kpiRisks"></div></div>
        </div>
        <div class="m-card">${section('✅ FTR Report', 'Client-wise & monthly')}<div id="kpiFtr"></div></div>
        <div class="m-card">${section('🏥 Project Health Status')}<div id="kpiHealth"></div></div>`;

    // Build a mapping of which columns to show based on kpiMetric filter
    const kpiCols = [
        { key: 'timeliness', label: 'Timeliness (SLA%)', match: 'Timeliness' },
        { key: 'utilization', label: 'Utilization (Effort %)', match: 'Utilization' },
        { key: 'quality', label: 'Quality (FTR %)', match: 'Quality' },
        { key: 'csatFrequency', label: 'CSAT (Frequency)', match: 'CSAT (Frequency)' },
        { key: 'csatRating', label: 'CSAT (Rating)', match: 'CSAT (Rating)' },
        { key: 'clientTraining', label: 'Client Training %', match: 'Client Training' },
    ];
    const visibleKpiCols = kpiMetric
        ? kpiCols.filter(c => c.match.toLowerCase().includes(kpiMetric.split('(')[0].trim().toLowerCase()) || kpiMetric.toLowerCase().includes(c.match.split('(')[0].trim().toLowerCase()))
        : kpiCols;

    const breakdownHeaders = ['Client', 'Project', 'PM', 'Month', ...visibleKpiCols.map(c => c.label)];
    const breakdownRows = scRows.map(r => [
        esc(r.client), `<strong>${esc(r.project)}</strong>`, esc(r.pm), esc(r.month),
        ...visibleKpiCols.map(c => {
            const val = r[c.key];
            if (!val) return '<span class="muted">—</span>';
            // Color code percentage values
            if (val.endsWith('%')) {
                const n = parseFloat(val);
                const cls = n >= 95 ? 'b-green' : n >= 85 ? 'b-amber' : n > 0 ? 'b-red' : 'b-grey';
                return `<span class="badge ${cls}">${esc(val)}</span>`;
            }
            return esc(val);
        }),
    ]);
    $('#kpiBreakdown').innerHTML = dataTable(breakdownHeaders, breakdownRows, {
        empty: 'No KPI scorecard data available. Upload a KPI file with a "Web Analytics KPI Summary" sheet.',
        note: 'Data from Web Analytics KPI Summary sheet.',
    });

    widgetKpiScorecard($('#kpiScore'));
    widgetKpiPivot($('#kpiPivot'));
    widgetKpiUtilPivot($('#kpiUtil'));
    const pf = { project: proj, client: client };
    widgetRiskList($('#kpiRisks'), pf, 10);
    widgetFtrReport($('#kpiFtr'), pf);
    widgetProjectHealth($('#kpiHealth'), pf);
}

// ══════════════════════════════════════════════════════════════════
//  TAB: FTR
// ══════════════════════════════════════════════════════════════════

function tabFtr(el) {
    const ftr = state.data.ftr || {};
    const client = getFilter('client'), month = getFilter('month'), proj = getFilter('project');

    let qa = ftr.qaMetrics || [];
    if (month) qa = qa.filter(q => monthLabel(q.qaDate) === month);
    if (client) {
        const clientProjs = getClientProjects(client);
        qa = qa.filter(q => {
            if (!q.project) return false;
            const lp = q.project.toLowerCase();
            if (clientProjs.has(lp)) return true;
            if (lp.includes(client.toLowerCase()) || client.toLowerCase().includes(lp.split(' ')[0])) return true;
            return false;
        });
    }
    if (proj) qa = qa.filter(q => (q.project || '').toLowerCase().includes(proj.toLowerCase()));

    const ftrAvgRating = qa.length ? Math.round(qa.reduce((s, q) => s + (q.passRate || 0), 0) / qa.length) : null;
    const totalQaRecords = qa.length;

    let accounts = ftr.accounts || [];
    if (client) accounts = accounts.filter(a => a.account === client);
    if (proj) {
        const sowProjects = state.data.sow?.projects || [];
        const matchedProj = sowProjects.find(p => p.projectName && p.projectName.toLowerCase() === proj.toLowerCase());
        if (matchedProj && matchedProj.client) {
            accounts = accounts.filter(a => a.account && a.account.toLowerCase() === matchedProj.client.toLowerCase());
        }
    }
    const accountsCount = accounts.length;

    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('✅', ftrAvgRating != null ? ftrAvgRating + '%' : '—', 'Avg FTR Rating')}
            ${statCard('📂', accountsCount, 'Accounts')}
            ${statCard('🔍', totalQaRecords, 'QA Records')}
        </div>
        <div class="m-card">${section('✅ FTR Report — Client-wise & Monthly')}<div id="ftrReport"></div></div>
        <div class="m-card">${section('🏷️ FTR by Account (Client-wise)')}<div id="ftrAccounts"></div></div>
        <div class="m-card">${section('🔍 QA Metrics Detail')}<div id="ftrQa"></div></div>`;

    widgetFtrReport($('#ftrReport'));

    $('#ftrAccounts').innerHTML = dataTable(
        ['Account', 'QA Tracking', 'Expected Projects', 'Web Team', 'Avg Rating', 'FTR Pass'],
        accounts.map(a => [
            `<strong>${esc(a.account)}</strong>`, esc(a.qaTracking), esc(a.expectedProjects),
            esc(a.webTeam), ratingCell(a.avgRating), esc(a.ftrPass),
        ]),
        { empty: 'No accounts match the filter.' }
    );

    $('#ftrQa').innerHTML = dataTable(
        ['Project', 'QA Date', 'Month', 'Total Tags', 'Failed', 'Pass Rate'],
        qa.map(q => [
            esc(q.project), esc(q.qaDate), esc(monthLabel(q.qaDate)),
            q.total, q.failed, ratingCell(q.passRate),
        ]),
        { empty: 'No QA records match the filters.' }
    );
}

// ══════════════════════════════════════════════════════════════════
//  TAB: LEAVE
// ══════════════════════════════════════════════════════════════════

function tabLeave(el) {
    const leave = state.data.leave || {};
    const cm = leave.currentMonth || {};
    const rm = getFilter('rm'), proj = getFilter('project'), client = getFilter('client');

    // Get team members filtered by RM, client and project for filtering leave data
    let teamMembers = state.data.team?.activeMembers || [];
    if (rm) teamMembers = teamMembers.filter(m => rmMatches(m.manager, rm));
    if (client) teamMembers = membersForPMs(teamMembers, pmsForClient(client));
    if (proj) teamMembers = membersForPMs(teamMembers, pmsForProject(proj));
    const teamNameSet = (rm || proj || client) ? new Set(teamMembers.map(m => m.name.toLowerCase())) : null;

    // Build a name→manager lookup from all team members
    const allMembers = state.data.team?.activeMembers || [];
    const managerLookup = {};
    allMembers.forEach(m => { managerLookup[m.name.toLowerCase()] = canonicalRM(m.manager) || '—'; });

    let onToday = cm.onLeaveToday || [];
    if (teamNameSet) onToday = onToday.filter(p => teamNameSet.has(p.name.toLowerCase()));

    let byPerson = Object.entries(cm.byPerson || {})
        .filter(([, c]) => c > 0).sort((a, b) => b[1] - a[1]);
    if (teamNameSet) byPerson = byPerson.filter(([n]) => teamNameSet.has(n.toLowerCase()));

    const totalLeavesThisMonth = byPerson.reduce((sum, [, count]) => sum + count, 0);
    const uniqueRMs = [...new Set(byPerson.map(([n]) => managerLookup[n.toLowerCase()] || '—').filter(v => v !== '—'))];

    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('🏖️', onToday.length, 'On Leave Today', onToday.length > 0 ? 'kpi-card-alert' : '')}
            ${statCard('📅', totalLeavesThisMonth, 'Leaves This Month')}
            ${statCard('👔', rm ? '1' : uniqueRMs.length, rm ? esc(rm) : 'RMs with Leave')}
            ${statCard('🗓️', (leave.months || []).length, 'Months Tracked')}
        </div>
        <div class="grid-2">
            <div class="m-card">${section('🏖️ On Leave Today', rm ? 'Reportees of ' + rm : '')}<div id="lvToday"></div></div>
            <div class="m-card">${section('🏆 Top Leave Takers — This Month', rm ? 'Under RM: ' + rm : '')}<div id="lvTop"></div></div>
        </div>
        <div class="m-card">${section('👔 Leave by Reporting Manager', 'Reportee-wise leave breakdown per RM')}<div id="lvByRM"></div></div>
        <div class="m-card">${section('🗓️ Leave Tracker — Months Available')}<div id="lvMonths"></div></div>`;

    $('#lvToday').innerHTML = dataTable(
        ['Member', 'Manager (RM)', 'Leave Type'],
        onToday.map(p => [
            `<strong>${esc(p.name)}</strong>`,
            esc(managerLookup[p.name.toLowerCase()] || '—'),
            esc(p.type),
        ]),
        { empty: 'Nobody is on leave today. ✅' }
    );

    $('#lvTop').innerHTML = dataTable(
        ['Member', 'Manager (RM)', 'Leave Days'],
        byPerson.slice(0, 20).map(([n, c]) => [
            `<strong>${esc(n)}</strong>`,
            esc(managerLookup[n.toLowerCase()] || '—'),
            badge(c + ' day(s)', 'b-amber'),
        ]),
        { empty: 'No leave records for the current month.' }
    );

    // Group leave by RM
    const leaveByRM = {};
    byPerson.forEach(([n, c]) => {
        const mgr = managerLookup[n.toLowerCase()] || 'Unknown';
        if (!leaveByRM[mgr]) leaveByRM[mgr] = [];
        leaveByRM[mgr].push({ name: n, days: c });
    });
    const rmRows = [];
    Object.entries(leaveByRM).sort((a, b) => {
        const totalA = a[1].reduce((s, r) => s + r.days, 0);
        const totalB = b[1].reduce((s, r) => s + r.days, 0);
        return totalB - totalA;
    }).forEach(([mgr, reportees]) => {
        const totalDays = reportees.reduce((s, r) => s + r.days, 0);
        reportees.forEach((r, idx) => {
            rmRows.push([
                idx === 0 ? `<strong>${esc(mgr)}</strong>` : '',
                `<strong>${esc(r.name)}</strong>`,
                badge(r.days + ' day(s)', 'b-amber'),
                idx === 0 ? badge(totalDays + ' total day(s)', 'b-blue') : '',
            ]);
        });
    });
    $('#lvByRM').innerHTML = dataTable(
        ['Reporting Manager', 'Reportee', 'Leave Days', 'RM Total'],
        rmRows,
        { empty: 'No leave data available for RM breakdown.' }
    );

    $('#lvMonths').innerHTML = dataTable(
        ['Month Sheet', 'Year'],
        (leave.months || []).map(m => [esc(m.sheet), esc(m.year)]),
        { empty: 'No month sheets found.', note: 'Leave file is tracked per monthly sheet; per-person detail is available for the current month above.' }
    );
}

// ══════════════════════════════════════════════════════════════════
//  TAB: ACTIVE RISKS
// ══════════════════════════════════════════════════════════════════

function tabRisks(el) {
    const risks = state.data.gov?.risks || [];
    const client = getFilter('client'), proj = getFilter('project'), month = getFilter('month');

    let pv = risks;
    if (proj) pv = pv.filter(r => r.project === proj);
    if (month) pv = pv.filter(r => r.month === month);
    if (client) {
        const clientProjs = getClientProjects(client);
        pv = pv.filter(r => r.project && clientProjs.has(r.project.toLowerCase()));
    }

    const activeCount = pv.filter(r => (r.status || '').toLowerCase() === 'ongoing').length;
    const totalCount = pv.length;

    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('⚠️', activeCount, 'Active Risks', activeCount > 0 ? 'kpi-card-alert' : '')}
            ${statCard('📋', totalCount, 'Total Risk Records')}
        </div>
        <div class="m-card">${section('📊 Risks — Project-wise & Monthly', 'Count of risk records')}<div id="rkPivot"></div></div>
        <div class="m-card">${section('⚠️ Risk Details')}<div id="rkList"></div></div>`;

    const rowKeys = uniqueSorted(pv.map(r => r.project || 'Unknown'));
    const colKeys = uniqueSorted(pv.map(r => r.month || 'Unknown'));
    $('#rkPivot').innerHTML = pivotTable(rowKeys, colKeys, (rk, ck) => {
        const n = pv.filter(r => (r.project || 'Unknown') === rk && (r.month || 'Unknown') === ck).length;
        return n ? `<span class="badge b-red">${n}</span>` : '';
    }, { corner: 'Project', empty: 'No risks match the filters.' });

    $('#rkList').innerHTML = dataTable(
        ['Project', 'Month', 'PM', 'Risk / Description', 'Impact', 'Mitigation', 'Status'],
        pv.map(r => [
            `<strong>${esc(r.project)}</strong>`, esc(r.month), esc(r.pm),
            esc(r.description || r.risk || ''), impactBadge(r.impact),
            esc(r.mitigation), statusBadge(r.status),
        ]),
        { empty: 'No risks match the filters.' }
    );
}

// ══════════════════════════════════════════════════════════════════
//  TAB: POV / POC
// ══════════════════════════════════════════════════════════════════

function tabPoc(el) {
    const poc = state.data.poc || {};
    let pocs = poc.pocs || [];
    const spoc = getFilter('spoc');
    if (spoc) pocs = pocs.filter(p => p.spoc === spoc);

    // Compute dynamic metrics
    const totalCount = pocs.length;
    const byStatus = {};
    pocs.forEach(p => {
        const s = p.status || 'Unknown';
        byStatus[s] = (byStatus[s] || 0) + 1;
    });

    let completedCount = 0;
    let inProgressCount = 0;
    Object.entries(byStatus).forEach(([k, v]) => {
        const lowerK = k.toLowerCase();
        if (lowerK.includes('complete') || lowerK === 'done') {
            completedCount += v;
        }
        if (['progress', 'wip', 'ongoing'].some(w => lowerK.includes(w))) {
            inProgressCount += v;
        }
    });

    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('🧪', totalCount, 'Total Use Cases')}
            ${statCard('✅', completedCount, 'Completed')}
            ${statCard('🔧', inProgressCount, 'In Progress')}
            ${statCard('🤖', (poc.aiUsecases || []).length, 'AI Use Cases')}
        </div>
        <div class="grid-2">
            <div class="m-card">${section('🚦 PoC Status Breakdown')}${chartCanvas('chPocStatus')}</div>
            <div class="m-card">${section('👔 PoCs by Reporting Manager')}${chartCanvas('chPocMgr')}</div>
        </div>
        <div class="m-card">${section('🧪 PoC Use Case Details')}<div id="pocTable"></div></div>
        <div class="m-card">${section('🤖 AI Use Case Backlog')}<div id="aiTable"></div></div>`;

    $('#pocTable').innerHTML = dataTable(
        ['#', 'PoC Title', 'Reporting Manager', 'SPOC', 'Team', 'Status', 'Last Worked', 'Comments / Links'],
        pocs.map(p => [
            esc(p.sno), `<strong>${esc(p.title)}</strong>`, esc(p.manager), esc(p.spoc),
            esc(p.team), statusBadge(p.status), esc(p.lastWorked), esc(p.comments),
        ]),
        { empty: 'No PoC use cases match the filters.' }
    );

    $('#aiTable').innerHTML = dataTable(
        ['Idea ID', 'Business Function', 'Pain Point', 'Proposed AI Solution', 'Priority', 'Impact', 'Idea By'],
        (poc.aiUsecases || []).map(a => [
            esc(a.ideaId), esc(a.businessFunction), esc(a.painPoint), esc(a.solution),
            priorityBadge(a.priority), priorityBadge(a.impact), esc(a.ideaBy),
        ]),
        { empty: 'No AI use cases recorded.' }
    );

    doughnutChart('chPocStatus', Object.keys(byStatus), Object.values(byStatus));
    const mgrCount = {};
    pocs.forEach(p => { mgrCount[p.manager || 'Unknown'] = (mgrCount[p.manager || 'Unknown'] || 0) + 1; });
    barChart('chPocMgr', Object.keys(mgrCount), Object.values(mgrCount), 'PoCs', '#06b6d4');
}

// ══════════════════════════════════════════════════════════════════
//  REUSABLE WIDGETS
// ══════════════════════════════════════════════════════════════════

function widgetProjectHealth(el, filter, limit) {
    let rows = state.data.health || [];
    if (filter) {
        if (filter.project) rows = rows.filter(p => p.projectName === filter.project);
        if (filter.client) rows = rows.filter(p => p.client === filter.client);
        if (filter.rm) {
            const rmProjs = getRmProjects(filter.rm);
            rows = rows.filter(p => p.projectName && rmProjs.has(p.projectName.toLowerCase()));
        }
    }
    if (limit) rows = rows.slice(0, limit);
    el.innerHTML = dataTable(
        ['Project', 'Client', 'PM', 'Status', 'SOW', 'PO', 'Health', 'Notes'],
        rows.map(p => [
            `<strong>${esc(p.projectName)}</strong>`, esc(p.client), esc(p.pm),
            esc(p.status), statusBadge(p.sowStatus), statusBadge(p.poStatus),
            healthBadge(p.health), `<span class="muted-sm">${esc(p.reasons)}</span>`,
        ]),
        { empty: 'No project health data.' }
    );
}

function widgetRiskList(el, filter, limit) {
    let risks = (state.data.gov?.risks || []).filter(r => (r.status || '').toLowerCase() === 'ongoing');
    if (filter) {
        if (filter.project) risks = risks.filter(r => r.project === filter.project);
        if (filter.month) risks = risks.filter(r => r.month === filter.month);
        if (filter.client) {
            const clientProjs = getClientProjects(filter.client);
            risks = risks.filter(r => r.project && clientProjs.has(r.project.toLowerCase()));
        }
        if (filter.rm) {
            const rmProjs = getRmProjects(filter.rm);
            risks = risks.filter(r => r.project && rmProjs.has(r.project.toLowerCase()));
        }
    }
    if (limit) risks = risks.slice(0, limit);
    if (!risks.length) { el.innerHTML = '<div class="empty-note">No active risks. ✅</div>'; return; }
    el.innerHTML = risks.map(r => `
        <div class="risk-item">
            <div class="risk-project">${esc(r.project)} <span class="muted">— ${esc(r.pm)}</span>
                ${impactBadge(r.impact)}</div>
            <div class="risk-text">${esc(r.description || r.risk || '')}</div>
            ${r.mitigation ? `<div class="risk-mit">🛡️ ${esc(r.mitigation)}</div>` : ''}
        </div>`).join('');
}

function widgetHighlights(el, limit, filter) {
    let items = (state.data.gov?.highlights || []).filter(h => h.highlight);
    if (filter) {
        if (typeof filter === 'string') {
            items = items.filter(h => h.project === filter);
        } else {
            if (filter.project) items = items.filter(h => h.project === filter.project);
            if (filter.month) items = items.filter(h => h.month === filter.month);
            if (filter.client) {
                const clientProjs = getClientProjects(filter.client);
                items = items.filter(h => h.project && clientProjs.has(h.project.toLowerCase()));
            }
            if (filter.rm) {
                const rmProjs = getRmProjects(filter.rm);
                items = items.filter(h => h.project && rmProjs.has(h.project.toLowerCase()));
            }
        }
    }
    items = items.slice(0, limit || 10);
    el.innerHTML = items.length ? items.map(h => `
        <div class="hl-item">
            <div class="hl-project">${esc(h.project)} <span class="muted">— ${esc(h.month)}</span></div>
            <div class="hl-text">${esc(h.highlight)}</div>
        </div>`).join('') : '<div class="empty-note">No highlights reported.</div>';
}

function widgetLowlights(el, limit, filter) {
    let items = (state.data.gov?.lowlights || []).filter(h => h.lowlight);
    if (filter) {
        if (typeof filter === 'string') {
            items = items.filter(h => h.project === filter);
        } else {
            if (filter.project) items = items.filter(h => h.project === filter.project);
            if (filter.month) items = items.filter(h => h.month === filter.month);
            if (filter.client) {
                const clientProjs = getClientProjects(filter.client);
                items = items.filter(h => h.project && clientProjs.has(h.project.toLowerCase()));
            }
            if (filter.rm) {
                const rmProjs = getRmProjects(filter.rm);
                items = items.filter(h => h.project && rmProjs.has(h.project.toLowerCase()));
            }
        }
    }
    items = items.slice(0, limit || 10);
    el.innerHTML = items.length ? items.map(h => `
        <div class="hl-item ll-item">
            <div class="hl-project">${esc(h.project)} <span class="muted">— ${esc(h.month)}</span></div>
            <div class="hl-text">${esc(h.lowlight)}</div>
        </div>`).join('') : '<div class="empty-note">No lowlights reported.</div>';
}

function widgetUtilTrendChart(canvasId) {
    const gov = state.data.gov || {};
    const fteData = gov.fte || [];
    const client = getFilter('client'), project = getFilter('project');
    
    let filteredFte = fteData;
    if (client) {
        filteredFte = fteData.filter(entry => entry.account && entry.account.toLowerCase() === client.toLowerCase());
    } else if (project) {
        const sowProjects = state.data.sow?.projects || [];
        const matchedProj = sowProjects.find(p => p.projectName && p.projectName.toLowerCase() === project.toLowerCase());
        if (matchedProj && matchedProj.client) {
            filteredFte = fteData.filter(entry => entry.account && entry.account.toLowerCase() === matchedProj.client.toLowerCase());
        } else {
            filteredFte = fteData.filter(entry => entry.account && (
                project.toLowerCase().includes(entry.account.toLowerCase()) ||
                entry.account.toLowerCase().includes(project.toLowerCase())
            ));
        }
    }
    
    const fteByMonth = {};
    filteredFte.forEach(entry => {
        Object.entries(entry.monthlyFTE || {}).forEach(([month, val]) => {
            fteByMonth[month] = (fteByMonth[month] || 0) + val;
        });
    });
    
    const trend = Object.entries(fteByMonth)
        .map(([month, totalFTE]) => ({ month, totalFTE }))
        .sort((a, b) => a.month.localeCompare(b.month));

    if (!trend.length) {
        const c = document.getElementById(canvasId);
        if (c) c.replaceWith(Object.assign(document.createElement('div'),
            { className: 'empty-note', textContent: 'No FTE trend data.' }));
        return;
    }
    lineChart(canvasId, trend.map(t => t.month), trend.map(t => t.totalFTE), 'Total FTE');
}

function widgetTopLeaveTakers(el, filter) {
    const leave = state.data.leave || {};
    const cm = leave.currentMonth || {};
    let teamMembers = state.data.team?.activeMembers || [];
    if (filter) {
        if (filter.rm) teamMembers = teamMembers.filter(m => rmMatches(m.manager, filter.rm));
        if (filter.client) teamMembers = membersForPMs(teamMembers, pmsForClient(filter.client));
        if (filter.project) teamMembers = membersForPMs(teamMembers, pmsForProject(filter.project));
    }
    const teamNameSet = (filter && (filter.rm || filter.project || filter.client)) ? new Set(teamMembers.map(m => m.name.toLowerCase())) : null;
    
    let byPerson = Object.entries(cm.byPerson || {})
        .filter(([, c]) => c > 0).sort((a, b) => b[1] - a[1]);
    if (teamNameSet) byPerson = byPerson.filter(([n]) => teamNameSet.has(n.toLowerCase()));
    
    const sliced = byPerson.slice(0, 10);
    el.innerHTML = dataTable(
        ['Member', 'Leave Days'],
        sliced.map(([n, c]) => [`<strong>${esc(n)}</strong>`, badge(c + ' day(s)', 'b-amber')]),
        { empty: 'No leave records this month.' }
    );
}

function widgetLeaveImpact(el, filter) {
    const leave = state.data.leave || {};
    const cm = leave.currentMonth || {};
    let teamMembers = state.data.team?.activeMembers || [];
    if (filter) {
        if (filter.rm) teamMembers = teamMembers.filter(m => rmMatches(m.manager, filter.rm));
        if (filter.client) teamMembers = membersForPMs(teamMembers, pmsForClient(filter.client));
        if (filter.project) teamMembers = membersForPMs(teamMembers, pmsForProject(filter.project));
    }
    const teamNameSet = (filter && (filter.rm || filter.project || filter.client)) ? new Set(teamMembers.map(m => m.name.toLowerCase())) : null;
    
    let onToday = cm.onLeaveToday || [];
    if (teamNameSet) onToday = onToday.filter(p => teamNameSet.has(p.name.toLowerCase()));
    
    let byPerson = Object.entries(cm.byPerson || {})
        .filter(([, c]) => c > 0).sort((a, b) => b[1] - a[1]);
    if (teamNameSet) byPerson = byPerson.filter(([n]) => teamNameSet.has(n.toLowerCase()));
    const totalLeavesThisMonth = byPerson.reduce((sum, [, count]) => sum + count, 0);

    el.innerHTML = `
        <div class="impact-row">
            <div class="impact-stat"><span class="impact-num">${onToday.length}</span> on leave today</div>
            <div class="impact-stat"><span class="impact-num">${totalLeavesThisMonth}</span> leaves this month</div>
        </div>
        ${onToday.length ? dataTable(['Member', 'Type'],
        onToday.map(p => [esc(p.name), esc(p.type)])) : '<div class="empty-note">Nobody on leave today. ✅</div>'}`;
}

function widgetKpiScorecard(el) {
    let metrics = state.data.kpi?.metrics || [];
    const client = getFilter('client'), month = getFilter('month'), proj = getFilter('project');
    if (client) metrics = metrics.filter(m => m.account === client);
    if (month) metrics = metrics.filter(m => m.month === month);
    if (proj) metrics = metrics.filter(m => (m.project || '').toLowerCase().includes(proj.toLowerCase()));
    el.innerHTML = dataTable(
        ['Client', 'Project', 'Metric', 'Month', 'Target', 'Actual', 'Status'],
        metrics.map(m => [
            esc(m.account), esc(m.project), esc(m.name), esc(m.month),
            esc(m.targetRaw), esc(m.actualRaw),
            m.status === 'Met' ? badge('Met', 'b-green') : badge('Not Met', 'b-red'),
        ]),
        { empty: 'No KPI metrics match the filters.' }
    );
}

function widgetKpiPivot(el) {
    let metrics = state.data.kpi?.metrics || [];
    const client = getFilter('client'), month = getFilter('month'), proj = getFilter('project');
    if (client) metrics = metrics.filter(m => m.account === client);
    if (month) metrics = metrics.filter(m => m.month === month);
    if (proj) metrics = metrics.filter(m => (m.project || '').toLowerCase().includes(proj.toLowerCase()));
    const rowKeys = uniqueSorted(metrics.map(m => m.account));
    const colKeys = uniqueSorted(metrics.map(m => m.month));
    el.innerHTML = pivotTable(rowKeys, colKeys, (rk, ck) => {
        const cell = metrics.filter(m => m.account === rk && m.month === ck);
        if (!cell.length) return '';
        const avg = cell.reduce((s, m) => s + m.actual, 0) / cell.length;
        const pct = Math.round(avg * 100);
        const cls = pct >= 95 ? 'b-green' : pct >= 85 ? 'b-amber' : 'b-red';
        return `<span class="badge ${cls}">${pct}%</span>`;
    }, { corner: 'Client', empty: 'No KPI data for this view.',
         note: 'Cell = average KPI score across all of that client\'s projects for the month.' });
}

function widgetKpiUtilPivot(el) {
    let metrics = (state.data.kpi?.metrics || []).filter(m => /utiliz/i.test(m.name));
    const client = getFilter('client'), month = getFilter('month'), proj = getFilter('project');
    if (client) metrics = metrics.filter(m => m.account === client);
    if (month) metrics = metrics.filter(m => m.month === month);
    if (proj) metrics = metrics.filter(m => (m.project || '').toLowerCase().includes(proj.toLowerCase()));
    const rowKeys = uniqueSorted(metrics.map(m => m.account));
    const colKeys = uniqueSorted(metrics.map(m => m.month));
    el.innerHTML = pivotTable(rowKeys, colKeys, (rk, ck) => {
        const cell = metrics.filter(m => m.account === rk && m.month === ck);
        if (!cell.length) return '';
        const avg = cell.reduce((s, m) => s + m.actual, 0) / cell.length;
        return `<span class="badge b-blue">${Math.round(avg * 100)}%</span>`;
    }, { corner: 'Client', empty: 'No FTE utilization data.' });
}

function widgetFtrReport(el, filter) {
    const qa = state.data.ftr?.qaMetrics || [];
    const accounts = (state.data.ftr?.accounts || []).map(a => a.account);
    function clientOf(project) {
        const lp = (project || '').toLowerCase();
        const hit = accounts.find(a => a && lp.includes(a.toLowerCase()));
        return hit || 'Unmapped';
    }
    let recs = qa.map(q => ({ ...q, client: clientOf(q.project), month: monthLabel(q.qaDate) }));
    const client = filter?.client || getFilter('client');
    const month = filter?.month || getFilter('month');
    const project = filter?.project || getFilter('project');
    if (client) recs = recs.filter(r => r.client === client);
    if (month) recs = recs.filter(r => r.month === month);
    if (project) recs = recs.filter(r => (r.project || '').toLowerCase().includes(project.toLowerCase()));
    const rowKeys = uniqueSorted(recs.map(r => r.client));
    const colKeys = uniqueSorted(recs.map(r => r.month));
    el.innerHTML = pivotTable(rowKeys, colKeys, (rk, ck) => {
        const cell = recs.filter(r => r.client === rk && r.month === ck);
        if (!cell.length) return '';
        const avg = cell.reduce((s, r) => s + (r.passRate || 0), 0) / cell.length;
        const cls = avg >= 95 ? 'b-green' : avg >= 85 ? 'b-amber' : 'b-red';
        return `<span class="badge ${cls}">${avg.toFixed(0)}%</span>`;
    }, { corner: 'Client', empty: 'No FTR QA records for this view.',
         note: 'Cell = average QA pass rate for that client in the month.' });
}

// ══════════════════════════════════════════════════════════════════
//  BADGE / CELL HELPERS
// ══════════════════════════════════════════════════════════════════

function statusBadge(s) {
    s = (s || '').toString().trim();
    if (!s) return '<span class="muted">—</span>';
    const l = s.toLowerCase();
    let cls = 'b-grey';
    if (/(received|complete|done|met|green|active)/.test(l)) cls = 'b-green';
    else if (/(progress|ytr|yet to|pending|amber|hold|wip)/.test(l)) cls = 'b-amber';
    else if (/(not |delay|red|miss|escalat|blocked|ideation)/.test(l)) cls = 'b-red';
    return badge(s, cls);
}

function healthBadge(h) {
    const cls = h === 'Green' ? 'b-green' : h === 'Red' ? 'b-red' : 'b-amber';
    return badge(h || '—', cls);
}

function impactBadge(i) {
    i = (i || '').toString();
    const l = i.toLowerCase();
    const cls = l.includes('loss') ? 'b-red' : l.includes('delay') ? 'b-amber' : 'b-grey';
    return badge(i || 'Unknown', cls);
}

function priorityBadge(p) {
    p = (p || '').toString();
    const l = p.toLowerCase();
    const cls = l === 'high' ? 'b-red' : l === 'medium' ? 'b-amber' : l === 'low' ? 'b-green' : 'b-grey';
    return p ? badge(p, cls) : '<span class="muted">—</span>';
}

function ratingCell(v) {
    if (v == null || v === '') return '<span class="muted">—</span>';
    const n = Number(v);
    const cls = n >= 95 ? 'b-green' : n >= 85 ? 'b-amber' : 'b-red';
    return `<span class="badge ${cls}">${n.toFixed(1)}%</span>`;
}

// ══════════════════════════════════════════════════════════════════
//  CHART HELPERS
// ══════════════════════════════════════════════════════════════════

function barChart(id, labels, data, label, color) {
    makeChart(id, {
        type: 'bar',
        data: { labels, datasets: [{ label, data, backgroundColor: color, borderRadius: 4 }] },
        options: { responsive: true, plugins: { legend: { display: false } }, scales: CHART_AXES },
    });
}

function lineChart(id, labels, data, label) {
    makeChart(id, {
        type: 'line',
        data: {
            labels, datasets: [{
                label, data, borderColor: '#06b6d4', backgroundColor: 'rgba(6,182,212,.12)',
                fill: true, tension: .4, pointRadius: 4, pointBackgroundColor: '#06b6d4',
            }],
        },
        options: { responsive: true, plugins: { legend: CHART_LEGEND }, scales: CHART_AXES },
    });
}

function doughnutChart(id, labels, data) {
    const colors = ['#22c55e', '#f59e0b', '#06b6d4', '#a855f7', '#ef4444', '#ec4899', '#64748b'];
    makeChart(id, {
        type: 'doughnut',
        data: { labels, datasets: [{ data, backgroundColor: colors.slice(0, labels.length), borderWidth: 0 }] },
        options: {
            responsive: true,
            plugins: { legend: { position: 'bottom', labels: { color: '#94a3b8', font: { size: 10 }, padding: 10 } } },
        },
    });
}

// ══════════════════════════════════════════════════════════════════
//  UI BINDINGS — manager toolbar + data-source modal
// ══════════════════════════════════════════════════════════════════

function bindUI() {
    $('#mdRefresh')?.addEventListener('click', () => {
        const btn = $('#mdRefresh');
        btn.classList.add('refreshing');
        fetchAllExecData().then(() => setTimeout(() => btn.classList.remove('refreshing'), 800));
    });
    $('#mdPrint')?.addEventListener('click', () => window.print());
    $('#mdSources')?.addEventListener('click', () => $('#mdSourceModal')?.classList.remove('hidden'));
    $('#mdCloseModal')?.addEventListener('click', () => $('#mdSourceModal')?.classList.add('hidden'));
    $('#mdSourceModal')?.addEventListener('click', e => {
        if (e.target === $('#mdSourceModal')) $('#mdSourceModal').classList.add('hidden');
    });
    $$('#mdSourceModal .tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            $$('#mdSourceModal .tab-btn').forEach(b => b.classList.remove('active'));
            $$('#mdSourceModal .tab-content').forEach(c => c.classList.remove('active'));
            btn.classList.add('active');
            $(`#tab${capitalize(btn.dataset.tab)}`)?.classList.add('active');
        });
    });
    $('#mdUploadAll')?.addEventListener('click', uploadFiles);
    $('#mdConnectUrls')?.addEventListener('click', connectUrls);
}

function capitalize(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

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
            method: 'POST', headers: { Authorization: 'Bearer ' + getToken() }, body: formData,
        });
        const data = await res.json();
        if (data.success) {
            statusEl.className = 'source-status-box success';
            statusEl.textContent = '✓ ' + data.message;
            setTimeout(() => { fetchAllExecData(); $('#mdSourceModal').classList.add('hidden'); }, 1200);
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
    const keys = ['ftr', 'team', 'sow', 'governance', 'leave', 'kpi', 'poc'];
    const spIds = ['urlFtr', 'urlTeam', 'urlSow', 'urlGovernance', 'urlLeave', 'urlKpi', 'urlPoc'];
    const gsIds = ['urlFtrGoogle', 'urlTeamGoogle', 'urlSowGoogle', 'urlGovernanceGoogle', 'urlLeaveGoogle', 'urlKpiGoogle', 'urlPocGoogle'];
    keys.forEach((key, i) => {
        // Check Google Sheets first, then SharePoint
        const gsEl = $(`#${gsIds[i]}`);
        const spEl = $(`#${spIds[i]}`);
        const gsUrl = gsEl ? gsEl.value.trim() : '';
        const spUrl = spEl ? spEl.value.trim() : '';
        if (gsUrl) {
            sources[key] = { url: gsUrl, type: 'google' };
        } else if (spUrl) {
            sources[key] = { url: spUrl, type: 'sharepoint' };
        }
    });
    const statusEl = $('#execSourceStatus');
    if (Object.keys(sources).length === 0) {
        statusEl.className = 'source-status-box error';
        statusEl.textContent = 'Please enter at least one Google Sheets or SharePoint URL.';
        statusEl.classList.remove('hidden');
        return;
    }
    statusEl.className = 'source-status-box';
    statusEl.textContent = 'Connecting & downloading data…';
    statusEl.classList.remove('hidden');
    try {
        const res = await fetch('/api/exec/connect-sources', {
            method: 'POST', headers: authHeaders(), body: JSON.stringify({ sources }),
        });
        const data = await res.json();
        if (data.success) {
            statusEl.className = 'source-status-box success';
            statusEl.textContent = '✓ ' + data.message + ' Data will auto-refresh every 2 minutes.';
            setTimeout(() => { fetchAllExecData(); checkSourceConnection(); $('#mdSourceModal').classList.add('hidden'); }, 1500);
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

const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function monthLabel(dateStr) {
    if (!dateStr) return 'Unknown';
    const m = String(dateStr).match(/^(\d{4})-(\d{2})/);
    if (m) return `${MONTH_NAMES[parseInt(m[2], 10) - 1]} ${m[1]}`;
    const d = new Date(dateStr);
    if (!isNaN(d.getTime())) return `${MONTH_NAMES[d.getMonth()]} ${d.getFullYear()}`;
    return 'Unknown';
}

// ══════════════════════════════════════════════════════════════════
//  CHATBOT — DE Assistant
// ══════════════════════════════════════════════════════════════════

function initChatbot() {
    const widget = document.getElementById('chatbotWidget');
    const toggle = document.getElementById('chatbotToggle');
    const panel = document.getElementById('chatbotPanel');
    const closeBtn = document.getElementById('chatbotClose');
    const input = document.getElementById('chatbotInput');
    const sendBtn = document.getElementById('chatbotSend');
    const messages = document.getElementById('chatbotMessages');
    if (!widget || !toggle) return;

    // Show widget
    widget.style.display = 'block';

    // Toggle panel
    toggle.addEventListener('click', () => {
        const isHidden = panel.classList.contains('hidden');
        panel.classList.toggle('hidden');
        if (isHidden) input.focus();
    });

    closeBtn.addEventListener('click', () => {
        panel.classList.add('hidden');
    });

    // Clear Chat Logic
    const clearChatBtn = document.getElementById('chatbotClearChat');
    if (clearChatBtn) {
        clearChatBtn.addEventListener('click', () => {
            messages.innerHTML = `
                <div class="chat-msg bot">
                  <span class="chat-avatar">🤖</span>
                  <div class="chat-bubble">Chat history cleared! Ask me anything about your dashboard data.<br><br>
                  <strong>Try asking:</strong><br>
                  • "How many team members?"<br>
                  • "Who is on leave today?"<br>
                  • "Show active risks"<br>
                  • "KPI performance summary"<br>
                  • "SOW status for [project]"
                  </div>
                </div>
            `;
        });
    }

    // API Key Settings Panel Toggle & Logic
    const settingsToggle = document.getElementById('chatbotSettingsToggle');
    const settingsPanel = document.getElementById('chatbotSettingsPanel');
    const apiKeyInput = document.getElementById('chatbotApiKeyInput');
    const saveKeyBtn = document.getElementById('chatbotSaveKey');
    const clearKeyBtn = document.getElementById('chatbotClearKey');
    const keyStatusText = document.getElementById('chatbotKeyStatus');

    if (settingsToggle && settingsPanel) {
        settingsToggle.addEventListener('click', (e) => {
            e.stopPropagation();
            settingsPanel.classList.toggle('hidden');
            if (!settingsPanel.classList.contains('hidden')) {
                updateKeySettingsUI();
            }
        });

        saveKeyBtn.addEventListener('click', () => {
            const key = apiKeyInput.value.trim();
            if (key) {
                localStorage.setItem('gemini_api_key', key);
                keyStatusText.innerHTML = 'Status: <span style="color:#10b981; font-weight:600;">Custom key saved!</span>';
                setTimeout(() => { settingsPanel.classList.add('hidden'); }, 1200);
            } else {
                alert('Please enter a valid key or click Clear.');
            }
        });

        clearKeyBtn.addEventListener('click', () => {
            localStorage.removeItem('gemini_api_key');
            apiKeyInput.value = '';
            keyStatusText.innerHTML = 'Status: <span style="color:#94a3b8;">Using default mode</span>';
        });

        function updateKeySettingsUI() {
            const savedKey = localStorage.getItem('gemini_api_key') || '';
            apiKeyInput.value = savedKey;
            if (savedKey) {
                keyStatusText.innerHTML = 'Status: <span style="color:#22d3ee; font-weight:600;">Custom key configured</span>';
            } else {
                keyStatusText.innerHTML = 'Status: <span style="color:#94a3b8;">Using default mode</span>';
            }
        }
        
        // Run once on load
        updateKeySettingsUI();
    }

    // Send message
    sendBtn.addEventListener('click', () => sendChatMessage());
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') sendChatMessage();
    });

    async function sendChatMessage() {
        const question = input.value.trim();
        if (!question) return;

        // Add user message
        appendMessage('user', question);
        input.value = '';

        // Show typing indicator
        const typing = document.createElement('div');
        typing.className = 'chat-msg bot';
        typing.innerHTML = `<span class="chat-avatar">🤖</span><div class="chat-bubble"><div class="chat-typing"><span></span><span></span><span></span></div></div>`;
        messages.appendChild(typing);
        messages.scrollTop = messages.scrollHeight;

        try {
            const token = sessionStorage.getItem(TOKEN_KEY);
            const userApiKey = localStorage.getItem('gemini_api_key') || '';
            const res = await fetch('/api/exec/chat', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer ' + token,
                    'x-gemini-api-key': userApiKey
                },
                body: JSON.stringify({ question }),
            });
            const data = await res.json();
            // Remove typing indicator
            typing.remove();
            // Add bot response
            appendMessage('bot', data.answer || 'Sorry, no answer available.');
        } catch (err) {
            typing.remove();
            appendMessage('bot', '❌ Error communicating with the server. Please try again.');
        }
    }

    function appendMessage(type, content) {
        const msg = document.createElement('div');
        msg.className = `chat-msg ${type}`;
        const avatar = type === 'bot' ? '🤖' : '👤';
        msg.innerHTML = `<span class="chat-avatar">${avatar}</span><div class="chat-bubble">${type === 'user' ? escHtml(content) : content}</div>`;
        messages.appendChild(msg);
        messages.scrollTop = messages.scrollHeight;
    }

    function escHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }
}

function showChatbot() {
    const w = document.getElementById('chatbotWidget');
    if (w) w.style.display = 'block';
}

function hideChatbot() {
    const w = document.getElementById('chatbotWidget');
    if (w) w.style.display = 'none';
}

})();
