/* ═══════════════════════════════════════════════════════════════
   Manager Dashboard — Tabbed Curated Views (executive.js)
   Embedded inside the Daily Standup page (index.html).
   Wrapped in an IIFE so nothing collides with app.js globals.
   Exposes: window.ManagerDash = { open, close }
   ═══════════════════════════════════════════════════════════════ */

(function () {
'use strict';

const TOKEN_KEY = 'exec_token';
const REFRESH_MS = 5 * 60 * 1000;
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
    fetchAllExecData();
    if (!state.intervalId) state.intervalId = setInterval(fetchAllExecData, REFRESH_MS);
}

function close() {
    if (state.intervalId) { clearInterval(state.intervalId); state.intervalId = null; }
    destroyCharts();
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

async function fetchAllExecData() {
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
        renderTab();
    } catch (err) {
        console.error('Manager dashboard fetch error:', err);
        if (err.message && err.message.includes('401')) { onUnauthorized(); return; }
    }
    showLoading(false);
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
            renderTab();
            $('#execSection')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
        });
    });
}

// ══════════════════════════════════════════════════════════════════
//  FILTER BAR
// ══════════════════════════════════════════════════════════════════

const TAB_FILTERS = {
    team: ['office', 'designation'],
    projects: ['project'],
    sow: ['client'],
    kpis: ['client', 'month'],
    ftr: ['client', 'month'],
    leave: ['month'],
    risks: ['project', 'month'],
    poc: ['pocStatus', 'manager'],
};

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
        case 'project':
            return uniqueSorted((d.health || []).map(p => p.projectName));
        case 'client':
            return uniqueSorted([
                ...(d.sow?.projects || []).map(p => p.client),
                ...(d.kpi?.metrics || []).map(m => m.account),
                ...(d.ftr?.accounts || []).map(a => a.account),
            ]);
        case 'month':
            return uniqueSorted([
                ...(d.kpi?.metrics || []).map(m => m.month),
                ...(d.gov?.risks || []).map(r => r.month),
                ...(d.ftr?.qaMetrics || []).map(q => monthLabel(q.qaDate)),
            ]);
        case 'pocStatus':
            return uniqueSorted((d.poc?.pocs || []).map(p => p.status));
        case 'manager':
            return uniqueSorted((d.poc?.pocs || []).map(p => p.manager));
        default:
            return [];
    }
}

const FILTER_LABELS = {
    office: '🏢 Office', role: '👤 Role', designation: '👤 Designation', project: '📂 Project', client: '🏷️ Client',
    month: '📅 Month', pocStatus: '🚦 Status', manager: '👔 Manager',
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
        const opts = filterOptions(key);
        html += `<label class="filter-ctrl">${FILTER_LABELS[key] || key}
            <select data-filter="${key}">
                <option value="">All</option>
                ${opts.map(o => `<option value="${esc(o)}" ${cur[key] === o ? 'selected' : ''}>${esc(o)}</option>`).join('')}
            </select></label>`;
    });
    html += '<button class="filter-clear" id="btnClearFilters">Clear</button>';
    bar.innerHTML = html;

    bar.querySelectorAll('select[data-filter]').forEach(sel => {
        sel.addEventListener('change', () => {
            cur[sel.dataset.filter] = sel.value;
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

function renderTab() {
    renderFilterBar();
    destroyCharts();
    const el = $('#tabContent');
    if (!el) return;
    el.innerHTML = '';
    const fn = {
        overview: tabOverview, team: tabTeam, projects: tabProjects, sow: tabSow,
        kpis: tabKpis, ftr: tabFtr, leave: tabLeave, risks: tabRisks, poc: tabPoc,
    }[state.activeTab];
    if (fn) fn(el);
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

function tabOverview(el) {
    const s = state.data.summary || {};
    const poc = state.data.poc?.summary || {};
    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('👥', s.teamSize || 0, 'Team Size')}
            ${statCard('📂', s.activeProjects || 0, 'Active Projects')}
            ${statCard('🎯', s.kpiMetRate ? s.kpiMetRate + '%' : '—', 'KPI Met Rate')}
            ${statCard('✅', s.ftrAvgRating ? s.ftrAvgRating + '%' : '—', 'Avg FTR Rating')}
            ${statCard('🏖️', s.onLeaveToday || 0, 'On Leave Today', 'kpi-card-alert')}
            ${statCard('⚠️', s.activeRisks || 0, 'Active Risks', 'kpi-card-alert')}
            ${statCard('🧪', poc.total || 0, 'PoC Use Cases')}
        </div>
        <div class="grid-2">
            <div class="m-card">${section('🏥 Project Health Snapshot')}<div id="ovHealth"></div></div>
            <div class="m-card">${section('⚠️ Active Risk Alerts')}<div id="ovRisks"></div></div>
        </div>
        <div class="grid-2">
            <div class="m-card">${section('🌟 Recent Highlights')}<div id="ovHL"></div></div>
            <div class="m-card">${section('⬇️ Recent Lowlights')}<div id="ovLL"></div></div>
        </div>`;
    widgetProjectHealth($('#ovHealth'), {}, 6);
    widgetRiskList($('#ovRisks'), {}, 6);
    widgetHighlights($('#ovHL'), 6);
    widgetLowlights($('#ovLL'), 6);
}

// ══════════════════════════════════════════════════════════════════
//  TAB: TEAM DETAILS
// ══════════════════════════════════════════════════════════════════

function tabTeam(el) {
    const t = state.data.team || {};
    let members = t.activeMembers || [];
    const office = getFilter('office'), designation = getFilter('designation');
    if (office) members = members.filter(m => m.location === office);
    if (designation) members = members.filter(m => m.designation === designation);

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
        ['Name', 'Role', 'Designation', 'Emp ID', 'Office / Location', 'Email', 'Date of Joining'],
        members.map(m => [
            `<strong>${esc(m.name)}</strong>`, esc(m.role), esc(m.designation),
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
    const proj = getFilter('project');
    el.innerHTML = `
        <div class="m-card">${section('🏥 Project Health Status')}<div id="pjHealth"></div></div>
        <div class="grid-2">
            <div class="m-card">${section('🌟 Recent Highlights')}<div id="pjHL"></div></div>
            <div class="m-card">${section('⬇️ Recent Lowlights')}<div id="pjLL"></div></div>
        </div>
        <div class="grid-2">
            <div class="m-card">${section('📈 FTE Utilization Trend')}${chartCanvas('chPjUtil')}</div>
            <div class="m-card">${section('⚠️ Project Risks')}<div id="pjRisks"></div></div>
        </div>
        <div class="m-card">${section('✅ FTR Report', 'Client-wise & monthly')}<div id="pjFtr"></div></div>
        <div class="grid-2">
            <div class="m-card">${section('🏆 Top Leave Takers')}<div id="pjLeaveTop"></div></div>
            <div class="m-card">${section('🏖️ Leave Impact — Today & This Month')}<div id="pjLeaveImpact"></div></div>
        </div>`;

    const pf = proj ? { project: proj } : {};
    widgetProjectHealth($('#pjHealth'), pf);
    widgetHighlights($('#pjHL'), 10, proj);
    widgetLowlights($('#pjLL'), 10, proj);
    widgetUtilTrendChart('chPjUtil');
    widgetRiskList($('#pjRisks'), pf, 10);
    widgetFtrReport($('#pjFtr'));
    widgetTopLeaveTakers($('#pjLeaveTop'));
    widgetLeaveImpact($('#pjLeaveImpact'));
}

// ══════════════════════════════════════════════════════════════════
//  TAB: SOW & PO
// ══════════════════════════════════════════════════════════════════

function tabSow(el) {
    let projects = state.data.sow?.projects || [];
    const client = getFilter('client');
    if (client) projects = projects.filter(p => p.client === client);

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
    const sm = kpi.summary || {};
    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('🎯', sm.metRate != null ? sm.metRate + '%' : '—', 'KPI Met Rate')}
            ${statCard('📊', sm.totalMetrics || 0, 'Total Metrics')}
            ${statCard('✅', sm.metCount || 0, 'Targets Met')}
        </div>
        <div class="m-card">${section('🎯 KPI Scorecard', 'Target vs Actual')}<div id="kpiScore"></div></div>
        <div class="m-card">${section('🏷️ KPI — Client-wise & Monthly View', 'Avg score across all projects')}<div id="kpiPivot"></div></div>
        <div class="grid-2">
            <div class="m-card">${section('📈 FTE Utilization — Client-wise & Monthly')}<div id="kpiUtil"></div></div>
            <div class="m-card">${section('⚠️ Project Risks')}<div id="kpiRisks"></div></div>
        </div>
        <div class="m-card">${section('✅ FTR Report', 'Client-wise & monthly')}<div id="kpiFtr"></div></div>
        <div class="m-card">${section('🏥 Project Health Status')}<div id="kpiHealth"></div></div>`;

    widgetKpiScorecard($('#kpiScore'));
    widgetKpiPivot($('#kpiPivot'));
    widgetKpiUtilPivot($('#kpiUtil'));
    widgetRiskList($('#kpiRisks'), {}, 10);
    widgetFtrReport($('#kpiFtr'));
    widgetProjectHealth($('#kpiHealth'), {});
}

// ══════════════════════════════════════════════════════════════════
//  TAB: FTR
// ══════════════════════════════════════════════════════════════════

function tabFtr(el) {
    const ftr = state.data.ftr || {};
    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('✅', ftr.summary?.avgRating ? ftr.summary.avgRating + '%' : '—', 'Avg FTR Rating')}
            ${statCard('📂', (ftr.accounts || []).length, 'Accounts')}
            ${statCard('🔍', (ftr.qaMetrics || []).length, 'QA Records')}
        </div>
        <div class="m-card">${section('✅ FTR Report — Client-wise & Monthly')}<div id="ftrReport"></div></div>
        <div class="m-card">${section('🏷️ FTR by Account (Client-wise)')}<div id="ftrAccounts"></div></div>
        <div class="m-card">${section('🔍 QA Metrics Detail')}<div id="ftrQa"></div></div>`;

    widgetFtrReport($('#ftrReport'));

    let accounts = ftr.accounts || [];
    const client = getFilter('client');
    if (client) accounts = accounts.filter(a => a.account === client);
    $('#ftrAccounts').innerHTML = dataTable(
        ['Account', 'QA Tracking', 'Expected Projects', 'Web Team', 'Avg Rating', 'FTR Pass'],
        accounts.map(a => [
            `<strong>${esc(a.account)}</strong>`, esc(a.qaTracking), esc(a.expectedProjects),
            esc(a.webTeam), ratingCell(a.avgRating), esc(a.ftrPass),
        ]),
        { empty: 'No accounts match the filter.' }
    );

    let qa = ftr.qaMetrics || [];
    const month = getFilter('month');
    if (month) qa = qa.filter(q => monthLabel(q.qaDate) === month);
    if (client) qa = qa.filter(q => (q.project || '').toLowerCase().includes(client.toLowerCase()));
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
    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('🏖️', (cm.onLeaveToday || []).length, 'On Leave Today', 'kpi-card-alert')}
            ${statCard('📅', cm.totalLeaves || 0, 'Leaves This Month')}
            ${statCard('🗓️', (leave.months || []).length, 'Months Tracked')}
        </div>
        <div class="grid-2">
            <div class="m-card">${section('🏖️ On Leave Today')}<div id="lvToday"></div></div>
            <div class="m-card">${section('🏆 Top Leave Takers — This Month')}<div id="lvTop"></div></div>
        </div>
        <div class="m-card">${section('🗓️ Leave Tracker — Months Available')}<div id="lvMonths"></div></div>`;

    const onToday = cm.onLeaveToday || [];
    $('#lvToday').innerHTML = dataTable(
        ['Member', 'Leave Type'],
        onToday.map(p => [`<strong>${esc(p.name)}</strong>`, esc(p.type)]),
        { empty: 'Nobody is on leave today. ✅' }
    );

    const byPerson = Object.entries(cm.byPerson || {})
        .filter(([, c]) => c > 0).sort((a, b) => b[1] - a[1]);
    $('#lvTop').innerHTML = dataTable(
        ['Member', 'Leave Days'],
        byPerson.slice(0, 20).map(([n, c]) => [`<strong>${esc(n)}</strong>`, badge(c + ' day(s)', 'b-amber')]),
        { empty: 'No leave records for the current month.' }
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
    const active = risks.filter(r => (r.status || '').toLowerCase() === 'ongoing');
    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('⚠️', active.length, 'Active Risks', 'kpi-card-alert')}
            ${statCard('📋', risks.length, 'Total Risk Records')}
        </div>
        <div class="m-card">${section('📊 Risks — Project-wise & Monthly', 'Count of risk records')}<div id="rkPivot"></div></div>
        <div class="m-card">${section('⚠️ Risk Details')}<div id="rkList"></div></div>`;

    const proj = getFilter('project'), month = getFilter('month');
    let pv = risks;
    if (proj) pv = pv.filter(r => r.project === proj);
    if (month) pv = pv.filter(r => r.month === month);
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
    const sm = poc.summary || {};
    const st = getFilter('pocStatus'), mgr = getFilter('manager');
    if (st) pocs = pocs.filter(p => p.status === st);
    if (mgr) pocs = pocs.filter(p => p.manager === mgr);

    el.innerHTML = `
        <div class="kpi-row">
            ${statCard('🧪', sm.total || 0, 'Total Use Cases')}
            ${statCard('✅', sm.completed || 0, 'Completed')}
            ${statCard('🔧', sm.inProgress || 0, 'In Progress')}
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

    doughnutChart('chPocStatus', Object.keys(sm.byStatus || {}), Object.values(sm.byStatus || {}));
    const mgrCount = {};
    (poc.pocs || []).forEach(p => { mgrCount[p.manager || 'Unknown'] = (mgrCount[p.manager || 'Unknown'] || 0) + 1; });
    barChart('chPocMgr', Object.keys(mgrCount), Object.values(mgrCount), 'PoCs', '#06b6d4');
}

// ══════════════════════════════════════════════════════════════════
//  REUSABLE WIDGETS
// ══════════════════════════════════════════════════════════════════

function widgetProjectHealth(el, filter, limit) {
    let rows = state.data.health || [];
    if (filter && filter.project) rows = rows.filter(p => p.projectName === filter.project);
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
    if (filter && filter.project) risks = risks.filter(r => r.project === filter.project);
    if (filter && filter.month) risks = risks.filter(r => r.month === filter.month);
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

function widgetHighlights(el, limit, project) {
    let items = (state.data.gov?.highlights || []).filter(h => h.highlight);
    if (project) items = items.filter(h => h.project === project);
    items = items.slice(0, limit || 10);
    el.innerHTML = items.length ? items.map(h => `
        <div class="hl-item">
            <div class="hl-project">${esc(h.project)} <span class="muted">— ${esc(h.month)}</span></div>
            <div class="hl-text">${esc(h.highlight)}</div>
        </div>`).join('') : '<div class="empty-note">No highlights reported.</div>';
}

function widgetLowlights(el, limit, project) {
    let items = (state.data.gov?.lowlights || []).filter(h => h.lowlight);
    if (project) items = items.filter(h => h.project === project);
    items = items.slice(0, limit || 10);
    el.innerHTML = items.length ? items.map(h => `
        <div class="hl-item ll-item">
            <div class="hl-project">${esc(h.project)} <span class="muted">— ${esc(h.month)}</span></div>
            <div class="hl-text">${esc(h.lowlight)}</div>
        </div>`).join('') : '<div class="empty-note">No lowlights reported.</div>';
}

function widgetUtilTrendChart(canvasId) {
    const trend = state.data.gov?.fteTrend || [];
    if (!trend.length) {
        const c = document.getElementById(canvasId);
        if (c) c.replaceWith(Object.assign(document.createElement('div'),
            { className: 'empty-note', textContent: 'No FTE trend data.' }));
        return;
    }
    lineChart(canvasId, trend.map(t => t.month), trend.map(t => t.totalFTE), 'Total FTE');
}

function widgetTopLeaveTakers(el) {
    const byPerson = Object.entries(state.data.leave?.currentMonth?.byPerson || {})
        .filter(([, c]) => c > 0).sort((a, b) => b[1] - a[1]).slice(0, 10);
    el.innerHTML = dataTable(
        ['Member', 'Leave Days'],
        byPerson.map(([n, c]) => [`<strong>${esc(n)}</strong>`, badge(c + ' day(s)', 'b-amber')]),
        { empty: 'No leave records this month.' }
    );
}

function widgetLeaveImpact(el) {
    const cm = state.data.leave?.currentMonth || {};
    const onToday = cm.onLeaveToday || [];
    el.innerHTML = `
        <div class="impact-row">
            <div class="impact-stat"><span class="impact-num">${onToday.length}</span> on leave today</div>
            <div class="impact-stat"><span class="impact-num">${cm.totalLeaves || 0}</span> leaves this month</div>
        </div>
        ${onToday.length ? dataTable(['Member', 'Type'],
        onToday.map(p => [esc(p.name), esc(p.type)])) : '<div class="empty-note">Nobody on leave today. ✅</div>'}`;
}

function widgetKpiScorecard(el) {
    let metrics = state.data.kpi?.metrics || [];
    const client = getFilter('client'), month = getFilter('month');
    if (client) metrics = metrics.filter(m => m.account === client);
    if (month) metrics = metrics.filter(m => m.month === month);
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
    const client = getFilter('client'), month = getFilter('month');
    if (client) metrics = metrics.filter(m => m.account === client);
    if (month) metrics = metrics.filter(m => m.month === month);
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
    const client = getFilter('client'), month = getFilter('month');
    if (client) metrics = metrics.filter(m => m.account === client);
    if (month) metrics = metrics.filter(m => m.month === month);
    const rowKeys = uniqueSorted(metrics.map(m => m.account));
    const colKeys = uniqueSorted(metrics.map(m => m.month));
    el.innerHTML = pivotTable(rowKeys, colKeys, (rk, ck) => {
        const cell = metrics.filter(m => m.account === rk && m.month === ck);
        if (!cell.length) return '';
        const avg = cell.reduce((s, m) => s + m.actual, 0) / cell.length;
        return `<span class="badge b-blue">${Math.round(avg * 100)}%</span>`;
    }, { corner: 'Client', empty: 'No FTE utilization data.' });
}

function widgetFtrReport(el) {
    const qa = state.data.ftr?.qaMetrics || [];
    const accounts = (state.data.ftr?.accounts || []).map(a => a.account);
    function clientOf(project) {
        const lp = (project || '').toLowerCase();
        const hit = accounts.find(a => a && lp.includes(a.toLowerCase()));
        return hit || 'Unmapped';
    }
    let recs = qa.map(q => ({ ...q, client: clientOf(q.project), month: monthLabel(q.qaDate) }));
    const client = getFilter('client'), month = getFilter('month');
    if (client) recs = recs.filter(r => r.client === client);
    if (month) recs = recs.filter(r => r.month === month);
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
    const ids = ['urlFtr', 'urlTeam', 'urlSow', 'urlGovernance', 'urlLeave', 'urlKpi', 'urlPoc'];
    keys.forEach((key, i) => {
        const el = $(`#${ids[i]}`);
        const url = el ? el.value.trim() : '';
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
            method: 'POST', headers: authHeaders(), body: JSON.stringify({ sources }),
        });
        const data = await res.json();
        if (data.success) {
            statusEl.className = 'source-status-box success';
            statusEl.textContent = '✓ ' + data.message;
            setTimeout(() => { fetchAllExecData(); $('#mdSourceModal').classList.add('hidden'); }, 1200);
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

})();
