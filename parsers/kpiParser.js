/**
 * KPI Parser
 * Sheets: "KPI" (main KPI data, R0=header), "Sheet2" (Highlights/Lowlights — same format as Governance)
 * KPI headers: No., BU Name, Department Name, Service Line, Account Name, Metric Name, KPI/Metric Name, Description, Formula, Target, Actual, Status, etc.
 */
const XLSX = require('xlsx');

function parseKPI(filePath) {
    const wb = XLSX.readFile(filePath, { cellDates: true });
    const result = { metrics: [], kpis: [], scorecardProjects: [], summary: { metRate: 0, totalMetrics: 0, metCount: 0 }, highlights: [], lowlights: [] };

    // Find the correct KPI sheet (preferring "KPI Master" or "KPI" over other sheet names)
    let kpiSheetName = wb.SheetNames.find(name => {
        const lower = name.toLowerCase().trim();
        return lower === 'kpi master' || lower === 'kpi';
    });
    if (!kpiSheetName) {
        kpiSheetName = wb.SheetNames.find(name => name.toLowerCase().includes('kpi'));
    }
    if (!kpiSheetName) {
        kpiSheetName = wb.SheetNames[0];
    }

    const kpiSheet = wb.Sheets[kpiSheetName];
    if (kpiSheet) {
        const raw = XLSX.utils.sheet_to_json(kpiSheet, { header: 1, defval: '', cellDates: true });
        if (raw.length < 2) return result;

        // Find header row (contains "Account" or "Metric")
        let headerIdx = 0;
        for (let i = 0; i < Math.min(10, raw.length); i++) {
            const joined = raw[i].map(c => String(c).toLowerCase()).join('|');
            if (joined.includes('metric') || joined.includes('account')) { headerIdx = i; break; }
        }

        const headers = raw[headerIdx].map(h => String(h).trim());
        const col = name => headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));

        const accountIdx = col('Account Name');
        const metricIdx = col('Metric Name');
        const kpiNameIdx = col('KPI');
        const targetIdx = col('Target');
        const actualIdx = col('Actual');
        const statusIdx = col('Status');
        const projectIdx = col('Project Name') >= 0 ? col('Project Name') : (col('Project') >= 0 ? col('Project') : col('Service Line'));
        const dateIdx = col('Start Date') >= 0 ? col('Start Date') : col('Date');

        for (let i = headerIdx + 1; i < raw.length; i++) {
            const r = raw[i];

            const account = accountIdx >= 0 ? String(r[accountIdx] || '').trim() : '';
            const metric = metricIdx >= 0 ? String(r[metricIdx] || '').trim() : '';
            if (!account && !metric) continue;

            const target = targetIdx >= 0 ? r[targetIdx] : '';
            const actual = actualIdx >= 0 ? r[actualIdx] : '';
            const status = statusIdx >= 0 ? String(r[statusIdx] || '').trim() : '';

            const kpiName = kpiNameIdx >= 0 ? String(r[kpiNameIdx] || '').trim() : metric;
            const description = col('Metric Description') >= 0 ? String(r[col('Metric Description')] || '').trim() : '';
            const formula = col('Formula') >= 0 ? String(r[col('Formula')] || '').trim() : '';
            const project = projectIdx >= 0 ? String(r[projectIdx] || '').trim() : '';

            let month = '';
            if (dateIdx >= 0 && r[dateIdx]) {
                const rawDate = r[dateIdx];
                if (rawDate instanceof Date) {
                    month = rawDate.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
                } else if (typeof rawDate === 'number') {
                    try {
                        const dateObj = new Date((rawDate - 25569) * 86400 * 1000);
                        month = dateObj.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
                    } catch (e) {}
                } else {
                    const d = new Date(String(rawDate));
                    if (!isNaN(d.getTime())) {
                        month = d.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
                    }
                }
            }

            result.metrics.push({
                account, metric, name: kpiName, description, formula,
                target: parseTarget(target), actual: parseTarget(actual),
                targetRaw: String(target), actualRaw: String(actual),
                status,
                project,
                month,
            });
            result.kpis.push({
                account, metric: kpiName || metric,
                target: parseTarget(target), actual: parseTarget(actual), status,
                project,
                month,
            });
        }

        // Summary
        result.summary.totalMetrics = result.metrics.length;
        result.summary.metCount = result.metrics.filter(m =>
            m.status.toLowerCase().includes('met') || m.status.toLowerCase().includes('achieved') ||
            m.status.toLowerCase().includes('green') || (m.actual >= m.target && m.target > 0)
        ).length;
        result.summary.metRate = result.summary.totalMetrics > 0
            ? Math.round(result.summary.metCount / result.summary.totalMetrics * 100)
            : 0;
    }

    // ── Sheet2: Highlights/Lowlights (if present) ──
    const sheet2 = wb.Sheets[wb.SheetNames[1]];
    if (sheet2) {
        const rows = XLSX.utils.sheet_to_json(sheet2, { defval: '' });
        let currentPM = '', currentProject = '', currentMonth = '';
        rows.forEach(r => {
            const pm = r['Project Manager'] || ''; if (pm.trim()) currentPM = pm.trim();
            const proj = r['Project'] || ''; if (proj.trim()) currentProject = proj.trim();
            const month = r['Month'] || ''; if (month.trim()) currentMonth = month.trim();
            const hl = String(r['Highlight'] || '').trim();
            const ll = String(r['Lowlight'] || '').trim();
            if (hl) result.highlights.push({ pm: currentPM, project: currentProject, month: currentMonth, highlight: hl });
            if (ll) result.lowlights.push({ pm: currentPM, project: currentProject, month: currentMonth, lowlight: ll });
        });
    }

    // ── Parse Web Analytics KPI Summary for project-wise monthly breakdown ──
    let scSheetName = wb.SheetNames.find(n => {
        const l = n.toLowerCase().trim();
        return l.includes('web analytics') || l.includes('kpi summary');
    });
    if (!scSheetName) {
        // Fallback: find sheet with timeliness + quality + client + project in first 8 rows
        for (const sn of wb.SheetNames) {
            const s = wb.Sheets[sn];
            const peek = XLSX.utils.sheet_to_json(s, { header: 1, defval: '', range: 'A1:L8' });
            const blob = peek.map(r => r.map(c => String(c).toLowerCase()).join('|')).join('|');
            if (blob.includes('timeliness') && blob.includes('quality') && blob.includes('client') && blob.includes('project')) {
                scSheetName = sn; break;
            }
        }
    }
    if (scSheetName) {
        const scSheet = wb.Sheets[scSheetName];
        const scRaw = XLSX.utils.sheet_to_json(scSheet, { header: 1, defval: '', cellDates: true });
        // Find header row
        let scHdr = 0;
        for (let i = 0; i < Math.min(10, scRaw.length); i++) {
            const joined = scRaw[i].map(c => String(c).toLowerCase()).join('|');
            if (joined.includes('client') && joined.includes('project')) { scHdr = i; break; }
        }
        const scHeaders = scRaw[scHdr].map(h => String(h).toLowerCase().trim());
        const scCol = (...subs) => {
            for (let idx = 0; idx < scHeaders.length; idx++) {
                if (subs.some(s => scHeaders[idx].includes(s))) return idx;
            }
            return -1;
        };
        const ci = { client: scCol('client'), project: scCol('project name', 'project'), pm: scCol('pm', 'project manager'),
            date: scCol('start date', 'measurement start', 'period'),
            timeliness: scCol('timeliness'), utilization: scCol('utli', 'utili'), quality: scCol('quality'),
            csatFreq: scCol('csat ( frequency', 'csat (frequency', 'csat frequency'),
            csatRating: scCol('csat ( rating', 'csat (rating', 'csat rating'),
            clientTraining: scCol('client training'), comments: scCol('comments') };

        const fmtPct = v => { const n = parseTarget(v); return n ? (n <= 1 ? (n * 100).toFixed(1) + '%' : n.toFixed(1) + '%') : ''; };

        for (let i = scHdr + 1; i < scRaw.length; i++) {
            const r = scRaw[i];
            const client = ci.client >= 0 ? String(r[ci.client] || '').trim() : '';
            const project = ci.project >= 0 ? String(r[ci.project] || '').trim().replace(/\s+/g, ' ') : '';
            if (!client && !project) continue;
            const pm = ci.pm >= 0 ? String(r[ci.pm] || '').trim() : '';

            let month = '';
            if (ci.date >= 0 && r[ci.date]) {
                const rd = r[ci.date];
                if (rd instanceof Date) month = rd.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
                else { const d = new Date(String(rd)); if (!isNaN(d.getTime())) month = d.toLocaleDateString('en-US', { month: 'short', year: 'numeric' }); }
            }

            result.scorecardProjects.push({
                client, project, pm, month,
                timeliness: ci.timeliness >= 0 ? fmtPct(r[ci.timeliness]) : '',
                utilization: ci.utilization >= 0 ? fmtPct(r[ci.utilization]) : '',
                quality: ci.quality >= 0 ? fmtPct(r[ci.quality]) : '',
                csatFrequency: ci.csatFreq >= 0 ? String(r[ci.csatFreq] || '').trim() : '',
                csatRating: ci.csatRating >= 0 ? String(r[ci.csatRating] || '').trim() : '',
                clientTraining: ci.clientTraining >= 0 ? fmtPct(r[ci.clientTraining]) : '',
                comments: ci.comments >= 0 ? String(r[ci.comments] || '').trim() : '',
            });
        }
    }

    return result;
}

function parseTarget(v) {
    if (!v && v !== 0) return 0;
    const s = String(v).replace(/[%$,]/g, '').trim();
    const n = Number(s);
    return isNaN(n) ? 0 : n;
}

module.exports = { parseKPI };
