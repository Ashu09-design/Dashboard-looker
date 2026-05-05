/**
 * KPI Parser
 * Sheets: "KPI" (main KPI data, R0=header), "Sheet2" (Highlights/Lowlights — same format as Governance)
 * KPI headers: No., BU Name, Department Name, Service Line, Account Name, Metric Name, KPI/Metric Name, Description, Formula, Target, Actual, Status, etc.
 */
const XLSX = require('xlsx');

function parseKPI(filePath) {
    const wb = XLSX.readFile(filePath);
    const result = { metrics: [], kpis: [], summary: { metRate: 0, totalMetrics: 0, metCount: 0 }, highlights: [], lowlights: [] };

    // ── Main KPI sheet ──
    const kpiSheet = wb.Sheets[wb.SheetNames[0]];
    if (kpiSheet) {
        const raw = XLSX.utils.sheet_to_json(kpiSheet, { header: 1, defval: '' });
        if (raw.length < 2) return result;

        // Find header row (contains "Account" or "Metric")
        let headerIdx = 0;
        for (let i = 0; i < Math.min(5, raw.length); i++) {
            const joined = raw[i].map(c => String(c).toLowerCase()).join('|');
            if (joined.includes('metric') || joined.includes('account')) { headerIdx = i; break; }
        }

        const headers = raw[headerIdx].map(h => String(h).trim());
        const col = name => headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));

        for (let i = headerIdx + 1; i < raw.length; i++) {
            const r = raw[i];
            const accountIdx = col('Account Name');
            const metricIdx = col('Metric Name');
            const kpiNameIdx = col('KPI');
            const targetIdx = col('Target');
            const actualIdx = col('Actual');
            const statusIdx = col('Status');

            const account = accountIdx >= 0 ? String(r[accountIdx] || '').trim() : '';
            const metric = metricIdx >= 0 ? String(r[metricIdx] || '').trim() : '';
            if (!account && !metric) continue;

            const target = targetIdx >= 0 ? r[targetIdx] : '';
            const actual = actualIdx >= 0 ? r[actualIdx] : '';
            const status = statusIdx >= 0 ? String(r[statusIdx] || '').trim() : '';

            const kpiName = kpiNameIdx >= 0 ? String(r[kpiNameIdx] || '').trim() : metric;
            const description = col('Metric Description') >= 0 ? String(r[col('Metric Description')] || '').trim() : '';
            const formula = col('Formula') >= 0 ? String(r[col('Formula')] || '').trim() : '';

            result.metrics.push({
                account, metric, name: kpiName, description, formula,
                target: parseTarget(target), actual: parseTarget(actual),
                targetRaw: String(target), actualRaw: String(actual),
                status,
            });
            result.kpis.push({
                account, metric: kpiName || metric,
                target: parseTarget(target), actual: parseTarget(actual), status,
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

    return result;
}

function parseTarget(v) {
    if (!v && v !== 0) return 0;
    const s = String(v).replace(/[%$,]/g, '').trim();
    const n = Number(s);
    return isNaN(n) ? 0 : n;
}

module.exports = { parseKPI };
