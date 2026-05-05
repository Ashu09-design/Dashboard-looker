/**
 * FTR Tracker Parser
 * Reads: "Account Details" (engagement info), "Bayer Quality Metrics " (QA data per project)
 * Headers at Row 0 in both sheets.
 */
const XLSX = require('xlsx');

function parseFTR(filePath) {
    const wb = XLSX.readFile(filePath);
    const result = { accounts: [], qaMetrics: [], summary: { avgRating: 0, totalProjects: 0, totalPassed: 0 } };

    // ── Sheet 1: Account Details ──
    const accSheet = wb.Sheets[wb.SheetNames[0]];
    if (accSheet) {
        const rows = XLSX.utils.sheet_to_json(accSheet, { defval: '' });
        rows.forEach(r => {
            const name = r['Engagement Name'] || r[Object.keys(r)[0]] || '';
            if (!name || !name.trim()) return;
            result.accounts.push({
                account: name.trim(),
                qaTracking: r['QA Metrics Tracking'] || '',
                expectedProjects: r['Expected projects per month'] || '',
                scrapping: r['Datalayer/DOM scrapping'] || '',
                webTeam: r['Web team- Internal/External'] || '',
                comments: r['Comments'] || '',
            });
        });
    }

    // ── Sheet 2+: Quality Metrics (any sheet with "Quality" or "Metrics" in name) ──
    wb.SheetNames.forEach(sn => {
        if (sn === wb.SheetNames[0]) return; // skip Account Details
        const ws = wb.Sheets[sn];
        if (!ws) return;
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        rows.forEach(r => {
            const proj = r['Project Name'] || r[Object.keys(r)[0]] || '';
            if (!proj || !proj.trim()) return;

            const global = parseNum(r['Global']);
            const globalFailed = parseNum(r['Global Failed']);
            const inpage = parseNum(r['Inpage']);
            const inpageFailed = parseNum(r['Inpage Failed']);
            const form = parseNum(r['Form']);
            const formFailed = parseNum(r['Form Failed']);
            const videos = parseNum(r['Videos']);
            const videosFailed = parseNum(r['Videos Failed']);

            const total = global + inpage + form + videos;
            const failed = globalFailed + inpageFailed + formFailed + videosFailed;
            const passRate = total > 0 ? Math.round(((total - failed) / total) * 10000) / 100 : 100;

            result.qaMetrics.push({
                project: proj.trim(),
                qaDate: excelDate(r['QA Date']),
                global, globalFailed, inpage, inpageFailed,
                form, formFailed, videos, videosFailed,
                total, failed, passRate,
            });
        });
    });

    // Summary
    if (result.qaMetrics.length > 0) {
        const rates = result.qaMetrics.map(m => m.passRate);
        result.summary.avgRating = Math.round(rates.reduce((a, b) => a + b, 0) / rates.length * 100) / 100;
        result.summary.totalProjects = result.qaMetrics.length;
        result.summary.totalPassed = result.qaMetrics.filter(m => m.failed === 0).length;
    }

    // Map avgRating to accounts
    result.accounts.forEach(acc => {
        const metrics = result.qaMetrics.filter(m =>
            m.project.toLowerCase().includes(acc.account.toLowerCase().split(' ')[0])
        );
        acc.avgRating = metrics.length > 0
            ? Math.round(metrics.reduce((s, m) => s + m.passRate, 0) / metrics.length * 100) / 100
            : 0;
        acc.ftrPass = metrics.filter(m => m.failed === 0).length + '/' + metrics.length;
    });

    return result;
}

function parseNum(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function excelDate(v) {
    if (!v) return '';
    if (typeof v === 'number') {
        const d = new Date((v - 25569) * 86400000);
        return d.toISOString().slice(0, 10);
    }
    return String(v);
}

module.exports = { parseFTR };
