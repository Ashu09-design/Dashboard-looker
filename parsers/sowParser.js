/**
 * SOW & PO Tracker Parser
 * Sheets: "FTE Projects SOW & PO" (R0=header), "POC and Unitized Projects SOW" (R0=header), "Sheet1" (R0=header)
 * All sheets have: PM Name, Client, Sub Project, Project Name, Total SOW Value, SOW Status, etc.
 */
const XLSX = require('xlsx');

function parseSOW(filePath) {
    const wb = XLSX.readFile(filePath);
    const result = { projects: [], summary: { activeProjects: 0, totalSOWValue: 0, sowReceived: 0, poReceived: 0 } };

    wb.SheetNames.forEach(sn => {
        const ws = wb.Sheets[sn];
        if (!ws) return;
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (raw.length < 2) return;

        // Find header row (first row that has PM/Client/Project in it)
        let headerIdx = -1;
        for (let i = 0; i < Math.min(5, raw.length); i++) {
            const joined = raw[i].map(c => String(c).toLowerCase()).join('|');
            if (joined.includes('pm') && (joined.includes('client') || joined.includes('project'))) {
                headerIdx = i; break;
            }
        }
        if (headerIdx < 0) return;

        const headers = raw[headerIdx].map(h => String(h).trim());
        const col = name => {
            const idx = headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
            return idx;
        };

        for (let i = headerIdx + 1; i < raw.length; i++) {
            const r = raw[i];
            const pm = String(r[col('PM')] || '').trim();
            const client = String(r[col('Client')] || '').trim();
            const projectName = String(r[col('Project Name')] || r[col('Sub Project')] || '').trim();
            if (!projectName && !client) continue;

            const sowValue = parseNum(r[col('Total SOW')]);
            const sowStatus = String(r[col('SOW Status')] || '').trim();
            const poStatus = String(r[col('PO Status')] || '').trim();
            const projectStatus = String(r[col('Project Status')] || '').trim();

            result.projects.push({
                pm, client, projectName: projectName || client,
                subProject: String(r[col('Sub Project')] || '').trim(),
                mysheetsPID: String(r[col('MySheets')] || '').trim(),
                sowValue,
                deRevenueShare: parseNum(r[col('DE Revenue')]),
                startDate: excelDate(r[col('Project Start')]),
                endDate: excelDate(r[col('Project End')]),
                sowStatus: sowStatus || 'Unknown',
                poStatus: poStatus || 'Unknown',
                projectStatus: projectStatus || 'Active',
                sheet: sn,
            });

            result.summary.totalSOWValue += sowValue;
            if (sowStatus.toLowerCase() === 'received') result.summary.sowReceived++;
            if (poStatus.toLowerCase() === 'received') result.summary.poReceived++;
        }
    });

    result.summary.activeProjects = result.projects.length;
    return result;
}

function parseNum(v) { const n = Number(String(v).replace(/[$,]/g, '')); return isNaN(n) ? 0 : n; }
function excelDate(v) {
    if (!v) return '';
    if (typeof v === 'number') return new Date((v - 25569) * 86400000).toISOString().slice(0, 10);
    return String(v);
}

module.exports = { parseSOW };
