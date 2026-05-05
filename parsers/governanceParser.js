/**
 * Governance Parser
 * Sheets: "Highlights & Low lights", "QBR", "FTE", "Risks and Mitigation", "Audits"
 * All sheets have headers at R0.
 */
const XLSX = require('xlsx');

function parseGovernance(filePath) {
    const wb = XLSX.readFile(filePath);
    const result = {
        highlights: [], lowlights: [], qbr: [], fte: [],
        risks: [], audits: [],
        summary: { activeRisks: 0, totalAudits: 0 },
    };

    // ── Highlights & Lowlights ──
    const hlSheet = findSheet(wb, ['Highlights & Low lights', 'Highlights', 'Highlights & Lowlights']);
    if (hlSheet) {
        const rows = XLSX.utils.sheet_to_json(hlSheet, { defval: '' });
        let currentPM = '', currentProject = '', currentMonth = '';
        rows.forEach(r => {
            const pm = r['Project Manager'] || r[Object.keys(r)[0]] || '';
            if (pm.trim()) currentPM = pm.trim();
            const proj = r['Project'] || '';
            if (proj.trim()) currentProject = proj.trim();
            const month = r['Month'] || '';
            if (month.trim()) currentMonth = month.trim();

            const hl = String(r['Highlight'] || '').trim();
            const ll = String(r['Lowlight'] || '').trim();
            if (hl) result.highlights.push({ pm: currentPM, project: currentProject, month: currentMonth, highlight: hl });
            if (ll) result.lowlights.push({ pm: currentPM, project: currentProject, month: currentMonth, lowlight: ll });
        });
    }

    // ── QBR ──
    const qbrSheet = findSheet(wb, ['QBR']);
    if (qbrSheet) {
        const rows = XLSX.utils.sheet_to_json(qbrSheet, { defval: '' });
        rows.forEach(r => {
            const account = r['Account'] || '';
            if (!account.trim()) return;
            result.qbr.push({
                account: account.trim(),
                pms: r['PMs'] || '',
                offshoreLeads: r['Offshore Leads'] || '',
                deliveryLeads: r['Delivery Leads'] || '',
                startDate: excelDate(r['Project Start date']),
                qbrStartDate: excelDate(r['QBR Start Date']),
                deliveryType: r['QBR Delivery Type'] || '',
                comments: r['Comments'] || '',
            });
        });
    }

    // ── FTE ──
    const fteSheet = findSheet(wb, ['FTE']);
    if (fteSheet) {
        const raw = XLSX.utils.sheet_to_json(fteSheet, { header: 1, defval: '' });
        if (raw.length >= 3) {
            // R0: labels with month serial numbers, R1: "Total FTE on ground" repeating, R2+: account data
            const monthHeaders = raw[0].slice(1).map(v => {
                if (typeof v === 'number') return excelDate(v);
                return String(v);
            });
            for (let i = 2; i < raw.length; i++) {
                const account = String(raw[i][0] || '').trim();
                if (!account) continue;
                const monthlyFTE = {};
                raw[i].slice(1).forEach((v, idx) => {
                    if (monthHeaders[idx]) monthlyFTE[monthHeaders[idx]] = parseNum(v);
                });
                result.fte.push({ account, monthlyFTE });
            }
        }
    }

    // ── Risks & Mitigation ──
    const riskSheet = findSheet(wb, ['Risks and Mitigation', 'Risks']);
    if (riskSheet) {
        const rows = XLSX.utils.sheet_to_json(riskSheet, { defval: '' });
        rows.forEach(r => {
            const risk = r['Risks'] || r['Risk'] || '';
            const project = r['Project'] || '';
            if (!project.trim()) return;
            result.risks.push({
                month: r['Month'] || '',
                week: r['Week'] || '',
                pm: r['Project Manager'] || '',
                project: project.trim(),
                description: risk.trim(),
                owner: r['Owner'] || '',
                rootCause: r['Root cause'] || '',
                mitigation: r['Mitigation Steps'] || '',
                impact: r['Impact'] || '',
                status: r['Status'] || '',
            });
        });
        result.summary.activeRisks = result.risks.filter(r =>
            r.status.toLowerCase() === 'ongoing' && r.description.toLowerCase() !== 'no risk'
        ).length;
    }

    // ── Audits ──
    const auditSheet = findSheet(wb, ['Audits', 'Audit']);
    if (auditSheet) {
        const rows = XLSX.utils.sheet_to_json(auditSheet, { defval: '' });
        rows.forEach(r => {
            const pm = r['Project Manager'] || '';
            if (!pm.trim()) return;
            result.audits.push({
                month: r['Month'] || '',
                pm: pm.trim(),
                project: r['Project'] || '',
                details: r['Audit details'] || '',
                findings: r['Details'] || '',
                closureDate: excelDate(r['Audit Closure Date']),
            });
        });
        result.summary.totalAudits = result.audits.length;
    }

    return result;
}

function findSheet(wb, names) {
    for (const n of names) {
        const found = wb.SheetNames.find(s => s.toLowerCase().trim() === n.toLowerCase().trim());
        if (found) return wb.Sheets[found];
    }
    // Fuzzy match
    for (const n of names) {
        const found = wb.SheetNames.find(s => s.toLowerCase().includes(n.toLowerCase()));
        if (found) return wb.Sheets[found];
    }
    return null;
}

function parseNum(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function excelDate(v) {
    if (!v) return '';
    if (typeof v === 'number') return new Date((v - 25569) * 86400000).toISOString().slice(0, 10);
    return String(v);
}

module.exports = { parseGovernance };
