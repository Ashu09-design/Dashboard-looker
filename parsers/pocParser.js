const XLSX = require('xlsx');

function parsePoC(filePath) {
    const wb = XLSX.readFile(filePath);
    const result = {
        pocs: [],
        aiUsecases: [],
        summary: { total: 0, byStatus: {}, completed: 0, inProgress: 0 }
    };

    // Find the sheets
    let ucSheet = null;
    let aiSheet = null;

    wb.SheetNames.forEach(sn => {
        const lower = sn.toLowerCase().trim();
        if (['use case details', 'poc details', 'use case'].includes(lower)) {
            ucSheet = wb.Sheets[sn];
        } else if (['ai usecase', 'ai use case', 'ai'].includes(lower)) {
            aiSheet = wb.Sheets[sn];
        }
    });

    // Fallback if sheet names are slightly different
    if (!ucSheet) {
        // Find first sheet that has 'title' or 'poc' in its header rows
        for (let sn of wb.SheetNames) {
            const ws = wb.Sheets[sn];
            const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
            const limit = Math.min(8, raw.length);
            let found = false;
            for (let i = 0; i < limit; i++) {
                const joined = raw[i].map(c => String(c).toLowerCase()).join('|');
                if (joined.includes('title') || joined.includes('poc')) {
                    ucSheet = ws;
                    found = true;
                    break;
                }
            }
            if (found) break;
        }
    }

    if (ucSheet) {
        const raw = XLSX.utils.sheet_to_json(ucSheet, { header: 1, defval: '' });
        let hdrIdx = 0;
        for (let i = 0; i < Math.min(8, raw.length); i++) {
            const joined = raw[i].map(c => String(c).toLowerCase()).join('|');
            if (joined.includes('title') || joined.includes('poc')) {
                hdrIdx = i;
                break;
            }
        }

        const headers = (raw[hdrIdx] || []).map(h => String(h).trim().toLowerCase());
        const col = (subs) => {
            return headers.findIndex(h => subs.some(s => h.includes(s)));
        };

        const snoCol = col(['s. no', 's.no', 'sno']);
        const managerCol = col(['reporting manager', 'manager']);
        const titleCol = col(['title']);
        const descCol = col(['description']);
        const spocCol = col(['spoc']);
        const teamCol = col(['team details', 'team']);
        const statusCol = col(['status']);
        const lastWorkedCol = col(['last worked', 'date']);
        const expCompCol = col(['expected completion', 'completion']);
        const artifactsCol = col(['artifacts']);
        const commentsCol = col(['comments', 'links']);

        for (let i = hdrIdx + 1; i < raw.length; i++) {
            const r = raw[i];
            // Skip rows where both title and manager are empty
            const title = titleCol !== -1 && titleCol < r.length ? String(r[titleCol] || '').trim() : '';
            const manager = managerCol !== -1 && managerCol < r.length ? String(r[managerCol] || '').trim() : '';
            if (!title && !manager) continue;

            const statusRaw = statusCol !== -1 && statusCol < r.length ? String(r[statusCol] || '').trim() : '';
            const status = cleanStatus(statusRaw);

            result.pocs.push({
                sno: snoCol !== -1 && snoCol < r.length ? String(r[snoCol] || '').trim() : '',
                manager,
                title,
                description: descCol !== -1 && descCol < r.length ? String(r[descCol] || '').trim() : '',
                spoc: spocCol !== -1 && spocCol < r.length ? String(r[spocCol] || '').trim() : '',
                team: teamCol !== -1 && teamCol < r.length ? String(r[teamCol] || '').trim() : '',
                status,
                statusRaw,
                lastWorked: lastWorkedCol !== -1 && lastWorkedCol < r.length ? excelDate(r[lastWorkedCol]) : '',
                expectedCompletion: expCompCol !== -1 && expCompCol < r.length ? excelDate(r[expCompCol]) : '',
                artifacts: artifactsCol !== -1 && artifactsCol < r.length ? String(r[artifactsCol] || '').trim() : '',
                comments: commentsCol !== -1 && commentsCol < r.length ? String(r[commentsCol] || '').trim() : ''
            });
        }
    }

    if (aiSheet) {
        const raw = XLSX.utils.sheet_to_json(aiSheet, { header: 1, defval: '' });
        if (raw.length > 0) {
            const headers = raw[0].map(h => String(h).trim().toLowerCase());
            const acol = (subs) => {
                return headers.findIndex(h => subs.some(s => h.includes(s)));
            };

            const amap = {
                ideaId: acol(['idea id', 'id']),
                businessFunction: acol(['business function', 'function']),
                currentProcess: acol(['current process', 'process']),
                painPoint: acol(['pain point', 'obstacle']),
                effortLevel: acol(['manual effort', 'effort']),
                frequency: acol(['frequency']),
                solution: acol(['proposed ai', 'solution']),
                benefit: acol(['expected benefit', 'benefit']),
                tools: acol(['tools', 'systems']),
                priority: acol(['priority']),
                impact: acol(['estimated impact', 'impact']),
                ideaBy: acol(['idea given by', 'given by'])
            };

            for (let i = 1; i < raw.length; i++) {
                const r = raw[i];
                const vals = {};
                let hasValue = false;
                for (const [k, idx] of Object.entries(amap)) {
                    const val = idx !== -1 && idx < r.length ? String(r[idx] || '').trim() : '';
                    vals[k] = val;
                    if (val) hasValue = true;
                }
                if (hasValue) {
                    result.aiUsecases.push(vals);
                }
            }
        }
    }

    // Build status and counts summary
    const byStatus = {};
    result.pocs.forEach(p => {
        const s = p.status || 'Unknown';
        byStatus[s] = (byStatus[s] || 0) + 1;
    });

    result.summary.total = result.pocs.length;
    result.summary.byStatus = byStatus;

    let completedCount = 0;
    let inProgressCount = 0;
    Object.entries(byStatus).forEach(([k, v]) => {
        if (k.toLowerCase().includes('complete')) {
            completedCount += v;
        }
        if (['progress', 'wip', 'ongoing'].some(w => k.toLowerCase().includes(w))) {
            inProgressCount += v;
        }
    });
    result.summary.completed = completedCount;
    result.summary.inProgress = inProgressCount;

    return result;
}

function cleanStatus(val) {
    const s = String(val || '');
    // Strip emoji/symbols from a status string, keep the words.
    const out = s.replace(/[^a-zA-Z0-9\s\/\-&]/g, '').trim();
    return out || s;
}

function excelDate(v) {
    if (!v) return '';
    if (typeof v === 'number') {
        try {
            // Excel dates start on Jan 1, 1900
            const d = new Date((v - 25569) * 86400000);
            if (!isNaN(d.getTime())) {
                return d.toISOString().slice(0, 10);
            }
        } catch (e) {
            return String(v);
        }
    }
    return String(v).slice(0, 10);
}

module.exports = { parsePoC };
