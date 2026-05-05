/**
 * Team Details Parser — READS ALL SHEETS
 * Sheets: "Team details 2026" (current team), "Team details 2024/2023" (historical),
 *         "Skillset Expertise 2025/Skillset Expertise" (skills matrix),
 *         "OneTrust/Google/Adobe/Tealium/GA4 Certification" (cert data),
 *         "Expertise/Expertise 2024" (expertise scores), "Exit resources" (attrition)
 */
const XLSX = require('xlsx');

function parseTeamDetails(filePath) {
    const wb = XLSX.readFile(filePath);
    const result = {
        activeMembers: [],
        exitResources: [],
        roleDistribution: {},
        designationDistribution: {},
        totalHeadcount: 0,
        certifications: { ga4: {}, oneTrust: {}, google: {}, adobe: {}, tealium: {} },
        skillMatrix: [],
        expertise: [],
        historicalTeamSize: {},
    };

    // ═══════════════════════════════════════════════
    // 1. CURRENT TEAM — "Team details 2026"
    // Headers R0: Role | Name | Emp ID | Designation | Email-ID | DOJ | Phone No | Birthday | Location
    // ═══════════════════════════════════════════════
    const teamSheet2026 = findSheet(wb, ['Team details 2026']);
    if (teamSheet2026) {
        const rows = XLSX.utils.sheet_to_json(teamSheet2026, { defval: '' });
        let lastRole = '';
        rows.forEach(r => {
            const name = String(r['Name'] || '').trim();
            if (!name) return;
            const role = String(r['Role'] || '').trim();
            if (role) lastRole = role;
            const designation = String(r['Designation'] || '').trim();
            result.activeMembers.push({
                name, role: lastRole, designation,
                empId: String(r['Emp ID'] || ''),
                email: String(r['Email-ID'] || ''),
                doj: excelDate(r['DOJ']),
                phone: String(r['Phone No:'] || r['Phone No'] || ''),
                birthday: excelDate(r['Birthday (Month/date)']),
                location: String(r['Current residing Location'] || ''),
            });
            const dKey = designation || 'Unspecified';
            result.designationDistribution[dKey] = (result.designationDistribution[dKey] || 0) + 1;
            const rKey = lastRole || 'Unspecified';
            result.roleDistribution[rKey] = (result.roleDistribution[rKey] || 0) + 1;
        });
    }

    // Fallback to older sheets if 2026 not found
    if (result.activeMembers.length === 0) {
        const fallback = findSheet(wb, ['Team details 2024', 'Team details 2023', 'Team details']);
        if (fallback) {
            const raw = XLSX.utils.sheet_to_json(fallback, { header: 1, defval: '' });
            const headerIdx = raw.findIndex(r => r.some(c => String(c).trim().toLowerCase() === 'name'));
            if (headerIdx >= 0) {
                const headers = raw[headerIdx].map(h => String(h).trim());
                const ci = name => headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
                for (let i = headerIdx + 1; i < raw.length; i++) {
                    const r = raw[i];
                    const name = String(r[ci('Name')] || '').trim();
                    if (!name) continue;
                    const desig = String(r[ci('Designation')] || '').trim();
                    result.activeMembers.push({
                        name, role: desig, designation: desig,
                        empId: String(r[ci('Emp ID')] || ''), email: String(r[ci('Email-ID')] || ''),
                        doj: excelDate(r[ci('DOJ')]), phone: String(r[ci('Phone No')] || ''),
                        birthday: '', location: ''
                    });
                    result.designationDistribution[desig || 'Unspecified'] = (result.designationDistribution[desig || 'Unspecified'] || 0) + 1;
                }
            }
        }
    }

    // ═══════════════════════════════════════════════
    // 2. HISTORICAL TEAM SIZES (2023, 2024)
    // ═══════════════════════════════════════════════
    ['Team details 2023', 'Team details 2024'].forEach(sheetName => {
        const ws = findSheet(wb, [sheetName]);
        if (!ws) return;
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        const headerIdx = raw.findIndex(r => r.some(c => String(c).trim().toLowerCase() === 'name'));
        if (headerIdx < 0) return;
        let count = 0;
        for (let i = headerIdx + 1; i < raw.length; i++) {
            if (String(raw[i][raw[headerIdx].findIndex(h => String(h).trim().toLowerCase() === 'name')] || '').trim()) count++;
        }
        const year = sheetName.match(/\d+/)?.[0] || '';
        result.historicalTeamSize[year] = count;
    });
    result.historicalTeamSize['2026'] = result.activeMembers.length;

    // ═══════════════════════════════════════════════
    // 3. EXIT RESOURCES
    // Headers R0: Name | Emp ID | Designation | Email-ID | DOJ | Phone No | Birthday | Last Working Day
    // ═══════════════════════════════════════════════
    const exitSheet = findSheet(wb, ['Exit resources']);
    if (exitSheet) {
        const rows = XLSX.utils.sheet_to_json(exitSheet, { defval: '' });
        rows.forEach(r => {
            const name = String(r['Name'] || '').trim();
            if (!name) return;
            result.exitResources.push({
                name, empId: String(r['Emp ID'] || ''), designation: String(r['Designation'] || ''),
                email: String(r['Email-ID'] || ''), lastWorkingDay: excelDate(r['Last Working Day']),
            });
        });
    }

    // ═══════════════════════════════════════════════
    // 4. SKILLSET EXPERTISE 2025 / Skillset Expertise
    // Headers at R4/R5: Name | Role | Verified | Launch | GTM | Tealium iQ | ...
    // ═══════════════════════════════════════════════
    const skillSheet = findSheet(wb, ['Skillset Expertise 2025', 'Skillset Expertise']);
    if (skillSheet) {
        const raw = XLSX.utils.sheet_to_json(skillSheet, { header: 1, defval: '' });
        const headerIdx = raw.findIndex(r => r.some(c => String(c).toLowerCase().includes('name')));
        if (headerIdx >= 0) {
            const headers = raw[headerIdx].map(h => String(h).replace(':', '').trim());
            for (let i = headerIdx + 1; i < raw.length; i++) {
                const r = raw[i];
                const name = String(r[0] || '').trim();
                if (!name) continue;
                const skills = {};
                for (let c = 2; c < headers.length; c++) {
                    if (headers[c]) skills[headers[c]] = String(r[c] || '').trim();
                }
                result.skillMatrix.push({ name, role: String(r[1] || '').trim(), skills });
            }
        }
    }

    // ═══════════════════════════════════════════════
    // 5. GA4 CERTIFICATION
    // Headers R0: Name | GA4 Completion | GA4 Certification link | Expiration | Comments
    // ═══════════════════════════════════════════════
    const ga4Sheet = findSheet(wb, ['GA4 Certification']);
    if (ga4Sheet) {
        const rows = XLSX.utils.sheet_to_json(ga4Sheet, { defval: '' });
        let certified = 0, total = 0;
        rows.forEach(r => {
            const name = String(r['Name'] || '').trim();
            if (!name) return;
            total++;
            if (String(r['GA4 Completion'] || '').toLowerCase() === 'yes') certified++;
        });
        result.certifications.ga4 = { certified, total, rate: total > 0 ? Math.round(certified / total * 100) : 0 };
    }

    // ═══════════════════════════════════════════════
    // 6. ONETRUST CERTIFICATION
    // Headers at R3: Resources | Expertise | Cookie Consent | Consent & Preference | Data Privacy Professional
    // ═══════════════════════════════════════════════
    const otSheet = findSheet(wb, ['OneTrust Certification']);
    if (otSheet) {
        const raw = XLSX.utils.sheet_to_json(otSheet, { header: 1, defval: '' });
        const headerIdx = raw.findIndex(r => r.some(c => String(c).toLowerCase().includes('resources')));
        if (headerIdx >= 0) {
            const headers = raw[headerIdx].map(h => String(h).trim());
            let certified = 0, total = 0;
            for (let i = headerIdx + 1; i < raw.length; i++) {
                const name = String(raw[i][1] || '').trim(); // Name in col B
                if (!name) continue;
                total++;
                // Check "Verified" column
                const verifiedCol = headers.findIndex(h => h.toLowerCase().includes('verified'));
                if (verifiedCol >= 0 && String(raw[i][verifiedCol] || '').toLowerCase() === 'yes') certified++;
            }
            result.certifications.oneTrust = { certified, total, rate: total > 0 ? Math.round(certified / total * 100) : 0 };
        }
    }

    // ═══════════════════════════════════════════════
    // 7. GOOGLE CERTIFICATION
    // Headers at R2: Resources | Certifications > GA | GTM | GDS | Verified
    // ═══════════════════════════════════════════════
    const googleSheet = findSheet(wb, ['Google Certification']);
    if (googleSheet) {
        const raw = XLSX.utils.sheet_to_json(googleSheet, { header: 1, defval: '' });
        const headerIdx = raw.findIndex(r => r.some(c => String(c).toLowerCase().includes('resources')));
        if (headerIdx >= 0) {
            let certified = 0, total = 0;
            const verCol = raw[headerIdx].findIndex(h => String(h).toLowerCase().includes('verified'));
            for (let i = headerIdx + 1; i < raw.length; i++) {
                const name = String(raw[i][0] || raw[i][1] || '').trim();
                if (!name || name.toLowerCase() === 'lead') continue;
                total++;
                if (verCol >= 0 && String(raw[i][verCol] || '').toLowerCase() === 'yes') certified++;
            }
            result.certifications.google = { certified, total, rate: total > 0 ? Math.round(certified / total * 100) : 0 };
        }
    }

    // ═══════════════════════════════════════════════
    // 8. TEALIUM CERTIFICATION
    // Headers at R1: Resources | Tealium CDP | Tealium IQ | Verified
    // ═══════════════════════════════════════════════
    const tealSheet = findSheet(wb, ['Tealium Certification']);
    if (tealSheet) {
        const raw = XLSX.utils.sheet_to_json(tealSheet, { header: 1, defval: '' });
        const headerIdx = raw.findIndex(r => r.some(c => String(c).toLowerCase().includes('resources')));
        if (headerIdx >= 0) {
            let certified = 0, total = 0;
            for (let i = headerIdx + 1; i < raw.length; i++) {
                const name = String(raw[i][1] || raw[i][0] || '').trim();
                if (!name || name.toLowerCase() === 'lead') continue;
                total++;
                // Check for any cert value that isn't "No"
                const hasAnyCert = raw[i].slice(2, 4).some(c => String(c).trim() && String(c).toLowerCase() !== 'no');
                if (hasAnyCert) certified++;
            }
            result.certifications.tealium = { certified, total, rate: total > 0 ? Math.round(certified / total * 100) : 0 };
        }
    }

    // ═══════════════════════════════════════════════
    // 9. EXPERTISE SCORES (Expertise 2024, Expertise)
    // Headers at R3/R0: Adobe | GA | GTM | Tealium | Matomo | Verified
    // ═══════════════════════════════════════════════
    const expSheet = findSheet(wb, ['Expertise  2024', 'Expertise']);
    if (expSheet) {
        const raw = XLSX.utils.sheet_to_json(expSheet, { header: 1, defval: '' });
        const headerIdx = raw.findIndex(r => r.some(c => ['adobe', 'ga', 'gtm', 'tealium'].includes(String(c).toLowerCase().trim())));
        if (headerIdx >= 0) {
            const headers = raw[headerIdx].map(h => String(h).trim());
            for (let i = headerIdx + 1; i < raw.length; i++) {
                const name = String(raw[i][0] || '').trim();
                if (!name) continue;
                const scores = {};
                headers.forEach((h, idx) => {
                    if (h && idx > 0) scores[h] = String(raw[i][idx] || '').trim();
                });
                result.expertise.push({ name, scores });
            }
        }
    }

    result.totalHeadcount = result.activeMembers.length;
    return result;
}

function findSheet(wb, names) {
    for (const n of names) {
        const found = wb.SheetNames.find(s => s.trim().toLowerCase() === n.toLowerCase().trim());
        if (found) return wb.Sheets[found];
    }
    for (const n of names) {
        const found = wb.SheetNames.find(s => s.toLowerCase().includes(n.toLowerCase()));
        if (found) return wb.Sheets[found];
    }
    return null;
}

function excelDate(v) {
    if (!v) return '';
    if (typeof v === 'number') return new Date((v - 25569) * 86400000).toISOString().slice(0, 10);
    return String(v);
}

module.exports = { parseTeamDetails };
