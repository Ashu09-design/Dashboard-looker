/**
 * Leave Tracker Parser (.xlsb)
 * Format: monthly sheets named like "Jan 26", "Feb26", "Dec 25", etc.
 * Recent sheets: Headers at R0 (Date | Day | Person1 | Person2 | ...)
 * Older sheets: Headers at R4 (Date | Day | Person1 | Person2 | ...)
 * Cell values: "Sick", "PL", "UL", "Floater", number 1/2, etc. = leave
 */
const XLSX = require('xlsx');

function parseLeaveTracker(filePath) {
    const wb = XLSX.readFile(filePath);
    const result = {
        months: [],
        currentMonth: { byPerson: {}, onLeaveToday: [], totalLeaves: 0 },
        allPersons: new Set(),
    };

    const today = new Date();
    const todaySerial = Math.floor((today.getTime() / 86400000) + 25569);

    // Find current month sheet — match latest year/month
    const monthOrder = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const recentSheets = [];

    wb.SheetNames.forEach(sn => {
        // Skip utility sheets
        const snl = sn.toLowerCase().trim();
        if (['legend', 'template', 'checklist', 'holiday'].some(k => snl.includes(k))) return;

        const ws = wb.Sheets[sn];
        if (!ws) return;
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (raw.length < 3) return;

        // Find the header row containing "Date" and "Day"
        let headerIdx = -1;
        for (let i = 0; i < Math.min(6, raw.length); i++) {
            const cells = raw[i].map(c => String(c).toLowerCase().trim());
            if (cells.includes('date') && cells.includes('day')) {
                headerIdx = i; break;
            }
        }
        if (headerIdx < 0) return;

        const headers = raw[headerIdx].map(c => String(c).trim());
        const dateCol = headers.findIndex(h => h.toLowerCase() === 'date');
        const dayCol = headers.findIndex(h => h.toLowerCase() === 'day');
        const personCols = [];
        for (let c = 2; c < headers.length; c++) {
            if (headers[c] && headers[c] !== '' && c !== dateCol && c !== dayCol) {
                personCols.push({ idx: c, name: headers[c] });
            }
        }

        // Determine month/year from sheet name or first date
        let sheetMonth = sn;
        let sheetYear = 2026;
        const monthMatch = sn.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i);
        const yearMatch = sn.match(/(\d{2,4})/);
        if (yearMatch) {
            sheetYear = yearMatch[1].length === 2 ? 2000 + parseInt(yearMatch[1]) : parseInt(yearMatch[1]);
        }
        const monthIdx = monthMatch ? monthOrder.findIndex(m => m.toLowerCase() === monthMatch[1].toLowerCase()) : -1;
        const sheetSortKey = sheetYear * 100 + (monthIdx >= 0 ? monthIdx : 0);

        const monthData = { sheet: sn, year: sheetYear, month: monthIdx, sortKey: sheetSortKey, persons: {}, onLeaveToday: [] };

        for (let i = headerIdx + 1; i < raw.length; i++) {
            const row = raw[i];
            const dateVal = row[dateCol];
            if (!dateVal) continue;

            let isToday = false;
            if (typeof dateVal === 'number') {
                isToday = Math.abs(dateVal - todaySerial) < 1;
            }

            personCols.forEach(pc => {
                const val = String(row[pc.idx] || '').trim().toLowerCase();
                if (!val) return;
                result.allPersons.add(pc.name);

                // Count as leave if: Sick, PL, UL, Floater, Planned, number 1 or 2
                const isLeave = ['sick', 'pl', 'ul', 'floater', 'planned', '1', '2',
                    'planned/earned leave', 'sick/casual leave'].some(l => val.includes(l))
                    || (val.length <= 3 && !isNaN(Number(val)) && Number(val) > 0);

                if (isLeave) {
                    monthData.persons[pc.name] = (monthData.persons[pc.name] || 0) + 1;
                    if (isToday) monthData.onLeaveToday.push(pc.name);
                }
            });
        }

        recentSheets.push(monthData);
    });

    // Sort by most recent
    recentSheets.sort((a, b) => b.sortKey - a.sortKey);
    result.months = recentSheets.map(m => ({ sheet: m.sheet, year: m.year, month: m.month }));

    // Current month = most recent sheet
    if (recentSheets.length > 0) {
        const latest = recentSheets[0];
        result.currentMonth.byPerson = latest.persons;
        result.currentMonth.onLeaveToday = latest.onLeaveToday;
        result.currentMonth.totalLeaves = Object.values(latest.persons).reduce((a, b) => a + b, 0);
    }

    result.allPersons = [...result.allPersons];
    return result;
}

module.exports = { parseLeaveTracker };
