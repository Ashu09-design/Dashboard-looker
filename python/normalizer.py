"""
Excel Normalizer for Dashboard Looker
Reads messy Excel files, detects headers smartly, fuzzy-matches columns,
and outputs clean JSON ready for Node.js consumption.

Usage:
    python normalizer.py --source team --input path/to/file.xlsx --output data/team.json
    python normalizer.py --source all --config sources.json --outdir data/
"""

import argparse
import json
import os
import sys
import re
import math
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd
import openpyxl

# ══════════════════════════════════════════════════════════════════
#  SMART EXCEL READER
# ══════════════════════════════════════════════════════════════════

ALIASES_PATH = os.path.join(os.path.dirname(__file__), 'column_aliases.json')

def load_aliases():
    with open(ALIASES_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

ALIASES = load_aliases()


def fuzzy_match(needle, haystack, threshold=0.6):
    """Find the best fuzzy match for needle in a list of haystack strings."""
    needle_lower = needle.lower().strip()
    best_score = 0
    best_match = None
    for h in haystack:
        h_lower = h.lower().strip()
        # Exact match
        if needle_lower == h_lower:
            return h
        # Contains match
        if needle_lower in h_lower or h_lower in needle_lower:
            score = 0.85
            if score > best_score:
                best_score = score
                best_match = h
                continue
        # Sequence matcher
        score = SequenceMatcher(None, needle_lower, h_lower).ratio()
        if score > best_score:
            best_score = score
            best_match = h
    if best_score >= threshold:
        return best_match
    return None


def find_header_row(df_raw, expected_headers, max_scan=30):
    """
    Scan rows to find the one that best matches expected column headers.
    Returns (header_row_index, matched_columns_dict).

    Scores each row by:
    1. How many expected headers it matches (fuzzy)
    2. How many non-empty string cells it has
    3. Penalizes rows that are mostly numeric
    """
    best_row = 0
    best_score = -1
    best_mapping = {}

    all_aliases = []
    for canonical, aliases in expected_headers.items():
        for a in aliases:
            all_aliases.append((canonical, a))

    num_rows = min(max_scan, len(df_raw))

    for idx in range(num_rows):
        row = df_raw.iloc[idx]
        row_values = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]

        if len(row_values) < 2:
            continue

        # Count how many are purely numeric — headers shouldn't be mostly numbers
        numeric_count = sum(1 for v in row_values if is_numeric(v))
        string_count = len(row_values) - numeric_count

        if string_count < 2:
            continue

        # Try to match each expected column
        mapping = {}
        matched = 0
        for canonical, alias in all_aliases:
            if canonical in mapping:
                continue
            match = fuzzy_match(alias, row_values, threshold=0.6)
            if match:
                # Find the actual column index
                for col_idx, cell in enumerate(row):
                    if pd.notna(cell) and str(cell).strip() == match:
                        mapping[canonical] = col_idx
                        matched += 1
                        break

        # Score: matched columns weighted heavily + unique string values
        unique_strings = len(set(row_values) - {''})
        score = matched * 10 + unique_strings - numeric_count

        if score > best_score:
            best_score = score
            best_row = idx
            best_mapping = mapping

    return best_row, best_mapping


def is_numeric(val):
    """Check if a string value is numeric."""
    try:
        float(str(val).replace(',', '').replace('$', '').replace('%', ''))
        return True
    except (ValueError, TypeError):
        return False


def excel_date_to_str(val):
    """Convert Excel serial date or various date formats to YYYY-MM-DD string."""
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return ''
    if isinstance(val, datetime):
        return val.strftime('%Y-%m-%d')
    if isinstance(val, (int, float)) and not math.isnan(val):
        try:
            # Excel serial date: days since 1900-01-01 (with the 1900 leap year bug)
            if 1 < val < 200000:
                base = datetime(1899, 12, 30)
                return (base + timedelta(days=int(val))).strftime('%Y-%m-%d')
        except (ValueError, OverflowError):
            pass
    s = str(val).strip()
    if not s:
        return ''
    # Try common date patterns
    for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%Y/%m/%d', '%d.%m.%Y']:
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    return s


def safe_str(val):
    """Convert value to clean string."""
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return ''
    return str(val).replace('\u200b', '').strip()


def safe_num(val):
    """Convert value to number, stripping currency/percent symbols."""
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return 0
    if isinstance(val, (int, float)):
        return 0 if math.isnan(val) else val
    s = str(val).replace('$', '').replace(',', '').replace('%', '').strip()
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0


def read_excel_sheets(file_path):
    """
    Read all sheets from an Excel file. Handles .xlsx, .xls, and .xlsb.
    Returns dict of {sheet_name: DataFrame (raw, no header inference)}.
    """
    ext = Path(file_path).suffix.lower()
    sheets = {}

    if ext == '.xlsb':
        # Use pyxlsb engine via pandas
        xls = pd.ExcelFile(file_path, engine='pyxlsb')
        for name in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=name, header=None, dtype=object)
                sheets[name] = df
            except Exception:
                continue
    else:
        # Use openpyxl for .xlsx, xlrd for .xls
        engine = 'openpyxl' if ext == '.xlsx' else None
        try:
            xls = pd.ExcelFile(file_path, engine=engine)
        except Exception:
            # Fallback: try without specifying engine
            xls = pd.ExcelFile(file_path)
        for name in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=name, header=None, dtype=object)
                sheets[name] = df
            except Exception:
                continue

    return sheets


def find_sheet(sheets, names):
    """Find a sheet by trying exact match first, then substring match."""
    # Exact match (case-insensitive)
    for n in names:
        for sn in sheets:
            if sn.strip().lower() == n.lower().strip():
                return sn, sheets[sn]
    # Substring match
    for n in names:
        for sn in sheets:
            if n.lower() in sn.lower():
                return sn, sheets[sn]
    return None, None


def extract_table(df_raw, expected_headers, max_header_scan=30):
    """
    Given a raw DataFrame and expected column aliases, find the header row,
    map columns, and return a list of dicts with canonical keys.
    Also returns diagnostic info.
    """
    header_row, col_mapping = find_header_row(df_raw, expected_headers, max_header_scan)

    diagnostics = {
        'header_row': header_row,
        'columns_found': list(col_mapping.keys()),
        'columns_missing': [k for k in expected_headers if k not in col_mapping],
        'total_expected': len(expected_headers),
        'total_found': len(col_mapping),
    }

    rows = []
    for idx in range(header_row + 1, len(df_raw)):
        row = df_raw.iloc[idx]
        record = {}
        for canonical, col_idx in col_mapping.items():
            record[canonical] = row.iloc[col_idx] if col_idx < len(row) else None
        rows.append(record)

    return rows, col_mapping, diagnostics


# ══════════════════════════════════════════════════════════════════
#  SOURCE-SPECIFIC NORMALIZERS
# ══════════════════════════════════════════════════════════════════

def normalize_team(file_path):
    """Parse Team Details Excel file."""
    sheets = read_excel_sheets(file_path)
    aliases = ALIASES.get('team', {})
    result = {
        'activeMembers': [],
        'exitResources': [],
        'roleDistribution': {},
        'designationDistribution': {},
        'totalHeadcount': 0,
        'certifications': {'ga4': {}, 'oneTrust': {}, 'google': {}, 'adobe': {}, 'tealium': {}},
        'skillMatrix': [],
        'expertise': [],
        'historicalTeamSize': {},
    }
    diagnostics = {'sheets_found': list(sheets.keys()), 'sections': {}}

    # 1. Current team — "Team details 2026" (fallback to 2024, 2023)
    team_aliases = {k: v for k, v in aliases.items() if k in ['name', 'role', 'designation', 'emp_id', 'email', 'doj', 'phone', 'birthday', 'location']}

    for sheet_name_pattern in ['Team details 2026', 'Team details 2025', 'Team details 2024', 'Team details 2023', 'Team details']:
        sn, df = find_sheet(sheets, [sheet_name_pattern])
        if df is None:
            continue

        rows, mapping, diag = extract_table(df, team_aliases)
        diagnostics['sections']['current_team'] = {**diag, 'sheet': sn}

        last_role = ''
        for r in rows:
            name = safe_str(r.get('name'))
            if not name:
                continue
            role = safe_str(r.get('role'))
            if role:
                last_role = role
            designation = safe_str(r.get('designation'))

            result['activeMembers'].append({
                'name': name,
                'role': last_role,
                'designation': designation,
                'empId': safe_str(r.get('emp_id')),
                'email': safe_str(r.get('email')),
                'doj': excel_date_to_str(r.get('doj')),
                'phone': safe_str(r.get('phone')),
                'birthday': excel_date_to_str(r.get('birthday')),
                'location': safe_str(r.get('location')),
            })

            d_key = designation or 'Unspecified'
            result['designationDistribution'][d_key] = result['designationDistribution'].get(d_key, 0) + 1
            r_key = last_role or 'Unspecified'
            result['roleDistribution'][r_key] = result['roleDistribution'].get(r_key, 0) + 1

        if result['activeMembers']:
            break  # Found data, no need to try older sheets

    # 2. Historical team sizes
    for year_label in ['2023', '2024', '2025']:
        sn, df = find_sheet(sheets, [f'Team details {year_label}'])
        if df is None:
            continue
        rows, _, _ = extract_table(df, {'name': aliases.get('name', ['Name'])})
        count = sum(1 for r in rows if safe_str(r.get('name')))
        result['historicalTeamSize'][year_label] = count
    result['historicalTeamSize']['2026'] = len(result['activeMembers'])

    # 3. Exit resources
    exit_aliases = {k: v for k, v in aliases.items() if k in ['name', 'emp_id', 'designation', 'email', 'last_working_day']}
    sn, df = find_sheet(sheets, ['Exit resources', 'Exit Resources', 'Exits'])
    if df is not None:
        rows, _, diag = extract_table(df, exit_aliases)
        diagnostics['sections']['exit_resources'] = {**diag, 'sheet': sn}
        for r in rows:
            name = safe_str(r.get('name'))
            if not name:
                continue
            result['exitResources'].append({
                'name': name,
                'empId': safe_str(r.get('emp_id')),
                'designation': safe_str(r.get('designation')),
                'email': safe_str(r.get('email')),
                'lastWorkingDay': excel_date_to_str(r.get('last_working_day')),
            })

    # 4. Skillset expertise
    sn, df = find_sheet(sheets, ['Skillset Expertise 2025', 'Skillset Expertise 2026', 'Skillset Expertise'])
    if df is not None:
        # For skill matrix, headers are dynamic (tool names), so we just find the row with "Name"
        skill_aliases = {'name': aliases.get('name', ['Name']), 'role': aliases.get('role', ['Role'])}
        header_row, mapping = find_header_row(df, skill_aliases)
        diagnostics['sections']['skillset'] = {'sheet': sn, 'header_row': header_row}

        if header_row >= 0:
            headers = [safe_str(df.iloc[header_row, c]) for c in range(len(df.columns))]
            for idx in range(header_row + 1, len(df)):
                row = df.iloc[idx]
                name = safe_str(row.iloc[0])
                if not name:
                    continue
                skills = {}
                for c in range(2, len(headers)):
                    if headers[c]:
                        skills[headers[c]] = safe_str(row.iloc[c]) if c < len(row) else ''
                result['skillMatrix'].append({
                    'name': name,
                    'role': safe_str(row.iloc[1]) if len(row) > 1 else '',
                    'skills': skills,
                })

    # 5. Certifications — GA4, OneTrust, Google, Adobe, Tealium
    cert_configs = [
        ('ga4', ['GA4 Certification'], 'GA4 Completion', 'yes'),
        ('oneTrust', ['OneTrust Certification'], 'verified', 'yes'),
        ('google', ['Google Certification'], 'verified', 'yes'),
        ('adobe', ['Adobe Certification'], 'verified', 'yes'),
        ('tealium', ['Tealium Certification'], None, None),
    ]
    for cert_key, sheet_names, check_col, check_val in cert_configs:
        sn, df = find_sheet(sheets, sheet_names)
        if df is None:
            continue

        # Find header row with "Name" or "Resources"
        name_aliases = {'name': aliases.get('name', ['Name']) + ['Resources']}
        header_row, mapping = find_header_row(df, name_aliases)
        diagnostics['sections'][f'cert_{cert_key}'] = {'sheet': sn, 'header_row': header_row}

        if header_row < 0:
            continue

        headers = [safe_str(df.iloc[header_row, c]).lower() for c in range(len(df.columns))]
        certified = 0
        total = 0

        for idx in range(header_row + 1, len(df)):
            row = df.iloc[idx]
            # Find name in first few columns
            name = ''
            for c in range(min(3, len(row))):
                val = safe_str(row.iloc[c])
                if val and val.lower() not in ('lead', 'total', 'grand total', ''):
                    name = val
                    break
            if not name:
                continue
            total += 1

            if check_col:
                # Find the check column
                check_idx = -1
                for c, h in enumerate(headers):
                    if check_col.lower() in h:
                        check_idx = c
                        break
                if check_idx >= 0 and safe_str(row.iloc[check_idx]).lower() == check_val:
                    certified += 1
            elif cert_key == 'tealium':
                # Tealium: any cert value that isn't "No" or empty
                has_cert = False
                for c in range(2, min(len(row), 5)):
                    val = safe_str(row.iloc[c]).lower()
                    if val and val != 'no':
                        has_cert = True
                        break
                if has_cert:
                    certified += 1

        result['certifications'][cert_key] = {
            'certified': certified,
            'total': total,
            'rate': round(certified / total * 100) if total > 0 else 0,
        }

    # 6. Expertise scores
    sn, df = find_sheet(sheets, ['Expertise  2024', 'Expertise 2024', 'Expertise 2025', 'Expertise'])
    if df is not None:
        # Find header row with tool names like Adobe, GA, GTM, Tealium
        tool_aliases = {
            'adobe': ['Adobe'], 'ga': ['GA', 'Google Analytics'],
            'gtm': ['GTM', 'Google Tag Manager'], 'tealium': ['Tealium'],
        }
        header_row, _ = find_header_row(df, tool_aliases)
        diagnostics['sections']['expertise'] = {'sheet': sn, 'header_row': header_row}

        if header_row >= 0:
            headers = [safe_str(df.iloc[header_row, c]) for c in range(len(df.columns))]
            for idx in range(header_row + 1, len(df)):
                row = df.iloc[idx]
                name = safe_str(row.iloc[0])
                if not name:
                    continue
                scores = {}
                for c in range(1, len(headers)):
                    if headers[c]:
                        scores[headers[c]] = safe_str(row.iloc[c])
                result['expertise'].append({'name': name, 'scores': scores})

    result['totalHeadcount'] = len(result['activeMembers'])

    return result, diagnostics


def normalize_ftr(file_path):
    """Parse FTR Tracker Excel file."""
    sheets = read_excel_sheets(file_path)
    aliases = ALIASES.get('ftr', {})
    result = {
        'accounts': [],
        'qaMetrics': [],
        'summary': {'avgRating': 0, 'totalProjects': 0, 'totalPassed': 0},
    }
    diagnostics = {'sheets_found': list(sheets.keys()), 'sections': {}}

    # Sheet 1: Account Details
    sheet_names = list(sheets.keys())
    if sheet_names:
        first_df = sheets[sheet_names[0]]
        account_aliases = {k: v for k, v in aliases.items() if k in ['engagement_name', 'qa_tracking', 'expected_projects', 'scrapping', 'web_team', 'comments']}
        rows, _, diag = extract_table(first_df, account_aliases)
        diagnostics['sections']['accounts'] = {**diag, 'sheet': sheet_names[0]}

        for r in rows:
            name = safe_str(r.get('engagement_name'))
            if not name:
                continue
            result['accounts'].append({
                'account': name,
                'qaTracking': safe_str(r.get('qa_tracking')),
                'expectedProjects': safe_str(r.get('expected_projects')),
                'scrapping': safe_str(r.get('scrapping')),
                'webTeam': safe_str(r.get('web_team')),
                'comments': safe_str(r.get('comments')),
            })

    # Remaining sheets: QA Metrics
    qa_aliases = {k: v for k, v in aliases.items() if k in [
        'project_name', 'qa_date', 'global', 'global_failed',
        'inpage', 'inpage_failed', 'form', 'form_failed', 'videos', 'videos_failed'
    ]}

    for sn in sheet_names[1:]:
        df = sheets[sn]
        rows, _, diag = extract_table(df, qa_aliases)
        diagnostics['sections'][f'qa_{sn}'] = {**diag, 'sheet': sn}

        for r in rows:
            proj = safe_str(r.get('project_name'))
            if not proj:
                continue

            g = safe_num(r.get('global'))
            gf = safe_num(r.get('global_failed'))
            ip = safe_num(r.get('inpage'))
            ipf = safe_num(r.get('inpage_failed'))
            f = safe_num(r.get('form'))
            ff = safe_num(r.get('form_failed'))
            v = safe_num(r.get('videos'))
            vf = safe_num(r.get('videos_failed'))

            total = g + ip + f + v
            failed = gf + ipf + ff + vf
            pass_rate = round(((total - failed) / total) * 10000) / 100 if total > 0 else 100

            result['qaMetrics'].append({
                'project': proj,
                'qaDate': excel_date_to_str(r.get('qa_date')),
                'global': g, 'globalFailed': gf,
                'inpage': ip, 'inpageFailed': ipf,
                'form': f, 'formFailed': ff,
                'videos': v, 'videosFailed': vf,
                'total': total, 'failed': failed, 'passRate': pass_rate,
            })

    # Summary
    if result['qaMetrics']:
        rates = [m['passRate'] for m in result['qaMetrics']]
        result['summary']['avgRating'] = round(sum(rates) / len(rates) * 100) / 100
        result['summary']['totalProjects'] = len(result['qaMetrics'])
        result['summary']['totalPassed'] = sum(1 for m in result['qaMetrics'] if m['failed'] == 0)

    # Map avgRating to accounts
    for acc in result['accounts']:
        metrics = [m for m in result['qaMetrics']
                   if acc['account'].lower().split(' ')[0] in m['project'].lower()]
        acc['avgRating'] = round(sum(m['passRate'] for m in metrics) / len(metrics) * 100) / 100 if metrics else 0
        acc['ftrPass'] = f"{sum(1 for m in metrics if m['failed'] == 0)}/{len(metrics)}"

    return result, diagnostics


def normalize_kpi(file_path):
    """Parse KPI Excel file."""
    sheets = read_excel_sheets(file_path)
    aliases = ALIASES.get('kpi', {})
    result = {
        'metrics': [],
        'kpis': [],
        'summary': {'metRate': 0, 'totalMetrics': 0, 'metCount': 0},
        'highlights': [],
        'lowlights': [],
    }
    diagnostics = {'sheets_found': list(sheets.keys()), 'sections': {}}

    # Main KPI sheet (first sheet)
    sheet_names = list(sheets.keys())
    if sheet_names:
        first_df = sheets[sheet_names[0]]
        kpi_aliases = {k: v for k, v in aliases.items() if k in [
            'account_name', 'metric_name', 'kpi_name', 'description',
            'formula', 'target', 'actual', 'status'
        ]}
        rows, _, diag = extract_table(first_df, kpi_aliases)
        diagnostics['sections']['kpi_main'] = {**diag, 'sheet': sheet_names[0]}

        for r in rows:
            account = safe_str(r.get('account_name'))
            metric = safe_str(r.get('metric_name'))
            if not account and not metric:
                continue

            target = safe_num(r.get('target'))
            actual = safe_num(r.get('actual'))
            status = safe_str(r.get('status'))
            kpi_name = safe_str(r.get('kpi_name')) or metric

            result['metrics'].append({
                'account': account,
                'metric': metric,
                'name': kpi_name,
                'description': safe_str(r.get('description')),
                'formula': safe_str(r.get('formula')),
                'target': target,
                'actual': actual,
                'targetRaw': safe_str(r.get('target')),
                'actualRaw': safe_str(r.get('actual')),
                'status': status,
            })
            result['kpis'].append({
                'account': account,
                'metric': kpi_name or metric,
                'target': target,
                'actual': actual,
                'status': status,
            })

        # Summary
        result['summary']['totalMetrics'] = len(result['metrics'])
        result['summary']['metCount'] = sum(1 for m in result['metrics']
            if any(kw in m['status'].lower() for kw in ['met', 'achieved', 'green'])
            or (m['actual'] >= m['target'] and m['target'] > 0))
        result['summary']['metRate'] = round(
            result['summary']['metCount'] / result['summary']['totalMetrics'] * 100
        ) if result['summary']['totalMetrics'] > 0 else 0

    # Sheet2: Highlights/Lowlights
    if len(sheet_names) > 1:
        second_df = sheets[sheet_names[1]]
        hl_aliases = {k: v for k, v in aliases.items() if k in ['project_manager', 'project', 'month', 'highlight', 'lowlight']}
        rows, _, diag = extract_table(second_df, hl_aliases)
        diagnostics['sections']['highlights_lowlights'] = {**diag, 'sheet': sheet_names[1]}

        current_pm = ''
        current_project = ''
        current_month = ''
        for r in rows:
            pm = safe_str(r.get('project_manager'))
            if pm:
                current_pm = pm
            proj = safe_str(r.get('project'))
            if proj:
                current_project = proj
            month = safe_str(r.get('month'))
            if month:
                current_month = month

            hl = safe_str(r.get('highlight'))
            ll = safe_str(r.get('lowlight'))
            if hl:
                result['highlights'].append({'pm': current_pm, 'project': current_project, 'month': current_month, 'highlight': hl})
            if ll:
                result['lowlights'].append({'pm': current_pm, 'project': current_project, 'month': current_month, 'lowlight': ll})

    return result, diagnostics


def normalize_leave(file_path):
    """Parse Leave Tracker (.xlsb or .xlsx)."""
    sheets = read_excel_sheets(file_path)
    result = {
        'months': [],
        'currentMonth': {'byPerson': {}, 'onLeaveToday': [], 'totalLeaves': 0},
        'allPersons': [],
    }
    diagnostics = {'sheets_found': list(sheets.keys()), 'sections': {}}

    today = datetime.now()
    # Excel serial date for today
    today_serial = (today - datetime(1899, 12, 30)).days

    month_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    leave_keywords = ['sick', 'pl', 'ul', 'floater', 'planned', 'planned/earned leave', 'sick/casual leave']

    all_persons = set()
    recent_sheets = []

    skip_keywords = ['legend', 'template', 'checklist', 'holiday', 'summary']

    for sn, df in sheets.items():
        snl = sn.lower().strip()
        if any(k in snl for k in skip_keywords):
            continue
        if len(df) < 3:
            continue

        # Find header row with "Date" and "Day"
        leave_aliases = {'date': ['Date', 'Cal Date', 'Day Date'], 'day': ['Day', 'Week Day', 'Day Name']}
        header_row, mapping = find_header_row(df, leave_aliases, max_scan=8)

        if 'date' not in mapping:
            continue

        date_col = mapping['date']
        day_col = mapping.get('day', -1)

        # Person columns: everything after date/day columns
        headers = [safe_str(df.iloc[header_row, c]) for c in range(len(df.columns))]
        person_cols = []
        for c in range(len(headers)):
            if c == date_col or c == day_col:
                continue
            if headers[c] and headers[c] not in ('', 'Date', 'Day'):
                person_cols.append((c, headers[c]))

        # Determine month/year from sheet name
        sheet_year = 2026
        month_match = re.search(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)', sn, re.IGNORECASE)
        year_match = re.search(r'(\d{2,4})', sn)
        if year_match:
            yr = year_match.group(1)
            sheet_year = 2000 + int(yr) if len(yr) == 2 else int(yr)
        month_idx = -1
        if month_match:
            month_idx = next((i for i, m in enumerate(month_order) if m.lower() == month_match.group(1).lower()), -1)
        sort_key = sheet_year * 100 + (month_idx if month_idx >= 0 else 0)

        month_data = {
            'sheet': sn, 'year': sheet_year, 'month': month_idx,
            'sortKey': sort_key, 'persons': {}, 'onLeaveToday': [],
        }

        for idx in range(header_row + 1, len(df)):
            row = df.iloc[idx]
            date_val = row.iloc[date_col] if date_col < len(row) else None
            if date_val is None or (isinstance(date_val, float) and math.isnan(date_val)):
                continue

            is_today = False
            if isinstance(date_val, (int, float)) and not math.isnan(date_val):
                is_today = abs(date_val - today_serial) < 1
            elif isinstance(date_val, datetime):
                is_today = date_val.date() == today.date()

            for pc_idx, pc_name in person_cols:
                if pc_idx >= len(row):
                    continue
                val = safe_str(row.iloc[pc_idx]).lower()
                if not val:
                    continue
                all_persons.add(pc_name)

                is_leave = any(kw in val for kw in leave_keywords)
                if not is_leave:
                    # Check if it's a small positive number (1 or 2)
                    try:
                        num_val = float(val)
                        if 0 < num_val <= 3 and len(val) <= 3:
                            is_leave = True
                    except ValueError:
                        pass

                if is_leave:
                    month_data['persons'][pc_name] = month_data['persons'].get(pc_name, 0) + 1
                    if is_today:
                        month_data['onLeaveToday'].append(pc_name)

        recent_sheets.append(month_data)
        diagnostics['sections'][sn] = {'header_row': header_row, 'persons_found': len(person_cols)}

    # Sort by most recent
    recent_sheets.sort(key=lambda m: m['sortKey'], reverse=True)
    result['months'] = [{'sheet': m['sheet'], 'year': m['year'], 'month': m['month']} for m in recent_sheets]

    if recent_sheets:
        latest = recent_sheets[0]
        result['currentMonth']['byPerson'] = latest['persons']
        result['currentMonth']['onLeaveToday'] = latest['onLeaveToday']
        result['currentMonth']['totalLeaves'] = sum(latest['persons'].values())

    result['allPersons'] = sorted(all_persons)

    return result, diagnostics


def normalize_sow(file_path):
    """Parse SOW & PO Tracker Excel file."""
    sheets = read_excel_sheets(file_path)
    aliases = ALIASES.get('sow', {})
    result = {
        'projects': [],
        'summary': {'activeProjects': 0, 'totalSOWValue': 0, 'sowReceived': 0, 'poReceived': 0},
    }
    diagnostics = {'sheets_found': list(sheets.keys()), 'sections': {}}

    for sn, df in sheets.items():
        if len(df) < 2:
            continue

        rows, _, diag = extract_table(df, aliases)
        diagnostics['sections'][sn] = {**diag, 'sheet': sn}

        # Skip sheets that matched very few expected columns (probably wrong sheet)
        if diag['total_found'] < 3:
            continue

        for r in rows:
            pm = safe_str(r.get('pm'))
            client = safe_str(r.get('client'))
            project_name = safe_str(r.get('project_name')) or safe_str(r.get('sub_project'))
            if not project_name and not client:
                continue

            sow_value = safe_num(r.get('total_sow_value'))
            sow_status = safe_str(r.get('sow_status')) or 'Unknown'
            po_status = safe_str(r.get('po_status')) or 'Unknown'
            project_status = safe_str(r.get('project_status')) or 'Active'

            result['projects'].append({
                'pm': pm,
                'client': client,
                'projectName': project_name or client,
                'subProject': safe_str(r.get('sub_project')),
                'mysheetsPID': safe_str(r.get('mysheets_pid')),
                'sowValue': sow_value,
                'deRevenueShare': safe_num(r.get('de_revenue_share')),
                'startDate': excel_date_to_str(r.get('project_start')),
                'endDate': excel_date_to_str(r.get('project_end')),
                'sowStatus': sow_status,
                'poStatus': po_status,
                'projectStatus': project_status,
                'sheet': sn,
            })

            result['summary']['totalSOWValue'] += sow_value
            if sow_status.lower() == 'received':
                result['summary']['sowReceived'] += 1
            if po_status.lower() == 'received':
                result['summary']['poReceived'] += 1

    result['summary']['activeProjects'] = len(result['projects'])

    return result, diagnostics


def normalize_governance(file_path):
    """Parse Governance Excel file."""
    sheets = read_excel_sheets(file_path)
    aliases = ALIASES.get('governance', {})
    result = {
        'highlights': [], 'lowlights': [], 'qbr': [], 'fte': [],
        'risks': [], 'audits': [],
        'summary': {'activeRisks': 0, 'totalAudits': 0},
    }
    diagnostics = {'sheets_found': list(sheets.keys()), 'sections': {}}

    # Highlights & Lowlights
    hl_aliases = {k: v for k, v in aliases.items() if k in ['project_manager', 'project', 'month', 'highlight', 'lowlight']}
    sn, df = find_sheet(sheets, ['Highlights & Low lights', 'Highlights', 'Highlights & Lowlights'])
    if df is not None:
        rows, _, diag = extract_table(df, hl_aliases)
        diagnostics['sections']['highlights'] = {**diag, 'sheet': sn}

        current_pm = ''
        current_project = ''
        current_month = ''
        for r in rows:
            pm = safe_str(r.get('project_manager'))
            if pm:
                current_pm = pm
            proj = safe_str(r.get('project'))
            if proj:
                current_project = proj
            month = safe_str(r.get('month'))
            if month:
                current_month = month

            hl = safe_str(r.get('highlight'))
            ll = safe_str(r.get('lowlight'))
            if hl:
                result['highlights'].append({'pm': current_pm, 'project': current_project, 'month': current_month, 'highlight': hl})
            if ll:
                result['lowlights'].append({'pm': current_pm, 'project': current_project, 'month': current_month, 'lowlight': ll})

    # QBR
    qbr_aliases = {k: v for k, v in aliases.items() if k in [
        'account', 'pms', 'offshore_leads', 'delivery_leads',
        'project_start_date', 'qbr_start_date', 'qbr_delivery_type', 'comments'
    ]}
    # Add comments alias if not present
    if 'comments' not in qbr_aliases:
        qbr_aliases['comments'] = ['Comments', 'Notes', 'Remarks']

    sn, df = find_sheet(sheets, ['QBR'])
    if df is not None:
        rows, _, diag = extract_table(df, qbr_aliases)
        diagnostics['sections']['qbr'] = {**diag, 'sheet': sn}

        for r in rows:
            account = safe_str(r.get('account'))
            if not account:
                continue
            result['qbr'].append({
                'account': account,
                'pms': safe_str(r.get('pms')),
                'offshoreLeads': safe_str(r.get('offshore_leads')),
                'deliveryLeads': safe_str(r.get('delivery_leads')),
                'startDate': excel_date_to_str(r.get('project_start_date')),
                'qbrStartDate': excel_date_to_str(r.get('qbr_start_date')),
                'deliveryType': safe_str(r.get('qbr_delivery_type')),
                'comments': safe_str(r.get('comments')),
            })

    # FTE
    sn, df = find_sheet(sheets, ['FTE'])
    if df is not None:
        diagnostics['sections']['fte'] = {'sheet': sn}
        if len(df) >= 3:
            # Row 0: month serial numbers, Row 1: labels, Row 2+: account data
            month_headers = []
            for c in range(1, len(df.columns)):
                val = df.iloc[0, c]
                if isinstance(val, (int, float)) and not (isinstance(val, float) and math.isnan(val)):
                    month_headers.append(excel_date_to_str(val))
                else:
                    month_headers.append(safe_str(val))

            for idx in range(2, len(df)):
                account = safe_str(df.iloc[idx, 0])
                if not account:
                    continue
                monthly_fte = {}
                for c in range(1, len(df.columns)):
                    header_idx = c - 1
                    if header_idx < len(month_headers) and month_headers[header_idx]:
                        monthly_fte[month_headers[header_idx]] = safe_num(df.iloc[idx, c])
                result['fte'].append({'account': account, 'monthlyFTE': monthly_fte})

    # Risks & Mitigation
    risk_aliases = {k: v for k, v in aliases.items() if k in [
        'month', 'week', 'project_manager', 'project', 'risks',
        'owner', 'root_cause', 'mitigation', 'impact', 'status'
    ]}
    sn, df = find_sheet(sheets, ['Risks and Mitigation', 'Risks'])
    if df is not None:
        rows, _, diag = extract_table(df, risk_aliases)
        diagnostics['sections']['risks'] = {**diag, 'sheet': sn}

        for r in rows:
            project = safe_str(r.get('project'))
            if not project:
                continue
            result['risks'].append({
                'month': safe_str(r.get('month')),
                'week': safe_str(r.get('week')),
                'pm': safe_str(r.get('project_manager')),
                'project': project,
                'description': safe_str(r.get('risks')),
                'owner': safe_str(r.get('owner')),
                'rootCause': safe_str(r.get('root_cause')),
                'mitigation': safe_str(r.get('mitigation')),
                'impact': safe_str(r.get('impact')),
                'status': safe_str(r.get('status')),
            })

        result['summary']['activeRisks'] = sum(1 for r in result['risks']
            if r['status'].lower() == 'ongoing' and r['description'].lower() != 'no risk')

    # Audits
    audit_aliases = {k: v for k, v in aliases.items() if k in [
        'month', 'project_manager', 'project', 'audit_details', 'findings', 'audit_closure_date'
    ]}
    sn, df = find_sheet(sheets, ['Audits', 'Audit'])
    if df is not None:
        rows, _, diag = extract_table(df, audit_aliases)
        diagnostics['sections']['audits'] = {**diag, 'sheet': sn}

        for r in rows:
            pm = safe_str(r.get('project_manager'))
            if not pm:
                continue
            result['audits'].append({
                'month': safe_str(r.get('month')),
                'pm': pm,
                'project': safe_str(r.get('project')),
                'details': safe_str(r.get('audit_details')),
                'findings': safe_str(r.get('findings')),
                'closureDate': excel_date_to_str(r.get('audit_closure_date')),
            })
        result['summary']['totalAudits'] = len(result['audits'])

    return result, diagnostics


# ══════════════════════════════════════════════════════════════════
#  MAIN CLI
# ══════════════════════════════════════════════════════════════════

NORMALIZERS = {
    'team': normalize_team,
    'ftr': normalize_ftr,
    'kpi': normalize_kpi,
    'leave': normalize_leave,
    'sow': normalize_sow,
    'governance': normalize_governance,
}


def run_single(source, input_path, output_path):
    """Run a single normalizer and write output JSON."""
    if source not in NORMALIZERS:
        return {'success': False, 'error': f'Unknown source: {source}'}

    if not os.path.exists(input_path):
        return {'success': False, 'error': f'File not found: {input_path}'}

    try:
        normalizer = NORMALIZERS[source]
        data, diagnostics = normalizer(input_path)

        output = {
            'success': True,
            'source': source,
            'data': data,
            'diagnostics': diagnostics,
            'timestamp': datetime.now().isoformat(),
            'input_file': input_path,
        }

        # Ensure output directory exists
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output, f, indent=2, default=str, ensure_ascii=False)

        return output
    except Exception as e:
        error_output = {
            'success': False,
            'source': source,
            'error': str(e),
            'error_type': type(e).__name__,
            'timestamp': datetime.now().isoformat(),
            'input_file': input_path,
        }
        # Still write error to output file so Node can read it
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(error_output, f, indent=2, default=str)

        return error_output


def main():
    parser = argparse.ArgumentParser(description='Excel Normalizer for Dashboard Looker')
    parser.add_argument('--source', required=True, help='Data source: team, ftr, kpi, leave, sow, governance, or all')
    parser.add_argument('--input', help='Input Excel file path (required unless --source=all)')
    parser.add_argument('--output', help='Output JSON file path (required unless --source=all)')
    parser.add_argument('--config', help='JSON config file mapping sources to input paths (for --source=all)')
    parser.add_argument('--outdir', default='data', help='Output directory (for --source=all)')

    args = parser.parse_args()

    if args.source == 'all':
        if not args.config:
            print(json.dumps({'success': False, 'error': '--config required when --source=all'}))
            sys.exit(1)

        with open(args.config, 'r') as f:
            config = json.load(f)

        results = {}
        for source, input_path in config.items():
            if source in NORMALIZERS and input_path:
                output_path = os.path.join(args.outdir, f'{source}.json')
                results[source] = run_single(source, input_path, output_path)
                status = 'OK' if results[source].get('success') else 'FAILED'
                print(f'  [{status}] {source}: {input_path}', file=sys.stderr)

        # Write combined diagnostics
        summary = {
            'success': all(r.get('success', False) for r in results.values()),
            'sources': {k: {'success': v.get('success'), 'error': v.get('error')} for k, v in results.items()},
            'timestamp': datetime.now().isoformat(),
        }
        print(json.dumps(summary))
    else:
        if not args.input or not args.output:
            print(json.dumps({'success': False, 'error': '--input and --output required'}))
            sys.exit(1)

        result = run_single(args.source, args.input, args.output)
        # Print summary to stdout (Node reads this)
        summary = {
            'success': result.get('success', False),
            'source': args.source,
            'error': result.get('error'),
        }
        if result.get('success'):
            data = result.get('data', {})
            # Add useful counts to summary
            if args.source == 'team':
                summary['counts'] = {'activeMembers': len(data.get('activeMembers', []))}
            elif args.source == 'ftr':
                summary['counts'] = {'accounts': len(data.get('accounts', [])), 'qaMetrics': len(data.get('qaMetrics', []))}
            elif args.source == 'kpi':
                summary['counts'] = {'metrics': len(data.get('metrics', [])), 'kpis': len(data.get('kpis', []))}
            elif args.source == 'leave':
                summary['counts'] = {'months': len(data.get('months', [])), 'persons': len(data.get('allPersons', []))}
            elif args.source == 'sow':
                summary['counts'] = {'projects': len(data.get('projects', []))}
            elif args.source == 'governance':
                summary['counts'] = {'risks': len(data.get('risks', [])), 'audits': len(data.get('audits', []))}

        print(json.dumps(summary))


if __name__ == '__main__':
    main()
