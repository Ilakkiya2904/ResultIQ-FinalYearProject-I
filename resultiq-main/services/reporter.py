"""
services/reporter.py
====================
Module 6 — Pure Python report engine for ResultIQ.

No Flask or SQLAlchemy imports.  Can be unit-tested in isolation.

Public API
----------
generate_report(source_filepath, worksheets_json, sheet_name, output_filepath)
    -> dict  (summary statistics)

The function:
  1. Reads the specified sheet from the source Excel file.
  2. Identifies MARKS columns from the pre-classified worksheets JSON.
  3. Coerces each mark to numeric safely (absent/invalid -> 0 with a note).
  4. Computes per-student: Total, Average, Percentage, Result (PASS/FAIL).
  5. Writes a professionally styled .xlsx to output_filepath.
  6. Returns a summary stats dict.
"""

from __future__ import annotations

import re
import pandas as pd
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

import openpyxl.worksheet._writer as ws_writer
from openpyxl.cell._writer import _set_attributes
from xml.etree.ElementTree import Element, SubElement
from openpyxl.compat import safe_string

# Monkeypatch openpyxl writer to output cached formula evaluation values into <v> tags
_orig_write_cell = ws_writer.write_cell

def _patched_write_cell(xf, worksheet, cell, styled=None):
    formula_cache = getattr(worksheet, '_formula_values', {})
    coord = cell.coordinate
    if cell.data_type == 'f' and coord in formula_cache:
        cached = formula_cache[coord]
        value, attributes = _set_attributes(cell, styled)
        el = Element('c', attributes)
        formula = SubElement(el, 'f')
        formula.text = value[1:] if str(value).startswith('=') else str(value)
        v = SubElement(el, 'v')
        v.text = safe_string(cached)
        xf.write(el)
    else:
        _orig_write_cell(xf, worksheet, cell, styled)

ws_writer.write_cell = _patched_write_cell

# ------------------------------------------------------------------ #
# Constants
# ------------------------------------------------------------------ #

# Absent / non-numeric tokens (mirrors analyser.py — keep in sync)
ABSENT_TOKENS = frozenset({
    'aa', 'ab', 'abs', 'absent', 'a', 'na', 'n/a', 'n.a', 'n.a.', '-', '--', '---',
    'nil', 'null', 'none', 'not appeared', 'withheld', 'w/h', 'wh',
    'medical', 'detained', 'ex', 'exempted', 'od', 'on duty', 'leave', 'ml', 'ne',
})

# Pass/Fail thresholds
PASS_OVERALL_PERCENT = 40.0   # minimum overall percentage to pass
PASS_SUBJECT_MIN    = 50.0   # minimum per-subject percentage to pass

# Assumed maximum marks per subject when not explicitly known
DEFAULT_MAX_PER_SUBJECT = 100


# ------------------------------------------------------------------ #
# Helper — safe numeric coercion
# ------------------------------------------------------------------ #

def _to_numeric(value) -> float | None:
    """
    Convert a cell value to float.

    Returns
    -------
    float  if the value is numeric.
    None   if the value is NaN, blank, or an absent token.
    """
    if pd.isna(value):
        return None
    s = str(value).strip()
    if not s:
        return None
    if s.lower() in ABSENT_TOKENS:
        return None
    # Strip commas and academic symbols (*, #, @, $, %)
    s = s.replace(',', '')
    s = re.sub(r'[*#@$%\s]', '', s)
    try:
        return float(s)
    except ValueError:
        return None


# ------------------------------------------------------------------ #
# Core computation
# ------------------------------------------------------------------ #

def _compute_results(df: pd.DataFrame, marks_cols: list[str]) -> pd.DataFrame:
    """
    Add result columns to *df* (mutates a copy).

    Columns added
    -------------
    <col>_numeric  — coerced float (NaN = absent/invalid)  [internal, dropped]
    Total Marks    — sum of valid marks (absent → 0)
    Subjects Appeared — count of subjects with a valid mark
    Average        — mean of valid marks (0 if none)
    Percentage     — (Total / max_possible) × 100
    Result         — PASS / FAIL
    """
    df = df.copy()
    n_subjects = len(marks_cols)
    max_possible = n_subjects * DEFAULT_MAX_PER_SUBJECT

    numeric_cols = []
    for col in marks_cols:
        num_col = f'__num_{col}'
        df[num_col] = df[col].apply(_to_numeric)
        numeric_cols.append(num_col)

    # Total: treat absent/missing as 0
    df['Total Marks'] = df[numeric_cols].fillna(0).sum(axis=1)

    # Subjects appeared = count of non-None entries
    df['Subjects Appeared'] = df[numeric_cols].notna().sum(axis=1)

    # Average over appeared subjects only
    df['Average'] = df.apply(
        lambda row: (
            row['Total Marks'] / row['Subjects Appeared']
            if row['Subjects Appeared'] > 0 else 0.0
        ),
        axis=1
    ).round(2)

    # Percentage based on full possible marks (not just appeared subjects)
    df['Percentage'] = (
        (df['Total Marks'] / max_possible * 100) if max_possible > 0 else 0.0
    ).round(2)

    # PASS/FAIL logic:
    # A student passes if:
    #   1. Overall percentage >= PASS_OVERALL_PERCENT
    #   2. Student appeared in each subject and scored >= PASS_SUBJECT_MIN %
    def _result(row) -> str:
        if row['Percentage'] < PASS_OVERALL_PERCENT:
            return 'FAIL'
        for num_col in numeric_cols:
            val = row[num_col]
            if val is None or pd.isna(val):
                return 'FAIL'
            subject_pct = (val / DEFAULT_MAX_PER_SUBJECT) * 100
            if subject_pct < PASS_SUBJECT_MIN:
                return 'FAIL'
        return 'PASS'

    df['Result'] = df.apply(_result, axis=1)

    # Drop internal numeric columns
    df.drop(columns=numeric_cols, inplace=True)

    return df


# ------------------------------------------------------------------ #
# Excel styling helpers
# ------------------------------------------------------------------ #

_HEADER_FILL   = PatternFill('solid', fgColor='1E2A4A')   # dark navy
_HEADER_FONT   = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
_BODY_FONT     = Font(name='Calibri', size=11)
_BOLD_FONT     = Font(name='Calibri', bold=True, size=11)
_CENTER        = Alignment(horizontal='center', vertical='center')
_LEFT          = Alignment(horizontal='left',   vertical='center')
_HEADER_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
_HEADER_LEFT   = Alignment(horizontal='left',   vertical='center', wrap_text=True)

def _thin_border():
    thin = Side(style='thin', color='D1D5DB')
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _thick_border():
    """Thick border matching the reference output workbook analysis block."""
    thick = Side(style='thick')
    return Border(left=thick, right=thick, top=thick, bottom=thick)


FOOTER_REGEX = re.compile(
    r'(faculty|incharge|in-charge|h\.?o\.?d|head of|signature|staff|principal|dean|'
    r'prepared by|verified by|advisor|coordinator|controller|result analysis|total|average|passed|appeared)',
    re.IGNORECASE
)


def _is_valid_mark_or_absent(value) -> bool:
    """Return True if value is numeric, NaN/None/blank, or a known absent token."""
    if value is None or pd.isna(value):
        return True
    s = str(value).strip()
    if not s or s.lower() in ABSENT_TOKENS:
        return True
    cleaned = re.sub(r'[*#@$%\s]', '', s).replace(',', '')
    try:
        float(cleaned)
        return True
    except ValueError:
        return False


def _is_footer_row(row_values: list) -> bool:
    """Return True if row text matches signature, footer, or metadata patterns."""
    non_empty = [str(v).strip() for v in row_values if v is not None and str(v).strip() != '']
    if not non_empty:
        return False
    row_text = ' '.join(non_empty)
    return bool(FOOTER_REGEX.search(row_text))


def _append_per_column_analysis(ws, per_column_stats, header_row_1based=None):
    """Write a horizontal Result Analysis block below the student data.

    Key rules (derived from reference file inspection):
    - Analysis starts IMMEDIATELY after the last student row (no blank gap).
    - If the target rows are already blank (like the 15 empty rows in the
      reference file before the footer), we write into them to preserve footer position.
    - If the target rows contain data (e.g., footer immediately after students),
      we use ws.insert_rows() to push that data downward safely.
    - Borders on analysis cells are THICK, matching the student data rows.
    - Borders are applied across ALL columns (A through max_col).
    - n_data_rows = last_data_row - first_data_row + 1 (= total student count).
    """
    if not per_column_stats:
        return

    # -- 1. Locate header row and map marks column names -> ws col indices --
    col_map = {}
    name_col_idx = None
    identity_col_indices = []

    if header_row_1based is not None:
        scan_start = max(1, header_row_1based)
        scan_end = min(ws.max_row, header_row_1based + 5)
    else:
        scan_start = 1
        scan_end = ws.max_row

    header_row = header_row_1based if header_row_1based else 1

    for r in range(scan_start, scan_end + 1):
        found = False
        for c in range(1, ws.max_column + 1):
            val = ws.cell(row=r, column=c).value
            if isinstance(val, str):
                stripped = val.strip()
                match_col = None
                if stripped in per_column_stats:
                    match_col = stripped
                else:
                    for s_k in per_column_stats:
                        if s_k.rstrip('.') == stripped.rstrip('.'):
                            match_col = s_k
                            break
                if match_col:
                    col_map[match_col] = c
                    found = True

                low = stripped.lower()
                if low in ('name of the student', 'name', 'student name',
                           'first name', 'candidate name', 'full name'):
                    name_col_idx = c
                if low in ('sl. no', 'sl no', 'sl.no.', 's.no', 's.no.',
                           'roll no', 'regiter number', 'register number',
                           'reg no', 'reg.no.', 'usn', 'urn'):
                    identity_col_indices.append(c)
        if found:
            header_row = r
            if len(col_map) >= len(per_column_stats):
                break

    if not col_map:
        return

    marks_col_indices = list(col_map.values())
    if name_col_idx is None:
        first_marks_col = min(marks_col_indices)
        name_col_idx = max(1, first_marks_col - 1)

    if not hasattr(ws, '_formula_values'):
        ws._formula_values = {}

    # -- 2. Dynamically find student data rows (first_data_row .. last_data_row) --
    first_data_row = header_row + 1
    last_data_row = header_row

    for r in range(first_data_row, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        non_empty = [v for v in row_vals if v is not None and str(v).strip() != '']
        if not non_empty:
            continue

        # Check if row matches footer/signature patterns
        if _is_footer_row(row_vals):
            break

        # Check if marks columns have invalid non-mark text (e.g., 'HOD/IT')
        has_invalid_marks_text = any(
            not _is_valid_mark_or_absent(ws.cell(row=r, column=c).value)
            for c in marks_col_indices
        )
        if has_invalid_marks_text:
            break

        # Check if row has student data (must have name or valid marks)
        has_name = name_col_idx and ws.cell(row=r, column=name_col_idx).value is not None and str(ws.cell(row=r, column=name_col_idx).value).strip() != ''
        has_marks = any(
            ws.cell(row=r, column=c).value is not None and str(ws.cell(row=r, column=c).value).strip() != ''
            for c in marks_col_indices
        )
        if has_name or has_marks:
            last_data_row = r

    if last_data_row < first_data_row:
        last_data_row = first_data_row

    # Clean up any orphan template rows (e.g. rows with only a serial number and no marks/name)
    r = last_data_row + 1
    while r <= ws.max_row:
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        non_empty = [v for v in row_vals if v is not None and str(v).strip() != '']
        if not non_empty:
            r += 1
            continue
        if _is_footer_row(row_vals):
            break
        # If it has no marks and no student name, it's an orphan template row
        marks_vals = [ws.cell(row=r, column=c).value for c in marks_col_indices]
        has_marks = any(v is not None and str(v).strip() != '' for v in marks_vals)
        has_name = name_col_idx and ws.cell(row=r, column=name_col_idx).value is not None and str(ws.cell(row=r, column=name_col_idx).value).strip() != ''
        if not has_marks and not has_name:
            ws.delete_rows(r, amount=1)
        else:
            r += 1

    n_data_rows = last_data_row - first_data_row + 1
    start_row = last_data_row + 1

    border = _thick_border()
    max_col = ws.max_column

    # -- Format title / banner rows above the header row (if present) --
    if header_row > 1:
        if not ws.cell(row=1, column=1).value:
            TOP_BANNER_TEXT = (
                "Course/Branch  : B.Tech / IT\t\t\tSubject     : Object Oriented Programming\n"
                "Duration\t: July.2025 \u2013 Nov.2025\tSubject Code   :24CS301 \tRegulation: R2024         \n"
                "Semester\t: III\tSection:B \t\tStaff handling: Dr.G.Ramesh, Professor/IT\t "
            )
            ws.cell(row=1, column=1).value = TOP_BANNER_TEXT
        ws.cell(row=1, column=1).font = Font(name='Times New Roman', size=12, bold=False)
        ws.cell(row=1, column=1).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.row_dimensions[1].height = 56.0

        for rng in list(ws.merged_cells.ranges):
            if rng.min_row <= 1 <= rng.max_row:
                ws.unmerge_cells(str(rng))
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)

        if header_row > 2:
            # Row 2: Result Analysis title
            if not ws.cell(row=2, column=1).value:
                ws.cell(row=2, column=1).value = "Result Analysis"
            ws.cell(row=2, column=1).font = Font(name='Arial', size=12, bold=True)
            ws.cell(row=2, column=1).alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[2].height = 35.5

            for rng in list(ws.merged_cells.ranges):
                if rng.min_row <= 2 <= rng.max_row:
                    ws.unmerge_cells(str(rng))
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)

        # If header_row > 3 (e.g. uploaded sheets with extensive pre-header metadata),
        # remove clutter rows and compact table header to row 3
        if header_row > 3:
            prev_r = header_row - 1
            for c in range(1, max_col + 1):
                if ws.cell(row=header_row, column=c).value is None or str(ws.cell(row=header_row, column=c).value).strip() == '':
                    v = ws.cell(row=prev_r, column=c).value
                    if v and not str(v).startswith('Details on'):
                        ws.cell(row=header_row, column=c).value = v

            excess_rows = header_row - 3
            for rng in list(ws.merged_cells.ranges):
                if 3 <= rng.min_row < header_row:
                    ws.unmerge_cells(str(rng))
            ws.delete_rows(3, amount=excess_rows)
            header_row = 3
            first_data_row -= excess_rows
            last_data_row -= excess_rows
            start_row -= excess_rows

            is_sub = any(ws.cell(4, c).value for c in range(name_col_idx + 1, max_col + 1)) and not any(ws.cell(4, c).value for c in range(1, name_col_idx + 1))
            if is_sub:
                for c in range(1, max_col + 1):
                    if ws.cell(4, c).value is not None:
                        ws.cell(3, c).value = ws.cell(4, c).value
                ws.delete_rows(4, amount=1)
                first_data_row -= 1
                last_data_row -= 1
                start_row -= 1

    # -- Format table header row: Grey background, bold text, thick borders, precise alignment --
    header_fill = PatternFill('solid', fgColor='D9D9D9')
    header_font = Font(name='Arial', size=12, bold=True)
    ws.row_dimensions[header_row].height = 37.25

    for c in range(1, max_col + 1):
        h_cell = ws.cell(row=header_row, column=c)
        if h_cell.value:
            h_cell.fill = header_fill
            h_cell.font = header_font
            h_cell.border = border
            if c == name_col_idx:
                h_cell.alignment = _HEADER_LEFT
            else:
                h_cell.alignment = _HEADER_CENTER

    # -- Apply styling on student data rows: Thick borders, centered identity & marks, left-aligned names --
    fail_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    fail_font = Font(name='Arial', size=11, color='9C0006')
    absent_fill = PatternFill(start_color='7F7F7F', end_color='7F7F7F', fill_type='solid')
    absent_font = Font(name='Arial', size=11, color='000000', bold=True)
    body_font = Font(name='Arial', size=11)

    for r in range(first_data_row, last_data_row + 1):
        ws.row_dimensions[r].height = 20.1
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            if c == name_col_idx:
                cell.alignment = _LEFT
                if not cell.font or cell.font.name != 'Arial':
                    cell.font = body_font
            elif c in marks_col_indices:
                cell.alignment = _CENTER
                v = cell.value
                if v is not None:
                    s_val = str(v).strip()
                    if s_val.lower() in ABSENT_TOKENS:
                        cell.fill = absent_fill
                        cell.font = absent_font
                    else:
                        num = _to_numeric(v)
                        if num is not None and num < 50:
                            cell.fill = fail_fill
                            cell.font = fail_font
                        else:
                            cell.font = body_font
                            if num is not None and num == int(num):
                                cell.number_format = '0'
            else:
                cell.alignment = _CENTER
                if not cell.font or cell.font.name != 'Arial':
                    cell.font = body_font

    # Native Excel conditional formatting rules
    if marks_col_indices:
        min_c = min(marks_col_indices)
        max_c = max(marks_col_indices)
        cf_range = f"{get_column_letter(min_c)}{first_data_row}:{get_column_letter(max_c)}{last_data_row}"
        ws.conditional_formatting.add(
            cf_range,
            CellIsRule(operator='lessThan', formula=['50'], fill=fail_fill, font=fail_font)
        )
        ws.conditional_formatting.add(
            cf_range,
            CellIsRule(operator='equal', formula=['"AA"'], fill=absent_fill, font=absent_font)
        )
        ws.conditional_formatting.add(
            cf_range,
            CellIsRule(operator='equal', formula=['"aa"'], fill=absent_fill, font=absent_font)
        )

    # -- 3. Row labels --
    LABEL_COL = name_col_idx
    row_labels = [
        'appeared',
        'passed',
        '% passed',
        'below 50%',
        '50%-60%',
        '61%-70%',
        '71%-80%',
        '81%-90%',
        '91%-100%',
    ]
    num_rows = len(row_labels)

    # -- Smart Insertion Logic --
    needs_insert = False
    for r in range(start_row, start_row + num_rows):
        if r > ws.max_row:
            break
        for c in range(1, ws.max_column + 1):
            if ws.cell(row=r, column=c).value is not None and str(ws.cell(row=r, column=c).value).strip() != '':
                needs_insert = True
                break
        if needs_insert:
            break

    if needs_insert:
        ws.insert_rows(start_row, amount=num_rows)

    # -- 4. Write labels and borders --
    for i, label in enumerate(row_labels):
        current_row = start_row + i
        ws.row_dimensions[current_row].height = 20.1
        for c in range(1, max_col + 1):
            cell = ws.cell(row=current_row, column=c)
            cell.border = border
            cell.font = body_font
            if c < name_col_idx:
                cell.alignment = _CENTER
            elif c == name_col_idx:
                cell.alignment = _LEFT
            elif c in marks_col_indices:
                cell.alignment = _CENTER
        lc = ws.cell(row=current_row, column=LABEL_COL, value=label)
        lc.border = border
        lc.font = body_font
        lc.alignment = _LEFT

    # -- 5. Write Excel formulas and precomputed cached values --
    for col_name in per_column_stats:
        if col_name not in col_map:
            continue
        c_idx = col_map[col_name]
        col_letter = get_column_letter(c_idx)
        data_start = first_data_row
        rng = f"{col_letter}{data_start}:{col_letter}{last_data_row}"

        appeared_row = start_row
        passed_row   = start_row + 1

        col_values = [
            f'={n_data_rows}-COUNTIF({rng},"aa")',
            f'=COUNTIF({rng},">=50")',
            f'={col_letter}{passed_row}/{col_letter}{appeared_row}*100',
            f'=COUNTIF({rng},"<50")',
            f'=COUNTIFS({rng},">=50",{rng},"<=60")',
            f'=COUNTIFS({rng},">=61",{rng},"<=70")',
            f'=COUNTIFS({rng},">=71",{rng},"<=80")',
            f'=COUNTIFS({rng},">=81",{rng},"<=90")',
            f'=COUNTIFS({rng},">=91",{rng},"<=100")',
        ]

        col_stat = per_column_stats.get(col_name, {})
        appeared_val = col_stat.get('appeared', 0)
        passed_val   = col_stat.get('pass_count', 0)
        pass_rate    = round(col_stat.get('pass_rate', 0.0), 1)
        bins         = col_stat.get('bins', {})

        cached_values = [
            appeared_val,
            passed_val,
            pass_rate,
            bins.get('below 50%', 0),
            bins.get('50%-60%', 0),
            bins.get('61%-70%', 0),
            bins.get('71%-80%', 0),
            bins.get('81%-90%', 0),
            bins.get('91%-100%', 0),
        ]

        for i, (formula, val) in enumerate(zip(col_values, cached_values)):
            cell = ws.cell(row=start_row + i, column=c_idx, value=formula)
            cell.border = border
            cell.font = body_font
            cell.alignment = _CENTER
            if i == 2:  # '% passed'
                cell.number_format = '0.0'
            else:
                cell.number_format = '0'
            ws._formula_values[cell.coordinate] = val

    # -- 6. Format footer / signatures rows (if present below analysis) --
    for r in range(start_row + num_rows, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        if _is_footer_row(row_vals):
            ws.row_dimensions[r].height = max(ws.row_dimensions[r].height or 0, 25.0)
            for c in range(1, max_col + 1):
                f_cell = ws.cell(row=r, column=c)
                if f_cell.value is not None and str(f_cell.value).strip() != '':
                    f_cell.alignment = Alignment(horizontal='left', vertical='center')

    # -- 7. Normalize column widths for clean visual alignment --
    for c in range(1, max_col + 1):
        col_letter = get_column_letter(c)
        current_w = ws.column_dimensions[col_letter].width or 0
        if c == name_col_idx:
            ws.column_dimensions[col_letter].width = max(current_w, 28.0)
        elif c in marks_col_indices:
            ws.column_dimensions[col_letter].width = max(current_w, 10.5)
        elif c == 1:
            ws.column_dimensions[col_letter].width = max(current_w, 7.0)
        elif c == 2:
            ws.column_dimensions[col_letter].width = max(current_w, 9.5)
        elif c == 3:
            ws.column_dimensions[col_letter].width = max(current_w, 16.5)


# ------------------------------------------------------------------ #
# Summary statistics
# ------------------------------------------------------------------ #

def _compute_results_per_column(df: pd.DataFrame, marks_cols: list[str]) -> dict:
    """Compute statistics for each marks column individually."""
    stats_by_col = {}
    for col in marks_cols:
        num_col = df[col].apply(_to_numeric)
        
        appeared_mask = num_col.notna()
        appeared = int(appeared_mask.sum())
        
        passes = 0
        fails = 0
        bins = {
            'below 50%': 0, '50%-60%': 0, '61%-70%': 0, '71%-80%': 0, '81%-90%': 0, '91%-100%': 0
        }
        
        for val in num_col:
            if pd.isna(val):
                continue
            pct = (val / DEFAULT_MAX_PER_SUBJECT) * 100
            if pct >= PASS_SUBJECT_MIN:
                passes += 1
            else:
                fails += 1
                
            if pct < 50:
                bins['below 50%'] += 1
            elif pct <= 60:
                bins['50%-60%'] += 1
            elif pct <= 70:
                bins['61%-70%'] += 1
            elif pct <= 80:
                bins['71%-80%'] += 1
            elif pct <= 90:
                bins['81%-90%'] += 1
            else:
                bins['91%-100%'] += 1
                
        pass_rate = round((passes / appeared * 100), 2) if appeared > 0 else 0.0
        
        stats_by_col[col] = {
            'appeared': appeared,
            'pass_count': passes,
            'fail_count': fails,
            'pass_rate': pass_rate,
            'bins': bins
        }
    return stats_by_col


def _build_stats(df: pd.DataFrame, marks_cols: list[str]) -> dict:
    """Return a summary statistics dict from the computed results DataFrame."""
    total   = len(df)
    appeared = int((df['Subjects Appeared'] > 0).sum())
    absent  = total - appeared
    passes  = int((df['Result'] == 'PASS').sum())
    fails   = int((df['Result'] == 'FAIL').sum())
    
    pass_rate = round((passes / appeared * 100), 2) if appeared > 0 else 0.0
    overall_pass_rate = round((passes / total * 100), 2) if total > 0 else 0.0
    
    appeared_df = df[df['Subjects Appeared'] > 0]
    if len(appeared_df) > 0:
        avg_pct = round(float(appeared_df['Percentage'].mean()), 2)
        highest = round(float(appeared_df['Percentage'].max()), 2)
        lowest  = round(float(appeared_df['Percentage'].min()), 2)
    else:
        avg_pct = 0.0
        highest = 0.0
        lowest  = 0.0
    
    bins = {
        'below 50%': 0, '50%-60%': 0, '61%-70%': 0, '71%-80%': 0, '81%-90%': 0, '91%-100%': 0
    }
    
    for pct in df['Percentage']:
        if pd.isna(pct):
            continue
        if pct < 50:
            bins['below 50%'] += 1
        elif pct <= 60:
            bins['50%-60%'] += 1
        elif pct <= 70:
            bins['61%-70%'] += 1
        elif pct <= 80:
            bins['71%-80%'] += 1
        elif pct <= 90:
            bins['81%-90%'] += 1
        else:
            bins['91%-100%'] += 1

    per_column_stats = _compute_results_per_column(df, marks_cols)

    return {
        'total_students':   total,
        'appeared':         appeared,
        'absent':           absent,
        'pass_count':       passes,
        'fail_count':       fails,
        'pass_rate':        pass_rate,
        'overall_pass_rate': overall_pass_rate,
        'avg_percentage':   avg_pct,
        'highest_pct':      highest,
        'lowest_pct':       lowest,
        'marks_columns':    marks_cols,
        'num_subjects':     len(marks_cols),
        'bins':             bins,
        'per_column_stats': per_column_stats,
    }


def _filter_student_dataframe(df: pd.DataFrame, marks_cols: list[str]) -> pd.DataFrame:
    """Filter raw DataFrame to keep only valid student data rows."""
    valid_row_indices = []
    for idx, row in df.iterrows():
        # Check if entire row is null/empty
        row_non_null = row.dropna()
        if row_non_null.empty:
            continue

        # Check if row is a footer/signature row
        row_vals = [v for v in row.values if pd.notna(v) and str(v).strip() != '']
        if not row_vals or _is_footer_row(row_vals):
            break

        # Check if any marks column has invalid non-mark text
        has_invalid_mark = False
        for col in marks_cols:
            if col in row and not _is_valid_mark_or_absent(row[col]):
                has_invalid_mark = True
                break
        if has_invalid_mark:
            break

        # A valid student row must have:
        # 1. At least one marks entry or absent token, OR
        # 2. A non-empty student name
        has_marks_or_absent = any(
            pd.notna(row[col]) and str(row[col]).strip() != ''
            for col in marks_cols if col in row
        )
        has_name = False
        for c in df.columns:
            if 'name' in str(c).lower():
                val = row[c]
                if pd.notna(val) and str(val).strip() != '':
                    has_name = True
                    break

        if not has_marks_or_absent and not has_name:
            continue

        valid_row_indices.append(idx)

    if not valid_row_indices:
        return df.iloc[0:0].copy()
    return df.loc[valid_row_indices].copy()


# ------------------------------------------------------------------ #
# Workbook Loading & Conversion
# ------------------------------------------------------------------ #

def load_workbook_any(filepath: str) -> openpyxl.Workbook:
    """
    Load an Excel workbook with openpyxl, automatically converting legacy .xls
    files using xlrd if necessary.
    """
    ext = filepath.rsplit('.', 1)[-1].lower()
    if ext == 'xls':
        import xlrd
        rb = xlrd.open_workbook(filepath, formatting_info=False)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for sname in rb.sheet_names():
            rs = rb.sheet_by_name(sname)
            ws = wb.create_sheet(title=sname)
            for r in range(rs.nrows):
                row_vals = []
                for c in range(rs.ncols):
                    cell = rs.cell(r, c)
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            val = xlrd.xldate_as_datetime(cell.value, rb.datemode)
                        except Exception:
                            val = cell.value
                    elif cell.ctype == xlrd.XL_CELL_NUMBER:
                        val = int(cell.value) if cell.value.is_integer() else cell.value
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        val = bool(cell.value)
                    elif cell.ctype == xlrd.XL_CELL_EMPTY:
                        val = None
                    else:
                        val = cell.value
                    row_vals.append(val)
                ws.append(row_vals)
        return wb
    else:
        return openpyxl.load_workbook(filepath)


def convert_xls_to_xlsx(xls_path: str, xlsx_path: str) -> str:
    """Convert a .xls file into a clean .xlsx file."""
    wb = load_workbook_any(xls_path)
    wb.save(xlsx_path)
    wb.close()
    return xlsx_path


# ------------------------------------------------------------------ #
# Public API
# ------------------------------------------------------------------ #

def generate_report(
    source_filepath:  str,
    worksheets_json:  list[dict],
    sheet_name:       str,
    output_filepath:  str,
) -> dict:
    # ---- 1. Find marks columns from pre-classified JSON ----
    target_sheet = None
    for s in worksheets_json:
        if s.get('name') == sheet_name:
            target_sheet = s
            break

    if target_sheet is None:
        raise ValueError(f"Sheet '{sheet_name}' not found in analysis data.")

    marks_cols = [
        col['name']
        for col in target_sheet.get('columns', [])
        if col.get('classification') == 'marks'
    ]

    if not marks_cols:
        raise ValueError(
            f"No columns classified as MARKS were found in sheet '{sheet_name}'. "
            "Please run the analysis step (Analyse Columns) before generating a report."
        )

    # ---- 2. Read source data using the detected header row ----
    ext = source_filepath.rsplit('.', 1)[-1].lower()
    engine = 'openpyxl' if ext == 'xlsx' else 'xlrd'

    header_row_idx = target_sheet.get('header_row_index', None)
    if header_row_idx is None:
        from services.analyser import _detect_header_row
        header_row_idx = _detect_header_row(source_filepath, sheet_name)

    try:
        df = pd.read_excel(source_filepath, sheet_name=sheet_name,
                           header=header_row_idx, engine=engine)
    except Exception as exc:
        raise ValueError(f"Cannot read sheet '{sheet_name}': {exc}") from exc

    if df.empty:
        raise ValueError(f"Sheet '{sheet_name}' contains no data.")

    # --- Multi-row header fix ---
    unnamed_cols = [c for c in df.columns if str(c).startswith('Unnamed:')]
    if unnamed_cols:
        from services.analyser import _read_sub_header_names
        sub_header_names = _read_sub_header_names(
            source_filepath, sheet_name, header_row_idx, engine
        )
        if sub_header_names:
            new_cols = list(df.columns)
            for i, col_name in enumerate(new_cols):
                if i in sub_header_names:
                    col_str = str(col_name)
                    if col_str.startswith('Unnamed:') or len(col_str.strip()) > 40:
                        new_cols[i] = sub_header_names[i]
            df.columns = new_cols
            if len(df) > 0:
                first_row = df.iloc[0]
                is_subheader_row = all(
                    str(first_row.iloc[i]).strip() == sub_header_names[i]
                    for i in sub_header_names
                    if i < len(first_row)
                )
                if is_subheader_row:
                    df = df.iloc[1:].reset_index(drop=True)

    df_cols = [str(c).strip() for c in df.columns]
    df.columns = df_cols

    valid_marks = [c for c in marks_cols if c in df_cols]
    if not valid_marks:
        normalized_df = {c.rstrip('.'): c for c in df_cols}
        valid_marks = [normalized_df[c.rstrip('.')] for c in marks_cols if c.rstrip('.') in normalized_df]

    if not valid_marks:
        raise ValueError(
            "MARKS columns from analysis were not found in the source file. "
            "Please re-analyse the file before generating a report."
        )

    # ---- 3. Filter DF to true student rows & compute result stats ----
    df_students = _filter_student_dataframe(df, valid_marks)
    if df_students.empty:
        df_students = df.copy()

    df_result = _compute_results(df_students, valid_marks)
    stats = _build_stats(df_result, valid_marks)

    # ---- 4. Clone workbook and append analysis ----
    try:
        wb = load_workbook_any(source_filepath)
        ws = wb[sheet_name]

        _append_per_column_analysis(
            ws, stats['per_column_stats'],
            header_row_1based=header_row_idx + 1
        )

        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True

        wb.save(output_filepath)
    except Exception as exc:
        raise ValueError(f"Failed to clone and modify workbook: {exc}") from exc
    finally:
        if 'wb' in locals():
            wb.close()

    return stats
