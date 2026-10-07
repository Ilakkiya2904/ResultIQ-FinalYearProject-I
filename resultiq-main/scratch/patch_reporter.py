"""
Patch script: replaces _append_per_column_analysis in services/reporter.py
with the corrected version.
"""
import re

REPORTER_PATH = r'c:\Users\TEST\Documents\ResultIQ-new\services\reporter.py'

with open(REPORTER_PATH, 'r', encoding='utf-8') as f:
    source = f.read()

# Find the start/end of _append_per_column_analysis
start_marker = 'def _append_per_column_analysis('
# Find the next top-level 'def ' or '# ---' that follows
# Use a regex to find the function block

# Find start
start_idx = source.index(start_marker)

# Find the next top-level def or comment block after that function
# The function ends before the double newline + comment block
end_marker = '\n\n\n# --'
end_idx = source.index(end_marker, start_idx)

old_func = source[start_idx:end_idx]

new_func = '''\
def _append_per_column_analysis(ws, per_column_stats):
    """Write a horizontal Result Analysis block below the student data.

    Key rules derived from reference file inspection:
    - Analysis starts IMMEDIATELY after the last student row (no blank gap).
    - Writes into EXISTING rows — never inserts new rows (preserves footer).
    - Borders on analysis cells are THICK, matching the student data rows.
    - Borders applied across ALL columns (A through max_col).
    - n_data_rows = last_data_row - header_row (total student count for formula).
    """
    if not per_column_stats:
        return

    # -- 1. Locate header row and map marks column names -> worksheet col indices --
    header_row = 1
    col_map = {}           # marks col name -> ws column index (1-based)
    name_col_idx = None    # worksheet column index for the label text

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        found = False
        for c_idx, val in enumerate(row, start=1):
            if isinstance(val, str):
                stripped = val.strip()
                if stripped in per_column_stats:
                    col_map[stripped] = c_idx
                    found = True
                if stripped.lower() in ('name of the student', 'name', 'student name'):
                    name_col_idx = c_idx
        if found:
            header_row = row_idx
            break

    if not col_map:
        return

    # Fallback: label column is immediately left of the first marks column
    if name_col_idx is None:
        first_marks_col = min(col_map.values())
        name_col_idx = max(1, first_marks_col - 1)

    # -- 2. Find the last student data row --
    last_data_row = header_row
    for r_idx in range(ws.max_row, header_row, -1):
        if any(ws.cell(row=r_idx, column=c).value is not None
               for c in col_map.values()):
            last_data_row = r_idx
            break

    # Total students = rows between header and last data row.
    # e.g.: header_row=3, last_data_row=49  ->  n_data_rows = 46
    n_data_rows = last_data_row - header_row

    # Analysis immediately follows last student row (no blank separator).
    # Reference: row 49 = last student, row 50 = appeared.
    start_row = last_data_row + 1

    # -- 3. Row labels matching reference output exactly --
    LABEL_COL = name_col_idx
    row_labels = [
        'appeared',
        'passed ',      # trailing space present in reference
        '% passed',
        'below 50%',
        '50%-60%',
        '61%-70%',
        '71%-80%',
        '81%-90%',
        '91%-100%',
    ]

    border = _thick_border()
    max_col = ws.max_column

    # -- 4. Write into existing rows; apply thick border across ALL columns --
    for i, label in enumerate(row_labels):
        current_row = start_row + i
        for c in range(1, max_col + 1):
            cell = ws.cell(row=current_row, column=c)
            cell.border = border
            cell.font = _BODY_FONT
            cell.alignment = _CENTER
        # Label text goes in the name column, left-aligned
        lc = ws.cell(row=current_row, column=LABEL_COL, value=label)
        lc.border = border
        lc.font = _BODY_FONT
        lc.alignment = _LEFT

    # -- 5. Write each marks column's Excel formulas --
    for col_name in per_column_stats:
        if col_name not in col_map:
            continue
        c_idx = col_map[col_name]
        col_letter = get_column_letter(c_idx)
        data_start = header_row + 1
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

        for i, value in enumerate(col_values):
            cell = ws.cell(row=start_row + i, column=c_idx, value=value)
            cell.border = border
            cell.font = _BODY_FONT
            cell.alignment = _CENTER
            if i == 2:
                cell.number_format = '0.00'\
'''

patched = source[:start_idx] + new_func + source[end_idx:]

with open(REPORTER_PATH, 'w', encoding='utf-8') as f:
    f.write(patched)

print(f'Patched OK. Old func length={len(old_func)}, new func length={len(new_func)}')
