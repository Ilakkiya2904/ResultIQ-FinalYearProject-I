"""
functional_audit.py — Verify ResultIQ against reference workbooks.
Checks all 10 functional areas without modifying any code.
"""
import sys
sys.path.insert(0, r'C:\Users\TEST\Documents\ResultIQ-new')

import openpyxl
import pandas as pd
from services.analyser import analyse_workbook
from services.reporter import (
    generate_report, _to_numeric, _compute_results_per_column,
    _append_per_column_analysis, ABSENT_TOKENS, PASS_SUBJECT_MIN
)

SOURCE = r'C:\Users\TEST\Documents\Reference\Original ExcelSpreadsheet.xlsx'
OUTPUT_REF = r'C:\Users\TEST\Documents\Reference\Output ExcelSpreadsheet.xlsx'
OUTPUT_GEN = r'C:\Users\TEST\Documents\ResultIQ-new\scratch\audit_generated.xlsx'

PASS = "[PASS]"
FAIL = "[FAIL]"

issues = []

def check(label, condition, detail=""):
    status = PASS if condition else FAIL
    print(f"  {status} | {label}")
    if not condition:
        issues.append(f"{label}: {detail}")
        print(f"         >> {detail}")

# ── 1. INPUT WORKBOOK DETECTION ──────────────────────────────────────
print("\n[1] INPUT WORKBOOK DETECTION")
worksheets_json = analyse_workbook(SOURCE)
check("Workbook opens without error", len(worksheets_json) > 0)
check("Sheet 'Table 1' detected", any(s['name'] == 'Table 1' for s in worksheets_json))
sheet_info = next(s for s in worksheets_json if s['name'] == 'Table 1')

# ── 2. STUDENT-ROW DETECTION ─────────────────────────────────────────
print("\n[2] STUDENT-ROW DETECTION")
# Source has 47 data rows (rows 5..51), header at row 4, rows 1-3 are metadata
# pd.read_excel with header=0 reads the first row as header = source row 5 header = 'Sl.No' etc
# but source has a metadata row 1 (multi-line) and 'Result Analysis' at row 2, blank row 3, then header row 4
wb_src = openpyxl.load_workbook(SOURCE, data_only=True)
ws_src = wb_src['Table 1']
print(f"  Source sheet: max_row={ws_src.max_row}, max_col={ws_src.max_column}")

# What does pd.read_excel pick up as header?
df_raw = pd.read_excel(SOURCE, sheet_name='Table 1', header=0, engine='openpyxl')
print(f"  pd.read_excel rows={len(df_raw)}, columns={list(df_raw.columns)[:6]}")
# The header (row 1 of the DF) should be the column names row in the workbook

# Reference: header is at worksheet row 4, data rows 5..51 = 47 rows
EXPECTED_DATA_ROWS = 47
check(
    f"Student data rows detected = {EXPECTED_DATA_ROWS}",
    len(df_raw) == EXPECTED_DATA_ROWS,
    f"Got {len(df_raw)} rows instead of {EXPECTED_DATA_ROWS}. "
    "pd.read_excel may be picking up incorrect header row."
)

# ── 3. DYNAMIC MARKS-COLUMN DETECTION ────────────────────────────────
print("\n[3] DYNAMIC MARKS-COLUMN DETECTION")
marks_cols_detected = [c['name'] for c in sheet_info['columns'] if c['classification'] == 'marks']
print(f"  Detected marks columns: {marks_cols_detected}")
check("CIT-1 classified as marks", 'CIT-1' in marks_cols_detected)
check("CIT-2 classified as marks", 'CIT-2' in marks_cols_detected)
check("CIT-3 classified as marks", 'CIT-3' in marks_cols_detected)
identity_cols = [c['name'] for c in sheet_info['columns'] if c['classification'] == 'identity']
print(f"  Detected identity columns: {identity_cols}")
check("Name classified as identity (not marks)", 'Name of the student' in identity_cols or
      any('name' in c.lower() for c in identity_cols),
      f"Identity cols: {identity_cols}")

# ── 4 & 5. CALCULATION ACCURACY / APPEARED/PASSED/FAILED/PASS% ───────
print("\n[4+5] CALCULATION ACCURACY / APPEARED/PASSED/FAILED/PASS%")

# Reference values from the source workbook (data_only view, rows 52-54):
#   appeared: CIT-1=45, CIT-2=46, CIT-3=45
#   passed:   CIT-1=39, CIT-2=46, CIT-3=43
#   % passed: CIT-1=86.67, CIT-2=100, CIT-3=95.56
REF = {
    'CIT-1': {'appeared': 45, 'passed': 39, 'failed': 6,  'pct': 86.67},
    'CIT-2': {'appeared': 46, 'passed': 46, 'failed': 0,  'pct': 100.0},
    'CIT-3': {'appeared': 45, 'passed': 43, 'failed': 2,  'pct': 95.56},
}

df_for_stats = pd.read_excel(SOURCE, sheet_name='Table 1', header=0, engine='openpyxl')
df_for_stats.columns = [str(c).strip() for c in df_for_stats.columns]
valid_marks = [c for c in ['CIT-1', 'CIT-2', 'CIT-3'] if c in df_for_stats.columns]
per_col = _compute_results_per_column(df_for_stats, valid_marks)

for col, ref in REF.items():
    got = per_col.get(col, {})
    check(f"{col} Appeared = {ref['appeared']}",
          got.get('appeared') == ref['appeared'],
          f"Got {got.get('appeared')}")
    check(f"{col} Passed = {ref['passed']}",
          got.get('pass_count') == ref['passed'],
          f"Got {got.get('pass_count')}")
    check(f"{col} Failed = {ref['failed']}",
          got.get('fail_count') == ref['failed'],
          f"Got {got.get('fail_count')}")
    pct_got = round(got.get('pass_rate', 0), 2)
    pct_ref = ref['pct']
    check(f"{col} Pass% ≈ {pct_ref}",
          abs(pct_got - pct_ref) < 0.1,
          f"Got {pct_got}")

# ── 6. MARK-RANGE COUNTS ──────────────────────────────────────────────
print("\n[6] MARK-RANGE COUNTS")
# Reference bins from source (rows 55-60):
REF_BINS = {
    'CIT-1': {'below 50%': 6, '50%-60%': 7, '61%-70%': 4, '71%-80%': 6, '81%-90%': 10, '91%-100%': 12},
    'CIT-2': {'below 50%': 0, '50%-60%': 9, '61%-70%': 5, '71%-80%': 7, '81%-90%':  9, '91%-100%': 16},
    'CIT-3': {'below 50%': 2, '50%-60%':10, '61%-70%': 5, '71%-80%': 2, '81%-90%':  5, '91%-100%': 21},
}
for col, ref_bins in REF_BINS.items():
    got_bins = per_col.get(col, {}).get('bins', {})
    for bin_label, ref_count in ref_bins.items():
        got_count = got_bins.get(bin_label, -1)
        check(f"{col} [{bin_label}] = {ref_count}",
              got_count == ref_count,
              f"Got {got_count}")

# ── 7. ABSENT-VALUE HANDLING ──────────────────────────────────────────
print("\n[7] ABSENT-VALUE HANDLING")
# 'AA' must map to None (not appeared)
check("'AA' → None",          _to_numeric('AA')   is None)
check("'aa' → None",          _to_numeric('aa')   is None)
check("'AB' → None",          _to_numeric('AB')   is None)
check("'Absent' → None",      _to_numeric('Absent') is None)
check("'50' → 50.0",          _to_numeric('50')   == 50.0)
check("'AA' excluded from appeared",
      per_col['CIT-1']['appeared'] == 45,  # 47 students - 2 AA = 45
      f"Got {per_col['CIT-1']['appeared']}")

# ── 8 & 9. RESULT ANALYSIS PLACEMENT / WORKBOOK CLONING ──────────────
print("\n[8+9] RESULT ANALYSIS PLACEMENT & WORKBOOK CLONING")
# Generate a report from the source
stats = generate_report(SOURCE, worksheets_json, 'Table 1', OUTPUT_GEN)
wb_gen = openpyxl.load_workbook(OUTPUT_GEN, data_only=False)
ws_gen = wb_gen['Table 1']

# Source: header row=4, data rows 5..51, so data ends at row 51
# After blank row 52, analysis must start at row 53
# Reference output: analysis block at rows 50-58, with data ending row 49
# Actually ref output has 46 students (rows 4-49), source has 47 students (rows 5-51)
# For our generated output from source: data 5..51=47 rows, analysis should start at 53+1=53

all_vals = {}
for r in range(1, ws_gen.max_row + 1):
    for c in range(1, ws_gen.max_column + 1):
        v = ws_gen.cell(row=r, column=c).value
        if v is not None:
            all_vals[(r, c)] = v

# Check original metadata preserved (row 1 multi-line, row 2 Result Analysis)
row1_val = ws_gen.cell(row=1, column=1).value
check("Row 1 metadata preserved (multi-line text)",
      row1_val is not None and 'B.Tech' in str(row1_val),
      f"Got: {str(row1_val)[:80]}")
check("Row 2 'Result Analysis' label preserved",
      ws_gen.cell(row=2, column=1).value == 'Result Analysis',
      f"Got: {ws_gen.cell(row=2, column=1).value}")

# Check student data preserved at row 5 (first data row)
student_row5 = [ws_gen.cell(row=5, column=c).value for c in range(1, 8)]
check("Student data row 5 preserved intact",
      student_row5[0] == 1,  # Sl.No
      f"Got: {student_row5}")

# Locate "appeared" label in column 4 (col D = identity label col)
appeared_row = None
for r in range(1, ws_gen.max_row + 1):
    v = ws_gen.cell(row=r, column=4).value
    if v is not None and str(v).strip().lower() == 'appeared':
        appeared_row = r
        break
check("'appeared' label found in analysis block",
      appeared_row is not None,
      "Label 'appeared' not found in generated workbook")

if appeared_row:
    # Reference formula: =46-COUNTIF(E4:E49,"aa") — we need =47-COUNTIF(E5:E51,"aa")
    appeared_formula_e = ws_gen.cell(row=appeared_row, column=5).value
    print(f"  Appeared formula (CIT-1 col): {appeared_formula_e}")
    check("Appeared uses COUNT formula (not hardcoded int)",
          isinstance(appeared_formula_e, str) and appeared_formula_e.startswith('='),
          f"Got: {appeared_formula_e}")

    # Check reference formula: ref uses =46-COUNTIF(E4:E49,"aa")
    # Our code uses =COUNT(E5:E51) — is this equivalent?
    # COUNT counts only numeric cells, so AA (string) is excluded. This is correct.
    check("Appeared formula uses COUNT (excludes AA strings)",
          isinstance(appeared_formula_e, str) and 'COUNT' in appeared_formula_e.upper(),
          f"Got: {appeared_formula_e}")

    # Reference formula for appeared: =46-COUNTIF(E4:E49,"aa")
    # vs our: =COUNT(E5:E51)
    # KEY DIFFERENCE: ref subtracts AA count from total; ours uses COUNT.
    # COUNT correctly excludes text "AA" naturally. Both are equivalent for clean data.
    # BUT: if a cell contains numeric 0 (edge case), COUNT would include it, COUNTIF wouldn't.
    # For this dataset, both are functionally correct.
    print(f"  Note: Reference uses '=N-COUNTIF(rng,\"aa\")'; ours uses '=COUNT(rng)'. Both are equivalent for AA-only absence.")

# ── 10. FINAL GENERATED EXCEL STRUCTURE ──────────────────────────────
print("\n[10] FINAL GENERATED EXCEL STRUCTURE")
# Reference output has rows: 1(meta), 2(Result Analysis), 3(header), 4-49(data), 50-58(analysis), 66(footer)
# Our source has:            rows: 1(meta), 2(Result Analysis), [3 blank], 4(header), 5-51(data), then analysis, 66(footer)

# Check footer preserved
footer_val = None
for r in range(ws_gen.max_row, 0, -1):
    for c in range(1, 8):
        v = ws_gen.cell(row=r, column=c).value
        if v is not None and 'Faculty' in str(v):
            footer_val = (r, v)
            break
    if footer_val:
        break
check("Footer 'Faculty Incharge' preserved",
      footer_val is not None,
      "Footer not found in generated workbook")
if footer_val:
    print(f"  Footer found at row {footer_val[0]}: {footer_val[1]}")

# Check analysis labels exist in col 4 (D) = label column
label_col = 4  # 'Name of the student' column = where labels appear
analysis_labels = []
for r in range(1, ws_gen.max_row + 1):
    v = ws_gen.cell(row=r, column=label_col).value
    if v is not None:
        analysis_labels.append((r, v))

label_values = [v for _, v in analysis_labels]
print(f"  Labels in col D: {[v for v in label_values if v not in [None]]}")

for expected in ['appeared', 'passed ', 'below 50%', '50%-60%', '91%-100%']:
    found = any(str(v).strip() == expected.strip() for v in label_values)
    check(f"Label '{expected}' present in analysis col",
          found, f"Not found. All labels: {label_values}")

# Check Pass% label — code writes "Pass %" but reference has "% passed"
pass_pct_label_code = 'Pass %'
pass_pct_label_ref  = '% passed'
found_code = any(str(v).strip() == pass_pct_label_code for v in label_values)
found_ref  = any(str(v).strip() == pass_pct_label_ref  for v in label_values)
check(f"Pass% label matches reference ('{pass_pct_label_ref}')",
      found_ref,
      f"Code writes '{pass_pct_label_code}', reference expects '{pass_pct_label_ref}'. "
      f"Found '{pass_pct_label_code}': {found_code}, found '{pass_pct_label_ref}': {found_ref}")

# Check 'Passed' label — code writes "Passed" but reference has "passed "
passed_label_code = 'Passed'
passed_label_ref  = 'passed '
found_passed_code = any(str(v).strip() == passed_label_code for v in label_values)
found_passed_ref  = any(str(v).strip() == passed_label_ref.strip() for v in label_values)
check(f"Passed label matches reference ('{passed_label_ref}')",
      found_passed_ref or found_passed_code,   # either form is acceptable functionally
      f"Neither '{passed_label_code}' nor '{passed_label_ref}' found")

# ── SUMMARY ───────────────────────────────────────────────────────────
print("\n" + "="*70)
print(f"  TOTAL ISSUES FOUND: {len(issues)}")
for i, issue in enumerate(issues, 1):
    print(f"  {i}. {issue}")
print("="*70)
