"""
Phase 6 TEST SUITE — three test scenarios:

TEST A: Reference Original -> output should match reference Output behavior
TEST B: Different workbook (dummy with different structure)
TEST C: Another different workbook (Book2.xlsx from Documents if exists)
"""
import sys, os, shutil
sys.path.insert(0, r'c:\Users\TEST\Documents\ResultIQ-new')

import openpyxl
import pandas as pd
from services.analyser import analyse_workbook
from services.reporter import generate_report

ORIG_REF  = r'C:\Users\TEST\Documents\ResultIQ-new\Reference\Original_1 ExcelSpreadsheet.xlsx'
OUT_DIR   = r'C:\Users\TEST\Documents\ResultIQ-new\scratch'

# ─────────────────────────────────────────────────────────────────────────────
# TEST A: Reference Original -> verify analysis block matches reference output
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 60)
print("TEST A: Reference original workbook")
print("=" * 60)

ws_json = analyse_workbook(ORIG_REF)
sheet_name = ws_json[0]['name']
marks = [c['name'] for c in ws_json[0]['columns'] if c.get('classification') == 'marks']
header_idx = ws_json[0]['header_row_index']
print(f"  Sheet: {sheet_name}")
print(f"  Marks columns detected: {marks}")
print(f"  Header row index (0-based): {header_idx}")

out_a = os.path.join(OUT_DIR, 'test_a_output.xlsx')
stats_a = generate_report(ORIG_REF, ws_json, sheet_name, out_a)
print(f"  Stats: appeared={stats_a['appeared']}, pass={stats_a['pass_count']}, fail={stats_a['fail_count']}")

# Inspect output
wb_a = openpyxl.load_workbook(out_a, data_only=False)
ws_a = wb_a.active
print(f"  Output max_row={ws_a.max_row}")

# Check analysis block (should start at row 50 for this workbook)
print("  Analysis rows (checking D column and formula cols E,F,G):")
for r in range(50, 60):
    d_val = ws_a.cell(r, 4).value
    e_val = ws_a.cell(r, 5).value
    f_val = ws_a.cell(r, 6).value
    g_val = ws_a.cell(r, 7).value
    if d_val or e_val:
        print(f"    row {r}: D={repr(d_val)}, E={repr(e_val)}, F={repr(f_val)}, G={repr(g_val)}")

# Check footer still at row 66
footer_val = ws_a.cell(66, 2).value
print(f"  Footer row 66 col B: {repr(footer_val)}  (expected 'Faculty Incharge')")

# Check border style of analysis row (row 50)
e50_border = ws_a.cell(50, 5).border.left.border_style
print(f"  Border style of E50: {e50_border}  (expected 'thick')")

wb_a.close()

# ─────────────────────────────────────────────────────────────────────────────
# TEST B: Different dummy workbook with 3 students, 2 subjects, different names
# ─────────────────────────────────────────────────────────────────────────────
print()
print("=" * 60)
print("TEST B: Different workbook — 3 students, 2 assessment columns")
print("=" * 60)

# Build a fresh workbook
wb_b = openpyxl.Workbook()
ws_b = wb_b.active
ws_b.title = 'Results'

# Row 1: course info header (merged)
ws_b.merge_cells('A1:E1')
ws_b['A1'] = 'Department of Physics  |  Subject: Quantum Mechanics  |  Semester: IV  |  Section: A'
# Row 2: table header
ws_b.append(['Sl.No', 'Roll No', 'Name', 'Internal 1', 'Internal 2'])
# Rows 3-5: students
ws_b.append([1, 'PH001', 'Ananya Sharma', 78, 85])
ws_b.append([2, 'PH002', 'Rahul Verma', 'AA', 60])
ws_b.append([3, 'PH003', 'Priya Singh', 45, 70])
# Row 6: footer
ws_b.append([None, None, None, None, None])  # blank
ws_b.append([None, 'Staff Incharge', None, None, 'HOD/Physics'])

dummy_b = os.path.join(OUT_DIR, 'test_b_input.xlsx')
wb_b.save(dummy_b)
print(f"  Created dummy B at: {dummy_b}")

ws_json_b = analyse_workbook(dummy_b)
sheet_b = ws_json_b[0]['name']
marks_b = [c['name'] for c in ws_json_b[0]['columns'] if c.get('classification') == 'marks']
print(f"  Sheet: {sheet_b}")
print(f"  Marks columns detected: {marks_b}")
print(f"  Header row index: {ws_json_b[0]['header_row_index']}")

out_b = os.path.join(OUT_DIR, 'test_b_output.xlsx')
stats_b = generate_report(dummy_b, ws_json_b, sheet_b, out_b)
print(f"  Stats: appeared={stats_b['appeared']}, pass={stats_b['pass_count']}, fail={stats_b['fail_count']}")

wb_out_b = openpyxl.load_workbook(out_b, data_only=False)
ws_out_b = wb_out_b.active
print(f"  Output max_row={ws_out_b.max_row}")
print("  Analysis area:")
for r in range(1, ws_out_b.max_row + 1):
    row_vals = [ws_out_b.cell(r, c).value for c in range(1, 6)]
    if any(v is not None for v in row_vals):
        print(f"    row {r}: {row_vals}")

wb_out_b.close()

# Verify TEST B output does NOT contain TEST A data
wb_check = openpyxl.load_workbook(out_b, data_only=True)
ws_check = wb_check.active
all_text = ' '.join(
    str(ws_check.cell(r, c).value)
    for r in range(1, ws_check.max_row + 1)
    for c in range(1, ws_check.max_column + 1)
    if ws_check.cell(r, c).value is not None
)
has_ref_name = 'SWETHA' in all_text or 'CIT-1' in all_text
wb_check.close()
print(f"  TEST B output contains reference workbook data: {has_ref_name}  (expected False)")

print()
print("=" * 60)
print("ALL TESTS COMPLETE")
print("=" * 60)
