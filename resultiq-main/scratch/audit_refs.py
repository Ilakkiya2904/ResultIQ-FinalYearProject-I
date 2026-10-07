"""
audit_refs.py — Deep inspection of both reference workbooks.
Prints everything needed to verify functional correctness.
"""
import openpyxl
from openpyxl.utils import get_column_letter

SOURCE = r'C:\Users\TEST\Documents\Reference\Original ExcelSpreadsheet.xlsx'
OUTPUT = r'C:\Users\TEST\Documents\Reference\Output ExcelSpreadsheet.xlsx'

def inspect_wb(path, label):
    print(f"\n{'='*70}")
    print(f"  {label}: {path}")
    print(f"{'='*70}")
    wb = openpyxl.load_workbook(path, data_only=True)
    for sname in wb.sheetnames:
        ws = wb[sname]
        print(f"\n--- Sheet: '{sname}'  (max_row={ws.max_row}, max_col={ws.max_column}) ---")
        # Print all rows, marking formulas vs values
        for r in range(1, ws.max_row + 1):
            row_vals = []
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                v = cell.value
                row_vals.append(v)
            # Skip completely empty rows in the middle of data
            if any(v is not None for v in row_vals):
                print(f"  Row {r:3d}: {row_vals}")
    print()

def inspect_wb_raw(path, label):
    """Read with data_only=False to capture formula strings."""
    print(f"\n{'='*70}")
    print(f"  {label} [RAW FORMULAS]: {path}")
    print(f"{'='*70}")
    wb = openpyxl.load_workbook(path, data_only=False)
    for sname in wb.sheetnames:
        ws = wb[sname]
        print(f"\n--- Sheet: '{sname}' ---")
        for r in range(1, ws.max_row + 1):
            row_vals = []
            for c in range(1, ws.max_column + 1):
                v = ws.cell(row=r, column=c).value
                row_vals.append(v)
            if any(v is not None for v in row_vals):
                print(f"  Row {r:3d}: {row_vals}")

inspect_wb(SOURCE, "SOURCE (Original)")
inspect_wb(OUTPUT, "OUTPUT (Reference)")
inspect_wb_raw(OUTPUT, "OUTPUT")
