import openpyxl
from openpyxl.utils import get_column_letter

out_path = r"c:\Users\TEST\Documents\Reference\Output ExcelSpreadsheet.xlsx"

def analyze_wb(path):
    print(f"=== Analyzing {path} ===")
    wb = openpyxl.load_workbook(path, data_only=False)
    for sheet_name in wb.sheetnames:
        print(f"\nSheet: {sheet_name}")
        ws = wb[sheet_name]
        
        print("\n--- Formulas (if any) ---")
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    print(f"{get_column_letter(col)}{row}: {cell.value}")
        
        print("\n--- Result Analysis rows (last 20 rows) ---")
        start_row = max(1, ws.max_row - 20)
        for row in range(start_row, ws.max_row + 1):
            row_data = []
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                val = cell.value
                if val is not None:
                    row_data.append(f"{get_column_letter(col)}{row}: {val}")
            if row_data:
                print(" | ".join(row_data))

analyze_wb(out_path)
