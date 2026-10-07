import openpyxl
from openpyxl.utils import get_column_letter

orig_path = r"c:\Users\TEST\Documents\Reference\Original ExcelSpreadsheet.xlsx"
out_path = r"c:\Users\TEST\Documents\Reference\Output ExcelSpreadsheet.xlsx"

def analyze_wb(path):
    print(f"=== Analyzing {path} ===")
    wb = openpyxl.load_workbook(path, data_only=False)
    for sheet_name in wb.sheetnames:
        print(f"\nSheet: {sheet_name}")
        ws = wb[sheet_name]
        print(f"Max row: {ws.max_row}, Max col: {ws.max_column}")
        
        for row in range(1, min(ws.max_row + 1, 50)): # Print first 50 rows
            row_data = []
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                val = cell.value
                if val is not None:
                    row_data.append(f"{get_column_letter(col)}{row}: {val}")
            if row_data:
                print(" | ".join(row_data))

analyze_wb(orig_path)
print("\n" + "="*50 + "\n")
analyze_wb(out_path)
