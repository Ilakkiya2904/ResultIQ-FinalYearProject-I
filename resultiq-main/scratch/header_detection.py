"""
header_detection.py — Diagnose how analyser.py reads the reference workbook,
and what the correct header row is.
"""
import sys
sys.path.insert(0, r'C:\Users\TEST\Documents\ResultIQ-new')

import openpyxl
import pandas as pd

SOURCE = r'C:\Users\TEST\Documents\Reference\Original ExcelSpreadsheet.xlsx'

print("=== Raw workbook structure (first 8 rows) ===")
wb = openpyxl.load_workbook(SOURCE, data_only=True)
ws = wb['Table 1']
for r in range(1, 9):
    row = [ws.cell(row=r, column=c).value for c in range(1, 8)]
    print(f"  Row {r}: {row}")

print("\n=== pd.read_excel with header=0 (default) ===")
df0 = pd.read_excel(SOURCE, sheet_name='Table 1', header=0, engine='openpyxl')
print(f"  Columns: {list(df0.columns)}")
print(f"  Rows: {len(df0)}")
print(f"  First few rows:\n{df0.head(5).to_string()}")

print("\n=== pd.read_excel with header=3 (row 4 = actual header) ===")
df3 = pd.read_excel(SOURCE, sheet_name='Table 1', header=3, engine='openpyxl')
print(f"  Columns: {list(df3.columns)}")
print(f"  Rows: {len(df3)}")
print(f"  First 3 rows:\n{df3.head(3).to_string()}")

print("\n=== analyser.py classify results with header=3 ===")
from services.analyser import classify_column
for col in df3.columns:
    classification, _ = classify_column(df3[col], str(col))
    print(f"  '{col}' -> {classification}")

print("\n=== What does analyse_workbook() return for this file? ===")
from services.analyser import analyse_workbook
sheets = analyse_workbook(SOURCE)
for s in sheets:
    print(f"  Sheet: {s['name']}")
    for c in s['columns']:
        print(f"    {c['name']!r:40} classification={c['classification']}")
