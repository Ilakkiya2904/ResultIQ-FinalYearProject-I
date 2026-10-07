"""Generate VISUAL_VERIFY report and verify values from stats dict."""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from services.analyser import analyse_workbook
from services.reporter import generate_report

SOURCE = r'C:\Users\TEST\Documents\ResultIQ-new\reports\VISUAL_VERIFY_source.xlsx'
GEN_OUTPUT = r'C:\Users\TEST\Documents\ResultIQ-new\reports\VISUAL_VERIFY_report.xlsx'

sheets = analyse_workbook(SOURCE)
target_sheet = sheets[0]['name']

stats = generate_report(SOURCE, sheets, target_sheet, GEN_OUTPUT)

print("\n=== FINAL VALUES FOR VISUAL VERIFY ===")
for col, cs in stats['per_column_stats'].items():
    print(f"{col}: Appeared {cs['appeared']}, Passed {cs['pass_count']}, Failed {cs['fail_count']}, Pass {cs['pass_rate']}%")
