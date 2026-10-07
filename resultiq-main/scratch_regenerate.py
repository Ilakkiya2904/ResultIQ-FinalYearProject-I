import sys
sys.path.append('C:/Users/TEST/Documents/ResultIQ-new')
from services.analyser import analyse_workbook
from services.reporter import generate_report

source_filepath = 'C:/Users/TEST/Documents/ResultIQ-new/reports/VISUAL_VERIFY_source.xlsx'
out_filepath = 'C:/Users/TEST/Documents/ResultIQ-new/reports/VISUAL_VERIFY_report.xlsx'

worksheets_json = analyse_workbook(source_filepath)
sheet_name = worksheets_json[0]['name']

stats = generate_report(source_filepath, worksheets_json, sheet_name, out_filepath)
for col, s in stats['per_column_stats'].items():
    print(f"{col}: Appeared {s['appeared']}, Passed {s['pass_count']}, Failed {s['fail_count']}, Pass {s['pass_rate']:.2f}%")
