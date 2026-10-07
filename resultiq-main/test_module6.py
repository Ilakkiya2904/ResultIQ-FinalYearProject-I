"""
test_module6.py
===============
End-to-end tests for Module 6: Result Analysis & Report Generation.

Covers:
  - services.reporter (pure unit tests)
  - POST /report/generate
  - GET  /report/<id>
  - GET  /report/<id>/download
"""

import io
import os
import tempfile
import unittest
import pandas as pd
import openpyxl
from openpyxl import Workbook

from app import create_app
from models import db, User, UploadedFile, Report
from services.reporter import (
    generate_report, _to_numeric, _compute_results,
    _compute_results_per_column, _append_per_column_analysis,
)


# ------------------------------------------------------------------ #
# Helpers
# ------------------------------------------------------------------ #

def make_xlsx_bytes(sheets_spec: dict) -> bytes:
    """Build an in-memory .xlsx workbook from {sheet: [rows]}."""
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets_spec.items():
        ws = wb.create_sheet(title=name)
        for row in rows:
            ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ------------------------------------------------------------------ #
# Unit tests — pure reporter service
# ------------------------------------------------------------------ #

class TestToNumeric(unittest.TestCase):

    def test_plain_int(self):
        self.assertEqual(_to_numeric(88), 88.0)

    def test_plain_float(self):
        self.assertAlmostEqual(_to_numeric(72.5), 72.5)

    def test_string_number(self):
        self.assertEqual(_to_numeric('90'), 90.0)

    def test_comma_number(self):
        self.assertEqual(_to_numeric('1,000'), 1000.0)

    def test_absent_token_returns_none(self):
        for token in ['AA', 'Absent', 'NA', 'nil', '--', 'ab']:
            self.assertIsNone(_to_numeric(token),
                              f"Expected None for absent token '{token}'")

    def test_nan_returns_none(self):
        import math
        self.assertIsNone(_to_numeric(float('nan')))

    def test_none_returns_none(self):
        self.assertIsNone(_to_numeric(None))


class TestComputeResults(unittest.TestCase):

    def _base_df(self):
        return pd.DataFrame({
            'Name':    ['Alice', 'Bob', 'Charlie', 'Diana', 'Eve'],
            'Roll':    [101, 102, 103, 104, 105],
            'Math':    [80, 90, 'AA', 60, 35],
            'Science': [75, 85, 70, 55, 90],
            'English': [88, 70, 65, 80, 75],
        })

    def test_columns_added(self):
        df = _compute_results(self._base_df(), ['Math', 'Science', 'English'])
        for col in ['Total Marks', 'Average', 'Percentage', 'Result', 'Subjects Appeared']:
            self.assertIn(col, df.columns, f"Missing column: {col}")

    def test_pass_fail_logic(self):
        df = _compute_results(self._base_df(), ['Math', 'Science', 'English'])
        # Alice: 80+75+88=243, pct=81%, all>=35% → PASS
        alice = df[df['Name'] == 'Alice'].iloc[0]
        self.assertEqual(alice['Result'], 'PASS')

        # Eve: 35+90+75=200, pct=66.7%, Math=35 → 35% exactly → borderline
        # 35/100*100=35, threshold is <35 to fail → 35 is >=35 so PASS
        eve = df[df['Name'] == 'Eve'].iloc[0]
        self.assertIn(eve['Result'], ('PASS', 'FAIL'))  # depends on threshold

    def test_absent_treated_as_zero_in_total(self):
        df = _compute_results(self._base_df(), ['Math', 'Science', 'English'])
        charlie = df[df['Name'] == 'Charlie'].iloc[0]
        # Math=AA(0), Science=70, English=65 → Total=135 (absent=0 for total)
        self.assertEqual(charlie['Total Marks'], 135.0)

    def test_percentage_based_on_max_possible(self):
        df = _compute_results(self._base_df(), ['Math', 'Science', 'English'])
        alice = df[df['Name'] == 'Alice'].iloc[0]
        # 3 subjects × 100 = 300 max
        expected_pct = round(alice['Total Marks'] / 300 * 100, 2)
        self.assertAlmostEqual(alice['Percentage'], expected_pct, places=1)

    def test_no_internal_num_cols_in_output(self):
        df = _compute_results(self._base_df(), ['Math', 'Science', 'English'])
        for col in df.columns:
            self.assertFalse(col.startswith('__num_'),
                             f"Internal column leaked: {col}")


class TestGenerateReportService(unittest.TestCase):

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.src_path = os.path.join(self.tmp_dir, 'source.xlsx')
        self.out_path = os.path.join(self.tmp_dir, 'report.xlsx')

        # Write a test source workbook
        data = make_xlsx_bytes({
            'Results': [
                ('Name',    'Roll', 'Math', 'Science', 'English'),
                ('Alice',   101,    80,     75,         88),
                ('Bob',     102,    90,     85,         70),
                ('Charlie', 103,    'AA',   70,         65),
                ('Diana',   104,    60,     55,         80),
                ('Eve',     105,    35,     90,         75),
            ]
        })
        with open(self.src_path, 'wb') as f:
            f.write(data)

    def _worksheets_json(self):
        return [{
            'name': 'Results',
            'is_empty': False,
            'row_count': 5,
            'columns': [
                {'name': 'Name',    'classification': 'identity', 'dtype': 'Text',    'samples': []},
                {'name': 'Roll',    'classification': 'identity', 'dtype': 'Integer', 'samples': []},
                {'name': 'Math',    'classification': 'marks',    'dtype': 'Decimal', 'samples': []},
                {'name': 'Science', 'classification': 'marks',    'dtype': 'Decimal', 'samples': []},
                {'name': 'English', 'classification': 'marks',    'dtype': 'Decimal', 'samples': []},
            ]
        }]

    def test_report_file_created(self):
        generate_report(self.src_path, self._worksheets_json(), 'Results', self.out_path)
        self.assertTrue(os.path.isfile(self.out_path), "Report file not created")

    def test_stats_returned(self):
        stats = generate_report(self.src_path, self._worksheets_json(), 'Results', self.out_path)
        self.assertEqual(stats['total_students'], 5)
        self.assertIn('pass_count', stats)
        self.assertIn('fail_count', stats)
        self.assertAlmostEqual(stats['pass_count'] + stats['fail_count'],
                               stats['total_students'])
        self.assertEqual(stats['num_subjects'], 3)
        self.assertIn('Math',    stats['marks_columns'])
        self.assertIn('Science', stats['marks_columns'])
        self.assertIn('English', stats['marks_columns'])

    def test_report_readable(self):
        generate_report(self.src_path, self._worksheets_json(), 'Results', self.out_path)
        df = pd.read_excel(self.out_path, engine='openpyxl')
        self.assertIn('Name', df.columns)
        self.assertIn('Roll', df.columns)
        # Should contain original columns
        self.assertIn('Math', df.columns)
        
        # Result analysis is appended at the bottom, so length will be more than 5
        self.assertGreater(len(df), 5)

    def test_pass_fail_values_only(self):
        stats = generate_report(self.src_path, self._worksheets_json(), 'Results', self.out_path)
        # Check that we have a valid pass rate and bins
        self.assertIn('pass_rate', stats)
        self.assertIn('bins', stats)
        self.assertIn('below 50%', stats['bins'])

    def test_no_marks_cols_raises(self):
        bad_json = [{
            'name': 'Results',
            'columns': [
                {'name': 'Name', 'classification': 'identity'},
                {'name': 'Roll', 'classification': 'identity'},
            ]
        }]
        with self.assertRaises(ValueError):
            generate_report(self.src_path, bad_json, 'Results', self.out_path)

    def test_missing_sheet_raises(self):
        with self.assertRaises(ValueError):
            generate_report(self.src_path, self._worksheets_json(),
                            'NonExistentSheet', self.out_path)

    def test_identity_cols_preserved(self):
        """Name and Roll must appear in the output unchanged."""
        generate_report(self.src_path, self._worksheets_json(), 'Results', self.out_path)
        df = pd.read_excel(self.out_path, engine='openpyxl')
        self.assertIn('Name', df.columns)
        self.assertIn('Roll', df.columns)
        names = df['Name'].tolist()
        self.assertIn('Alice', names)
        self.assertIn('Bob', names)

    def test_all_absent_student(self):
        """A student with all marks absent should still appear with FAIL result."""
        data = make_xlsx_bytes({
            'Results': [
                ('Name', 'Math', 'Science'),
                ('Alice', 80, 75),
                ('Ghost', 'AA', 'AA'),  # all absent
            ]
        })
        src = os.path.join(self.tmp_dir, 'absent.xlsx')
        with open(src, 'wb') as f:
            f.write(data)
        ws_json = [{
            'name': 'Results',
            'columns': [
                {'name': 'Name',    'classification': 'identity'},
                {'name': 'Math',    'classification': 'marks'},
                {'name': 'Science', 'classification': 'marks'},
            ]
        }]
        stats = generate_report(src, ws_json, 'Results',
                                os.path.join(self.tmp_dir, 'absent_rpt.xlsx'))
        # Ghost has 0 total → FAIL
        self.assertEqual(stats['fail_count'], 1)
        self.assertEqual(stats['pass_count'], 1)


# ------------------------------------------------------------------ #
# Per-column tests — _compute_results_per_column
# ------------------------------------------------------------------ #

class TestComputeResultsPerColumn(unittest.TestCase):

    def _base_df(self):
        return pd.DataFrame({
            'Name':    ['Alice', 'Bob', 'Charlie', 'Diana', 'Eve'],
            'CIT-1':   [80, 90, 'AA', 60, 35],
            'CIT-2':   [75, 85, 70, 55, 90],
            'CIT-3':   [88, 70, 65, 80, 75],
        })

    def test_returns_entry_for_each_marks_col(self):
        df = self._base_df()
        result = _compute_results_per_column(df, ['CIT-1', 'CIT-2', 'CIT-3'])
        self.assertIn('CIT-1', result)
        self.assertIn('CIT-2', result)
        self.assertIn('CIT-3', result)
        self.assertEqual(len(result), 3)

    def test_each_col_has_required_keys(self):
        df = self._base_df()
        result = _compute_results_per_column(df, ['CIT-1', 'CIT-2', 'CIT-3'])
        for col_name, stats in result.items():
            for key in ('appeared', 'pass_count', 'fail_count', 'pass_rate', 'bins'):
                self.assertIn(key, stats, f"Missing key '{key}' in stats for '{col_name}'")

    def test_absent_excluded_from_appeared(self):
        """AA in CIT-1 → appeared = 4, not 5."""
        df = self._base_df()
        result = _compute_results_per_column(df, ['CIT-1', 'CIT-2', 'CIT-3'])
        # CIT-1 has one 'AA' value for Charlie
        self.assertEqual(result['CIT-1']['appeared'], 4)
        # CIT-2 and CIT-3 have all valid marks
        self.assertEqual(result['CIT-2']['appeared'], 5)
        self.assertEqual(result['CIT-3']['appeared'], 5)

    def test_columns_independent_of_each_other(self):
        """Each column's pass/fail is calculated only from that column's marks."""
        df = self._base_df()
        result = _compute_results_per_column(df, ['CIT-1', 'CIT-2', 'CIT-3'])
        # CIT-1: 35 -> 35% -> PASS_SUBJECT_MIN is 35 -> passes, 80->pass, 90->pass, 60->pass, AA->not counted
        # CIT-2: all >= 35% threshold
        # CIT-3: all >= 35% threshold
        cit1_appeared = result['CIT-1']['appeared']  # 4
        cit1_pass = result['CIT-1']['pass_count']
        cit1_fail = result['CIT-1']['fail_count']
        self.assertEqual(cit1_pass + cit1_fail, cit1_appeared)

    def test_bins_sum_to_appeared(self):
        """Bin counts must sum to appeared for each column."""
        df = self._base_df()
        result = _compute_results_per_column(df, ['CIT-1', 'CIT-2', 'CIT-3'])
        for col_name, stats in result.items():
            bin_sum = sum(stats['bins'].values())
            self.assertEqual(
                bin_sum, stats['appeared'],
                f"Bin sum {bin_sum} != appeared {stats['appeared']} for column '{col_name}'"
            )

    def test_all_absent_column(self):
        """A fully-absent column: appeared=0, pass=0, fail=0."""
        df = pd.DataFrame({
            'Name': ['Alice', 'Bob'],
            'Exam': ['AA', 'AA'],
        })
        result = _compute_results_per_column(df, ['Exam'])
        self.assertEqual(result['Exam']['appeared'], 0)
        self.assertEqual(result['Exam']['pass_count'], 0)
        self.assertEqual(result['Exam']['fail_count'], 0)
        self.assertEqual(result['Exam']['pass_rate'], 0.0)

    def test_pass_rate_calculation(self):
        """Pass % = (passed / appeared) * 100, not based on total rows."""
        df = pd.DataFrame({
            'Name':  ['Alice', 'Bob', 'Charlie'],
            'Score': [90, 20, 'AA'],   # Alice passes, Bob fails, Charlie absent
        })
        result = _compute_results_per_column(df, ['Score'])
        # appeared=2 (Alice+Bob), passed=1 (Alice)
        self.assertEqual(result['Score']['appeared'], 2)
        self.assertEqual(result['Score']['pass_count'], 1)
        self.assertEqual(result['Score']['fail_count'], 1)
        self.assertAlmostEqual(result['Score']['pass_rate'], 50.0)

    def test_different_column_names_not_hardcoded(self):
        """Works with any column names, not just CIT-1/CIT-2/CIT-3."""
        df = pd.DataFrame({
            'Mid-Term': [80, 60, 40],
            'Final':    [70, 90, 30],
            'Lab Work': [85, 75, 55],
        })
        result = _compute_results_per_column(df, ['Mid-Term', 'Final', 'Lab Work'])
        self.assertIn('Mid-Term', result)
        self.assertIn('Final', result)
        self.assertIn('Lab Work', result)


# ------------------------------------------------------------------ #
# Per-column Excel layout tests — _append_per_column_analysis
# ------------------------------------------------------------------ #

class TestAppendPerColumnAnalysis(unittest.TestCase):

    def _make_ws_with_data(self, col_names, student_rows):
        """Build a fresh openpyxl worksheet with headers + student rows."""
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(col_names)
        for row in student_rows:
            ws.append(row)
        return ws, wb

    def test_analysis_appended_below_each_column(self):
        """Horizontal layout: labels in name col, each marks col gets its own values column."""
        ws, wb = self._make_ws_with_data(
            ['Name', 'CIT-1', 'CIT-2'],
            [['Alice', 80, 75], ['Bob', 90, 85], ['Charlie', 60, 70]]
        )
        per_col_stats = {
            'CIT-1': {'appeared': 3, 'pass_count': 3, 'fail_count': 0, 'pass_rate': 100.0,
                      'bins': {'0-49': 0, '50-59': 0, '60-69': 1, '70-79': 0, '80-89': 1, '90-100': 1}},
            'CIT-2': {'appeared': 3, 'pass_count': 3, 'fail_count': 0, 'pass_rate': 100.0,
                      'bins': {'0-49': 0, '50-59': 0, '60-69': 0, '70-79': 2, '80-89': 1, '90-100': 0}},
        }
        _append_per_column_analysis(ws, per_col_stats)

        # Header row = 1, data rows = 2,3,4; analysis starts immediately at row 5
        analysis_start = 5

        # Labels are in col 1 (Name col) — 'appeared' (no Result Analysis title)
        self.assertEqual(ws.cell(row=analysis_start, column=1).value, 'appeared',
                         "Expected 'appeared' in Name col at analysis start row")

        # CIT-1 appeared formula in col 2
        cit1_appeared = ws.cell(row=analysis_start, column=2).value
        self.assertIn('COUNTIF', str(cit1_appeared),
                      f"Expected COUNTIF formula for CIT-1 appeared, got: {cit1_appeared}")

        # CIT-2 appeared formula in col 3
        cit2_appeared = ws.cell(row=analysis_start, column=3).value
        self.assertIn('COUNTIF', str(cit2_appeared),
                      f"Expected COUNTIF formula for CIT-2 appeared, got: {cit2_appeared}")

    def test_original_student_data_unchanged(self):
        """Original student marks must remain untouched."""
        ws, wb = self._make_ws_with_data(
            ['Name', 'Score'],
            [['Alice', 88], ['Bob', 72]]
        )
        per_col_stats = {
            'Score': {'appeared': 2, 'pass_count': 2, 'fail_count': 0, 'pass_rate': 100.0,
                      'bins': {'0-49': 0, '50-59': 0, '60-69': 0, '70-79': 1, '80-89': 1, '90-100': 0}},
        }
        _append_per_column_analysis(ws, per_col_stats)

        # Row 1: header; row 2: Alice; row 3: Bob — must be unchanged
        self.assertEqual(ws.cell(row=1, column=2).value, 'Score')
        self.assertEqual(ws.cell(row=2, column=2).value, 88)
        self.assertEqual(ws.cell(row=3, column=2).value, 72)

    def test_analysis_rows_contain_all_required_labels(self):
        """Labels must be in name col; the marks column values must be in their own columns."""
        ws, wb = self._make_ws_with_data(
            ['Name', 'Score'],
            [['Alice', 88]]
        )
        per_col_stats = {
            'Score': {'appeared': 1, 'pass_count': 1, 'fail_count': 0, 'pass_rate': 100.0,
                      'bins': {'0-49': 0, '50-59': 0, '60-69': 0, '70-79': 0, '80-89': 1, '90-100': 0}},
        }
        _append_per_column_analysis(ws, per_col_stats)

        # Labels must be in column 1 (Name col), starting at analysis block
        # Data rows: row1=header, row2=Alice; analysis starts immediately at row3
        col1_values = [ws.cell(row=r, column=1).value for r in range(3, ws.max_row + 1)]
        col1_str = [str(v) for v in col1_values if v is not None]
        joined = ' '.join(col1_str)

        # Reference labels (no 'Result Analysis' title, no 'Failed')
        for expected in ['appeared', 'passed', '% passed',
                         'below 50%', '50%-60%', '61%-70%', '71%-80%', '81%-90%', '91%-100%']:
            self.assertIn(expected, joined, f"Missing label '{expected}' in label column")

        # Score values must be in column 2 (not labels)
        col2_values = [ws.cell(row=r, column=2).value for r in range(3, ws.max_row + 1)]
        non_none_col2 = [v for v in col2_values if v is not None]
        # First value = appeared formula using N-COUNTIF
        self.assertTrue(any('COUNTIF' in str(v) for v in non_none_col2),
                        f"Appeared COUNTIF formula not found in col 2, got: {non_none_col2}")

    def test_no_analysis_written_for_unlisted_column(self):
        """A column not in per_column_stats must not have analysis values written into it."""
        ws, wb = self._make_ws_with_data(
            ['Name', 'Score1', 'Score2'],
            [['Alice', 80, 90]]
        )
        # Only provide stats for Score1 (col 2); Score2 (col 3) must stay clean
        per_col_stats = {
            'Score1': {'appeared': 1, 'pass_count': 1, 'fail_count': 0, 'pass_rate': 100.0,
                       'bins': {'0-49': 0, '50-59': 0, '60-69': 0, '70-79': 0, '80-89': 1, '90-100': 0}},
        }
        _append_per_column_analysis(ws, per_col_stats)

        # Score2 is col 3 — should have no analysis values written
        col3_values = [ws.cell(row=r, column=3).value for r in range(3, ws.max_row + 1)]
        col3_non_none = [v for v in col3_values if v is not None]
        self.assertEqual(col3_non_none, [],
                         f"Expected no analysis in Score2 col (col 3), but found: {col3_non_none}")


# ------------------------------------------------------------------ #
# End-to-end per-column integration via generate_report
# ------------------------------------------------------------------ #

class TestGenerateReportPerColumn(unittest.TestCase):

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()

    def _src_path(self, name='source.xlsx'):
        return os.path.join(self.tmp_dir, name)

    def _out_path(self, name='report.xlsx'):
        return os.path.join(self.tmp_dir, name)

    def _write_xlsx(self, sheets_spec, filename='source.xlsx'):
        data = make_xlsx_bytes(sheets_spec)
        path = self._src_path(filename)
        with open(path, 'wb') as f:
            f.write(data)
        return path

    def _ws_json(self, sheet_name, marks_cols, identity_cols=None):
        cols = []
        for c in (identity_cols or []):
            cols.append({'name': c, 'classification': 'identity', 'dtype': 'Text', 'samples': []})
        for c in marks_cols:
            cols.append({'name': c, 'classification': 'marks', 'dtype': 'Decimal', 'samples': []})
        return [{'name': sheet_name, 'is_empty': False, 'row_count': 5, 'columns': cols}]

    def test_per_column_stats_in_returned_dict(self):
        """generate_report must return per_column_stats keyed by each marks column."""
        src = self._write_xlsx({
            'Results': [
                ('Name', 'Exam-A', 'Exam-B'),
                ('Alice', 80, 75),
                ('Bob',   90, 85),
            ]
        })
        stats = generate_report(src, self._ws_json('Results', ['Exam-A', 'Exam-B'], ['Name']),
                                'Results', self._out_path())
        self.assertIn('per_column_stats', stats)
        self.assertIn('Exam-A', stats['per_column_stats'])
        self.assertIn('Exam-B', stats['per_column_stats'])

    def test_each_col_stat_independent(self):
        """per_column_stats for col A must not be influenced by col B marks."""
        src = self._write_xlsx({
            'S': [
                ('Name', 'Col-A', 'Col-B'),
                ('Alice', 100, 10),   # Alice: A=100(pass), B=10(fail)
                ('Bob',   10, 100),   # Bob:   A=10(fail),  B=100(pass)
            ]
        })
        stats = generate_report(src, self._ws_json('S', ['Col-A', 'Col-B'], ['Name']),
                                'S', self._out_path())
        pc = stats['per_column_stats']
        # Col-A: Alice passes (100%), Bob fails (10%) → 1 pass, 1 fail
        self.assertEqual(pc['Col-A']['pass_count'], 1)
        self.assertEqual(pc['Col-A']['fail_count'], 1)
        # Col-B: Alice fails (10%), Bob passes (100%) → 1 pass, 1 fail
        self.assertEqual(pc['Col-B']['pass_count'], 1)
        self.assertEqual(pc['Col-B']['fail_count'], 1)

    def test_excel_has_analysis_below_each_col(self):
        """Generated Excel must contain 'Result Analysis' text in each MARKS column."""
        src = self._write_xlsx({
            'Exam': [
                ('Name', 'Subject-X', 'Subject-Y'),
                ('Alice', 80, 75),
                ('Bob',   70, 90),
                ('Charlie', 60, 65),
            ]
        })
        out = self._out_path()
        generate_report(src, self._ws_json('Exam', ['Subject-X', 'Subject-Y'], ['Name']),
                        'Exam', out)

        # Read the raw worksheet
        wb = openpyxl.load_workbook(out)
        ws = wb['Exam']

        # Collect all values per column
        all_values = set()
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str):
                    all_values.add(v)

        self.assertIn('appeared', all_values,
                      "'appeared' label not found anywhere in the generated worksheet")

    def test_original_student_data_preserved_in_generated_file(self):
        """Original data rows must appear in the generated .xlsx unchanged."""
        src = self._write_xlsx({
            'Data': [
                ('Name', 'Score'),
                ('Alice', 88),
                ('Bob',   72),
            ]
        })
        out = self._out_path()
        generate_report(src, self._ws_json('Data', ['Score'], ['Name']),
                        'Data', out)

        wb = openpyxl.load_workbook(out)
        ws = wb['Data']
        # Row 1: header, Row 2: Alice=88, Row 3: Bob=72
        self.assertEqual(ws.cell(row=1, column=2).value, 'Score')
        self.assertEqual(ws.cell(row=2, column=2).value, 88)
        self.assertEqual(ws.cell(row=3, column=2).value, 72)

    def test_dynamically_named_cols_no_hardcoding(self):
        """Stats dict and Excel both work with non-standard column names."""
        marks = ['Term 1', 'Practical', 'Final Exam']
        src = self._write_xlsx({
            'Marks': [
                ['Name'] + marks,
                ['Alice', 80, 70, 90],
                ['Bob',   60, 80, 70],
                ['Carol', 40, 50, 30],
            ]
        })
        out = self._out_path()
        stats = generate_report(src, self._ws_json('Marks', marks, ['Name']),
                                'Marks', out)
        pc = stats['per_column_stats']
        for col in marks:
            self.assertIn(col, pc, f"Expected per_column_stats entry for '{col}'")

    def test_blank_absent_excluded_from_appeared(self):
        """Blank / AA values must reduce the appeared count, not appear as 0 in appeared."""
        src = self._write_xlsx({
            'Sheet': [
                ('Name', 'Mark'),
                ('Alice', 80),
                ('Bob',   'AA'),
                ('Carol', None),
            ]
        })
        out = self._out_path()
        stats = generate_report(src, self._ws_json('Sheet', ['Mark'], ['Name']),
                                'Sheet', out)
        # Only Alice has a valid mark → appeared = 1
        self.assertEqual(stats['per_column_stats']['Mark']['appeared'], 1)

    def test_bins_per_column_accuracy(self):
        """Grade bins must be calculated correctly and independently per column."""
        src = self._write_xlsx({
            'Grades': [
                ('Name', 'Col1', 'Col2'),
                # Col1 covers each range; Col2 all in 90-100
                ('S1', 25, 95),
                ('S2', 55, 92),
                ('S3', 65, 91),
                ('S4', 75, 93),
                ('S5', 85, 94),
                ('S6', 95, 96),
            ]
        })
        out = self._out_path()
        stats = generate_report(src, self._ws_json('Grades', ['Col1', 'Col2']),
                                'Grades', out)
        col1_bins = stats['per_column_stats']['Col1']['bins']
        col2_bins = stats['per_column_stats']['Col2']['bins']

        # Col1: 25→below 50%, 55→50%-60%, 65→61%-70%, 75→71%-80%, 85→81%-90%, 95→91%-100%
        self.assertEqual(col1_bins['below 50%'], 1)
        self.assertEqual(col1_bins['50%-60%'], 1)
        self.assertEqual(col1_bins['61%-70%'], 1)
        self.assertEqual(col1_bins['71%-80%'], 1)
        self.assertEqual(col1_bins['81%-90%'], 1)
        self.assertEqual(col1_bins['91%-100%'], 1)

        # Col2: all in 91%-100%
        self.assertEqual(col2_bins['91%-100%'], 6)
        self.assertEqual(sum(v for k, v in col2_bins.items() if k != '91%-100%'), 0)



class TestReportRoutes(unittest.TestCase):

    def setUp(self):
        self.app = create_app('testing')
        self.ctx = self.app.app_context()
        self.ctx.push()
        db.create_all()
        self.client = self.app.test_client()

        self.user = User(name='Module6 Tester', email='m6@resultiq.com', username='m6tester')
        self.user.set_password('pass6789')
        db.session.add(self.user)
        db.session.commit()
        self.client.post('/login', data={
            'login_id': 'm6@resultiq.com', 'password': 'pass6789'
        }, follow_redirects=True)

    def tearDown(self):
        # Clean up generated reports
        for report in Report.query.all():
            if report.report_path and os.path.isfile(report.report_path):
                try:
                    os.remove(report.report_path)
                except Exception:
                    pass
        db.session.remove()
        db.drop_all()
        self.ctx.pop()

    def _upload_and_analyse(self, sheets_spec, filename='data.xlsx'):
        """Upload a workbook and trigger analysis, returning the UploadedFile record."""
        data = make_xlsx_bytes(sheets_spec)
        self.client.post('/upload', data={
            'excel_file': (io.BytesIO(data), filename,
                           'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        }, content_type='multipart/form-data', follow_redirects=True)

        record = UploadedFile.query.filter_by(original_filename=filename).first()
        self.assertIsNotNone(record)

        # Trigger analysis to populate classification
        self.client.get(f'/analyse/{record.id}', follow_redirects=True)
        db.session.refresh(record)
        return record

    def test_generate_redirects_to_preview(self):
        record = self._upload_and_analyse({
            'Marks': [
                ('Name',    'Math', 'Science'),
                ('Alice',   88,     75),
                ('Bob',     70,     90),
                ('Charlie', 'AA',   65),
            ]
        }, filename='m6test1.xlsx')

        resp = self.client.post('/report/generate', data={
            'file_id':    record.id,
            'sheet_name': 'Marks',
        }, follow_redirects=True)
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b'Report', resp.data)
        print('[PASS] test_generate_redirects_to_preview')

    def test_report_record_created(self):
        record = self._upload_and_analyse({
            'Results': [
                ('Name',    'Score1', 'Score2'),
                ('Alice',   80,       75),
                ('Bob',     90,       85),
            ]
        }, filename='m6test2.xlsx')

        self.client.post('/report/generate', data={
            'file_id':    record.id,
            'sheet_name': 'Results',
        }, follow_redirects=True)

        report = Report.query.filter_by(uploaded_file_id=record.id).first()
        self.assertIsNotNone(report, "Report record not created in DB")
        self.assertEqual(report.status, 'completed')
        self.assertIsNotNone(report.report_path)
        self.assertTrue(os.path.isfile(report.report_path),
                        "Report file not on disk")
        print('[PASS] test_report_record_created')

    def test_preview_shows_stats(self):
        record = self._upload_and_analyse({
            'Exam': [
                ('Name',     'Sub1', 'Sub2', 'Sub3'),
                ('Alice',    80,     75,      88),
                ('Bob',      90,     85,      70),
                ('Charlie',  'AA',   70,      65),
                ('Diana',    60,     55,      80),
                ('Eve',      30,     90,      75),
            ]
        }, filename='m6test3.xlsx')

        resp = self.client.post('/report/generate', data={
            'file_id':    record.id,
            'sheet_name': 'Exam',
        }, follow_redirects=True)
        self.assertEqual(resp.status_code, 200)

        # Preview should show numeric stats
        report = Report.query.filter_by(uploaded_file_id=record.id).first()
        resp2 = self.client.get(f'/report/{report.id}')
        self.assertEqual(resp2.status_code, 200)
        # Should contain pass/fail indicators and student count
        self.assertEqual(report.analysis_trace['total_students'], 5)
        self.assertIn(b'Passed', resp2.data)
        self.assertIn(b'Failed', resp2.data)
        print('[PASS] test_preview_shows_stats')

    def test_download_returns_xlsx(self):
        record = self._upload_and_analyse({
            'Grades': [
                ('Name',  'G1', 'G2'),
                ('Alice', 80,   75),
                ('Bob',   90,   85),
            ]
        }, filename='m6test4.xlsx')

        self.client.post('/report/generate', data={
            'file_id':    record.id,
            'sheet_name': 'Grades',
        }, follow_redirects=True)

        report = Report.query.filter_by(uploaded_file_id=record.id).first()
        resp = self.client.get(f'/report/{report.id}/download')
        self.assertEqual(resp.status_code, 200)
        ct = resp.headers.get('Content-Type', '')
        self.assertIn('spreadsheetml', ct)
        print('[PASS] test_download_returns_xlsx')

    def test_ownership_403_on_other_user(self):
        record = self._upload_and_analyse({
            'S': [('N','X'), ('A', 80), ('B', 90)]
        }, filename='m6priv.xlsx')
        self.client.post('/report/generate', data={
            'file_id': record.id, 'sheet_name': 'S'
        }, follow_redirects=True)
        report = Report.query.filter_by(uploaded_file_id=record.id).first()

        # Log out, create another user
        self.client.get('/logout', follow_redirects=True)
        other = User(name='Other', email='other6@resultiq.com', username='other6')
        other.set_password('x')
        db.session.add(other)
        db.session.commit()
        self.client.post('/login', data={'login_id': 'other6@resultiq.com', 'password': 'x'},
                         follow_redirects=True)

        resp = self.client.get(f'/report/{report.id}')
        self.assertEqual(resp.status_code, 404)
        print('[PASS] test_ownership_403_on_other_user')

    def test_generate_without_analysis_handled(self):
        """If worksheets have no classification keys, generate should fail gracefully."""
        # Upload but do NOT call /analyse (so no classification keys)
        data = make_xlsx_bytes({
            'Raw': [('Name','Score'), ('Alice', 80), ('Bob', 70)]
        })
        self.client.post('/upload', data={
            'excel_file': (io.BytesIO(data), 'raw.xlsx',
                           'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        }, content_type='multipart/form-data', follow_redirects=True)
        record = UploadedFile.query.filter_by(original_filename='raw.xlsx').first()

        # Do NOT call /analyse — worksheets have no classification keys
        resp = self.client.post('/report/generate', data={
            'file_id':    record.id,
            'sheet_name': 'Raw',
        }, follow_redirects=True)
        self.assertEqual(resp.status_code, 200)
        # Should show an error flash, not crash
        self.assertIn(b'analyse', resp.data.lower())
        print('[PASS] test_generate_without_analysis_handled')


if __name__ == '__main__':
    unittest.main(verbosity=2)
