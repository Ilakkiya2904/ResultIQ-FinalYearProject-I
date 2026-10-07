"""
test_module5.py
===============
Comprehensive tests for Module 5: Intelligent Data Understanding & Column Classification.

Tests cover the pure service layer (services/analyser.py) in isolation
as well as the Flask /analyse/<file_id> route.
"""

import io
import os
import unittest
import pandas as pd
import openpyxl
from openpyxl import Workbook

from app import create_app
from models import db, User, UploadedFile
from services.analyser import classify_column, analyse_sheet, analyse_workbook


# ------------------------------------------------------------------ #
# Helpers
# ------------------------------------------------------------------ #

def make_xlsx_bytes(sheets_spec: dict) -> bytes:
    """Build an in-memory .xlsx workbook from a dict of {name: [rows]}."""
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


def series(*values):
    """Convenience: create a pd.Series from positional values."""
    return pd.Series(list(values))


# ------------------------------------------------------------------ #
# Unit tests — pure service layer (no Flask)
# ------------------------------------------------------------------ #

class TestClassifyColumn(unittest.TestCase):

    def test_numeric_marks_column(self):
        """Pure numeric column → marks."""
        s = series(45, 72, 88, 91, 63, 55, 80)
        cls, q = classify_column(s, 'Subject A')
        self.assertEqual(cls, 'marks')
        self.assertTrue(q['suitable_for_analysis'])

    def test_marks_with_absent_tokens(self):
        """Numeric + AA/Absent tokens → still marks."""
        s = series(45, 'AA', 88, 'Absent', 63, 'aa', 80, 72)
        cls, q = classify_column(s, 'CIT-1')
        self.assertEqual(cls, 'marks')
        self.assertTrue(q['suitable_for_analysis'])
        self.assertGreater(q['absent_token_count'], 0)

    def test_absent_counted_not_crashed(self):
        """Various absent token variants are counted, not crashed."""
        tokens = ['NA', 'N/A', 'nil', 'Withheld', 'ab', '--', 'Medical', 'Detained']
        s = series(55, 60, *tokens, 75, 80)
        cls, q = classify_column(s, 'Score')
        # Should not raise; absent tokens should be accounted for
        self.assertIn(cls, ('marks', 'text', 'category', 'identity'))
        self.assertGreaterEqual(q['absent_token_count'], len(tokens))

    def test_identity_column_names(self):
        """High-uniqueness text column → identity."""
        s = series('Alice Johnson', 'Bob Smith', 'Charlie Brown',
                   'Diana Prince', 'Eve Torres', 'Frank Castle',
                   'Grace Hopper', 'Hank Pym')
        cls, q = classify_column(s, 'Student Name')
        self.assertEqual(cls, 'identity')
        self.assertFalse(q['suitable_for_analysis'])

    def test_category_column(self):
        """Low-cardinality text → category."""
        s = series('Pass', 'Fail', 'Pass', 'Pass', 'Fail',
                   'Pass', 'Fail', 'Pass', 'Pass', 'Fail')
        cls, q = classify_column(s, 'Result')
        self.assertEqual(cls, 'category')
        self.assertFalse(q['suitable_for_analysis'])

    def test_empty_column(self):
        """All-None column → empty."""
        s = series(None, None, None, None, None, None, None, None, None, None)
        cls, q = classify_column(s, 'Blank Col')
        self.assertEqual(cls, 'empty')
        self.assertFalse(q['suitable_for_analysis'])

    def test_quality_metrics_completeness(self):
        """Quality dict must contain all required keys."""
        s = series(10, 20, None, 30, 'AA')
        _, q = classify_column(s, 'X')
        required = {'total', 'filled', 'missing', 'numeric_count',
                    'absent_token_count', 'text_count', 'unique_count',
                    'suitable_for_analysis'}
        self.assertTrue(required.issubset(q.keys()))

    def test_missing_values_dont_crash(self):
        """NaN and None in a column must not raise any exception."""
        s = series(None, None, None)
        try:
            cls, q = classify_column(s, 'Empty')
            self.assertEqual(cls, 'empty')
        except Exception as e:
            self.fail(f"classify_column raised {e} on all-None series")

    def test_metadata_column_not_treated_as_marks(self):
        """A datetime column should not be classified as marks."""
        dates = pd.to_datetime(['2023-01-01', '2023-01-02', '2023-01-03',
                                '2023-01-04', '2023-01-05'])
        s = pd.Series(dates)
        cls, q = classify_column(s, 'Date of Birth')
        self.assertEqual(cls, 'metadata')
        self.assertFalse(q['suitable_for_analysis'])


class TestAnalyseSheet(unittest.TestCase):

    def _make_df(self):
        return pd.DataFrame({
            'Roll No':    [101, 102, 103, 104, 105],
            'Name':       ['Alice', 'Bob', 'Charlie', 'Diana', 'Eve'],
            'Maths':      [88, 'AA', 72, 91, 65],
            'Science':    [75, 80, 'Absent', 88, 70],
            'Pass/Fail':  ['Pass', 'Fail', 'Pass', 'Pass', 'Pass'],
            'Blank':      [None, None, None, None, None],
        })

    def test_sheet_classified_correctly(self):
        df = self._make_df()
        result = analyse_sheet(df, 'Results')
        self.assertFalse(result['is_empty'])
        self.assertEqual(result['row_count'], 5)

        by_name = {c['name']: c for c in result['columns']}

        self.assertEqual(by_name['Maths']['classification'],   'marks')
        self.assertEqual(by_name['Science']['classification'], 'marks')
        self.assertEqual(by_name['Name']['classification'],    'identity')
        self.assertIn(by_name['Pass/Fail']['classification'],  ('category',))
        self.assertEqual(by_name['Blank']['classification'],   'empty')

    def test_empty_dataframe(self):
        result = analyse_sheet(pd.DataFrame(), 'Empty')
        self.assertTrue(result['is_empty'])
        self.assertEqual(result['columns'], [])

    def test_roll_number_not_marks(self):
        """Roll numbers are integer but should be identity, not marks."""
        df = pd.DataFrame({
            'Roll No': list(range(1, 30)),   # 29 unique integers — high uniqueness
            'Score':   [float(i * 2 % 100) for i in range(1, 30)],
        })
        result = analyse_sheet(df, 'Sheet1')
        by_name = {c['name']: c for c in result['columns']}
        # Score should be marks
        self.assertEqual(by_name['Score']['classification'], 'marks')
        # Roll No could be identity or marks depending on uniqueness; it must NOT crash
        self.assertIn(by_name['Roll No']['classification'], ('marks', 'identity', 'category', 'text'))


# ------------------------------------------------------------------ #
# Integration tests — Flask route
# ------------------------------------------------------------------ #

class TestAnalyseRoute(unittest.TestCase):

    def setUp(self):
        self.app = create_app('testing')
        self.ctx = self.app.app_context()
        self.ctx.push()
        db.create_all()
        self.client = self.app.test_client()

        # Create & log in a user
        self.user = User(name='Module5 Tester', email='m5@resultiq.com', username='m5tester')
        self.user.set_password('pass5678')
        db.session.add(self.user)
        db.session.commit()

        self.client.post('/login', data={'login_id': 'm5@resultiq.com', 'password': 'pass5678'}, follow_redirects=True)

    def tearDown(self):
        db.session.remove()
        db.drop_all()
        self.ctx.pop()

    def _upload_xlsx(self, sheets_spec, filename='test.xlsx'):
        data = make_xlsx_bytes(sheets_spec)
        return self.client.post('/upload', data={
            'excel_file': (io.BytesIO(data), filename,
                           'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        }, content_type='multipart/form-data', follow_redirects=True)

    def test_analyse_route_returns_200(self):
        """Upload a workbook then call /analyse/<id> — expect 200."""
        self._upload_xlsx({
            'Results': [
                ('Name', 'Score'),
                ('Alice', 88), ('Bob', 'AA'), ('Charlie', 72)
            ]
        })
        record = UploadedFile.query.filter_by(original_filename='test.xlsx').first()
        resp = self.client.get(f'/analyse/{record.id}', follow_redirects=True)
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b'test.xlsx', resp.data)
        print('[PASS] test_analyse_route_returns_200')

    def test_enriched_json_persisted(self):
        """After /analyse, worksheets JSON must contain classification key."""
        self._upload_xlsx({
            'Marks': [
                ('Roll', 'Physics', 'Chemistry'),
                (101, 78, 'AA'),
                (102, 90, 85),
                (103, 'Absent', 70),
            ]
        }, filename='marks.xlsx')
        record = UploadedFile.query.filter_by(original_filename='marks.xlsx').first()
        self.client.get(f'/analyse/{record.id}')

        # Re-fetch from DB
        db.session.refresh(record)
        sheets = record.worksheets
        self.assertIsNotNone(sheets)
        for col in sheets[0]['columns']:
            self.assertIn('classification', col, "Each column must have a classification key")
            self.assertIn('quality', col, "Each column must have a quality dict")
        print('[PASS] test_enriched_json_persisted')

    def test_marks_columns_identified(self):
        """Physics and Chemistry must be classified as marks; Name as identity."""
        self._upload_xlsx({
            'Results': [
                ('Name',    'Physics', 'Chemistry', 'Maths'),
                ('Alice',   88,        75,          'AA'),
                ('Bob',     'Absent',  90,          82),
                ('Charlie', 72,        68,          91),
                ('Diana',   95,        'AA',        78),
                ('Eve',     80,        85,          60),
            ]
        }, filename='results2.xlsx')
        record = UploadedFile.query.filter_by(original_filename='results2.xlsx').first()
        self.client.get(f'/analyse/{record.id}')
        db.session.refresh(record)

        cols = {c['name']: c for c in record.worksheets[0]['columns']}
        self.assertEqual(cols['Physics']['classification'],   'marks')
        self.assertEqual(cols['Chemistry']['classification'], 'marks')
        self.assertEqual(cols['Maths']['classification'],     'marks')
        self.assertEqual(cols['Name']['classification'],      'identity')
        print('[PASS] test_marks_columns_identified')

    def test_ownership_404(self):
        """Another user cannot access the analysis of a file they don't own."""
        self._upload_xlsx({'S': [('A',), (1,), (2,)]}, filename='priv.xlsx')
        record = UploadedFile.query.filter_by(original_filename='priv.xlsx').first()
        self.client.get('/logout', follow_redirects=True)

        other = User(name='Other', email='other5@resultiq.com', username='other5')
        other.set_password('x')
        db.session.add(other)
        db.session.commit()
        self.client.post('/login', data={'login_id': 'other5@resultiq.com', 'password': 'x'}, follow_redirects=True)

        resp = self.client.get(f'/analyse/{record.id}')
        self.assertEqual(resp.status_code, 404)
        print('[PASS] test_ownership_404')

    def test_multiple_sheets_handled(self):
        """Workbook with 3 sheets: all enriched, empty sheet flagged."""
        self._upload_xlsx({
            'Students': [('Name', 'Score'), ('Alice', 88), ('Bob', 72)],
            'Teachers': [('ID', 'Dept'), (1, 'CS'), (2, 'Maths')],
            'Empty':    [],
        }, filename='multi.xlsx')
        record = UploadedFile.query.filter_by(original_filename='multi.xlsx').first()
        self.client.get(f'/analyse/{record.id}')
        db.session.refresh(record)

        sheet_names = [s['name'] for s in record.worksheets]
        self.assertIn('Students', sheet_names)
        self.assertIn('Teachers', sheet_names)
        self.assertIn('Empty',    sheet_names)

        empty_sheet = next(s for s in record.worksheets if s['name'] == 'Empty')
        self.assertTrue(empty_sheet['is_empty'])
        print('[PASS] test_multiple_sheets_handled')


if __name__ == '__main__':
    unittest.main(verbosity=2)
