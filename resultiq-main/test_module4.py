"""
Module 4 validation tests.
Creates test workbooks, uploads them via the Flask test client,
and verifies the inspection results against the database.
"""
import os
import io
import tempfile
import unittest
import openpyxl
from openpyxl import Workbook
from app import create_app
from models import db, UploadedFile, User


def make_workbook_bytes(sheets_spec):
    """
    Build an in-memory .xlsx workbook from a dict:
      { 'SheetName': [ (col1, col2, ...), (val1, val2, ...), ... ] }
    Returns bytes.
    """
    wb = Workbook()
    wb.remove(wb.active)           # remove default sheet
    for name, rows in sheets_spec.items():
        ws = wb.create_sheet(title=name)
        for row in rows:
            ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


class Module4Tests(unittest.TestCase):

    def setUp(self):
        self.app = create_app('testing')
        self.ctx = self.app.app_context()
        self.ctx.push()
        db.create_all()
        self.client = self.app.test_client()

        # Create and log in a test user
        self.user = User(name='Test Faculty', email='m4test@resultiq.com', username='m4testuser')
        self.user.set_password('pass1234')
        db.session.add(self.user)
        db.session.commit()

        # Authenticate
        self.client.post('/login', data={
            'login_id': 'm4test@resultiq.com',
            'password': 'pass1234'
        }, follow_redirects=True)

    def tearDown(self):
        db.session.remove()
        db.drop_all()
        self.ctx.pop()

    # ------------------------------------------------------------------ #
    def _upload(self, filename, data, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'):
        return self.client.post('/upload', data={
            'excel_file': (io.BytesIO(data), filename, content_type)
        }, content_type='multipart/form-data', follow_redirects=True)

    # ------------------------------------------------------------------ #
    def test_valid_multi_sheet_workbook(self):
        """Upload a 3-sheet workbook with mixed data types and verify inspection."""
        xls = make_workbook_bytes({
            'Students': [
                ('Roll No', 'Student Name', 'Score', 'Grade'),
                (101, 'Alice', 88.5, 'A'),
                (102, 'Bob', 72.0, 'B'),
                (103, 'Charlie', 91.0, 'A+'),
            ],
            'Teachers': [
                ('ID', 'Name', 'Department'),
                (1, 'Dr. Smith', 'Physics'),
                (2, 'Prof. Rao', 'Maths'),
            ],
            'EmptySheet': []
        })

        resp = self._upload('results.xlsx', xls)
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b'results.xlsx', resp.data)   # inspect page shows filename

        # Verify DB record
        record = UploadedFile.query.filter_by(original_filename='results.xlsx').first()
        self.assertIsNotNone(record, "UploadedFile record should be created")
        self.assertTrue(record.is_valid)

        sheets = record.worksheets
        sheet_names = [s['name'] for s in sheets]
        self.assertIn('Students', sheet_names)
        self.assertIn('Teachers', sheet_names)
        self.assertIn('EmptySheet', sheet_names)

        students = next(s for s in sheets if s['name'] == 'Students')
        col_names = [c['name'] for c in students['columns']]
        self.assertIn('Roll No', col_names)
        self.assertIn('Student Name', col_names)
        self.assertIn('Score', col_names)

        empty = next(s for s in sheets if s['name'] == 'EmptySheet')
        self.assertTrue(empty['is_empty'], "EmptySheet should be flagged as empty")

        print("[PASS] test_valid_multi_sheet_workbook")

    def test_different_column_names(self):
        """Ensure column names are taken verbatim from the workbook."""
        xls = make_workbook_bytes({
            'Survey': [
                ('Respondent ID', 'Age Group', 'Response Text', 'Rating (1-5)'),
                (1, '18-25', 'Excellent', 5),
                (2, '26-35', 'Good', 4),
            ]
        })
        self._upload('survey.xlsx', xls)
        record = UploadedFile.query.filter_by(original_filename='survey.xlsx').first()
        col_names = [c['name'] for c in record.worksheets[0]['columns']]
        self.assertIn('Respondent ID', col_names)
        self.assertIn('Rating (1-5)', col_names)
        print("[PASS] test_different_column_names")

    def test_invalid_file_type(self):
        """Uploading a .txt file should be rejected with a user-friendly error."""
        resp = self._upload('notes.txt', b'hello world', content_type='text/plain')
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b'Only .xlsx and .xls files are accepted', resp.data)
        print("[PASS] test_invalid_file_type")

    def test_corrupted_workbook(self):
        """Uploading random bytes named .xlsx should fail gracefully."""
        resp = self._upload('corrupt.xlsx', b'PK this is not a real xlsx file!!!')
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b'Could not read the workbook', resp.data)
        # No DB record should be created
        record = UploadedFile.query.filter_by(original_filename='corrupt.xlsx').first()
        self.assertIsNone(record, "Corrupt files must not be saved to the database")
        print("[PASS] test_corrupted_workbook")

    def test_no_file_submitted(self):
        """Submitting with no file selected should show an error."""
        resp = self.client.post('/upload',
                                data={'excel_file': (io.BytesIO(b''), '', 'application/octet-stream')},
                                content_type='multipart/form-data',
                                follow_redirects=True)
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b'Please select a file before uploading', resp.data)
        print("[PASS] test_no_file_submitted")

    def test_inspect_route_ownership(self):
        """Cannot access another user's uploaded file."""
        xls = make_workbook_bytes({'Sheet1': [('A', 'B'), (1, 2)]})
        self._upload('priv.xlsx', xls)
        record = UploadedFile.query.filter_by(original_filename='priv.xlsx').first()

        # Log out and create another user
        self.client.get('/logout', follow_redirects=True)
        other = User(name='Other', email='other@resultiq.com', username='otherusr')
        other.set_password('pass9999')
        db.session.add(other)
        db.session.commit()
        self.client.post('/login', data={'login_id': 'other@resultiq.com', 'password': 'pass9999'}, follow_redirects=True)

        resp = self.client.get(f'/upload/inspect/{record.id}')
        self.assertEqual(resp.status_code, 404)
        print("[PASS] test_inspect_route_ownership")


if __name__ == '__main__':
    unittest.main(verbosity=2)
