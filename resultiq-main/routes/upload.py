import os
import uuid
import pandas as pd
from flask import Blueprint, render_template, redirect, url_for, flash, request, current_app, abort
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename
from models import db, UploadedFile, ActivityLog

upload_bp = Blueprint('upload', __name__)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


def allowed_file(filename):
    """Check if the file extension is permitted."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def infer_dtype_label(series):
    """Return a human-readable data type label for a pandas Series."""
    dtype = series.dtype
    if pd.api.types.is_integer_dtype(dtype):
        return 'Integer'
    elif pd.api.types.is_float_dtype(dtype):
        return 'Decimal'
    elif pd.api.types.is_bool_dtype(dtype):
        return 'Boolean'
    elif pd.api.types.is_datetime64_any_dtype(dtype):
        return 'Date/Time'
    else:
        return 'Text'


def inspect_workbook(filepath):
    """
    Intelligently inspect an Excel workbook.

    Returns a list of sheet dicts:
        [
          {
            'name': 'Sheet1',
            'is_empty': False,
            'row_count': 120,
            'columns': [
              {
                'name': 'Student Name',
                'dtype': 'Text',
                'samples': ['Alice', 'Bob', 'Charlie']
              },
              ...
            ]
          },
          ...
        ]

    Raises ValueError for corrupt / unreadable workbooks.
    """
    ext = filepath.rsplit('.', 1)[1].lower()
    try:
        engine = 'openpyxl' if ext == 'xlsx' else 'xlrd'
        xf = pd.ExcelFile(filepath, engine=engine)
    except Exception as exc:
        raise ValueError(f"Cannot read workbook: {exc}") from exc

    sheets = []
    try:
        for sheet_name in xf.sheet_names:
            try:
                df = xf.parse(sheet_name, header=0)
            except Exception:
                sheets.append({'name': sheet_name, 'is_empty': True, 'row_count': 0, 'columns': []})
                continue

            if df.empty or len(df.columns) == 0:
                sheets.append({'name': sheet_name, 'is_empty': True, 'row_count': 0, 'columns': []})
                continue

            columns = []
            for col in df.columns:
                series = df[col].dropna()
                col_str = str(col).strip()
                if col_str.startswith('Unnamed') and series.empty:
                    continue
                samples = [str(v) for v in series.head(3).tolist()]
                columns.append({
                    'name': col_str,
                    'dtype': infer_dtype_label(df[col]),
                    'samples': samples
                })

            sheets.append({
                'name': sheet_name,
                'is_empty': len(columns) == 0,
                'row_count': len(df),
                'columns': columns
            })
    finally:
        xf.close()   # always close the file handle

    return sheets


@upload_bp.route('/upload', methods=['GET', 'POST'])
@login_required
def upload():
    if request.method == 'POST':
        # --- Presence check ---
        if 'excel_file' not in request.files:
            flash('No file was submitted.', 'danger')
            return redirect(url_for('upload.upload'))

        file = request.files['excel_file']
        if file.filename == '':
            flash('Please select a file before uploading.', 'danger')
            return redirect(url_for('upload.upload'))

        # --- Extension check ---
        if not allowed_file(file.filename):
            flash('Only .xlsx and .xls files are accepted.', 'danger')
            return redirect(url_for('upload.upload'))

        # --- Size check (enforce Flask's MAX_CONTENT_LENGTH gracefully) ---
        file.seek(0, 2)   # seek to end
        file_size = file.tell()
        file.seek(0)
        max_size = current_app.config.get('MAX_CONTENT_LENGTH', 16 * 1024 * 1024)
        if file_size > max_size:
            flash(f'File exceeds the {max_size // (1024*1024)} MB size limit.', 'danger')
            return redirect(url_for('upload.upload'))

        # --- Save securely ---
        original_filename = secure_filename(file.filename)
        ext = original_filename.rsplit('.', 1)[1].lower()
        upload_folder = current_app.config['UPLOAD_FOLDER']

        if ext == 'xls':
            temp_xls_filename = f"{uuid.uuid4().hex}.xls"
            temp_xls_path = os.path.join(upload_folder, temp_xls_filename)
            file.save(temp_xls_path)

            stored_filename = f"{uuid.uuid4().hex}.xlsx"
            filepath = os.path.join(upload_folder, stored_filename)
            try:
                from services.reporter import convert_xls_to_xlsx
                convert_xls_to_xlsx(temp_xls_path, filepath)
            finally:
                if os.path.exists(temp_xls_path):
                    try:
                        os.remove(temp_xls_path)
                    except OSError:
                        pass
            file_size = os.path.getsize(filepath)
        else:
            stored_filename = f"{uuid.uuid4().hex}.{ext}"
            filepath = os.path.join(upload_folder, stored_filename)
            file.save(filepath)

        # --- Inspect workbook ---
        try:
            sheets = inspect_workbook(filepath)
        except ValueError as exc:
            os.remove(filepath)   # don't keep the corrupt file
            flash(f'Could not read the workbook: {exc}', 'danger')
            return redirect(url_for('upload.upload'))

        # --- Persist to database ---
        record = UploadedFile(
            user_id=current_user.id,
            original_filename=original_filename,
            stored_filename=stored_filename,
            file_size=file_size,
            file_path=filepath,
            is_valid=True,
            worksheets=sheets   # full JSON metadata stored here
        )
        log = ActivityLog(
            user_id=current_user.id,
            activity_type='FILE_UPLOADED',
            description=f'Uploaded workbook: {original_filename}'
        )
        try:
            db.session.add(record)
            db.session.add(log)
            db.session.commit()
        except Exception as exc:
            db.session.rollback()
            os.remove(filepath)
            flash('A database error occurred. Please try again.', 'danger')
            current_app.logger.error(f'Upload DB error: {exc}')
            return redirect(url_for('upload.upload'))

        flash(f'"{original_filename}" uploaded and inspected successfully.', 'success')
        return redirect(url_for('upload.inspect', file_id=record.id))

    # GET — render the upload form
    # Show recently uploaded files for this user
    recent_files = (UploadedFile.query
                    .filter_by(user_id=current_user.id, is_valid=True)
                    .order_by(UploadedFile.uploaded_at.desc())
                    .limit(5)
                    .all())
    return render_template('upload/index.html', recent_files=recent_files)


@upload_bp.route('/upload/inspect/<file_id>')
@login_required
def inspect(file_id):
    record = db.session.get(UploadedFile, file_id)
    if record is None or record.user_id != current_user.id:
        abort(404)
    sheets = record.worksheets or []
    return render_template('upload/inspect.html', record=record, sheets=sheets)
