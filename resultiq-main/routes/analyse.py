"""
routes/analyse.py
=================
Module 5 — Intelligent Column Classification route.

GET  /analyse/<file_id>
    Calls services.analyser.analyse_workbook() to enrich column metadata,
    persists the enriched JSON back to UploadedFile.worksheets, then
    renders the results page which includes the Generate Report form.
"""

from flask import Blueprint, render_template, redirect, url_for, flash, abort, current_app
from flask_login import login_required, current_user

from models import db, UploadedFile, ActivityLog
from services.analyser import analyse_workbook

analyse_bp = Blueprint('analyse', __name__)


@analyse_bp.route('/analyse/<file_id>')
@login_required
def analyse(file_id: str):
    """Enrich column metadata and render the classification results page."""

    # --- Ownership check ---
    record = db.session.get(UploadedFile, file_id)
    if record is None or record.user_id != current_user.id:
        abort(404)

    # --- Run analysis ---
    try:
        enriched_sheets = analyse_workbook(record.file_path)
    except ValueError as exc:
        flash(f'Analysis failed: {exc}', 'danger')
        return redirect(url_for('upload.inspect', file_id=file_id))
    except Exception as exc:
        current_app.logger.error(f'Unexpected analysis error for file {file_id}: {exc}')
        flash('An unexpected error occurred during analysis. Please try again.', 'danger')
        return redirect(url_for('upload.inspect', file_id=file_id))

    # --- Persist enriched JSON ---
    try:
        record.worksheets = enriched_sheets
        log = ActivityLog(
            user_id=current_user.id,
            activity_type='FILE_ANALYSED',
            description=f'Analysed columns for: {record.original_filename}'
        )
        db.session.add(log)
        db.session.commit()
    except Exception as exc:
        db.session.rollback()
        current_app.logger.error(f'DB persist error after analysis for file {file_id}: {exc}')
        flash('Could not save analysis results. Please try again.', 'danger')
        return redirect(url_for('upload.inspect', file_id=file_id))

    # --- Prepare summary counts for the template ---
    all_columns = []
    for sheet in enriched_sheets:
        all_columns.extend(sheet.get('columns', []))

    classification_counts = {}
    for col in all_columns:
        cls = col.get('classification', 'unknown')
        classification_counts[cls] = classification_counts.get(cls, 0) + 1

    marks_count = classification_counts.get('marks', 0)

    return render_template(
        'analyse/results.html',
        record=record,
        sheets=enriched_sheets,
        classification_counts=classification_counts,
        marks_count=marks_count,
    )
