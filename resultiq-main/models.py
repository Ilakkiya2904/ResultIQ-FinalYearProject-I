from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
import uuid
from datetime import datetime, timezone

db = SQLAlchemy()

def generate_uuid():
    return str(uuid.uuid4())

class User(UserMixin, db.Model):
    __tablename__ = 'users'

    id = db.Column(db.String(36), primary_key=True, default=generate_uuid)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False, index=True)
    username = db.Column(db.String(50), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    department = db.Column(db.String(100))
    designation = db.Column(db.String(100))
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))

    # Relationships
    uploaded_files = db.relationship('UploadedFile', backref='user', lazy=True, cascade='all, delete-orphan')
    reports = db.relationship('Report', backref='user', lazy=True, cascade='all, delete-orphan')
    activity_logs = db.relationship('ActivityLog', backref='user', lazy=True, cascade='all, delete-orphan')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

class UploadedFile(db.Model):
    __tablename__ = 'uploaded_files'

    id = db.Column(db.String(36), primary_key=True, default=generate_uuid)
    user_id = db.Column(db.String(36), db.ForeignKey('users.id'), nullable=False, index=True)
    original_filename = db.Column(db.String(255), nullable=False)
    stored_filename = db.Column(db.String(255), nullable=False, unique=True)
    file_size = db.Column(db.Integer)
    file_path = db.Column(db.String(512), nullable=False)
    uploaded_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    is_valid = db.Column(db.Boolean, default=False)
    validation_errors = db.Column(db.Text)
    worksheets = db.Column(db.JSON)  # Using JSON type for structured data

    # Relationships
    reports = db.relationship('Report', backref='uploaded_file', lazy=True, cascade='all, delete-orphan')

class Report(db.Model):
    __tablename__ = 'reports'

    id = db.Column(db.String(36), primary_key=True, default=generate_uuid)
    user_id = db.Column(db.String(36), db.ForeignKey('users.id'), nullable=False, index=True)
    uploaded_file_id = db.Column(db.String(36), db.ForeignKey('uploaded_files.id'), nullable=False, index=True)
    original_filename = db.Column(db.String(255), nullable=False)
    report_name = db.Column(db.String(255), nullable=False)
    sheet_name = db.Column(db.String(100))
    requested_columns = db.Column(db.JSON)
    report_path = db.Column(db.String(512))
    generated_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    downloaded_at = db.Column(db.DateTime)
    status = db.Column(db.String(50), default='pending') # pending, processing, completed, failed
    error_message = db.Column(db.Text)
    analysis_trace = db.Column(db.JSON)

class ActivityLog(db.Model):
    __tablename__ = 'activity_logs'

    id = db.Column(db.String(36), primary_key=True, default=generate_uuid)
    user_id = db.Column(db.String(36), db.ForeignKey('users.id'), nullable=False, index=True)
    activity_type = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text)
    timestamp = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
