from flask import Blueprint, render_template, redirect, url_for, flash, request, session
from flask_login import login_user, logout_user, login_required, current_user
from sqlalchemy import func
from models import db, User, Report

auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/register', methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('auth.dashboard'))
        
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        email = request.form.get('email', '').strip().lower()
        username = request.form.get('username', '').strip().lower()
        department = request.form.get('department', '').strip()
        designation = request.form.get('designation', '').strip()
        password = request.form.get('password', '')
        confirm_password = request.form.get('confirm_password', '')

        if not all([name, email, username, password, confirm_password]):
            flash('All fields are required.', 'danger')
            return redirect(url_for('auth.register'))

        if password != confirm_password:
            flash('Passwords do not match.', 'danger')
            return redirect(url_for('auth.register'))

        import re
        if len(password) < 8:
            flash('Password must be at least 8 characters long.', 'danger')
            return redirect(url_for('auth.register'))
        if not re.search(r'[A-Z]', password):
            flash('Password must contain at least 1 uppercase letter.', 'danger')
            return redirect(url_for('auth.register'))
        if not re.search(r'[a-z]', password):
            flash('Password must contain at least 1 lowercase letter.', 'danger')
            return redirect(url_for('auth.register'))
        if not re.search(r'[0-9]', password):
            flash('Password must contain at least 1 number.', 'danger')
            return redirect(url_for('auth.register'))
        if not re.search(r'[^A-Za-z0-9\s]', password):
            flash('Password must contain at least 1 special character.', 'danger')
            return redirect(url_for('auth.register'))
        if ' ' in password:
            flash('Password should not contain spaces.', 'danger')
            return redirect(url_for('auth.register'))

        existing_email = User.query.filter(func.lower(User.email) == email).first()
        if existing_email:
            flash('Email already registered. Please sign in.', 'danger')
            return redirect(url_for('auth.login'))

        existing_user = User.query.filter(func.lower(User.username) == username).first()
        if existing_user:
            flash('Username already taken.', 'danger')
            return redirect(url_for('auth.register'))

        user = User(
            name=name,
            email=email,
            username=username,
            department=department,
            designation=designation
        )
        user.set_password(password)
        
        try:
            db.session.add(user)
            db.session.commit()
            flash('Registration successful! You can now log in.', 'success')
            return redirect(url_for('auth.login'))
        except Exception as e:
            db.session.rollback()
            flash('An error occurred during registration. Please try again.', 'danger')
            
    return render_template('auth/register.html')

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('auth.dashboard'))
        
    if request.method == 'POST':
        login_id = request.form.get('login_id', '').strip().lower()  # Can be email or username
        password = request.form.get('password', '')
        remember = request.form.get('remember', 'on') in ('on', 'true', '1')

        if not login_id or not password:
            flash('Please provide both username/email and password.', 'danger')
            return redirect(url_for('auth.login'))

        user = User.query.filter(
            (func.lower(User.email) == login_id) | 
            (func.lower(User.username) == login_id)
        ).first()

        password_matches = False
        if user:
            # Check exact password first; fallback to stripped password if user accidentally included extra spaces
            password_matches = user.check_password(password) or (
                bool(password.strip()) and user.check_password(password.strip())
            )

        if user and password_matches:
            # Mark session as permanent to keep user logged in across app switches and browser sessions
            session.permanent = True
            login_user(user, remember=remember)
            flash('Logged in successfully.', 'success')
            next_page = request.args.get('next')
            return redirect(next_page or url_for('auth.dashboard'))
        elif not user:
            flash('No registered account found with that email or username. Please check your spelling or register an account.', 'danger')
        else:
            flash('Incorrect password. Please verify and try again.', 'danger')

    return render_template('auth/login.html')

@auth_bp.route('/logout')
@login_required
def logout():
    session.clear()
    logout_user()
    flash('You have been logged out.', 'info')
    return redirect(url_for('index'))

@auth_bp.route('/dashboard')
@login_required
def dashboard():
    recent_reports = (Report.query
                      .filter_by(user_id=current_user.id)
                      .order_by(Report.generated_at.desc())
                      .limit(5)
                      .all())
    return render_template('dashboard.html', recent_reports=recent_reports)
