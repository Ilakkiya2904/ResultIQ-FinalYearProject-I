"""
app.py
======
Flask application factory for ResultIQ.
All modules register their Blueprints here.
"""

import os
from flask import Flask, render_template
from flask_login import LoginManager

from config import config
from models import db, User


def create_app(config_name: str = 'default') -> Flask:
    """
    Create and configure the Flask application.

    Parameters
    ----------
    config_name : str
        One of 'development', 'testing', 'production', 'default'.

    Returns
    -------
    Flask application instance.
    """
    app = Flask(__name__)
    app.config.from_object(config[config_name])

    # ------------------------------------------------------------------ #
    # Ensure required directories exist
    # ------------------------------------------------------------------ #
    for folder_key in ('UPLOAD_FOLDER', 'REPORT_FOLDER'):
        folder = app.config.get(folder_key)
        if folder:
            os.makedirs(folder, exist_ok=True)

    # ------------------------------------------------------------------ #
    # Extensions
    # ------------------------------------------------------------------ #
    db.init_app(app)

    login_manager = LoginManager()
    login_manager.init_app(app)
    login_manager.login_view = 'auth.login'
    login_manager.login_message = 'Please log in to access this page.'
    login_manager.login_message_category = 'warning'

    @login_manager.user_loader
    def load_user(user_id: str):
        return db.session.get(User, user_id)

    # ------------------------------------------------------------------ #
    # Jinja2 custom filters / globals
    # ------------------------------------------------------------------ #
    app.jinja_env.filters['enumerate'] = enumerate

    # ------------------------------------------------------------------ #
    # Blueprints
    # ------------------------------------------------------------------ #
    from routes.auth import auth_bp
    from routes.upload import upload_bp
    from routes.analyse import analyse_bp
    from routes.report import report_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(upload_bp)
    app.register_blueprint(analyse_bp)
    app.register_blueprint(report_bp)

    # ------------------------------------------------------------------ #
    # Root route
    # ------------------------------------------------------------------ #
    @app.route('/')
    def index():
        return render_template('index.html')

    # ------------------------------------------------------------------ #
    # Error handlers
    # ------------------------------------------------------------------ #
    @app.errorhandler(404)
    def page_not_found(e):
        return render_template('404.html'), 404

    @app.errorhandler(500)
    def internal_error(e):
        db.session.rollback()
        return render_template('500.html'), 500

    return app


# ------------------------------------------------------------------ #
# Dev-server entry point
# ------------------------------------------------------------------ #
if __name__ == '__main__':
    app = create_app('development')
    with app.app_context():
        db.create_all()
    app.run(debug=True)
