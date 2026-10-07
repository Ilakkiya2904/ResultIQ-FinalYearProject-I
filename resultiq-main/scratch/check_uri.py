import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from app import create_app
from models import db

app = create_app('development')
with app.app_context():
    print("SQLALCHEMY_DATABASE_URI:", app.config['SQLALCHEMY_DATABASE_URI'])
    print("db.engine.url.database:", db.engine.url.database)
