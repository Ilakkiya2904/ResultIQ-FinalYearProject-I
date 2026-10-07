import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from app import create_app
from models import db, User

app = create_app('testing')
client = app.test_client()

with app.app_context():
    db.create_all()
    
    # 1. Register with a trailing space
    response = client.post('/register', data={
        'name': 'Trace User',
        'email': 'trace@example.com ',
        'username': 'traceuser ',
        'department': 'CSE',
        'designation': 'Student',
        'password': 'password123',
        'confirm_password': 'password123'
    }, follow_redirects=True)
    
    user = User.query.filter_by(username='traceuser').first()
    print("User saved as:", repr(user.username) if user else "Not Found")
    
    # 2. Login without a trailing space
    response = client.post('/login', data={
        'login_id': 'traceuser',
        'password': 'password123'
    }, follow_redirects=True)
    
    print("Login success:", b'Logged in successfully' in response.data)
