import sqlite3
import os

db_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'resultiq.db')
print("DB Path:", db_path)

if not os.path.exists(db_path):
    print("Database file does not exist!")
else:
    conn = sqlite3.connect(db_path)
    tables = conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
    print("Tables:", tables)
    
    if ('users',) in tables:
        users = conn.execute("SELECT id, username, email FROM users;").fetchall()
        print("Users:", users)
    else:
        print("Users table not found.")
