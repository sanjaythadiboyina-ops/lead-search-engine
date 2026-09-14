import os
import getpass
import sys
import psycopg
from dotenv import load_dotenv

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(PROJECT_DIR, '.env'))
DATABASE_URL = os.getenv('DATABASE_URL')
ADMIN_EMAIL = 'sanjaythadiboyina@gmail.com'
ADMIN_NAME = 'SANJAY'

if not DATABASE_URL:
    raise SystemExit('DATABASE_URL is missing in .env')

print('WARNING: This permanently clears the current Lead Search Engine database data.')
confirm = input('Type RESET to continue: ').strip()
if confirm != 'RESET':
    raise SystemExit('Cancelled.')

password = getpass.getpass('Enter the new Admin password: ')
if len(password) < 6:
    raise SystemExit('Password must be at least 6 characters.')

sys.path.insert(0, os.path.join(PROJECT_DIR, 'backend'))
import main

# Remove application data completely.
with psycopg.connect(DATABASE_URL) as conn:
    with conn.cursor() as cur:
        cur.execute('''
            DROP TABLE IF EXISTS activity_logs, call_logs, group_leads, lead_groups,
            search_history, leads, users CASCADE
        ''')
    conn.commit()

# Recreate the application's current schema.
main.init_db()

# Defensive schema check: the reset must work even if a stale backend file was used.
with psycopg.connect(DATABASE_URL) as conn:
    with conn.cursor() as cur:
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS active BOOLEAN DEFAULT TRUE")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS team_leader_id INTEGER REFERENCES users(id) ON DELETE SET NULL")
        cur.execute("UPDATE users SET active=TRUE")
        conn.commit()

password_hash, salt = main.hash_password(password)

with psycopg.connect(DATABASE_URL) as conn:
    with conn.cursor() as cur:
        # The dedicated Admin is always this account. All other future accounts are Agents.
        cur.execute("DELETE FROM users WHERE email=%s", (ADMIN_EMAIL,))
        cur.execute(
            """INSERT INTO users(email,name,password_hash,salt,role,team_leader_id,active)
               VALUES(%s,%s,%s,%s,'Admin',NULL,TRUE) RETURNING id""",
            (ADMIN_EMAIL, ADMIN_NAME, password_hash, salt),
        )
        admin_id = cur.fetchone()[0]
        cur.execute("UPDATE users SET role='Agent', team_leader_id=NULL, active=TRUE WHERE id<>%s", (admin_id,))
    conn.commit()

print('\nRESET COMPLETE.')
print('Dedicated Admin account created successfully.')
print('Email:', ADMIN_EMAIL)
print('Password: the password you entered above')
