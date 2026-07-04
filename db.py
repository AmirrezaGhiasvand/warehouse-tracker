"""
db.py
Database layer for the Warehouse Tracker app.
Uses a single local SQLite file (warehouse.db) stored next to the app.
"""

import sqlite3
import os
import sys
from datetime import datetime


def get_app_dir():
    """Return the directory the app is running from (works both as
    a plain .py script and as a PyInstaller-built .exe)."""
    if getattr(sys, "frozen", False):
        # Running as a PyInstaller exe
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


DB_PATH = os.path.join(get_app_dir(), "warehouse.db")


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS parts (
            code TEXT PRIMARY KEY,
            name TEXT,
            model TEXT,
            current_location TEXT,
            last_updated TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS location_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL,
            location TEXT,
            changed_at TEXT NOT NULL,
            source_file TEXT,
            FOREIGN KEY (code) REFERENCES parts(code)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS uploads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            uploaded_at TEXT,
            row_count INTEGER,
            new_parts INTEGER,
            location_changes INTEGER
        )
    """)

    cur.execute("CREATE INDEX IF NOT EXISTS idx_parts_name ON parts(name)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_history_code ON location_history(code)")

    conn.commit()
    conn.close()


def upsert_part(cur, code, name, model, location, source_file):
    """
    Insert or update a single part. Returns a tuple:
    (is_new_part: bool, location_changed: bool)
    """
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute("SELECT * FROM parts WHERE code = ?", (code,))
    row = cur.fetchone()

    if row is None:
        # Brand new part
        cur.execute(
            """INSERT INTO parts (code, name, model, current_location, last_updated)
               VALUES (?, ?, ?, ?, ?)""",
            (code, name, model, location, now),
        )
        cur.execute(
            """INSERT INTO location_history (code, location, changed_at, source_file)
               VALUES (?, ?, ?, ?)""",
            (code, location, now, source_file),
        )
        return True, True

    location_changed = (row["current_location"] or "") != (location or "")

    cur.execute(
        """UPDATE parts SET name = ?, model = ?, current_location = ?, last_updated = ?
           WHERE code = ?""",
        (name, model, location, now, code),
    )

    if location_changed:
        cur.execute(
            """INSERT INTO location_history (code, location, changed_at, source_file)
               VALUES (?, ?, ?, ?)""",
            (code, location, now, source_file),
        )

    return False, location_changed


def log_upload(cur, filename, row_count, new_parts, location_changes):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        """INSERT INTO uploads (filename, uploaded_at, row_count, new_parts, location_changes)
           VALUES (?, ?, ?, ?, ?)""",
        (filename, now, row_count, new_parts, location_changes),
    )


def search_parts(query):
    """Search parts by code or name (partial, case-insensitive).

    We lower-case both the column and the query ourselves rather than
    relying on SQLite's built-in LIKE case-folding, since that only
    reliably covers ASCII a-z by default - fine for Latin part codes,
    but we do it explicitly here to be safe.
    """
    conn = get_connection()
    cur = conn.cursor()
    like_query = f"%{query.lower()}%"
    cur.execute(
        """SELECT * FROM parts
           WHERE LOWER(code) LIKE ? OR LOWER(name) LIKE ?
           ORDER BY name COLLATE NOCASE""",
        (like_query, like_query),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_history(code):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """SELECT * FROM location_history
           WHERE code = ?
           ORDER BY changed_at DESC""",
        (code,),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_upload_log(limit=20):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """SELECT * FROM uploads ORDER BY uploaded_at DESC LIMIT ?""",
        (limit,),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_stats():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) AS c FROM parts")
    total_parts = cur.fetchone()["c"]
    cur.execute("SELECT COUNT(*) AS c FROM location_history")
    total_history = cur.fetchone()["c"]
    conn.close()
    return total_parts, total_history
