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
            created_at TEXT,
            last_updated TEXT
        )
    """)

    # Migration: older databases were created before "created_at" existed.
    # Add it and backfill using each product's earliest history entry,
    # since that row is created at the same moment the product first
    # appears (see upsert_part below).
    cur.execute("PRAGMA table_info(parts)")
    existing_columns = [row["name"] for row in cur.fetchall()]
    if "created_at" not in existing_columns:
        cur.execute("ALTER TABLE parts ADD COLUMN created_at TEXT")
        cur.execute("""
            UPDATE parts
            SET created_at = (
                SELECT MIN(changed_at) FROM location_history
                WHERE location_history.code = parts.code
            )
            WHERE created_at IS NULL
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
            """INSERT INTO parts (code, name, model, current_location, created_at, last_updated)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (code, name, model, location, now, now),
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


def get_part(code):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM parts WHERE code = ?", (code,))
    row = cur.fetchone()
    conn.close()
    return row

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
    """Return location history for a product, newest first.

    Entries where the location is empty are skipped - if a product
    currently has no location on file, we don't want a blank row
    cluttering the history. Older entries that DO have a location are
    still shown.
    """
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """SELECT * FROM location_history
           WHERE code = ? AND location IS NOT NULL AND TRIM(location) != ''
           ORDER BY changed_at DESC""",
        (code,),
    )
    rows = cur.fetchall()
    conn.close()
    return rows

def get_new_parts_by_date(gregorian_date_str):
    """Return parts first added to the system on the given Gregorian
    date (format 'YYYY-MM-DD'), along with the location they were
    placed in AT THAT TIME - not today's current location, which may
    have changed since.
    """
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """SELECT p.code AS code, p.name AS name, p.model AS model,
                  lh.location AS location
           FROM parts p
           JOIN location_history lh ON lh.code = p.code AND lh.changed_at = p.created_at
           WHERE date(p.created_at) = ?
           ORDER BY p.name COLLATE NOCASE""",
        (gregorian_date_str,),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_location_changes_by_date(gregorian_date_str):
    """Return location changes (joined with the product name) that
    happened on the given Gregorian date.

    Also returns the previous location before the change.
    """
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            curr.code,
            p.name,
            curr.location,
            curr.changed_at,
            (
                SELECT prev.location
                FROM location_history prev
                WHERE prev.code = curr.code
                  AND prev.changed_at < curr.changed_at
                ORDER BY prev.changed_at DESC
                LIMIT 1
            ) AS previous_location
        FROM location_history curr
        JOIN parts p
            ON p.code = curr.code
        WHERE date(curr.changed_at) = ?
          AND curr.location IS NOT NULL
          AND TRIM(curr.location) != ''
          AND date(p.created_at) != date(curr.changed_at)
        ORDER BY p.name COLLATE NOCASE
    """, (gregorian_date_str,))

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
