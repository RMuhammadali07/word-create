import json
import sqlite3
from contextlib import contextmanager
from pathlib import Path


SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    user_id INTEGER PRIMARY KEY,
    phone TEXT,
    username TEXT,
    free_additions_used INTEGER NOT NULL DEFAULT 0,
    registered_at TEXT NOT NULL,
    last_seen_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS roster_sessions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    class_name TEXT,
    template_path TEXT,
    state TEXT NOT NULL DEFAULT 'awaiting_template',
    current_student INTEGER NOT NULL DEFAULT 0,
    correction_document_id INTEGER,
    correction_base_count INTEGER NOT NULL DEFAULT 0,
    correction_is_free INTEGER NOT NULL DEFAULT 0,
    correction_max_additions INTEGER NOT NULL DEFAULT 3,
    updated_at TEXT NOT NULL,
    FOREIGN KEY(user_id) REFERENCES users(user_id)
);

CREATE TABLE IF NOT EXISTS students (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id INTEGER NOT NULL,
    position INTEGER NOT NULL,
    photo_path TEXT,
    name TEXT,
    age TEXT,
    gender TEXT,
    phone TEXT,
    extra_json TEXT NOT NULL DEFAULT '{}',
    FOREIGN KEY(session_id) REFERENCES roster_sessions(id)
);

CREATE TABLE IF NOT EXISTS submissions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    kind TEXT NOT NULL,
    text TEXT,
    file_path TEXT,
    created_at TEXT NOT NULL,
    FOREIGN KEY(user_id) REFERENCES users(user_id)
);

CREATE TABLE IF NOT EXISTS credits (
    user_id INTEGER PRIMARY KEY,
    balance INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY(user_id) REFERENCES users(user_id)
);

CREATE TABLE IF NOT EXISTS payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    receipt_path TEXT NOT NULL,
    amount INTEGER,
    credits INTEGER NOT NULL DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'pending',
    reviewed_by INTEGER,
    reviewed_at TEXT,
    created_at TEXT NOT NULL,
    FOREIGN KEY(user_id) REFERENCES users(user_id)
);

CREATE TABLE IF NOT EXISTS payment_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    payment_id INTEGER NOT NULL,
    actor_id INTEGER NOT NULL,
    action TEXT NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY(payment_id) REFERENCES payments(id)
);

CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    class_name TEXT NOT NULL,
    file_path TEXT NOT NULL,
    session_id INTEGER,
    free_additions_used INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    FOREIGN KEY(user_id) REFERENCES users(user_id)
);

CREATE TABLE IF NOT EXISTS promo_codes (
    code TEXT PRIMARY KEY,
    credits INTEGER NOT NULL,
    max_uses INTEGER NOT NULL,
    uses INTEGER NOT NULL DEFAULT 0,
    created_by INTEGER NOT NULL,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS promo_redemptions (
    code TEXT NOT NULL,
    user_id INTEGER NOT NULL,
    redeemed_at TEXT NOT NULL,
    PRIMARY KEY(code,user_id),
    FOREIGN KEY(code) REFERENCES promo_codes(code),
    FOREIGN KEY(user_id) REFERENCES users(user_id)
);

CREATE TABLE IF NOT EXISTS admin_grants (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    admin_id INTEGER NOT NULL,
    uses INTEGER NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY(user_id) REFERENCES users(user_id)
);

CREATE TABLE IF NOT EXISTS app_settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
INSERT OR IGNORE INTO app_settings(key,value) VALUES('shared_template_enabled','0');

CREATE INDEX IF NOT EXISTS idx_students_session ON students(session_id);
CREATE INDEX IF NOT EXISTS idx_documents_owner ON documents(user_id);
"""


class Database:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as con:
            con.executescript(SCHEMA)
            required_columns = {
                "users": {
                    "free_additions_used": "INTEGER NOT NULL DEFAULT 0",
                },
                "students": {
                    "extra_json": "TEXT NOT NULL DEFAULT '{}'",
                },
                "roster_sessions": {
                    "correction_document_id": "INTEGER",
                    "correction_base_count": "INTEGER NOT NULL DEFAULT 0",
                    "correction_is_free": "INTEGER NOT NULL DEFAULT 0",
                    "correction_max_additions": "INTEGER NOT NULL DEFAULT 3",
                },
                "documents": {
                    "free_additions_used": "INTEGER NOT NULL DEFAULT 0",
                },
            }
            for table, columns in required_columns.items():
                existing = {
                    row["name"] for row in con.execute(f"PRAGMA table_info({table})")
                }
                for name, declaration in columns.items():
                    if name not in existing:
                        con.execute(
                            f"ALTER TABLE {table} ADD COLUMN {name} {declaration}"
                        )

    @contextmanager
    def connect(self):
        con = sqlite3.connect(self.path)
        con.row_factory = sqlite3.Row
        con.execute("PRAGMA foreign_keys = ON")
        try:
            yield con
            con.commit()
        except Exception:
            con.rollback()
            raise
        finally:
            con.close()

    def get_setting(self, key, default=None):
        with self.connect() as con:
            row = con.execute(
                "SELECT value FROM app_settings WHERE key=?",
                (key,),
            ).fetchone()
        return row["value"] if row else default

    def set_setting(self, key, value):
        with self.connect() as con:
            con.execute(
                "INSERT INTO app_settings(key,value) VALUES(?,?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, value),
            )

