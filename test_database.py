from pathlib import Path
import sqlite3
import tempfile
from contextlib import closing

from database import Database


import unittest


class DatabaseTest(unittest.TestCase):
    def test_database_schema_and_credit_transaction(self):
        with tempfile.TemporaryDirectory() as tmp:
            db = Database(Path(tmp) / "test.sqlite3")
            self.assertEqual(db.get_setting("shared_template_enabled"), "0")
            db.set_setting("shared_template_enabled", "1")
            self.assertEqual(db.get_setting("shared_template_enabled"), "1")
            with db.connect() as con:
                con.execute(
                    "INSERT INTO users(user_id,phone,registered_at,last_seen_at) "
                    "VALUES(?,?,?,?)",
                    (123, "+998000000000", "2026-01-01", "2026-01-01"),
                )
                con.execute("INSERT INTO credits(user_id,balance) VALUES(?,1)", (123,))
            with db.connect() as con:
                row = con.execute("SELECT balance FROM credits WHERE user_id=123").fetchone()
                self.assertEqual(row["balance"], 1)

    def test_existing_database_gets_new_payment_and_correction_columns(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "legacy.sqlite3"
            with closing(sqlite3.connect(path)) as con:
                con.executescript(
                    """
                    CREATE TABLE users (
                        user_id INTEGER PRIMARY KEY, phone TEXT, username TEXT,
                        registered_at TEXT NOT NULL, last_seen_at TEXT NOT NULL
                    );
                    CREATE TABLE roster_sessions (
                        id INTEGER PRIMARY KEY, user_id INTEGER, class_name TEXT,
                        template_path TEXT, state TEXT, current_student INTEGER,
                        updated_at TEXT
                    );
                    CREATE TABLE students (
                        id INTEGER PRIMARY KEY, session_id INTEGER, position INTEGER,
                        photo_path TEXT, name TEXT, age TEXT, gender TEXT, phone TEXT
                    );
                    CREATE TABLE documents (
                        id INTEGER PRIMARY KEY, user_id INTEGER, class_name TEXT,
                        file_path TEXT, session_id INTEGER, created_at TEXT
                    );
                    """
                )
            Database(path)
            with closing(sqlite3.connect(path)) as con:
                users_columns = {row[1] for row in con.execute("PRAGMA table_info(users)")}
                session_columns = {row[1] for row in con.execute("PRAGMA table_info(roster_sessions)")}
                document_columns = {row[1] for row in con.execute("PRAGMA table_info(documents)")}
                promo_table = con.execute(
                    "SELECT 1 FROM sqlite_master WHERE type='table' AND name='promo_codes'"
                ).fetchone()
                grant_table = con.execute(
                    "SELECT 1 FROM sqlite_master WHERE type='table' AND name='admin_grants'"
                ).fetchone()
                template_setting = con.execute(
                    "SELECT value FROM app_settings WHERE key='shared_template_enabled'"
                ).fetchone()
            self.assertIn("free_additions_used", users_columns)
            self.assertIn("correction_document_id", session_columns)
            self.assertIn("free_additions_used", document_columns)
            self.assertIsNotNone(promo_table)
            self.assertIsNotNone(grant_table)
            self.assertEqual(template_setting[0], "0")


if __name__ == "__main__":
    unittest.main()
