import tempfile
import unittest
from pathlib import Path

from admin_user_tools import clear_user_data, grant_word_uses, load_admin_user_data
from database import Database


class AdminUserToolsTest(unittest.TestCase):
    def test_admin_grant_adds_uses_and_records_admin(self):
        with tempfile.TemporaryDirectory() as tmp:
            db = Database(Path(tmp) / "bot.sqlite3")
            with db.connect() as con:
                con.execute(
                    "INSERT INTO users(user_id,phone,username,registered_at,last_seen_at) "
                    "VALUES(?,?,?,?,?)",
                    (123, "+998000000000", "student", "now", "now"),
                )
                con.execute("INSERT INTO credits(user_id,balance) VALUES(?,1)", (123,))

            self.assertEqual(grant_word_uses(db, 123, 900, 2), ("ok", 3))
            self.assertEqual(grant_word_uses(db, 123, 900, 0), ("invalid_amount", 0))
            user_data = load_admin_user_data(db, 123)
            self.assertEqual(user_data["remaining_uses"], 3)
            self.assertEqual(user_data["grants"][0]["admin_id"], 900)
            with db.connect() as con:
                grant = con.execute(
                    "SELECT user_id,admin_id,uses FROM admin_grants"
                ).fetchone()
                balance = con.execute(
                    "SELECT balance FROM credits WHERE user_id=123"
                ).fetchone()["balance"]
            self.assertEqual(tuple(grant), (123, 900, 2))
            self.assertEqual(balance, 3)

    def test_clear_user_data_anonymizes_and_deletes_private_content_but_keeps_audit(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            private = root / "private"
            private.mkdir()
            photo = private / "student.jpg"
            template = private / "template.docx"
            output = private / "class.docx"
            receipt = private / "receipt.jpg"
            shared_template = root / "6D.docx"
            for path in (photo, template, output, receipt, shared_template):
                path.write_bytes(b"file")
            db = Database(root / "bot.sqlite3")
            with db.connect() as con:
                con.execute(
                    "INSERT INTO users(user_id,phone,username,registered_at,last_seen_at) "
                    "VALUES(?,?,?,?,?)",
                    (123, "+998000000000", "student", "now", "now"),
                )
                con.execute("INSERT INTO credits(user_id,balance) VALUES(?,4)", (123,))
                session = con.execute(
                    "INSERT INTO roster_sessions(user_id,class_name,template_path,state,updated_at) "
                    "VALUES(?,?,?,?,?)",
                    (123, "7A", str(template), "complete", "now"),
                ).lastrowid
                con.execute(
                    "INSERT INTO students(session_id,position,photo_path,name) VALUES(?,?,?,?)",
                    (session, 1, str(photo), "Ali"),
                )
                con.execute(
                    "INSERT INTO submissions(user_id,kind,text,file_path,created_at) "
                    "VALUES(?,?,?,?,?)",
                    (123, "photo", "caption", str(photo), "now"),
                )
                con.execute(
                    "INSERT INTO documents(user_id,class_name,file_path,session_id,created_at) "
                    "VALUES(?,?,?,?,?)",
                    (123, "7A", str(output), session, "now"),
                )
                payment = con.execute(
                    "INSERT INTO payments(user_id,receipt_path,amount,credits,status,created_at) "
                    "VALUES(?,?,?,?,?,?)",
                    (123, str(receipt), 5000, 1, "approved", "now"),
                ).lastrowid
                con.execute(
                    "INSERT INTO payment_events(payment_id,actor_id,action,created_at) "
                    "VALUES(?,?,?,?)",
                    (payment, 900, "approved", "now"),
                )
                con.execute(
                    "INSERT INTO admin_grants(user_id,admin_id,uses,created_at) "
                    "VALUES(?,?,?,?)",
                    (123, 900, 2, "now"),
                )
                con.execute(
                    "INSERT INTO promo_codes(code,credits,max_uses,created_by,created_at) "
                    "VALUES(?,?,?,?,?)",
                    ("KEEP", 1, 10, 900, "now"),
                )
                con.execute(
                    "INSERT INTO promo_redemptions(code,user_id,redeemed_at) VALUES(?,?,?)",
                    ("KEEP", 123, "now"),
                )
                con.execute(
                    "INSERT INTO roster_sessions(user_id,class_name,template_path,state,updated_at) "
                    "VALUES(?,?,?,?,?)",
                    (123, "6D", str(shared_template), "complete", "now"),
                )

            status, deleted, failed = clear_user_data(db, 123, private)
            self.assertEqual((status, deleted, failed), ("ok", 4, 0))
            self.assertTrue(shared_template.is_file())
            self.assertFalse(any(path.exists() for path in (photo, template, output, receipt)))
            with db.connect() as con:
                user = con.execute(
                    "SELECT phone,username FROM users WHERE user_id=123"
                ).fetchone()
                counts = (
                    con.execute("SELECT COUNT(*) FROM students").fetchone()[0],
                    con.execute("SELECT COUNT(*) FROM roster_sessions").fetchone()[0],
                    con.execute("SELECT COUNT(*) FROM submissions").fetchone()[0],
                    con.execute("SELECT COUNT(*) FROM documents").fetchone()[0],
                )
                balance = con.execute(
                    "SELECT balance FROM credits WHERE user_id=123"
                ).fetchone()["balance"]
                payment_count = con.execute("SELECT COUNT(*) FROM payments").fetchone()[0]
                event_count = con.execute("SELECT COUNT(*) FROM payment_events").fetchone()[0]
                grant_count = con.execute("SELECT COUNT(*) FROM admin_grants").fetchone()[0]
                promo_count = con.execute(
                    "SELECT COUNT(*) FROM promo_redemptions WHERE user_id=123"
                ).fetchone()[0]
            self.assertIsNone(user["phone"])
            self.assertEqual(user["username"], "")
            self.assertEqual(counts, (0, 0, 0, 0))
            self.assertEqual(balance, 0)
            self.assertEqual(payment_count, 1)
            self.assertEqual(event_count, 1)
            self.assertEqual(grant_count, 1)
            self.assertEqual(promo_count, 1)


if __name__ == "__main__":
    unittest.main()
