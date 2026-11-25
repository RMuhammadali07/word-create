import unittest
from pathlib import Path
from types import SimpleNamespace
import tempfile

from bot import (
    approve_payment,
    can_review_member,
    create_promo_code,
    image_file_id,
    parse_student_details,
    redeem_promo_code,
)
from docx import Document
from PIL import Image
from sinf_tayyorlash import sinfni_yoz
from telegram.ext import CommandHandler
from database import Database


class StudentInputTest(unittest.TestCase):
    def test_caption_lines_map_name_and_phone(self):
        parsed = parse_student_details(
            "Ali Valiyev\n+998901234567",
            {"name": "", "age": "", "gender": "", "phone": ""},
        )
        self.assertEqual(parsed["name"], "Ali Valiyev")
        self.assertEqual(parsed["phone"], "+998901234567")

    def test_text_after_photo_fills_phone_without_overwriting_name(self):
        parsed = parse_student_details(
            "+998901234567",
            {"name": "Ali Valiyev", "age": "", "gender": "", "phone": ""},
        )
        self.assertEqual(parsed["name"], "Ali Valiyev")
        self.assertEqual(parsed["phone"], "+998901234567")

    def test_labeled_caption_fields_can_arrive_in_any_order(self):
        parsed = parse_student_details(
            "Telefon: +998901234567\nIsm: Ali Valiyev",
            {"name": "", "age": "", "gender": "", "phone": ""},
        )
        self.assertEqual(parsed["name"], "Ali Valiyev")
        self.assertEqual(parsed["phone"], "+998901234567")

    def test_caption_lines_follow_template_headers(self):
        headers = [
            "T/R",
            "F.I.SH",
            "Sinf",
            "Sinf harfi",
            "Telefon raqam",
            "Yashash joyi",
            "O'quvchi rasmi",
            "Jinsi",
        ]
        parsed = parse_student_details(
            "Ali Valiyev\n+998901234567\nToshkent\nO'g'il",
            {"name": "", "age": "", "gender": "", "phone": ""},
            headers,
        )
        self.assertEqual(parsed["name"], "Ali Valiyev")
        self.assertEqual(parsed["phone"], "+998901234567")
        self.assertEqual(parsed["gender"], "O'g'il")
        self.assertEqual(parsed["extra"]["Yashash joyi"], "Toshkent")

    def test_attached_caption_name_phone_gender_ignores_optional_address(self):
        headers = ["F.I.SH", "Telefon raqam", "Yashash joyi", "Rasm", "Jinsi"]
        parsed = parse_student_details(
            "Teshaboyev Eshmat\n123456789\nerkak",
            {"name": "", "age": "", "gender": "", "phone": "", "extra_json": "{}"},
            headers,
        )
        self.assertEqual(parsed["name"], "Teshaboyev Eshmat")
        self.assertEqual(parsed["phone"], "123456789")
        self.assertEqual(parsed["gender"], "erkak")
        self.assertNotIn("Yashash joyi", parsed["extra"])

    def test_custom_template_header_receives_caption_value(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            template = root / "template.docx"
            doc = Document()
            table = doc.add_table(rows=2, cols=6)
            headers = ["T/R", "F.I.SH", "Sinf", "Telefon raqam", "Rasm", "Qo'shimcha"]
            for cell, header in zip(table.rows[0].cells, headers):
                cell.text = header
            doc.save(template)
            image = root / "student.png"
            Image.new("RGB", (20, 20), "white").save(image)
            output = sinfni_yoz(
                template,
                ("7", "A"),
                [{
                    "ism": "Ali Valiyev",
                    "tel": "+998901234567",
                    "rasm": image,
                    "extra": {"Qo'shimcha": "Ma'lumot"},
                }],
                root / "out",
                False,
            )
            result = Document(output)
            cells = result.tables[0].rows[1].cells
            self.assertEqual(cells[1].text, "Ali Valiyev")
            self.assertEqual(cells[2].text, "7")
            self.assertEqual(cells[3].text, "+998901234567")
            self.assertTrue(cells[4]._tc.xpath(".//a:blip"))
            self.assertEqual(cells[5].text, "Ma'lumot")

    def test_photo_documents_are_accepted_for_supported_image_types(self):
        image_doc = SimpleNamespace(
            document=SimpleNamespace(
                file_id="photo-id",
                file_name="student.png",
                mime_type="image/png",
            ),
            photo=[],
        )
        self.assertEqual(image_file_id(image_doc), ("photo-id", ".png"))

    def test_all_review_group_members_can_review_but_nonmembers_cannot(self):
        self.assertTrue(can_review_member(SimpleNamespace(status="member", is_member=True)))
        self.assertTrue(can_review_member(SimpleNamespace(status="administrator", is_member=True)))
        self.assertTrue(can_review_member(SimpleNamespace(status="creator", is_member=True)))
        self.assertFalse(can_review_member(SimpleNamespace(status="left", is_member=False)))
        self.assertFalse(can_review_member(SimpleNamespace(status="restricted", is_member=False)))

    def test_admin_panel_command_is_valid_for_telegram(self):
        CommandHandler("admin_jrgrumua", lambda update, context: None)

    def test_promo_adds_credits_once_and_obeys_total_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            db = Database(Path(tmp) / "promos.sqlite3")
            with db.connect() as con:
                for user_id in (101, 202):
                    con.execute(
                        "INSERT INTO users(user_id,phone,registered_at,last_seen_at) "
                        "VALUES(?,?,?,?)",
                        (user_id, "+998000000000", "now", "now"),
                    )
            create_promo_code(db, "TEST10", 2, 1, 999)
            self.assertEqual(redeem_promo_code(db, 101, "TEST10"), ("ok", 2))
            self.assertEqual(redeem_promo_code(db, 101, "TEST10"), ("already_used", 0))
            self.assertEqual(redeem_promo_code(db, 202, "TEST10"), ("exhausted", 0))
            with db.connect() as con:
                balance = con.execute(
                    "SELECT balance FROM credits WHERE user_id=101"
                ).fetchone()["balance"]
            self.assertEqual(balance, 2)

    def test_payment_approval_converts_confirmed_sum_to_whole_word_uses(self):
        with tempfile.TemporaryDirectory() as tmp:
            db = Database(Path(tmp) / "payments.sqlite3")
            with db.connect() as con:
                con.execute(
                    "INSERT INTO users(user_id,phone,registered_at,last_seen_at) "
                    "VALUES(?,?,?,?)",
                    (303, "+998000000000", "now", "now"),
                )
                con.execute("INSERT INTO credits(user_id,balance) VALUES(?,1)", (303,))
                cur = con.execute(
                    "INSERT INTO payments(user_id,receipt_path,amount,credits,status,created_at) "
                    "VALUES(?,?,?,?,?,?)",
                    (303, "receipt.jpg", None, 0, "pending", "now"),
                )
                payment_id = cur.lastrowid

            self.assertEqual(
                approve_payment(db, payment_id, 900, 7500),
                ("invalid_amount", 0, None),
            )
            self.assertEqual(
                approve_payment(db, payment_id, 900, 10000),
                ("approved", 2, 303),
            )
            with db.connect() as con:
                payment = con.execute(
                    "SELECT amount,credits,status FROM payments WHERE id=?",
                    (payment_id,),
                ).fetchone()
                uses = con.execute(
                    "SELECT balance FROM credits WHERE user_id=303"
                ).fetchone()["balance"]
            self.assertEqual(tuple(payment), (10000, 2, "approved"))
            self.assertEqual(uses, 3)


if __name__ == "__main__":
    unittest.main()
