import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from PIL import Image

from admin_exports import create_user_archive, create_user_report


class AdminExportsTest(unittest.TestCase):
    def test_word_and_pdf_reports_include_data_and_full_file_archive(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            private = root / "private"
            private.mkdir()
            photo = private / "student.png"
            Image.new("RGB", (40, 30), "white").save(photo)
            template = private / "template.docx"
            template.write_bytes(b"template")
            output = private / "class.docx"
            output.write_bytes(b"generated")
            receipt = private / "receipt.png"
            Image.new("RGB", (30, 20), "white").save(receipt)
            shared_template = root / "6D.docx"
            shared_template.write_bytes(b"shared")
            data = {
                "user": {
                    "user_id": 321,
                    "username": "student",
                    "phone": "+998000000000",
                },
                "remaining_uses": 2,
                "sessions": [{
                    "class_name": "7A",
                    "state": "complete",
                    "template_path": str(template),
                }, {
                    "class_name": "6D",
                    "state": "complete",
                    "template_path": str(shared_template),
                }],
                "students": [{
                    "class_name": "7A",
                    "position": 1,
                    "name": "Ali",
                    "age": "12",
                    "gender": "erkak",
                    "phone": "123456789",
                    "extra_json": '{"Manzil":"Toshkent"}',
                    "photo_path": str(photo),
                }],
                "submissions": [{
                    "kind": "student_details",
                    "text": "Ali, 12",
                    "file_path": None,
                    "created_at": "now",
                }],
                "documents": [{
                    "class_name": "7A",
                    "file_path": str(output),
                    "created_at": "now",
                }],
                "payments": [{
                    "id": 8,
                    "amount": 10000,
                    "credits": 2,
                    "status": "approved",
                    "receipt_path": str(receipt),
                    "created_at": "now",
                }],
                "payment_events": [{
                    "payment_id": 8,
                    "actor_id": 900,
                    "action": "approved",
                    "created_at": "now",
                }],
                "grants": [{"admin_id": 900, "uses": 1, "created_at": "now"}],
                "promos": [{"code": "TEST", "redeemed_at": "now"}],
            }

            word_path = root / "user.docx"
            create_user_report(data, "docx", word_path, private, shared_template)
            word = Document(word_path)
            paragraphs = [paragraph.text for paragraph in word.paragraphs]
            self.assertIn("Foydalanuvchi hisoboti", paragraphs[0])
            self.assertIn("Ali", "\n".join(paragraphs))
            self.assertGreaterEqual(len(word.inline_shapes), 2)

            pdf_path = root / "user.pdf"
            create_user_report(data, "pdf", pdf_path, private, shared_template)
            self.assertTrue(pdf_path.read_bytes().startswith(b"%PDF-"))

            archive_path = root / "user-files.zip"
            self.assertTrue(
                create_user_archive(data, archive_path, private, shared_template)
            )
            with zipfile.ZipFile(archive_path) as archive:
                names = archive.namelist()
                self.assertTrue(any("student.png" in name for name in names))
                self.assertTrue(any("template.docx" in name for name in names))
                self.assertTrue(any("class.docx" in name for name in names))
                self.assertTrue(any("receipt.png" in name for name in names))
                self.assertTrue(any("6D.docx" in name for name in names))


if __name__ == "__main__":
    unittest.main()
