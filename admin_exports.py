import html
import io
import json
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Inches
from PIL import Image, ImageOps


def _value(value):
    if value in (None, ""):
        return "-"
    return _safe_text(value)


def _safe_text(value):
    text = str(value)
    return "".join(
        char for char in text
        if (
            char in "\t\n\r"
            or 0x20 <= ord(char) <= 0xD7FF
            or 0xE000 <= ord(char) <= 0xFFFD
            or 0x10000 <= ord(char) <= 0x10FFFF
        )
    )


def _extra_fields(student):
    try:
        extra = json.loads(student.get("extra_json") or "{}")
    except (TypeError, json.JSONDecodeError):
        extra = {}
    return extra if isinstance(extra, dict) else {}


def _asset_items(data):
    for student in data["students"]:
        if student.get("photo_path"):
            yield "Rasmlar", f"O'quvchi {student['position']} rasmi", student["photo_path"]
    for submission in data["submissions"]:
        if submission.get("file_path"):
            yield "Yuborilgan fayllar", submission.get("kind") or "Material", submission["file_path"]
    for session in data["sessions"]:
        if session.get("template_path"):
            yield "Shablonlar", session.get("class_name") or "Shablon", session["template_path"]
    for document in data["documents"]:
        if document.get("file_path"):
            yield "Tayyor Wordlar", document.get("class_name") or "Word", document["file_path"]
    for payment in data["payments"]:
        if payment.get("receipt_path"):
            yield "To'lov cheklari", f"To'lov {payment['id']}", payment["receipt_path"]


def _allowed_file(raw_path, private_root, shared_template):
    try:
        path = Path(raw_path).resolve()
        roots = (
            [Path(item).resolve() for item in private_root]
            if isinstance(private_root, (list, tuple, set))
            else [Path(private_root).resolve()]
        )
        shared = Path(shared_template).resolve() if shared_template else None
    except (OSError, TypeError, ValueError):
        return None
    if any(root in path.parents for root in roots) or path == shared:
        return path if path.is_file() else None
    return None


def create_user_archive(data, output_path, private_root, shared_template=None):
    items = [
        (category, label, _allowed_file(path, private_root, shared_template))
        for category, label, path in _asset_items(data)
    ]
    items = [(category, label, path) for category, label, path in items if path]
    if not items:
        return False
    manifest = ["Asl foydalanuvchi fayllari", ""]
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for index, (category, label, path) in enumerate(items, start=1):
            archive_name = f"{category}/{index:04d}_{path.name}"
            archive.write(path, archive_name)
            manifest.append(f"{category}: {label} ({path.name})")
        archive.writestr("MANIFEST.txt", "\n".join(manifest))
    return True


def _thumbnail(path):
    try:
        with Image.open(path) as source:
            image = ImageOps.exif_transpose(source)
            image.thumbnail((1000, 1000))
            if image.mode not in ("RGB", "L"):
                image = image.convert("RGB")
            output = io.BytesIO()
            image.save(output, format="JPEG", quality=78, optimize=True)
            return output.getvalue()
    except (OSError, ValueError):
        return None


def _report_sections(data):
    user = data["user"]
    yield "Hisob", [
        f"Telegram ID: {user['user_id']}",
        f"Username: @{user.get('username') or '-'}",
        f"Telefon: {_value(user.get('phone'))}",
        f"Qolgan Word tayyorlashlar: {data['remaining_uses']}",
        f"Sinf sessiyalari: {len(data['sessions'])}",
        f"O'quvchi yozuvlari: {len(data['students'])}",
        f"Yuborilgan materiallar: {len(data['submissions'])}",
        f"Tayyor Word fayllar: {len(data['documents'])}",
        f"To'lovlar: {len(data['payments'])}",
        f"Foydalanilgan promokodlar: {len(data['promos'])}",
        f"Admin bergan Wordlar: {sum(item['uses'] for item in data['grants'])}",
    ]
    yield "Sinf sessiyalari", [
        f"Sinf: {_value(item.get('class_name'))} | Holat: {item['state']} | "
        f"Shablon: {Path(item['template_path']).name if item.get('template_path') else '-'}"
        for item in data["sessions"]
    ]
    yield "O'quvchilar", [
        " | ".join(
            [
                f"Sinf: {_value(student.get('class_name'))}",
                f"#{student['position']}: {_value(student.get('name'))}",
                f"Yosh: {_value(student.get('age'))}",
                f"Jinsi: {_value(student.get('gender'))}",
                f"Telefon: {_value(student.get('phone'))}",
                *(
                    f"{key}: {_value(value)}"
                    for key, value in _extra_fields(student).items()
                ),
            ]
        )
        for student in data["students"]
    ]
    yield "Yuborilgan matnlar", [
        f"{_value(submission.get('created_at'))} | {_value(submission.get('text'))}"
        for submission in data["submissions"]
        if submission.get("text")
    ]
    yield "Tayyor Word fayllar", [
        f"{item['class_name']} | {item['created_at']} | {Path(item['file_path']).name}"
        for item in data["documents"]
    ]
    yield "To'lovlar", [
        f"#{item['id']} | {item['amount'] if item['amount'] is not None else 'summa kutilmoqda'} "
        f"| {item['credits']} ta Word | {item['status']} | {item['created_at']}"
        for item in data["payments"]
    ]
    yield "To'lov qarorlari", [
        f"#{item['payment_id']} | {item['action']} | Admin ID: {item['actor_id']} "
        f"| {item['created_at']}"
        for item in data["payment_events"]
    ]
    yield "Foydalanilgan promokodlar", [
        f"{item['code']} | {item['redeemed_at']}"
        for item in data["promos"]
    ]
    yield "Admin bergan Wordlar", [
        f"{item['uses']} ta | Admin ID: {item['admin_id']} | {item['created_at']}"
        for item in data["grants"]
    ]


def _image_items(data):
    seen = set()
    for student in data["students"]:
        path = student.get("photo_path")
        if path and path not in seen:
            seen.add(path)
            yield f"O'quvchi {student['position']} rasmi", path
    for submission in data["submissions"]:
        path = submission.get("file_path")
        if submission.get("kind") == "photo" and path and path not in seen:
            seen.add(path)
            yield "Yuborilgan rasm", path
    for payment in data["payments"]:
        path = payment.get("receipt_path")
        if path and path not in seen:
            seen.add(path)
            yield f"To'lov #{payment['id']} cheki", path


def create_word_report(data, output_path, private_root, shared_template=None):
    document = Document()
    document.add_heading("Foydalanuvchi hisoboti", 0)
    for heading, lines in _report_sections(data):
        document.add_heading(heading, level=1)
        if lines:
            for line in lines:
                document.add_paragraph(_safe_text(line))
        else:
            document.add_paragraph("Ma'lumot yo'q.")

    document.add_heading("Rasmlar va cheklar", level=1)
    has_images = False
    for caption, raw_path in _image_items(data):
        path = _allowed_file(raw_path, private_root, shared_template)
        if not path:
            continue
        image = _thumbnail(path)
        if not image:
            document.add_paragraph(
                _safe_text(f"{caption}: rasmni ochib bo'lmadi ({path.name}).")
            )
            continue
        document.add_paragraph(_safe_text(caption))
        document.add_picture(io.BytesIO(image), width=Inches(2))
        has_images = True
    if not has_images:
        document.add_paragraph("Rasm yoki chek mavjud emas.")
    document.save(output_path)


def create_pdf_report(data, output_path, private_root, shared_template=None):
    from reportlab import __file__ as reportlab_file
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import Image as PdfImage
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    fonts_dir = Path(reportlab_file).parent / "fonts"
    regular_path = fonts_dir / "Vera.ttf"
    bold_path = fonts_dir / "VeraBd.ttf"
    if not regular_path.is_file() or not bold_path.is_file():
        raise RuntimeError("PDF uchun reportlab Vera shriftlari topilmadi.")
    if "UserReport" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("UserReport", str(regular_path)))
        pdfmetrics.registerFont(TTFont("UserReport-Bold", str(bold_path)))

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="UserReportTitle",
        parent=styles["Title"],
        fontName="UserReport-Bold",
        alignment=TA_CENTER,
    ))
    styles.add(ParagraphStyle(
        name="UserReportHeading",
        parent=styles["Heading2"],
        fontName="UserReport-Bold",
        textColor=colors.HexColor("#174A6E"),
    ))
    styles.add(ParagraphStyle(
        name="UserReportBody",
        parent=styles["BodyText"],
        fontName="UserReport",
        leading=14,
        wordWrap="CJK",
    ))
    story = [Paragraph("Foydalanuvchi hisoboti", styles["UserReportTitle"]), Spacer(1, 12)]
    for heading, lines in _report_sections(data):
        story.append(Paragraph(html.escape(heading), styles["UserReportHeading"]))
        if not lines:
            story.append(Paragraph("Ma'lumot yo'q.", styles["UserReportBody"]))
            continue
        table = Table(
            [
                [Paragraph(html.escape(_safe_text(line)), styles["UserReportBody"])]
                for line in lines
            ],
            colWidths=[6.6 * inch],
            repeatRows=0,
        )
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F2F6F9")),
            ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#D7E0E7")),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D7E0E7")),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
            ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.extend([table, Spacer(1, 8)])

    story.append(Paragraph("Rasmlar va cheklar", styles["UserReportHeading"]))
    image_count = 0
    for caption, raw_path in _image_items(data):
        path = _allowed_file(raw_path, private_root, shared_template)
        if not path:
            continue
        image_data = _thumbnail(path)
        if not image_data:
            story.append(Paragraph(
                f"{html.escape(_safe_text(caption))}: rasmni ochib bo'lmadi "
                f"({html.escape(_safe_text(path.name))}).",
                styles["UserReportBody"],
            ))
            continue
        from reportlab.lib.utils import ImageReader

        reader = ImageReader(io.BytesIO(image_data))
        width, height = reader.getSize()
        scale = min(5.5 * inch / width, 5 * inch / height, 1)
        story.extend([
            Paragraph(html.escape(_safe_text(caption)), styles["UserReportBody"]),
            PdfImage(io.BytesIO(image_data), width=width * scale, height=height * scale),
            Spacer(1, 8),
        ])
        image_count += 1
    if not image_count:
        story.append(Paragraph("Rasm yoki chek mavjud emas.", styles["UserReportBody"]))

    SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=0.8 * inch,
        leftMargin=0.8 * inch,
        topMargin=0.7 * inch,
        bottomMargin=0.7 * inch,
        title="Foydalanuvchi hisoboti",
    ).build(story)


def create_user_report(data, format_name, output_path, private_root, shared_template=None):
    if format_name == "docx":
        create_word_report(data, output_path, private_root, shared_template)
    elif format_name == "pdf":
        create_pdf_report(data, output_path, private_root, shared_template)
    else:
        raise ValueError("Eksport formati docx yoki pdf bo'lishi kerak.")
