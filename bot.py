#!/usr/bin/env python3
import json
import logging
import os
import re
import sqlite3
import tempfile
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path

from docx import Document
from dotenv import load_dotenv
from telegram import (
    ForceReply,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
    Update,
)
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)
from telegram.error import TelegramError

from admin_exports import create_user_archive, create_user_report
from admin_user_tools import (
    clear_user_data,
    grant_word_uses,
    load_admin_user_data,
)
from database import Database
from sinf_tayyorlash import sarlavha_normal, sarlavha_turi, sinfni_yoz


BASE = Path(__file__).parent
load_dotenv(BASE / ".env")
DATA = Path(os.environ.get("DATA_DIR", BASE / "data"))
DATA.mkdir(parents=True, exist_ok=True)
PRIVATE = DATA / "private"
PRIVATE.mkdir(parents=True, exist_ok=True)
USER_FILE_ROOTS = (PRIVATE, DATA / "rasmlar", DATA / "natija")
DB = Database(DATA / "bot.sqlite3")
SHABLON = Path(os.environ.get("SHABLON", PRIVATE / "templates" / "6D.docx"))
WORD_PRICE = 5000


def now():
    return datetime.now(timezone.utc).isoformat()


def migrate_legacy_json():
    legacy = BASE / "data" / "oquvchilar.json"
    marker = DATA / "migration_v1.done"
    if marker.exists() or not legacy.exists():
        return
    try:
        data = json.loads(legacy.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        logging.warning("Legacy migration skipped: %s", exc)
        return
    imported = 0
    with DB.connect() as con:
        for class_name, rows in data.items():
            if not isinstance(rows, list):
                continue
            for position, row in enumerate(rows, start=1):
                owner = row.get("kim")
                if isinstance(owner, int):
                    con.execute(
                        "INSERT OR IGNORE INTO users(user_id,phone,username,registered_at,last_seen_at) "
                        "VALUES(?,?,?,?,?)",
                        (owner, None, "", now(), now()),
                    )
                rasm = row.get("rasm")
                if not rasm or not Path(rasm).is_file():
                    continue
                if isinstance(owner, int):
                    session = con.execute(
                        "SELECT id FROM roster_sessions WHERE user_id=? AND class_name=? ORDER BY id DESC LIMIT 1",
                        (owner, class_name),
                    ).fetchone()
                    if not session:
                        cur = con.execute(
                            "INSERT INTO roster_sessions(user_id,class_name,state,updated_at) "
                            "VALUES(?,?,?,?)",
                            (owner, class_name, "complete", now()),
                        )
                        session_id = cur.lastrowid
                    else:
                        session_id = session["id"]
                    con.execute(
                        "INSERT OR IGNORE INTO students(session_id,position,photo_path,name,age,gender,phone) "
                        "VALUES(?,?,?,?,?,?,?)",
                        (
                            session_id,
                            position,
                            rasm,
                            row.get("ism") or "",
                            "",
                            row.get("jins") or "",
                            row.get("tel") or "",
                        ),
                    )
                    imported += 1
    marker.write_text(
        json.dumps({"imported": imported, "at": now()}, ensure_ascii=False),
        encoding="utf-8",
    )


def id_set(name):
    return {int(x.strip()) for x in os.environ.get(name, "").split(",") if x.strip()}


def registered(user_id):
    with DB.connect() as con:
        row = con.execute("SELECT 1 FROM users WHERE user_id=? AND phone IS NOT NULL", (user_id,)).fetchone()
    return bool(row)


def get_active_session(user_id):
    with DB.connect() as con:
        return con.execute(
            "SELECT * FROM roster_sessions WHERE user_id=? ORDER BY id DESC LIMIT 1",
            (user_id,),
        ).fetchone()


def set_session(user_id, session_id=None, **fields):
    with DB.connect() as con:
        if session_id:
            if fields:
                names = ", ".join(f"{key}=?" for key in fields)
                con.execute(
                    f"UPDATE roster_sessions SET {names}, updated_at=? WHERE id=?",
                    [*fields.values(), now(), session_id],
                )
        else:
            cur = con.execute(
                "INSERT INTO roster_sessions(user_id,state,updated_at) VALUES(?,?,?)",
                (user_id, "awaiting_template", now()),
            )
            return cur.lastrowid


def credit_balance(user_id):
    with DB.connect() as con:
        row = con.execute("SELECT balance FROM credits WHERE user_id=?", (user_id,)).fetchone()
    return row["balance"] if row else 0


def approve_payment(db, payment_id, reviewer_id, amount):
    if amount < WORD_PRICE or amount > (2**63 - 1) or amount % WORD_PRICE:
        return "invalid_amount", 0, None
    word_count = amount // WORD_PRICE
    with db.connect() as con:
        con.execute("BEGIN IMMEDIATE")
        payment = con.execute(
            "SELECT user_id,status FROM payments WHERE id=?",
            (payment_id,),
        ).fetchone()
        if not payment:
            return "not_found", 0, None
        if payment["status"] != "pending":
            return "already_reviewed", 0, payment["user_id"]
        updated = con.execute(
            "UPDATE payments SET status='approved',amount=?,credits=?,reviewed_by=?,"
            "reviewed_at=? WHERE id=? AND status='pending'",
            (amount, word_count, reviewer_id, now(), payment_id),
        )
        if updated.rowcount != 1:
            return "already_reviewed", 0, payment["user_id"]
        con.execute(
            "INSERT INTO credits(user_id,balance) VALUES(?,?) "
            "ON CONFLICT(user_id) DO UPDATE SET balance=balance+excluded.balance",
            (payment["user_id"], word_count),
        )
        con.execute(
            "INSERT INTO payment_events(payment_id,actor_id,action,created_at) "
            "VALUES(?,?,?,?)",
            (payment_id, reviewer_id, "approved", now()),
        )
        return "approved", word_count, payment["user_id"]


def shared_template_enabled():
    return DB.get_setting("shared_template_enabled", "0") == "1"


def admin_menu_markup():
    template_enabled = shared_template_enabled()
    template_label = (
        "✅ Umumiy shablon: yoqilgan"
        if template_enabled
        else "⚪ Umumiy shablon: o'chirilgan"
    )
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("👥 Foydalanuvchilar", callback_data="admin:users:0")],
        [InlineKeyboardButton("📊 Statistika", callback_data="admin:stats")],
        [InlineKeyboardButton("💳 To'lovlar", callback_data="admin:payments")],
        [InlineKeyboardButton("🎟 Promokod yaratish", callback_data="admin:promo")],
        [InlineKeyboardButton(template_label, callback_data="admin:template")],
    ])


def admin_user_markup(user_id):
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("⬇ Word hisobot", callback_data=f"admin:export:docx:{user_id}"),
            InlineKeyboardButton("⬇ PDF hisobot", callback_data=f"admin:export:pdf:{user_id}"),
        ],
        [InlineKeyboardButton("➕ Word berish", callback_data=f"admin:grant:{user_id}")],
        [InlineKeyboardButton("🧹 Ma'lumotlarni tozalash", callback_data=f"admin:clear:{user_id}")],
        [InlineKeyboardButton("⬅ Foydalanuvchilar", callback_data="admin:users:0")],
    ])


def create_promo_code(db, code, credits, max_uses, created_by):
    with db.connect() as con:
        con.execute(
            "INSERT INTO promo_codes(code,credits,max_uses,created_by,created_at) "
            "VALUES(?,?,?,?,?)",
            (code, credits, max_uses, created_by, now()),
        )


def redeem_promo_code(db, user_id, code):
    with db.connect() as con:
        con.execute("BEGIN IMMEDIATE")
        promo = con.execute(
            "SELECT credits,max_uses,uses FROM promo_codes WHERE code=?",
            (code,),
        ).fetchone()
        if not promo:
            return "not_found", 0
        already_used = con.execute(
            "SELECT 1 FROM promo_redemptions WHERE code=? AND user_id=?",
            (code, user_id),
        ).fetchone()
        if already_used:
            return "already_used", 0
        if promo["uses"] >= promo["max_uses"]:
            return "exhausted", 0
        con.execute(
            "INSERT INTO promo_redemptions(code,user_id,redeemed_at) VALUES(?,?,?)",
            (code, user_id, now()),
        )
        con.execute("UPDATE promo_codes SET uses=uses+1 WHERE code=?", (code,))
        con.execute(
            "INSERT INTO credits(user_id,balance) VALUES(?,?) "
            "ON CONFLICT(user_id) DO UPDATE SET balance=balance+excluded.balance",
            (user_id, promo["credits"]),
        )
        return "ok", promo["credits"]


def is_contact_owner(update):
    contact = update.message.contact
    return contact and contact.user_id == update.effective_user.id


async def require_private_chat(update):
    if update.effective_chat and update.effective_chat.type == "private":
        return True
    if update.effective_message:
        await update.effective_message.reply_text(
            "Shaxsiy ma'lumotlar va rasmlar uchun botga shaxsiy chatdan yozing."
        )
    return False


def main_menu():
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("📝 Word tayyorlash"), KeyboardButton("📁 Fayllarim")],
            [KeyboardButton("👤 Profil"), KeyboardButton("💳 To'lov")],
            [KeyboardButton("🎟 Promokod"), KeyboardButton("➕ O'quvchi qo'shish")],
            [KeyboardButton("❌ Bekor qilish")],
        ],
        resize_keyboard=True,
    )


async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    ctx.user_data.pop("awaiting_promo_code", None)
    user = update.effective_user
    with DB.connect() as con:
        row = con.execute("SELECT phone FROM users WHERE user_id=?", (user.id,)).fetchone()
    if row and row["phone"]:
        await update.message.reply_text(
            "👋 Xush kelibsiz!\n\n"
            "📝 Word ro'yxati tayyorlash\n"
            "📁 Tayyor fayllarni ko'rish\n"
            "👤 Profil va foydalanishlar soni\n"
            "💳 Word tayyorlash: 5 000 so'm / fayl\n"
            "🎟 Promokod: /promo KOD\n"
            "➕ Kech qolgan o'quvchilarni qo'shish: /qoshimcha\n"
            "❌ Jarayonni bekor qilish: /bekor\n\n"
            f"📊 Qolgan Word tayyorlashlar: {credit_balance(user.id)} ta.",
            reply_markup=main_menu(),
        )
        return
    kb = ReplyKeyboardMarkup(
        [[KeyboardButton("📱 Telefon raqamini ulashish", request_contact=True)]],
        resize_keyboard=True,
        one_time_keyboard=True,
    )
    await update.message.reply_text(
        "👋 Xush kelibsiz!\nBotdan foydalanishni boshlash uchun Telegram telefon raqamingizni ulashing.",
        reply_markup=kb,
    )


async def contact(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    user = update.effective_user
    if not is_contact_owner(update):
        await update.message.reply_text("Faqat o'zingizning Telegram telefon raqamingizni ulashing.")
        return
    phone = update.message.contact.phone_number
    with DB.connect() as con:
        con.execute(
            "INSERT INTO users(user_id,phone,username,registered_at,last_seen_at) "
            "VALUES(?,?,?,?,?) ON CONFLICT(user_id) DO UPDATE "
            "SET phone=excluded.phone, username=excluded.username, last_seen_at=excluded.last_seen_at",
            (user.id, phone, user.username or "", now(), now()),
        )
        con.execute(
            "INSERT OR IGNORE INTO credits(user_id,balance) VALUES(?,0)",
            (user.id,),
        )
    await update.message.reply_text(
        "✅ Ro'yxatdan o'tdingiz!\n\n"
        "📝 Word ro'yxati tayyorlang — har bir fayl 5 000 so'm.\n"
        "📁 Tayyor fayllar faqat sizga ko'rinadi.\n"
        "Kerakli bo'limni pastdagi menyudan tanlang.",
        reply_markup=main_menu(),
    )


async def profile(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    user = update.effective_user
    with DB.connect() as con:
        record = con.execute(
            "SELECT phone,username,free_additions_used FROM users WHERE user_id=?",
            (user.id,),
        ).fetchone()
        docs = con.execute(
            "SELECT COUNT(*) FROM documents WHERE user_id=?",
            (user.id,),
        ).fetchone()[0]
    if not record or not record["phone"]:
        await start(update, ctx)
        return
    free_used = record["free_additions_used"]
    await update.message.reply_text(
        "👤 Profil\n\n"
        f"Username: @{record['username'] or '-'}\n"
        f"📱 Telefon: {record['phone']}\n"
        f"📝 Qolgan Word tayyorlashlar: {credit_balance(user.id)} ta\n"
        f"📁 Tayyorlangan fayllar: {docs} ta\n"
        f"🎁 Bepul qo'shilgan o'quvchilar: {free_used}/3",
        reply_markup=main_menu(),
    )


async def promo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    if len(ctx.args) != 1:
        await update.message.reply_text("Promokodni ishlatish: /promo KOD")
        return
    await apply_promo_code(update, ctx, ctx.args[0])


async def apply_promo_code(update: Update, ctx: ContextTypes.DEFAULT_TYPE, raw_code: str):
    user = update.effective_user
    if not registered(user.id):
        await start(update, ctx)
        return
    code = raw_code.strip().upper()
    status, credits = redeem_promo_code(DB, user.id, code)
    if status == "not_found":
        await update.message.reply_text("❌ Bunday promokod topilmadi.")
    elif status == "exhausted":
        await update.message.reply_text("❌ Bu promokodning foydalanish limiti tugagan.")
    elif status == "already_used":
        await update.message.reply_text("ℹ️ Siz bu promokodni oldin ishlatgansiz.")
    else:
        ctx.user_data.pop("awaiting_promo_code", None)
        ctx.user_data.pop("awaiting_receipt", None)
        await update.message.reply_text(
            f"🎉 Promokod qabul qilindi!\n"
            f"🏆 Mukofotingiz: {credits} ta bepul Word tayyorlash.\n"
            f"Qolgan Word tayyorlashlar: {credit_balance(user.id)} ta.",
            reply_markup=main_menu(),
        )


async def create_promo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in id_set("ADMIN_IDS"):
        await update.message.reply_text("Bu buyruq faqat asosiy administrator uchun.")
        return
    if update.effective_chat.type != "private":
        await update.message.reply_text("Promokodni bot bilan shaxsiy chatda yarating.")
        return
    if len(ctx.args) != 3:
        await update.message.reply_text(
            "Format: /promo_yarat KOD WORD_SONI FOYDALANISH_LIMI\n"
            "Masalan: /promo_yarat MAKTAB10 2 50"
        )
        return
    code = ctx.args[0].strip().upper()
    if not re.fullmatch(r"[A-Z0-9_-]{3,32}", code):
        await update.message.reply_text("Kod 3-32 ta lotin harfi, raqam, tire yoki pastki chiziqdan iborat bo'lsin.")
        return
    try:
        credits, max_uses = int(ctx.args[1]), int(ctx.args[2])
    except ValueError:
        await update.message.reply_text("Word soni va foydalanish limiti musbat son bo'lishi kerak.")
        return
    if credits < 1 or max_uses < 1:
        await update.message.reply_text("Word soni va foydalanish limiti kamida 1 bo'lishi kerak.")
        return
    try:
        create_promo_code(DB, code, credits, max_uses, update.effective_user.id)
    except sqlite3.IntegrityError:
        await update.message.reply_text("Bu promokod avval yaratilgan.")
        return
    await update.message.reply_text(
        f"Promokod yaratildi: {code}\nHar foydalanuvchiga: {credits} ta Word tayyorlash\n"
        f"Jami foydalanish limiti: {max_uses}"
    )


async def roster_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    user = update.effective_user
    if not registered(user.id):
        await start(update, ctx)
        return
    active = get_active_session(user.id)
    if active and active["state"] not in ("complete", "cancelled"):
        await update.message.reply_text(
            "Sizda tugallanmagan ro'yxat bor. Uni davom ettiring yoki /bekor qiling."
        )
        return
    balance = credit_balance(user.id)
    if balance < 1:
        await show_payment_instructions(update, ctx)
        return
    ctx.user_data.pop("awaiting_receipt", None)
    session_id = set_session(user.id)
    if shared_template_enabled():
        if not SHABLON.is_file():
            set_session(user.id, session_id, state="cancelled")
            await update.message.reply_text(
                "Umumiy shablon topilmadi. Administrator shablon faylini tekshirsin."
            )
            return
        set_session(
            user.id,
            session_id,
            template_path=str(SHABLON),
            state="awaiting_class",
        )
        await update.message.reply_text(
            "Umumiy shablon tanlandi. Sinfni yozing, masalan: 7A.",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("/bekor")]], resize_keyboard=True),
        )
        return
    await update.message.reply_text(
        f"Sizda {balance} ta Word tayyorlash qolgan. Har bir tayyor faylda bittadan foydalanish "
        "hisoblanadi.\nAvval Word namuna/shablonini .docx formatida yuboring.",
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("/bekor")]], resize_keyboard=True),
    )


async def qoshimcha(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    user = update.effective_user
    if not registered(user.id):
        await start(update, ctx)
        return
    active = get_active_session(user.id)
    if active and active["state"] not in ("complete", "cancelled"):
        await update.message.reply_text("Avval joriy ro'yxatni tugating yoki /bekor qiling.")
        return
    await show_additional_document_options(update, user.id, 0)


async def show_additional_document_options(update_or_query, user_id: int, page: int):
    page_size = 10
    with DB.connect() as con:
        total = con.execute(
            "SELECT COUNT(*) FROM documents d "
            "JOIN roster_sessions s ON s.id=d.session_id AND s.user_id=d.user_id "
            "WHERE d.user_id=? AND d.session_id IS NOT NULL "
            "AND COALESCE(s.template_path,'')!=''",
            (user_id,),
        ).fetchone()[0]
        documents = con.execute(
            "SELECT d.id,d.class_name,d.created_at FROM documents d "
            "JOIN roster_sessions s ON s.id=d.session_id AND s.user_id=d.user_id "
            "WHERE d.user_id=? AND d.session_id IS NOT NULL "
            "AND COALESCE(s.template_path,'')!='' "
            "ORDER BY d.created_at DESC,d.id DESC LIMIT ? OFFSET ?",
            (user_id, page_size, page * page_size),
        ).fetchall()
    if not documents:
        message = "Qo'shimcha kiritish uchun botda tayyorlangan Word fayl topilmadi."
        if update_or_query.callback_query:
            await update_or_query.callback_query.answer()
            await update_or_query.callback_query.edit_message_text(message)
        else:
            await update_or_query.message.reply_text(message)
        return
    buttons = [
        [InlineKeyboardButton(
            f"{doc['class_name']}.docx · {doc['created_at'][:10]} · #{doc['id']}",
            callback_data=f"addstudent:select:{doc['id']}",
        )]
        for doc in documents
    ]
    navigation = []
    if page > 0:
        navigation.append(
            InlineKeyboardButton("⬅ Oldingi", callback_data=f"addstudent:page:{page - 1}")
        )
    if (page + 1) * page_size < total:
        navigation.append(
            InlineKeyboardButton("Keyingi ➡", callback_data=f"addstudent:page:{page + 1}")
        )
    if navigation:
        buttons.append(navigation)
    text = "Qaysi Word faylga o'quvchi qo'shmoqchisiz?"
    markup = InlineKeyboardMarkup(buttons)
    if update_or_query.callback_query:
        await update_or_query.callback_query.answer()
        await update_or_query.callback_query.edit_message_text(text, reply_markup=markup)
    else:
        await update_or_query.message.reply_text(text, reply_markup=markup)


async def start_additional_student(
    ctx: ContextTypes.DEFAULT_TYPE,
    user_id: int,
    document_id: int,
    query=None,
):
    active = get_active_session(user_id)
    if active and active["state"] not in ("complete", "cancelled"):
        message = "Avval joriy ro'yxatni tugating yoki /bekor qiling."
        if query:
            await query.answer(message, show_alert=True)
        else:
            await ctx.bot.send_message(chat_id=user_id, text=message)
        return
    with DB.connect() as con:
        document = con.execute(
            "SELECT * FROM documents WHERE id=? AND user_id=?",
            (document_id, user_id),
        ).fetchone()
        user_row = con.execute(
            "SELECT free_additions_used FROM users WHERE user_id=?",
            (user_id,),
        ).fetchone()
        base_session = None
        students = []
        if document and document["session_id"]:
            base_session = con.execute(
                "SELECT * FROM roster_sessions WHERE id=? AND user_id=?",
                (document["session_id"], user_id),
            ).fetchone()
            if base_session:
                students = con.execute(
                    "SELECT * FROM students WHERE session_id=? ORDER BY position",
                    (base_session["id"],),
                ).fetchall()

    if not document or not document["session_id"]:
        message = "Qo'shimcha kiritish uchun botda tayyorlangan Word faylni tanlang."
        if query:
            await query.answer(message, show_alert=True)
        else:
            await ctx.bot.send_message(chat_id=user_id, text=message)
        return
    if not base_session or not base_session["template_path"]:
        message = "Asl shablon topilmadi; bu faylga qo'shimcha kiritib bo'lmaydi."
        if query:
            await query.answer(message, show_alert=True)
        else:
            await ctx.bot.send_message(chat_id=user_id, text=message)
        return
    if not Path(document["file_path"]).is_file():
        message = "Avvalgi Word fayli topilmadi."
        if query:
            await query.answer(message, show_alert=True)
        else:
            await ctx.bot.send_message(chat_id=user_id, text=message)
        return
    free_left = max(0, 3 - user_row["free_additions_used"])
    is_free = free_left > 0
    max_additions = free_left if is_free else 3
    balance = credit_balance(user_id)
    if not is_free and balance < 1:
        message = (
            "🎁 Bepul qo'shimcha o'quvchi limitingiz tugadi. "
            "Yana Word tayyorlash uchun to'lov qiling."
        )
        if query:
            await query.answer(message, show_alert=True)
        await ctx.bot.send_message(chat_id=user_id, text=message)
        await send_payment_instructions_to_user(ctx, user_id)
        return
    with DB.connect() as con:
        cur = con.execute(
            "INSERT INTO roster_sessions("
            "user_id,class_name,template_path,state,current_student,correction_document_id,"
            "correction_base_count,correction_is_free,correction_max_additions,updated_at"
            ") VALUES(?,?,?,?,?,?,?,?,?,?)",
            (
                user_id,
                document["class_name"],
                base_session["template_path"],
                "collecting",
                len(students) + 1,
                document["id"],
                len(students),
                int(is_free),
                max_additions,
                now(),
            ),
        )
        correction_session_id = cur.lastrowid
        for student in students:
            con.execute(
                "INSERT INTO students("
                "session_id,position,photo_path,name,age,gender,phone,extra_json"
                ") VALUES(?,?,?,?,?,?,?,?)",
                (
                    correction_session_id,
                    student["position"],
                    student["photo_path"],
                    student["name"],
                    student["age"],
                    student["gender"],
                    student["phone"],
                    student["extra_json"],
                ),
            )
    if query:
        await query.answer()
        await query.edit_message_text(f"✅ {document['class_name']}.docx tanlandi.")
    if is_free:
        prompt = (
            f"Bu bepul qo'shimcha. Umumiy bepul limitingizdan {free_left} ta o'quvchi qolgan. "
            f"Eng ko'pi bilan {max_additions} ta o'quvchi qo'shing. "
            "Har birining rasmini captiondagi ma'lumotlari bilan yuboring, so'ng /yakunlash."
        )
    else:
        prompt = (
            f"Bu yangilash 1 ta Word tayyorlashdan foydalanadi. Qolgan foydalanish: {balance} ta. "
            f"Eng ko'pi bilan {max_additions} ta o'quvchi qo'shing; har birining rasmini caption bilan "
            "yuboring, so'ng /yakunlash."
        )
    await ctx.bot.send_message(
        chat_id=user_id,
        text=prompt,
        reply_markup=ReplyKeyboardMarkup([[KeyboardButton("/bekor")]], resize_keyboard=True),
    )


async def add_student_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    if not await require_private_chat(update):
        return
    user_id = update.effective_user.id
    if not registered(user_id):
        await query.answer("Avval ro'yxatdan o'ting.", show_alert=True)
        return
    parts = query.data.split(":")
    if len(parts) != 3:
        await query.answer("So'rov noto'g'ri.", show_alert=True)
        return
    if parts[1] == "page":
        try:
            page = int(parts[2])
        except ValueError:
            await query.answer("Bo'lim raqami noto'g'ri.", show_alert=True)
            return
        if page < 0:
            await query.answer("Bo'lim raqami noto'g'ri.", show_alert=True)
            return
        await show_additional_document_options(update, user_id, page)
        return
    if parts[1] != "select":
        await query.answer("So'rov noto'g'ri.", show_alert=True)
        return
    try:
        document_id = int(parts[2])
    except ValueError:
        await query.answer("Fayl raqami noto'g'ri.", show_alert=True)
        return
    await start_additional_student(ctx, user_id, document_id, query)


async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    user = update.effective_user
    ctx.user_data.pop("awaiting_receipt", None)
    ctx.user_data.pop("awaiting_promo_code", None)
    with DB.connect() as con:
        session = con.execute(
            "SELECT id FROM roster_sessions WHERE user_id=? AND state!='complete' ORDER BY id DESC LIMIT 1",
            (user.id,),
        ).fetchone()
        if session:
            con.execute("UPDATE roster_sessions SET state='cancelled',updated_at=? WHERE id=?", (now(), session["id"]))
    await update.message.reply_text("Ro'yxat sessiyasi bekor qilindi.", reply_markup=main_menu())


async def document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    user = update.effective_user
    session = get_active_session(user.id)
    if not session or session["state"] not in ("awaiting_template",):
        await update.message.reply_text("Avval /royxatga_olish buyrug'ini bosing.")
        return
    if not document_file(update):
        await update.message.reply_text("Faqat .docx Word namunasini yuboring.")
        return
    file = await ctx.bot.get_file(update.message.document.file_id)
    dest = PRIVATE / f"templates/{uuid.uuid4().hex}.docx"
    dest.parent.mkdir(parents=True, exist_ok=True)
    await file.download_to_drive(dest)
    set_session(user.id, session["id"], template_path=str(dest), state="awaiting_class")
    await update.message.reply_text("Shablon qabul qilindi. Sinfni yozing, masalan: 7A.")


def document_file(update):
    doc = update.message.document
    return doc and (doc.file_name or "").lower().endswith(".docx")


def session_headers(session):
    template = Path(session["template_path"])
    doc = Document(template)
    if not doc.tables or not doc.tables[0].rows:
        raise RuntimeError(f"{template.name} faylida sarlavhali jadval topilmadi.")
    return [cell.text for cell in doc.tables[0].rows[0].cells]


def parse_student_details(raw_text, student, headers=None):
    """Parse labeled values or assign unlabelled lines in template column order."""
    values = {key: student[key] or "" for key in ("name", "age", "gender", "phone")}
    try:
        extra = json.loads(student["extra_json"] or "{}")
    except (KeyError, TypeError, json.JSONDecodeError):
        extra = {}
    headers = headers or ["F.I.SH", "Telefon raqam"]
    editable_headers = [
        header for header in headers
        if sarlavha_turi(header) not in ("tr", "class_number", "class_letter", "photo")
    ]

    def store(header, value):
        category = sarlavha_turi(header)
        if category in values:
            values[category] = value
        else:
            extra[header] = value

    unlabelled = []
    for line in raw_text.splitlines():
        line = line.strip()
        if not line:
            continue
        if ":" in line:
            label, value = line.split(":", 1)
            label_key = sarlavha_normal(label)
            matching_header = next(
                (header for header in editable_headers if sarlavha_normal(header) == label_key),
                None,
            )
            if matching_header and value.strip():
                store(matching_header, value.strip())
                continue
            category = sarlavha_turi(label)
            matching_header = next(
                (header for header in editable_headers if sarlavha_turi(header) == category),
                None,
            )
            if matching_header and category != "extra" and value.strip():
                store(matching_header, value.strip())
                continue
            if value.strip():
                extra[label.strip()] = value.strip()
                continue
        unlabelled.append(line)

    for line in unlabelled:
        normalized_line = sarlavha_normal(line).replace("'", "")
        gender_header = next(
            (header for header in editable_headers if sarlavha_turi(header) == "gender"),
            None,
        )
        if gender_header and normalized_line in {
            "erkak", "ogil", "o g il", "male", "man", "boy",
            "ayol", "qiz", "female", "woman", "girl",
        }:
            store(gender_header, line)
            continue
        phone_header = next(
            (header for header in editable_headers if sarlavha_turi(header) == "phone"),
            None,
        )
        if phone_header and not values["phone"] and re.search(r"\d{7,15}", line):
            store(phone_header, line)
            continue
        target = next(
            (
                header for header in editable_headers
                if (
                    not values.get(sarlavha_turi(header), "")
                    if sarlavha_turi(header) in values
                    else not extra.get(header, "")
                )
            ),
            None,
        )
        if target:
            store(target, line)
        else:
            return None
    values["extra"] = extra
    return values


def image_file_id(message):
    if message.photo:
        return message.photo[-1].file_id, ".jpg"
    document = message.document
    if not document or not (document.mime_type or "").startswith("image/"):
        return None
    suffix = Path(document.file_name or "").suffix.lower()
    if suffix not in (".jpg", ".jpeg", ".png"):
        return None
    return document.file_id, suffix


async def prompt_for_student_details(update, ctx, session, student, position):
    if not student["name"]:
        await update.message.reply_text("Ism-familiyani yozing (rasm izohida yuborgan bo'lsangiz, qayta yozish shart emas).")
        return
    headers = session_headers(session)
    phone_required = any(sarlavha_turi(header) == "phone" for header in headers)
    if phone_required and not student["phone"]:
        await update.message.reply_text("Telefon raqamini yozing.")
        return
    set_session(session["user_id"], session["id"], current_student=position + 1)
    if session["correction_document_id"] and position + 1 > (
        session["correction_base_count"] + session["correction_max_additions"]
    ):
        await update.message.reply_text(
            "Qo'shimcha o'quvchilar limiti tugadi. Word faylni yangilash uchun /yakunlash ni bosing."
        )
        return
    await update.message.reply_text(
        f"{position}-o'quvchi saqlandi. Keyingi o'quvchining rasmini captiondagi "
        "ma'lumotlari bilan, shablon ustunlari tartibida yuboring.",
        reply_markup=ReplyKeyboardMarkup(
            [[KeyboardButton("/yakunlash"), KeyboardButton("/bekor")]],
            resize_keyboard=True,
        ),
    )


async def text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    user = update.effective_user
    text_value = (update.message.text or "").strip()
    grant_prompts = ctx.chat_data.get("admin_grant_prompts", {})
    reply_to = update.message.reply_to_message
    if reply_to and reply_to.message_id in grant_prompts:
        if user.id not in id_set("ADMIN_IDS"):
            await update.message.reply_text("Bu amal faqat asosiy administrator uchun.")
            return
        await admin_grant_reply(
            update,
            ctx,
            grant_prompts[reply_to.message_id],
            reply_to.message_id,
        )
        return
    menu_actions = {
        "📝 Word tayyorlash": roster_start,
        "📁 Fayllarim": my_files,
        "👤 Profil": profile,
        "💳 To'lov": purchase,
        "➕ O'quvchi qo'shish": qoshimcha,
        "❌ Bekor qilish": cancel,
    }
    if text_value in menu_actions:
        ctx.user_data.pop("awaiting_promo_code", None)
        await menu_actions[text_value](update, ctx)
        return
    if text_value == "🎟 Promokod":
        ctx.user_data["awaiting_promo_code"] = True
        await update.message.reply_text("🎟 Promokod matnini yuboring:")
        return
    if ctx.user_data.get("awaiting_promo_code"):
        await apply_promo_code(update, ctx, text_value)
        return
    if text_value.startswith("/"):
        return
    if not registered(user.id):
        await update.message.reply_text("Avval telefon raqamingizni ulashing: /start")
        return
    session = get_active_session(user.id)
    if not session:
        await update.message.reply_text("Ro'yxat uchun /royxatga_olish buyrug'ini yuboring.")
        return

    if session["state"] == "awaiting_class":
        if not re.fullmatch(r"\d{1,2}[A-Za-zА-Яа-я]{1,2}", text_value):
            await update.message.reply_text("Sinfni 7A yoki 3YE ko'rinishida yozing.")
            return
        set_session(user.id, session["id"], class_name=text_value.upper(), state="collecting", current_student=1)
        await update.message.reply_text(
            "Har bir o'quvchining rasmini izohi (caption) bilan birga yuboring. "
            "Sinf raqami va harfini yozmang — ular tanlangan sinfdan olinadi. "
            "Qolgan ma'lumotlarni Word shablonidagi ustunlar ketma-ketligida, "
            "har birini yangi qatordan yozing; ular tegishli ustunlarga joylanadi.\n"
            "Masalan, shablonda ism, yosh, telefon va jins ustunlari shu tartibda bo'lsa, "
            "captionga ham ism, yosh, telefon va jinsni shu ketma-ketlikda kiriting.\n"
            "Ro'yxat tugagach /yakunlash buyrug'ini yuboring."
        )
        return

    if session["state"] != "collecting":
        return

    position = session["current_student"]
    with DB.connect() as con:
        student = con.execute(
            "SELECT * FROM students WHERE session_id=? AND position=?",
            (session["id"], position),
        ).fetchone()
        if not student:
            await update.message.reply_text("Avval shu o'quvchining rasmini yuboring.")
            return
    try:
        headers = session_headers(session)
    except (OSError, ValueError, RuntimeError) as exc:
        await update.message.reply_text(f"Word shablon jadvali o'qilmadi: {exc}")
        return
    parsed = parse_student_details(text_value, student, headers)
    if parsed is None:
        await update.message.reply_text("Ma'lumotni qatorlarga ajratib yuboring: ism, telefon raqami.")
        return
    with DB.connect() as con:
        con.execute(
            "UPDATE students SET name=?,age=?,gender=?,phone=?,extra_json=? "
            "WHERE session_id=? AND position=?",
            (
                parsed["name"],
                parsed["age"],
                parsed["gender"],
                parsed["phone"],
                json.dumps(parsed["extra"], ensure_ascii=False),
                session["id"],
                position,
            ),
        )
        con.execute(
            "INSERT INTO submissions(user_id,kind,text,created_at) VALUES(?,?,?,?)",
            (user.id, "student_details", text_value, now()),
        )
        student = con.execute(
            "SELECT * FROM students WHERE session_id=? AND position=?",
            (session["id"], position),
        ).fetchone()
    await prompt_for_student_details(update, ctx, session, student, position)


async def photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    user = update.effective_user
    if not registered(user.id):
        await update.message.reply_text("Avval telefon raqamingizni ulashing: /start")
        return
    session = get_active_session(user.id)
    if not session or session["state"] != "collecting":
        await update.message.reply_text("Avval /royxatga_olish buyrug'i bilan sessiyani boshlang.")
        return
    image = image_file_id(update.message)
    if not image:
        await update.message.reply_text(
            "Rasmni Telegram fotosi yoki .jpg/.jpeg/.png fayli ko'rinishida yuboring."
        )
        return
    position = session["current_student"]
    if session["correction_document_id"] and position > (
        session["correction_base_count"] + session["correction_max_additions"]
    ):
        await update.message.reply_text(
            "Qo'shimcha o'quvchilar limiti tugadi. Word faylni yangilash uchun /yakunlash ni bosing."
        )
        return
    file = await ctx.bot.get_file(image[0])
    dest = PRIVATE / "photos" / f"{uuid.uuid4().hex}{image[1]}"
    dest.parent.mkdir(parents=True, exist_ok=True)
    await file.download_to_drive(dest)
    caption = (update.message.caption or "").strip()
    with DB.connect() as con:
        existing = con.execute(
            "SELECT * FROM students WHERE session_id=? AND position=?",
            (session["id"], position),
        ).fetchone()
        if existing:
            con.execute(
                "UPDATE students SET photo_path=? WHERE session_id=? AND position=?",
                (str(dest), session["id"], position),
            )
        else:
            con.execute(
                "INSERT INTO students(session_id,position,photo_path) VALUES(?,?,?)",
                (session["id"], position, str(dest)),
            )
        con.execute(
            "INSERT INTO submissions(user_id,kind,text,file_path,created_at) VALUES(?,?,?,?,?)",
            (user.id, "photo", caption or None, str(dest), now()),
        )
        student = con.execute(
            "SELECT * FROM students WHERE session_id=? AND position=?",
            (session["id"], position),
        ).fetchone()
    if caption:
        try:
            headers = session_headers(session)
        except (OSError, ValueError, RuntimeError) as exc:
            await update.message.reply_text(f"Word shablon jadvali o'qilmadi: {exc}")
            return
        parsed = parse_student_details(caption, student, headers)
        if parsed is None:
            await update.message.reply_text("Rasm qabul qilindi, lekin izohni tushunmadim. Ism va telefonni alohida qatorlarda yozing.")
            return
        with DB.connect() as con:
            con.execute(
                "UPDATE students SET name=?,age=?,gender=?,phone=?,extra_json=? "
                "WHERE session_id=? AND position=?",
                (
                    parsed["name"],
                    parsed["age"],
                    parsed["gender"],
                    parsed["phone"],
                    json.dumps(parsed["extra"], ensure_ascii=False),
                    session["id"],
                    position,
                ),
            )
            student = con.execute(
                "SELECT * FROM students WHERE session_id=? AND position=?",
                (session["id"], position),
            ).fetchone()
    await prompt_for_student_details(update, ctx, session, student, position)


def sinf_normal(name):
    match = re.fullmatch(r"(\d{1,2})([A-Za-zА-Яа-я]{1,2})", name)
    return (match.group(1), match.group(2).upper()) if match else None


async def finish(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    user = update.effective_user
    session = get_active_session(user.id)
    if not session or session["state"] != "collecting":
        await update.message.reply_text("Yakunlanadigan ro'yxat sessiyasi yo'q.")
        return
    with DB.connect() as con:
        students = con.execute(
            "SELECT * FROM students WHERE session_id=? ORDER BY position",
            (session["id"],),
        ).fetchall()
    try:
        headers = session_headers(session)
    except (OSError, ValueError, RuntimeError) as exc:
        await update.message.reply_text(f"Word shablon jadvali o'qilmadi: {exc}")
        return
    phone_required = any(sarlavha_turi(header) == "phone" for header in headers)
    if not students:
        await update.message.reply_text("Hali to'liq o'quvchi qo'shilmagan.")
        return
    correction = session["correction_document_id"] is not None
    added_count = len(students) - session["correction_base_count"] if correction else 0
    if correction and added_count < 1:
        await update.message.reply_text("Hali yangi o'quvchi qo'shilmadi.")
        return
    if correction and added_count > session["correction_max_additions"]:
        await update.message.reply_text("Bu yangilash uchun ruxsat etilgan o'quvchi sonidan oshib ketdi.")
        return
    incomplete = [
        student["position"]
        for student in students
        if (
            not student["photo_path"]
            or not student["name"]
            or (phone_required and not student["phone"])
        )
    ]
    if incomplete:
        await update.message.reply_text(
            "Quyidagi o'quvchilarning rasm, ism yoki telefon raqami yetishmaydi: "
            + ", ".join(map(str, incomplete))
            + ". Ma'lumotlarni to'ldiring yoki /bekor qiling."
        )
        return
    is_free_correction = correction and bool(session["correction_is_free"])
    balance = credit_balance(user.id)
    if not is_free_correction and balance < 1:
        await update.message.reply_text(
            "ℹ️ Word tayyorlash uchun foydalanish qolmagan. Fayl tayyorlanmadi."
        )
        await show_payment_instructions(update, ctx)
        return

    sinf = sinf_normal(session["class_name"])
    if not sinf:
        await update.message.reply_text("Sinf nomi noto'g'ri.")
        return
    template = Path(session["template_path"])
    if not template.is_file():
        await update.message.reply_text("Shablon fayli topilmadi. Ro'yxatni qaytadan boshlang.")
        return
    rows = [
        {
            "ism": s["name"],
            "rasm": Path(s["photo_path"]),
            "tel": s["phone"] or "",
            "joy": s["age"] or "",
            "yosh": s["age"] or "",
            "jins": s["gender"] or "",
            "extra": json.loads(s["extra_json"] or "{}"),
        }
        for s in students
    ]
    output_dir = PRIVATE / "documents" / str(user.id)
    output_dir.mkdir(parents=True, exist_ok=True)
    output = sinfni_yoz(template, sinf, rows, output_dir, False)
    with DB.connect() as con:
        if not is_free_correction:
            charged = con.execute(
                "UPDATE credits SET balance=balance-1 WHERE user_id=? AND balance>=1",
                (user.id,),
            )
            if charged.rowcount != 1:
                raise RuntimeError("Word tayyorlashlar soni o'zgargani sabab foydalanish qayd etilmadi.")
        if correction:
            if is_free_correction:
                free_update = con.execute(
                    "UPDATE users SET free_additions_used=free_additions_used+? "
                    "WHERE user_id=? AND free_additions_used+?<=3",
                    (added_count, user.id, added_count),
                )
                if free_update.rowcount != 1:
                    raise RuntimeError("Bepul qo'shimchalar limiti o'zgardi; Word yangilanmadi.")
            updated_document = con.execute(
                "UPDATE documents SET file_path=?,session_id=?,"
                "free_additions_used=free_additions_used+? WHERE id=? AND user_id=?",
                (
                    str(output),
                    session["id"],
                    added_count if is_free_correction else 0,
                    session["correction_document_id"],
                    user.id,
                ),
            )
            if updated_document.rowcount != 1:
                raise RuntimeError("Yangilanadigan Word yozuvi topilmadi.")
        else:
            con.execute(
                "INSERT INTO documents(user_id,class_name,file_path,session_id,created_at) "
                "VALUES(?,?,?,?,?)",
                (user.id, session["class_name"], str(output), session["id"], now()),
            )
        con.execute("UPDATE roster_sessions SET state='complete',updated_at=? WHERE id=?", (now(), session["id"]))
    balance = credit_balance(user.id)
    if is_free_correction:
        result_text = f"✅ Word bepul yangilandi; {added_count} o'quvchi qo'shildi. Qolgan Word tayyorlashlar: {balance} ta."
    elif correction:
        result_text = f"✅ Word yangilandi; {added_count} o'quvchi qo'shildi. Qolgan Word tayyorlashlar: {balance} ta."
    else:
        result_text = f"✅ Word tayyorlandi. Qolgan Word tayyorlashlar: {balance} ta."
    await update.message.reply_text(
        result_text,
        reply_markup=main_menu(),
    )
    with open(output, "rb") as f:
        await update.message.reply_document(f, filename=f"{session['class_name']}.docx")


async def purchase(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    user = update.effective_user
    if not registered(user.id):
        await start(update, ctx)
        return
    await show_payment_instructions(update, ctx)


async def show_payment_instructions(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await send_payment_instructions_to_user(ctx, update.effective_user.id)


async def send_payment_instructions_to_user(ctx: ContextTypes.DEFAULT_TYPE, user_id: int):
    card = os.environ.get("PAYMENT_CARD_NUMBER", "").strip()
    if not card:
        await ctx.bot.send_message(
            chat_id=user_id,
            text="To'lov kartasi hali sozlanmagan. Administrator bilan bog'laning.",
        )
        return
    ctx.user_data.pop("awaiting_receipt", None)
    ctx.user_data["awaiting_receipt"] = True
    await ctx.bot.send_message(
        chat_id=user_id,
        text=(
            "💳 Word tayyorlash uchun to'lov\n\n"
            f"💰 Narxi: {WORD_PRICE:,} so'm / 1 ta Word\n"
            f"💳 Karta raqami: {card}\n\n"
            "Word soni chekdagi summa asosida hisoblanadi; to'lov summasi 5 000 so'mga karrali bo'lsin.\n\n"
            "To'lovni amalga oshirgach, chek rasmini shu chatga yuboring. "
            "Admin chekdagi summani tasdiqlaydi, Word tayyorlashlar soni avtomatik hisoblanadi."
        ),
    )


async def receipt(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    user = update.effective_user
    if not ctx.user_data.get("awaiting_receipt"):
        return
    image = image_file_id(update.message)
    if not image:
        await update.message.reply_text(
            "📷 Chekni Telegram fotosi yoki .jpg/.jpeg/.png rasm fayli qilib yuboring."
        )
        return
    file = await ctx.bot.get_file(image[0])
    dest = PRIVATE / "receipts" / f"{uuid.uuid4().hex}{image[1]}"
    dest.parent.mkdir(parents=True, exist_ok=True)
    await file.download_to_drive(dest)
    with DB.connect() as con:
        cur = con.execute(
            "INSERT INTO payments(user_id,receipt_path,amount,credits,status,created_at) "
            "VALUES(?,?,?,?,?,?)",
            (user.id, str(dest), None, 0, "pending", now()),
        )
        payment_id = cur.lastrowid
    ctx.user_data.pop("awaiting_receipt", None)
    group_id = os.environ.get("ADMIN_REVIEW_GROUP_ID", "").strip()
    if not group_id:
        await update.message.reply_text(
            "✅ Chek qabul qilindi. To'lov guruhi sozlanmagan, admin tekshiradi."
        )
        return
    nickname = f"@{user.username}" if user.username else user.full_name
    keyboard = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"payapprove:{payment_id}"),
            InlineKeyboardButton("❌ Rad etish", callback_data=f"payreject:{payment_id}"),
        ]
    ])
    with open(dest, "rb") as f:
        await ctx.bot.send_photo(
            chat_id=int(group_id),
            photo=f,
            caption=(
                f"🧾 TO'LOV CHEKI  |  #{payment_id}\n"
                f"👤 Foydalanuvchi: {nickname} (ID: {user.id})\n\n"
                f"💰 1 ta Word: {WORD_PRICE:,} so'm\n"
                "Chekdagi summani tekshiring. Tasdiqlash uchun "
                "chekdagi to'langan summani kiriting."
            ),
            reply_markup=keyboard,
        )
    await update.message.reply_text("✅ Chek adminga yuborildi. Tasdiqni kuting.")


async def review_group_member(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    group_value = os.environ.get("ADMIN_REVIEW_GROUP_ID", "").strip()
    try:
        group_id = int(group_value)
    except ValueError:
        logging.error("ADMIN_REVIEW_GROUP_ID is missing or invalid.")
        await query.answer("Tekshiruv guruhi sozlanmagan.", show_alert=True)
        return False
    if not query.message or query.message.chat.id != group_id:
        await query.answer("Bu tugma faqat to'lov tekshiruv guruhida ishlaydi.", show_alert=True)
        return False
    try:
        member = await ctx.bot.get_chat_member(group_id, update.effective_user.id)
    except TelegramError:
        logging.exception("Could not verify review-group membership.")
        await query.answer(
            "Guruh a'zoligini tekshirib bo'lmadi. Bot guruhda administrator bo'lishi kerak.",
            show_alert=True,
        )
        return False
    if not can_review_member(member):
        await query.answer("Bu amal faqat tekshiruv guruhi a'zolari uchun.", show_alert=True)
        return False
    return True


def can_review_member(member):
    return member.status in ("member", "administrator", "creator") or (
        member.status == "restricted" and member.is_member
    )


async def payment_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    if not await review_group_member(update, ctx):
        return
    actor = update.effective_user.id
    action, payment_id_text = query.data.split(":", 1)
    payment_id = int(payment_id_text)
    if action == "payapprove":
        with DB.connect() as con:
            payment = con.execute(
                "SELECT user_id,status FROM payments WHERE id=?",
                (payment_id,),
            ).fetchone()
        if not payment:
            await query.answer("To'lov topilmadi.", show_alert=True)
            return
        if payment["status"] != "pending":
            await query.answer("Bu to'lov allaqachon ko'rib chiqilgan.", show_alert=True)
            return
        prompt = await query.message.reply_text(
            f"💰 Chek #{payment_id}: shu xabarga javob qilib, chekdagi to'langan summani so'mda yuboring.\n"
            f"1 ta Word = {WORD_PRICE:,} so'm. Masalan: 10000",
            reply_markup=ForceReply(input_field_placeholder="Masalan: 10000"),
        )
        ctx.chat_data.setdefault("payment_amount_prompts", {})[prompt.message_id] = {
            "payment_id": payment_id,
            "review_message_id": query.message.message_id,
            "caption": query.message.caption or "",
        }
        await query.answer("Summani yuborish uchun guruhdagi savolga javob bering.")
        return

    with DB.connect() as con:
        payment = con.execute("SELECT * FROM payments WHERE id=?", (payment_id,)).fetchone()
        if not payment:
            await query.answer("To'lov topilmadi.", show_alert=True)
            return
        status = payment["status"]
        if action == "payreject" and status == "pending":
            con.execute("UPDATE payments SET status='rejected',reviewed_by=?,reviewed_at=? WHERE id=?", (actor, now(), payment_id))
            con.execute(
                "INSERT INTO payment_events(payment_id,actor_id,action,created_at) VALUES(?,?,?,?)",
                (payment_id, actor, "rejected", now()),
            )
            result_text = "Rad etildi"
        elif action == "payreject" and status in ("approved", "rejected"):
            await query.answer("Bu to'lov allaqachon qaror qilingan.", show_alert=True)
            return
        else:
            await query.answer("Noto'g'ri to'lov holati.", show_alert=True)
            return
    await query.answer(result_text)
    try:
        await ctx.bot.send_message(payment["user_id"], "❌ To'lovingiz rad etildi.")
    except TelegramError:
        logging.exception("Payment user notification failed for payment_id=%s", payment_id)
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("↩️ Qarorni 5 daqiqa ichida qaytarish", callback_data=f"payreverse:{payment_id}")]
    ])
    await query.edit_message_caption(
        caption=f"{query.message.caption}\n\n❌ Holat: {result_text}",
        reply_markup=keyboard,
    )


async def payment_amount_input(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    message = update.message
    if not message or not message.reply_to_message:
        return
    prompts = ctx.chat_data.get("payment_amount_prompts", {})
    prompt_id = message.reply_to_message.message_id
    review = prompts.get(prompt_id)
    if not review:
        return
    group_value = os.environ.get("ADMIN_REVIEW_GROUP_ID", "").strip()
    try:
        group_id = int(group_value)
    except ValueError:
        logging.error("ADMIN_REVIEW_GROUP_ID is missing or invalid.")
        await message.reply_text("To'lov guruhi sozlanmagan.")
        return
    if message.chat.id != group_id:
        return
    try:
        member = await ctx.bot.get_chat_member(group_id, update.effective_user.id)
    except TelegramError:
        logging.exception("Could not verify reviewer for payment amount.")
        await message.reply_text("Guruh a'zoligini tekshirib bo'lmadi.")
        return
    if not can_review_member(member):
        await message.reply_text("Bu amal faqat guruh a'zolari uchun.")
        return
    raw_amount = re.sub(r"[\s,]", "", message.text or "")
    if not raw_amount.isdigit():
        await message.reply_text("Summani faqat raqam bilan yuboring. Masalan: 10000")
        return
    if len(raw_amount) > 19:
        await message.reply_text("Summa juda katta. Chekdagi summani qayta tekshiring.")
        return
    amount = int(raw_amount)
    status, word_count, user_id = approve_payment(
        DB,
        review["payment_id"],
        update.effective_user.id,
        amount,
    )
    if status == "invalid_amount":
        await message.reply_text(
            f"❌ Summa kamida {WORD_PRICE:,} so'm va {WORD_PRICE:,} so'mga karrali bo'lishi kerak. "
            "Chekdagi summani qayta yuboring."
        )
        return
    prompts.pop(prompt_id, None)
    if status == "not_found":
        await message.reply_text("❌ To'lov topilmadi.")
        return
    if status == "already_reviewed":
        await message.reply_text("ℹ️ Bu to'lov allaqachon ko'rib chiqilgan.")
        return

    await message.reply_text(
        f"✅ To'lov tasdiqlandi: {amount:,} so'm → {word_count} ta Word."
    )
    try:
        await ctx.bot.send_message(
            user_id,
            f"✅ To'lovingiz tasdiqlandi!\n"
            f"💰 To'langan summa: {amount:,} so'm\n"
            f"📝 Qo'shilgan Word tayyorlashlar: {word_count} ta\n"
            f"Qolgan foydalanishlar: {credit_balance(user_id)} ta.\n"
            "Davom etish uchun /royxatga_olish yoki /qoshimcha buyrug'ini yuboring.",
        )
    except TelegramError:
        logging.exception("Payment user notification failed for payment_id=%s", review["payment_id"])
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton(
            "↩️ Qarorni 5 daqiqa ichida qaytarish",
            callback_data=f"payreverse:{review['payment_id']}",
        )]
    ])
    caption = (
        f"{review['caption']}\n\n"
        f"✅ TASDIQLANDI\n"
        f"💰 Summa: {amount:,} so'm\n"
        f"📝 Word: {word_count} ta"
    )
    try:
        await ctx.bot.edit_message_caption(
            chat_id=group_id,
            message_id=review["review_message_id"],
            caption=caption,
            reply_markup=keyboard,
        )
    except TelegramError:
        logging.exception("Could not update receipt message for payment_id=%s", review["payment_id"])
        await message.reply_text(
            "To'lov tasdiqlandi, ammo chek xabarini yangilashda Telegram xatosi yuz berdi."
        )


async def reverse_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    if not await review_group_member(update, ctx):
        return
    actor = update.effective_user.id
    _, payment_id_text = query.data.split(":", 1)
    payment_id = int(payment_id_text)
    with DB.connect() as con:
        payment = con.execute("SELECT * FROM payments WHERE id=?", (payment_id,)).fetchone()
        if not payment or payment["status"] not in ("approved", "rejected"):
            await query.answer("Qaytarish mumkin emas.", show_alert=True)
            return
        if not payment["reviewed_at"]:
            await query.answer("Qaror vaqti topilmadi.", show_alert=True)
            return
        reviewed = datetime.fromisoformat(payment["reviewed_at"])
        if datetime.now(timezone.utc) - reviewed > timedelta(minutes=5):
            await query.answer("5 daqiqalik qaytarish muddati tugagan.", show_alert=True)
            return
        existing = con.execute(
            "SELECT 1 FROM payment_events WHERE payment_id=? AND action='reversed'",
            (payment_id,),
        ).fetchone()
        if existing:
            await query.answer("Bu qaror allaqachon qaytarilgan.", show_alert=True)
            return
        previous = payment["status"]
        con.execute("UPDATE payments SET status='pending',reviewed_by=NULL,reviewed_at=NULL WHERE id=?", (payment_id,))
        if previous == "approved":
            con.execute("UPDATE credits SET balance=MAX(balance-?,0) WHERE user_id=?", (payment["credits"], payment["user_id"]))
        con.execute(
            "INSERT INTO payment_events(payment_id,actor_id,action,created_at) VALUES(?,?,?,?)",
            (payment_id, actor, "reversed", now()),
        )
    await query.answer("Qaror qaytarildi.")
    await query.edit_message_caption(caption=f"{query.message.caption}\n\nHolat: Qaytarildi")


async def my_files(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not await require_private_chat(update):
        return
    user = update.effective_user
    if not registered(user.id):
        await start(update, ctx)
        return
    await show_my_files(update, user.id, 0)


async def show_my_files(update_or_query, user_id: int, page: int):
    page_size = 10
    with DB.connect() as con:
        total = con.execute(
            "SELECT COUNT(*) FROM documents WHERE user_id=?",
            (user_id,),
        ).fetchone()[0]
        docs = con.execute(
            "SELECT id,class_name,created_at FROM documents "
            "WHERE user_id=? ORDER BY created_at DESC,id DESC LIMIT ? OFFSET ?",
            (user_id, page_size, page * page_size),
        ).fetchall()
    if not docs:
        text = "Sizga tegishli tayyor Word fayl yo'q."
        if update_or_query.callback_query:
            await update_or_query.callback_query.edit_message_text(text)
        else:
            await update_or_query.message.reply_text(text)
        return
    buttons = [
        [InlineKeyboardButton(
            f"{doc['class_name']}.docx",
            callback_data=f"myfiles:file:{doc['id']}",
        )]
        for doc in docs
    ]
    navigation = []
    if page > 0:
        navigation.append(
            InlineKeyboardButton("⬅ Oldingi", callback_data=f"myfiles:page:{page - 1}")
        )
    if (page + 1) * page_size < total:
        navigation.append(
            InlineKeyboardButton("Keyingi ➡", callback_data=f"myfiles:page:{page + 1}")
        )
    if navigation:
        buttons.append(navigation)
    markup = InlineKeyboardMarkup(buttons)
    text = f"📁 Fayllarim ({total} ta) · {page + 1}-bo'lim"
    if update_or_query.callback_query:
        await update_or_query.callback_query.edit_message_text(text, reply_markup=markup)
    else:
        await update_or_query.message.reply_text(text, reply_markup=markup)


async def my_files_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    if not await require_private_chat(update):
        return
    if not registered(update.effective_user.id):
        await query.answer("Avval ro'yxatdan o'ting.", show_alert=True)
        return
    parts = query.data.split(":")
    if len(parts) != 3:
        await query.answer("So'rov noto'g'ri.", show_alert=True)
        return
    if parts[1] == "page":
        try:
            page = int(parts[2])
        except ValueError:
            await query.answer("Bo'lim raqami noto'g'ri.", show_alert=True)
            return
        if page < 0:
            await query.answer("Bo'lim raqami noto'g'ri.", show_alert=True)
            return
        await query.answer()
        await show_my_files(update, update.effective_user.id, page)
        return
    if parts[1] != "file":
        await query.answer("So'rov noto'g'ri.", show_alert=True)
        return
    try:
        document_id = int(parts[2])
    except ValueError:
        await query.answer("Fayl raqami noto'g'ri.", show_alert=True)
        return
    with DB.connect() as con:
        doc = con.execute(
            "SELECT class_name,file_path FROM documents WHERE id=? AND user_id=?",
            (document_id, update.effective_user.id),
        ).fetchone()
    if not doc:
        await query.answer("Fayl topilmadi yoki sizga tegishli emas.", show_alert=True)
        return
    path = Path(doc["file_path"])
    if not path.is_file():
        await query.answer("Fayl saqlangan joyidan topilmadi.", show_alert=True)
        return
    await query.answer()
    with open(path, "rb") as file:
        await ctx.bot.send_document(
            chat_id=update.effective_chat.id,
            document=file,
            filename=f"{doc['class_name']}.docx",
        )


async def admin(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in id_set("ADMIN_IDS"):
        await update.message.reply_text("Bu buyruq faqat asosiy administrator uchun.")
        return
    if update.effective_chat.type != "private":
        await update.message.reply_text("Admin panelni bot bilan shaxsiy chatda oching.")
        return
    await update.message.reply_text("Admin panel", reply_markup=admin_menu_markup())


async def admin_panel_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    if update.effective_user.id not in id_set("ADMIN_IDS"):
        await query.answer("Bu panel faqat asosiy administrator uchun.", show_alert=True)
        return
    if not query.message or query.message.chat.type != "private":
        await query.answer("Admin panelni shaxsiy chatda ishlating.", show_alert=True)
        return
    parts = query.data.split(":")
    action = parts[1]
    if action != "template":
        await query.answer()
    if action == "users":
        page = int(parts[2]) if len(parts) > 2 else 0
        await show_admin_users(query, page)
    elif action == "user":
        await query.edit_message_text("Foydalanuvchi ma'lumotlari yuklanmoqda...")
        await send_admin_user_data(ctx.bot, query.message.chat.id, int(parts[2]))
    elif action == "grant":
        user_id = int(parts[2])
        prompt = await query.message.reply_text(
            f"Foydalanuvchi {user_id} ga nechta Word tayyorlash berilsin? "
            "Musbat butun sonni javob qilib yuboring.",
            reply_markup=ForceReply(input_field_placeholder="Masalan: 3"),
        )
        ctx.chat_data.setdefault("admin_grant_prompts", {})[prompt.message_id] = user_id
    elif action == "clear":
        user_id = int(parts[2])
        await query.edit_message_text(
            f"⚠️ Foydalanuvchi {user_id} ning profil ma'lumotlari, rasmlari, matnlari, "
            "shablonlari va tayyor Word fayllari tozalanadi.\n"
            "To'lov, promokod va admin tomonidan berilgan Wordlar auditi saqlanadi.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🧹 Ha, tozalash", callback_data=f"admin:clearconfirm:{user_id}")],
                [InlineKeyboardButton("Bekor qilish", callback_data=f"admin:clearcancel:{user_id}")],
            ]),
        )
    elif action == "clearconfirm":
        user_id = int(parts[2])
        status, deleted_files, failed_files = clear_user_data(DB, user_id, USER_FILE_ROOTS)
        if status == "not_found":
            await query.edit_message_text("Foydalanuvchi topilmadi.")
            return
        message = (
            f"✅ Foydalanuvchi {user_id} ning shaxsiy va ishchi ma'lumotlari tozalandi.\n"
            "To'lov, promokod va admin berish auditi saqlandi.\n"
            f"O'chirilgan fayllar: {deleted_files} ta."
        )
        if failed_files:
            message += f"\n⚠️ {failed_files} ta faylni o'chirib bo'lmadi; logni tekshiring."
        await query.edit_message_text(
            message,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("Saqlangan auditni ko'rish", callback_data=f"admin:user:{user_id}")],
                [InlineKeyboardButton("⬅ Foydalanuvchilar", callback_data="admin:users:0")],
            ]),
        )
    elif action == "clearcancel":
        user_id = int(parts[2])
        await query.edit_message_text(
            "Tozalash bekor qilindi.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅ Foydalanuvchi statistikasi", callback_data=f"admin:user:{user_id}")],
            ]),
        )
    elif action == "export":
        format_name, user_id = parts[2], int(parts[3])
        if format_name not in ("docx", "pdf"):
            await query.edit_message_text("Eksport formati noto'g'ri.")
            return
        await send_admin_user_export(
            ctx.bot,
            query.message.chat.id,
            user_id,
            format_name,
        )
    elif action == "stats":
        await query.edit_message_text(
            await admin_statistics(),
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅ Admin panel", callback_data="admin:menu")]
            ]),
        )
    elif action == "payments":
        await query.edit_message_text(
            await admin_payment_summary(),
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅ Admin panel", callback_data="admin:menu")]
            ]),
        )
    elif action == "promo":
        await query.edit_message_text(
            "Promokod yaratish:\n"
            "/promo_yarat KOD WORD_SONI FOYDALANISH_LIMI\n"
            "Masalan: /promo_yarat MAKTAB10 2 50\n"
            "Har foydalanuvchi bir kodni bir marta ishlatishi mumkin.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅ Admin panel", callback_data="admin:menu")]
            ]),
        )
    elif action == "template":
        enabled = shared_template_enabled()
        if not enabled:
            if not SHABLON.is_file():
                await query.answer("Umumiy shablon topilmadi.", show_alert=True)
                return
            DB.set_setting("shared_template_enabled", "1")
            await query.answer("Umumiy shablon yoqildi.")
        else:
            DB.set_setting("shared_template_enabled", "0")
            await query.answer("Umumiy shablon o'chirildi.")
        await query.edit_message_text("Admin panel", reply_markup=admin_menu_markup())
    elif action == "menu":
        await query.edit_message_text("Admin panel", reply_markup=admin_menu_markup())


async def show_admin_users(query, page):
    page_size = 8
    with DB.connect() as con:
        total = con.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]
        rows = con.execute(
            "SELECT user_id,username,phone FROM users ORDER BY registered_at DESC LIMIT ? OFFSET ?",
            (page_size, page * page_size),
        ).fetchall()
    if not total:
        await query.edit_message_text(
            "Foydalanuvchilar hali yo'q.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅ Admin panel", callback_data="admin:menu")]
            ]),
        )
        return
    buttons = [
        [InlineKeyboardButton(
            f"{row['username'] or row['phone'] or 'Foydalanuvchi'} · {row['user_id']}",
            callback_data=f"admin:user:{row['user_id']}",
        )]
        for row in rows
    ]
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("⬅", callback_data=f"admin:users:{page - 1}"))
    if (page + 1) * page_size < total:
        nav.append(InlineKeyboardButton("➡", callback_data=f"admin:users:{page + 1}"))
    if nav:
        buttons.append(nav)
    buttons.append([InlineKeyboardButton("⬅ Admin panel", callback_data="admin:menu")])
    await query.edit_message_text(
        f"Foydalanuvchilar: {total} ta (sahifa {page + 1})",
        reply_markup=InlineKeyboardMarkup(buttons),
    )


async def admin_statistics():
    with DB.connect() as con:
        users_count = con.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        registered_count = con.execute(
            "SELECT COUNT(*) FROM users WHERE phone IS NOT NULL"
        ).fetchone()[0]
        student_count = con.execute("SELECT COUNT(*) FROM students").fetchone()[0]
        document_count = con.execute("SELECT COUNT(*) FROM documents").fetchone()[0]
        payments = con.execute(
            "SELECT status,COUNT(*) AS n FROM payments GROUP BY status"
        ).fetchall()
        credits = con.execute("SELECT COALESCE(SUM(balance),0) FROM credits").fetchone()[0]
        promo_count = con.execute("SELECT COUNT(*) FROM promo_codes").fetchone()[0]
        promo_uses = con.execute("SELECT COALESCE(SUM(uses),0) FROM promo_codes").fetchone()[0]
    payment_counts = {row["status"]: row["n"] for row in payments}
    return (
        "Bot statistikasi\n"
        f"Foydalanuvchilar: {users_count}\n"
        f"Ro'yxatdan o'tganlar: {registered_count}\n"
        f"O'quvchi yozuvlari: {student_count}\n"
        f"Tayyorlangan Word fayllar: {document_count}\n"
        f"Foydalanuvchilardagi qolgan Word tayyorlashlar: {credits}\n"
        f"Promokodlar: {promo_count} ta, foydalanilgan {promo_uses} marta\n"
        f"To'lovlar: kutilmoqda {payment_counts.get('pending', 0)}, "
        f"tasdiqlangan {payment_counts.get('approved', 0)}, "
        f"rad etilgan {payment_counts.get('rejected', 0)}"
    )


async def admin_payment_summary():
    with DB.connect() as con:
        rows = con.execute(
            "SELECT id,user_id,amount,status,created_at FROM payments ORDER BY id DESC LIMIT 20"
        ).fetchall()
    if not rows:
        return "To'lovlar hali yo'q."
    lines = ["So'nggi to'lovlar:"]
    for row in rows:
        amount = f"{row['amount']:,} so'm" if row["amount"] is not None else "summa kutilmoqda"
        lines.append(f"#{row['id']} | user {row['user_id']} | {amount} | {row['status']}")
    return "\n".join(lines)


async def send_admin_user_data(bot, chat_id, user_id):
    data = load_admin_user_data(DB, user_id)
    if not data:
        await bot.send_message(chat_id, "Foydalanuvchi topilmadi.")
        return
    user = data["user"]
    await bot.send_message(
        chat_id,
        "👤 Foydalanuvchi statistikasi\n\n"
        f"🆔 ID: {user_id}\n"
        f"Username: @{user['username'] or '-'}\n"
        f"📱 Telefon: {user['phone'] or '-'}\n"
        f"📝 Qolgan Word tayyorlashlar: {data['remaining_uses']} ta\n"
        f"📚 Sinf sessiyalari: {len(data['sessions'])}\n"
        f"👨‍🎓 O'quvchi yozuvlari: {len(data['students'])}\n"
        f"📨 Yuborilgan materiallar: {len(data['submissions'])}\n"
        f"📄 Tayyor Word fayllar: {len(data['documents'])}\n"
        f"💳 To'lovlar: {len(data['payments'])}\n"
        f"🎟 Ishlatilgan promokodlar: {len(data['promos'])}\n"
        f"🎁 Admin bergan Wordlar: {sum(item['uses'] for item in data['grants'])} ta",
        reply_markup=admin_user_markup(user_id),
    )


async def send_admin_user_export(bot, chat_id, user_id, format_name):
    data = load_admin_user_data(DB, user_id)
    if not data:
        await bot.send_message(chat_id, "Foydalanuvchi topilmadi.")
        return
    with tempfile.TemporaryDirectory(prefix=f"admin-export-{user_id}-", dir=PRIVATE) as temp:
        root = Path(temp)
        report_path = root / f"foydalanuvchi_{user_id}.{format_name}"
        create_user_report(data, format_name, report_path, USER_FILE_ROOTS, SHABLON)
        archive_path = root / f"foydalanuvchi_{user_id}_asl_fayllar.zip"
        has_archive = create_user_archive(data, archive_path, USER_FILE_ROOTS, SHABLON)
        with open(report_path, "rb") as report:
            await bot.send_document(
                chat_id,
                document=report,
                filename=report_path.name,
                caption=f"📊 Foydalanuvchi {user_id} hisoboti ({format_name.upper()})",
            )
        if has_archive:
            with open(archive_path, "rb") as archive:
                await bot.send_document(
                    chat_id,
                    document=archive,
                    filename=archive_path.name,
                    caption="📦 Asl rasmlar, cheklar, shablonlar va Word fayllar",
                )


async def admin_grant_reply(update: Update, ctx: ContextTypes.DEFAULT_TYPE, user_id, prompt_id):
    text_value = (update.message.text or "").strip()
    if not text_value.isdigit():
        await update.message.reply_text("Word sonini musbat butun son qilib yuboring.")
        return
    if len(text_value) > 19:
        await update.message.reply_text("Word soni ruxsat etilgan maksimal qiymatdan oshib ketdi.")
        return
    amount = int(text_value)
    status, remaining = grant_word_uses(
        DB,
        user_id,
        update.effective_user.id,
        amount,
    )
    if status == "invalid_amount":
        await update.message.reply_text("Word soni 1 dan katta va juda katta bo'lmagan butun son bo'lsin.")
        return
    if status == "too_large":
        await update.message.reply_text("Word soni ruxsat etilgan maksimal qiymatdan oshib ketdi.")
        return
    ctx.chat_data.get("admin_grant_prompts", {}).pop(prompt_id, None)
    if status == "not_found":
        await update.message.reply_text("Foydalanuvchi topilmadi.")
        return
    await update.message.reply_text(
        f"✅ Foydalanuvchi {user_id} ga {amount} ta Word tayyorlash berildi.\n"
        f"Qolgan Word tayyorlashlar: {remaining} ta.",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("⬅ Statistika", callback_data=f"admin:user:{user_id}")],
        ]),
    )
    try:
        await ctx.bot.send_message(
            user_id,
            f"🎁 Administrator sizga {amount} ta Word tayyorlash berdi.\n"
            f"Qolgan Word tayyorlashlar: {remaining} ta.",
        )
    except TelegramError:
        logging.exception("Could not notify user about admin Word grant: user_id=%s", user_id)
        await update.message.reply_text(
            "⚠️ Berish saqlandi, lekin foydalanuvchiga xabar yetkazilmadi."
        )


async def users(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in id_set("ADMIN_IDS"):
        await update.message.reply_text("Administrator huquqi yo'q.")
        return
    if update.effective_chat.type != "private":
        await update.message.reply_text("Foydalanuvchi ma'lumotlarini shaxsiy chatda ko'ring.")
        return
    await admin(update, ctx)


async def user_detail(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in id_set("ADMIN_IDS"):
        await update.message.reply_text("Administrator huquqi yo'q.")
        return
    if update.effective_chat.type != "private":
        await update.message.reply_text("Foydalanuvchi ma'lumotlarini shaxsiy chatda ko'ring.")
        return
    if not ctx.args:
        await update.message.reply_text("Foydalanuvchi ID sini yozing: /user 123456")
        return
    try:
        user_id = int(ctx.args[0])
    except ValueError:
        await update.message.reply_text("ID raqam bo'lishi kerak.")
        return
    await send_admin_user_data(ctx.bot, update.effective_chat.id, user_id)


async def payments(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in id_set("ADMIN_IDS"):
        await update.message.reply_text("Administrator huquqi yo'q.")
        return
    if update.effective_chat.type != "private":
        await update.message.reply_text("To'lovlar haqidagi ma'lumotni shaxsiy chatda ko'ring.")
        return
    with DB.connect() as con:
        rows = con.execute(
            "SELECT id,user_id,amount,status,created_at FROM payments ORDER BY id DESC LIMIT 30"
        ).fetchall()
    if not rows:
        await update.message.reply_text("To'lovlar yo'q.")
        return
    lines = []
    for row in rows:
        amount = f"{row['amount']:,} so'm" if row["amount"] is not None else "summa kutilmoqda"
        lines.append(f"{row['id']} | user {row['user_id']} | {amount} | {row['status']}")
    await update.message.reply_text("\n".join(lines))


def main():
    token = os.environ.get("BOT_TOKEN", "").strip()
    if not token:
        raise RuntimeError(
            "BOT_TOKEN topilmadi. Loyiha papkasidagi .env.example faylidan .env nusxa yarating "
            "va yangi BOT_TOKEN ni .env ichiga kiriting. .env.example avtomatik yuklanmaydi."
        )
    migrate_legacy_json()
    admin_ids = id_set("ADMIN_IDS")
    if not admin_ids:
        logging.warning("ADMIN_IDS is empty; admin commands are disabled.")
    if not os.environ.get("ADMIN_REVIEW_GROUP_ID"):
        logging.warning("ADMIN_REVIEW_GROUP_ID is empty; receipt forwarding is disabled.")

    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler(["start", "yordam"], start))
    app.add_handler(CommandHandler("royxatga_olish", roster_start))
    app.add_handler(CommandHandler("bekor", cancel))
    app.add_handler(CommandHandler(["yakunlash", "tayyor"], finish))
    app.add_handler(CommandHandler(["sotib_olish", "purchase"], purchase))
    app.add_handler(CommandHandler(["fayllarim", "files"], my_files))
    app.add_handler(CommandHandler("profil", profile))
    app.add_handler(CommandHandler("promo", promo))
    app.add_handler(CommandHandler("promo_yarat", create_promo))
    app.add_handler(CommandHandler("qoshimcha", qoshimcha))
    app.add_handler(CommandHandler(["admin", "admin_jrgrumua"], admin))
    app.add_handler(CommandHandler(["users", "foydalanuvchilar"], users))
    app.add_handler(CommandHandler("user", user_detail))
    app.add_handler(CommandHandler("payments", payments))
    app.add_handler(MessageHandler(filters.CONTACT, contact))
    app.add_handler(MessageHandler(filters.Document.DOCX, document))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.IMAGE, _dispatch_photo))
    review_group = os.environ.get("ADMIN_REVIEW_GROUP_ID", "").strip()
    try:
        review_group_id = int(review_group)
    except ValueError:
        review_group_id = None
    if review_group_id is not None:
        app.add_handler(
            MessageHandler(
                filters.Chat(chat_id=review_group_id)
                & filters.TEXT
                & filters.REPLY
                & ~filters.COMMAND,
                payment_amount_input,
            )
        )
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text))
    app.add_handler(CallbackQueryHandler(payment_callback, pattern=r"^pay(approve|reject):"))
    app.add_handler(CallbackQueryHandler(reverse_callback, pattern=r"^payreverse:"))
    app.add_handler(CallbackQueryHandler(my_files_callback, pattern=r"^myfiles:"))
    app.add_handler(CallbackQueryHandler(add_student_callback, pattern=r"^addstudent:"))
    app.add_handler(CallbackQueryHandler(admin_panel_callback, pattern=r"^admin:"))
    logging.basicConfig(level=logging.INFO)
    app.run_polling()


async def _dispatch_photo(update, ctx):
    if ctx.user_data.get("awaiting_receipt"):
        await receipt(update, ctx)
    else:
        await photo(update, ctx)


if __name__ == "__main__":
    main()
