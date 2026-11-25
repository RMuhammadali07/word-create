import logging
from datetime import datetime, timezone
from pathlib import Path


MAX_SQLITE_INT = 2**63 - 1


def _now():
    return datetime.now(timezone.utc).isoformat()


def grant_word_uses(db, user_id, admin_id, amount):
    if not isinstance(amount, int) or amount < 1 or amount > MAX_SQLITE_INT:
        return "invalid_amount", 0
    with db.connect() as con:
        con.execute("BEGIN IMMEDIATE")
        if not con.execute("SELECT 1 FROM users WHERE user_id=?", (user_id,)).fetchone():
            return "not_found", 0
        current = con.execute(
            "SELECT balance FROM credits WHERE user_id=?",
            (user_id,),
        ).fetchone()
        current_uses = current["balance"] if current else 0
        if amount > MAX_SQLITE_INT - current_uses:
            return "too_large", current_uses
        con.execute(
            "INSERT INTO credits(user_id,balance) VALUES(?,?) "
            "ON CONFLICT(user_id) DO UPDATE SET balance=balance+excluded.balance",
            (user_id, amount),
        )
        con.execute(
            "INSERT INTO admin_grants(user_id,admin_id,uses,created_at) VALUES(?,?,?,?)",
            (user_id, admin_id, amount, _now()),
        )
        return "ok", current_uses + amount


def load_admin_user_data(db, user_id):
    with db.connect() as con:
        user = con.execute("SELECT * FROM users WHERE user_id=?", (user_id,)).fetchone()
        if not user:
            return None
        sessions = con.execute(
            "SELECT id,class_name,state,template_path,updated_at "
            "FROM roster_sessions WHERE user_id=? ORDER BY id",
            (user_id,),
        ).fetchall()
        students = con.execute(
            "SELECT s.*,r.class_name FROM students s "
            "JOIN roster_sessions r ON r.id=s.session_id "
            "WHERE r.user_id=? ORDER BY s.session_id,s.position",
            (user_id,),
        ).fetchall()
        submissions = con.execute(
            "SELECT kind,text,file_path,created_at FROM submissions "
            "WHERE user_id=? ORDER BY id",
            (user_id,),
        ).fetchall()
        documents = con.execute(
            "SELECT class_name,file_path,created_at FROM documents "
            "WHERE user_id=? ORDER BY id",
            (user_id,),
        ).fetchall()
        payments = con.execute(
            "SELECT id,amount,credits,status,receipt_path,created_at "
            "FROM payments WHERE user_id=? ORDER BY id",
            (user_id,),
        ).fetchall()
        payment_events = con.execute(
            "SELECT e.payment_id,e.actor_id,e.action,e.created_at "
            "FROM payment_events e JOIN payments p ON p.id=e.payment_id "
            "WHERE p.user_id=? ORDER BY e.id",
            (user_id,),
        ).fetchall()
        grants = con.execute(
            "SELECT admin_id,uses,created_at FROM admin_grants "
            "WHERE user_id=? ORDER BY id",
            (user_id,),
        ).fetchall()
        promos = con.execute(
            "SELECT code,redeemed_at FROM promo_redemptions "
            "WHERE user_id=? ORDER BY redeemed_at",
            (user_id,),
        ).fetchall()
        credit = con.execute(
            "SELECT balance FROM credits WHERE user_id=?",
            (user_id,),
        ).fetchone()
    return {
        "user": dict(user),
        "remaining_uses": credit["balance"] if credit else 0,
        "sessions": [dict(row) for row in sessions],
        "students": [dict(row) for row in students],
        "submissions": [dict(row) for row in submissions],
        "documents": [dict(row) for row in documents],
        "payments": [dict(row) for row in payments],
        "payment_events": [dict(row) for row in payment_events],
        "grants": [dict(row) for row in grants],
        "promos": [dict(row) for row in promos],
    }


def clear_user_data(db, user_id, private_root):
    roots = (
        [Path(item).resolve() for item in private_root]
        if isinstance(private_root, (list, tuple, set))
        else [Path(private_root).resolve()]
    )
    with db.connect() as con:
        con.execute("BEGIN IMMEDIATE")
        if not con.execute("SELECT 1 FROM users WHERE user_id=?", (user_id,)).fetchone():
            return "not_found", 0, 0

        paths = set()
        queries = (
            "SELECT photo_path AS path FROM students WHERE session_id IN "
            "(SELECT id FROM roster_sessions WHERE user_id=?)",
            "SELECT file_path AS path FROM submissions WHERE user_id=?",
            "SELECT template_path AS path FROM roster_sessions WHERE user_id=?",
            "SELECT file_path AS path FROM documents WHERE user_id=?",
            "SELECT receipt_path AS path FROM payments WHERE user_id=?",
        )
        for sql in queries:
            paths.update(
                row["path"]
                for row in con.execute(sql, (user_id,))
                if row["path"]
            )

        con.execute(
            "DELETE FROM students WHERE session_id IN "
            "(SELECT id FROM roster_sessions WHERE user_id=?)",
            (user_id,),
        )
        con.execute("DELETE FROM roster_sessions WHERE user_id=?", (user_id,))
        con.execute("DELETE FROM submissions WHERE user_id=?", (user_id,))
        con.execute("DELETE FROM documents WHERE user_id=?", (user_id,))
        con.execute(
            "UPDATE users SET phone=NULL,username='',free_additions_used=0 WHERE user_id=?",
            (user_id,),
        )
        con.execute("UPDATE credits SET balance=0 WHERE user_id=?", (user_id,))

        other_paths = set()
        other_queries = (
            "SELECT file_path AS path FROM submissions WHERE user_id!=?",
            "SELECT template_path AS path FROM roster_sessions WHERE user_id!=?",
            "SELECT file_path AS path FROM documents WHERE user_id!=?",
            "SELECT receipt_path AS path FROM payments WHERE user_id!=?",
            "SELECT s.photo_path AS path FROM students s "
            "JOIN roster_sessions r ON r.id=s.session_id WHERE r.user_id!=?",
        )
        for sql in other_queries:
            other_paths.update(
                row["path"]
                for row in con.execute(sql, (user_id,))
                if row["path"]
            )

    shared = set()
    for raw_path in other_paths:
        try:
            shared.add(str(Path(raw_path).resolve()))
        except OSError:
            logging.exception("Could not resolve another user's file reference.")

    deleted = 0
    failed = 0
    for raw_path in paths:
        try:
            path = Path(raw_path).resolve()
            if (
                not any(root in path.parents for root in roots)
                or str(path) in shared
                or not path.is_file()
            ):
                continue
            path.unlink()
            deleted += 1
        except OSError:
            logging.exception("Could not remove private data file.")
            failed += 1
    return "ok", deleted, failed
