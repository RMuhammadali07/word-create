#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Telegram guruh eksportidan o'quvchilarni (I.F.Sh + rasm) Word shablonga joylaydi.
Har bir sinf uchun alohida .docx fayl yaratiladi (7A.docx, 7B.docx, 3YE.docx ...).

Ishlatish:
    pip install python-docx pillow
    python sinf_tayyorlash.py ChatExport_1 ChatExport_2=7B 7A.docx
    (ChatExport_2=7B: bu chatda sinf yozilmagan, hammasi 7B deb olinadi)
    (papka va shablon ko'rsatilmasa, dastur o'zi so'raydi)

Qo'shimcha:
    --chiqish natija     natija papkasi (standart: natija)
    --jins               F.I.Sh dagi "qizi/ovna" yoki "o'g'li/ovich" bo'yicha Jinsi ustunini to'ldiradi
    --faqat 7A 7B 3YE    faqat shu sinflarni tayyorlaydi
"""
import argparse
import copy
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image

# ---- sozlamalar -------------------------------------------------------
RASM_BALANDLIK = Inches(1.1979)   # shablondagi rasm balandligi
RASM_MAX_ENI = Inches(1.05)       # katak eniga sig'ishi uchun
COL_TR, COL_FISH, COL_SINF, COL_HARF, COL_TEL, COL_JOY, COL_RASM, COL_JINS = range(8)

# Kirill harflarini lotinga o'tkazish (sinf harfi uchun: "7А" kirill A bo'lishi mumkin)
KIRILL_LOTIN = {"А": "A", "В": "V", "Б": "B", "Г": "G", "Д": "D", "Е": "E",
                "а": "A", "в": "V", "б": "B", "г": "G", "д": "D", "е": "E"}

SINF_RE = re.compile(r"^\s*(\d{1,2})\s*[-–—]?\s*([A-Za-zА-Яа-я]{1,2})\s*(?:[- ]?sinf)?\s*[:.\-–—]?\s*(.*)$",
                     re.S | re.I)
SINF_NOM_RE = re.compile(r"^\s*(\d{1,2})\s*[-–]?\s*([A-Za-zА-Яа-я]{1,2})\s*$", re.I)


def harfni_normal(harf):
    return "".join(KIRILL_LOTIN.get(h, h.upper()) for h in harf)


# ---- Telegram eksportini o'qish --------------------------------------
def matn_ol(msg):
    t = msg.get("text", "")
    if isinstance(t, list):
        t = "".join(p if isinstance(p, str) else p.get("text", "") for p in t)
    return t.strip()


def rasm_yoli(msg, papka):
    p = msg.get("photo")
    if not p and msg.get("media_type") not in (None, "photo") and \
            not str(msg.get("mime_type", "")).startswith("image/"):
        return None
    p = p or (msg.get("file") if str(msg.get("mime_type", "")).startswith("image/") else None)
    if not p or p.startswith("("):      # "(File not included...)"
        return None
    f = papka / p
    return f if f.exists() else None


def sinf_ajrat(matn):
    """'7A' yoki '3YE Aliyev Vali' -> (('7','A'), 'Aliyev Vali'); aks holda (None, matn)."""
    m = SINF_RE.match(matn)
    if not m:
        return None, matn
    raqam, harf, qolgan = m.group(1), m.group(2), m.group(3).strip()
    harf = harfni_normal(harf)
    nxt = matn[m.end(2):m.end(2) + 1]
    if qolgan and nxt and nxt.isalpha():          # "7Aliyev" -> sinf emas
        return None, matn
    return (raqam, harf), qolgan


def eksportni_oqi(papka, boshlangich=None, sinflar=None):
    data = json.loads((papka / "result.json").read_text(encoding="utf-8"))
    if "messages" in data:                         # bitta chat eksporti
        chatlar = [data["messages"]]
    else:                                          # butun akkaunt eksporti (ko'p chat)
        royxat = data.get("chats", {}).get("list", [])
        chatlar = [c.get("messages", []) for c in royxat]
    if sinflar is None:
        sinflar = {}             # ("7","A") -> [ {"ism":..., "rasm": Path|None}, ... ]
    for xabarlar in chatlar:
        chatni_oqi(xabarlar, papka, boshlangich, sinflar)
    return sinflar


def chatni_oqi(xabarlar, papka, boshlangich, sinflar):
    joriy = boshlangich
    kutilayotgan = None          # ismi hali topilmagan rasm yozuvi

    def yoz(ism, rasm):
        sinflar.setdefault(joriy, []).append({"ism": ism, "rasm": rasm})
        return sinflar[joriy][-1]

    for m in xabarlar:
        if m.get("type") != "message":
            continue
        matn = matn_ol(m)
        rasm = rasm_yoli(m, papka)
        sinf, qolgan = sinf_ajrat(matn) if matn else (None, "")

        if sinf and not rasm:                      # faqat sinf sarlavhasi: "7A"
            joriy, kutilayotgan = sinf, None
            continue

        if rasm:
            if sinf:                               # sarlavha rasm tagsuzi sifatida
                joriy = sinf
                matn = qolgan
            if joriy is None:
                continue
            kutilayotgan = yoz(" ".join(matn.split()), rasm)
            if matn:
                kutilayotgan = None                # ism bor - tamom
            continue

        # oddiy matn: rasmning ismi bo'lishi mumkin
        if matn and kutilayotgan is not None and not kutilayotgan["ism"]:
            kutilayotgan["ism"] = " ".join(matn.split())
            kutilayotgan = None


# ---- Word shablonni to'ldirish ---------------------------------------
def jins_ol(ism):
    s = ism.lower().replace("’", "'").replace("`", "'").replace("ʻ", "'")
    if re.search(r"(qizi|kizi|ovna|evna|yevna)\b", s):
        return "Qiz"
    if re.search(r"(o'g'li|og'li|ogli|ugli|o'gli|ovich|evich|yevich)\b", s):
        return "O'g'il"
    return ""


def katakni_tozala(tc):
    for p in tc.findall(qn("w:p"))[1:]:
        tc.remove(p)
    p = tc.find(qn("w:p"))
    for ch in list(p):
        p.remove(ch)
    return p


def matn_yoz(cell, matn, markaz=False):
    tc = cell._tc
    katakni_tozala(tc)
    par = cell.paragraphs[0]
    if markaz:
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = par.add_run(matn)
    run.font.name = "Calibri"
    run.font.size = Pt(11)
    rpr = run._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), "Calibri")


def rasm_qo_y(cell, yol):
    tc = cell._tc
    # eski rasm aloqalarini yig'ib, keyin o'chiramiz
    katakni_tozala(tc)
    par = cell.paragraphs[0]
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    with Image.open(yol) as im:
        w, h = im.size
    balandlik = RASM_BALANDLIK
    eni = int(balandlik * w / h)
    if eni > RASM_MAX_ENI:
        eni = int(RASM_MAX_ENI)
        balandlik = int(eni * h / w)
    par.add_run().add_picture(str(yol), width=Emu(eni), height=Emu(balandlik))


def sarlavha_normal(sarlavha):
    s = str(sarlavha).lower()
    for apostrof in ("'", "’", "ʻ", "`", "�"):
        s = s.replace(apostrof, " ")
    return " ".join(re.sub(r"[^a-z0-9а-яё]+", " ", s).split())


def sarlavha_turi(sarlavha):
    s = sarlavha_normal(sarlavha)
    if s in ("t r", "tartib raqami", "raqam", "no", "№"):
        return "tr"
    if any(word in s for word in ("rasm", "foto", "photo", "image")):
        return "photo"
    if "harf" in s or "letter" in s:
        return "class_letter"
    if "sinf" in s or "class" in s:
        return "class_number"
    if any(word in s.split() for word in ("telefon", "tel", "phone", "mobile")):
        return "phone"
    if (
        any(word in s.split() for word in ("ism", "fio", "name", "familiya"))
        or "f i sh" in s
        or s in ("fish", "full name")
    ):
        return "name"
    if "yosh" in s or "age" in s:
        return "age"
    if "jins" in s or "gender" in s:
        return "gender"
    if "manzil" in s or "joy" in s or "address" in s:
        return "address"
    return "extra"


def sinfni_yoz(shablon, sinf, oquvchilar, chiqish, jins_kerak):
    doc = Document(str(shablon))
    if not doc.tables or not doc.tables[0].rows:
        raise RuntimeError(f"{Path(shablon).name} ichida jadval topilmadi.")
    jadval = doc.tables[0]
    satrlar = jadval.rows
    if len(satrlar) < 2:
        raise RuntimeError(f"{Path(shablon).name} jadvalida sarlavha ostida namuna qatori bo'lishi kerak.")
    sarlavhalar = [cell.text for cell in satrlar[0].cells]
    namuna_tr = copy.deepcopy(satrlar[1]._tr)     # 1-o'quvchi qatori namuna sifatida

    # kerak bo'lsa qatorlarni ko'paytiramiz, ortiqchasini o'chiramiz
    bor = len(satrlar) - 1
    kerak = len(oquvchilar)
    for _ in range(kerak - bor):
        jadval._tbl.append(copy.deepcopy(namuna_tr))
    for r in list(jadval.rows)[1 + kerak:]:
        jadval._tbl.remove(r._tr)

    raqam, harf = sinf
    for i, (o, row) in enumerate(zip(oquvchilar, list(jadval.rows)[1:]), start=1):
        c = row.cells
        extra = {
            sarlavha_normal(k): v
            for k, v in o.get("extra", {}).items()
        }
        for index, (header, cell) in enumerate(zip(sarlavhalar, c)):
            kind = sarlavha_turi(header)
            if kind == "tr":
                matn_yoz(cell, str(i))
            elif kind == "name":
                matn_yoz(cell, o.get("ism", ""))
            elif kind == "class_number":
                matn_yoz(cell, raqam)
            elif kind == "class_letter":
                matn_yoz(cell, harf)
            elif kind == "phone":
                matn_yoz(cell, o.get("tel", ""))
            elif kind == "age":
                matn_yoz(cell, o.get("yosh", ""))
            elif kind == "address":
                matn_yoz(cell, o.get("joy", "") or extra.get(sarlavha_normal(header), ""))
            elif kind == "gender":
                jins = jins_ol(o.get("ism", "")) if jins_kerak else o.get("jins", "")
                matn_yoz(cell, jins)
            elif kind == "photo":
                rasm_qo_y(cell, o["rasm"])
            elif extra.get(sarlavha_normal(header)) is not None:
                matn_yoz(cell, str(extra[sarlavha_normal(header)]))

    # shablondagi eski, ishlatilmay qolgan rasm fayllarini paketdan tashlaymiz
    ishlatilgan = set(doc.element.body.xpath("//a:blip/@r:embed"))
    for rid, rel in list(doc.part.rels.items()):
        if rel.reltype.endswith("/image") and rid not in ishlatilgan:
            doc.part.drop_rel(rid)

    chiqish.mkdir(parents=True, exist_ok=True)
    nom = chiqish / f"{raqam}{harf}.docx"
    doc.save(str(nom))
    return nom


# ---- asosiy ----------------------------------------------------------
def sinf_normal(t):
    m = SINF_NOM_RE.fullmatch(t)
    if not m:
        sys.exit(f"Sinf noto'g'ri yozilgan: {t} (masalan 7A yoki 3YE)")
    return m.group(1), harfni_normal(m.group(2))


def sinf_tartib(nom):
    raqam, harf = sinf_normal(nom)
    return int(raqam), harf


def main():
    ap = argparse.ArgumentParser(description="Telegram -> Word shablon")
    ap.add_argument("manba", nargs="*",
                    help="eksport papkalari va shablon .docx. Chatda sinf yozilmagan bo'lsa: papka=7A")
    ap.add_argument("--chiqish", default="natija")
    ap.add_argument("--jins", action="store_true")
    ap.add_argument("--faqat", nargs="*", help="faqat shu sinflar, masalan: 7A 10B")
    a = ap.parse_args()

    manba = a.manba
    if not manba:
        manba = [input("Telegram eksport papkasi yo'li: ").strip().strip('"'),
                 input("Word shablon fayli (masalan 7A.docx): ").strip().strip('"')]
    shablon = next((Path(x) for x in manba if x.lower().endswith(".docx")), None)
    papkalar = [x for x in manba if not x.lower().endswith(".docx")]
    if not shablon or not shablon.exists():
        sys.exit("Xato: shablon (.docx) ko'rsatilmagan yoki topilmadi.")
    if not papkalar:
        sys.exit("Xato: eksport papkasi ko'rsatilmagan.")

    sinflar = {}
    for x in papkalar:
        yol, _, sinf = x.partition("=")
        yol = Path(yol)
        if not (yol / "result.json").exists():
            sys.exit(f"Xato: {yol}/result.json topilmadi. Eksportni JSON formatda qiling.")
        eksportni_oqi(yol, sinf_normal(sinf) if sinf else None, sinflar)
    if not sinflar:
        sys.exit("Hech qanday sinf/rasm topilmadi. Eksportda Photos belgilanganini tekshiring.")

    tanlangan = {sinf_normal(s) for s in (a.faqat or [])}
    for sinf, lst in sinflar.items():
        if tanlangan and sinf not in tanlangan:
            continue
        nom = sinfni_yoz(shablon, sinf, lst, Path(a.chiqish), a.jins)
        ismsiz = [i + 1 for i, o in enumerate(lst) if not o["ism"]]
        print(f"{sinf[0]}{sinf[1]}: {len(lst)} ta o'quvchi -> {nom}")
        if ismsiz:
            print(f"   DIQQAT: {ismsiz} -tartibdagi rasmlarda ism topilmadi (bo'sh qoldi)")


if __name__ == "__main__":
    main()
