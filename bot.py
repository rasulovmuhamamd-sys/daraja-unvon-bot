import asyncio
import io
import logging
import os
import re
import sqlite3
import zipfile
from datetime import datetime
from html import escape

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    BufferedInputFile, CallbackQuery, InlineKeyboardButton, InlineKeyboardMarkup, Message,
)
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

# ================== SOZLAMALAR ==================
BOT_TOKEN = os.getenv("BOT_TOKEN")

ADMINS = [5137395017, 399336424, 8320070405, 8576050673]

# Majburiy obuna kanallari. Bot har bir kanalda ADMIN bo'lishi shart.
# "id" -> @username (ochiq kanal) yoki -100... (yopiq kanal), "url" -> havola
CHANNELS = [
    {"id": "@kanal_1", "title": "1-kanal", "url": "https://t.me/kanal_1"},
    {"id": "@kanal_2", "title": "2-kanal", "url": "https://t.me/kanal_2"},
]

DB_PATH = "applications.db"

DESCRIPTION = (
    "Bu bot Qo'qon universitetiga ilmiy daraja (PhD, DSc) va ilmiy unvon "
    "(dotsent, professor, akademik) bo'yicha ariza topshirish uchun mo'ljallangan.\n\n"
    "Botdan foydalanish uchun avval quyidagi kanallarga obuna bo'ling, "
    "so'ng START tugmasini bosing."
)
SHORT_DESCRIPTION = "Qo'qon universiteti daraja va unvonlariga ariza topshirish boti"
# =================================================

logging.basicConfig(level=logging.INFO)
router = Router()


# ---------------------- DB ----------------------
def db():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with db() as con:
        con.execute(
            """CREATE TABLE IF NOT EXISTS applications(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER, username TEXT,
                ism TEXT, familya TEXT, otasi TEXT,
                tur TEXT, unvon TEXT, kafedra TEXT,
                diplom TEXT, sana TEXT, bayon TEXT,
                file_id TEXT,
                status TEXT DEFAULT 'pending',   -- pending | rejecting | accepted | rejected
                reason TEXT, decided_by INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )"""
        )
        con.execute(
            """CREATE TABLE IF NOT EXISTS admin_msgs(
                app_id INTEGER, admin_id INTEGER, message_id INTEGER
            )"""
        )


def get_app(app_id):
    with db() as con:
        return con.execute("SELECT * FROM applications WHERE id=?", (app_id,)).fetchone()


# ---------------------- FSM ----------------------
class Form(StatesGroup):
    ism = State()
    familya = State()
    otasi = State()
    tur = State()
    unvon = State()
    daraja = State()
    kafedra = State()
    diplom = State()
    sana = State()
    bayon = State()
    pdf = State()


class Reject(StatesGroup):
    reason = State()


# ---------------------- Klaviaturalar ----------------------
def sub_kb():
    rows = [[InlineKeyboardButton(text=f"📢 {c['title']}", url=c["url"])] for c in CHANNELS]
    rows.append([InlineKeyboardButton(text="✅ Obunani tekshirish", callback_data="check_sub")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def tur_kb():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🎓 Ilmiy daraja", callback_data="tur:daraja")],
            [InlineKeyboardButton(text="🏅 Ilmiy unvon", callback_data="tur:unvon")],
        ]
    )


def unvon_kb():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="Dotsent", callback_data="unvon:Dotsent")],
            [InlineKeyboardButton(text="Professor", callback_data="unvon:Professor")],
            [InlineKeyboardButton(text="Akademik", callback_data="unvon:Akademik")],
        ]
    )


def daraja_kb():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="PhD", callback_data="daraja:PhD"),
                InlineKeyboardButton(text="DSc", callback_data="daraja:DSc"),
            ]
        ]
    )


def admin_kb(app_id):
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="✅ Qabul qilish", callback_data=f"ok:{app_id}"),
                InlineKeyboardButton(text="❌ Rad etish", callback_data=f"no:{app_id}"),
            ]
        ]
    )


# ---------------------- Yordamchilar ----------------------
async def is_subscribed(bot: Bot, user_id: int) -> bool:
    return True
    for c in CHANNELS:
        try:
            m = await bot.get_chat_member(c["id"], user_id)
            if m.status in ("left", "kicked"):
                return False
        except Exception as e:
            logging.warning("Obunani tekshirib bo'lmadi (%s): %s", c["id"], e)
            return False
    return True


def app_text(a, status_line=""):
    fio = f"{a['familya']} {a['ism']} {a['otasi']}"
    t = (
        f"📄 <b>Ariza №{a['id']}</b>\n\n"
        f"👤 F.I.Sh.: <b>{escape(fio)}</b>\n"
        f"📌 Turi: {escape(a['tur'])}\n"
        f"🎖 Tanlangan: <b>{escape(a['unvon'])}</b>\n"
        f"🏛 Kafedra: {escape(a['kafedra'])}\n"
        f"🔢 Diplom raqami: {escape(a['diplom'])}\n"
        f"📅 Berilgan sana: {escape(a['sana'])}\n"
        f"📝 Bayon: {escape(a['bayon'])}\n"
    )
    if a["username"]:
        t += f"🔗 Telegram: @{escape(a['username'])}\n"
    t += f"🆔 User ID: <code>{a['user_id']}</code>"
    if status_line:
        t += f"\n\n{status_line}"
    return t


def safe_name(text: str) -> str:
    return re.sub(r"[^\w]+", "_", text, flags=re.UNICODE).strip("_") or "ariza"


def build_docx(a) -> bytes:
    """Foydalanuvchi to'ldirgan ma'lumotlardan Word (docx) ariza yaratadi."""
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(13)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run("Qo'qon universiteti\nIlmiy kengashiga").bold = True

    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(f"ARIZA №{a['id']}")
    r.bold = True
    r.font.size = Pt(16)

    fio = f"{a['familya']} {a['ism']} {a['otasi']}"
    rows = [
        ("F.I.Sh.", fio),
        ("Ariza turi", a["tur"]),
        ("Tanlangan daraja / unvon", a["unvon"]),
        ("Ishlayotgan kafedra", a["kafedra"]),
        ("Diplom raqami", a["diplom"]),
        ("Diplom berilgan sana", a["sana"]),
        ("Bayon nomi va raqami", a["bayon"]),
        ("Telegram", f"@{a['username']}" if a["username"] else "-"),
    ]
    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    for k, v in rows:
        cells = table.add_row().cells
        cells[0].text = k
        cells[1].text = str(v)
        for run in cells[0].paragraphs[0].runs:
            run.bold = True

    doc.add_paragraph()
    doc.add_paragraph(f"Ariza topshirilgan sana: {datetime.now():%d.%m.%Y}")
    doc.add_paragraph("Imzo: ____________________")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_zip(a, pdf_bytes: bytes) -> bytes:
    base = f"{safe_name(a['familya'])}_{safe_name(a['ism'])}"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"Ariza_{base}.docx", build_docx(a))
        z.writestr(f"Diplom_{base}.pdf", pdf_bytes)
    return buf.getvalue()


async def start_form(msg: Message, state: FSMContext):
    await state.clear()
    await state.set_state(Form.ism)
    await msg.answer("Arizani boshlaymiz.\n\n1️⃣ <b>Ismingizni</b> kiriting:")


async def finalize_for_admins(bot: Bot, app_id: int, status_line: str):
    """Barcha adminlardagi xabarni yangilaydi va tugmalarni olib tashlaydi."""
    a = get_app(app_id)
    with db() as con:
        msgs = con.execute("SELECT * FROM admin_msgs WHERE app_id=?", (app_id,)).fetchall()
    for m in msgs:
        try:
            await bot.edit_message_caption(
                chat_id=m["admin_id"],
                message_id=m["message_id"],
                caption=app_text(a, status_line),
                reply_markup=None,
            )
        except Exception as e:
            logging.warning("Admin xabarini yangilab bo'lmadi: %s", e)


# ---------------------- /start va obuna ----------------------
@router.message(CommandStart())
async def cmd_start(msg: Message, state: FSMContext, bot: Bot):
    await state.clear()
    if not await is_subscribed(bot, msg.from_user.id):
        await msg.answer(
            "Botdan foydalanish uchun quyidagi kanallarga obuna bo'ling, "
            "so'ng «Obunani tekshirish» tugmasini bosing 👇",
            reply_markup=sub_kb(),
        )
        return
    await start_form(msg, state)


@router.callback_query(F.data == "check_sub")
async def check_sub(cb: CallbackQuery, state: FSMContext, bot: Bot):
    if not await is_subscribed(bot, cb.from_user.id):
        await cb.answer("Siz hali barcha kanallarga obuna bo'lmadingiz!", show_alert=True)
        return
    await cb.answer("Obuna tasdiqlandi ✅")
    try:
        await cb.message.delete()
    except Exception:
        pass
    await start_form(cb.message, state)


# ---------------------- Ariza shakli ----------------------
@router.message(Form.ism, F.text)
async def f_ism(msg: Message, state: FSMContext):
    await state.update_data(ism=msg.text.strip())
    await state.set_state(Form.familya)
    await msg.answer("2️⃣ <b>Familiyangizni</b> kiriting:")


@router.message(Form.familya, F.text)
async def f_familya(msg: Message, state: FSMContext):
    await state.update_data(familya=msg.text.strip())
    await state.set_state(Form.otasi)
    await msg.answer("3️⃣ <b>Otangizning ismini</b> kiriting:")


@router.message(Form.otasi, F.text)
async def f_otasi(msg: Message, state: FSMContext):
    await state.update_data(otasi=msg.text.strip())
    await state.set_state(Form.tur)
    await msg.answer("4️⃣ Quyidagilardan birini tanlang:", reply_markup=tur_kb())


@router.callback_query(Form.tur, F.data.startswith("tur:"))
async def f_tur(cb: CallbackQuery, state: FSMContext):
    kind = cb.data.split(":")[1]
    await cb.answer()
    if kind == "unvon":
        await state.update_data(tur="Ilmiy unvon")
        await state.set_state(Form.unvon)
        await cb.message.edit_text("5️⃣ Ilmiy unvonni tanlang:", reply_markup=unvon_kb())
    else:
        await state.update_data(tur="Ilmiy daraja")
        await state.set_state(Form.daraja)
        await cb.message.edit_text("5️⃣ Ilmiy darajani tanlang:", reply_markup=daraja_kb())


@router.callback_query(Form.unvon, F.data.startswith("unvon:"))
@router.callback_query(Form.daraja, F.data.startswith("daraja:"))
async def f_unvon_daraja(cb: CallbackQuery, state: FSMContext):
    value = cb.data.split(":")[1]
    await cb.answer()
    await state.update_data(unvon=value)
    await cb.message.edit_text(f"5️⃣ Tanlandi: <b>{value}</b>")
    await state.set_state(Form.kafedra)
    await cb.message.answer("6️⃣ Ishlayotgan <b>kafedra nomini</b> kiriting:")


@router.message(Form.kafedra, F.text)
async def f_kafedra(msg: Message, state: FSMContext):
    await state.update_data(kafedra=msg.text.strip())
    await state.set_state(Form.diplom)
    await msg.answer("7️⃣ <b>Diplom raqamini</b> kiriting:")


@router.message(Form.diplom, F.text)
async def f_diplom(msg: Message, state: FSMContext):
    await state.update_data(diplom=msg.text.strip())
    await state.set_state(Form.sana)
    await msg.answer("8️⃣ Diplom <b>berilgan sanani</b> kiriting (masalan: 15.06.2020):")


@router.message(Form.sana, F.text)
async def f_sana(msg: Message, state: FSMContext):
    text = msg.text.strip()
    if not re.fullmatch(r"\d{1,2}[./-]\d{1,2}[./-]\d{4}", text):
        await msg.answer("❗️ Sana formati noto'g'ri. Masalan: <b>15.06.2020</b>")
        return
    await state.update_data(sana=text)
    await state.set_state(Form.bayon)
    await msg.answer("9️⃣ <b>Bayon nomi va raqamini</b> kiriting:")


@router.message(Form.bayon, F.text)
async def f_bayon(msg: Message, state: FSMContext):
    await state.update_data(bayon=msg.text.strip())
    await state.set_state(Form.pdf)
    await msg.answer("🔟 Diplomingizni <b>PDF</b> holatida yuboring:")


@router.message(Form.pdf, F.document)
async def f_pdf(msg: Message, state: FSMContext, bot: Bot):
    d = msg.document
    is_pdf = d.mime_type == "application/pdf" or (d.file_name or "").lower().endswith(".pdf")
    if not is_pdf:
        await msg.answer("❗️ Faqat <b>PDF</b> fayl qabul qilinadi. Qaytadan yuboring.")
        return
    if d.file_size and d.file_size > 20 * 1024 * 1024:
        await msg.answer("❗️ Fayl hajmi 20 MB dan oshmasligi kerak. Kichikroq PDF yuboring.")
        return

    # PDF ni yuklab olamiz
    try:
        pdf_buf = io.BytesIO()
        await bot.download(d, destination=pdf_buf)
        pdf_bytes = pdf_buf.getvalue()
    except Exception as e:
        logging.exception("PDF yuklab olinmadi: %s", e)
        await msg.answer("❗️ Faylni qabul qilib bo'lmadi. Qaytadan yuboring.")
        return

    data = await state.get_data()
    await state.clear()

    with db() as con:
        cur = con.execute(
            """INSERT INTO applications
               (user_id, username, ism, familya, otasi, tur, unvon, kafedra, diplom, sana, bayon, file_id)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                msg.from_user.id, msg.from_user.username,
                data["ism"], data["familya"], data["otasi"],
                data["tur"], data["unvon"], data["kafedra"],
                data["diplom"], data["sana"], data["bayon"], d.file_id,
            ),
        )
        app_id = cur.lastrowid
    a = get_app(app_id)

    # Word + PDF -> ZIP
    zip_name = f"{safe_name(a['familya'])}_{safe_name(a['ism'])}.zip"
    zip_file = BufferedInputFile(build_zip(a, pdf_bytes), filename=zip_name)

    # Foydalanuvchiga nusxa (ZIP bir marta yuklanadi, keyin file_id qayta ishlatiladi)
    sent_user = await msg.answer_document(
        zip_file,
        caption=app_text(a, "⏳ Arizangiz adminlarga yuborildi. Natija haqida xabar beramiz."),
    )
    zip_id = sent_user.document.file_id

    # 4 ta adminga
    for admin_id in ADMINS:
        try:
            sent = await bot.send_document(
                admin_id, zip_id, caption=app_text(a), reply_markup=admin_kb(app_id)
            )
            with db() as con:
                con.execute(
                    "INSERT INTO admin_msgs(app_id, admin_id, message_id) VALUES (?,?,?)",
                    (app_id, admin_id, sent.message_id),
                )
        except Exception as e:
            logging.warning("Adminga (%s) yuborib bo'lmadi: %s", admin_id, e)


@router.message(Form.pdf)
async def f_pdf_wrong(msg: Message):
    await msg.answer("❗️ Iltimos, diplomni <b>PDF fayl</b> sifatida yuboring.")


# ---------------------- Admin: qabul / rad ----------------------
@router.callback_query(F.data.startswith("ok:"))
async def admin_accept(cb: CallbackQuery, bot: Bot):
    if cb.from_user.id not in ADMINS:
        await cb.answer("Ruxsat yo'q", show_alert=True)
        return
    app_id = int(cb.data.split(":")[1])
    with db() as con:
        cur = con.execute(
            "UPDATE applications SET status='accepted', decided_by=? WHERE id=? AND status='pending'",
            (cb.from_user.id, app_id),
        )
        changed = cur.rowcount
    if not changed:
        await cb.answer("Bu ariza bo'yicha qaror allaqachon qabul qilingan.", show_alert=True)
        a = get_app(app_id)
        try:
            await cb.message.edit_reply_markup(reply_markup=None)
        except Exception:
            pass
        return
    await cb.answer("Qabul qilindi ✅")
    await finalize_for_admins(bot, app_id, f"✅ <b>QABUL QILINDI</b> (admin: <code>{cb.from_user.id}</code>)")
    a = get_app(app_id)
    try:
        await bot.send_message(
            a["user_id"], f"✅ Hurmatli {escape(a['ism'])}, <b>Ariza №{app_id}</b> qabul qilindi."
        )
    except Exception:
        pass


@router.callback_query(F.data.startswith("no:"))
async def admin_reject(cb: CallbackQuery, state: FSMContext):
    if cb.from_user.id not in ADMINS:
        await cb.answer("Ruxsat yo'q", show_alert=True)
        return
    app_id = int(cb.data.split(":")[1])
    with db() as con:
        cur = con.execute(
            "UPDATE applications SET status='rejecting', decided_by=? WHERE id=? AND status='pending'",
            (cb.from_user.id, app_id),
        )
        changed = cur.rowcount
    if not changed:
        await cb.answer("Bu ariza bo'yicha qaror allaqachon qabul qilingan yoki ko'rib chiqilmoqda.", show_alert=True)
        return
    await cb.answer()
    await state.set_state(Reject.reason)
    await state.update_data(app_id=app_id)
    await cb.message.answer(
        f"❌ Ariza №{app_id} rad etilmoqda.\n<b>Rad etish sababini</b> yozing "
        f"(bekor qilish: /cancel):"
    )


@router.message(Reject.reason, Command("cancel"))
async def reject_cancel(msg: Message, state: FSMContext):
    data = await state.get_data()
    await state.clear()
    with db() as con:
        con.execute(
            "UPDATE applications SET status='pending', decided_by=NULL WHERE id=? AND status='rejecting'",
            (data["app_id"],),
        )
    await msg.answer("Bekor qilindi. Ariza yana ko'rib chiqish uchun ochiq.")


@router.message(Reject.reason, F.text)
async def reject_reason(msg: Message, state: FSMContext, bot: Bot):
    data = await state.get_data()
    app_id = data["app_id"]
    reason = msg.text.strip()
    await state.clear()
    with db() as con:
        con.execute(
            "UPDATE applications SET status='rejected', reason=?, decided_by=? WHERE id=?",
            (reason, msg.from_user.id, app_id),
        )
    a = get_app(app_id)
    await finalize_for_admins(
        bot, app_id,
        f"❌ <b>RAD ETILDI</b> (admin: <code>{msg.from_user.id}</code>)\n"
        f"Sabab: {escape(reason)}",
    )
    await msg.answer("Ariza rad etildi, sabab foydalanuvchiga yuborildi.")
    try:
        await bot.send_message(
            a["user_id"],
            f"❌ Hurmatli {escape(a['ism'])}, <b>Ariza №{app_id}</b> rad etildi.\n\n"
            f"<b>Sabab:</b> {escape(reason)}\n\n"
            f"Xatolarni to'g'rilab, /start orqali qaytadan ariza topshirishingiz mumkin.",
        )
    except Exception:
        pass


# ---------------------- Ishga tushirish ----------------------
async def main():
    init_db()
    bot = Bot(BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)

    # Start bosilishidan oldin yuqorida ko'rinadigan matn ("Bu bot nima qila oladi")
    try:
        await bot.set_my_description(DESCRIPTION)
        await bot.set_my_short_description(SHORT_DESCRIPTION)
    except Exception as e:
        logging.warning("Tavsifni o'rnatib bo'lmadi: %s", e)

    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
    