import asyncio
import io
import math
import os
import sqlite3
from datetime import datetime
from pathlib import Path

from aiogram import Bot, Dispatcher, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import BufferedInputFile, CallbackQuery, Message
from aiogram.utils.keyboard import InlineKeyboardBuilder
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from docx import Document
from docx.shared import Pt

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
DB_PATH = os.getenv("DB_PATH", "data.sqlite3")
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is not set in .env")

Path(DB_PATH).parent.mkdir(parents=True, exist_ok=True)
conn = sqlite3.connect(DB_PATH, check_same_thread=False)
conn.row_factory = sqlite3.Row


def db_init():
    conn.execute("PRAGMA foreign_keys=ON")
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, username TEXT DEFAULT '', first_name TEXT DEFAULT '')
    CREATE TABLE IF NOT EXISTS objects (
        id INTEGER PRIMARY KEY AUTOINCREMENT, owner_id INTEGER NOT NULL, name TEXT NOT NULL,
        address TEXT DEFAULT '', customer TEXT DEFAULT '', phone TEXT DEFAULT '', created_at TEXT NOT NULL,
        FOREIGN KEY(owner_id) REFERENCES users(id) ON DELETE CASCADE
    )
    CREATE TABLE IF NOT EXISTS rooms (
        id INTEGER PRIMARY KEY AUTOINCREMENT, object_id INTEGER NOT NULL, name TEXT NOT NULL,
        length REAL NOT NULL, width REAL NOT NULL, height REAL NOT NULL,
        FOREIGN KEY(object_id) REFERENCES objects(id) ON DELETE CASCADE
    )
    CREATE TABLE IF NOT EXISTS openings (
        id INTEGER PRIMARY KEY AUTOINCREMENT, room_id INTEGER NOT NULL, kind TEXT NOT NULL,
        width REAL NOT NULL, height REAL NOT NULL, qty INTEGER NOT NULL DEFAULT 1,
        FOREIGN KEY(room_id) REFERENCES rooms(id) ON DELETE CASCADE
    )
    CREATE TABLE IF NOT EXISTS work_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT, owner_id INTEGER NOT NULL, name TEXT NOT NULL,
        unit TEXT NOT NULL, price REAL NOT NULL,
        UNIQUE(owner_id, name), FOREIGN KEY(owner_id) REFERENCES users(id) ON DELETE CASCADE
    )
    CREATE TABLE IF NOT EXISTS materials (
        id INTEGER PRIMARY KEY AUTOINCREMENT, owner_id INTEGER NOT NULL, name TEXT NOT NULL,
        sale_unit TEXT NOT NULL, pack_qty REAL NOT NULL, pack_price REAL NOT NULL,
        consumption REAL NOT NULL DEFAULT 0, consumption_unit TEXT DEFAULT '',
        UNIQUE(owner_id, name), FOREIGN KEY(owner_id) REFERENCES users(id) ON DELETE CASCADE
    )
    CREATE TABLE IF NOT EXISTS estimate_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT, object_id INTEGER NOT NULL,
        kind TEXT NOT NULL CHECK(kind IN ('work','material')), name TEXT NOT NULL,
        qty REAL NOT NULL, unit TEXT NOT NULL, unit_price REAL NOT NULL, total REAL NOT NULL,
        note TEXT DEFAULT '', FOREIGN KEY(object_id) REFERENCES objects(id) ON DELETE CASCADE
    )
    """)
    conn.commit()


def q(sql, args=(), one=False):
    cur = conn.execute(sql, args)
    return cur.fetchone() if one else cur.fetchall()


def exec_sql(sql, args=()):
    cur = conn.execute(sql, args)
    conn.commit()
    return cur.lastrowid


def money(v):
    return f"{v:,.0f} ₽".replace(",", " ")


def num(v):
    return f"{v:g}"


def parse_float(text):
    s = (text or "").strip().replace(" ", "").replace(",", ".")
    return float(s)


def valid_positive(v):
    return math.isfinite(v) and v > 0


def user_upsert(message: Message):
    u = message.from_user
    exec_sql("INSERT INTO users(id,username,first_name) VALUES(?,?,?) ON CONFLICT(id) DO UPDATE SET username=excluded.username, first_name=excluded.first_name", (u.id, u.username or "", u.first_name or ""))


def owner_object(user_id, oid):
    return q("SELECT * FROM objects WHERE id=? AND owner_id=?", (oid, user_id), True)


def owner_room(user_id, rid):
    return q("SELECT r.*,o.owner_id,o.name AS object_name FROM rooms r JOIN objects o ON o.id=r.object_id WHERE r.id=? AND o.owner_id=?", (rid, user_id), True)


def room_metrics(rid):
    r = q("SELECT * FROM rooms WHERE id=?", (rid,), True)
    if not r:
        return None
    raw_wall = 2 * (r["length"] + r["width"]) * r["height"]
    floor = r["length"] * r["width"]
    openings = q("SELECT width,height,qty FROM openings WHERE room_id=?", (rid,))
    deducted = sum(x["width"] * x["height"] * x["qty"] for x in openings)
    return {"wall": max(0, raw_wall - deducted), "raw_wall": raw_wall, "floor": floor, "ceiling": floor, "openings": deducted}


def object_totals(oid):
    rows = q("SELECT kind,total FROM estimate_items WHERE object_id=?", (oid,))
    work = sum(r["total"] for r in rows if r["kind"] == "work")
    mat = sum(r["total"] for r in rows if r["kind"] == "material")
    return work, mat


class NewObject(StatesGroup):
    name = State(); address = State(); customer = State(); phone = State()


class RoomForm(StatesGroup):
    object_id = State(); name = State(); length = State(); width = State(); height = State()


class OpeningForm(StatesGroup):
    room_id = State(); kind = State(); width = State(); height = State(); qty = State()


class WorkForm(StatesGroup):
    object_id = State(); name = State(); qty = State(); unit = State(); price = State()


class MaterialForm(StatesGroup):
    object_id = State(); material_id = State(); area = State(); waste = State()


class WorkPriceForm(StatesGroup):
    name = State(); unit = State(); price = State()


class MaterialCatalogForm(StatesGroup):
    name = State(); sale_unit = State(); pack_qty = State(); pack_price = State(); consumption = State(); consumption_unit = State()


class QuickForm(StatesGroup):
    name = State(); qty = State(); unit = State(); price = State()


class EditObjectForm(StatesGroup):
    oid = State(); field = State(); value = State()


def main_kb():
    b = InlineKeyboardBuilder()
    b.button(text="➕ Новый объект", callback_data="new_object")
    b.button(text="📂 Мои объекты", callback_data="objects")
    b.button(text="📐 Быстрый расчёт", callback_data="quick")
    b.button(text="🧑‍🔧 Цены работ", callback_data="work_prices")
    b.button(text="🧱 Материалы", callback_data="materials")
    b.adjust(1, 1, 2, 1)
    return b.as_markup()


def back_home_kb():
    b = InlineKeyboardBuilder(); b.button(text="⬅️ Главное меню", callback_data="home"); return b.as_markup()


def object_kb(oid):
    b = InlineKeyboardBuilder()
    b.button(text="📐 Комнаты", callback_data=f"rooms:{oid}")
    b.button(text="➕ Работа", callback_data=f"addwork:{oid}")
    b.button(text="🧱 Материал", callback_data=f"addmat:{oid}")
    b.button(text="📋 Смета", callback_data=f"estimate:{oid}")
    b.button(text="✏️ Объект", callback_data=f"editobj:{oid}")
    b.button(text="📊 Excel", callback_data=f"excel:{oid}")
    b.button(text="📝 Word", callback_data=f"word:{oid}")
    b.button(text="🗑 Удалить", callback_data=f"delobj:{oid}")
    b.button(text="⬅️ Объекты", callback_data="objects")
    b.adjust(2, 2, 2, 2, 1)
    return b.as_markup()


def cancel_kb():
    b = InlineKeyboardBuilder(); b.button(text="❌ Отмена", callback_data="cancel"); return b.as_markup()


async def safe_answer(callback, text=""):
    try:
        await callback.answer(text)
    except Exception:
        pass


async def send_error(message, exc):
    # Do not leak internal DB paths/SQL to the user.
    print(f"ERROR: {type(exc).__name__}: {exc}")
    await message.answer("⚠️ Не удалось выполнить действие. Попробуйте ещё раз или нажмите /cancel.")


async def start(message: Message, state: FSMContext):
    await state.clear(); user_upsert(message)
    await message.answer("🏗 <b>СтройСмета</b>\n\nРаботы + материалы + автоматический расчёт.\n\nВыберите действие:", reply_markup=main_kb())


async def home(callback: CallbackQuery, state: FSMContext):
    await state.clear(); await callback.message.edit_text("🏗 <b>СтройСмета</b>\n\nВыберите действие:", reply_markup=main_kb()); await safe_answer(callback)


async def cancel(callback: CallbackQuery, state: FSMContext):
    await state.clear(); await callback.message.answer("❌ Действие отменено.", reply_markup=main_kb()); await safe_answer(callback)


async def cancel_cmd(message: Message, state: FSMContext):
    await state.clear(); await message.answer("❌ Действие отменено.", reply_markup=main_kb())


# ---------- Objects ----------
async def new_object(callback: CallbackQuery, state: FSMContext):
    await state.set_state(NewObject.name); await callback.message.answer("🏠 Название объекта?", reply_markup=cancel_kb()); await safe_answer(callback)


async def new_object_name(message, state):
    text = message.text.strip()
    if not text: return await message.answer("Название не может быть пустым.")
    await state.update_data(name=text); await state.set_state(NewObject.address); await message.answer("📍 Адрес?\nМожно написать —, чтобы пропустить.")


async def new_object_address(message, state):
    await state.update_data(address="" if message.text.strip()=="—" else message.text.strip()); await state.set_state(NewObject.customer); await message.answer("👤 Заказчик?\nМожно —.")


async def new_object_customer(message, state):
    await state.update_data(customer="" if message.text.strip()=="—" else message.text.strip()); await state.set_state(NewObject.phone); await message.answer("📞 Телефон?\nМожно —.")


async def new_object_phone(message, state):
    try:
        d=await state.get_data(); phone="" if message.text.strip()=="—" else message.text.strip()
        oid=exec_sql("INSERT INTO objects(owner_id,name,address,customer,phone,created_at) VALUES(?,?,?,?,?,?)",(message.from_user.id,d["name"],d["address"],d["customer"],phone,datetime.now().isoformat()))
        await state.clear(); await message.answer(f"✅ Объект создан: <b>{d['name']}</b>", reply_markup=object_kb(oid))
    except Exception as e: await send_error(message,e)


async def objects(callback):
    rows=q("SELECT * FROM objects WHERE owner_id=? ORDER BY id DESC",(callback.from_user.id,))
    b=InlineKeyboardBuilder()
    for r in rows: b.button(text=f"🏠 {r['name']}",callback_data=f"openobj:{r['id']}")
    b.button(text="➕ Новый объект",callback_data="new_object"); b.button(text="⬅️ Главное меню",callback_data="home"); b.adjust(1)
    await callback.message.edit_text("📂 <b>Мои объекты</b>\n\n" + ("Выберите объект:" if rows else "Пока нет объектов."),reply_markup=b.as_markup()); await safe_answer(callback)


async def open_obj(callback):
    try:
        oid=int(callback.data.split(":")[1]); obj=owner_object(callback.from_user.id,oid)
        if not obj:return await callback.answer("Нет доступа",show_alert=True)
        rooms=q("SELECT id FROM rooms WHERE object_id=?",(oid,)); work,mat=object_totals(oid)
        text=f"🏠 <b>{obj['name']}</b>\n\n📍 {obj['address'] or 'Адрес не указан'}\n👤 {obj['customer'] or 'Заказчик не указан'}\n📐 Комнат: {len(rooms)}\n\n🧑‍🔧 Работа: <b>{money(work)}</b>\n🧱 Материалы: <b>{money(mat)}</b>\n💰 Итого: <b>{money(work+mat)}</b>"
        await callback.message.edit_text(text,reply_markup=object_kb(oid)); await safe_answer(callback)
    except Exception as e: await safe_answer(callback,"Ошибка"); print(e)


async def edit_object(callback,state):
    oid=int(callback.data.split(":")[1]); obj=owner_object(callback.from_user.id,oid)
    if not obj:return await callback.answer("Нет доступа",show_alert=True)
    b=InlineKeyboardBuilder()
    for key,label in [("name","🏠 Название"),("address","📍 Адрес"),("customer","👤 Заказчик"),("phone","📞 Телефон")]: b.button(text=label,callback_data=f"editfield:{oid}:{key}")
    b.button(text="⬅️ Объект",callback_data=f"openobj:{oid}"); b.adjust(2,2,1)
    await callback.message.edit_text("✏️ Что изменить?",reply_markup=b.as_markup()); await safe_answer(callback)


async def edit_field(callback,state):
    _,oid,field=callback.data.split(":"); oid=int(oid)
    if not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    await state.set_state(EditObjectForm.value); await state.update_data(oid=oid,field=field)
    labels={"name":"название","address":"адрес","customer":"заказчика","phone":"телефон"}
    await callback.message.answer(f"Введите новое значение: {labels[field]}.\nДля пустого значения — напишите —.",reply_markup=cancel_kb()); await safe_answer(callback)


async def edit_field_value(message,state):
    d=await state.get_data(); oid=d["oid"]; field=d["field"]
    if not owner_object(message.from_user.id,oid): await state.clear(); return await message.answer("Нет доступа.")
    val="" if message.text.strip()=="—" else message.text.strip()
    exec_sql(f"UPDATE objects SET {field}=? WHERE id=? AND owner_id=?",(val,oid,message.from_user.id)); await state.clear(); await message.answer("✅ Изменено.",reply_markup=object_kb(oid))


# ---------- Rooms and openings ----------
async def rooms(callback):
    oid=int(callback.data.split(":")[1]); obj=owner_object(callback.from_user.id,oid)
    if not obj:return await callback.answer("Нет доступа",show_alert=True)
    rows=q("SELECT * FROM rooms WHERE object_id=? ORDER BY id",(oid,)); b=InlineKeyboardBuilder()
    for r in rows:
        m=room_metrics(r["id"]); b.button(text=f"📐 {r['name']} — {m['floor']:.1f} м²",callback_data=f"room:{r['id']}")
    b.button(text="➕ Добавить комнату",callback_data=f"addroom:{oid}"); b.button(text="⬅️ Объект",callback_data=f"openobj:{oid}"); b.adjust(1)
    await callback.message.edit_text("📐 <b>Комнаты</b>\n\nВыберите комнату или добавьте новую.",reply_markup=b.as_markup()); await safe_answer(callback)


async def add_room(callback,state):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    await state.update_data(object_id=oid); await state.set_state(RoomForm.name); await callback.message.answer("Название комнаты? Например: Гостиная",reply_markup=cancel_kb()); await safe_answer(callback)


async def room_name(message,state): await state.update_data(name=message.text.strip()); await state.set_state(RoomForm.length); await message.answer("📏 Длина, м?")
async def room_length(message,state):
    try:
        v=parse_float(message.text)
        if not valid_positive(v): raise ValueError
        await state.update_data(length=v); await state.set_state(RoomForm.width); await message.answer("📏 Ширина, м?")
    except: await message.answer("Введите положительное число, например 6 или 5,5.")
async def room_width(message,state):
    try:
        v=parse_float(message.text)
        if not valid_positive(v): raise ValueError
        await state.update_data(width=v); await state.set_state(RoomForm.height); await message.answer("📏 Высота, м?")
    except: await message.answer("Введите положительное число, например 2,7.")
async def room_height(message,state):
    try:
        v=parse_float(message.text)
        if not valid_positive(v): raise ValueError
        d=await state.get_data(); rid=exec_sql("INSERT INTO rooms(object_id,name,length,width,height) VALUES(?,?,?,?,?)",(d['object_id'],d['name'],d['length'],d['width'],v)); await state.clear(); m=room_metrics(rid)
        await message.answer(f"✅ <b>{d['name']}</b> добавлена.\n\nПол: {m['floor']:.2f} м²\nПотолок: {m['ceiling']:.2f} м²\nСтены: {m['wall']:.2f} м²\n\nМожно добавить двери/окна или вернуться к объекту.",reply_markup=room_kb(rid))
    except: await message.answer("Введите положительное число, например 2,7.")


def room_kb(rid):
    r=owner_room_by_id(rid)
    oid=r["object_id"] if r else 0
    b=InlineKeyboardBuilder(); b.button(text="🚪 Дверь/окно",callback_data=f"addopening:{rid}"); b.button(text="🗑 Удалить комнату",callback_data=f"delroom:{rid}"); b.button(text="⬅️ Комнаты",callback_data=f"rooms:{oid}"); b.adjust(1)
    return b.as_markup()


def owner_room_by_id(rid): return q("SELECT * FROM rooms WHERE id=?",(rid,),True)


async def room_view(callback):
    rid=int(callback.data.split(":")[1]); r=owner_room(callback.from_user.id,rid)
    if not r:return await callback.answer("Нет доступа",show_alert=True)
    m=room_metrics(rid); opens=q("SELECT * FROM openings WHERE room_id=? ORDER BY id",(rid,))
    text=f"📐 <b>{r['name']}</b>\n\nПол: {m['floor']:.2f} м²\nПотолок: {m['ceiling']:.2f} м²\nСтены: {m['wall']:.2f} м²\nПроёмы: {m['openings']:.2f} м²"
    if opens:text += "\n\n🚪 Проёмы:\n" + "\n".join(f"• {o['kind']} {o['width']}×{o['height']} × {o['qty']}" for o in opens)
    await callback.message.edit_text(text,reply_markup=room_kb(rid)); await safe_answer(callback)


async def add_opening(callback,state):
    rid=int(callback.data.split(":")[1]); r=owner_room(callback.from_user.id,rid)
    if not r:return await callback.answer("Нет доступа",show_alert=True)
    await state.update_data(room_id=rid); await state.set_state(OpeningForm.kind); b=InlineKeyboardBuilder(); b.button(text="🚪 Дверь",callback_data="openingkind:Дверь"); b.button(text="🪟 Окно",callback_data="openingkind:Окно"); b.adjust(2); await callback.message.answer("Что добавить?",reply_markup=b.as_markup()); await safe_answer(callback)


async def opening_kind(callback,state):
    await state.update_data(kind=callback.data.split(":",1)[1]); await state.set_state(OpeningForm.width); await callback.message.answer("Ширина, м? Например 0,9"); await safe_answer(callback)
async def opening_width(message,state):
    try:
        v=parse_float(message.text)
        if not valid_positive(v): raise ValueError
        await state.update_data(width=v); await state.set_state(OpeningForm.height); await message.answer("Высота, м?")
    except: await message.answer("Введите положительное число.")
async def opening_height(message,state):
    try:
        v=parse_float(message.text)
        if not valid_positive(v): raise ValueError
        await state.update_data(height=v); await state.set_state(OpeningForm.qty); await message.answer("Количество, шт? Например 1")
    except: await message.answer("Введите положительное число.")
async def opening_qty(message,state):
    try:
        qty=int(parse_float(message.text));
        if qty<1: raise ValueError
        d=await state.get_data(); exec_sql("INSERT INTO openings(room_id,kind,width,height,qty) VALUES(?,?,?,?,?)",(d['room_id'],d['kind'],d['width'],d['height'],qty)); await state.clear(); await message.answer("✅ Проём добавлен.",reply_markup=room_kb(d['room_id']))
    except: await message.answer("Введите целое количество, например 1.")


async def del_room(callback):
    rid=int(callback.data.split(":")[1]); r=owner_room(callback.from_user.id,rid)
    if not r:return await callback.answer("Нет доступа",show_alert=True)
    oid=r['object_id']; exec_sql("DELETE FROM rooms WHERE id=?",(rid,)); await callback.message.edit_text("🗑 Комната удалена.",reply_markup=object_kb(oid)); await safe_answer(callback)


# ---------- Work prices / work items ----------
async def work_prices(callback):
    rows=q("SELECT * FROM work_prices WHERE owner_id=? ORDER BY name",(callback.from_user.id,)); b=InlineKeyboardBuilder()
    for r in rows:b.button(text=f"🧑‍🔧 {r['name']} — {money(r['price'])}/{r['unit']}",callback_data=f"wpview:{r['id']}")
    b.button(text="➕ Добавить/изменить цену",callback_data="addworkprice"); b.button(text="⬅️ Назад",callback_data="home"); b.adjust(1)
    await callback.message.edit_text("🧑‍🔧 <b>Цены на работы</b>\n\nЭти цены можно использовать при добавлении работы.",reply_markup=b.as_markup()); await safe_answer(callback)


async def add_work_price(callback,state): await state.set_state(WorkPriceForm.name); await callback.message.answer("Название работы?",reply_markup=cancel_kb()); await safe_answer(callback)
async def wp_name(message,state): await state.update_data(name=message.text.strip()); await state.set_state(WorkPriceForm.unit); await message.answer("Единица? Например м², м.п., шт")
async def wp_unit(message,state): await state.update_data(unit=message.text.strip()); await state.set_state(WorkPriceForm.price); await message.answer("Цена, ₽?")
async def wp_price(message,state):
    try:
        p=parse_float(message.text)
        if p<0: raise ValueError
        d=await state.get_data(); exec_sql("INSERT INTO work_prices(owner_id,name,unit,price) VALUES(?,?,?,?) ON CONFLICT(owner_id,name) DO UPDATE SET unit=excluded.unit,price=excluded.price",(message.from_user.id,d['name'],d['unit'],p)); await state.clear(); await message.answer("✅ Цена работы сохранена.",reply_markup=main_kb())
    except: await message.answer("Введите цену числом, например 650.")


async def wp_view(callback):
    wid=int(callback.data.split(":")[1]); r=q("SELECT * FROM work_prices WHERE id=? AND owner_id=?",(wid,callback.from_user.id),True)
    if not r:return await callback.answer("Нет доступа",show_alert=True)
    b=InlineKeyboardBuilder(); b.button(text="🗑 Удалить",callback_data=f"delwp:{wid}"); b.button(text="⬅️ Цены",callback_data="work_prices"); b.adjust(1)
    await callback.message.edit_text(f"🧑‍🔧 <b>{r['name']}</b>\n\n{money(r['price'])} / {r['unit']}",reply_markup=b.as_markup()); await safe_answer(callback)


async def del_wp(callback):
    wid=int(callback.data.split(":")[1]); exec_sql("DELETE FROM work_prices WHERE id=? AND owner_id=?",(wid,callback.from_user.id)); await callback.message.edit_text("🗑 Цена удалена.",reply_markup=back_home_kb()); await safe_answer(callback)


async def add_work(callback,state):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    rows=q("SELECT * FROM work_prices WHERE owner_id=? ORDER BY name",(callback.from_user.id,)); b=InlineKeyboardBuilder()
    for r in rows:b.button(text=f"🧑‍🔧 {r['name']} — {money(r['price'])}/{r['unit']}",callback_data=f"usework:{oid}:{r['id']}")
    b.button(text="✍️ Ввести работу вручную",callback_data=f"manualwork:{oid}"); b.button(text="⬅️ Объект",callback_data=f"openobj:{oid}"); b.adjust(1)
    await callback.message.answer("Выберите работу из прайса или введите вручную:",reply_markup=b.as_markup()); await safe_answer(callback)


async def use_work(callback,state):
    _,oid,wid=callback.data.split(":"); oid=int(oid); wid=int(wid); r=q("SELECT * FROM work_prices WHERE id=? AND owner_id=?",(wid,callback.from_user.id),True)
    if not r or not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    await state.update_data(object_id=oid,name=r['name'],unit=r['unit'],price=r['price']); await state.set_state(WorkForm.qty); await callback.message.answer(f"{r['name']}\nЦена: {money(r['price'])}/{r['unit']}\n\nКоличество?"); await safe_answer(callback)


async def manual_work(callback,state):
    oid=int(callback.data.split(":")[1]); await state.update_data(object_id=oid); await state.set_state(WorkForm.name); await callback.message.answer("Название работы?"); await safe_answer(callback)


async def work_name(message,state): await state.update_data(name=message.text.strip()); await state.set_state(WorkForm.qty); await message.answer("Количество?")
async def work_qty(message,state):
    try:
        v=parse_float(message.text)
        if not valid_positive(v): raise ValueError
        await state.update_data(qty=v)
        d=await state.get_data()
        if d.get('unit') and d.get('price') is not None:
            await save_work(message,state,d)
        else:
            await state.set_state(WorkForm.unit); await message.answer("Единица? Например м², м.п., шт")
    except: await message.answer("Введите положительное число.")
async def work_unit(message,state): await state.update_data(unit=message.text.strip()); await state.set_state(WorkForm.price); await message.answer("Цена работы за единицу, ₽?")
async def work_price(message,state):
    try:
        p=parse_float(message.text)
        if p<0: raise ValueError
        d=await state.get_data(); d['price']=p; await save_work(message,state,d)
    except: await message.answer("Введите цену числом.")


async def save_work(message,state,d):
    total=d['qty']*d['price']; exec_sql("INSERT INTO estimate_items(object_id,kind,name,qty,unit,unit_price,total) VALUES(?,?,?,?,?,?,?)",(d['object_id'],'work',d['name'],d['qty'],d['unit'],d['price'],total)); await state.clear(); await message.answer(f"✅ Работа добавлена.\n\n{d['name']}\n{num(d['qty'])} {d['unit']} × {money(d['price'])}\n= <b>{money(total)}</b>",reply_markup=object_kb(d['object_id']))


# ---------- Materials ----------
async def materials(callback):
    rows=q("SELECT * FROM materials WHERE owner_id=? ORDER BY name",(callback.from_user.id,)); b=InlineKeyboardBuilder()
    for r in rows:b.button(text=f"🧱 {r['name']} — {money(r['pack_price'])}/{r['sale_unit']}",callback_data=f"matview:{r['id']}")
    b.button(text="➕ Добавить/изменить материал",callback_data="addmaterial"); b.button(text="⬅️ Назад",callback_data="home"); b.adjust(1)
    await callback.message.edit_text("🧱 <b>Материалы</b>\n\nХранится цена упаковки, размер упаковки и расход.",reply_markup=b.as_markup()); await safe_answer(callback)


async def add_material(callback,state): await state.set_state(MaterialCatalogForm.name); await callback.message.answer("Название материала?",reply_markup=cancel_kb()); await safe_answer(callback)
async def mat_name(message,state): await state.update_data(name=message.text.strip()); await state.set_state(MaterialCatalogForm.sale_unit); await message.answer("Единица продажи? Например мешок, банка, коробка, шт")
async def mat_unit(message,state): await state.update_data(sale_unit=message.text.strip()); await state.set_state(MaterialCatalogForm.pack_qty); await message.answer("Количество в одной упаковке? Например 30 кг. Введите только число: 30")
async def mat_pack_qty(message,state):
    try:
        v=parse_float(message.text)
        if not valid_positive(v): raise ValueError
        await state.update_data(pack_qty=v); await state.set_state(MaterialCatalogForm.pack_price); await message.answer("Цена одной упаковки, ₽?")
    except: await message.answer("Введите положительное число.")
async def mat_pack_price(message,state):
    try:
        v=parse_float(message.text)
        if v<0: raise ValueError
        await state.update_data(pack_price=v); await state.set_state(MaterialCatalogForm.consumption); await message.answer("Расход на 1 м²? Если не используется — 0")
    except: await message.answer("Введите цену числом.")
async def mat_consumption(message,state):
    try:
        v=parse_float(message.text)
        if v<0: raise ValueError
        await state.update_data(consumption=v); await state.set_state(MaterialCatalogForm.consumption_unit); await message.answer("Единица расхода? Например кг/м², л/м². Если не нужна — —")
    except: await message.answer("Введите число.")
async def mat_consumption_unit(message,state):
    d=await state.get_data(); cu="" if message.text.strip()=="—" else message.text.strip()
    exec_sql("INSERT INTO materials(owner_id,name,sale_unit,pack_qty,pack_price,consumption,consumption_unit) VALUES(?,?,?,?,?,?,?) ON CONFLICT(owner_id,name) DO UPDATE SET sale_unit=excluded.sale_unit,pack_qty=excluded.pack_qty,pack_price=excluded.pack_price,consumption=excluded.consumption,consumption_unit=excluded.consumption_unit",(message.from_user.id,d['name'],d['sale_unit'],d['pack_qty'],d['pack_price'],d['consumption'],cu)); await state.clear(); await message.answer("✅ Материал сохранён.",reply_markup=main_kb())


async def mat_view(callback):
    mid=int(callback.data.split(":")[1]); m=q("SELECT * FROM materials WHERE id=? AND owner_id=?",(mid,callback.from_user.id),True)
    if not m:return await callback.answer("Нет доступа",show_alert=True)
    b=InlineKeyboardBuilder(); b.button(text="🗑 Удалить",callback_data=f"delmat:{mid}"); b.button(text="⬅️ Материалы",callback_data="materials"); b.adjust(1)
    text=f"🧱 <b>{m['name']}</b>\n\nУпаковка: {num(m['pack_qty'])} {m['sale_unit']}\nЦена: {money(m['pack_price'])}\nРасход: {num(m['consumption'])} {m['consumption_unit'] or 'на м²'}"
    await callback.message.edit_text(text,reply_markup=b.as_markup()); await safe_answer(callback)


async def del_mat(callback):
    mid=int(callback.data.split(":")[1]); exec_sql("DELETE FROM materials WHERE id=? AND owner_id=?",(mid,callback.from_user.id)); await callback.message.edit_text("🗑 Материал удалён.",reply_markup=back_home_kb()); await safe_answer(callback)


async def add_material_to_object(callback,state):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    rows=q("SELECT * FROM materials WHERE owner_id=? ORDER BY name",(callback.from_user.id,)); b=InlineKeyboardBuilder()
    for r in rows:b.button(text=f"🧱 {r['name']} — {money(r['pack_price'])}",callback_data=f"usemat:{oid}:{r['id']}")
    b.button(text="➕ Создать материал",callback_data="addmaterial"); b.button(text="⬅️ Объект",callback_data=f"openobj:{oid}"); b.adjust(1)
    await callback.message.answer("Выберите материал:",reply_markup=b.as_markup()); await safe_answer(callback)


async def use_material(callback,state):
    _,oid,mid=callback.data.split(":"); oid=int(oid); mid=int(mid); m=q("SELECT * FROM materials WHERE id=? AND owner_id=?",(mid,callback.from_user.id),True)
    if not m or not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    rooms=q("SELECT id,name FROM rooms WHERE object_id=? ORDER BY id",(oid,)); b=InlineKeyboardBuilder()
    b.button(text="🏠 Все стены",callback_data=f"matarea:{oid}:{mid}:wall")
    b.button(text="⬆️ Все потолки",callback_data=f"matarea:{oid}:{mid}:ceiling")
    b.button(text="🪵 Все полы",callback_data=f"matarea:{oid}:{mid}:floor")
    for r in rooms:b.button(text=f"📐 {r['name']}",callback_data=f"matroom:{oid}:{mid}:{r['id']}")
    b.button(text="✍️ Своя площадь",callback_data=f"matcustom:{oid}:{mid}"); b.adjust(1)
    await callback.message.answer("На какую площадь считать материал?",reply_markup=b.as_markup()); await safe_answer(callback)


async def set_material_area(callback,state):
    _,oid,mid,typ=callback.data.split(":"); oid=int(oid); mid=int(mid)
    if not owner_object(callback.from_user.id,oid): return await callback.answer("Нет доступа",show_alert=True)
    rooms=q("SELECT id FROM rooms WHERE object_id=?",(oid,)); area=sum(room_metrics(r['id'])[typ] for r in rooms)
    await state.update_data(object_id=oid,material_id=mid,area=area); await state.set_state(MaterialForm.waste); await callback.message.answer(f"Площадь: <b>{area:.2f} м²</b>\n\nЗапас, %? Например 10",reply_markup=cancel_kb()); await safe_answer(callback)


async def set_material_room(callback,state):
    _,oid,mid,rid=callback.data.split(":"); oid=int(oid); mid=int(mid); rid=int(rid)
    r=owner_room(callback.from_user.id,rid)
    if not r or not owner_object(callback.from_user.id,oid): return await callback.answer("Нет доступа",show_alert=True)
    await state.update_data(object_id=oid,material_id=mid,area=room_metrics(rid)["wall"]); await state.set_state(MaterialForm.waste); await callback.message.answer(f"Комната: <b>{r['name']}</b>\nПлощадь стен: <b>{room_metrics(rid)['wall']:.2f} м²</b>\n\nЗапас, %? Например 10",reply_markup=cancel_kb()); await safe_answer(callback)


async def set_material_custom(callback,state):
    _,oid,mid=callback.data.split(":"); oid=int(oid); mid=int(mid)
    if not owner_object(callback.from_user.id,oid): return await callback.answer("Нет доступа",show_alert=True)
    await state.update_data(object_id=oid,material_id=mid); await state.set_state(MaterialForm.area); await callback.message.answer("Введите площадь в м².",reply_markup=cancel_kb()); await safe_answer(callback)


async def material_area_message(message,state):
    try:
        area=parse_float(message.text)
        if not valid_positive(area): raise ValueError
        await state.update_data(area=area); await state.set_state(MaterialForm.waste); await message.answer("Запас, %? Например 10")
    except: await message.answer("Введите положительную площадь.")


async def material_waste(message,state):
    try:
        waste=parse_float(message.text)
        if waste<0: raise ValueError
        d=await state.get_data(); m=q("SELECT * FROM materials WHERE id=? AND owner_id=?",(d['material_id'],message.from_user.id),True)
        if not m or not owner_object(message.from_user.id,d['object_id']): raise PermissionError
        area=d['area']; cons=m['consumption']; needed=area*cons if cons>0 else area; needed*=1+waste/100
        packs=math.ceil(needed/m['pack_qty']); total=packs*m['pack_price']
        note=f"Площадь {area:.2f} м²; расход {cons:g} {m['consumption_unit'] or 'ед./м²'}; запас {waste:g}%"
        exec_sql("INSERT INTO estimate_items(object_id,kind,name,qty,unit,unit_price,total,note) VALUES(?,?,?,?,?,?,?,?)",(d['object_id'],'material',m['name'],packs,m['sale_unit'],m['pack_price'],total,note))
        await state.clear(); await message.answer(f"🧱 <b>{m['name']}</b>\n\nПлощадь: {area:.2f} м²\nПотребность: {needed:.2f}\nУпаковок: <b>{packs}</b>\nЦена упаковки: {money(m['pack_price'])}\n💰 <b>{money(total)}</b>",reply_markup=object_kb(d['object_id']))
    except PermissionError: await state.clear(); await message.answer("Нет доступа к объекту или материалу.")
    except Exception as e: await send_error(message,e)


# ---------- Estimate / exports ----------
async def estimate(callback):
    oid=int(callback.data.split(":")[1]); obj=owner_object(callback.from_user.id,oid)
    if not obj:return await callback.answer("Нет доступа",show_alert=True)
    rows=q("SELECT * FROM estimate_items WHERE object_id=? ORDER BY id",(oid,)); works=[r for r in rows if r['kind']=='work']; mats=[r for r in rows if r['kind']=='material']; tw,tm=object_totals(oid)
    text=f"📋 <b>СМЕТА</b>\n\n🏠 {obj['name']}\n\n🧑‍🔧 <b>Работы</b>\n" + ("\n".join(f"• {r['name']} — {num(r['qty'])} {r['unit']} × {money(r['unit_price'])} = {money(r['total'])}" for r in works) or "—") + "\n\n🧱 <b>Материалы</b>\n" + ("\n".join(f"• {r['name']} — {num(r['qty'])} {r['unit']} × {money(r['unit_price'])} = {money(r['total'])}" for r in mats) or "—") + f"\n\n━━━━━━━━━━━━\nРаботы: <b>{money(tw)}</b>\nМатериалы: <b>{money(tm)}</b>\n💰 <b>ИТОГО: {money(tw+tm)}</b>"
    b=InlineKeyboardBuilder(); b.button(text="📊 Excel",callback_data=f"excel:{oid}"); b.button(text="📝 Word",callback_data=f"word:{oid}"); b.button(text="🗑 Очистить позиции",callback_data=f"clearitems:{oid}"); b.button(text="⬅️ Объект",callback_data=f"openobj:{oid}"); b.adjust(2,1,1)
    await callback.message.edit_text(text,reply_markup=b.as_markup()); await safe_answer(callback)


async def clear_items(callback):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid): return await callback.answer("Нет доступа",show_alert=True)
    exec_sql("DELETE FROM estimate_items WHERE object_id=?",(oid,)); await callback.message.edit_text("🗑 Позиции сметы очищены.",reply_markup=object_kb(oid)); await safe_answer(callback)


async def make_excel(oid,user_id):
    obj=owner_object(user_id,oid); rows=q("SELECT * FROM estimate_items WHERE object_id=? ORDER BY id",(oid,)); wb=Workbook(); ws=wb.active; ws.title="Смета"
    ws.append(["СМЕТА",obj['name']]); ws.append([]); ws.append(["Тип","Позиция","Количество","Ед.","Цена","Сумма","Комментарий"])
    for r in rows: ws.append(["Работа" if r['kind']=='work' else "Материал",r['name'],r['qty'],r['unit'],r['unit_price'],r['total'],r['note']])
    tw,tm=object_totals(oid); ws.append([]); ws.append(["Работы",tw]); ws.append(["Материалы",tm]); ws.append(["ИТОГО",tw+tm])
    ws["A1"].font=Font(bold=True,size=16); ws["A3"].font=Font(bold=True)
    for row in ws.iter_rows():
        for cell in row: cell.alignment=Alignment(vertical="top",wrap_text=True)
    for col in ws.columns:
        letter=col[0].column_letter; ws.column_dimensions[letter].width=min(max(len(str(c.value or "")) for c in col)+2,45)
    bio=io.BytesIO(); wb.save(bio); bio.seek(0); return bio


async def make_word(oid,user_id):
    obj=owner_object(user_id,oid); rows=q("SELECT * FROM estimate_items WHERE object_id=? ORDER BY id",(oid,)); doc=Document(); p=doc.add_paragraph(); rr=p.add_run("СМЕТА НА РЕМОНТ"); rr.bold=True; rr.font.size=Pt(18); doc.add_paragraph(f"Объект: {obj['name']}")
    if obj['address']: doc.add_paragraph(f"Адрес: {obj['address']}")
    if obj['customer']: doc.add_paragraph(f"Заказчик: {obj['customer']}")
    table=doc.add_table(rows=1,cols=6); table.style="Table Grid"
    for i,v in enumerate(["Тип","Позиция","Количество","Ед.","Цена","Сумма"]): table.rows[0].cells[i].text=v
    for r in rows:
        vals=["Работа" if r['kind']=='work' else "Материал",r['name'],num(r['qty']),r['unit'],money(r['unit_price']),money(r['total'])]
        cells=table.add_row().cells
        for i,v in enumerate(vals): cells[i].text=v
    tw,tm=object_totals(oid); doc.add_paragraph(f"Работы: {money(tw)}"); doc.add_paragraph(f"Материалы: {money(tm)}"); p=doc.add_paragraph(); rr=p.add_run(f"ИТОГО: {money(tw+tm)}"); rr.bold=True; rr.font.size=Pt(14)
    bio=io.BytesIO(); doc.save(bio); bio.seek(0); return bio


async def export_excel(callback):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    bio=await make_excel(oid,callback.from_user.id); await callback.message.answer_document(BufferedInputFile(bio.getvalue(),filename=f"smeta_{oid}.xlsx")); await safe_answer(callback,"Готово")


async def export_word(callback):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    bio=await make_word(oid,callback.from_user.id); await callback.message.answer_document(BufferedInputFile(bio.getvalue(),filename=f"smeta_{oid}.docx")); await safe_answer(callback,"Готово")


async def del_obj(callback):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    exec_sql("DELETE FROM objects WHERE id=? AND owner_id=?",(oid,callback.from_user.id)); await callback.message.edit_text("🗑 Объект удалён.",reply_markup=main_kb()); await safe_answer(callback)


# ---------- Quick calculator ----------
async def quick(callback,state):
    await state.set_state(QuickForm.name); await callback.message.answer("📐 <b>Быстрый расчёт</b>\n\nНазвание работы? Например: покраска стен",reply_markup=cancel_kb()); await safe_answer(callback)
async def quick_name(message,state): await state.update_data(name=message.text.strip()); await state.set_state(QuickForm.qty); await message.answer("Количество?")
async def quick_qty(message,state):
    try:
        v=parse_float(message.text)
        if not valid_positive(v): raise ValueError
        await state.update_data(qty=v); await state.set_state(QuickForm.unit); await message.answer("Единица? Например м²")
    except: await message.answer("Введите положительное число.")
async def quick_unit(message,state): await state.update_data(unit=message.text.strip()); await state.set_state(QuickForm.price); await message.answer("Цена за единицу, ₽?")
async def quick_price(message,state):
    try:
        p=parse_float(message.text); d=await state.get_data(); total=d['qty']*p; await state.clear(); await message.answer(f"📐 <b>Расчёт</b>\n\n{d['name']}\n{num(d['qty'])} {d['unit']} × {money(p)}\n\n💰 <b>{money(total)}</b>",reply_markup=main_kb())
    except: await message.answer("Введите цену числом.")


async def main():
    db_init(); bot=Bot(BOT_TOKEN,default=DefaultBotProperties(parse_mode=ParseMode.HTML)); dp=Dispatcher()
    dp.message.register(start,CommandStart()); dp.message.register(start,Command("menu")); dp.message.register(cancel_cmd,Command("cancel"))
    # callbacks
    dp.callback_query.register(cancel,F.data=="cancel"); dp.callback_query.register(home,F.data=="home"); dp.callback_query.register(new_object,F.data=="new_object"); dp.callback_query.register(objects,F.data=="objects"); dp.callback_query.register(open_obj,F.data.startswith("openobj:")); dp.callback_query.register(edit_object,F.data.startswith("editobj:")); dp.callback_query.register(edit_field,F.data.startswith("editfield:")); dp.callback_query.register(rooms,F.data.startswith("rooms:")); dp.callback_query.register(add_room,F.data.startswith("addroom:")); dp.callback_query.register(room_view,F.data.startswith("room:")); dp.callback_query.register(add_opening,F.data.startswith("addopening:")); dp.callback_query.register(opening_kind,F.data.startswith("openingkind:")); dp.callback_query.register(del_room,F.data.startswith("delroom:")); dp.callback_query.register(add_work,F.data.startswith("addwork:")); dp.callback_query.register(use_work,F.data.startswith("usework:")); dp.callback_query.register(manual_work,F.data.startswith("manualwork:")); dp.callback_query.register(work_prices,F.data=="work_prices"); dp.callback_query.register(add_work_price,F.data=="addworkprice"); dp.callback_query.register(wp_view,F.data.startswith("wpview:")); dp.callback_query.register(del_wp,F.data.startswith("delwp:")); dp.callback_query.register(materials,F.data=="materials"); dp.callback_query.register(add_material,F.data=="addmaterial"); dp.callback_query.register(mat_view,F.data.startswith("matview:")); dp.callback_query.register(del_mat,F.data.startswith("delmat:")); dp.callback_query.register(add_material_to_object,F.data.startswith("addmat:")); dp.callback_query.register(use_material,F.data.startswith("usemat:")); dp.callback_query.register(set_material_area,F.data.startswith("matarea:")); dp.callback_query.register(set_material_room,F.data.startswith("matroom:")); dp.callback_query.register(set_material_custom,F.data.startswith("matcustom:")); dp.callback_query.register(estimate,F.data.startswith("estimate:")); dp.callback_query.register(clear_items,F.data.startswith("clearitems:")); dp.callback_query.register(export_excel,F.data.startswith("excel:")); dp.callback_query.register(export_word,F.data.startswith("word:")); dp.callback_query.register(del_obj,F.data.startswith("delobj:")); dp.callback_query.register(quick,F.data=="quick")
    # messages / FSM
    dp.message.register(new_object_name,NewObject.name); dp.message.register(new_object_address,NewObject.address); dp.message.register(new_object_customer,NewObject.customer); dp.message.register(new_object_phone,NewObject.phone)
    dp.message.register(room_name,RoomForm.name); dp.message.register(room_length,RoomForm.length); dp.message.register(room_width,RoomForm.width); dp.message.register(room_height,RoomForm.height)
    dp.message.register(opening_width,OpeningForm.width); dp.message.register(opening_height,OpeningForm.height); dp.message.register(opening_qty,OpeningForm.qty)
    dp.message.register(work_name,WorkForm.name); dp.message.register(work_qty,WorkForm.qty); dp.message.register(work_unit,WorkForm.unit); dp.message.register(work_price,WorkForm.price)
    dp.message.register(mat_name,MaterialCatalogForm.name); dp.message.register(mat_unit,MaterialCatalogForm.sale_unit); dp.message.register(mat_pack_qty,MaterialCatalogForm.pack_qty); dp.message.register(mat_pack_price,MaterialCatalogForm.pack_price); dp.message.register(mat_consumption,MaterialCatalogForm.consumption); dp.message.register(mat_consumption_unit,MaterialCatalogForm.consumption_unit)
    dp.message.register(material_area_message,MaterialForm.area); dp.message.register(material_waste,MaterialForm.waste)
    dp.message.register(wp_name,WorkPriceForm.name); dp.message.register(wp_unit,WorkPriceForm.unit); dp.message.register(wp_price,WorkPriceForm.price)
    dp.message.register(edit_field_value,EditObjectForm.value)
    dp.message.register(quick_name,QuickForm.name); dp.message.register(quick_qty,QuickForm.qty); dp.message.register(quick_unit,QuickForm.unit); dp.message.register(quick_price,QuickForm.price)
    await dp.start_polling(bot)


if __name__ == "__main__": asyncio.run(main())
