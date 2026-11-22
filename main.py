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
from aiogram.types import BufferedInputFile, CallbackQuery, InlineKeyboardButton, Message
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
    conn.executescript("""
    PRAGMA foreign_keys=ON;
    CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, username TEXT, first_name TEXT);
    CREATE TABLE IF NOT EXISTS objects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        owner_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        address TEXT DEFAULT '',
        customer TEXT DEFAULT '',
        phone TEXT DEFAULT '',
        created_at TEXT NOT NULL,
        FOREIGN KEY(owner_id) REFERENCES users(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS rooms (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        object_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        length REAL NOT NULL,
        width REAL NOT NULL,
        height REAL NOT NULL,
        FOREIGN KEY(object_id) REFERENCES objects(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS openings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        room_id INTEGER NOT NULL,
        kind TEXT NOT NULL,
        width REAL NOT NULL,
        height REAL NOT NULL,
        qty INTEGER NOT NULL DEFAULT 1,
        FOREIGN KEY(room_id) REFERENCES rooms(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS work_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        owner_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        unit TEXT NOT NULL,
        price REAL NOT NULL,
        FOREIGN KEY(owner_id) REFERENCES users(id) ON DELETE CASCADE,
        UNIQUE(owner_id, name)
    );
    CREATE TABLE IF NOT EXISTS materials (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        owner_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        unit TEXT NOT NULL,
        pack_qty REAL NOT NULL,
        pack_price REAL NOT NULL,
        consumption REAL DEFAULT 0,
        consumption_unit TEXT DEFAULT '',
        FOREIGN KEY(owner_id) REFERENCES users(id) ON DELETE CASCADE,
        UNIQUE(owner_id, name)
    );
    CREATE TABLE IF NOT EXISTS estimate_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        object_id INTEGER NOT NULL,
        kind TEXT NOT NULL CHECK(kind IN ('work','material')),
        name TEXT NOT NULL,
        qty REAL NOT NULL,
        unit TEXT NOT NULL,
        unit_price REAL NOT NULL,
        total REAL NOT NULL,
        note TEXT DEFAULT '',
        FOREIGN KEY(object_id) REFERENCES objects(id) ON DELETE CASCADE
    );
    """)
    conn.commit()


def q(sql, args=(), one=False):
    cur = conn.execute(sql, args)
    rows = cur.fetchone() if one else cur.fetchall()
    return rows


def exec_sql(sql, args=()):
    cur = conn.execute(sql, args)
    conn.commit()
    return cur.lastrowid


def money(v):
    return f"{v:,.0f} ₽".replace(",", " ")


def user_upsert(message: Message):
    exec_sql("INSERT INTO users(id, username, first_name) VALUES(?,?,?) ON CONFLICT(id) DO UPDATE SET username=excluded.username, first_name=excluded.first_name",
             (message.from_user.id, message.from_user.username or "", message.from_user.first_name or ""))


def owner_object(user_id, object_id):
    return q("SELECT * FROM objects WHERE id=? AND owner_id=?", (object_id, user_id), True)


def room_metrics(room_id):
    r = q("SELECT * FROM rooms WHERE id=?", (room_id,), True)
    if not r:
        return None
    wall = 2 * (r["length"] + r["width"]) * r["height"]
    floor = r["length"] * r["width"]
    ceiling = floor
    openings = q("SELECT width,height,qty FROM openings WHERE room_id=?", (room_id,))
    deducted = sum(x["width"] * x["height"] * x["qty"] for x in openings)
    return {"wall": max(0, wall - deducted), "floor": floor, "ceiling": ceiling, "raw_wall": wall, "openings": deducted}


class NewObject(StatesGroup):
    name = State(); address = State(); customer = State(); phone = State()

class RoomForm(StatesGroup):
    object_id = State(); name = State(); length = State(); width = State(); height = State()

class WorkForm(StatesGroup):
    object_id = State(); name = State(); qty = State(); unit = State(); price = State()

class MaterialForm(StatesGroup):
    object_id = State(); material_id = State(); area_type = State(); thickness = State(); waste = State(); override_price = State()

class PriceWorkForm(StatesGroup):
    name = State(); unit = State(); price = State()

class MaterialCatalogForm(StatesGroup):
    name = State(); unit = State(); pack_qty = State(); pack_price = State(); consumption = State(); consumption_unit = State()


def main_kb():
    b = InlineKeyboardBuilder()
    b.button(text="➕ Новый объект", callback_data="new_object")
    b.button(text="📂 Мои объекты", callback_data="objects")
    b.button(text="📐 Быстрый расчёт", callback_data="quick")
    b.button(text="💰 Цены на работы", callback_data="work_prices")
    b.button(text="🧱 Материалы", callback_data="materials")
    b.adjust(1, 1, 2, 1)
    return b.as_markup()


def object_kb(obj_id):
    b = InlineKeyboardBuilder()
    b.button(text="📐 Комнаты", callback_data=f"rooms:{obj_id}")
    b.button(text="➕ Добавить работу", callback_data=f"addwork:{obj_id}")
    b.button(text="🧱 Добавить материал", callback_data=f"addmat:{obj_id}")
    b.button(text="📋 Смета", callback_data=f"estimate:{obj_id}")
    b.button(text="📊 Excel", callback_data=f"excel:{obj_id}")
    b.button(text="📝 Word", callback_data=f"word:{obj_id}")
    b.button(text="🗑 Удалить объект", callback_data=f"delobj:{obj_id}")
    b.button(text="⬅️ Главное меню", callback_data="home")
    b.adjust(2, 2, 2, 1, 1)
    return b.as_markup()


def parse_float(text):
    return float(text.replace(" ", "").replace(",", "."))


async def start(message: Message, state: FSMContext):
    await state.clear(); user_upsert(message)
    await message.answer("🏗 <b>СтройСмета</b>\n\nРасчёт работ и материалов по объектам.\n\nВыберите действие:", reply_markup=main_kb())


async def home(callback: CallbackQuery, state: FSMContext):
    await state.clear(); await callback.message.edit_text("🏗 <b>СтройСмета</b>\n\nВыберите действие:", reply_markup=main_kb()); await callback.answer()


async def new_object(callback: CallbackQuery, state: FSMContext):
    await state.set_state(NewObject.name); await callback.message.answer("🏠 Название объекта?"); await callback.answer()


async def new_object_name(message: Message, state: FSMContext):
    await state.update_data(name=message.text.strip()); await state.set_state(NewObject.address); await message.answer("📍 Адрес?\nМожно написать <b>—</b>, чтобы пропустить.")

async def new_object_address(message: Message, state: FSMContext):
    await state.update_data(address="" if message.text.strip()=="—" else message.text.strip()); await state.set_state(NewObject.customer); await message.answer("👤 Заказчик?\nМожно <b>—</b>.")

async def new_object_customer(message: Message, state: FSMContext):
    await state.update_data(customer="" if message.text.strip()=="—" else message.text.strip()); await state.set_state(NewObject.phone); await message.answer("📞 Телефон?\nМожно <b>—</b>.")

async def new_object_phone(message: Message, state: FSMContext):
    d=await state.get_data(); phone="" if message.text.strip()=="—" else message.text.strip()
    oid=exec_sql("INSERT INTO objects(owner_id,name,address,customer,phone,created_at) VALUES(?,?,?,?,?,?)",(message.from_user.id,d["name"],d["address"],d["customer"],phone,datetime.now().isoformat()))
    await state.clear(); obj=owner_object(message.from_user.id,oid)
    await message.answer(f"✅ Объект создан: <b>{obj['name']}</b>",reply_markup=object_kb(oid))

async def objects(callback: CallbackQuery):
    rows=q("SELECT * FROM objects WHERE owner_id=? ORDER BY id DESC",(callback.from_user.id,))
    b=InlineKeyboardBuilder()
    for r in rows: b.button(text=f"🏠 {r['name']}",callback_data=f"openobj:{r['id']}")
    b.button(text="➕ Новый объект",callback_data="new_object"); b.button(text="⬅️ Назад",callback_data="home"); b.adjust(1)
    await callback.message.edit_text("📂 <b>Мои объекты</b>\n\n" + ("Выберите объект:" if rows else "Пока нет объектов."),reply_markup=b.as_markup()); await callback.answer()

async def open_obj(callback: CallbackQuery):
    obj=owner_object(callback.from_user.id,int(callback.data.split(":")[1]))
    if not obj: return await callback.answer("Нет доступа",show_alert=True)
    await callback.message.edit_text(f"🏠 <b>{obj['name']}</b>\n\n📍 {obj['address'] or 'Адрес не указан'}\n👤 {obj['customer'] or 'Заказчик не указан'}",reply_markup=object_kb(obj['id'])); await callback.answer()

async def rooms(callback: CallbackQuery):
    oid=int(callback.data.split(":")[1]); obj=owner_object(callback.from_user.id,oid)
    if not obj:return await callback.answer("Нет доступа",show_alert=True)
    rows=q("SELECT * FROM rooms WHERE object_id=? ORDER BY id",(oid,)); b=InlineKeyboardBuilder()
    for r in rows:
        m=room_metrics(r['id']); b.button(text=f"📐 {r['name']} — {m['floor']:.1f} м²",callback_data=f"room:{r['id']}")
    b.button(text="➕ Добавить комнату",callback_data=f"addroom:{oid}"); b.button(text="⬅️ Объект",callback_data=f"openobj:{oid}"); b.adjust(1)
    await callback.message.edit_text("📐 <b>Комнаты</b>",reply_markup=b.as_markup()); await callback.answer()

async def add_room(callback: CallbackQuery,state:FSMContext):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid): return await callback.answer("Нет доступа",show_alert=True)
    await state.update_data(object_id=oid); await state.set_state(RoomForm.name); await callback.message.answer("Название комнаты? Например: <b>Гостиная</b>"); await callback.answer()

async def room_name(message:Message,state:FSMContext): await state.update_data(name=message.text.strip()); await state.set_state(RoomForm.length); await message.answer("📏 Длина, м?")
async def room_length(message:Message,state:FSMContext):
    try: v=parse_float(message.text); await state.update_data(length=v); await state.set_state(RoomForm.width); await message.answer("📏 Ширина, м?")
    except: await message.answer("Введите число, например 6 или 5.5")
async def room_width(message:Message,state:FSMContext):
    try: v=parse_float(message.text); await state.update_data(width=v); await state.set_state(RoomForm.height); await message.answer("📏 Высота, м?")
    except: await message.answer("Введите число, например 2.7")
async def room_height(message:Message,state:FSMContext):
    try:
        v=parse_float(message.text); d=await state.get_data(); rid=exec_sql("INSERT INTO rooms(object_id,name,length,width,height) VALUES(?,?,?,?,?)",(d['object_id'],d['name'],d['length'],d['width'],v)); await state.clear(); m=room_metrics(rid); await message.answer(f"✅ <b>{d['name']}</b> добавлена.\n\nПол: {m['floor']:.2f} м²\nПотолок: {m['ceiling']:.2f} м²\nСтены до вычета проёмов: {m['raw_wall']:.2f} м²",reply_markup=object_kb(d['object_id']))
    except: await message.answer("Введите число, например 2.7")

async def room_view(callback:CallbackQuery):
    rid=int(callback.data.split(":")[1]); r=q("SELECT r.*,o.owner_id FROM rooms r JOIN objects o ON o.id=r.object_id WHERE r.id=?",(rid,),True)
    if not r or r['owner_id']!=callback.from_user.id:return await callback.answer("Нет доступа",show_alert=True)
    m=room_metrics(rid); await callback.message.edit_text(f"📐 <b>{r['name']}</b>\n\nПол: {m['floor']:.2f} м²\nПотолок: {m['ceiling']:.2f} м²\nСтены: {m['wall']:.2f} м²\nПроёмы: {m['openings']:.2f} м²",reply_markup=InlineKeyboardBuilder().button(text="⬅️ Комнаты",callback_data=f"rooms:{r['object_id']}").as_markup()); await callback.answer()

async def add_work(callback:CallbackQuery,state:FSMContext):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    await state.update_data(object_id=oid); await state.set_state(WorkForm.name); await callback.message.answer("🧑‍🔧 Название работы? Например: <b>Штукатурка стен</b>")
    await callback.answer()
async def work_name(message:Message,state:FSMContext): await state.update_data(name=message.text.strip()); await state.set_state(WorkForm.qty); await message.answer("Количество?")
async def work_qty(message:Message,state:FSMContext):
    try: await state.update_data(qty=parse_float(message.text)); await state.set_state(WorkForm.unit); await message.answer("Единица? Например: м², м.п., шт")
    except: await message.answer("Введите число.")
async def work_unit(message:Message,state:FSMContext): await state.update_data(unit=message.text.strip()); await state.set_state(WorkForm.price); await message.answer("Цена работы за единицу, ₽?")
async def work_price(message:Message,state:FSMContext):
    try:
        d=await state.get_data(); p=parse_float(message.text); total=d['qty']*p; exec_sql("INSERT INTO estimate_items(object_id,kind,name,qty,unit,unit_price,total) VALUES(?,?,?,?,?,?,?)",(d['object_id'],'work',d['name'],d['qty'],d['unit'],p,total)); await state.clear(); await message.answer(f"✅ Работа добавлена.\n\n{d['name']}\n{d['qty']} {d['unit']} × {money(p)}\n= <b>{money(total)}</b>",reply_markup=object_kb(d['object_id']))
    except: await message.answer("Введите цену числом.")

async def materials(callback:CallbackQuery):
    rows=q("SELECT * FROM materials WHERE owner_id=? ORDER BY name",(callback.from_user.id,)); b=InlineKeyboardBuilder()
    for r in rows:b.button(text=f"🧱 {r['name']} — {money(r['pack_price'])}",callback_data=f"matview:{r['id']}")
    b.button(text="➕ Добавить материал",callback_data="addmaterial");b.button(text="⬅️ Назад",callback_data="home");b.adjust(1)
    await callback.message.edit_text("🧱 <b>Материалы</b>\n\nЦена хранится отдельно от работ. Можно менять её в любой момент.",reply_markup=b.as_markup());await callback.answer()

async def add_material(callback:CallbackQuery,state:FSMContext):
    await state.set_state(MaterialCatalogForm.name); await callback.message.answer("Название материала?"); await callback.answer()
async def mat_name(message:Message,state:FSMContext): await state.update_data(name=message.text.strip());await state.set_state(MaterialCatalogForm.unit);await message.answer("Единица продажи? Например: мешок, л, м², шт")
async def mat_unit(message:Message,state:FSMContext): await state.update_data(unit=message.text.strip());await state.set_state(MaterialCatalogForm.pack_qty);await message.answer("Количество в упаковке? Например 30 для мешка 30 кг. Для шт поставьте 1.")
async def mat_pack_qty(message:Message,state:FSMContext):
    try: await state.update_data(pack_qty=parse_float(message.text));await state.set_state(MaterialCatalogForm.pack_price);await message.answer("Цена одной упаковки, ₽?")
    except: await message.answer("Введите число.")
async def mat_pack_price(message:Message,state:FSMContext):
    try: await state.update_data(pack_price=parse_float(message.text));await state.set_state(MaterialCatalogForm.consumption);await message.answer("Расход на 1 м²? Если не нужен — 0.")
    except: await message.answer("Введите число.")
async def mat_consumption(message:Message,state:FSMContext):
    try: await state.update_data(consumption=parse_float(message.text));await state.set_state(MaterialCatalogForm.consumption_unit);await message.answer("Единица расхода (кг/м², л/м² и т.п.)? Можно —")
    except: await message.answer("Введите число.")
async def mat_consumption_unit(message:Message,state:FSMContext):
    d=await state.get_data(); cu="" if message.text.strip()=="—" else message.text.strip(); exec_sql("INSERT INTO materials(owner_id,name,unit,pack_qty,pack_price,consumption,consumption_unit) VALUES(?,?,?,?,?,?,?) ON CONFLICT(owner_id,name) DO UPDATE SET unit=excluded.unit,pack_qty=excluded.pack_qty,pack_price=excluded.pack_price,consumption=excluded.consumption,consumption_unit=excluded.consumption_unit",(message.from_user.id,d['name'],d['unit'],d['pack_qty'],d['pack_price'],d['consumption'],cu));await state.clear();await message.answer("✅ Материал сохранён.",reply_markup=main_kb())

async def add_material_to_object(callback:CallbackQuery,state:FSMContext):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    rows=q("SELECT * FROM materials WHERE owner_id=? ORDER BY name",(callback.from_user.id,)); b=InlineKeyboardBuilder()
    for r in rows:b.button(text=f"🧱 {r['name']}",callback_data=f"usemat:{oid}:{r['id']}")
    b.button(text="➕ Создать материал",callback_data="addmaterial");b.adjust(1);await callback.message.answer("Выберите материал:",reply_markup=b.as_markup());await callback.answer()

async def use_material(callback:CallbackQuery,state:FSMContext):
    _,oid,mid=callback.data.split(":");oid=int(oid);mid=int(mid);m=q("SELECT * FROM materials WHERE id=? AND owner_id=?",(mid,callback.from_user.id),True)
    if not owner_object(callback.from_user.id,oid) or not m:return await callback.answer("Нет доступа",show_alert=True)
    await state.update_data(object_id=oid,material_id=mid);await state.set_state(MaterialForm.area_type)
    b=InlineKeyboardBuilder();b.button(text="🧱 Все стены",callback_data="area:wall");b.button(text="⬆️ Все потолки",callback_data="area:ceiling");b.button(text="🪵 Все полы",callback_data="area:floor");b.button(text="✍️ Ввести площадь",callback_data="area:custom");b.adjust(1)
    await callback.message.answer("На какую площадь считать материал?",reply_markup=b.as_markup());await callback.answer()

async def material_area(callback:CallbackQuery,state:FSMContext):
    d=await state.get_data();oid=d['object_id']; typ=callback.data.split(":")[1]
    if typ=="custom": await state.set_state(MaterialForm.thickness);await state.update_data(area_custom=True);await callback.message.answer("Введите площадь в м².")
    else:
        rooms=q("SELECT id FROM rooms WHERE object_id=?",(oid,)); area=sum(room_metrics(r['id'])[typ] for r in rooms);await state.update_data(area=area,area_custom=False);await state.set_state(MaterialForm.thickness);await callback.message.answer(f"Площадь: <b>{area:.2f} м²</b>\n\nВведите толщину/слой, мм. Если расход уже задан как готовый на м² — напишите 0.")
    await callback.answer()

async def material_thickness(message:Message,state:FSMContext):
    try:
        v=parse_float(message.text);d=await state.get_data()
        if d.get('area_custom'):d['area']=v
        d['thickness']=v;await state.update_data(**d);await state.set_state(MaterialForm.waste);await message.answer("Запас, %? Например 10")
    except:await message.answer("Введите число.")

async def material_waste(message:Message,state:FSMContext):
    try:
        waste=parse_float(message.text);d=await state.get_data();m=q("SELECT * FROM materials WHERE id=? AND owner_id=?",(d['material_id'],message.from_user.id),True);area=d['area'];cons=m['consumption'];
        # consumption is interpreted as amount per m²; thickness is informational for catalog items
        needed=area*cons if cons>0 else area
        needed*=1+waste/100
        packs=math.ceil(needed/m['pack_qty']) if m['pack_qty']>0 else 0
        total=packs*m['pack_price']
        exec_sql("INSERT INTO estimate_items(object_id,kind,name,qty,unit,unit_price,total,note) VALUES(?,?,?,?,?,?,?,?)",(d['object_id'],'material',m['name'],packs,'упак.',m['pack_price'],total,f"Площадь {area:.2f} м²; расход {cons} {m['consumption_unit'] or m['unit']}/м²; запас {waste}%"))
        await state.clear();await message.answer(f"🧱 <b>{m['name']}</b>\n\nПлощадь: {area:.2f} м²\nРасчётная потребность: {needed:.2f} {m['unit']}\nУпаковок: {packs}\nЦена: {money(m['pack_price'])}\n\n💰 <b>{money(total)}</b>",reply_markup=object_kb(d['object_id']))
    except Exception as e: await message.answer(f"Ошибка расчёта: {e}")

async def estimate(callback:CallbackQuery):
    oid=int(callback.data.split(":")[1]);obj=owner_object(callback.from_user.id,oid)
    if not obj:return await callback.answer("Нет доступа",show_alert=True)
    rows=q("SELECT * FROM estimate_items WHERE object_id=? ORDER BY id",(oid,));works=[r for r in rows if r['kind']=='work'];mats=[r for r in rows if r['kind']=='material'];tw=sum(r['total'] for r in works);tm=sum(r['total'] for r in mats)
    text=f"📋 <b>СМЕТА</b>\n\n🏠 {obj['name']}\n\n🧑‍🔧 <b>Работы:</b>\n"+("\n".join(f"• {r['name']} — {r['qty']:g} {r['unit']} × {money(r['unit_price'])} = {money(r['total'])}" for r in works) or "—")+f"\n\n🧱 <b>Материалы:</b>\n"+("\n".join(f"• {r['name']} — {r['qty']:g} {r['unit']} × {money(r['unit_price'])} = {money(r['total'])}" for r in mats) or "—")+f"\n\n━━━━━━━━━━━━\nРаботы: <b>{money(tw)}</b>\nМатериалы: <b>{money(tm)}</b>\n💰 <b>ИТОГО: {money(tw+tm)}</b>"
    b=InlineKeyboardBuilder();b.button(text="📊 Excel",callback_data=f"excel:{oid}");b.button(text="📝 Word",callback_data=f"word:{oid}");b.button(text="⬅️ Объект",callback_data=f"openobj:{oid}");b.adjust(2,1)
    await callback.message.edit_text(text,reply_markup=b.as_markup());await callback.answer()

async def make_excel(oid,user_id):
    obj=owner_object(user_id,oid);rows=q("SELECT * FROM estimate_items WHERE object_id=? ORDER BY id",(oid,));wb=Workbook();ws=wb.active;ws.title="Смета";ws.append(["СМЕТА",obj['name']]);ws.append([]);ws.append(["Тип","Позиция","Количество","Ед.","Цена","Сумма","Комментарий"])
    for r in rows:ws.append(["Работа" if r['kind']=='work' else "Материал",r['name'],r['qty'],r['unit'],r['unit_price'],r['total'],r['note']])
    works=sum(r['total'] for r in rows if r['kind']=='work');mats=sum(r['total'] for r in rows if r['kind']=='material');ws.append([]);ws.append(["Работы",works]);ws.append(["Материалы",mats]);ws.append(["ИТОГО",works+mats]);
    for cell in ws[1]:cell.font=Font(bold=True,size=16)
    for cell in ws[3]:cell.font=Font(bold=True)
    for col in ws.columns: ws.column_dimensions[col[0].column_letter].width=min(max(len(str(c.value or "")) for c in col)+2,40)
    bio=io.BytesIO();wb.save(bio);bio.seek(0);return bio

async def make_word(oid,user_id):
    obj=owner_object(user_id,oid);rows=q("SELECT * FROM estimate_items WHERE object_id=? ORDER BY id",(oid,));doc=Document();p=doc.add_paragraph();run=p.add_run("СМЕТА НА РЕМОНТ");run.bold=True;run.font.size=Pt(18);doc.add_paragraph(f"Объект: {obj['name']}");
    if obj['address']:doc.add_paragraph(f"Адрес: {obj['address']}")
    table=doc.add_table(rows=1,cols=6);hdr=table.rows[0].cells
    for i,v in enumerate(["Тип","Позиция","Количество","Ед.","Цена","Сумма"]):hdr[i].text=v
    for r in rows:
        cells=table.add_row().cells
        vals=["Работа" if r['kind']=='work' else "Материал",r['name'],f"{r['qty']:g}",r['unit'],money(r['unit_price']),money(r['total'])]
        for i,v in enumerate(vals):cells[i].text=v
    works=sum(r['total'] for r in rows if r['kind']=='work');mats=sum(r['total'] for r in rows if r['kind']=='material');doc.add_paragraph(f"Работы: {money(works)}");doc.add_paragraph(f"Материалы: {money(mats)}");p=doc.add_paragraph();rr=p.add_run(f"ИТОГО: {money(works+mats)}");rr.bold=True;rr.font.size=Pt(14)
    bio=io.BytesIO();doc.save(bio);bio.seek(0);return bio

async def export_excel(callback:CallbackQuery):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    bio=await make_excel(oid,callback.from_user.id);await callback.message.answer_document(BufferedInputFile(bio.getvalue(),filename="smeta.xlsx"));await callback.answer()

async def export_word(callback:CallbackQuery):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    bio=await make_word(oid,callback.from_user.id);await callback.message.answer_document(BufferedInputFile(bio.getvalue(),filename="smeta.docx"));await callback.answer()

async def del_obj(callback:CallbackQuery):
    oid=int(callback.data.split(":")[1]);
    if not owner_object(callback.from_user.id,oid):return await callback.answer("Нет доступа",show_alert=True)
    exec_sql("DELETE FROM objects WHERE id=? AND owner_id=?",(oid,callback.from_user.id));await callback.message.edit_text("🗑 Объект удалён.",reply_markup=main_kb());await callback.answer()

async def quick(callback:CallbackQuery):
    await callback.message.answer("📐 Быстрый расчёт\n\nДля быстрого расчёта сейчас используйте создание объекта. В следующей версии сюда можно добавить отдельные калькуляторы: стены, пол, потолок, плитка, краска и т.д.");await callback.answer()

async def work_prices(callback:CallbackQuery):
    rows=q("SELECT * FROM work_prices WHERE owner_id=? ORDER BY name",(callback.from_user.id,));b=InlineKeyboardBuilder()
    for r in rows:b.button(text=f"🧑‍🔧 {r['name']} — {money(r['price'])}/{r['unit']}",callback_data=f"wp:{r['id']}")
    b.button(text="➕ Добавить цену",callback_data="addworkprice");b.button(text="⬅️ Назад",callback_data="home");b.adjust(1)
    await callback.message.edit_text("💰 <b>Цены на работы</b>",reply_markup=b.as_markup());await callback.answer()

async def add_work_price(callback:CallbackQuery,state:FSMContext): await state.set_state(PriceWorkForm.name);await callback.message.answer("Название работы?");await callback.answer()
async def wp_name(message:Message,state:FSMContext):await state.update_data(name=message.text.strip());await state.set_state(PriceWorkForm.unit);await message.answer("Единица?")
async def wp_unit(message:Message,state:FSMContext):await state.update_data(unit=message.text.strip());await state.set_state(PriceWorkForm.price);await message.answer("Цена, ₽?")
async def wp_price(message:Message,state:FSMContext):
    try:
        d=await state.get_data();p=parse_float(message.text);exec_sql("INSERT INTO work_prices(owner_id,name,unit,price) VALUES(?,?,?,?) ON CONFLICT(owner_id,name) DO UPDATE SET unit=excluded.unit,price=excluded.price",(message.from_user.id,d['name'],d['unit'],p));await state.clear();await message.answer("✅ Цена работы сохранена.",reply_markup=main_kb())
    except:await message.answer("Введите цену числом.")

async def main():
    db_init();bot=Bot(BOT_TOKEN,default=DefaultBotProperties(parse_mode=ParseMode.HTML));dp=Dispatcher()
    dp.message.register(start,CommandStart());dp.message.register(start,Command("menu"))
    dp.callback_query.register(home,F.data=="home");dp.callback_query.register(new_object,F.data=="new_object");dp.callback_query.register(objects,F.data=="objects");dp.callback_query.register(open_obj,F.data.startswith("openobj:"));dp.callback_query.register(rooms,F.data.startswith("rooms:"));dp.callback_query.register(add_room,F.data.startswith("addroom:"));dp.callback_query.register(room_view,F.data.startswith("room:"));dp.callback_query.register(add_work,F.data.startswith("addwork:"));dp.callback_query.register(add_material_to_object,F.data.startswith("addmat:"));dp.callback_query.register(materials,F.data=="materials");dp.callback_query.register(add_material,F.data=="addmaterial");dp.callback_query.register(use_material,F.data.startswith("usemat:"));dp.callback_query.register(material_area,F.data.startswith("area:"));dp.callback_query.register(estimate,F.data.startswith("estimate:"));dp.callback_query.register(export_excel,F.data.startswith("excel:"));dp.callback_query.register(export_word,F.data.startswith("word:"));dp.callback_query.register(del_obj,F.data.startswith("delobj:"));dp.callback_query.register(quick,F.data=="quick");dp.callback_query.register(work_prices,F.data=="work_prices");dp.callback_query.register(add_work_price,F.data=="addworkprice")
    dp.message.register(new_object_name,NewObject.name);dp.message.register(new_object_address,NewObject.address);dp.message.register(new_object_customer,NewObject.customer);dp.message.register(new_object_phone,NewObject.phone)
    dp.message.register(room_name,RoomForm.name);dp.message.register(room_length,RoomForm.length);dp.message.register(room_width,RoomForm.width);dp.message.register(room_height,RoomForm.height)
    dp.message.register(work_name,WorkForm.name);dp.message.register(work_qty,WorkForm.qty);dp.message.register(work_unit,WorkForm.unit);dp.message.register(work_price,WorkForm.price)
    dp.message.register(mat_name,MaterialCatalogForm.name);dp.message.register(mat_unit,MaterialCatalogForm.unit);dp.message.register(mat_pack_qty,MaterialCatalogForm.pack_qty);dp.message.register(mat_pack_price,MaterialCatalogForm.pack_price);dp.message.register(mat_consumption,MaterialCatalogForm.consumption);dp.message.register(mat_consumption_unit,MaterialCatalogForm.consumption_unit)
    dp.message.register(material_thickness,MaterialForm.thickness);dp.message.register(material_waste,MaterialForm.waste)
    dp.message.register(wp_name,PriceWorkForm.name);dp.message.register(wp_unit,PriceWorkForm.unit);dp.message.register(wp_price,PriceWorkForm.price)
    await dp.start_polling(bot)

if __name__=="__main__": asyncio.run(main())
