import asyncio
import re
import os
from aiogram import Bot, Dispatcher, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ContentType, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.utils import executor
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.contrib.fsm_storage.memory import MemoryStorage
import gspread
from oauth2client.service_account import ServiceAccountCredentials

API_TOKEN = '7530739654:AAETWPKYaMWI21BnvDj6f0T70XXZfEWLDVI'

# Google Sheets
scope = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

spreadsheet_sources = {
    '1IZhZ0nbDFmbH-dbS91gmZmkZdw7t9PnltE1DOwvvN_0': 'Решенные инциденты',
    '1sBvdvXXoNERLcLYbwQAw_onKRu5oTsVjf8jIWm64CEM': 'Передача проблемных товаров',
    '1g0o2D13UHmn4w_HIqIgSRtM-BFO6iMuujijffFY4ux4': 'Новая таблица передачи товаров',
    '1Wda6N00YgoHTX88MHU0tiCfjMel5hR-pK1zvqjnArY4': 'VP OPV таблица',
    '1PgdQLqdg1zbc7dTaXudQBlwBJR9r4htKKYCn3FstbJ8': 'Отчет OPV 2025',
    '1CJZeLK-Q_oDVaFinJzdWycymF3IWZkX4QqRmGoJM_m8': 'ТСТ таблица'
}

EMPLOYEES_SHEET_ID = '1QGXxe3TYXpFEMcglbaHSJrF55EYdZOyxY2ScNhmiQu4'
REFRESH_INTERVAL_MINUTES = 3

bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)

sheets = []
employees = []

if not os.path.exists('employees'):
    os.makedirs('employees')

main_keyboard = ReplyKeyboardMarkup(resize_keyboard=True)
main_keyboard.add(
    KeyboardButton("🔄 Обновить таблицы"),
    KeyboardButton("➕ Добавить сотрудницу")
)

class EmployeeStates(StatesGroup):
    waiting_for_photo = State()
    waiting_for_fio = State()
    waiting_for_confirm = State()

def normalize(text):
    if not isinstance(text, str):
        return ''
    return re.sub(r'\W+', '', text).lower()

def contains_cyrillic(text):
    return bool(re.search(r'[А-Яа-яЁё]', text))

async def load_employees_from_sheet():
    global employees
    employees = []
    try:
        sheet = client.open_by_key(EMPLOYEES_SHEET_ID).sheet1
        rows = sheet.get_all_values()
        for row in rows[1:]:
            if len(row) >= 2:
                fio = row[0].strip()
                photo_filename = row[1].strip()
                photo_path = os.path.join('employees', photo_filename)
                if fio and os.path.exists(photo_path):
                    employees.append({"fio": fio, "photo_path": photo_path})
    except Exception as e:
        print("Ошибка загрузки сотрудниц:", e)

async def save_employee_to_sheet(fio, photo_filename):
    try:
        sheet = client.open_by_key(EMPLOYEES_SHEET_ID).sheet1
        sheet.append_row([fio, photo_filename])
    except Exception as e:
        print("Ошибка сохранения сотрудницы:", e)

async def refresh_sheets_once():
    global sheets
    new_sheets = []
    for sheet_id, name in spreadsheet_sources.items():
        try:
            sheet = client.open_by_key(sheet_id).sheet1
            new_sheets.append((sheet, name))
        except Exception as e:
            print(f"Ошибка загрузки таблицы {name}: {e}")
    sheets = new_sheets

async def refresh_sheets_loop():
    while True:
        await refresh_sheets_once()
        await load_employees_from_sheet()
        await asyncio.sleep(REFRESH_INTERVAL_MINUTES * 60)

@dp.message_handler(commands=['start'])
async def start_cmd(message: types.Message):
    await message.reply(
        "👋 Привет! Введи номер заказа, ШК, акта или имя сотрудницы для поиска.\n\n"
        "Или нажми ➕ Добавить сотрудницу, чтобы загрузить фото и ФИО.",
        reply_markup=main_keyboard
    )

@dp.message_handler(lambda message: message.text == "🔄 Обновить таблицы")
async def refresh_tables_handler(message: types.Message):
    await message.reply("🔄 Обновляю таблицы...")
    await refresh_sheets_once()
    await load_employees_from_sheet()
    await message.reply("✅ Таблицы обновлены.")

@dp.message_handler(lambda message: message.text == "➕ Добавить сотрудницу")
async def add_employee_start(message: types.Message):
    await message.reply("📸 Отправьте фото сотрудницы.")
    await EmployeeStates.waiting_for_photo.set()

@dp.message_handler(content_types=ContentType.PHOTO, state=EmployeeStates.waiting_for_photo)
async def employee_photo_received(message: types.Message, state: FSMContext):
    file = await bot.get_file(message.photo[-1].file_id)
    photo_bytes = await bot.download_file(file.file_path)
    await state.update_data(photo_bytes=photo_bytes.read())
    await message.reply("Фото получено. Теперь отправьте ФИО сотрудницы (только латинскими буквами).")
    await EmployeeStates.waiting_for_fio.set()

@dp.message_handler(state=EmployeeStates.waiting_for_fio)
async def employee_fio_received(message: types.Message, state: FSMContext):
    fio = message.text.strip()

    if contains_cyrillic(fio):
        await message.reply("❌ Пишите ФИО только латинскими буквами.")
        return

    data = await state.get_data()
    photo_bytes = data.get('photo_bytes')

    if not fio or not photo_bytes:
        await message.reply("Ошибка: не получены данные.")
        return

    photo_filename = f"{fio}.jpg"
    photo_path = os.path.join('employees', photo_filename)

    with open(photo_path, 'wb') as f:
        f.write(photo_bytes)

    employees.append({"fio": fio, "photo_path": photo_path})
    await save_employee_to_sheet(fio, photo_filename)

    await message.reply(f"✅ Сотрудница '{fio}' добавлена.")
    await state.finish()

@dp.message_handler(state=EmployeeStates.waiting_for_confirm)
async def confirm_add_employee(message: types.Message, state: FSMContext):
    text = message.text.strip().lower()
    if text == "да":
        await message.reply("📸 Отправьте фото новой сотрудницы.")
        await EmployeeStates.waiting_for_photo.set()
    else:
        await message.reply("❌ Добавление отменено.")
        await state.finish()

@dp.callback_query_handler(lambda c: c.data and c.data.startswith('add_employee_yes:'))
async def process_add_employee_yes(callback_query: types.CallbackQuery):
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.from_user.id, "📸 Отправьте фото новой сотрудницы.")
    await EmployeeStates.waiting_for_photo.set()

@dp.callback_query_handler(lambda c: c.data and c.data.startswith('add_employee_no:'))
async def process_add_employee_no(callback_query: types.CallbackQuery):
    await bot.answer_callback_query(callback_query.id)
    await bot.send_message(callback_query.from_user.id, "❌ Добавление отменено.")

@dp.message_handler()
async def search_handler(message: types.Message):
    query_raw = message.text.strip()
    if not query_raw:
        await message.reply("Пустой запрос.")
        return

    query = query_raw.lower()

    # --- Поиск по сотрудницам ---
    matches = [e for e in employees if query in e['fio'].lower()]
    if matches:
        for e in matches:
            if e.get('photo_path') and os.path.exists(e['photo_path']):
                with open(e['photo_path'], 'rb') as photo:
                    await message.reply_photo(photo=photo, caption=e['fio'])
            else:
                await message.reply(f"👤 {e['fio']} (фото не найдено)")
        return

    # --- Поиск по таблицам ---
    if not sheets:
        await message.reply("❌ Таблицы не загружены.")
        return

    found_results = []
    qnorm = normalize(query_raw)
    searching_msg = await message.reply("🔍 Ищу информацию по таблицам...")

    try:
        for sheet, table_name in sheets:
            all_data = sheet.get_all_values()
            if not all_data or len(all_data) < 2:
                continue

            headers = all_data[0]
            rows = all_data[1:]
            for row in rows:
                for i, cell in enumerate(row):
                    if qnorm and qnorm in normalize(cell):
                        info = "\n".join(f"*{headers[j]}*: {row[j] if j < len(row) else ''}" for j in range(len(headers)))
                        result = f"📋 *Источник:* {table_name}\n\n{info}"
                        found_results.append(result)
                        break
    except Exception as e:
        print("Ошибка поиска в таблицах:", e)

    try:
        await bot.delete_message(chat_id=message.chat.id, message_id=searching_msg.message_id)
    except Exception:
        pass

    if found_results:
        for res in found_results:
            await message.reply(res, parse_mode="Markdown")
    else:
        # Вместо текстового запроса — отправляем с кнопками Да/Нет
        keyboard = InlineKeyboardMarkup(row_width=2)
        keyboard.add(
            InlineKeyboardButton("✅ Да", callback_data=f"add_employee_yes:{query_raw}"),
            InlineKeyboardButton("❌ Нет", callback_data=f"add_employee_no:{query_raw}")
        )
        await message.reply(
            f"❌ Не нашёл сотрудницу под именем: {query_raw}\n"
            "❓ Хотите добавить новую сотрудницу?",
            reply_markup=keyboard
        )

if __name__ == '__main__':
    loop = asyncio.get_event_loop()
    loop.create_task(refresh_sheets_loop())
    loop.run_until_complete(refresh_sheets_once())
    loop.run_until_complete(load_employees_from_sheet())
    executor.start_polling(dp, skip_updates=True)
