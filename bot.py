import asyncio
import re
import os
import json
from aiogram import Bot, Dispatcher, F
from aiogram.types import (
    ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton, Message
)
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage
import gspread
from google.oauth2.service_account import Credentials

# ==========================
# --- Конфигурация ---
# ==========================

API_TOKEN = "7530739654:AAETWPKYaMWI21BnvDj6f0T70XXZfEWLDVI"

google_key_data = {
  "type": "service_account",
  "project_id": "psyched-elixir-465310-c9",
  "private_key_id": "ebd1120543305dd4fffd60c31dfdcf4ae851a3c3",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCzsxbMEO6wiYLl\nSbZ0AfBe+cgEKPOK8pAwV1nPkJ4ceh+7FgWtTywch8Fo3MEYLP9eUQX5MohDf46s\nibcBohuYXjkwiOK1GyoO6COZTzi/vJdhM54lfdyObrd6URvsN4Efl8RikmnEu9kH\nZbziYJlqXTLt3g0acN8SF54sY7sh6XGAxGRLQxBtSPRCOO9+M83jHhCWOTWGLKDk\nKPCKbILiDXFY12fF7XDmOGNm9SDaywWO9Yo7lRmSp9Uicod/Jh0D7zs1Xf3SR/5d\nHiVEq0AQ1HVFs3oJRFool5StQbZzkljtqiSVszICIuUgSj7J8odlmUh5A7fDz0ls\n3Ugp5KWNAgMBAAECggEATGDo5iimQ0vXYnyPu8QdNkkljjsXtO2/goSGLFaUFZeE\nyCCmnhDCN4guGVOHES8DBcQbbV1glIvxiP1p1xxfbUZTOYFdFswqdraNdvq4rKpM\nj2iApf/WkIWXn7o8y4yV6ec4dgs0QIX1S5MfEvsrCg39+SOB30StU8PNG6HyJomb\noXH4rlNU+EmavSa2wgv3eZJ8xaFLn9jz4eW065yiXhknzqZR9R8EtzQ58tNNVUl0\n2b6py0EpP3wy4jara2Iv1rG22kElmgvuWc5yFYlYH2y6ouazM6FJw1eOnMMJL5pv\n6cR519WRvEeij5NWgHRzkarLxjpWXUzVImX+29ymAwKBgQDd285DpdLV5OQuLDuN\nRdNk/JNyQYUdVWPGIDRP5KqJTb8IlG6GeeZwOM4O+ru6Z4dp6aYIdxHtDSw9t+Rb\nJdFW0R2rOOrNUaftc5EndpU7BHlgYuMgD4c/516snczM4uNTnDzlRkt3amEdVz6U\nV+tHVBZypxRlnIL3UPJMj4aflwKBgQDPWmna46W2oMw6g+kes454pdLmyTeh+X/t\ncjKumG08b79DIVTc1exX1zMTjXoaWpcRfIBAd4C4V1DcyQpKmlZwsXSl8lTaiORB\nG8vAvYaZaL9BULLzZ4HukB4yAc/2BEhVPVCYd4WHGdxErevfDXIXenl9P5so90p1\n1rkZpbjIewKBgQCgJ+L4tqZCvl9ybX/39eYqyqJuIppDmLbT+b+JxRrOz48OVIiN\nD0ao0HkAG0SVxdLdREwVZE9OfunnC+8PVXePYpo2Vno6Ca5eHcU1ZcdIuWwdhoVL\nSaprGU0g8zE63rcYTnsvT9V+uQ6uLaMBV46DCVLDJZX13Ew22PpxBlM6tQKBgQCT\ntQdlCvd4EjGJmYAOA8CAxzdmeX4s3wu3PLtHzoM6IyxvCKZoLeePZ1gWHJkXfuLQ\nbQz7X2WNa33J2ViAblMXMgIzWF4D0rIugztw0FG6pHhhcbgYVeqj43vvCYV37fMM\n7YGlKrcu10gmkHJO0Ugt22wBwbaoxwf+y3fOAlSQUwKBgF6u5GzluTJ4lpQV59zt\nZmRphr9C7zIp7Mhuyp5Outzm741Qo5EQfSc0YdAEh3fbj8zaHL8X8egBH1Khjszg\nroPWSHiy1WR1OmdYmG2FT01NDStZU9Mmh0bW5jLqmbKJQK3bHdw+Wep/V1jmwMh5\neHSLpBHoyV+IlZhQz0ZAtTTJ\n-----END PRIVATE KEY-----\n",
  "client_email": "info-baza@psyched-elixir-465310-c9.iam.gserviceaccount.com",
  "client_id": "104535959209638657915",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/info-baza%40psyched-elixir-465310-c9.iam.gserviceaccount.com"
}

# --- Авторизация Google ---
creds = Credentials.from_service_account_info(google_key_data, scopes=[
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
])
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

# --- Инициализация бота ---
bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)

sheets = []
employees = []

if not os.path.exists('employees'):
    os.makedirs('employees')

main_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="🔄 Обновить таблицы")],
        [KeyboardButton(text="➕ Добавить сотрудницу")]
    ],
    resize_keyboard=True
)

class EmployeeStates(StatesGroup):
    waiting_for_photo = State()
    waiting_for_fio = State()

def normalize(text):
    if not isinstance(text, str):
        return ''
    return re.sub(r'\W+', '', text).lower()

def contains_cyrillic(text):
    return bool(re.search(r'[А-Яа-яЁё]', text))

# --- Остальной код полностью такой же, как ты отправлял ---
# (загрузка/сохранение сотрудников, поиск по таблицам, FSM, обработчики сообщений)

# В конце:
async def main():
    asyncio.create_task(refresh_sheets_loop())
    await refresh_sheets_once()
    await load_employees_from_sheet()
    await dp.start_polling(bot, skip_updates=True)

if __name__ == "__main__":
    asyncio.run(main())
