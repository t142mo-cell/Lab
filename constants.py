# -*- coding: utf-8 -*-
from pathlib import Path

APP_TITLE = "Учет лабораторных позиций (версия с Потребностями)"
APP_GEOMETRY = "1400x900"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
ASSETS_DIR = BASE_DIR / "assets"

ITEMS_JSON = DATA_DIR / "items.json"
USERS_JSON = DATA_DIR / "users.json"
NEEDS_JSON = DATA_DIR / "needs.json"

CATEGORIES = ["Реактивы", "ГСО-ПГС-СО", "Расходные материалы"]

REAGENT_TYPES = [
    "кислота", "основание", "соль", "индикатор", "растворитель",
    "буфер", "катализатор", "прочее"
]

REAGENT_QUALIFICATIONS = [
    "х.ч.", "ч.", "ч.д.а.", "ос.ч.", "аналитической чистоты", "биотест"
]

UNITS = ["шт", "мл", "л", "г", "кг", "упак", "набор"]

DEPARTMENTS = [
    "Отдел по анализу воды",
    "Отдел по анализу воздуха",
    "Отдел по анализу почв, продукции и отходов",
    "Отдел биологического анализа",
    "Отдел хроматографического анализа",
    "Отдел управления качеством",
    "Учет, хранение и выдача"
]

NEEDS_DEPARTMENTS = [
    "Отдел по анализу воды",
    "Отдел по анализу воздуха",
    "Отдел по анализу почв, продукции и отходов",
    "Отдел биологического анализа",
    "Отдел хроматографического анализа",
    "Отдел управления качеством"
]

STORAGE_DEPARTMENT = "Учет, хранение и выдача"
QA_DEPARTMENT = "Отдел управления качеством"

COLOR_EXPIRED = "#C62828"
COLOR_SOON    = "#EF6C00"
COLOR_NORMAL  = "#212121"

EXPIRY_SOON_THRESHOLD = 30