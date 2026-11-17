import streamlit as st
import pandas as pd
import io
import datetime
import calendar
import zipfile
import os
import base64
import json
import re
import gc
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from supabase import create_client, Client

# ==========================================
# ДИЗАЙН: СКАНДИНАВСКИЙ МИНИМАЛИЗМ (NORDIC LIGHT)
# ==========================================
st.set_page_config(page_title="Учёт выработки", layout="wide")

custom_css = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        color: #0F172A;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    .stApp {
        background-color: #F8FAFC;
    }

    /* Боковая панель */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E2E8F0;
    }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] h4 {
        color: #0F172A;
        letter-spacing: -0.01em;
    }

    /* Вкладки */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #F1F5F9;
        padding: 5px;
        border-radius: 10px;
        border: 1px solid #E2E8F0;
        margin-bottom: 16px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        background-color: transparent;
        border: none;
        border-radius: 7px;
        color: #475569;
        font-weight: 500;
        font-size: 14px;
        padding: 0 18px;
        transition: all 0.15s ease;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: #0F172A;
        background-color: rgba(255, 255, 255, 0.6);
    }
    .stTabs [aria-selected="true"] {
        background-color: #FFFFFF !important;
        color: #2563EB !important;
        font-weight: 600;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.08);
    }
    .stTabs [data-baseweb="tab-highlight"] {
        display: none;
    }

    /* Кнопки */
    .stButton>button[kind="primary"] {
        background-color: #2563EB;
        color: #FFFFFF;
        border: 1px solid #1D4ED8;
        border-radius: 8px;
        font-weight: 500;
        padding: 0.45rem 1.1rem;
        box-shadow: 0 1px 2px rgba(37, 99, 235, 0.18);
        transition: all 0.15s ease;
    }
    .stButton>button[kind="primary"]:hover {
        background-color: #1D4ED8;
        border-color: #1E40AF;
        color: #FFFFFF;
    }
    .stButton>button {
        border-radius: 8px;
        border: 1px solid #CBD5E1;
        background-color: #FFFFFF;
        color: #1E293B;
        font-weight: 500;
    }
    .stButton>button:hover {
        border-color: #94A3B8;
        background-color: #F8FAFC;
    }

    /* Карточки метрик */
    [data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 12px 16px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03);
    }
    [data-testid="stMetricLabel"] {
        font-size: 12px;
        font-weight: 500;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }
    [data-testid="stMetricValue"] {
        font-size: 22px;
        font-weight: 700;
        color: #0F172A;
    }

    /* Поля ввода и таблицы */
    [data-testid="stDataFrame"], .stTable {
        border-radius: 10px;
        border: 1px solid #E2E8F0;
        background-color: #FFFFFF;
        overflow: hidden;
    }

    /* Выравнивание виджетов в строках табеля */
    [data-testid="stHorizontalBlock"] [data-testid="stCheckbox"] {
        margin-top: 0px !important;
        margin-bottom: 0px !important;
        display: flex;
        align-items: center;
    }
    [data-testid="stHorizontalBlock"] [data-testid="stNumberInput"],
    [data-testid="stHorizontalBlock"] [data-testid="stTextInput"] {
        margin-top: 0px !important;
        margin-bottom: 0px !important;
    }

    /* Шапка и плашки */
    .nordic-header-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 18px 22px;
        margin-bottom: 18px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03);
    }
    .nordic-title {
        font-size: 21px;
        font-weight: 700;
        color: #0F172A;
        margin: 0;
        letter-spacing: -0.02em;
    }
    .nordic-subtitle {
        font-size: 13px;
        color: #64748B;
        margin-top: 3px;
    }
    .nordic-period-badge {
        background-color: #EFF6FF;
        color: #1D4ED8;
        border: 1px solid #BFDBFE;
        padding: 6px 14px;
        border-radius: 999px;
        font-size: 13px;
        font-weight: 600;
    }
    .nordic-period-badge-empty {
        background-color: #F1F5F9;
        color: #475569;
        border: 1px solid #CBD5E1;
        padding: 6px 14px;
        border-radius: 999px;
        font-size: 13px;
        font-weight: 500;
    }
    .subgroup-banner {
        background-color: #F1F5F9;
        border-left: 3px solid #2563EB;
        padding: 7px 14px;
        border-radius: 0 8px 8px 0;
        font-size: 13px;
        font-weight: 600;
        color: #1E293B;
        margin: 14px 0 8px 0;
        letter-spacing: 0.01em;
    }
    .emp-label {
        font-size: 14px;
        font-weight: 500;
        color: #0F172A;
        line-height: 1.25;
        margin: 0;
    }
    .emp-post {
        font-size: 12px;
        color: #64748B;
        font-weight: 400;
        margin-top: 2px;
    }
    .row-divider {
        border-bottom: 1px solid #E2E8F0;
        margin: 4px 0 8px 0;
    }
    .empty-cell-dash {
        color: #94A3B8;
        font-size: 14px;
        text-align: center;
        padding: 6px 0;
    }
    .badge-shift {
        display: inline-block;
        background-color: #EFF6FF;
        color: #1D4ED8;
        border: 1px solid #BFDBFE;
        font-size: 11px;
        font-weight: 600;
        padding: 1px 7px;
        border-radius: 999px;
        margin-left: 6px;
    }
    .badge-sub {
        display: inline-block;
        background-color: #FFF7ED;
        color: #C2410C;
        border: 1px solid #FED7AA;
        font-size: 11px;
        font-weight: 600;
        padding: 1px 7px;
        border-radius: 999px;
        margin-left: 6px;
    }
    .log-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 10px 14px;
        margin-bottom: 8px;
        font-size: 13.5px;
        line-height: 1.5;
    }
    .pin-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 24px 28px;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.04);
        margin-top: 24px;
        margin-bottom: 12px;
    }
    .pin-tag {
        display: inline-block;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #2563EB;
        background-color: #EFF6FF;
        border: 1px solid #BFDBFE;
        padding: 3px 9px;
        border-radius: 6px;
        margin-bottom: 10px;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==========================================
# 1. СЛОВАРИ ОТДЕЛОВ, ПАРОЛЕЙ И ДОЛЖНОСТЕЙ
# ==========================================
DEPARTMENTS = {
    "pf": "Отделение Полуфабрикатов",
    "main_deboning": "Основная обвалка",
    "chicken_deboning": "Куриная обвалка",
    "offal": "Ливерное отделение",
    "delicatessen": "Деликатесное отделение",
    "formovka": "Отделение формовки",
    "farshesozdanie": "Отделение фаршесоставления"
}

DEPT_PASSWORDS = {
    "pf": "1145",
    "delicatessen": "1145",
    "formovka": "2278",
    "farshesozdanie": "2278",
    "main_deboning": "3591",
    "chicken_deboning": "4826",
    "offal": "5739"
}

FILE_PREFIXES = {
    "pf": "Полуфабрикаты",
    "main_deboning": "Основная_обвалка",
    "chicken_deboning": "Куриная_обвалка",
    "offal": "Ливерное",
    "delicatessen": "Деликатесы",
    "formovka": "Формовка",
    "farshesozdanie": "Фаршесоставление"
}

FORMOVKA_SUBGROUPS = ["Формовка", "Клипсатор", "Сосисочная линия"]
FORMOVKA_TABELS = {
    "Формовка": "tpl_tabel_formovka_formovka.xlsm",
    "Клипсатор": "tpl_tabel_formovka_klipsator.xlsm",
    "Сосисочная линия": "tpl_tabel_formovka_line.xlsm"
}
FORMOVKA_SUBS = {
    "Грузчики": "tpl_sub_formovka_gruzchiki.xlsm",
    "Клипсатор": "tpl_sub_formovka_klipsator.xlsm",
    "Сосисочная линия": "tpl_sub_formovka_line.xlsm"
}

MAIN_DEBONING_SUBGROUPS = ["Обвалка", "Жиловка", "Операторы"]
MAIN_DEBONING_TABELS = {
    "Обвалка": "tpl_tabel_main_deboning_obvalka.xlsm",
    "Жиловка": "tpl_tabel_main_deboning_zhilovka.xlsm",
    "Операторы": "tpl_tabel_main_deboning_operators.xlsm"
}
MAIN_DEBONING_POKAZ = {
    "Обвалка": "tpl_pokaz_main_deboning_obvalka.xlsx",
    "Жиловка": "tpl_pokaz_main_deboning_zhilovka.xlsx"
}

CHICKEN_DEBONING_SUBGROUPS = ["Обвалка", "Жиловка"]
CHICKEN_DEBONING_TABELS = {
    "Обвалка": "tpl_tabel_chicken_deboning_obvalka.xlsm",
    "Жиловка": "tpl_tabel_chicken_deboning_zhilovka.xlsm"
}
CHICKEN_DEBONING_SUBS = {
    "Жиловка": "tpl_sub_chicken_deboning_zhilovka.xlsm",
    "Грузчики": "tpl_sub_chicken_deboning_gruzchiki.xlsm"
}
CHICKEN_DEBONING_POKAZ = {
    "Жиловка": "tpl_pokaz_chicken_deboning_zhilovka.xlsm"
}

PAY_TYPES_DEBONING = ["По коэффициенту бригады", "Фиксированный (100%)"]

POSITIONS_BY_DEPT = {
    "pf": ["Изготовитель мясных п/ф", "Весовщик", "Грузчик", "Формовщик", "Жиловщик", "Оператор линии", "Бригадир", "Мастер"],
    "main_deboning": [
        "Обвальщик",
        "Жиловщик",
        "Жиловщик (шкура)",
        "Распиловщик",
        "Засольщик мяса и продуктов из мяса",
        "Оператор машины по изм. мяса",
        "Подсобный рабочий",
        "Весовщик-логист",
        "Мастер"
    ],
    "chicken_deboning": [
        "Обвальщик птицы",
        "Жиловщик птицы",
        "Подача",
        "Грузчик",
        "Мастер"
    ],
    "offal": ["Обработчик субпродуктов", "Весовщик", "Грузчик", "Мастер"],
    "delicatessen": ["Изготовитель деликатесов", "Формовщик", "Весовщик", "Грузчик", "Мастер"],
    "formovka": [
        "формовщик",
        "грузчик формовки колбасных изделий",
        "оператор по производству колбас",
        "грузчик линии по производству колбас",
        "оператор линии производства сосисок",
        "подсобный рабочий сосисочная линия",
        "Мастер"
    ],
    "farshesozdanie": ["Фаршесоставитель", "Весовщик", "Грузчик", "Бригадир", "Мастер"]
}

DEFAULT_RATES_BY_DEPT = {
    "pf": [
        {"post": "Весовщик", "grade": "", "rate": 5000},
        {"post": "Грузчик", "grade": "", "rate": 4300},
        {"post": "Изготовитель мясных п/ф", "grade": "1 разряд", "rate": 3700},
        {"post": "Изготовитель мясных п/ф", "grade": "2 разряд", "rate": 4500},
        {"post": "Мастер", "grade": "", "rate": 7370},
    ],
    "main_deboning": [
        {"post": "Обвальщик", "grade": "", "rate": 8200},
        {"post": "Жиловщик", "grade": "", "rate": 6460},
        {"post": "Жиловщик (шкура)", "grade": "", "rate": 5000},
        {"post": "Распиловщик", "grade": "", "rate": 4750},
        {"post": "Засольщик мяса и продуктов из мяса", "grade": "", "rate": 4850},
        {"post": "Оператор машины по изм. мяса", "grade": "", "rate": 4750},
        {"post": "Подсобный рабочий", "grade": "", "rate": 4000},
        {"post": "Весовщик-логист", "grade": "", "rate": 5000},
        {"post": "Мастер", "grade": "", "rate": 7370},
    ],
    "chicken_deboning": [
        {"post": "Обвальщик птицы", "grade": "", "rate": 5280},
        {"post": "Жиловщик птицы", "grade": "", "rate": 4440},
        {"post": "Подача", "grade": "", "rate": 5000},
        {"post": "Грузчик", "grade": "", "rate": 4300},
        {"post": "Мастер", "grade": "", "rate": 7370},
    ],
    "offal": [
        {"post": "Обработчик субпродуктов", "grade": "", "rate": 4500},
        {"post": "Весовщик", "grade": "", "rate": 4000},
        {"post": "Грузчик", "grade": "", "rate": 3800},
        {"post": "Мастер", "grade": "", "rate": 7370},
    ],
    "delicatessen": [
        {"post": "Изготовитель деликатесов", "grade": "1 разряд", "rate": 5200},
        {"post": "Изготовитель деликатесов", "grade": "2 разряд", "rate": 4600},
        {"post": "Формовщик", "grade": "", "rate": 4300},
        {"post": "Весовщик", "grade": "", "rate": 4000},
        {"post": "Грузчик", "grade": "", "rate": 3800},
        {"post": "Мастер", "grade": "", "rate": 7370},
    ],
    "formovka": [
        {"post": "формовщик", "grade": "", "rate": 4500},
        {"post": "грузчик формовки колбасных изделий", "grade": "", "rate": 4000},
        {"post": "оператор по производству колбас", "grade": "", "rate": 4800},
        {"post": "грузчик линии по производству колбас", "grade": "", "rate": 4000},
        {"post": "оператор линии производства сосисок", "grade": "", "rate": 4800},
        {"post": "подсобный рабочий сосисочная линия", "grade": "", "rate": 3800},
        {"post": "Мастер", "grade": "", "rate": 7370},
    ],
    "farshesozdanie": [
        {"post": "Фаршесоставитель", "grade": "1 разряд", "rate": 5200},
        {"post": "Фаршесоставитель", "grade": "2 разряд", "rate": 4600},
        {"post": "Весовщик", "grade": "", "rate": 4000},
        {"post": "Грузчик", "grade": "", "rate": 3800},
        {"post": "Мастер", "grade": "", "rate": 7370},
    ],
}

@st.cache_resource
def init_connection():
    try:
        return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    except Exception:
        return None

supabase = init_connection()

MONTHS_RU = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

ORANGE_FILL = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
BLUE_FILL = PatternFill(start_color="8EA9DB", end_color="8EA9DB", fill_type="solid")
NO_FILL = PatternFill(fill_type=None)

# ==========================================
# 2. ИНИЦИАЛИЗАЦИЯ СЕССИИ
# ==========================================
_defaults = {
    "current_dept": "",
    "unlocked_pins": [],
    "active_month": "", "active_year": "", "current_id": None,
    "shifts": {},
    "tonnage": {}, "subordinates": [],
    "brigade_percents": {},
    "emp_percents": {},
    "emp_remarks": {},
    "emp_remarks_2": {},
    "emp_hours": {},
    "staff_roster": [], "rates": []
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ==========================================
# 3. ХЕЛПЕРЫ
# ==========================================
def is_deboning_dept(dept_id):
    return dept_id in ("main_deboning", "chicken_deboning")

def is_master_post(post_name):
    return "мастер" in str(post_name or "").strip().lower()

def calc_master_shift_pay(shift_or_hourly_rate, hours):
    rate = float(shift_or_hourly_rate)
    hrs = float(hours)
    if rate <= 2000:
        return round(rate * hrs, 2)
    return round((rate / 11.0) * hrs, 2)

def get_current_positions():
    positions = []
    for r in st.session_state.rates:
        if r.get("post") and str(r["post"]).strip() not in positions:
            positions.append(str(r["post"]).strip())
    for p in POSITIONS_BY_DEPT.get(st.session_state.current_dept, []):
        if p not in positions:
            positions.append(p)
    if not positions:
        positions = ["Сотрудник", "Мастер"]
    return positions

def make_short_name(full_name):
    parts = str(full_name).strip().split()
    return f"{parts[0]} {parts[1][0]}." if len(parts) > 1 else (parts[0] if parts else "")

def get_rate_for(post, grade):
    for r in st.session_state.rates:
        if r["post"] == post and r["grade"] == grade:
            return r["rate"]
    for r in st.session_state.rates:
        if r["post"] == post:
            return r["rate"]
    if is_master_post(post):
        return 7370
    return 4500

def get_template_name(base_name, dept_id, ext):
    return f"{base_name}_{dept_id}.{ext}"

def match_sheet_by_fio(wb, full_name):
    target_full = str(full_name).strip().lower()
    target_short = make_short_name(full_name).strip().lower()
    target_surname = target_full.split()[0] if target_full else ""

    for sname in wb.sheetnames:
        low = sname.strip().lower()
        if low == target_full[:31] or low == target_short:
            return sname
    if target_surname:
        for sname in wb.sheetnames:
            low = sname.strip().lower()
            if low.startswith(target_surname + " ") or low == target_surname:
                return sname
    return None

def load_initial_staff(dept_id):
    res = []
    if dept_id == "formovka":
        for sub, filename in FORMOVKA_TABELS.items():
            if not os.path.exists(filename):
                continue
            with open(filename, "rb") as f:
                wb = load_workbook(io.BytesIO(f.read()), keep_vba=True, read_only=True)
            for sname in wb.sheetnames:
                if sname in ("Итог", "рубкг") or sname.startswith("[Уволен]"):
                    continue
                res.append({
                    "full_name": sname,
                    "short_name": make_short_name(sname),
                    "post": "Сотрудник",
                    "grade": "—",
                    "subgroup": sub
                })
            wb.close()
    elif dept_id == "main_deboning":
        skip_sheets = {"Итог", "Общий", "рубкг", "тоннаж"}
        for sub, filename in MAIN_DEBONING_TABELS.items():
            if not os.path.exists(filename):
                continue
            with open(filename, "rb") as f:
                wb = load_workbook(io.BytesIO(f.read()), keep_vba=True, read_only=True)
            for sname in wb.sheetnames:
                if sname.strip() in skip_sheets or sname.startswith("[Уволен]"):
                    continue
                low = sname.lower()
                if sub == "Обвалка":
                    if any(k in low for k in ["паронян", "дядик", "макеева"]):
                        post = "Мастер"
                        pay_type = "Фиксированный (100%)"
                    elif "арашов" in low:
                        post = "Весовщик-логист"
                        pay_type = "Фиксированный (100%)"
                    elif any(k in low for k in ["хафизов", "убайдуллоев", "махсудов"]):
                        post = "Распиловщик"
                        pay_type = "Фиксированный (100%)"
                    else:
                        post = "Обвальщик"
                        pay_type = "По коэффициенту бригады"
                elif sub == "Жиловка":
                    if any(k in low for k in ["ермишкин", "халимджонова", "ширинов"]):
                        post = "Жиловщик (шкура)"
                        pay_type = "Фиксированный (100%)"
                    else:
                        post = "Жиловщик"
                        pay_type = "По коэффициенту бригады"
                else:
                    post = "Оператор машины по изм. мяса"
                    pay_type = "Фиксированный (100%)"

                res.append({
                    "full_name": sname.strip(),
                    "short_name": make_short_name(sname),
                    "post": post,
                    "grade": "—",
                    "subgroup": sub,
                    "pay_type": pay_type
                })
            wb.close()
    elif dept_id == "chicken_deboning":
        skip_sheets = {"Итог", "Общий", "рубкг", "тоннаж"}
        for sub, filename in CHICKEN_DEBONING_TABELS.items():
            if not os.path.exists(filename):
                continue
            with open(filename, "rb") as f:
                wb = load_workbook(io.BytesIO(f.read()), keep_vba=True, read_only=True)
            for sname in wb.sheetnames:
                if sname.strip() in skip_sheets or sname.startswith("[Уволен]"):
                    continue
                low = sname.lower()
                if sub == "Обвалка":
                    if any(k in low for k in ["баринов", "зайцев"]):
                        post = "Мастер"
                        pay_type = "Фиксированный (100%)"
                    elif "богдагулов" in low:
                        post = "Подача"
                        pay_type = "Фиксированный (100%)"
                    elif "исоев" in low:
                        post = "Грузчик"
                        pay_type = "Фиксированный (100%)"
                    else:
                        post = "Обвальщик птицы"
                        pay_type = "По коэффициенту бригады"
                else:
                    post = "Жиловщик птицы"
                    pay_type = "По коэффициенту бригады"

                res.append({
                    "full_name": sname.strip(),
                    "short_name": make_short_name(sname),
                    "post": post,
                    "grade": "—",
                    "subgroup": sub,
                    "pay_type": pay_type
                })
            wb.close()
    else:
        tpl_name = get_template_name("tpl_tabel", dept_id, "xlsm")
        if not os.path.exists(tpl_name):
            return []
        with open(tpl_name, "rb") as f:
            wb = load_workbook(io.BytesIO(f.read()), keep_vba=True, read_only=True)
        for sname in wb.sheetnames:
            if sname in ("Итог", "рубкг") or sname.startswith("[Уволен]"):
                continue
            low = sname.lower()
            if "рузиева" in low:
                post = "Весовщик"
            elif "насруддинов" in low or "рузиев и" in low:
                post = "Грузчик"
            elif dept_id == "pf":
                post = "Изготовитель мясных п/ф"
            else:
                positions = POSITIONS_BY_DEPT.get(dept_id, ["Сотрудник"])
                post = positions[0] if positions else "Сотрудник"
            res.append({"full_name": sname, "short_name": make_short_name(sname), "post": post, "grade": "—"})
        wb.close()
    return res

# ==========================================
# 4. РАБОТА С БАЗОЙ ДАННЫХ
# ==========================================
def save_global_settings():
    if not supabase or not st.session_state.current_dept:
        return
    dept_id = st.session_state.current_dept
    rec_id = f"GLOBAL_SETTINGS_{dept_id}"
    try:
        payload = {"staff_roster": st.session_state.staff_roster, "rates": st.session_state.rates}
        json_str = json.dumps(payload, ensure_ascii=False)
        data = {"pf_base64": base64.b64encode(json_str.encode('utf-8')).decode("utf-8")}
        res = supabase.table("app_data").select("id").eq("id", rec_id).execute()
        if res.data:
            supabase.table("app_data").update(data).eq("id", rec_id).execute()
        else:
            data["id"] = rec_id
            supabase.table("app_data").insert(data).execute()
    except Exception as e:
        st.error(f"Ошибка сохранения настроек: {e}")

def load_global_settings():
    if not supabase or not st.session_state.current_dept:
        return
    dept_id = st.session_state.current_dept
    rec_id = f"GLOBAL_SETTINGS_{dept_id}"
    try:
        res = supabase.table("app_data").select("*").eq("id", rec_id).execute()
        if res.data and res.data[0]["pf_base64"]:
            raw_json = base64.b64decode(res.data[0]["pf_base64"]).decode('utf-8')
            payload = json.loads(raw_json)
            st.session_state.staff_roster = payload.get("staff_roster", [])
            st.session_state.rates = payload.get("rates", [])
            if not st.session_state.rates:
                st.session_state.rates = DEFAULT_RATES_BY_DEPT.get(dept_id, [{"post": "Сотрудник", "grade": "", "rate": 4500}])
        else:
            st.session_state.staff_roster = load_initial_staff(dept_id)
            st.session_state.rates = DEFAULT_RATES_BY_DEPT.get(dept_id, [{"post": "Сотрудник", "grade": "", "rate": 4500}])
            save_global_settings()
    except Exception as e:
        st.error(f"Ошибка загрузки настроек: {e}")

def save_month_to_db():
    if not supabase or not st.session_state.current_id:
        return
    try:
        payload = {
            "shifts": st.session_state.shifts,
            "tonnage": st.session_state.tonnage,
            "subordinates": st.session_state.subordinates,
            "brigade_percents": st.session_state.brigade_percents,
            "emp_percents": st.session_state.emp_percents,
            "emp_remarks": st.session_state.emp_remarks,
            "emp_remarks_2": st.session_state.emp_remarks_2,
            "emp_hours": st.session_state.emp_hours
        }
        json_str = json.dumps(payload, ensure_ascii=False)
        data = {"pf_base64": base64.b64encode(json_str.encode('utf-8')).decode("utf-8")}
        supabase.table("app_data").update(data).eq("id", st.session_state.current_id).execute()
    except Exception as e:
        st.error(f"Ошибка сохранения месяца: {e}")

def load_month_from_db(record_id):
    if not supabase:
        return False
    try:
        res = supabase.table("app_data").select("*").eq("id", record_id).execute()
        if res.data and res.data[0]["pf_base64"]:
            raw_json = base64.b64decode(res.data[0]["pf_base64"]).decode('utf-8')
            payload = json.loads(raw_json)
            raw_shifts = payload.get("shifts", {})
            st.session_state.shifts = {}
            for k, v in raw_shifts.items():
                if isinstance(v, list):
                    st.session_state.shifts[k] = {"timesheet": v, "graph_only": {}}
                else:
                    go = v.get("graph_only", {})
                    if isinstance(go, list):
                        go = {name: "" for name in go}
                    st.session_state.shifts[k] = {"timesheet": v.get("timesheet", []), "graph_only": go}
            st.session_state.tonnage = payload.get("tonnage", {})
            st.session_state.subordinates = payload.get("subordinates", [])
            st.session_state.brigade_percents = payload.get("brigade_percents", {})
            st.session_state.emp_percents = payload.get("emp_percents", {})
            st.session_state.emp_remarks = payload.get("emp_remarks", {})
            st.session_state.emp_remarks_2 = payload.get("emp_remarks_2", {})
            st.session_state.emp_hours = payload.get("emp_hours", {})
            st.session_state.current_id = record_id
            parts = record_id.split("_")
            st.session_state.active_month = parts[0]
            st.session_state.active_year = parts[1]
            return True
    except Exception as e:
        st.error(f"Ошибка загрузки месяца: {e}")
    return False

# ==========================================
# 5. ГЕНЕРАЦИЯ ФАЙЛОВ
# ==========================================
def generate_and_download(status_container):
    try:
        zip_buf = io.BytesIO()
        m_name = st.session_state.active_month
        y_val = st.session_state.active_year
        dept_id = st.session_state.current_dept

        prefix = FILE_PREFIXES.get(dept_id, "Отчет")

        tpl_gr_name = get_template_name("tpl_graph", dept_id, "xlsx")
        tpl_sh_name = "tpl_sh.xlsx"

        if not os.path.exists(tpl_gr_name) or not os.path.exists(tpl_sh_name):
            st.error("Не все основные шаблоны загружены на GitHub (График или Шашина).")
            return None

        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            # 5.1. ТАБЕЛИ
            status_container.write("Обрабатываем Табели...")
            if dept_id == "formovka":
                for sub_name, tpl_pf_name in FORMOVKA_TABELS.items():
                    if not os.path.exists(tpl_pf_name):
                        status_container.write(f"Файл {tpl_pf_name} не найден, пропускаем.")
                        continue

                    with open(tpl_pf_name, "rb") as f:
                        wb_pf = load_workbook(io.BytesIO(f.read()), keep_vba=True)

                    wb_pf.calculation.fullCalcOnLoad = True
                    tpl_pf_sheet = next((s for s in wb_pf.sheetnames if s not in ("Итог", "рубкг")), None)

                    for s_name in list(wb_pf.sheetnames):
                        if s_name not in ("Итог", "рубкг", tpl_pf_sheet):
                            del wb_pf[s_name]

                    sub_roster = [e for e in st.session_state.staff_roster if e.get("subgroup") == sub_name]
                    roster_map = {e["full_name"]: e for e in sub_roster}

                    for emp in sub_roster:
                        sheet_name = emp["full_name"][:31]
                        rate_val = get_rate_for(emp["post"], emp.get("grade", ""))
                        if is_master_post(emp["post"]):
                            rate_val = calc_master_shift_pay(rate_val, 11.0)
                        if sheet_name not in wb_pf.sheetnames and tpl_pf_sheet in wb_pf.sheetnames:
                            new_ws = wb_pf.copy_worksheet(wb_pf[tpl_pf_sheet])
                            new_ws.title = sheet_name
                            for r in range(2, 33):
                                if new_ws.cell(row=r, column=2).data_type != 'f':
                                    new_ws.cell(row=r, column=2).value = None
                                if new_ws.cell(row=r, column=3).data_type != 'f':
                                    new_ws.cell(row=r, column=3).value = rate_val

                    if tpl_pf_sheet in wb_pf.sheetnames and not any(e["full_name"][:31] == tpl_pf_sheet for e in sub_roster):
                        del wb_pf[tpl_pf_sheet]

                    for day_str, data in st.session_state.shifts.items():
                        day = int(day_str)
                        timesheet_emps = data.get("timesheet", [])
                        graph_only_emps = data.get("graph_only", {})
                        day_hours_map = st.session_state.emp_hours.get(day_str, {})

                        for emp_full_name in timesheet_emps:
                            if emp_full_name in graph_only_emps:
                                continue
                            sheet_name = emp_full_name[:31]
                            if sheet_name in wb_pf.sheetnames:
                                wb_pf[sheet_name].cell(row=day + 1, column=2, value=1)
                                emp_obj = roster_map.get(emp_full_name)
                                if emp_obj and is_master_post(emp_obj["post"]):
                                    hrs = float(day_hours_map.get(emp_full_name, 11.0))
                                    base_r = get_rate_for(emp_obj["post"], emp_obj.get("grade", ""))
                                    if wb_pf[sheet_name].cell(row=day + 1, column=3).data_type != 'f':
                                        wb_pf[sheet_name].cell(row=day + 1, column=3, value=calc_master_shift_pay(base_r, hrs))

                        if day_str in st.session_state.tonnage and "рубкг" in wb_pf.sheetnames:
                            wb_pf["рубкг"].cell(row=day + 1, column=6, value=st.session_state.tonnage[day_str])

                    if "Итог" in wb_pf.sheetnames:
                        ws_it = wb_pf["Итог"]
                        for r in range(2, 50):
                            ws_it.cell(row=r, column=1).value = None
                            ws_it.cell(row=r, column=2).value = None
                            ws_it.cell(row=r, column=3).value = None

                        for i, emp in enumerate(sub_roster):
                            r = 2 + i
                            sheet_name = emp["full_name"][:31]
                            ws_it.cell(row=r, column=1, value=emp["full_name"])
                            ws_it.cell(row=r, column=2, value=f"='{sheet_name}'!J33")
                            ws_it.cell(row=r, column=3, value=f"='{sheet_name}'!B33")

                        itogo_r = 2 + len(sub_roster)
                        ws_it.cell(row=itogo_r, column=1, value="ИТОГО:")
                        ws_it.cell(row=itogo_r, column=2, value=f"=SUM(B2:B{itogo_r-1})")
                        ws_it.cell(row=itogo_r, column=3, value=f"=SUM(C2:C{itogo_r-1})")
                        ws_it.delete_rows(itogo_r + 1, 50)

                    buf = io.BytesIO()
                    wb_pf.save(buf)
                    zf.writestr(f"Табель_{sub_name}_{m_name}_{y_val}.xlsm", buf.getvalue())
                    del wb_pf
                    gc.collect()

            elif is_deboning_dept(dept_id):
                service_sheets = {"Итог", "Общий", "рубкг", "тоннаж"}
                tabels_map = MAIN_DEBONING_TABELS if dept_id == "main_deboning" else CHICKEN_DEBONING_TABELS
                pokaz_map = MAIN_DEBONING_POKAZ if dept_id == "main_deboning" else CHICKEN_DEBONING_POKAZ

                for sub_name, tpl_pf_name in tabels_map.items():
                    if not os.path.exists(tpl_pf_name):
                        status_container.write(f"Файл {tpl_pf_name} не найден, пропускаем.")
                        continue

                    with open(tpl_pf_name, "rb") as f:
                        wb_pf = load_workbook(io.BytesIO(f.read()), keep_vba=True)

                    wb_pf.calculation.fullCalcOnLoad = True
                    tpl_pf_sheet = next((s for s in wb_pf.sheetnames if s not in service_sheets), None)

                    for s_name in list(wb_pf.sheetnames):
                        if s_name not in service_sheets and s_name != tpl_pf_sheet:
                            del wb_pf[s_name]

                    if "Общий" in wb_pf.sheetnames:
                        ws_common = wb_pf["Общий"]
                        for day_num in range(1, 32):
                            d_str = str(day_num)
                            b_data = st.session_state.brigade_percents.get(d_str, {})
                            if sub_name == "Обвалка":
                                pct = float(b_data.get("obvalka", 100.0))
                            elif sub_name == "Жиловка":
                                pct = float(b_data.get("zhilovka", 100.0))
                            else:
                                pct = 100.0
                            ws_common.cell(row=day_num + 1, column=2, value=round(pct / 100.0, 4))

                    sub_roster = [e for e in st.session_state.staff_roster if e.get("subgroup") == sub_name]

                    for emp in sub_roster:
                        sheet_name = emp["full_name"][:31]
                        rate_val = get_rate_for(emp["post"], emp.get("grade", ""))
                        if is_master_post(emp["post"]):
                            rate_val = calc_master_shift_pay(rate_val, 11.0)
                        if sheet_name not in wb_pf.sheetnames and tpl_pf_sheet in wb_pf.sheetnames:
                            new_ws = wb_pf.copy_worksheet(wb_pf[tpl_pf_sheet])
                            new_ws.title = sheet_name
                        if sheet_name in wb_pf.sheetnames:
                            ws_emp = wb_pf[sheet_name]
                            for r in range(2, 33):
                                if ws_emp.cell(row=r, column=2).data_type != 'f':
                                    ws_emp.cell(row=r, column=2).value = None
                                if ws_emp.cell(row=r, column=3).data_type != 'f':
                                    ws_emp.cell(row=r, column=3).value = rate_val
                                if ws_emp.cell(row=r, column=4).data_type != 'f':
                                    ws_emp.cell(row=r, column=4).value = 1

                    if tpl_pf_sheet in wb_pf.sheetnames and not any(e["full_name"][:31] == tpl_pf_sheet for e in sub_roster):
                        del wb_pf[tpl_pf_sheet]

                    roster_map = {e["full_name"]: e for e in sub_roster}

                    for day_str, data in st.session_state.shifts.items():
                        day = int(day_str)
                        timesheet_emps = data.get("timesheet", [])
                        graph_only_emps = data.get("graph_only", {})
                        day_emp_pcts = st.session_state.emp_percents.get(day_str, {})
                        day_hours_map = st.session_state.emp_hours.get(day_str, {})
                        b_data = st.session_state.brigade_percents.get(day_str, {})
                        brigade_pct = float(b_data.get("obvalka", 100.0)) if sub_name == "Обвалка" else (
                            float(b_data.get("zhilovka", 100.0)) if sub_name == "Жиловка" else 100.0
                        )

                        for emp_full_name in timesheet_emps:
                            if emp_full_name in graph_only_emps:
                                continue
                            if emp_full_name not in roster_map:
                                continue
                            sheet_name = emp_full_name[:31]
                            if sheet_name in wb_pf.sheetnames:
                                ws_emp = wb_pf[sheet_name]
                                ws_emp.cell(row=day + 1, column=2, value=1)
                                emp_obj = roster_map[emp_full_name]

                                if is_master_post(emp_obj["post"]):
                                    hrs = float(day_hours_map.get(emp_full_name, 11.0))
                                    base_r = get_rate_for(emp_obj["post"], emp_obj.get("grade", ""))
                                    if ws_emp.cell(row=day + 1, column=3).data_type != 'f':
                                        ws_emp.cell(row=day + 1, column=3, value=calc_master_shift_pay(base_r, hrs))
                                    ws_emp.cell(row=day + 1, column=4, value=1)
                                else:
                                    default_pct = brigade_pct if emp_obj.get("pay_type") == "По коэффициенту бригады" else 100.0
                                    actual_pct = float(day_emp_pcts.get(emp_full_name, default_pct))

                                    if "Общий" in wb_pf.sheetnames and abs(actual_pct - brigade_pct) < 0.001 and emp_obj.get("pay_type") == "По коэффициенту бригады":
                                        ws_emp.cell(row=day + 1, column=4, value=f"=Общий!B{day + 1}")
                                    else:
                                        ws_emp.cell(row=day + 1, column=4, value=round(actual_pct / 100.0, 4))

                        if day_str in st.session_state.tonnage:
                            if "тоннаж" in wb_pf.sheetnames:
                                wb_pf["тоннаж"].cell(row=day + 1, column=2, value=st.session_state.tonnage[day_str])
                            elif dept_id == "chicken_deboning" and "рубкг" in wb_pf.sheetnames:
                                if wb_pf["рубкг"].cell(row=day + 1, column=8).data_type != 'f':
                                    wb_pf["рубкг"].cell(row=day + 1, column=8, value=st.session_state.tonnage[day_str])

                    if "Итог" in wb_pf.sheetnames:
                        ws_it = wb_pf["Итог"]
                        for r in range(2, 50):
                            ws_it.cell(row=r, column=1).value = None
                            ws_it.cell(row=r, column=2).value = None
                            ws_it.cell(row=r, column=3).value = None

                        for i, emp in enumerate(sub_roster):
                            r = 2 + i
                            sheet_name = emp["full_name"][:31]
                            ws_it.cell(row=r, column=1, value=emp["full_name"])
                            ws_it.cell(row=r, column=2, value=f"='{sheet_name}'!E33")
                            ws_it.cell(row=r, column=3, value=f"='{sheet_name}'!B33")

                        itogo_r = 2 + len(sub_roster)
                        ws_it.cell(row=itogo_r, column=1, value="ИТОГО:")
                        ws_it.cell(row=itogo_r, column=2, value=f"=SUM(B2:B{itogo_r-1})")
                        ws_it.cell(row=itogo_r, column=3, value=f"=SUM(C2:C{itogo_r-1})")
                        ws_it.delete_rows(itogo_r + 1, 50)

                    buf = io.BytesIO()
                    wb_pf.save(buf)
                    zf.writestr(f"Табель_{sub_name}_{prefix}_{m_name}_{y_val}.xlsm", buf.getvalue())
                    del wb_pf
                    gc.collect()

                # Генерация файлов Показателей разделки
                status_container.write("Обрабатываем Показатели разделки...")
                for sub_name, tpl_pokaz_name in pokaz_map.items():
                    if not os.path.exists(tpl_pokaz_name):
                        status_container.write(f"Файл {tpl_pokaz_name} не найден, пропускаем.")
                        continue

                    with open(tpl_pokaz_name, "rb") as f:
                        wb_pk = load_workbook(io.BytesIO(f.read()))

                    wb_pk.calculation.fullCalcOnLoad = True
                    sub_roster = [e for e in st.session_state.staff_roster if e.get("subgroup") == sub_name]

                    for emp in sub_roster:
                        matched_sheet = match_sheet_by_fio(wb_pk, emp["full_name"])
                        if not matched_sheet:
                            continue
                        ws_pk = wb_pk[matched_sheet]
                        for day_num in range(1, 32):
                            d_str = str(day_num)
                            day_data = st.session_state.shifts.get(d_str, {})
                            is_working = (
                                emp["full_name"] in day_data.get("timesheet", [])
                                and emp["full_name"] not in day_data.get("graph_only", {})
                            )
                            rem_val = int(st.session_state.emp_remarks.get(d_str, {}).get(emp["full_name"], 0))
                            if ws_pk.cell(row=day_num + 1, column=2).data_type != 'f':
                                ws_pk.cell(row=day_num + 1, column=2).value = rem_val if (is_working and rem_val > 0) else None

                            if dept_id == "chicken_deboning" and ws_pk.max_column >= 3:
                                rem_val_2 = int(st.session_state.emp_remarks_2.get(d_str, {}).get(emp["full_name"], 0))
                                if ws_pk.cell(row=day_num + 1, column=3).data_type != 'f':
                                    ws_pk.cell(row=day_num + 1, column=3).value = rem_val_2 if (is_working and rem_val_2 > 0) else None

                    buf = io.BytesIO()
                    wb_pk.save(buf)
                    zf.writestr(f"Показатели_разделки_{sub_name}_{prefix}_{m_name}_{y_val}.xlsx", buf.getvalue())
                    del wb_pk
                    gc.collect()

            else:
                tpl_pf_name = get_template_name("tpl_tabel", dept_id, "xlsm")
                if not os.path.exists(tpl_pf_name):
                    st.error(f"Файл {tpl_pf_name} не найден.")
                    return None

                with open(tpl_pf_name, "rb") as f:
                    wb_pf = load_workbook(io.BytesIO(f.read()), keep_vba=True)

                wb_pf.calculation.fullCalcOnLoad = True
                tpl_pf_sheet = next((s for s in wb_pf.sheetnames if s not in ("Итог", "рубкг")), None)

                for s_name in list(wb_pf.sheetnames):
                    if s_name not in ("Итог", "рубкг", tpl_pf_sheet):
                        del wb_pf[s_name]

                roster_map = {e["full_name"]: e for e in st.session_state.staff_roster}

                for emp in st.session_state.staff_roster:
                    sheet_name = emp["full_name"][:31]
                    rate_val = get_rate_for(emp["post"], emp.get("grade", ""))
                    if is_master_post(emp["post"]):
                        rate_val = calc_master_shift_pay(rate_val, 11.0)
                    if sheet_name not in wb_pf.sheetnames and tpl_pf_sheet in wb_pf.sheetnames:
                        new_ws = wb_pf.copy_worksheet(wb_pf[tpl_pf_sheet])
                        new_ws.title = sheet_name
                        for r in range(2, 33):
                            if new_ws.cell(row=r, column=2).data_type != 'f':
                                new_ws.cell(row=r, column=2).value = None
                            if new_ws.cell(row=r, column=3).data_type != 'f':
                                new_ws.cell(row=r, column=3).value = rate_val

                if tpl_pf_sheet in wb_pf.sheetnames and not any(e["full_name"][:31] == tpl_pf_sheet for e in st.session_state.staff_roster):
                    del wb_pf[tpl_pf_sheet]

                for day_str, data in st.session_state.shifts.items():
                    day = int(day_str)
                    timesheet_emps = data.get("timesheet", [])
                    graph_only_emps = data.get("graph_only", {})
                    day_hours_map = st.session_state.emp_hours.get(day_str, {})

                    for emp_full_name in timesheet_emps:
                        if emp_full_name in graph_only_emps:
                            continue
                        sheet_name = emp_full_name[:31]
                        if sheet_name in wb_pf.sheetnames:
                            wb_pf[sheet_name].cell(row=day + 1, column=2, value=1)
                            emp_obj = roster_map.get(emp_full_name)
                            if emp_obj and is_master_post(emp_obj["post"]):
                                hrs = float(day_hours_map.get(emp_full_name, 11.0))
                                base_r = get_rate_for(emp_obj["post"], emp_obj.get("grade", ""))
                                if wb_pf[sheet_name].cell(row=day + 1, column=3).data_type != 'f':
                                    wb_pf[sheet_name].cell(row=day + 1, column=3, value=calc_master_shift_pay(base_r, hrs))

                    if day_str in st.session_state.tonnage and "рубкг" in wb_pf.sheetnames:
                        wb_pf["рубкг"].cell(row=day + 1, column=6, value=st.session_state.tonnage[day_str])

                if "Итог" in wb_pf.sheetnames:
                    ws_it = wb_pf["Итог"]
                    for r in range(2, 50):
                        ws_it.cell(row=r, column=1).value = None
                        ws_it.cell(row=r, column=2).value = None
                        ws_it.cell(row=r, column=3).value = None

                    for i, emp in enumerate(st.session_state.staff_roster):
                        r = 2 + i
                        sheet_name = emp["full_name"][:31]
                        ws_it.cell(row=r, column=1, value=emp["full_name"])
                        ws_it.cell(row=r, column=2, value=f"='{sheet_name}'!J33")
                        ws_it.cell(row=r, column=3, value=f"='{sheet_name}'!B33")

                    itogo_r = 2 + len(st.session_state.staff_roster)
                    ws_it.cell(row=itogo_r, column=1, value="ИТОГО:")
                    ws_it.cell(row=itogo_r, column=2, value=f"=SUM(B2:B{itogo_r-1})")
                    ws_it.cell(row=itogo_r, column=3, value=f"=SUM(C2:C{itogo_r-1})")
                    ws_it.delete_rows(itogo_r + 1, 50)

                buf = io.BytesIO()
                wb_pf.save(buf)
                zf.writestr(f"{prefix}_{m_name}_{y_val}.xlsm", buf.getvalue())
                del wb_pf
                gc.collect()

            # 5.2. ГРАФИК (ОДИН ДЛЯ ВСЕХ с динамическим месяцем и листом)
            status_container.write("Обрабатываем График...")
            with open(tpl_gr_name, "rb") as f:
                wb_gr = load_workbook(io.BytesIO(f.read()))

            ws_gr = wb_gr.active
            ws_gr['A1'] = m_name
            sheet_title_str = f"{m_name} {y_val}"
            try:
                ws_gr.title = sheet_title_str
            except Exception:
                pass

            m_idx = MONTHS_RU.index(m_name) + 1
            num_days = calendar.monthrange(int(y_val), m_idx)[1]

            itogo_col = None
            for c_idx in range(4, 55):
                for r_idx in [1, 2, 3]:
                    val = str(ws_gr.cell(row=r_idx, column=c_idx).value or "").strip().upper()
                    if "ИТОГО" in val:
                        itogo_col = c_idx
                        break
                if itogo_col:
                    break

            if not itogo_col:
                itogo_col = 3 + num_days + 1
                ws_gr.cell(row=2, column=itogo_col, value="ИТОГО")

            block_r = 4
            for r in range(1, 20):
                for c in range(1, 4):
                    cell_txt = str(ws_gr.cell(row=r, column=c).value or "").lower()
                    if "отделение" in cell_txt or "обвалка" in cell_txt or "подготовки" in cell_txt:
                        block_r = r
                        break
                if block_r > 4:
                    break

            last_emp_row = block_r + len(st.session_state.staff_roster)
            for r_idx in range(1, last_emp_row + 1):
                ws_gr.cell(row=r_idx, column=itogo_col).fill = BLUE_FILL

            for d in range(1, 32):
                col_idx = 3 + d
                if d <= num_days:
                    ws_gr.cell(row=2, column=col_idx, value=d)
                else:
                    ws_gr.cell(row=2, column=col_idx, value=None)
                for r_idx in [1, 2, 3]:
                    ws_gr.cell(row=r_idx, column=col_idx).fill = NO_FILL

            for r_idx in range(1, last_emp_row + 1):
                for c_idx in [1, 2, 3]:
                    ws_gr.cell(row=r_idx, column=c_idx).fill = NO_FILL

            for i, emp in enumerate(st.session_state.staff_roster):
                r = block_r + 1 + i
                ws_gr.cell(row=r, column=1, value=i + 1)
                ws_gr.cell(row=r, column=2, value=emp["full_name"])
                ws_gr.cell(row=r, column=3, value=emp["post"])

                for c_idx in range(4, 50):
                    cell = ws_gr.cell(row=r, column=c_idx)
                    cell.value = None
                    cell.fill = NO_FILL
                    cell.comment = None

                shifts_count = 0
                for day_str, data in st.session_state.shifts.items():
                    day = int(day_str)
                    if day > num_days:
                        continue
                    col_idx = 3 + day
                    day_hours_map = st.session_state.emp_hours.get(day_str, {})
                    if is_master_post(emp["post"]):
                        hrs_val = float(day_hours_map.get(emp["full_name"], 11.0))
                        hours = int(hrs_val) if hrs_val.is_integer() else hrs_val
                    else:
                        hours = 8 if "грузчик" in emp["post"].lower() or "подсобный" in emp["post"].lower() else 11

                    if emp["full_name"] in data.get("graph_only", {}):
                        dept_name = data["graph_only"][emp["full_name"]]
                        cell = ws_gr.cell(row=r, column=col_idx, value=hours)
                        cell.fill = ORANGE_FILL
                        cell.comment = Comment(text=f"Подработка: {dept_name}" if dept_name else "Подработка", author="Система")
                        shifts_count += 1
                    elif emp["full_name"] in data.get("timesheet", []):
                        ws_gr.cell(row=r, column=col_idx, value=hours)
                        shifts_count += 1

                ws_gr.cell(row=r, column=itogo_col, value=shifts_count)

            status_container.write("Сохраняем График...")
            buf = io.BytesIO()
            wb_gr.save(buf)
            zf.writestr(f"ГРАФИК_{prefix}_{m_name}_{y_val}.xlsx", buf.getvalue())
            del wb_gr
            gc.collect()

            # 5.3. ПОДРАБОТЧИКИ
            status_container.write("Обрабатываем Подработчиков...")
            if dept_id in ("formovka", "chicken_deboning"):
                sub_files_map = FORMOVKA_SUBS if dept_id == "formovka" else CHICKEN_DEBONING_SUBS
                sum_cell_ref = "E33" if dept_id == "chicken_deboning" else "J33"
                sub_skip = {"Итог", "Общий"}

                for group_name, file_name in sub_files_map.items():
                    if not os.path.exists(file_name):
                        status_container.write(f"Файл {file_name} не найден, пропускаем.")
                        continue

                    group_subs = [s for s in st.session_state.subordinates if s.get("subgroup") == group_name]
                    if not group_subs:
                        continue

                    with open(file_name, "rb") as f:
                        wb_sub = load_workbook(io.BytesIO(f.read()), keep_vba=True)

                    wb_sub.calculation.fullCalcOnLoad = True
                    tpl_sub_sheet = next((s for s in wb_sub.sheetnames if s not in sub_skip), None)

                    for s in list(wb_sub.sheetnames):
                        if s not in sub_skip and s != tpl_sub_sheet:
                            del wb_sub[s]

                    subs_by_fio = {}
                    for sub in group_subs:
                        sheet_name = sub["short_name"][:31]
                        if sheet_name not in wb_sub.sheetnames and tpl_sub_sheet:
                            new_ws = wb_sub.copy_worksheet(wb_sub[tpl_sub_sheet])
                            new_ws.title = sheet_name
                            for r in range(2, 33):
                                if new_ws.cell(row=r, column=2).data_type != 'f':
                                    new_ws.cell(row=r, column=2).value = None
                                if new_ws.cell(row=r, column=3).data_type != 'f':
                                    new_ws.cell(row=r, column=3).value = None
                                if dept_id == "chicken_deboning" and new_ws.cell(row=r, column=4).data_type != 'f':
                                    new_ws.cell(row=r, column=4).value = 1

                        day = int(sub["date"].split(".")[0])
                        wb_sub[sheet_name].cell(row=day + 1, column=2, value=1)
                        wb_sub[sheet_name].cell(row=day + 1, column=3, value=sub["amount"])
                        if dept_id == "chicken_deboning" and wb_sub[sheet_name].cell(row=day + 1, column=4).data_type != 'f':
                            wb_sub[sheet_name].cell(row=day + 1, column=4, value=1)

                        if sheet_name not in subs_by_fio:
                            subs_by_fio[sheet_name] = {"amount": 0.0, "shifts": 0}
                        subs_by_fio[sheet_name]["amount"] += float(sub["amount"])
                        subs_by_fio[sheet_name]["shifts"] += 1

                    if tpl_sub_sheet in wb_sub.sheetnames:
                        del wb_sub[tpl_sub_sheet]

                    if "Итог" in wb_sub.sheetnames:
                        ws_it = wb_sub["Итог"]
                        itogo_row = 50
                        for r in range(2, 50):
                            if str(ws_it.cell(row=r, column=1).value or "").strip().upper() == "ИТОГО:":
                                itogo_row = r
                                break
                            ws_it.cell(row=r, column=1).value = None
                            ws_it.cell(row=r, column=2).value = None
                            ws_it.cell(row=r, column=3).value = None

                        for i, (fio, data) in enumerate(subs_by_fio.items()):
                            r = 2 + i
                            if r < itogo_row:
                                ws_it.cell(row=r, column=1, value=fio)
                                ws_it.cell(row=r, column=2, value=f"='{fio}'!{sum_cell_ref}")
                                ws_it.cell(row=r, column=3, value=f"='{fio}'!B33")

                        if itogo_row < 50:
                            ws_it.cell(row=itogo_row, column=2, value=f"=SUM(B2:B{itogo_row-1})")
                            ws_it.cell(row=itogo_row, column=3, value=f"=SUM(C2:C{itogo_row-1})")

                    buf = io.BytesIO()
                    wb_sub.save(buf)
                    zf.writestr(f"Подработка_{group_name}_{prefix}_{m_name}_{y_val}.xlsm", buf.getvalue())
                    del wb_sub
                    gc.collect()
            else:
                tpl_sub_name = get_template_name("tpl_sub", dept_id, "xlsm")
                if os.path.exists(tpl_sub_name):
                    with open(tpl_sub_name, "rb") as f:
                        wb_sub = load_workbook(io.BytesIO(f.read()), keep_vba=True)

                    wb_sub.calculation.fullCalcOnLoad = True
                    sub_skip = {"Итог", "Общий"}
                    tpl_sub_sheet = next((s for s in wb_sub.sheetnames if s not in sub_skip), None)

                    for s in list(wb_sub.sheetnames):
                        if s not in sub_skip and s != tpl_sub_sheet:
                            del wb_sub[s]

                    subs_by_fio = {}
                    for sub in st.session_state.subordinates:
                        sheet_name = sub["short_name"][:31]
                        if sheet_name not in wb_sub.sheetnames and tpl_sub_sheet:
                            new_ws = wb_sub.copy_worksheet(wb_sub[tpl_sub_sheet])
                            new_ws.title = sheet_name
                            for r in range(2, 33):
                                if new_ws.cell(row=r, column=2).data_type != 'f':
                                    new_ws.cell(row=r, column=2).value = None
                                if new_ws.cell(row=r, column=3).data_type != 'f':
                                    new_ws.cell(row=r, column=3).value = None
                                if dept_id == "main_deboning" and new_ws.cell(row=r, column=4).data_type != 'f':
                                    new_ws.cell(row=r, column=4).value = 1

                        day = int(sub["date"].split(".")[0])
                        wb_sub[sheet_name].cell(row=day + 1, column=2, value=1)
                        wb_sub[sheet_name].cell(row=day + 1, column=3, value=sub["amount"])
                        if dept_id == "main_deboning" and wb_sub[sheet_name].cell(row=day + 1, column=4).data_type != 'f':
                            wb_sub[sheet_name].cell(row=day + 1, column=4, value=1)

                        if sheet_name not in subs_by_fio:
                            subs_by_fio[sheet_name] = {"amount": 0.0, "shifts": 0}
                        subs_by_fio[sheet_name]["amount"] += float(sub["amount"])
                        subs_by_fio[sheet_name]["shifts"] += 1

                    if tpl_sub_sheet in wb_sub.sheetnames:
                        del wb_sub[tpl_sub_sheet]

                    sum_cell_ref = "E33" if dept_id == "main_deboning" else "J33"
                    if "Итог" in wb_sub.sheetnames:
                        ws_it = wb_sub["Итог"]
                        itogo_row = 50
                        for r in range(2, 50):
                            if str(ws_it.cell(row=r, column=1).value or "").strip().upper() == "ИТОГО:":
                                itogo_row = r
                                break
                            ws_it.cell(row=r, column=1).value = None
                            ws_it.cell(row=r, column=2).value = None
                            ws_it.cell(row=r, column=3).value = None

                        for i, (fio, data) in enumerate(subs_by_fio.items()):
                            r = 2 + i
                            if r < itogo_row:
                                ws_it.cell(row=r, column=1, value=fio)
                                ws_it.cell(row=r, column=2, value=f"='{fio}'!{sum_cell_ref}")
                                ws_it.cell(row=r, column=3, value=f"='{fio}'!B33")

                        if itogo_row < 50:
                            ws_it.cell(row=itogo_row, column=2, value=f"=SUM(B2:B{itogo_row-1})")
                            ws_it.cell(row=itogo_row, column=3, value=f"=SUM(C2:C{itogo_row-1})")

                    buf = io.BytesIO()
                    wb_sub.save(buf)
                    zf.writestr(f"Подработчики_{prefix}_{m_name}_{y_val}.xlsm", buf.getvalue())
                    del wb_sub
                    gc.collect()

            # 5.4. ВЕДОМОСТЬ ШАШИНОЙ (Маппинг по вашим точным названиям)
            status_container.write("Обрабатываем Ведомость Шашиной...")
            with open(tpl_sh_name, "rb") as f:
                wb_sh = load_workbook(io.BytesIO(f.read()))

            sh_sheet_map = {
                "pf": "пф",
                "main_deboning": "обвалка",
                "chicken_deboning": "птица",
                "offal": "ло",
                "delicatessen": "до",
                "formovka": "формовка",
                "farshesozdanie": "фарши"
            }
            target_sh_name = sh_sheet_map.get(dept_id, "пф").strip().lower()

            actual_sheet_name = None
            for s_name in wb_sh.sheetnames:
                if s_name.replace(" ", "").strip().lower() == target_sh_name.replace(" ", ""):
                    actual_sheet_name = s_name
                    break

            if actual_sheet_name:
                status_container.write(f"Записываем {len(st.session_state.subordinates)} подработчиков на лист '{actual_sheet_name}'...")
                wb_sh.calculation.fullCalcOnLoad = True
                ws_sh = wb_sh[actual_sheet_name]
                start_row = 4
                for i, sub in enumerate(st.session_state.subordinates):
                    r = start_row + i
                    ws_sh.cell(row=r, column=2, value=i + 1)
                    ws_sh.cell(row=r, column=3, value=sub["short_name"])
                    ws_sh.cell(row=r, column=4, value=sub["post"])
                    ws_sh.cell(row=r, column=5, value=sub.get("grade", ""))
                    ws_sh.cell(row=r, column=6, value=11)
                    ws_sh.cell(row=r, column=8, value=sub["date"])
                    ws_sh.cell(row=r, column=9, value=sub["amount"])
            else:
                status_container.write(f"Лист '{target_sh_name}' не найден в файле Шашиной.")

            buf = io.BytesIO()
            wb_sh.save(buf)
            zf.writestr(f"Шашина_{m_name}_{y_val}.xlsx", buf.getvalue())
            del wb_sh
            gc.collect()

        return zip_buf.getvalue()
    except Exception as e:
        st.error(f"Критическая ошибка при создании файлов: {e}")
        return None

# ==========================================
# 6. ИНТЕРФЕЙС (SIDEBAR И АВТОРИЗАЦИЯ ПО PIN)
# ==========================================
st.sidebar.markdown("### Производственный учёт")
dept_options = [""] + list(DEPARTMENTS.keys())
dept_labels = ["— Выберите отделение —"] + list(DEPARTMENTS.values())

try:
    current_idx = dept_options.index(st.session_state.current_dept)
except ValueError:
    current_idx = 0

selected_idx = st.sidebar.selectbox(
    "Подразделение",
    range(len(dept_options)),
    format_func=lambda x: dept_labels[x],
    index=current_idx
)
selected_dept = dept_options[selected_idx]

if selected_dept != st.session_state.current_dept:
    st.session_state.current_dept = selected_dept
    st.session_state.active_month = ""
    st.session_state.current_id = None
    st.session_state.shifts = {}
    st.session_state.tonnage = {}
    st.session_state.subordinates = []
    st.session_state.brigade_percents = {}
    st.session_state.emp_percents = {}
    st.session_state.emp_remarks = {}
    st.session_state.emp_remarks_2 = {}
    st.session_state.emp_hours = {}
    if selected_dept:
        load_global_settings()
    st.rerun()

if not st.session_state.current_dept:
    st.markdown("""
    <div class="nordic-header-card">
        <div>
            <div class="nordic-title">Учёт выработки и табелирование</div>
            <div class="nordic-subtitle">Выберите ваше подразделение в меню слева, чтобы начать работу</div>
        </div>
        <div class="nordic-period-badge-empty">Отделение не выбрано</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Проверка 4-значного пароля доступа в выбранное отделение
required_pin = DEPT_PASSWORDS.get(st.session_state.current_dept, "")
if required_pin and required_pin not in st.session_state.unlocked_pins:
    _, mid_col, _ = st.columns([1.2, 1.6, 1.2])
    with mid_col:
        st.markdown(f"""
        <div class="pin-card">
            <div class="pin-tag">Доступ ограничен</div>
            <div style="font-size: 20px; font-weight: 700; color: #0F172A; margin-bottom: 6px;">
                {DEPARTMENTS[st.session_state.current_dept]}
            </div>
            <div style="font-size: 13.5px; color: #64748B;">
                Введите 4-значный PIN-код доступа вашего подразделения
            </div>
        </div>
        """, unsafe_allow_html=True)
        entered_pin = st.text_input(
            "PIN-код (4 цифры)",
            type="password",
            max_chars=4,
            key=f"pin_input_{st.session_state.current_dept}",
            placeholder="••••"
        )
        if st.button("Войти в подразделение", type="primary", use_container_width=True):
            if entered_pin.strip() == required_pin:
                st.session_state.unlocked_pins.append(required_pin)
                st.rerun()
            else:
                st.error("Неверный PIN-код. Проверьте код вашего подразделения.")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.markdown("#### Период и выгрузка")
if st.session_state.active_month:
    st.sidebar.info(f"**Активный период:** {st.session_state.active_month} {st.session_state.active_year}")

    if st.sidebar.button("Сформировать отчеты (ZIP)", type="primary", use_container_width=True):
        with st.status("Формируем Excel-файлы...", expanded=True) as status:
            zip_data = generate_and_download(status)
            if zip_data:
                status.update(label="Файлы успешно сформированы", state="complete", expanded=False)
                zip_prefix = FILE_PREFIXES.get(st.session_state.current_dept, "Отчет")
                st.sidebar.download_button(
                    label="Скачать ZIP-архив",
                    data=zip_data,
                    file_name=f"Итоги_{zip_prefix}_{st.session_state.active_month}_{st.session_state.active_year}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            else:
                status.update(label="Ошибка при создании файлов", state="error", expanded=True)
else:
    st.sidebar.caption("Месяц не выбран — создайте новый или откройте из архива ниже.")

st.sidebar.markdown("---")
if supabase:
    st.sidebar.markdown("#### Архив месяцев")
    try:
        dept_id = st.session_state.current_dept
        res = supabase.table("app_data").select("id").neq("id", f"GLOBAL_SETTINGS_{dept_id}").execute()
        saved = [r["id"] for r in res.data if r["id"].endswith(f"_{dept_id}")]
        if saved:
            saved.sort(reverse=True)
            sel = st.sidebar.selectbox(
                "Сохраненные периоды",
                saved,
                format_func=lambda x: x.replace(f"_{dept_id}", "").replace("_", " "),
                label_visibility="collapsed"
            )
            if st.sidebar.button("Открыть период", use_container_width=True):
                if load_month_from_db(sel):
                    st.sidebar.success("Период загружен")
                    st.rerun()
        else:
            st.sidebar.caption("Нет сохраненных месяцев.")
    except Exception:
        pass

period_badge_html = (
    f'<div class="nordic-period-badge">{st.session_state.active_month} {st.session_state.active_year}</div>'
    if st.session_state.active_month
    else '<div class="nordic-period-badge-empty">Период не выбран</div>'
)
st.markdown(f"""
<div class="nordic-header-card">
    <div>
        <div class="nordic-title">{DEPARTMENTS[st.session_state.current_dept]}</div>
        <div class="nordic-subtitle">Табелирование смен, расчёт выработки и формирование ведомостей</div>
    </div>
    {period_badge_html}
</div>
""", unsafe_allow_html=True)

tab0, tab1, tab2, tab3, tab4 = st.tabs(["Создать месяц", "Табель смен", "Подработчики", "Ставки", "Кадры"])

# ================= TAB 0: СОЗДАНИЕ =================
with tab0:
    st.subheader("Открытие нового расчетного периода")
    c1, c2 = st.columns(2)
    with c1:
        new_month = st.selectbox("Месяц", MONTHS_RU, index=datetime.date.today().month - 1)
    with c2:
        years_list = [str(y) for y in range(2024, 2036)]
        new_year = st.selectbox("Год", years_list, index=years_list.index(str(datetime.date.today().year)))

    if st.button("Создать расчетный месяц", type="primary"):
        dept_id = st.session_state.current_dept
        record_id = f"{new_month}_{new_year}_{dept_id}"
        if not supabase:
            st.error("Нет подключения к БД.")
        else:
            try:
                existing = supabase.table("app_data").select("id").eq("id", record_id).execute()
                if existing.data:
                    st.error(f"Месяц {new_month} {new_year} для этого отдела уже существует.")
                else:
                    st.session_state.shifts = {}
                    st.session_state.tonnage = {}
                    st.session_state.subordinates = []
                    st.session_state.brigade_percents = {}
                    st.session_state.emp_percents = {}
                    st.session_state.emp_remarks = {}
                    st.session_state.emp_remarks_2 = {}
                    st.session_state.emp_hours = {}
                    st.session_state.active_month = new_month
                    st.session_state.active_year = str(new_year)
                    st.session_state.current_id = record_id

                    empty_payload = json.dumps({
                        "shifts": {},
                        "tonnage": {},
                        "subordinates": [],
                        "brigade_percents": {},
                        "emp_percents": {},
                        "emp_remarks": {},
                        "emp_remarks_2": {},
                        "emp_hours": {}
                    })
                    data = {"id": record_id, "pf_base64": base64.b64encode(empty_payload.encode('utf-8')).decode("utf-8")}
                    supabase.table("app_data").insert(data).execute()

                    st.success(f"Месяц {new_month} {new_year} для отдела «{DEPARTMENTS[dept_id]}» создан.")
                    st.rerun()
            except Exception as e:
                st.error(f"Ошибка: {e}")

# ================= TAB 1: ТАБЕЛЬ =================
with tab1:
    if not st.session_state.active_month:
        st.info("Сначала создайте новый месяц или откройте существующий в левой панели.")
    else:
        cur_dept = st.session_state.current_dept
        is_deb = is_deboning_dept(cur_dept)
        is_chicken = (cur_dept == "chicken_deboning")

        if is_deb:
            c1, c2, c3, c4 = st.columns(4, vertical_alignment="bottom")
            with c1:
                shift_date = st.date_input("Дата смены", datetime.date.today())
            day_key_preview = str(shift_date.day)
            saved_b_pct = st.session_state.brigade_percents.get(day_key_preview, {})
            with c2:
                pct_obvalka = st.number_input(
                    "Процент Обвалки (%)",
                    min_value=0.0,
                    max_value=300.0,
                    value=float(saved_b_pct.get("obvalka", 100.0)),
                    step=1.0
                )
            with c3:
                pct_zhilovka = st.number_input(
                    "Процент Жиловки (%)",
                    min_value=0.0,
                    max_value=300.0,
                    value=float(saved_b_pct.get("zhilovka", 100.0)),
                    step=1.0
                )
            with c4:
                tonnage = st.number_input(
                    "Тоннаж цеха (кг)",
                    min_value=0.0,
                    value=float(st.session_state.tonnage.get(day_key_preview, 0.0)),
                    step=1.0
                )
        else:
            c1, c2 = st.columns(2, vertical_alignment="bottom")
            with c1:
                shift_date = st.date_input("Дата смены", datetime.date.today())
            day_key_preview = str(shift_date.day)
            with c2:
                tonnage = st.number_input(
                    "Тоннаж цеха (кг)",
                    min_value=0.0,
                    value=float(st.session_state.tonnage.get(day_key_preview, 0.0)),
                    step=1.0
                )

        day_saved_shifts = st.session_state.shifts.get(day_key_preview, {})
        day_saved_ts = day_saved_shifts.get("timesheet", [])
        day_saved_go = day_saved_shifts.get("graph_only", {})
        if isinstance(day_saved_go, list):
            day_saved_go = {n: "" for n in day_saved_go}
        day_saved_emp_pct = st.session_state.emp_percents.get(day_key_preview, {})
        day_saved_emp_rem = st.session_state.emp_remarks.get(day_key_preview, {})
        day_saved_emp_rem2 = st.session_state.emp_remarks_2.get(day_key_preview, {})
        day_saved_emp_hrs = st.session_state.emp_hours.get(day_key_preview, {})

        active_main_cnt = len([x for x in day_saved_ts if x not in day_saved_go])
        active_sub_cnt = len(day_saved_go)
        month_total_tonnage = sum(float(v) for v in st.session_state.tonnage.values())
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("В основной смене", f"{active_main_cnt} чел.")
        m2.metric("На подработке в др. цехах", f"{active_sub_cnt} чел.")
        m3.metric("Тоннаж за смену", f"{tonnage:,.0f} кг".replace(",", " "))
        m4.metric("Тоннаж за месяц", f"{month_total_tonnage:,.0f} кг".replace(",", " "))

        st.markdown("")
        f_col1, f_col2 = st.columns([2, 3], vertical_alignment="center")
        with f_col1:
            search_fio = st.text_input("Поиск сотрудника по фамилии", placeholder="Поиск по фамилии...", label_visibility="collapsed")
        with f_col2:
            st.caption("«Смена» — сотрудник попадает в табель и график (для Мастера указываются часы от ставки за 11ч). «Подработка» — только в график.")

        st.markdown("")
        if is_chicken:
            h1, h2, h3, h4, h5, h6, h7 = st.columns([2.8, 0.7, 1.35, 1.15, 1.15, 1.15, 2.0], vertical_alignment="center")
            h1.caption("СОТРУДНИК")
            h2.caption("СМЕНА")
            h3.caption("ПРОЦЕНТ / ЧАСЫ")
            h4.caption("ЗАМ. (КОСТИ)")
            h5.caption("ЗАМ. (ХРЯЩИ)")
            h6.caption("ПОДРАБОТКА")
            h7.caption("ЦЕХ ПОДРАБОТКИ")
        elif cur_dept == "main_deboning":
            h1, h2, h3, h4, h5, h6 = st.columns([3.2, 0.8, 1.45, 1.25, 1.25, 2.35], vertical_alignment="center")
            h1.caption("СОТРУДНИК")
            h2.caption("СМЕНА")
            h3.caption("ПРОЦЕНТ / ЧАСЫ")
            h4.caption("ЗАМЕЧАНИЯ")
            h5.caption("ПОДРАБОТКА")
            h6.caption("ЦЕХ ПОДРАБОТКИ")
        else:
            h1, h2, h3, h4, h5 = st.columns([3.4, 0.9, 1.5, 1.4, 2.8], vertical_alignment="center")
            h1.caption("СОТРУДНИК")
            h2.caption("СМЕНА")
            h3.caption("ЧАСЫ (МАСТЕР)")
            h4.caption("ПОДРАБОТКА")
            h5.caption("ЦЕХ ПОДРАБОТКИ")

        selected = {}
        last_subgroup = None

        for emp in st.session_state.staff_roster:
            if search_fio.strip() and search_fio.strip().lower() not in emp["full_name"].lower():
                continue

            cur_sub = emp.get("subgroup")
            if cur_sub and cur_sub != last_subgroup:
                st.markdown(f'<div class="subgroup-banner">{cur_sub}</div>', unsafe_allow_html=True)
                last_subgroup = cur_sub

            emp_is_master = is_master_post(emp["post"])
            default_main = emp["full_name"] in day_saved_ts
            default_other = emp["full_name"] in day_saved_go
            default_dept = day_saved_go.get(emp["full_name"], "")
            default_hrs = float(day_saved_emp_hrs.get(emp["full_name"], 11.0))

            status_badge = ""
            if default_other:
                status_badge = '<span class="badge-sub">Подработка</span>'
            elif default_main:
                status_badge = '<span class="badge-shift">В смене</span>'

            grade_txt = f" ({emp['grade']})" if emp.get('grade') and emp['grade'] != "—" else ""
            role_hint = f"{emp['post']}{grade_txt}"
            if emp_is_master:
                role_hint += " · часы"
            emp_html = f'<div class="emp-label">{emp["short_name"]}{status_badge}<div class="emp-post">{role_hint}</div></div>'

            if is_deb:
                if emp["full_name"] in day_saved_emp_pct:
                    init_pct = float(day_saved_emp_pct[emp["full_name"]])
                elif emp.get("pay_type") == "По коэффициенту бригады":
                    init_pct = pct_obvalka if emp.get("subgroup") == "Обвалка" else (
                        pct_zhilovka if emp.get("subgroup") == "Жиловка" else 100.0
                    )
                else:
                    init_pct = 100.0

                if is_chicken:
                    col1, col2, col3, col4, col5, col6, col7 = st.columns([2.8, 0.7, 1.35, 1.15, 1.15, 1.15, 2.0], vertical_alignment="center")
                    with col1:
                        st.markdown(emp_html, unsafe_allow_html=True)
                    with col2:
                        is_main = st.checkbox(
                            "Смена",
                            value=default_main,
                            key=f"main_{day_key_preview}_{emp['full_name']}",
                            label_visibility="collapsed"
                        )
                    with col3:
                        if emp_is_master:
                            emp_hrs_val = st.number_input(
                                "Часы",
                                min_value=0.0,
                                max_value=24.0,
                                value=default_hrs,
                                step=0.5,
                                key=f"hrs_{day_key_preview}_{emp['full_name']}",
                                label_visibility="collapsed"
                            )
                            emp_pct_val = 100.0
                        else:
                            emp_pct_val = st.number_input(
                                "%",
                                min_value=0.0,
                                max_value=300.0,
                                value=float(init_pct),
                                step=1.0,
                                key=f"pct_{day_key_preview}_{emp['full_name']}",
                                label_visibility="collapsed"
                            )
                            emp_hrs_val = 11.0
                    with col4:
                        emp_rem_val = st.number_input(
                            "Кости",
                            min_value=0,
                            max_value=50,
                            value=int(day_saved_emp_rem.get(emp["full_name"], 0)),
                            step=1,
                            key=f"rem_{day_key_preview}_{emp['full_name']}",
                            label_visibility="collapsed"
                        )
                    with col5:
                        emp_rem_val_2 = st.number_input(
                            "Хрящи",
                            min_value=0,
                            max_value=50,
                            value=int(day_saved_emp_rem2.get(emp["full_name"], 0)),
                            step=1,
                            key=f"rem2_{day_key_preview}_{emp['full_name']}",
                            label_visibility="collapsed"
                        )
                    with col6:
                        is_other = st.checkbox(
                            "Подработка",
                            value=default_other,
                            key=f"other_{day_key_preview}_{emp['full_name']}"
                        )
                    with col7:
                        dept_name = st.text_input(
                            "Цех",
                            value=default_dept,
                            key=f"dept_{day_key_preview}_{emp['full_name']}",
                            label_visibility="collapsed",
                            placeholder="Куда ушел..."
                        )
                    selected[emp["full_name"]] = {
                        "timesheet": is_main,
                        "graph_only": is_other,
                        "dept": dept_name,
                        "pct": emp_pct_val,
                        "remarks": emp_rem_val,
                        "remarks_2": emp_rem_val_2,
                        "hours": emp_hrs_val,
                        "is_master": emp_is_master
                    }
                else:
                    col1, col2, col3, col4, col5, col6 = st.columns([3.2, 0.8, 1.45, 1.25, 1.25, 2.35], vertical_alignment="center")
                    with col1:
                        st.markdown(emp_html, unsafe_allow_html=True)
                    with col2:
                        is_main = st.checkbox(
                            "Смена",
                            value=default_main,
                            key=f"main_{day_key_preview}_{emp['full_name']}",
                            label_visibility="collapsed"
                        )
                    with col3:
                        if emp_is_master:
                            emp_hrs_val = st.number_input(
                                "Часы",
                                min_value=0.0,
                                max_value=24.0,
                                value=default_hrs,
                                step=0.5,
                                key=f"hrs_{day_key_preview}_{emp['full_name']}",
                                label_visibility="collapsed"
                            )
                            emp_pct_val = 100.0
                        else:
                            emp_pct_val = st.number_input(
                                "%",
                                min_value=0.0,
                                max_value=300.0,
                                value=float(init_pct),
                                step=1.0,
                                key=f"pct_{day_key_preview}_{emp['full_name']}",
                                label_visibility="collapsed"
                            )
                            emp_hrs_val = 11.0
                    with col4:
                        emp_rem_val = st.number_input(
                            "Зам.",
                            min_value=0,
                            max_value=50,
                            value=int(day_saved_emp_rem.get(emp["full_name"], 0)),
                            step=1,
                            key=f"rem_{day_key_preview}_{emp['full_name']}",
                            label_visibility="collapsed"
                        )
                    with col5:
                        is_other = st.checkbox(
                            "Подработка",
                            value=default_other,
                            key=f"other_{day_key_preview}_{emp['full_name']}"
                        )
                    with col6:
                        dept_name = st.text_input(
                            "Цех",
                            value=default_dept,
                            key=f"dept_{day_key_preview}_{emp['full_name']}",
                            label_visibility="collapsed",
                            placeholder="Куда ушел..."
                        )
                    selected[emp["full_name"]] = {
                        "timesheet": is_main,
                        "graph_only": is_other,
                        "dept": dept_name,
                        "pct": emp_pct_val,
                        "remarks": emp_rem_val,
                        "remarks_2": 0,
                        "hours": emp_hrs_val,
                        "is_master": emp_is_master
                    }
            else:
                col1, col2, col3, col4, col5 = st.columns([3.4, 0.9, 1.5, 1.4, 2.8], vertical_alignment="center")
                with col1:
                    st.markdown(emp_html, unsafe_allow_html=True)
                with col2:
                    is_main = st.checkbox(
                        "Смена",
                        value=default_main,
                        key=f"main_{day_key_preview}_{emp['full_name']}",
                        label_visibility="collapsed"
                    )
                with col3:
                    if emp_is_master:
                        emp_hrs_val = st.number_input(
                            "Часы",
                            min_value=0.0,
                            max_value=24.0,
                            value=default_hrs,
                            step=0.5,
                            key=f"hrs_{day_key_preview}_{emp['full_name']}",
                            label_visibility="collapsed"
                        )
                    else:
                        st.markdown('<div class="empty-cell-dash">—</div>', unsafe_allow_html=True)
                        emp_hrs_val = 11.0
                with col4:
                    is_other = st.checkbox(
                        "Подработка",
                        value=default_other,
                        key=f"other_{day_key_preview}_{emp['full_name']}"
                    )
                with col5:
                    dept_name = st.text_input(
                        "Цех",
                        value=default_dept,
                        key=f"dept_{day_key_preview}_{emp['full_name']}",
                        label_visibility="collapsed",
                        placeholder="Куда ушел..."
                    )

                selected[emp["full_name"]] = {
                    "timesheet": is_main,
                    "graph_only": is_other,
                    "dept": dept_name,
                    "hours": emp_hrs_val,
                    "is_master": emp_is_master
                }

            st.markdown('<div class="row-divider"></div>', unsafe_allow_html=True)

        st.markdown("")
        if st.button("Провести и сохранить смену", type="primary"):
            day_str = str(shift_date.day)
            if day_str not in st.session_state.shifts:
                st.session_state.shifts[day_str] = {"timesheet": [], "graph_only": {}}
            if "graph_only" not in st.session_state.shifts[day_str]:
                st.session_state.shifts[day_str]["graph_only"] = {}
            if isinstance(st.session_state.shifts[day_str]["graph_only"], list):
                st.session_state.shifts[day_str]["graph_only"] = {n: "" for n in st.session_state.shifts[day_str]["graph_only"]}

            if day_str not in st.session_state.emp_hours:
                st.session_state.emp_hours[day_str] = {}

            if is_deb:
                st.session_state.brigade_percents[day_str] = {
                    "obvalka": float(pct_obvalka),
                    "zhilovka": float(pct_zhilovka)
                }
                if day_str not in st.session_state.emp_percents:
                    st.session_state.emp_percents[day_str] = {}
                if day_str not in st.session_state.emp_remarks:
                    st.session_state.emp_remarks[day_str] = {}
                if day_str not in st.session_state.emp_remarks_2:
                    st.session_state.emp_remarks_2[day_str] = {}

            for emp_full_name, checks in selected.items():
                if checks["timesheet"]:
                    if emp_full_name not in st.session_state.shifts[day_str]["timesheet"]:
                        st.session_state.shifts[day_str]["timesheet"].append(emp_full_name)
                else:
                    if emp_full_name in st.session_state.shifts[day_str]["timesheet"]:
                        st.session_state.shifts[day_str]["timesheet"].remove(emp_full_name)

                if checks["graph_only"]:
                    st.session_state.shifts[day_str]["graph_only"][emp_full_name] = checks["dept"]
                else:
                    if emp_full_name in st.session_state.shifts[day_str]["graph_only"]:
                        del st.session_state.shifts[day_str]["graph_only"][emp_full_name]

                if checks.get("is_master"):
                    st.session_state.emp_hours[day_str][emp_full_name] = float(checks["hours"])

                if is_deb:
                    st.session_state.emp_percents[day_str][emp_full_name] = float(checks["pct"])
                    st.session_state.emp_remarks[day_str][emp_full_name] = int(checks["remarks"])
                    st.session_state.emp_remarks_2[day_str][emp_full_name] = int(checks.get("remarks_2", 0))

            if tonnage > 0:
                st.session_state.tonnage[day_str] = float(tonnage)
            save_month_to_db()
            st.success(f"Смена на {shift_date.strftime('%d.%m.%Y')} сохранена.")
            st.rerun()

        st.markdown("---")
        st.markdown("#### Журнал проведенных смен за месяц")
        roster_lookup = {s["full_name"]: s for s in st.session_state.staff_roster}
        has_any_log = False
        for day, data in sorted(st.session_state.shifts.items(), key=lambda x: int(x[0])):
            go_dict = data.get("graph_only", {})
            if isinstance(go_dict, list):
                go_dict = {name: "" for name in go_dict}

            timesheet_names = []
            for e in data.get("timesheet", []):
                if e in go_dict:
                    continue
                emp_info = roster_lookup.get(e, {})
                short_n = emp_info.get("short_name", e)
                extra = []
                if is_master_post(emp_info.get("post", "")):
                    hrs_v = st.session_state.emp_hours.get(day, {}).get(e, 11.0)
                    extra.append(f"{hrs_v:g} ч.")
                elif is_deb:
                    pct_val = st.session_state.emp_percents.get(day, {}).get(e)
                    if pct_val is not None:
                        extra.append(f"{pct_val:g}%")
                if is_deb:
                    rem_val = st.session_state.emp_remarks.get(day, {}).get(e, 0)
                    rem_val2 = st.session_state.emp_remarks_2.get(day, {}).get(e, 0)
                    if is_chicken:
                        if rem_val:
                            extra.append(f"кости: {rem_val}")
                        if rem_val2:
                            extra.append(f"хрящи: {rem_val2}")
                    else:
                        if rem_val:
                            extra.append(f"зам: {rem_val}")
                if extra:
                    short_n += f" ({', '.join(extra)})"
                timesheet_names.append(short_n)

            graph_names = []
            for name, dept in go_dict.items():
                short_n = roster_lookup.get(name, {}).get("short_name", name)
                graph_names.append(f"{short_n}{f' (в {dept})' if dept else ''}")

            if not timesheet_names and not graph_names:
                continue
            has_any_log = True

            header_tag = f"<b>{day} число</b>"
            if is_deb and day in st.session_state.brigade_percents:
                bp = st.session_state.brigade_percents[day]
                header_tag += f" &nbsp;<span class='badge-shift'>Обв {bp.get('obvalka', 100):g}% · Жил {bp.get('zhilovka', 100):g}%</span>"
            if day in st.session_state.tonnage and st.session_state.tonnage[day] > 0:
                header_tag += f" &nbsp;<span style='color:#64748B; font-size:12px;'>({st.session_state.tonnage[day]:g} кг)</span>"

            lines = [header_tag]
            if timesheet_names:
                lines.append(f"<span style='color:#1E293B;'><b>Основная:</b> {', '.join(timesheet_names)}</span>")
            if graph_names:
                lines.append(f"<span style='color:#C2410C;'><b>Подработка:</b> {', '.join(graph_names)}</span>")

            st.markdown(f'<div class="log-card">{"<br>".join(lines)}</div>', unsafe_allow_html=True)

        if not has_any_log:
            st.caption("В этом месяце еще не проведено ни одной смены.")

# ================= TAB 2: ПОДРАБОТЧИКИ =================
with tab2:
    if not st.session_state.active_month:
        st.info("Сначала создайте месяц.")
    else:
        st.subheader("Внесение подработки (внешние сотрудники / пофамильно)")
        c1, c2 = st.columns(2)

        subgroup_val = None
        with c1:
            fio = st.text_input("ФИО подработчика")
            post = st.selectbox("Должность", get_current_positions())
            available_grades = ["—"] + [r["grade"] for r in st.session_state.rates if r["post"] == post and r["grade"]]
            grade = st.selectbox("Разряд", available_grades) if len(available_grades) > 1 else "—"

            if st.session_state.current_dept == "formovka":
                subgroup_val = st.selectbox("Куда вышел? (Подгруппа)", ["Грузчики", "Клипсатор", "Сосисочная линия"])
            elif st.session_state.current_dept == "chicken_deboning":
                subgroup_val = st.selectbox("Куда вышел? (Подгруппа)", ["Жиловка", "Грузчики"])

        with c2:
            sub_date = st.date_input("Дата выхода")
            default_sub_amt = 4300.0 if (st.session_state.current_dept == "chicken_deboning" and subgroup_val == "Грузчики") else (
                4440.0 if (st.session_state.current_dept == "chicken_deboning" and subgroup_val == "Жиловка") else 4500.0
            )
            amount = st.number_input("Сумма к выплате (руб)", min_value=0.0, value=default_sub_amt, step=100.0)

        if st.button("Добавить подработку", type="primary"):
            if not fio.strip():
                st.error("Укажите ФИО")
            else:
                new_sub = {"short_name": fio.strip(), "post": post, "grade": grade, "date": sub_date.strftime("%d.%m.%Y"), "amount": float(amount)}
                if subgroup_val:
                    new_sub["subgroup"] = subgroup_val
                st.session_state.subordinates.append(new_sub)
                save_month_to_db()
                st.success("Подработка добавлена.")
                st.rerun()

        st.markdown("---")
        if st.session_state.subordinates:
            df_sub = pd.DataFrame(st.session_state.subordinates).rename(columns={
                "short_name": "ФИО",
                "subgroup": "Подгруппа",
                "post": "Должность",
                "grade": "Разряд",
                "date": "Дата",
                "amount": "Сумма (₽)"
            })
            cols = ["ФИО", "Должность", "Разряд", "Дата", "Сумма (₽)"]
            if "Подгруппа" in df_sub.columns:
                cols.insert(1, "Подгруппа")
            if not df_sub.empty:
                st.dataframe(df_sub[cols], use_container_width=True, hide_index=True)
        else:
            st.caption("Пока нет записей по внешним подработчикам.")

# ================= TAB 3: СТАВКИ =================
with tab3:
    st.subheader(f"Справочник ставок — {DEPARTMENTS[st.session_state.current_dept]}")
    st.caption("Для должности «Мастер» указывается ставка за стандартную 11-часовую смену (например, 7370 ₽) — при вводе фактических часов в табеле сумма пересчитается пропорционально (ставка / 11 × часы).")

    if not st.session_state.rates:
        st.session_state.rates = [{"post": "", "grade": "", "rate": 0}]

    df_rates = pd.DataFrame(st.session_state.rates)
    edited_rates = st.data_editor(
        df_rates, num_rows="dynamic", use_container_width=True,
        column_config={"post": st.column_config.TextColumn("Должность", required=True), "grade": st.column_config.TextColumn("Разряд"), "rate": st.column_config.NumberColumn("Ставка (руб)", min_value=0, step=10, format="%d ₽")},
        key="rates_editor"
    )
    if st.button("Сохранить ставки", type="primary"):
        new_rates = []
        for _, row in edited_rates.iterrows():
            if pd.notna(row["post"]) and str(row["post"]).strip() != "":
                new_rates.append({"post": str(row["post"]).strip(), "grade": str(row["grade"]).strip() if pd.notna(row["grade"]) else "", "rate": int(row["rate"]) if pd.notna(row["rate"]) else 0})
        st.session_state.rates = new_rates
        save_global_settings()
        st.success("Ставки сохранены.")
        st.rerun()

# ================= TAB 4: КАДРЫ =================
with tab4:
    st.subheader(f"Штатное расписание — {DEPARTMENTS[st.session_state.current_dept]}")
    if st.session_state.staff_roster:
        all_grades = set()
        for r in st.session_state.rates:
            if r.get("grade") and str(r["grade"]).strip():
                all_grades.add(str(r["grade"]).strip())
        grade_options = ["—"] + sorted(list(all_grades))

        df_staff = pd.DataFrame(st.session_state.staff_roster).rename(columns={
            "full_name": "Полное ФИО",
            "short_name": "Короткое ФИО",
            "post": "Должность",
            "grade": "Разряд",
            "subgroup": "Подгруппа",
            "pay_type": "Тип оплаты"
        })

        if is_deboning_dept(st.session_state.current_dept) and "Тип оплаты" not in df_staff.columns:
            df_staff["Тип оплаты"] = "По коэффициенту бригады"

        column_config = {
            "Полное ФИО": st.column_config.TextColumn(required=True),
            "Короткое ФИО": st.column_config.TextColumn(required=True),
            "Должность": st.column_config.SelectboxColumn(options=get_current_positions(), required=True),
            "Разряд": st.column_config.SelectboxColumn(options=grade_options)
        }

        if st.session_state.current_dept == "formovka":
            column_config["Подгруппа"] = st.column_config.SelectboxColumn(options=FORMOVKA_SUBGROUPS, required=True)
        elif st.session_state.current_dept == "main_deboning":
            column_config["Подгруппа"] = st.column_config.SelectboxColumn(options=MAIN_DEBONING_SUBGROUPS, required=True)
            column_config["Тип оплаты"] = st.column_config.SelectboxColumn(options=PAY_TYPES_DEBONING, required=True)
        elif st.session_state.current_dept == "chicken_deboning":
            column_config["Подгруппа"] = st.column_config.SelectboxColumn(options=CHICKEN_DEBONING_SUBGROUPS, required=True)
            column_config["Тип оплаты"] = st.column_config.SelectboxColumn(options=PAY_TYPES_DEBONING, required=True)

        edited_df = st.data_editor(
            df_staff, num_rows="dynamic", use_container_width=True,
            column_config=column_config,
            key="staff_editor"
        )
        if st.button("Сохранить штат", type="primary"):
            new_roster = []
            for _, row in edited_df.iterrows():
                if pd.notna(row["Полное ФИО"]) and str(row["Полное ФИО"]).strip() != "":
                    new_emp = {
                        "full_name": str(row["Полное ФИО"]).strip(),
                        "short_name": str(row["Короткое ФИО"]).strip() if pd.notna(row["Короткое ФИО"]) else make_short_name(str(row["Полное ФИО"])),
                        "post": str(row["Должность"]).strip() if pd.notna(row["Должность"]) else get_current_positions()[0],
                        "grade": str(row["Разряд"]).strip() if pd.notna(row["Разряд"]) else "—"
                    }
                    if st.session_state.current_dept == "formovka" and "Подгруппа" in row and pd.notna(row["Подгруппа"]):
                        new_emp["subgroup"] = str(row["Подгруппа"])
                    elif is_deboning_dept(st.session_state.current_dept):
                        new_emp["subgroup"] = str(row["Подгруппа"]) if "Подгруппа" in row and pd.notna(row["Подгруппа"]) else "Обвалка"
                        new_emp["pay_type"] = str(row["Тип оплаты"]) if "Тип оплаты" in row and pd.notna(row["Тип оплаты"]) else "По коэффициенту бригады"
                    new_roster.append(new_emp)
            st.session_state.staff_roster = new_roster
            save_global_settings()
            st.success("Штат сохранен.")
            st.rerun()
    else:
        st.write("Штат пуст. Нажмите кнопку ниже, чтобы загрузить из шаблона.")
        if st.button("Загрузить сотрудников из шаблона"):
            st.session_state.staff_roster = load_initial_staff(st.session_state.current_dept)
            save_global_settings()
            st.rerun()
