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
# 🎨 ДИЗАЙН
# ==========================================
st.set_page_config(page_title="Учёт выработки", page_icon="🏭", layout="wide")

custom_css = """
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stTabs [data-baseweb="tab-list"] { gap: 30px; background-color: transparent; border-bottom: 2px solid #e5e7eb; padding-bottom: 0px; }
    .stTabs [data-baseweb="tab"] { height: 50px; background-color: transparent; border: none; font-weight: 500; font-size: 15px; }
    .stTabs [aria-selected="true"] { color: #0055ff !important; border-bottom: 3px solid #0055ff !important; font-weight: 600; }
    .stButton>button[kind="primary"] { background-color: #0055ff; color: white; border: none; border-radius: 6px; box-shadow: 0 2px 4px rgba(0, 85, 255, 0.2); }
    .stButton>button[kind="primary"]:hover { background-color: #0044cc; color: white; }
    .stButton>button { border-radius: 6px; border: 1px solid #d1d5db; }
    [data-testid="stDataFrame"] { border-radius: 8px; border: 1px solid #e5e7eb; background-color: #ffffff; }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==========================================
# 1. СЛОВАРИ ОТДЕЛОВ И ДОЛЖНОСТЕЙ
# ==========================================
DEPARTMENTS = {
    "pf": "Отделение Полуфабрикатов",
    "main_deboning": "Основная обвалка",
    "chicken_deboning": "Куриная обвалка",
    "thermal": "Термическое отделение",
    "offal": "Ливерное отделение",
    "delicatessen": "Деликатесное отделение"
}

FILE_PREFIXES = {
    "pf": "Полуфабрикаты",
    "main_deboning": "Основная_обвалка",
    "chicken_deboning": "Куриная_обвалка",
    "thermal": "Термичка",
    "offal": "Ливерное",
    "delicatessen": "Деликатесы"
}

POSITIONS_BY_DEPT = {
    "pf": ["Изготовитель мясных п/ф", "Весовщик", "Грузчик", "Формовщик", "Жиловщик", "Оператор линии", "Бригадир"],
    "main_deboning": ["Обвальщик", "Жиловщик", "Весовщик", "Грузчик", "Бригадир"],
    "chicken_deboning": ["Обвальщик курицы", "Жиловщик", "Весовщик", "Грузчик"],
    "thermal": ["Термист", "Оператор", "Весовщик", "Грузчик"],
    "offal": ["Обработчик субпродуктов", "Весовщик", "Грузчик"],
    "delicatessen": ["Изготовитель деликатесов", "Формовщик", "Весовщик", "Грузчик"]
}

DEFAULT_RATES_BY_DEPT = {
    "pf": [
        {"post": "Весовщик", "grade": "", "rate": 5000},
        {"post": "Грузчик", "grade": "", "rate": 4300},
        {"post": "Изготовитель мясных п/ф", "grade": "1 разряд", "rate": 3700},
        {"post": "Изготовитель мясных п/ф", "grade": "2 разряд", "rate": 4500},
    ],
    "main_deboning": [
        {"post": "Обвальщик", "grade": "1 разряд", "rate": 5500},
        {"post": "Обвальщик", "grade": "2 разряд", "rate": 4800},
        {"post": "Грузчик", "grade": "", "rate": 4300},
    ],
    "delicatessen": [
        {"post": "Изготовитель деликатесов", "grade": "1 разряд", "rate": 5200},
        {"post": "Изготовитель деликатесов", "grade": "2 разряд", "rate": 4600},
        {"post": "Формовщик", "grade": "", "rate": 4300},
        {"post": "Весовщик", "grade": "", "rate": 4000},
        {"post": "Грузчик", "grade": "", "rate": 3800},
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
    "active_month": "", "active_year": "", "current_id": None,
    "shifts": {}, 
    "tonnage": {}, "subordinates": [], 
    "staff_roster": [], "rates": []
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ==========================================
# 3. ХЕЛПЕРЫ
# ==========================================
def get_current_positions():
    positions = []
    for r in st.session_state.rates:
        if r.get("post") and str(r["post"]).strip() not in positions:
            positions.append(str(r["post"]).strip())
    if not positions:
        positions = POSITIONS_BY_DEPT.get(st.session_state.current_dept, ["Сотрудник"])
    return positions

def make_short_name(full_name):
    parts = str(full_name).strip().split()
    return f"{parts[0]} {parts[1][0]}." if len(parts) > 1 else (parts[0] if parts else "")

def get_rate_for(post, grade):
    for r in st.session_state.rates:
        if r["post"] == post and r["grade"] == grade: return r["rate"]
    for r in st.session_state.rates:
        if r["post"] == post: return r["rate"]
    return 4500

def get_template_name(base_name, dept_id, ext):
    return f"{base_name}_{dept_id}.{ext}"

def load_initial_staff(dept_id):
    tpl_name = get_template_name("tpl_tabel", dept_id, "xlsm")
    if not os.path.exists(tpl_name): return []
    with open(tpl_name, "rb") as f:
        wb = load_workbook(io.BytesIO(f.read()), keep_vba=True, read_only=True)
    res = []
    for sname in wb.sheetnames:
        if sname in ("Итог", "рубкг") or sname.startswith("[Уволен]"): continue
        low = sname.lower()
        if "рузиева" in low: post = "Весовщик"
        elif "насруддинов" in low or "рузиев и" in low: post = "Грузчик"
        elif dept_id == "pf": post = "Изготовитель мясных п/ф"
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
    if not supabase or not st.session_state.current_dept: return
    dept_id = st.session_state.current_dept
    rec_id = f"GLOBAL_SETTINGS_{dept_id}"
    try:
        payload = {"staff_roster": st.session_state.staff_roster, "rates": st.session_state.rates}
        json_str = json.dumps(payload, ensure_ascii=False)
        data = {"pf_base64": base64.b64encode(json_str.encode('utf-8')).decode("utf-8")}
        res = supabase.table("app_data").select("id").eq("id", rec_id).execute()
        if res.data: supabase.table("app_data").update(data).eq("id", rec_id).execute()
        else:
            data["id"] = rec_id
            supabase.table("app_data").insert(data).execute()
    except Exception as e: st.error(f"Ошибка сохранения настроек: {e}")

def load_global_settings():
    if not supabase or not st.session_state.current_dept: return
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
    except Exception as e: st.error(f"Ошибка загрузки настроек: {e}")

def save_month_to_db():
    if not supabase or not st.session_state.current_id: return
    try:
        payload = {"shifts": st.session_state.shifts, "tonnage": st.session_state.tonnage, "subordinates": st.session_state.subordinates}
        json_str = json.dumps(payload, ensure_ascii=False)
        data = {"pf_base64": base64.b64encode(json_str.encode('utf-8')).decode("utf-8")}
        supabase.table("app_data").update(data).eq("id", st.session_state.current_id).execute()
    except Exception as e: st.error(f"Ошибка сохранения месяца: {e}")

def load_month_from_db(record_id):
    if not supabase: return False
    try:
        res = supabase.table("app_data").select("*").eq("id", record_id).execute()
        if res.data and res.data[0]["pf_base64"]:
            raw_json = base64.b64decode(res.data[0]["pf_base64"]).decode('utf-8')
            payload = json.loads(raw_json)
            raw_shifts = payload.get("shifts", {})
            st.session_state.shifts = {}
            for k, v in raw_shifts.items():
                if isinstance(v, list): st.session_state.shifts[k] = {"timesheet": v, "graph_only": {}}
                else:
                    go = v.get("graph_only", {})
                    if isinstance(go, list): go = {name: "" for name in go}
                    st.session_state.shifts[k] = {"timesheet": v.get("timesheet", []), "graph_only": go}
            st.session_state.tonnage = payload.get("tonnage", {})
            st.session_state.subordinates = payload.get("subordinates", [])
            st.session_state.current_id = record_id
            parts = record_id.split("_")
            st.session_state.active_month = parts[0]
            st.session_state.active_year = parts[1]
            return True
    except Exception as e: st.error(f"Ошибка загрузки месяца: {e}")
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

        tpl_pf_name = get_template_name("tpl_tabel", dept_id, "xlsm")
        tpl_gr_name = get_template_name("tpl_graph", dept_id, "xlsx")
        tpl_sub_name = get_template_name("tpl_sub", dept_id, "xlsm")
        tpl_sh_name = "tpl_sh.xlsx"

        if not all(os.path.exists(f) for f in [tpl_pf_name, tpl_gr_name, tpl_sub_name, tpl_sh_name]):
            st.error(f"Не все шаблоны загружены на GitHub для отдела '{dept_id}'.")
            return None

        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            # 5.1. ТАБЕЛЬ
            status_container.write("Читаем шаблон Табеля...")
            with open(tpl_pf_name, "rb") as f:
                wb_pf = load_workbook(io.BytesIO(f.read()), keep_vba=True)
            
            status_container.write("Обрабатываем Табель...")
            wb_pf.calculation.fullCalcOnLoad = True
            tpl_pf_sheet = next((s for s in wb_pf.sheetnames if s not in ("Итог", "рубкг")), None)

            for s_name in list(wb_pf.sheetnames):
                if s_name not in ("Итог", "рубкг", tpl_pf_sheet): del wb_pf[s_name]

            for emp in st.session_state.staff_roster:
                sheet_name = emp["full_name"][:31]
                rate_val = get_rate_for(emp["post"], emp.get("grade", ""))
                if sheet_name not in wb_pf.sheetnames and tpl_pf_sheet in wb_pf.sheetnames:
                    new_ws = wb_pf.copy_worksheet(wb_pf[tpl_pf_sheet])
                    new_ws.title = sheet_name
                    for r in range(2, 33):
                        if new_ws.cell(row=r, column=2).data_type != 'f': new_ws.cell(row=r, column=2).value = None 
                        if new_ws.cell(row=r, column=3).data_type != 'f': new_ws.cell(row=r, column=3).value = rate_val 

            if tpl_pf_sheet in wb_pf.sheetnames and not any(e["full_name"][:31] == tpl_pf_sheet for e in st.session_state.staff_roster):
                del wb_pf[tpl_pf_sheet]

            for day_str, data in st.session_state.shifts.items():
                day = int(day_str)
                timesheet_emps = data.get("timesheet", [])
                graph_only_emps = data.get("graph_only", {})
                
                for emp_full_name in timesheet_emps:
                    # Если человек на подработке, не ставим ему смену в табель
                    if emp_full_name in graph_only_emps:
                        continue 
                        
                    sheet_name = emp_full_name[:31]
                    if sheet_name in wb_pf.sheetnames: 
                        wb_pf[sheet_name].cell(row=day + 1, column=2, value=1)
                        
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

            status_container.write("Сохраняем Табель...")
            buf = io.BytesIO(); wb_pf.save(buf)
            zf.writestr(f"{prefix}_{m_name}_{y_val}.xlsm", buf.getvalue())
            del wb_pf; gc.collect()

            # 5.2. ГРАФИК
            status_container.write("Читаем шаблон Графика...")
            with open(tpl_gr_name, "rb") as f:
                wb_gr = load_workbook(io.BytesIO(f.read()))
            
            status_container.write("Обрабатываем График...")
            ws_gr = wb_gr.active
            ws_gr['A1'] = m_name
            try: ws_gr.title = f"{m_name} {y_val}"
            except: pass
            
            m_idx = MONTHS_RU.index(m_name) + 1
            num_days = calendar.monthrange(int(y_val), m_idx)[1]
            
            itogo_col = None
            for c_idx in range(4, 50):
                for r_idx in [1, 2, 3]:
                    val = str(ws_gr.cell(row=r_idx, column=c_idx).value or "").strip().upper()
                    if "ИТОГО" in val:
                        itogo_col = c_idx; break
                if itogo_col: break
                
            if not itogo_col:
                itogo_col = 35 
                ws_gr.cell(row=2, column=itogo_col, value="ИТОГО")

            block_r = 4
            for r in range(1, 20):
                for c in range(1, 4):
                    if "отделение" in str(ws_gr.cell(row=r, column=c).value or "").lower(): block_r = r; break
                if block_r > 4: break
                
            last_emp_row = block_r + len(st.session_state.staff_roster)
            for r_idx in range(1, last_emp_row + 1):
                ws_gr.cell(row=r_idx, column=itogo_col).fill = BLUE_FILL

            for d in range(1, 32):
                ws_gr.cell(row=2, column=3 + d, value=d if d <= num_days else None)
                for r_idx in [1, 2, 3]:
                    ws_gr.cell(row=r_idx, column=3 + d).fill = NO_FILL
            
            for r_idx in range(1, last_emp_row + 1):
                for c_idx in [1, 2, 3]: ws_gr.cell(row=r_idx, column=c_idx).fill = NO_FILL
            
            for i, emp in enumerate(st.session_state.staff_roster):
                r = block_r + 1 + i
                ws_gr.cell(row=r, column=1, value=i + 1)
                ws_gr.cell(row=r, column=2, value=emp["full_name"])
                ws_gr.cell(row=r, column=3, value=emp["post"])

                for c_idx in range(4, 3 + num_days + 1):
                    cell = ws_gr.cell(row=r, column=c_idx)
                    cell.value = None; cell.fill = NO_FILL; cell.comment = None

                shifts_count = 0
                for day_str, data in st.session_state.shifts.items():
                    day = int(day_str); col_idx = 3 + day
                    hours = 8 if emp["post"] in ("Весовщик", "Грузчик") else 11
                    
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
            buf = io.BytesIO(); wb_gr.save(buf)
            zf.writestr(f"ГРАФИК_{prefix}_{m_name}_{y_val}.xlsx", buf.getvalue())
            del wb_gr; gc.collect()

            # 5.3. ПОДРАБОТЧИКИ
            status_container.write("Читаем шаблон Подработчиков...")
            with open(tpl_sub_name, "rb") as f:
                wb_sub = load_workbook(io.BytesIO(f.read()), keep_vba=True)
            
            status_container.write("Обрабатываем Подработчиков...")
            wb_sub.calculation.fullCalcOnLoad = True
            tpl_sub_sheet = next((s for s in wb_sub.sheetnames if s != "Итог"), None)
                    
            for s in list(wb_sub.sheetnames):
                if s not in ("Итог", tpl_sub_sheet): del wb_sub[s]
                    
            subs_by_fio = {}
            for sub in st.session_state.subordinates:
                sheet_name = sub["short_name"][:31]
                if sheet_name not in wb_sub.sheetnames and tpl_sub_sheet:
                    new_ws = wb_sub.copy_worksheet(wb_sub[tpl_sub_sheet])
                    new_ws.title = sheet_name
                    for r in range(2, 33): 
                        if new_ws.cell(row=r, column=2).data_type != 'f': new_ws.cell(row=r, column=2).value = None
                        if new_ws.cell(row=r, column=3).data_type != 'f': new_ws.cell(row=r, column=3).value = None
                
                day = int(sub["date"].split(".")[0])
                wb_sub[sheet_name].cell(row=day + 1, column=2, value=1)
                wb_sub[sheet_name].cell(row=day + 1, column=3, value=sub["amount"])
                
                if sheet_name not in subs_by_fio: subs_by_fio[sheet_name] = {"amount": 0.0, "shifts": 0}
                subs_by_fio[sheet_name]["amount"] += float(sub["amount"])
                subs_by_fio[sheet_name]["shifts"] += 1
                
            if tpl_sub_sheet in wb_sub.sheetnames: del wb_sub[tpl_sub_sheet]
            
            if "Итог" in wb_sub.sheetnames:
                ws_it = wb_sub["Итог"]
                itogo_row = 50
                for r in range(2, 50):
                    if str(ws_it.cell(row=r, column=1).value or "").strip().upper() == "ИТОГО:": itogo_row = r; break
                    ws_it.cell(row=r, column=1).value = None
                    ws_it.cell(row=r, column=2).value = None
                    ws_it.cell(row=r, column=3).value = None
                
                for i, (fio, data) in enumerate(subs_by_fio.items()):
                    r = 2 + i
                    if r < itogo_row:
                        ws_it.cell(row=r, column=1, value=fio)
                        ws_it.cell(row=r, column=2, value=f"='{fio}'!J33")
                        ws_it.cell(row=r, column=3, value=f"='{fio}'!B33")

                if itogo_row < 50:
                    ws_it.cell(row=itogo_row, column=2, value=f"=SUM(B2:B{itogo_row-1})")
                    ws_it.cell(row=itogo_row, column=3, value=f"=SUM(C2:C{itogo_row-1})")
                    
            status_container.write("Сохраняем Подработчиков...")
            buf = io.BytesIO(); wb_sub.save(buf)
            zf.writestr(f"Подработчики_{prefix}_{m_name}_{y_val}.xlsm", buf.getvalue())
            del wb_sub; gc.collect()

            # 5.4. ВЕДОМОСТЬ ШАШИНОЙ
            status_container.write("Читаем шаблон Ведомости Шашиной...")
            with open(tpl_sh_name, "rb") as f:
                wb_sh = load_workbook(io.BytesIO(f.read()))
            
            status_container.write("Обрабатываем Ведомость Шашиной...")
            wb_sh.calculation.fullCalcOnLoad = True
            sh_sheet_map = {
                "pf": "ПФ", 
                "main_deboning": "Обвалка", 
                "chicken_deboning": "Куриная обвалка",
                "thermal": "Термичка", 
                "offal": "ЛО", 
                "delicatessen": "ДО"
            }
            sheet_name_sh = sh_sheet_map.get(dept_id, "ПФ")
            
            if sheet_name_sh in wb_sh.sheetnames:
                ws_sh = wb_sh[sheet_name_sh]
                start_row = 4
                for i, sub in enumerate(st.session_state.subordinates):
                    r = start_row + i
                    ws_sh.cell(row=r, column=2, value=i+1)
                    ws_sh.cell(row=r, column=3, value=sub["short_name"])
                    ws_sh.cell(row=r, column=4, value=sub["post"])
                    ws_sh.cell(row=r, column=5, value=sub.get("grade", ""))
                    ws_sh.cell(row=r, column=6, value=11)
                    ws_sh.cell(row=r, column=8, value=sub["date"])
                    ws_sh.cell(row=r, column=9, value=sub["amount"])
                    
            status_container.write("Сохраняем Ведомость Шашиной...")
            buf = io.BytesIO(); wb_sh.save(buf)
            zf.writestr(f"Шашина_{m_name}_{y_val}.xlsx", buf.getvalue())
            del wb_sh; gc.collect()

        return zip_buf.getvalue()
    except Exception as e:
        st.error(f"Критическая ошибка при создании файлов: {e}")
        return None

# ==========================================
# 6. ИНТЕРФЕЙС (SIDEBAR)
# ==========================================
st.sidebar.title("🏭 Выбор отделения")
dept_options = [""] + list(DEPARTMENTS.keys())
dept_labels = ["--- Выберите отделение ---"] + list(DEPARTMENTS.values())

try: current_idx = dept_options.index(st.session_state.current_dept)
except ValueError: current_idx = 0

selected_idx = st.sidebar.selectbox(
    "В каком отделе вы работаете?", 
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
    if selected_dept: load_global_settings()
    st.rerun()

if not st.session_state.current_dept:
    st.title("🏭 Учёт выработки")
    st.info("👈 Пожалуйста, выберите ваше отделение в левой панели, чтобы начать работу.")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.header("🗓️ Текущий период")
if st.session_state.active_month:
    st.sidebar.success(f"### {st.session_state.active_month} {st.session_state.active_year}")
    
    if st.sidebar.button("⬇️ Сгенерировать и скачать файлы (ZIP)", type="primary", use_container_width=True):
        with st.status("Формируем Excel-файлы...", expanded=True) as status:
            zip_data = generate_and_download(status)
            if zip_data:
                status.update(label="✅ Файлы успешно сформированы!", state="complete", expanded=False)
                
                zip_prefix = FILE_PREFIXES.get(st.session_state.current_dept, "Отчет")
                
                st.sidebar.download_button(
                    label="💾 Сохранить на компьютер", 
                    data=zip_data, 
                    file_name=f"Итоги_{zip_prefix}_{st.session_state.active_month}_{st.session_state.active_year}.zip", 
                    mime="application/zip", use_container_width=True
                )
            else:
                status.update(label="❌ Ошибка при создании файлов", state="error", expanded=True)
else:
    st.sidebar.warning("### Месяц не выбран\nСоздайте новый или выберите старый.")

st.sidebar.markdown("---")
if supabase:
    st.sidebar.subheader("📂 Открыть месяц")
    try:
        dept_id = st.session_state.current_dept
        res = supabase.table("app_data").select("id").neq("id", f"GLOBAL_SETTINGS_{dept_id}").execute()
        saved = [r["id"] for r in res.data if r["id"].endswith(f"_{dept_id}")]
        if saved:
            saved.sort(reverse=True)
            sel = st.sidebar.selectbox("Выберите месяц", saved, format_func=lambda x: x.replace(f"_{dept_id}", "").replace("_", " "))
            if st.sidebar.button("Загрузить", use_container_width=True):
                if load_month_from_db(sel):
                    st.sidebar.success("Загружено!")
                    st.rerun()
        else:
            st.sidebar.caption("Нет сохраненных месяцев для этого отдела.")
    except: pass

st.title(f"📊 {DEPARTMENTS[st.session_state.current_dept]} — {st.session_state.active_month} {st.session_state.active_year}" if st.session_state.active_month else f"📊 {DEPARTMENTS[st.session_state.current_dept]} — Учёт выработки")

tab0, tab1, tab2, tab3, tab4 = st.tabs(["✨ Создать месяц", "📅 Табель смен", "👷 Подработчики", "💰 Ставки", "👥 Кадры"])

# ================= TAB 0: СОЗДАНИЕ =================
with tab0:
    st.subheader("Новый месяц")
    c1, c2 = st.columns(2)
    with c1: new_month = st.selectbox("Месяц", MONTHS_RU, index=datetime.date.today().month - 1)
    with c2: 
        years_list = [str(y) for y in range(2024, 2036)]
        new_year = st.selectbox("Год", years_list, index=years_list.index(str(datetime.date.today().year)))

    if st.button("✨ Создать месяц", type="primary"):
        dept_id = st.session_state.current_dept
        record_id = f"{new_month}_{new_year}_{dept_id}"
        if not supabase:
            st.error("Нет подключения к БД.")
        else:
            try:
                existing = supabase.table("app_data").select("id").eq("id", record_id).execute()
                if existing.data:
                    st.error(f"Месяц {new_month} {new_year} для этого отдела уже существует!")
                else:
                    st.session_state.shifts = {}
                    st.session_state.tonnage = {}
                    st.session_state.subordinates = []
                    st.session_state.active_month = new_month
                    st.session_state.active_year = str(new_year)
                    st.session_state.current_id = record_id
                    
                    empty_payload = json.dumps({"shifts": {}, "tonnage": {}, "subordinates": []})
                    data = {"id": record_id, "pf_base64": base64.b64encode(empty_payload.encode('utf-8')).decode("utf-8")}
                    supabase.table("app_data").insert(data).execute()
                    
                    st.success(f"✅ Месяц {new_month} {new_year} для отдела {DEPARTMENTS[dept_id]} создан!")
                    st.rerun()
            except Exception as e:
                st.error(f"Ошибка: {e}")

# ================= TAB 1: ТАБЕЛЬ =================
with tab1:
    if not st.session_state.active_month:
        st.info("👈 Сначала создайте месяц.")
    else:
        st.subheader("Отметка смен")
        st.info("Если сотрудник работает в своем отделе — ставьте «Смена». Если он ушел на подработку в другой цех — ставьте обе галочки («Смена» и «Подработка»), чтобы он попал в график, но не попал в табель.")
        
        c1, c2 = st.columns(2)
        with c1: shift_date = st.date_input("Дата смены", datetime.date.today())
        with c2: tonnage = st.number_input("Тоннаж цеха (кг)", min_value=0.0, value=0.0, step=1.0)

        st.markdown("**Кого отметить?**")
        selected = {}
        for emp in st.session_state.staff_roster:
            col1, col2, col3, col4 = st.columns([4, 2, 2, 3])
            with col1:
                label = f"{emp['short_name']} — {emp['post']}"
                if emp.get('grade') and emp['grade'] != "—": label += f" ({emp['grade']})"
                st.write(label)
            with col2:
                is_main = st.checkbox("Смена", key=f"main_{emp['full_name']}", label_visibility="collapsed")
            with col3:
                is_other = st.checkbox("Подработка", key=f"other_{emp['full_name']}")
            with col4:
                dept_name = st.text_input("Цех", key=f"dept_{emp['full_name']}", label_visibility="collapsed", placeholder="Введите цех...")
            
            selected[emp["full_name"]] = {"timesheet": is_main, "graph_only": is_other, "dept": dept_name}

        if st.button("💾 Провести смену", type="primary"):
            day_str = str(shift_date.day)
            if day_str not in st.session_state.shifts: st.session_state.shifts[day_str] = {"timesheet": [], "graph_only": {}}
            if "graph_only" not in st.session_state.shifts[day_str]: st.session_state.shifts[day_str]["graph_only"] = {}
            if isinstance(st.session_state.shifts[day_str]["graph_only"], list): st.session_state.shifts[day_str]["graph_only"] = {n: "" for n in st.session_state.shifts[day_str]["graph_only"]}
            
            for emp_full_name, checks in selected.items():
                # Независимая запись в Табель
                if checks["timesheet"]:
                    if emp_full_name not in st.session_state.shifts[day_str]["timesheet"]:
                        st.session_state.shifts[day_str]["timesheet"].append(emp_full_name)
                else:
                    if emp_full_name in st.session_state.shifts[day_str]["timesheet"]:
                        st.session_state.shifts[day_str]["timesheet"].remove(emp_full_name)
                
                # Независимая запись в График (Подработка)
                if checks["graph_only"]:
                    st.session_state.shifts[day_str]["graph_only"][emp_full_name] = checks["dept"]
                else:
                    if emp_full_name in st.session_state.shifts[day_str]["graph_only"]:
                        del st.session_state.shifts[day_str]["graph_only"][emp_full_name]
            
            if tonnage > 0: st.session_state.tonnage[day_str] = float(tonnage)
            save_month_to_db()
            st.success(f"Смена на {shift_date.strftime('%d.%m.%Y')} сохранена!")
            st.rerun()

        st.markdown("---")
        st.markdown("**Кто уже отмечен в этом месяце:**")
        for day, data in sorted(st.session_state.shifts.items(), key=lambda x: int(x[0])):
            timesheet_names = [next((s["short_name"] for s in st.session_state.staff_roster if s["full_name"] == e), e) for e in data.get("timesheet", [])]
            go_dict = data.get("graph_only", {})
            if isinstance(go_dict, list): go_dict = {name: "" for name in go_dict}
            graph_names = []
            for name, dept in go_dict.items():
                short_n = next((s["short_name"] for s in st.session_state.staff_roster if s["full_name"] == name), name)
                graph_names.append(f"{short_n}{f' (в {dept})' if dept else ''}")
            
            msg = f"**{day} число:** "
            if timesheet_names: 
                msg += f"Основная: {', '.join(timesheet_names)}. "
            if graph_names: 
                msg += f"Подработка: {', '.join(graph_names)}."
            st.write(msg)

# ================= TAB 2: ПОДРАБОТЧИКИ =================
with tab2:
    if not st.session_state.active_month:
        st.info("👈 Сначала создайте месяц.")
    else:
        st.subheader("Внесение подработки (Пофамильно)")
        c1, c2 = st.columns(2)
        with c1:
            fio = st.text_input("ФИО подработчика")
            post = st.selectbox("Должность", get_current_positions())
            available_grades = ["—"] + [r["grade"] for r in st.session_state.rates if r["post"] == post and r["grade"]]
            grade = st.selectbox("Разряд", available_grades) if len(available_grades) > 1 else "—"
        with c2:
            sub_date = st.date_input("Дата")
            amount = st.number_input("Сумма (руб)", min_value=0.0, value=4500.0, step=100.0)

        if st.button("💾 Добавить подработку", type="primary"):
            if not fio.strip(): st.error("Укажите ФИО")
            else:
                st.session_state.subordinates.append({"short_name": fio.strip(), "post": post, "grade": grade, "date": sub_date.strftime("%d.%m.%Y"), "amount": float(amount)})
                save_month_to_db()
                st.success("Подработка добавлена!")
                st.rerun()

        st.markdown("---")
        if st.session_state.subordinates:
            df_sub = pd.DataFrame(st.session_state.subordinates)
            if not df_sub.empty: st.table(df_sub[["short_name", "post", "grade", "date", "amount"]])
        else: st.write("Пока нет записей.")

# ================= TAB 3: СТАВКИ =================
with tab3:
    st.subheader(f"Справочник ставок ({DEPARTMENTS[st.session_state.current_dept]})")
    
    if not st.session_state.rates:
        st.session_state.rates = [{"post": "", "grade": "", "rate": 0}]
        
    df_rates = pd.DataFrame(st.session_state.rates)
    edited_rates = st.data_editor(
        df_rates, num_rows="dynamic", use_container_width=True,
        column_config={"post": st.column_config.TextColumn("Должность", required=True), "grade": st.column_config.TextColumn("Разряд"), "rate": st.column_config.NumberColumn("Ставка (руб)", min_value=0, step=100, format="%d ₽")},
        key="rates_editor"
    )
    if st.button("💾 Сохранить ставки", type="primary"):
        new_rates = []
        for _, row in edited_rates.iterrows():
            if pd.notna(row["post"]) and str(row["post"]).strip() != "":
                new_rates.append({"post": str(row["post"]).strip(), "grade": str(row["grade"]).strip() if pd.notna(row["grade"]) else "", "rate": int(row["rate"]) if pd.notna(row["rate"]) else 0})
        st.session_state.rates = new_rates
        save_global_settings()
        st.success("✅ Ставки сохранены!")
        st.rerun()

# ================= TAB 4: КАДРЫ =================
with tab4:
    st.subheader(f"Реестр сотрудников ({DEPARTMENTS[st.session_state.current_dept]})")
    if st.session_state.staff_roster:
        all_grades = set()
        for r in st.session_state.rates:
            if r.get("grade") and str(r["grade"]).strip(): all_grades.add(str(r["grade"]).strip())
        grade_options = ["—"] + sorted(list(all_grades))
        
        df_staff = pd.DataFrame(st.session_state.staff_roster).rename(columns={"full_name": "Полное ФИО", "short_name": "Короткое ФИО", "post": "Должность", "grade": "Разряд"})
        edited_df = st.data_editor(
            df_staff, num_rows="dynamic", use_container_width=True,
            column_config={"Полное ФИО": st.column_config.TextColumn(required=True), "Короткое ФИО": st.column_config.TextColumn(required=True), "Должность": st.column_config.SelectboxColumn(options=get_current_positions(), required=True), "Разряд": st.column_config.SelectboxColumn(options=grade_options)},
            key="staff_editor"
        )
        if st.button("💾 Сохранить штат", type="primary"):
            new_roster = []
            for _, row in edited_df.iterrows():
                if pd.notna(row["Полное ФИО"]) and str(row["Полное ФИО"]).strip() != "":
                    new_roster.append({"full_name": str(row["Полное ФИО"]).strip(), "short_name": str(row["Короткое ФИО"]).strip() if pd.notna(row["Короткое ФИО"]) else make_short_name(str(row["Полное ФИО"])), "post": str(row["Должность"]).strip() if pd.notna(row["Должность"]) else get_current_positions()[0], "grade": str(row["Разряд"]).strip() if pd.notna(row["Разряд"]) else "—"})
            st.session_state.staff_roster = new_roster
            save_global_settings()
            st.success("✅ Штат сохранен!")
            st.rerun()
    else:
        st.write("Штат пуст. Нажмите кнопку ниже, чтобы загрузить из шаблона.")
        if st.button("Загрузить сотрудников из шаблона"):
            st.session_state.staff_roster = load_initial_staff(st.session_state.current_dept)
            save_global_settings()
            st.rerun()
