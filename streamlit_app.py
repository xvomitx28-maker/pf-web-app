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
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from supabase import create_client, Client

# ==========================================
# 🎨 БЕЗОПАСНЫЙ ДИЗАЙН (CSS)
# ==========================================
st.set_page_config(page_title="Учёт ПФ", page_icon="📊", layout="wide")

custom_css = """
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 30px;
        background-color: transparent;
        border-bottom: 2px solid #e5e7eb;
        padding-bottom: 0px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        background-color: transparent;
        border: none;
        font-weight: 500;
        font-size: 15px;
    }
    .stTabs [aria-selected="true"] {
        color: #0055ff !important;
        border-bottom: 3px solid #0055ff !important;
        font-weight: 600;
    }
    
    .stButton>button[kind="primary"] {
        background-color: #0055ff;
        color: white;
        border: none;
        border-radius: 6px;
        box-shadow: 0 2px 4px rgba(0, 85, 255, 0.2);
    }
    .stButton>button[kind="primary"]:hover {
        background-color: #0044cc;
        color: white;
    }
    
    .stButton>button {
        border-radius: 6px;
        border: 1px solid #d1d5db;
    }
    
    [data-testid="stDataFrame"] {
        border-radius: 8px;
        border: 1px solid #e5e7eb;
        background-color: #ffffff;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==========================================
# 1. ПОДКЛЮЧЕНИЕ К БАЗЕ ДАННЫХ
# ==========================================
@st.cache_resource
def init_connection():
    try:
        return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    except Exception:
        return None

supabase = init_connection()

# ==========================================
# 2. НАЗВАНИЯ ФАЙЛОВ-ШАБЛОНОВ
# ==========================================
TPL_PF = "tpl_pf.xlsm"
TPL_SUB = "tpl_sub.xlsm"
TPL_GRAPH = "tpl_graph.xlsx"
TPL_SH = "tpl_sh.xlsx"

# ==========================================
# 3. БАЗОВЫЕ НАСТРОЙКИ И КОНСТАНТЫ
# ==========================================
POSITIONS_LIST = ["Изготовитель мясных п/ф", "Весовщик", "Грузчик", "Формовщик", "Жиловщик", "Оператор линии", "Бригадир"]

DEFAULT_RATES = [
    {"post": "Весовщик", "grade": "", "rate": 5000},
    {"post": "Грузчик", "grade": "", "rate": 4300},
    {"post": "Изготовитель мясных п/ф", "grade": "1 разряд", "rate": 3700},
    {"post": "Изготовитель мясных п/ф", "grade": "2 разряд", "rate": 4500},
    {"post": "Формовщик", "grade": "", "rate": 4500},
    {"post": "Жиловщик", "grade": "", "rate": 4500},
    {"post": "Оператор линии", "grade": "", "rate": 4500},
    {"post": "Бригадир", "grade": "", "rate": 4500},
]

MONTHS_RU = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

ORANGE_FILL = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
NO_FILL = PatternFill(fill_type=None) 

# ==========================================
# 4. ИНИЦИАЛИЗАЦИЯ СЕССИИ
# ==========================================
_defaults = {
    "active_month": "", "active_year": "", "current_id": None,
    "shifts": {}, 
    "tonnage": {}, "subordinates": [], 
    "staff_roster": [], "rates": DEFAULT_RATES
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ==========================================
# 5. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ==========================================
def clean_name(name):
    return re.sub(r'[^а-яё]', '', str(name).lower())

def make_short_name(full_name):
    parts = str(full_name).strip().split()
    if len(parts) > 1:
        return f"{parts[0]} {parts[1][0]}."
    return parts[0] if parts else ""

def get_rate_for(post, grade):
    for r in st.session_state.rates:
        if r["post"] == post and r["grade"] == grade:
            return r["rate"]
    for r in st.session_state.rates:
        if r["post"] == post:
            return r["rate"]
    return 4500

def load_initial_staff():
    if not os.path.exists(TPL_PF): return []
    wb = load_workbook(io.BytesIO(open(TPL_PF, "rb").read()), keep_vba=True)
    res = []
    for sname in wb.sheetnames:
        if sname in ("Итог", "рубкг") or sname.startswith("[Уволен]"):
            continue
        low = sname.lower()
        if "рузиева" in low: post = "Весовщик"
        elif "насруддинов" in low or "рузиев и" in low: post = "Грузчик"
        else: post = "Изготовитель мясных п/ф"
        
        res.append({
            "full_name": sname, 
            "short_name": make_short_name(sname), 
            "post": post,
            "grade": "—"
        })
    return res

# ==========================================
# 6. РАБОТА С БАЗОЙ ДАННЫХ (SUPABASE)
# ==========================================
def save_global_settings():
    if not supabase: return
    try:
        payload = {
            "staff_roster": st.session_state.staff_roster,
            "rates": st.session_state.rates
        }
        json_str = json.dumps(payload, ensure_ascii=False)
        data = {"pf_base64": base64.b64encode(json_str.encode('utf-8')).decode("utf-8")}
        
        res = supabase.table("app_data").select("id").eq("id", "GLOBAL_SETTINGS").execute()
        if res.data:
            supabase.table("app_data").update(data).eq("id", "GLOBAL_SETTINGS").execute()
        else:
            data["id"] = "GLOBAL_SETTINGS"
            supabase.table("app_data").insert(data).execute()
    except Exception as e:
        st.error(f"Ошибка сохранения глобальных настроек: {e}")

def load_global_settings():
    if not supabase: return
    try:
        res = supabase.table("app_data").select("*").eq("id", "GLOBAL_SETTINGS").execute()
        if res.data and res.data[0]["pf_base64"]:
            raw_json = base64.b64decode(res.data[0]["pf_base64"]).decode('utf-8')
            payload = json.loads(raw_json)
            st.session_state.staff_roster = payload.get("staff_roster", [])
            st.session_state.rates = payload.get("rates", DEFAULT_RATES)
        else:
            st.session_state.staff_roster = load_initial_staff()
            st.session_state.rates = DEFAULT_RATES
            save_global_settings()
    except Exception as e:
        st.error(f"Ошибка загрузки глобальных настроек: {e}")

def save_month_to_db():
    if not supabase or not st.session_state.current_id: return
    try:
        payload = {
            "shifts": st.session_state.shifts, 
            "tonnage": st.session_state.tonnage, 
            "subordinates": st.session_state.subordinates
        }
        json_str = json.dumps(payload, ensure_ascii=False)
        data = {"pf_base64": base64.b64encode(json_str.encode('utf-8')).decode("utf-8")}
        supabase.table("app_data").update(data).eq("id", st.session_state.current_id).execute()
    except Exception as e:
        st.error(f"Ошибка сохранения месяца: {e}")

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
                if isinstance(v, list): 
                    st.session_state.shifts[k] = {"timesheet": v, "graph_only": {}}
                else:
                    go = v.get("graph_only", {})
                    if isinstance(go, list):
                        go = {name: "" for name in go}
                    st.session_state.shifts[k] = {"timesheet": v.get("timesheet", []), "graph_only": go}

            st.session_state.tonnage = payload.get("tonnage", {})
            st.session_state.subordinates = payload.get("subordinates", [])
            
            st.session_state.current_id = record_id
            st.session_state.active_month, st.session_state.active_year = record_id.split("_")
            return True
    except Exception as e:
        st.error(f"Ошибка загрузки месяца: {e}")
    return False

# ==========================================
# 7. ФОРМИРОВАНИЕ EXCEL-ФАЙЛОВ (ГЕНЕРАЦИЯ)
# ==========================================
def generate_and_download():
    zip_buf = io.BytesIO()
    m_name = st.session_state.active_month
    y_val = st.session_state.active_year

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # ---------------------------------------------------------
        # 7.1. ТАБЕЛЬ (xlsm)
        # ---------------------------------------------------------
        wb_pf = load_workbook(io.BytesIO(open(TPL_PF, "rb").read()), keep_vba=True)
        wb_pf.calculation.fullCalcOnLoad = True
        wb_pf.calculation.calcMode = 'auto'

        tpl_pf_name = None
        for s_name in wb_pf.sheetnames:
            if s_name not in ("Итог", "рубкг"):
                tpl_pf_name = s_name
                break
        
        if not tpl_pf_name: return None

        for s_name in list(wb_pf.sheetnames):
            if s_name not in ("Итог", "рубкг", tpl_pf_name):
                del wb_pf[s_name]

        for emp in st.session_state.staff_roster:
            name = emp["full_name"]
            rate_val = get_rate_for(emp["post"], emp.get("grade", ""))
            sheet_name = name[:31]
            
            if sheet_name not in wb_pf.sheetnames:
                if tpl_pf_name in wb_pf.sheetnames:
                    new_ws = wb_pf.copy_worksheet(wb_pf[tpl_pf_name])
                    new_ws.title = sheet_name
                    for r in range(2, 33):
                        cell_b = new_ws.cell(row=r, column=2)
                        if cell_b.data_type != 'f': cell_b.value = None 
                        cell_c = new_ws.cell(row=r, column=3)
                        if cell_c.data_type != 'f': cell_c.value = rate_val 

        if tpl_pf_name in wb_pf.sheetnames and not any(e["full_name"][:31] == tpl_pf_name for e in st.session_state.staff_roster):
            del wb_pf[tpl_pf_name]

        for day_str, data in st.session_state.shifts.items():
            day = int(day_str)
            timesheet_emps = data.get("timesheet", [])
            for emp_full_name in timesheet_emps:
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
                name = emp["full_name"]
                sheet_name = name[:31]
                ws_it.cell(row=r, column=1, value=name)
                ws_it.cell(row=r, column=2, value=f"='{sheet_name}'!J33")
                ws_it.cell(row=r, column=3, value=f"='{sheet_name}'!B33")
                
            itogo_r = 2 + len(st.session_state.staff_roster)
            ws_it.cell(row=itogo_r, column=1, value="ИТОГО:")
            ws_it.cell(row=itogo_r, column=2, value=f"=SUM(B2:B{itogo_r-1})")
            ws_it.cell(row=itogo_r, column=3, value=f"=SUM(C2:C{itogo_r-1})")
            
            ws_it.delete_rows(itogo_r + 1, 50)

        buf_pf = io.BytesIO(); wb_pf.save(buf_pf)
        zf.writestr(f"Полуфабрикаты_{m_name}_{y_val}.xlsm", buf_pf.getvalue())

        # ---------------------------------------------------------
        # 7.2. ГРАФИК (xlsx) 
        # ---------------------------------------------------------
        wb_gr = load_workbook(io.BytesIO(open(TPL_GRAPH, "rb").read()))
        wb_gr.calculation.fullCalcOnLoad = True
        wb_gr.calculation.calcMode = 'auto'
        ws_gr = wb_gr.active
        ws_gr['A1'] = m_name
        try: ws_gr.title = f"{m_name} {y_val}"
        except: pass
        
        m_idx = MONTHS_RU.index(m_name) + 1
        num_days = calendar.monthrange(int(y_val), m_idx)[1]
        
        # 1. Ищем колонку "ИТОГО" в шаблоне ДО того, как ставить дни
        current_itogo_col = None
        for c_idx in range(4, 50):
            for r_idx in [1, 2, 3]:
                val = ws_gr.cell(row=r_idx, column=c_idx).value
                if val and str(val).strip().upper() == "ИТОГО":
                    current_itogo_col = c_idx
                    break
            if current_itogo_col: break
        
        # 2. Корректируем количество колонок, чтобы ИТОГО стояло ровно после последнего дня
        target_itogo_col = 4 + num_days
        
        if current_itogo_col:
            if current_itogo_col < target_itogo_col:
                ws_gr.insert_cols(current_itogo_col, target_itogo_col - current_itogo_col)
            elif current_itogo_col > target_itogo_col:
                ws_gr.delete_cols(target_itogo_col, current_itogo_col - target_itogo_col)
        
        # Теперь колонка ИТОГО гарантированно находится на позиции target_itogo_col
        itogo_col = target_itogo_col
        
        # 3. Ставим дни в шапку и очищаем лишние цвета
        for d in range(1, num_days + 1):
            ws_gr.cell(row=2, column=3 + d, value=d)
            for r_idx in [1, 2, 3]:
                cell_header = ws_gr.cell(row=r_idx, column=3 + d)
                cell_header.fill = NO_FILL
        
        # Очищаем заливку в A, B, C
        block_r = 4
        for r in range(1, 50):
            for c in range(1, 4):
                if "полуфабрикат" in str(ws_gr.cell(row=r, column=c).value or "").lower():
                    block_r = r
                    break
            if block_r > 4: break
            
        last_day_col_letter = get_column_letter(3 + num_days)
        last_emp_row = block_r + len(st.session_state.staff_roster)
        
        for r_idx in range(1, last_emp_row + 1):
            for c_idx in [1, 2, 3]:
                ws_gr.cell(row=r_idx, column=c_idx).fill = NO_FILL
        
        # 4. Заполняем сотрудников
        for i, emp in enumerate(st.session_state.staff_roster):
            r = block_r + 1 + i
            ws_gr.cell(row=r, column=1, value=i + 1)
            ws_gr.cell(row=r, column=2, value=emp["full_name"])
            ws_gr.cell(row=r, column=3, value=emp["post"])

            for c_idx in range(4, 3 + num_days + 1):
                cell_to_clear = ws_gr.cell(row=r, column=c_idx)
                cell_to_clear.value = None
                cell_to_clear.fill = NO_FILL 
                cell_to_clear.comment = None

            for day_str, data in st.session_state.shifts.items():
                timesheet_emps = data.get("timesheet", [])
                graph_only_dict = data.get("graph_only", {})
                
                day = int(day_str)
                col_idx = 3 + day
                hours = 8 if emp["post"] in ("Весовщик", "Грузчик") else 11
                
                if emp["full_name"] in graph_only_dict:
                    dept_name = graph_only_dict[emp["full_name"]]
                    cell = ws_gr.cell(row=r, column=col_idx, value=hours)
                    cell.fill = ORANGE_FILL 
                    
                    if dept_name:
                        comment_text = f"Подработка: {dept_name}"
                    else:
                        comment_text = "Подработка в другом цехе"
                    cell.comment = Comment(text=comment_text, author="Система учёта")
                    
                elif emp["full_name"] in timesheet_emps:
                    ws_gr.cell(row=r, column=col_idx, value=hours)
            
            ws_gr.cell(row=r, column=itogo_col, value=f"=COUNT(D{r}:{last_day_col_letter}{r})")
        
        start_delete_row = block_r + 1 + len(st.session_state.staff_roster)
        ws_gr.delete_rows(start_delete_row, 100)
        
        buf_gr = io.BytesIO(); wb_gr.save(buf_gr)
        zf.writestr(f"ГРАФИК_{m_name}_{y_val}.xlsx", buf_gr.getvalue())

        # ---------------------------------------------------------
        # 7.3. ПОДРАБОТЧИКИ (xlsm)
        # ---------------------------------------------------------
        wb_sub = load_workbook(io.BytesIO(open(TPL_SUB, "rb").read()), keep_vba=True)
        wb_sub.calculation.fullCalcOnLoad = True
        wb_sub.calculation.calcMode = 'auto'
        
        tpl_sub_name = None
        for s in wb_sub.sheetnames:
            if s != "Итог":
                tpl_sub_name = s
                break
                
        if not tpl_sub_name: return None
            
        for s in list(wb_sub.sheetnames):
            if s not in ("Итог", tpl_sub_name):
                del wb_sub[s]
                
        subs_by_fio = {}
        for sub in st.session_state.subordinates:
            sheet_name = sub["short_name"][:31]
            if sheet_name not in wb_sub.sheetnames:
                new_ws = wb_sub.copy_worksheet(wb_sub[tpl_sub_name])
                new_ws.title = sheet_name
                for r in range(2, 33): 
                    cell_b = new_ws.cell(row=r, column=2)
                    if cell_b.data_type != 'f': cell_b.value = None
                    cell_c = new_ws.cell(row=r, column=3)
                    if cell_c.data_type != 'f': cell_c.value = None
            
            day = int(sub["date"].split(".")[0])
            
            wb_sub[sheet_name].cell(row=day + 1, column=2, value=1)
            wb_sub[sheet_name].cell(row=day + 1, column=3, value=sub["amount"])
            
            if sheet_name not in subs_by_fio:
                subs_by_fio[sheet_name] = {"amount": 0.0, "shifts": 0}
            subs_by_fio[sheet_name]["amount"] += float(sub["amount"])
            subs_by_fio[sheet_name]["shifts"] += 1
            
        if tpl_sub_name in wb_sub.sheetnames:
            del wb_sub[tpl_sub_name]
        
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
                    ws_it.cell(row=r, column=2, value=f"='{fio}'!J33")
                    ws_it.cell(row=r, column=3, value=f"='{fio}'!B33")

            if itogo_row < 50:
                ws_it.cell(row=itogo_row, column=2, value=f"=SUM(B2:B{itogo_row-1})")
                ws_it.cell(row=itogo_row, column=3, value=f"=SUM(C2:C{itogo_row-1})")
                
        buf_sub = io.BytesIO(); wb_sub.save(buf_sub)
        zf.writestr(f"Подработчики_{m_name}_{y_val}.xlsm", buf_sub.getvalue())

        # ---------------------------------------------------------
        # 7.4. ВЕДОМОСТЬ ШАШИНОЙ (xlsx)
        # ---------------------------------------------------------
        wb_sh = load_workbook(io.BytesIO(open(TPL_SH, "rb").read()))
        wb_sh.calculation.fullCalcOnLoad = True
        if "ПФ" in wb_sh.sheetnames:
            ws_sh = wb_sh["ПФ"]
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
        buf_sh = io.BytesIO(); wb_sh.save(buf_sh)
        zf.writestr(f"Шашина_{m_name}_{y_val}.xlsx", buf_sh.getvalue())

    return zip_buf.getvalue()

# ==========================================
# 8. ЗАГРУЗКА ПРИ СТАРТЕ
# ==========================================
if "initialized" not in st.session_state:
    load_global_settings()
    st.session_state["initialized"] = True

# ==========================================
# 9. ИНТЕРФЕЙС (SIDEBAR)
# ==========================================
st.sidebar.header("🗓️ Текущий период")
if st.session_state.active_month:
    st.sidebar.success(f"### {st.session_state.active_month} {st.session_state.active_year}")
    if st.sidebar.button("⬇️ Сгенерировать и скачать файлы (ZIP)", type="primary", use_container_width=True):
        zip_data = generate_and_download()
        if zip_data:
            st.sidebar.download_button(
                label="💾 Сохранить на компьютер", 
                data=zip_data, 
                file_name=f"Итоги_{st.session_state.active_month}_{st.session_state.active_year}.zip", 
                mime="application/zip",
                use_container_width=True
            )
else:
    st.sidebar.warning("### Месяц не выбран\nСоздайте новый или выберите старый.")

st.sidebar.markdown("---")
if supabase:
    st.sidebar.subheader("📂 Открыть месяц")
    try:
        res = supabase.table("app_data").select("id").neq("id", "GLOBAL_SETTINGS").execute()
        saved = [r["id"] for r in res.data]
        if saved:
            saved.sort(reverse=True)
            sel = st.sidebar.selectbox("Выберите месяц", saved)
            if st.sidebar.button("Загрузить", use_container_width=True):
                if load_month_from_db(sel):
                    st.sidebar.success("Загружено!")
                    st.rerun()
    except: pass

# ==========================================
# 10. ОСНОВНОЙ ЭКРАН (ВКЛАДКИ)
# ==========================================
st.title(f"📊 Отделение Полуфабрикатов — {st.session_state.active_month} {st.session_state.active_year}" if st.session_state.active_month else "📊 Учёт выработки")

tab0, tab1, tab2, tab3, tab4 = st.tabs(["✨ Создать месяц", "📅 Табель смен", "👷 Подработчики", "💰 Ставки", "👥 Кадры"])

# ================= TAB 0: СОЗДАНИЕ =================
with tab0:
    st.subheader("Новый месяц")
    c1, c2 = st.columns(2)
    with c1: 
        new_month = st.selectbox("Месяц", MONTHS_RU, index=datetime.date.today().month - 1)
    with c2: 
        years_list = [str(y) for y in range(2024, 2036)]
        current_year_str = str(datetime.date.today().year)
        default_index = years_list.index(current_year_str) if current_year_str in years_list else 0
        new_year = st.selectbox("Год", years_list, index=default_index)

    if st.button("✨ Создать месяц", type="primary"):
        record_id = f"{new_month}_{new_year}"
        if not supabase:
            st.error("Нет подключения к БД.")
        else:
            try:
                existing = supabase.table("app_data").select("id").eq("id", record_id).execute()
                if existing.data:
                    st.error(f"Месяц {record_id} уже существует!")
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
                    
                    st.success("✅ Месяц создан!")
                    st.rerun()
            except Exception as e:
                st.error(f"Ошибка: {e}")

# ================= TAB 1: ТАБЕЛЬ =================
with tab1:
    if not st.session_state.active_month:
        st.info("👈 Сначала создайте месяц.")
    else:
        st.subheader("Отметка смен")
        c1, c2 = st.columns(2)
        with c1: shift_date = st.date_input("Дата смены", datetime.date.today())
        with c2: tonnage = st.number_input("Тоннаж цеха (кг)", min_value=0.0, value=0.0, step=1.0)

        st.markdown("**Кого отметить?**")
        selected = {}
        for emp in st.session_state.staff_roster:
            col_main, col_other, col_dept = st.columns([4, 3, 3])
            
            with col_main:
                label = f"{emp['short_name']} — {emp['post']}"
                if emp.get('grade') and emp['grade'] != "—":
                    label += f" ({emp['grade']})"
                is_main = st.checkbox(label, key=f"main_{emp['full_name']}")
            
            with col_other:
                is_other = st.checkbox("Подработка (другой цех)", key=f"other_{emp['full_name']}")
                
            with col_dept:
                dept_name = ""
                if is_other:
                    dept_name = st.text_input("Подразделение", key=f"dept_{emp['full_name']}", label_visibility="collapsed", placeholder="Введите цех...")
            
            selected[emp["full_name"]] = {"timesheet": is_main, "graph_only": is_other, "dept": dept_name}

        if st.button("💾 Провести смену", type="primary"):
            day_str = str(shift_date.day)
            if day_str not in st.session_state.shifts:
                st.session_state.shifts[day_str] = {"timesheet": [], "graph_only": {}}
            
            if "graph_only" not in st.session_state.shifts[day_str]:
                st.session_state.shifts[day_str]["graph_only"] = {}
            if isinstance(st.session_state.shifts[day_str]["graph_only"], list):
                st.session_state.shifts[day_str]["graph_only"] = {n: "" for n in st.session_state.shifts[day_str]["graph_only"]}
            
            for emp_full_name, checks in selected.items():
                if checks["timesheet"]:
                    if emp_full_name not in st.session_state.shifts[day_str]["timesheet"]:
                        st.session_state.shifts[day_str]["timesheet"].append(emp_full_name)
                
                if checks["graph_only"]:
                    st.session_state.shifts[day_str]["graph_only"][emp_full_name] = checks["dept"]
            
            if tonnage > 0:
                st.session_state.tonnage[day_str] = float(tonnage)
            
            save_month_to_db()
            st.success(f"Смена на {shift_date.strftime('%d.%m.%Y')} сохранена!")
            st.rerun()

        st.markdown("---")
        st.markdown("**Кто уже отмечен в этом месяце:**")
        for day, data in sorted(st.session_state.shifts.items(), key=lambda x: int(x[0])):
            timesheet_names = [next((s["short_name"] for s in st.session_state.staff_roster if s["full_name"] == e), e) for e in data.get("timesheet", [])]
            
            go_dict = data.get("graph_only", {})
            if isinstance(go_dict, list):
                go_dict = {name: "" for name in go_dict}
            
            graph_names = []
            for name, dept in go_dict.items():
                short_n = next((s["short_name"] for s in st.session_state.staff_roster if s["full_name"] == name), name)
                dept_str = f" (в {dept})" if dept else ""
                graph_names.append(f"{short_n}{dept_str}")
            
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
            fio = st.text_input("ФИО подработчика (например, Иванов И.)")
            post = st.selectbox("Должность", POSITIONS_LIST)
            
            available_grades = ["—"] + [r["grade"] for r in st.session_state.rates if r["post"] == post and r["grade"]]
            grade = st.selectbox("Разряд", available_grades) if len(available_grades) > 1 else "—"
            
        with c2:
            sub_date = st.date_input("Дата")
            amount = st.number_input("Сумма (руб)", min_value=0.0, value=4500.0, step=100.0)

        if st.button("💾 Добавить подработку", type="primary"):
            if not fio.strip():
                st.error("Укажите ФИО")
            else:
                st.session_state.subordinates.append({
                    "short_name": fio.strip(), 
                    "post": post, "grade": grade,
                    "date": sub_date.strftime("%d.%m.%Y"), "amount": float(amount)
                })
                save_month_to_db()
                st.success("Подработка добавлена!")
                st.rerun()

        st.markdown("---")
        st.markdown("**Список подработок за месяц:**")
        if st.session_state.subordinates:
            df_sub = pd.DataFrame(st.session_state.subordinates)
            if not df_sub.empty:
                st.table(df_sub[["short_name", "post", "grade", "date", "amount"]])
        else:
            st.write("Пока нет записей.")

# ================= TAB 3: СТАВКИ =================
with tab3:
    st.subheader("Глобальный справочник ставок")
    st.info("Эти ставки применяются ко всем месяцам. Измените данные и нажмите «Сохранить».")
    
    df_rates = pd.DataFrame(st.session_state.rates)
    
    edited_rates = st.data_editor(
        df_rates,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "post": st.column_config.TextColumn("Должность", required=True),
            "grade": st.column_config.TextColumn("Разряд"),
            "rate": st.column_config.NumberColumn("Ставка (руб)", min_value=0, step=100, format="%d ₽")
        },
        key="rates_editor"
    )
    
    if st.button("💾 Сохранить глобальные ставки", type="primary"):
        new_rates = []
        for _, row in edited_rates.iterrows():
            if pd.notna(row["post"]) and str(row["post"]).strip() != "":
                new_rates.append({
                    "post": str(row["post"]).strip(),
                    "grade": str(row["grade"]).strip() if pd.notna(row["grade"]) else "",
                    "rate": int(row["rate"]) if pd.notna(row["rate"]) else 0
                })
        st.session_state.rates = new_rates
        save_global_settings()
        st.success("✅ Ставки сохранены глобально!")
        st.rerun()

# ================= TAB 4: КАДРЫ =================
with tab4:
    st.subheader("Глобальный реестр сотрудников (Штат)")
    st.info("Этот список един для всех месяцев. Разряд можно выбрать из выпадающего списка (он берется из вкладки «Ставки»).")
    
    if st.session_state.staff_roster:
        all_grades = set()
        for r in st.session_state.rates:
            if r.get("grade") and str(r["grade"]).strip():
                all_grades.add(str(r["grade"]).strip())
                
        grade_options = ["—"] + sorted(list(all_grades))
        
        df_staff = pd.DataFrame(st.session_state.staff_roster)
        df_staff = df_staff.rename(columns={
            "full_name": "Полное ФИО", 
            "short_name": "Короткое ФИО", 
            "post": "Должность", 
            "grade": "Разряд"
        })
        
        edited_df = st.data_editor(
            df_staff,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Полное ФИО": st.column_config.TextColumn(required=True),
                "Короткое ФИО": st.column_config.TextColumn(required=True),
                "Должность": st.column_config.SelectboxColumn(options=POSITIONS_LIST, required=True),
                "Разряд": st.column_config.SelectboxColumn(options=grade_options)
            },
            key="staff_editor"
        )
        
        if st.button("💾 Сохранить глобальный штат", type="primary"):
            new_roster = []
            for _, row in edited_df.iterrows():
                if pd.notna(row["Полное ФИО"]) and str(row["Полное ФИО"]).strip() != "":
                    new_roster.append({
                        "full_name": str(row["Полное ФИО"]).strip(),
                        "short_name": str(row["Короткое ФИО"]).strip() if pd.notna(row["Короткое ФИО"]) else make_short_name(str(row["Полное ФИО"])),
                        "post": str(row["Должность"]).strip() if pd.notna(row["Должность"]) else "Изготовитель мясных п/ф",
                        "grade": str(row["Разряд"]).strip() if pd.notna(row["Разряд"]) else "—"
                    })
            
            st.session_state.staff_roster = new_roster
            save_global_settings()
            st.success("✅ Глобальный штат сохранен!")
            st.rerun()
    else:
        st.write("Штат пуст. Нажмите кнопку ниже, чтобы загрузить стандартный список.")
        if st.button("Загрузить сотрудников из шаблона"):
            st.session_state.staff_roster = load_initial_staff()
            save_global_settings()
            st.rerun()
