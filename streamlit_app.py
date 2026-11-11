import streamlit as st
import pandas as pd
import io
import datetime
import calendar
import zipfile
import os
import base64
import json
from openpyxl import load_workbook
from supabase import create_client, Client

st.set_page_config(page_title="Учёт ПФ", page_icon="📊", layout="wide")

# ---------- ПОДКЛЮЧЕНИЕ К БАЗЕ ----------
@st.cache_resource
def init_connection():
    try:
        return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    except Exception:
        return None

supabase = init_connection()

# ---------- НАЗВАНИЯ ШАБЛОНОВ ----------
TPL_PF = "tpl_pf.xlsm"
TPL_SUB = "tpl_sub.xlsm"
TPL_GRAPH = "tpl_graph.xlsx"
TPL_SH = "tpl_sh.xlsx"

# ---------- БАЗОВЫЕ НАСТРОЙКИ (Ставки) ----------
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

# ---------- ИНИЦИАЛИЗАЦИЯ ----------
_defaults = {
    "rates": DEFAULT_RATES,
    "active_month": "", "active_year": "", "current_id": None,
    "shifts": {}, "tonnage": {}, "subordinates": [], "staff_list": []
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ---------- ХЕЛПЕРЫ ----------
def get_rate(post, grade):
    for r in st.session_state.rates:
        if r["post"] == post and r["grade"] == grade: return r["rate"]
    return None

def get_positions(): return sorted(set(r["post"] for r in st.session_state.rates))
def get_grades(post): return [r["grade"] for r in st.session_state.rates if r["post"] == post]

def read_staff_from_template():
    if not os.path.exists(TPL_PF): return []
    wb = load_workbook(io.BytesIO(open(TPL_PF, "rb").read()), keep_vba=True)
    res = []
    for sname in wb.sheetnames:
        if sname in ("Итог", "рубкг") or sname.startswith("[Уволен]") or "вакансия" in sname.lower():
            continue
        low = sname.lower()
        if "рузиева" in low and "шазие" in low: post = "Весовщик"
        elif "насруддинов" in low or ("рузиев" in low and "ислом" in low): post = "Грузчик"
        else: post = "Изготовитель мясных п/ф"
        res.append({"sheet": sname, "fio": sname, "post": post})
    return res

# ---------- БАЗА ДАННЫХ ----------
def save_to_db():
    if not supabase or not st.session_state.current_id: return
    try:
        payload = {"shifts": st.session_state.shifts, "tonnage": st.session_state.tonnage, "subordinates": st.session_state.subordinates}
        json_str = json.dumps(payload, ensure_ascii=False)
        data = {"pf_base64": base64.b64encode(json_str.encode('utf-8')).decode("utf-8")}
        supabase.table("app_data").update(data).eq("id", st.session_state.current_id).execute()
    except Exception as e:
        st.error(f"Ошибка БД: {e}")

def load_from_db(record_id):
    if not supabase: return False
    try:
        res = supabase.table("app_data").select("*").eq("id", record_id).execute()
        if res.data and res.data[0]["pf_base64"]:
            raw_json = base64.b64decode(res.data[0]["pf_base64"]).decode('utf-8')
            payload = json.loads(raw_json)
            st.session_state.shifts = payload.get("shifts", {})
            st.session_state.tonnage = payload.get("tonnage", {})
            st.session_state.subordinates = payload.get("subordinates", [])
            st.session_state.current_id = record_id
            st.session_state.active_month, st.session_state.active_year = record_id.split("_")
            return True
    except Exception as e:
        st.error(f"Ошибка загрузки: {e}")
    return False

# ---------- ГЕНЕРАЦИЯ И СКАЧИВАНИЕ ----------
def generate_and_download():
    if not all(os.path.exists(f) for f in [TPL_PF, TPL_SUB, TPL_GRAPH, TPL_SH]):
        st.error("Шаблоны не найдены на GitHub!")
        return None

    zip_buf = io.BytesIO()
    m_name = st.session_state.active_month
    y_val = st.session_state.active_year

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # 1. ТАБЕЛЬ ПОЛУФАБРИКАТЫ (xlsm)
        wb_pf = load_workbook(io.BytesIO(open(TPL_PF, "rb").read()), keep_vba=True)
        wb_pf.calculation.fullCalcOnLoad = True

        for day_str, emps in st.session_state.shifts.items():
            day = int(day_str)
            for emp_sheet in emps:
                if emp_sheet in wb_pf.sheetnames:
                    wb_pf[emp_sheet].cell(row=day + 1, column=2, value=1)
            if day_str in st.session_state.tonnage and "рубкг" in wb_pf.sheetnames:
                wb_pf["рубкг"].cell(row=day + 1, column=6, value=st.session_state.tonnage[day_str])
        
        # Заполняем "Итог" в табеле значениями (без формул)
        if "Итог" in wb_pf.sheetnames:
            ws_it = wb_pf["Итог"]
            for r in range(2, 50):
                fio_cell = ws_it.cell(row=r, column=1).value
                if fio_cell and str(fio_cell).strip().lower() not in ("итого:", "", "none"):
                    sheet_name = str(fio_cell).strip()
                    if sheet_name in wb_pf.sheetnames:
                        shifts_count = sum(1 for d, emps in st.session_state.shifts.items() if sheet_name in emps)
                        rate = 4500
                        for emp in st.session_state.staff_list:
                            if emp["sheet"] == sheet_name:
                                rate = get_rate(emp["post"], emp.get("grade", "")) or 4500
                                break
                        ws_it.cell(row=r, column=2, value=shifts_count * rate)
                        ws_it.cell(row=r, column=3, value=shifts_count)
                elif fio_cell and str(fio_cell).strip().lower() == "итого:":
                    ws_it.cell(row=r, column=2, value=f"=SUM(B2:B{r-1})")
                    ws_it.cell(row=r, column=3, value=f"=SUM(C2:C{r-1})")

        buf_pf = io.BytesIO(); wb_pf.save(buf_pf)
        zf.writestr(f"Полуфабрикаты_{m_name}_{y_val}.xlsm", buf_pf.getvalue())

        # 2. ГРАФИК (xlsx)
        wb_gr = load_workbook(io.BytesIO(open(TPL_GRAPH, "rb").read()))
        wb_gr.calculation.fullCalcOnLoad = True
        ws_gr = wb_gr.active
        ws_gr['A1'] = m_name
        try: ws_gr.title = f"{m_name} {y_val}"
        except: pass
        
        m_idx = MONTHS_RU.index(m_name) + 1
        num_days = calendar.monthrange(int(y_val), m_idx)[1]
        for d in range(1, num_days + 1):
            ws_gr.cell(row=2, column=3 + d, value=d)
        
        block_r = 4
        for r in range(1, 120):
            if "полуфабрикат" in str(ws_gr.cell(row=r, column=1).value or "").lower():
                block_r = r; break
        
        for day_str, emps in st.session_state.shifts.items():
            day = int(day_str)
            col_idx = 3 + day
            for emp_sheet in emps:
                surname = emp_sheet.split()[0].lower()
                for r in range(block_r + 1, block_r + 40):
                    graph_fio = str(ws_gr.cell(row=r, column=2).value or "").lower()
                    if surname in graph_fio:
                        hours = 8 if "весовщик" in emp_sheet.lower() or "грузчик" in emp_sheet.lower() else 11
                        ws_gr.cell(row=r, column=col_idx, value=hours)
                        break
        
        buf_gr = io.BytesIO(); wb_gr.save(buf_gr)
        zf.writestr(f"ГРАФИК_{m_name}_{y_val}.xlsx", buf_gr.getvalue())

        # 3. ПОДРАБОТЧИКИ (xlsm)
        wb_sub = load_workbook(io.BytesIO(open(TPL_SUB, "rb").read()), keep_vba=True)
        wb_sub.calculation.fullCalcOnLoad = True
        
        tpl_sub = wb_sub[wb_sub.sheetnames[0] if wb_sub.sheetnames[0] != 'Итог' else wb_sub.sheetnames[1]]
        
        subs_by_fio = {}
        for sub in st.session_state.subordinates:
            title = sub["fio"][:31]
            if title not in wb_sub.sheetnames:
                new_ws = wb_sub.copy_worksheet(tpl_sub); new_ws.title = title
                for r in range(2, 33): 
                    new_ws.cell(row=r, column=2).value = None
                    new_ws.cell(row=r, column=3).value = None # Очищаем и суммы
            
            day = int(sub["date"].split(".")[0])
            wb_sub[title].cell(row=day + 1, column=2, value=1)
            wb_sub[title].cell(row=day + 1, column=3, value=sub["amount"])
            
            if title not in subs_by_fio:
                subs_by_fio[title] = {"amount": 0.0, "shifts": 0}
            subs_by_fio[title]["amount"] += float(sub["amount"])
            subs_by_fio[title]["shifts"] += 1
        
        # Заполняем "Итог" в подработчиках готовыми цифрами
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
                    ws_it.cell(row=r, column=2, value=data["amount"])
                    ws_it.cell(row=r, column=3, value=data["shifts"])

            if itogo_row < 50:
                ws_it.cell(row=itogo_row, column=2, value=f"=SUM(B2:B{itogo_row-1})")
                ws_it.cell(row=itogo_row, column=3, value=f"=SUM(C2:C{itogo_row-1})")
                
        buf_sub = io.BytesIO(); wb_sub.save(buf_sub)
        zf.writestr(f"Подработчики_{m_name}_{y_val}.xlsm", buf_sub.getvalue())

        # 4. ШАШИНА (xlsx)
        wb_sh = load_workbook(io.BytesIO(open(TPL_SH, "rb").read()))
        wb_sh.calculation.fullCalcOnLoad = True
        if "ПФ" in wb_sh.sheetnames:
            ws_sh = wb_sh["ПФ"]
            start_row = 4
            for i, sub in enumerate(st.session_state.subordinates):
                r = start_row + i
                ws_sh.cell(row=r, column=2, value=i+1)
                ws_sh.cell(row=r, column=3, value=sub["fio"])
                ws_sh.cell(row=r, column=4, value=sub["post"])
                ws_sh.cell(row=r, column=5, value=sub["grade"])
                ws_sh.cell(row=r, column=6, value=11)
                ws_sh.cell(row=r, column=8, value=sub["date"])
                ws_sh.cell(row=r, column=9, value=sub["amount"])
        buf_sh = io.BytesIO(); wb_sh.save(buf_sh)
        zf.writestr(f"Шашина_{m_name}_{y_val}.xlsx", buf_sh.getvalue())

    return zip_buf.getvalue()

# ---------- ИНТЕРФЕЙС ----------
st.sidebar.header("🗓️ Текущий период")
if st.session_state.active_month:
    st.sidebar.success(f"### {st.session_state.active_month} {st.session_state.active_year}")
    if st.sidebar.button("⬇️ Сгенерировать и скачать файлы (ZIP)", type="primary"):
        zip_data = generate_and_download()
        if zip_data:
            st.sidebar.download_button(
                label="💾 Сохранить на компьютер", 
                data=zip_data, 
                file_name=f"Итоги_{st.session_state.active_month}_{st.session_state.active_year}.zip", 
                mime="application/zip"
            )
else:
    st.sidebar.warning("### Месяц не выбран\nСоздайте новый или выберите старый.")

st.sidebar.markdown("---")
if supabase:
    st.sidebar.subheader("📂 Открыть месяц")
    try:
        res = supabase.table("app_data").select("id").execute()
        saved = [r["id"] for r in res.data]
        if saved:
            saved.sort(reverse=True)
            sel = st.sidebar.selectbox("Выберите месяц", saved)
            if st.sidebar.button("Загрузить"):
                if load_from_db(sel):
                    st.sidebar.success("Загружено!")
                    st.rerun()
    except: pass

# ---------- ОСНОВНОЙ ЭКРАН ----------
st.title(f"📊 Отделение Полуфабрикатов — {st.session_state.active_month} {st.session_state.active_year}" if st.session_state.active_month else "📊 Учёт выработки")

tab0, tab1, tab2, tab3 = st.tabs(["✨ Создать месяц", "📅 Табель смен", "👷 Подработчики", "💰 Ставки"])

# ================= TAB 0: СОЗДАНИЕ =================
with tab0:
    st.subheader("Новый месяц")
    c1, c2 = st.columns(2)
    with c1: new_month = st.selectbox("Месяц", MONTHS_RU, index=datetime.date.today().month - 1)
    with c2: new_year = st.number_input("Год", min_value=2024, max_value=2035, value=datetime.date.today().year)

    if st.button("✨ Создать месяц", type="primary"):
        record_id = f"{new_month}_{new_year}"
        if not supabase:
            st.error("Нет подключения к БД.")
        elif not os.path.exists(TPL_PF):
            st.error("Шаблоны не загружены на GitHub!")
        else:
            try:
                existing = supabase.table("app_data").select("id").eq("id", record_id).execute()
                if existing.data:
                    st.error(f"Месяц {record_id} уже существует! Удалите его в Supabase.")
                else:
                    st.session_state.shifts = {}
                    st.session_state.tonnage = {}
                    st.session_state.subordinates = []
                    st.session_state.active_month = new_month
                    st.session_state.active_year = str(new_year)
                    st.session_state.current_id = record_id
                    st.session_state.staff_list = read_staff_from_template()
                    
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
        if not st.session_state.staff_list:
            st.session_state.staff_list = read_staff_from_template()

        st.subheader("Отметка смен")
        c1, c2 = st.columns(2)
        with c1: shift_date = st.date_input("Дата смены", datetime.date.today())
        with c2: tonnage = st.number_input("Тоннаж цеха (кг)", min_value=0.0, value=0.0, step=1.0)

        st.markdown("**Кого отметить?**")
        selected = {}
        for emp in st.session_state.staff_list:
            selected[emp["sheet"]] = st.checkbox(f"{emp['fio']} — {emp['post']}", key=f"chk_{emp['sheet']}")

        if st.button("💾 Провести смену", type="primary"):
            day_str = str(shift_date.day)
            if day_str not in st.session_state.shifts:
                st.session_state.shifts[day_str] = []
            
            for emp_sheet, checked in selected.items():
                if checked and emp_sheet not in st.session_state.shifts[day_str]:
                    st.session_state.shifts[day_str].append(emp_sheet)
            
            if tonnage > 0:
                st.session_state.tonnage[day_str] = float(tonnage)
            
            save_to_db()
            st.success(f"Смена на {shift_date.strftime('%d.%m.%Y')} сохранена в базу!")
            st.rerun()

        st.markdown("---")
        st.markdown("**Кто уже отмечен в этом месяце:**")
        for day, emps in sorted(st.session_state.shifts.items(), key=lambda x: int(x[0])):
            st.write(f"**{day} число:** {', '.join(emps)}")

# ================= TAB 2: ПОДРАБОТЧИКИ =================
with tab2:
    if not st.session_state.active_month:
        st.info("👈 Сначала создайте месяц.")
    else:
        st.subheader("Внесение подработки")
        c1, c2 = st.columns(2)
        with c1:
            fio = st.text_input("ФИО")
            post = st.selectbox("Должность", get_positions())
            grades = get_grades(post)
            grade = st.selectbox("Разряд", grades) if grades else ""
        with c2:
            sub_date = st.date_input("Дата")
            amount = st.number_input("Сумма (руб)", min_value=0.0, value=4500.0, step=100.0)

        if st.button("💾 Добавить подработку", type="primary"):
            if not fio.strip():
                st.error("Укажите ФИО")
            else:
                st.session_state.subordinates.append({
                    "fio": fio.strip(), "post": post, "grade": grade,
                    "date": sub_date.strftime("%d.%m.%Y"), "amount": float(amount)
                })
                save_to_db()
                st.success("Подработка добавлена в базу!")
                st.rerun()

        st.markdown("---")
        st.markdown("**Список подработок за месяц:**")
        if st.session_state.subordinates:
            st.table(pd.DataFrame(st.session_state.subordinates))
        else:
            st.write("Пока нет записей.")

# ================= TAB 3: СТАВКИ =================
with tab3:
    st.subheader("Справочник ставок")
    df = pd.DataFrame(st.session_state.rates)
    df.columns = ["Должность", "Разряд", "Ставка"]
    st.data_editor(df, use_container_width=True, disabled=True)
