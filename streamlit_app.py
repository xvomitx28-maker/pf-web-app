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

# ---------- ИНИЦИАЛИЗАЦИЯ ----------
_defaults = {
    "active_month": "", "active_year": "", "current_id": None,
    "shifts": {}, "tonnage": {}, "subordinates": [], "staff_roster": []
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

MONTHS_RU = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

# ---------- ХЕЛПЕРЫ ----------
def clean_name(name):
    """Удаляет все лишние символы, оставляя только буквы для надежного поиска"""
    return re.sub(r'[^а-яё]', '', str(name).lower())

def make_short_name(full_name):
    """Автоматически создает короткое имя: Рузиева Шазие -> Рузиева Ш."""
    parts = str(full_name).strip().split()
    if len(parts) > 1:
        return f"{parts[0]} {parts[1][0]}."
    return parts[0] if parts else ""

def load_initial_staff():
    """Если штат пустой, пробуем вытащить его из шаблона Табеля"""
    if not os.path.exists(TPL_PF): return []
    wb = load_workbook(io.BytesIO(open(TPL_PF, "rb").read()), keep_vba=True)
    res = []
    for sname in wb.sheetnames:
        if sname in ("Итог", "рубкг") or sname.startswith("[Уволен]") or "вакансия" in sname.lower():
            continue
        low = sname.lower()
        if "рузиева" in low: post = "Весовщик"
        elif "насруддинов" in low or "рузиев и" in low: post = "Грузчик"
        else: post = "Изготовитель мясных п/ф"
        
        res.append({
            "full_name": sname, 
            "short_name": make_short_name(sname), 
            "post": post
        })
    return res

# ---------- БАЗА ДАННЫХ ----------
def save_to_db():
    if not supabase or not st.session_state.current_id: return
    try:
        payload = {
            "shifts": st.session_state.shifts, 
            "tonnage": st.session_state.tonnage, 
            "subordinates": st.session_state.subordinates,
            "staff_roster": st.session_state.staff_roster
        }
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
            st.session_state.staff_roster = payload.get("staff_roster", [])
            
            if not st.session_state.staff_roster:
                st.session_state.staff_roster = load_initial_staff()
                
            st.session_state.current_id = record_id
            st.session_state.active_month, st.session_state.active_year = record_id.split("_")
            return True
    except Exception as e:
        st.error(f"Ошибка загрузки: {e}")
    return False

# ---------- ГЕНЕРАЦИЯ ----------
def generate_and_download():
    zip_buf = io.BytesIO()
    m_name = st.session_state.active_month
    y_val = st.session_state.active_year

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # 1. ТАБЕЛЬ (xlsm) - Единый кадровый шаблон
        wb_pf = load_workbook(io.BytesIO(open(TPL_PF, "rb").read()), keep_vba=True)
        wb_pf.calculation.fullCalcOnLoad = True

        for day_str, emps in st.session_state.shifts.items():
            day = int(day_str)
            for emp_full_name in emps:
                if emp_full_name in wb_pf.sheetnames:
                    wb_pf[emp_full_name].cell(row=day + 1, column=2, value=1)
            if day_str in st.session_state.tonnage and "рубкг" in wb_pf.sheetnames:
                wb_pf["рубкг"].cell(row=day + 1, column=6, value=st.session_state.tonnage[day_str])
        
        if "Итог" in wb_pf.sheetnames:
            ws_it = wb_pf["Итог"]
            for r in range(2, 50):
                fio_cell = ws_it.cell(row=r, column=1).value
                if fio_cell and str(fio_cell).strip().lower() not in ("итого:", "", "none"):
                    sheet_name = str(fio_cell).strip()
                    if sheet_name in wb_pf.sheetnames:
                        shifts_count = sum(1 for d, emps in st.session_state.shifts.items() if sheet_name in emps)
                        ws_it.cell(row=r, column=2, value=shifts_count * 4500) 
                        ws_it.cell(row=r, column=3, value=shifts_count)
                elif fio_cell and str(fio_cell).strip().lower() == "итого:":
                    ws_it.cell(row=r, column=2, value=f"=SUM(B2:B{r-1})")
                    ws_it.cell(row=r, column=3, value=f"=SUM(C2:C{r-1})")

        buf_pf = io.BytesIO(); wb_pf.save(buf_pf)
        zf.writestr(f"Полуфабрикаты_{m_name}_{y_val}.xlsm", buf_pf.getvalue())

        # 2. ГРАФИК (xlsx) - Единый кадровый шаблон
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
            for emp_full_name in emps:
                clean_surname = clean_name(emp_full_name)
                for r in range(block_r + 1, block_r + 40):
                    graph_fio = clean_name(ws_gr.cell(row=r, column=2).value or "")
                    if clean_surname in graph_fio:
                        post = next((s["post"] for s in st.session_state.staff_roster if s["full_name"] == emp_full_name), "")
                        hours = 8 if post in ("Весовщик", "Грузчик") else 11
                        ws_gr.cell(row=r, column=col_idx, value=hours)
                        break
        
        buf_gr = io.BytesIO(); wb_gr.save(buf_gr)
        zf.writestr(f"ГРАФИК_{m_name}_{y_val}.xlsx", buf_gr.getvalue())

        # 3. ПОДРАБОТЧИКИ (xlsm) - Только пофамильно, динамически
        wb_sub = load_workbook(io.BytesIO(open(TPL_SUB, "rb").read()), keep_vba=True)
        wb_sub.calculation.fullCalcOnLoad = True
        
        # Находим лист-шаблон (все, что не Итог)
        tpl_sheet_name = None
        for s in wb_sub.sheetnames:
            if s != "Итог":
                tpl_sheet_name = s
                break
                
        if not tpl_sheet_name:
            st.error("В файле tpl_sub.xlsm нет листа-шаблона! Оставьте один пустой лист.")
            return None
            
        # Удаляем все старые листы (кроме Итога и Шаблона)
        for s in list(wb_sub.sheetnames):
            if s not in ("Итог", tpl_sheet_name):
                del wb_sub[s]
                
        subs_by_fio = {}
        for sub in st.session_state.subordinates:
            title = sub["short_name"][:31]
            if title not in wb_sub.sheetnames:
                new_ws = wb_sub.copy_worksheet(wb_sub[tpl_sheet_name])
                new_ws.title = title
                for r in range(2, 33): 
                    new_ws.cell(row=r, column=2).value = None
                    new_ws.cell(row=r, column=3).value = None
            
            day = int(sub["date"].split(".")[0])
            wb_sub[title].cell(row=day + 1, column=2, value=1)
            wb_sub[title].cell(row=day + 1, column=3, value=sub["amount"])
            
            if title not in subs_by_fio:
                subs_by_fio[title] = {"amount": 0.0, "shifts": 0}
            subs_by_fio[title]["amount"] += float(sub["amount"])
            subs_by_fio[title]["shifts"] += 1
            
        # Удаляем сам лист-шаблон, чтобы он не попал в итоговый файл
        if tpl_sheet_name in wb_sub.sheetnames:
            del wb_sub[tpl_sheet_name]
        
        # Заполняем Итог
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

        # 4. ШАШИНА (xlsx) - Только пофамильно
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

tab0, tab1, tab2, tab3, tab4 = st.tabs(["✨ Создать месяц", "📅 Табель смен", "👷 Подработчики", "💰 Ставки", "👥 Кадры"])

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
                    st.session_state.staff_roster = load_initial_staff()
                    
                    empty_payload = json.dumps({
                        "shifts": {}, "tonnage": {}, "subordinates": [], 
                        "staff_roster": st.session_state.staff_roster
                    })
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
        st.subheader("Отметка смен (Штат)")
        c1, c2 = st.columns(2)
        with c1: shift_date = st.date_input("Дата смены", datetime.date.today())
        with c2: tonnage = st.number_input("Тоннаж цеха (кг)", min_value=0.0, value=0.0, step=1.0)

        st.markdown("**Кого отметить?**")
        selected = {}
        for emp in st.session_state.staff_roster:
            selected[emp["full_name"]] = st.checkbox(f"{emp['short_name']} — {emp['post']}", key=f"chk_{emp['full_name']}")

        if st.button("💾 Провести смену", type="primary"):
            day_str = str(shift_date.day)
            if day_str not in st.session_state.shifts:
                st.session_state.shifts[day_str] = []
            
            for emp_full_name, checked in selected.items():
                if checked and emp_full_name not in st.session_state.shifts[day_str]:
                    st.session_state.shifts[day_str].append(emp_full_name)
            
            if tonnage > 0:
                st.session_state.tonnage[day_str] = float(tonnage)
            
            save_to_db()
            st.success(f"Смена на {shift_date.strftime('%d.%m.%Y')} сохранена в базу!")
            st.rerun()

        st.markdown("---")
        st.markdown("**Кто уже отмечен в этом месяце:**")
        for day, emps in sorted(st.session_state.shifts.items(), key=lambda x: int(x[0])):
            short_names = [next((s["short_name"] for s in st.session_state.staff_roster if s["full_name"] == e), e) for e in emps]
            st.write(f"**{day} число:** {', '.join(short_names)}")

# ================= TAB 2: ПОДРАБОТЧИКИ =================
with tab2:
    if not st.session_state.active_month:
        st.info("👈 Сначала создайте месяц.")
    else:
        st.subheader("Внесение подработки (Пофамильно)")
        c1, c2 = st.columns(2)
        with c1:
            fio = st.text_input("ФИО подработчика (например, Иванов И.И. или просто Иванов И.)")
            post = st.text_input("Должность", value="Изготовитель мясных п/ф")
            grade = st.text_input("Разряд", value="")
        with c2:
            sub_date = st.date_input("Дата")
            amount = st.number_input("Сумма (руб)", min_value=0.0, value=4500.0, step=100.0)

        if st.button("💾 Добавить подработку", type="primary"):
            if not fio.strip():
                st.error("Укажите ФИО")
            else:
                st.session_state.subordinates.append({
                    "short_name": fio.strip(), # Для подработок используем то, что ввели
                    "post": post, "grade": grade,
                    "date": sub_date.strftime("%d.%m.%Y"), "amount": float(amount)
                })
                save_to_db()
                st.success("Подработка добавлена!")
                st.rerun()

        st.markdown("---")
        st.markdown("**Список подработок за месяц:**")
        if st.session_state.subordinates:
            df_sub = pd.DataFrame(st.session_state.subordinates)
            if not df_sub.empty:
                st.table(df_sub[["short_name", "post", "date", "amount"]])
        else:
            st.write("Пока нет записей.")

# ================= TAB 3: СТАВКИ =================
with tab3:
    st.subheader("Справочник ставок")
    st.info("Ставки задаются в коде программы. Если нужно их поменять, обратитесь к разработчику.")

# ================= TAB 4: КАДРЫ =================
with tab4:
    st.subheader("Управление штатом")
    st.info("Задайте здесь эталонные ФИО. Программа использует Полное ФИО для Графика и Табеля, а Короткое — для удобства на экране.")
    
    # Форма добавления
    with st.expander("➕ Добавить нового сотрудника в штат"):
        c1, c2, c3 = st.columns(3)
        with c1: new_full = st.text_input("Полное ФИО (как в Табеле/Графике)")
        with c2: new_short = st.text_input("Короткое ФИО (для списков)")
        with c3: new_post = st.selectbox("Должность", ["Изготовитель мясных п/ф", "Весовщик", "Грузчик", "Формовщик", "Жиловщик"])
        
        if st.button("Добавить в штат"):
            if new_full and new_short:
                st.session_state.staff_roster.append({
                    "full_name": new_full.strip(), "short_name": new_short.strip(), "post": new_post
                })
                save_to_db()
                st.success("Сотрудник добавлен!")
                st.rerun()
            else:
                st.error("Заполните все поля!")
                
    # Таблица штата
    st.markdown("**Текущий штат:**")
    if st.session_state.staff_roster:
        df_staff = pd.DataFrame(st.session_state.staff_roster)
        df_staff = df_staff.rename(columns={"full_name": "Полное ФИО", "short_name": "Короткое ФИО", "post": "Должность"})
        
        # Удаление сотрудника
        to_delete = st.selectbox("Выберите сотрудника для увольнения:", [s["full_name"] for s in st.session_state.staff_roster])
        if st.button("🚫 Удалить из штата"):
            st.session_state.staff_roster = [s for s in st.session_state.staff_roster if s["full_name"] != to_delete]
            # Также удаляем его смены из текущего месяца
            for day in st.session_state.shifts:
                if to_delete in st.session_state.shifts[day]:
                    st.session_state.shifts[day].remove(to_delete)
            save_to_db()
            st.success(f"{to_delete} уволен.")
            st.rerun()
            
        st.dataframe(df_staff, use_container_width=True)
    else:
        st.write("Штат пуст. Добавьте сотрудников.")
