import streamlit as st
import pandas as pd
import io
import datetime
import calendar
import zipfile
import os
import base64
from openpyxl import load_workbook
from supabase import create_client, Client

st.set_page_config(page_title="Учёт ПФ", page_icon="📊", layout="wide")

# ---------- ПОДКЛЮЧЕНИЕ К БАЗЕ ДАННЫХ ----------
@st.cache_resource
def init_connection():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception as e:
        return None

supabase = init_connection()

# ---------- НАЗВАНИЯ ШАБЛОНОВ НА GITHUB ----------
TPL_PF = "tpl_pf.xlsm"
TPL_SUB = "tpl_sub.xlsm"
TPL_GRAPH = "tpl_graph.xlsx"
TPL_SH = "tpl_sh.xlsx"

# ---------- INIT SESSION ----------
_defaults = {
    "rates": [
        {"post": "Весовщик", "grade": "", "rate": 5000},
        {"post": "Грузчик", "grade": "", "rate": 4300},
        {"post": "Изготовитель мясных п/ф", "grade": "1 разряд", "rate": 3700},
        {"post": "Изготовитель мясных п/ф", "grade": "2 разряд", "rate": 4500},
        {"post": "Формовщик", "grade": "", "rate": 4500},
        {"post": "Жиловщик", "grade": "", "rate": 4500},
        {"post": "Оператор линии", "grade": "", "rate": 4500},
        {"post": "Бригадир", "grade": "", "rate": 4500},
    ],
    "staff_meta": {},
    "pf_bytes": None, "sub_bytes": None, "graph_bytes": None, "sh_bytes": None,
    "pf_name": "Полуфабрикаты.xlsm", "sub_name": "Подработчики_ПФ.xlsm",
    "graph_name": "ГРАФИК.xlsx", "sh_name": "Шашина.xlsx",
    "active_month": "", "active_year": "",
    "current_record_id": None
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

MONTHS_RU = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

# ---------- HELPERS ----------
def get_rate(post, grade):
    for r in st.session_state.rates:
        if r["post"] == post and r["grade"] == grade: return r["rate"]
    return None

def get_positions():
    return sorted(set(r["post"] for r in st.session_state.rates))

def get_grades(post):
    return [r["grade"] for r in st.session_state.rates if r["post"] == post]

def infer_post_grade(sname, sheet_rate):
    low = sname.lower()
    if "рузиева" in low and "шазие" in low: return ("Весовщик", "")
    if "насруддинов" in low or ("рузиев" in low and "ислом" in low): return ("Грузчик", "")
    if sheet_rate == 3700 or "никулин" in low or "ермакова" in low: return ("Изготовитель мясных п/ф", "1 разряд")
    return ("Изготовитель мясных п/ф", "2 разряд")

def load_wb(data):
    return load_workbook(io.BytesIO(data), keep_vba=True)

def wb_to_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def read_staff(wb):
    result = []
    for sname in wb.sheetnames:
        if sname in ("Итог", "рубкг"): continue
        ws = wb[sname]
        sheet_rate = None
        for r in range(2, 33):
            v = ws.cell(row=r, column=3).value
            if isinstance(v, (int, float)) and v > 0:
                sheet_rate = int(v); break
        meta = st.session_state.staff_meta.get(sname)
        if not meta:
            post, grade = infer_post_grade(sname, sheet_rate)
            meta = {"post": post, "grade": grade}
            st.session_state.staff_meta[sname] = meta
        post, grade = meta["post"], meta.get("grade", "")
        rate = get_rate(post, grade) or sheet_rate or 4500
        if not sname.startswith("[Уволен]") and "вакансия" not in sname.lower():
            result.append({"sheet": sname, "fio": sname, "post": post, "grade": grade, "rate": rate})
    return result

# ---------- ФУНКЦИИ БАЗЫ ДАННЫХ ----------
def save_to_db():
    if not supabase or not st.session_state.current_record_id: return
    try:
        data = {
            "pf_base64": base64.b64encode(st.session_state.pf_bytes).decode("utf-8") if st.session_state.pf_bytes else None,
            "sub_base64": base64.b64encode(st.session_state.sub_bytes).decode("utf-8") if st.session_state.sub_bytes else None,
            "graph_base64": base64.b64encode(st.session_state.graph_bytes).decode("utf-8") if st.session_state.graph_bytes else None,
            "sh_base64": base64.b64encode(st.session_state.sh_bytes).decode("utf-8") if st.session_state.sh_bytes else None,
        }
        supabase.table("app_data").update(data).eq("id", st.session_state.current_record_id).execute()
    except Exception as e:
        st.error(f"Ошибка сохранения в базу: {e}")

def load_from_db(record_id):
    if not supabase: return False
    try:
        res = supabase.table("app_data").select("*").eq("id", record_id).execute()
        if res.data:
            row = res.data[0]
            st.session_state.pf_bytes = base64.b64decode(row["pf_base64"]) if row["pf_base64"] else None
            st.session_state.sub_bytes = base64.b64decode(row["sub_base64"]) if row["sub_base64"] else None
            st.session_state.graph_bytes = base64.b64decode(row["graph_base64"]) if row["graph_base64"] else None
            st.session_state.sh_bytes = base64.b64decode(row["sh_base64"]) if row["sh_base64"] else None
            st.session_state.current_record_id = record_id
            st.session_state.active_month, st.session_state.active_year = record_id.split("_")
            return True
    except Exception as e:
        st.error(f"Ошибка загрузки из базы: {e}")
    return False

# ---------- SIDEBAR ----------
st.sidebar.header("🗓️ Текущий период")
if st.session_state.active_month and st.session_state.active_year:
    st.sidebar.success(f"### {st.session_state.active_month} {st.session_state.active_year}")
    st.sidebar.caption("Все изменения автоматически сохраняются в базе.")
    
    zip_prog_buf = io.BytesIO()
    with zipfile.ZipFile(zip_prog_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        if st.session_state.pf_bytes: zf.writestr(st.session_state.pf_name, st.session_state.pf_bytes)
        if st.session_state.graph_bytes: zf.writestr(st.session_state.graph_name, st.session_state.graph_bytes)
        if st.session_state.sub_bytes: zf.writestr(st.session_state.sub_name, st.session_state.sub_bytes)
        if st.session_state.sh_bytes: zf.writestr(st.session_state.sh_name, st.session_state.sh_bytes)
    
    st.sidebar.download_button(
        label="⬇️ Скачать готовые файлы (ZIP)", 
        data=zip_prog_buf.getvalue(), 
        file_name=f"Итоги_{st.session_state.active_month}_{st.session_state.active_year}.zip", 
        mime="application/zip",
        type="primary"
    )
else:
    st.sidebar.warning("### Месяц не выбран\nСоздайте новый или выберите старый.")

st.sidebar.markdown("---")
if supabase:
    st.sidebar.subheader("📂 Открыть сохраненный месяц")
    try:
        res = supabase.table("app_data").select("id").execute()
        saved_months = [row["id"] for row in res.data]
        if saved_months:
            saved_months.sort(reverse=True)
            selected_month = st.sidebar.selectbox("Выберите месяц", saved_months)
            if st.sidebar.button("Загрузить в программу"):
                if load_from_db(selected_month):
                    st.sidebar.success("Данные загружены!")
                    st.rerun()
        else:
            st.sidebar.caption("Сохраненных месяцев пока нет.")
    except Exception as e:
        st.sidebar.error(f"Ошибка БД: {e}")

# ---------- MAIN ----------
if st.session_state.active_month:
    st.title(f"📊 Отделение Полуфабрикатов — {st.session_state.active_month} {st.session_state.active_year}")
else:
    st.title("📊 Учёт выработки — Отделение Полуфабрикатов")

tab0, tab1, tab2, tab3, tab4 = st.tabs(
    ["✨ Создать месяц", "📅 Табель смен", "👷 Подработчики", "💰 Ставки и разряды", "👥 Кадры"]
)

# ================= TAB 0: СОЗДАНИЕ МЕСЯЦА =================
with tab0:
    st.subheader("Формирование нового месяца")
    
    c1, c2 = st.columns(2)
    with c1:
        new_month = st.selectbox("Месяц", MONTHS_RU, index=datetime.date.today().month - 1)
    with c2:
        new_year = st.number_input("Год", min_value=2024, max_value=2035, value=datetime.date.today().year)

    if st.button("✨ Создать месяц и сохранить в базу", type="primary"):
        record_id = f"{new_month}_{new_year}"
        if not all([os.path.exists(TPL_PF), os.path.exists(TPL_SUB), os.path.exists(TPL_GRAPH), os.path.exists(TPL_SH)]):
            st.error("❌ Файлы шаблонов не найдены на GitHub.")
        elif not supabase:
            st.error("❌ Нет подключения к базе данных. Проверьте настройки Secrets.")
        else:
            with st.spinner("Создаем файлы..."):
                try:
                    existing = supabase.table("app_data").select("id").eq("id", record_id).execute()
                    if existing.data:
                        st.error(f"❌ Месяц {record_id} уже существует! Удалите его из базы Supabase (Table Editor) и попробуйте снова.")
                    else:
                        m_idx = MONTHS_RU.index(new_month) + 1
                        num_days = calendar.monthrange(new_year, m_idx)[1]
                        
                        wb_pf = load_wb(open(TPL_PF, "rb").read())
                        if "рубкг" in wb_pf.sheetnames:
                            ws_rub = wb_pf["рубкг"]
                            for day_i in range(1, 32):
                                r = 1 + day_i
                                ws_rub.cell(row=r, column=1).value = datetime.datetime(new_year, m_idx, day_i) if day_i <= num_days else None
                                for c in [6, 8, 9]:
                                    cell = ws_rub.cell(row=r, column=c)
                                    if cell.data_type != 'f': cell.value = None
                        for sname in wb_pf.sheetnames:
                            if sname not in ["Итог", "рубкг"] and not sname.startswith("[Уволен]"):
                                ws_e = wb_pf[sname]
                                for r in range(2, 33):
                                    for c in [2, 4, 5, 6, 7]: 
                                        cell = ws_e.cell(row=r, column=c)
                                        if cell.data_type != 'f': cell.value = None
                        st.session_state.pf_bytes = wb_to_bytes(wb_pf)
                        st.session_state.pf_name = f"Полуфабрикаты_{new_month}_{new_year}.xlsm"

                        wb_sub = load_wb(open(TPL_SUB, "rb").read())
                        for sname in wb_sub.sheetnames:
                            if sname != "Итог":
                                ws_s = wb_sub[sname]
                                for r in range(2, 33):
                                    for c in [2, 3]:
                                        cell = ws_s.cell(row=r, column=c)
                                        if cell.data_type != 'f': cell.value = None
                        st.session_state.sub_bytes = wb_to_bytes(wb_sub)
                        st.session_state.sub_name = f"Подработчики_{new_month}_{new_year}.xlsm"

                        wb_graph = load_workbook(io.BytesIO(open(TPL_GRAPH, "rb").read()))
                        ws_gr = wb_graph.active
                        try: ws_gr.title = f"{new_month} {new_year}"
                        except: pass
                        ws_gr['A1'] = new_month
                        block_r = 4
                        for r in range(1, min(120, ws_gr.max_row + 1)):
                            for c in range(1, 4):
                                if "полуфабрикат" in str(ws_gr.cell(row=r, column=c).value or "").lower():
                                    block_r = r; break
                            if block_r > 4: break
                        for r in range(block_r + 1, min(block_r + 40, ws_gr.max_row + 1)):
                            fio_v = str(ws_gr.cell(row=r, column=2).value or "").strip().lower()
                            if "отделение" in fio_v and "полуфабрикат" not in fio_v: break
                            for c in range(4, 35):
                                cell_d = ws_gr.cell(row=r, column=c)
                                if cell_d.data_type != 'f' and not str(cell_d.value or '').startswith('='):
                                    cell_d.value = None
                        buf_graph = io.BytesIO(); wb_graph.save(buf_graph)
                        st.session_state.graph_bytes = buf_graph.getvalue()
                        st.session_state.graph_name = f"ГРАФИК_{new_month}_{new_year}.xlsx"

                        wb_sh = load_workbook(io.BytesIO(open(TPL_SH, "rb").read()))
                        for sname in wb_sh.sheetnames:
                            ws_s = wb_sh[sname]
                            for r in range(4, min(150, ws_s.max_row + 1)):
                                for c in range(2, 11):
                                    cell = ws_s.cell(row=r, column=c)
                                    if cell.data_type != 'f' and not str(cell.value or '').startswith('='): 
                                        cell.value = None
                        buf_sh = io.BytesIO(); wb_sh.save(buf_sh)
                        st.session_state.sh_bytes = buf_sh.getvalue()
                        st.session_state.sh_name = f"Шашина_{new_month}_{new_year}.xlsx"

                        st.session_state.active_month = new_month
                        st.session_state.active_year = str(new_year)
                        st.session_state.current_record_id = record_id
                        
                        data = {
                            "id": record_id,
                            "pf_base64": base64.b64encode(st.session_state.pf_bytes).decode("utf-8"),
                            "sub_base64": base64.b64encode(st.session_state.sub_bytes).decode("utf-8"),
                            "graph_base64": base64.b64encode(st.session_state.graph_bytes).decode("utf-8"),
                            "sh_base64": base64.b64encode(st.session_state.sh_bytes).decode("utf-8"),
                        }
                        supabase.table("app_data").insert(data).execute()
                        
                        st.success(f"✅ Месяц {new_month} {new_year} создан (формулы сохранены) и записан в базу!")
                        st.rerun()
                except Exception as e:
                    st.error(f"Ошибка: {e}")


# ================= TAB 1: ТАБЕЛЬ СМЕН =================
with tab1:
    if st.session_state.pf_bytes is None:
        st.info("👈 Создайте или загрузите месяц в боковой панели.")
    else:
        wb_pf = load_wb(st.session_state.pf_bytes)
        wb_graph = load_workbook(io.BytesIO(st.session_state.graph_bytes)) if st.session_state.graph_bytes else None
            
        staff = read_staff(wb_pf)
        
        # --- БЛОК ОТЛАДКИ ---
        with st.expander("🛠️ Проверка чтения файлов (Отладка)"):
            st.write(f"Найдено листов (сотрудников): **{len(staff)}**")
            if len(staff) > 0:
                st.write("Первые три сотрудника:")
                for s in staff[:3]:
                    st.write(f"- Лист: `{s['sheet']}`, ФИО: `{s['fio']}`, Должность: `{s['post']}`")
            if wb_graph:
                st.write("График прочитан успешно.")
            else:
                st.write("⚠️ График НЕ прочитан!")

        st.subheader("Отметка смен")
        c1, c2 = st.columns(2)
        with c1: shift_date = st.date_input("Дата смены", datetime.date.today(), key="shift_date")
        with c2: tonnage = st.number_input("Тоннаж цеха (кг)", min_value=0.0, value=0.0, step=1.0, key="tonnage")

        st.markdown("**Кого отметить?**")
        selected = {}
        for emp in staff:
            gr = f" ({emp['grade']})" if emp["grade"] else ""
            label = f"{emp['fio']} — {emp['post']}{gr} — {emp['rate']} ₽"
            selected[emp["sheet"]] = st.checkbox(label, key=f"chk_{emp['sheet']}")

        if st.button("💾 Провести смену", type="primary"):
            day = shift_date.day
            n = 0
            for sname, checked in selected.items():
                if checked and sname in wb_pf.sheetnames:
                    wb_pf[sname].cell(row=day + 1, column=2, value=1)
                    n += 1
            if tonnage > 0 and "рубкг" in wb_pf.sheetnames:
                wb_pf["рубкг"].cell(row=day + 1, column=6, value=float(tonnage))
                
            if wb_graph:
                ws_gr = wb_graph.active
                day_col = None
                for r_check in [2, 3, 1]:
                    for c_idx in range(4, 36):
                        val = ws_gr.cell(row=r_check, column=c_idx).value
                        if val is not None and str(val).strip() == str(day): 
                            day_col = c_idx
                            break
                    if day_col: break
                
                if day_col:
                    block_r = 4
                    for r in range(1, 120):
                        if "полуфабрикат" in str(ws_gr.cell(row=r, column=1).value or "").lower():
                            block_r = r; break
                    for sname, checked in selected.items():
                        if not checked: continue
                        surname = sname.split()[0].lower()
                        for r in range(block_r + 1, block_r + 40):
                            graph_fio = str(ws_gr.cell(row=r, column=2).value or "").lower()
                            if surname in graph_fio:
                                info = next((e for e in staff if e["sheet"] == sname), {})
                                hours = 8 if info.get("post") in ("Весовщик", "Грузчик") else 11
                                ws_gr.cell(row=r, column=day_col, value=hours)
                                break
                st.session_state.graph_bytes = wb_to_bytes(wb_graph)

            st.session_state.pf_bytes = wb_to_bytes(wb_pf)
            save_to_db()
            st.success(f"Отмечено {n} смен на {shift_date.strftime('%d.%m.%Y')} в Табеле и Графике (Сохранено в базу).")
            st.rerun()

# ================= TAB 2: ПОДРАБОТЧИКИ =================
with tab2:
    st.subheader("Внесение подработки")
    if st.session_state.sub_bytes is None or st.session_state.sh_bytes is None:
        st.info("👈 Создайте или загрузите месяц в боковой панели.")
    else:
        wb_sub = load_wb(st.session_state.sub_bytes)
        wb_sh = load_workbook(io.BytesIO(st.session_state.sh_bytes))
            
        c1, c2 = st.columns(2)
        with c1:
            fio = st.text_input("ФИО", key="sub_fio")
            post = st.selectbox("Должность", get_positions(), key="sub_post")
            grade_opts = get_grades(post)
            grade = st.selectbox("Разряд", grade_opts, key="sub_grade") if grade_opts else ""
        with c2:
            sub_date = st.date_input("Дата", datetime.date.today(), key="sub_date")
            amount = st.number_input("Сумма (руб)", min_value=0.0, value=4500.0, step=100.0, key="sub_amount")

        if st.button("💾 Сохранить подработку", type="primary"):
            if not fio.strip():
                st.error("Укажите ФИО")
            else:
                day = sub_date.day
                fio_clean = fio.strip()
                
                title = fio_clean[:31]
                if title not in wb_sub.sheetnames:
                    tpl = wb_sub[wb_sub.sheetnames[0] if wb_sub.sheetnames[0] != 'Итог' else wb_sub.sheetnames[1]]
                    new_ws = wb_sub.copy_worksheet(tpl); new_ws.title = title
                    for r in range(2, 33): new_ws.cell(row=r, column=2).value = None
                    target_ws = new_ws
                else:
                    target_ws = wb_sub[title]
                
                target_ws.cell(row=day + 1, column=2, value=1)
                if amount > 0: target_ws.cell(row=day + 1, column=3, value=float(amount))
                st.session_state.sub_bytes = wb_to_bytes(wb_sub)

                if "ПФ" in wb_sh.sheetnames:
                    ws_pf = wb_sh["ПФ"]
                    r = 4
                    while ws_pf.cell(row=r, column=3).value is not None and str(ws_pf.cell(row=r, column=3).value).strip().lower() != 'итого':
                        r += 1
                    ws_pf.cell(row=r, column=2, value=r - 3)
                    ws_pf.cell(row=r, column=3, value=fio_clean)
                    ws_pf.cell(row=r, column=4, value=post)
                    ws_pf.cell(row=r, column=5, value=grade)
                    ws_pf.cell(row=r, column=6, value=11)
                    ws_pf.cell(row=r, column=7, value="д")
                    ws_pf.cell(row=r, column=8, value=sub_date.strftime("%d.%m.%Y"))
                    if amount > 0: ws_pf.cell(row=r, column=9, value=float(amount))
                    
                    buf_sh = io.BytesIO(); wb_sh.save(buf_sh)
                    st.session_state.sh_bytes = buf_sh.getvalue()

                save_to_db()
                st.success(f"Подработка {fio_clean} внесена (Сохранено в базу).")
                st.rerun()

# ================= TAB 3: СТАВКИ И РАЗРЯДЫ =================
with tab3:
    st.subheader("Справочник ставок и разрядов")
    df = pd.DataFrame(st.session_state.rates)
    df.columns = ["Должность", "Разряд", "Ставка"]
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="rates_editor", column_config={"Должность": st.column_config.TextColumn(required=True), "Разряд": st.column_config.TextColumn(), "Ставка": st.column_config.NumberColumn(min_value=0, step=100, format="%d ₽")})

    if st.button("💾 Применить к табелю", type="primary"):
        if st.session_state.pf_bytes is None:
            st.error("Сначала создайте Табель.")
        else:
            new_rates = []
            for row in edited.to_dict("records"):
                if not row.get("Должность"): continue
                new_rates.append({"post": str(row["Должность"]).strip(), "grade": (str(row["Разряд"]).strip() if row.get("Разряд") else ""), "rate": int(row["Ставка"] or 0)})
            st.session_state.rates = new_rates
            updated = 0
            wb_pf = load_wb(st.session_state.pf_bytes)
            for sname, meta in st.session_state.staff_meta.items():
                if sname not in wb_pf.sheetnames: continue
                rt = get_rate(meta.get("post", ""), meta.get("grade", ""))
                if rt:
                    ws = wb_pf[sname]
                    for r in range(2, 33): ws.cell(row=r, column=3, value=rt)
                    updated += 1
            st.session_state.pf_bytes = wb_to_bytes(wb_pf)
            save_to_db()
            st.success(f"Ставки обновлены в Табеле (Сохранено в базу). Обновлено листов: {updated}.")
            st.rerun()

# ================= TAB 4: КАДРЫ =================
with tab4:
    st.subheader("Приём сотрудника")
    if st.session_state.pf_bytes is None:
        st.info("Сначала создайте Табель.")
    else:
        c1, c2, c3 = st.columns(3)
        with c1: hire_fio = st.text_input("ФИО нового сотрудника", key="hire_fio")
        with c2: hire_post = st.selectbox("Должность", get_positions(), key="hire_post")
        with c3:
            hire_grades = get_grades(hire_post)
            hire_grade = st.selectbox("Разряд", hire_grades, key="hire_grade") if hire_grades else ""

        hire_rate = get_rate(hire_post, hire_grade)
        if st.button("✅ Принять", type="primary", disabled=not hire_rate):
            fio_clean = hire_fio.strip()
            title = fio_clean[:31]
            wb_pf = load_wb(st.session_state.pf_bytes)
            if not fio_clean: st.error("Укажите ФИО")
            elif title in wb_pf.sheetnames: st.error("Сотрудник уже есть")
            else:
                tpl = None
                for s in wb_pf.sheetnames:
                    if s not in ("Итог", "рубкг") and not s.startswith("[Уволен]") and "вакансия" not in s.lower():
                        tpl = s; break
                if tpl is None: st.error("Нет шаблонного листа")
                else:
                    new_ws = wb_pf.copy_worksheet(wb_pf[tpl])
                    new_ws.title = title
                    for r in range(2, 33):
                        new_ws.cell(row=r, column=2, value=None)
                        new_ws.cell(row=r, column=3, value=hire_rate)
                    st.session_state.staff_meta[title] = {"post": hire_post, "grade": hire_grade}
                    st.session_state.pf_bytes = wb_to_bytes(wb_pf)
                    save_to_db()
                    st.success(f"Сотрудник {fio_clean} принят (Сохранено в базу).")
                    st.rerun()

        st.markdown("---")
        st.subheader("Увольнение")
        wb_pf = load_wb(st.session_state.pf_bytes)
        staff = read_staff(wb_pf)
        active_fio = [e["fio"] for e in staff]
        if active_fio:
            fire_fio = st.selectbox("Кого уволить", active_fio, key="fire_fio")
            if st.button("🚫 Уволить"):
                if fire_fio in wb_pf.sheetnames:
                    new_title = f"[Уволен] {fire_fio}"[:31]
                    wb_pf[fire_fio].title = new_title
                    if fire_fio in st.session_state.staff_meta:
                        st.session_state.staff_meta[new_title] = st.session_state.staff_meta.pop(fire_fio)
                    st.session_state.pf_bytes = wb_to_bytes(wb_pf)
                    save_to_db()
                    st.success(f"{fire_fio} уволен (Сохранено в базу).")
                    st.rerun()
        else:
            st.caption("Нет активных сотрудников")
