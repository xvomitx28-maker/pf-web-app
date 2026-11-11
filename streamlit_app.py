import streamlit as st
import pandas as pd
import io
import datetime
import calendar
import zipfile
import os
from openpyxl import load_workbook

st.set_page_config(page_title="Учёт ПФ", page_icon="📊", layout="wide")

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
    "graph_name": "ГРАФИК.xlsx", "sh_name": "Шашина.xlsx"
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

# ---------- SIDEBAR: ФАЙЛЫ ----------
st.sidebar.header("📁 Текущие файлы")
st.sidebar.caption("Если вы сформировали месяц на первой вкладке, файлы загрузятся автоматически. Либо загрузите их вручную.")

up_pf = st.sidebar.file_uploader("Табель (.xlsm)", type=["xlsm"])
if up_pf and st.sidebar.button("↪ Загрузить Табель"):
    st.session_state.pf_bytes = up_pf.read(); st.session_state.pf_name = up_pf.name
    st.sidebar.success("Табель готов")

up_graph = st.sidebar.file_uploader("График (.xlsx)", type=["xlsx"])
if up_graph and st.sidebar.button("↪ Загрузить График"):
    st.session_state.graph_bytes = up_graph.read(); st.session_state.graph_name = up_graph.name
    st.sidebar.success("График готов")

up_sub = st.sidebar.file_uploader("Подработчики (.xlsm)", type=["xlsm"])
if up_sub and st.sidebar.button("↪ Загрузить Подработчиков"):
    st.session_state.sub_bytes = up_sub.read(); st.session_state.sub_name = up_sub.name
    st.sidebar.success("Подработчики готовы")

up_sh = st.sidebar.file_uploader("Ведомость Шашиной (.xlsx)", type=["xlsx"])
if up_sh and st.sidebar.button("↪ Загрузить Ведомость"):
    st.session_state.sh_bytes = up_sh.read(); st.session_state.sh_name = up_sh.name
    st.sidebar.success("Ведомость готова")

st.sidebar.markdown("---")
st.sidebar.subheader("⬇️ Скачать результат")
if st.session_state.pf_bytes: st.sidebar.download_button("Скачать Табель", data=st.session_state.pf_bytes, file_name=st.session_state.pf_name, use_container_width=True)
if st.session_state.graph_bytes: st.sidebar.download_button("Скачать График", data=st.session_state.graph_bytes, file_name=st.session_state.graph_name, use_container_width=True)
if st.session_state.sub_bytes: st.sidebar.download_button("Скачать Подработчиков", data=st.session_state.sub_bytes, file_name=st.session_state.sub_name, use_container_width=True)
if st.session_state.sh_bytes: st.sidebar.download_button("Скачать Ведомость", data=st.session_state.sh_bytes, file_name=st.session_state.sh_name, use_container_width=True)

st.sidebar.warning("⚠️ Файлы живут в сессии. Скачивайте их перед закрытием вкладки!")

# ---------- MAIN ----------
st.title("📊 Учёт выработки — Отделение Полуфабрикатов")

tab0, tab1, tab2, tab3, tab4 = st.tabs(
    ["✨ Создать месяц", "📅 Табель смен", "👷 Подработчики", "💰 Ставки и разряды", "👥 Кадры"]
)

# ================= TAB 0: СОЗДАНИЕ МЕСЯЦА =================
with tab0:
    st.subheader("Формирование нового месяца")
    st.info("Программа возьмет шаблоны, зашитые в нее, очистит их и автоматически загрузит в приложение для работы.")
    
    c1, c2 = st.columns(2)
    with c1:
        new_month = st.selectbox("Месяц", MONTHS_RU, index=datetime.date.today().month - 1)
    with c2:
        new_year = st.number_input("Год", min_value=2024, max_value=2035, value=datetime.date.today().year)

    if st.button("✨ Сформировать файлы", type="primary"):
        if not all([os.path.exists(TPL_PF), os.path.exists(TPL_SUB), os.path.exists(TPL_GRAPH), os.path.exists(TPL_SH)]):
            st.error(f"❌ Файлы шаблонов не найдены! Убедитесь, что вы загрузили на GitHub файлы: {TPL_PF}, {TPL_SUB}, {TPL_GRAPH}, {TPL_SH}.")
        else:
            with st.spinner("Обработка шаблонов..."):
                try:
                    m_idx = MONTHS_RU.index(new_month) + 1
                    num_days = calendar.monthrange(new_year, m_idx)[1]
                    zip_buf = io.BytesIO()
                    
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        # 1. Табель ПФ
                        wb_pf = load_wb(open(TPL_PF, "rb").read())
                        if "рубкг" in wb_pf.sheetnames:
                            ws_rub = wb_pf["рубкг"]
                            for day_i in range(1, 32):
                                r = 1 + day_i
                                ws_rub.cell(row=r, column=1).value = datetime.datetime(new_year, m_idx, day_i) if day_i <= num_days else None
                                ws_rub.cell(row=r, column=6).value = None
                                ws_rub.cell(row=r, column=8).value = None
                                ws_rub.cell(row=r, column=9).value = None
                        for sname in wb_pf.sheetnames:
                            if sname not in ["Итог", "рубкг"] and not sname.startswith("[Уволен]"):
                                ws_e = wb_pf[sname]
                                for r in range(2, 33):
                                    for c in range(2, 8): ws_e.cell(row=r, column=c).value = None
                        pf_bytes = wb_to_bytes(wb_pf)
                        zf.writestr(f"Полуфабрикаты_{new_month}_{new_year}.xlsm", pf_bytes)
                        st.session_state.pf_bytes = pf_bytes
                        st.session_state.pf_name = f"Полуфабрикаты_{new_month}_{new_year}.xlsm"

                        # 2. Подработчики ПФ
                        wb_sub = load_wb(open(TPL_SUB, "rb").read())
                        for sname in wb_sub.sheetnames:
                            if sname != "Итог":
                                ws_s = wb_sub[sname]
                                for r in range(2, 33):
                                    ws_s.cell(row=r, column=2).value = None
                                    ws_s.cell(row=r, column=3).value = None
                        sub_bytes = wb_to_bytes(wb_sub)
                        zf.writestr(f"Подработчики_Полуфабрикаты_{new_month}_{new_year}.xlsm", sub_bytes)
                        st.session_state.sub_bytes = sub_bytes
                        st.session_state.sub_name = f"Подработчики_Полуфабрикаты_{new_month}_{new_year}.xlsm"

                        # 3. График ПФ
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
                                if cell_d.value is not None and not str(cell_d.value).startswith('='):
                                    cell_d.value = None
                        buf_graph = io.BytesIO(); wb_graph.save(buf_graph)
                        graph_bytes = buf_graph.getvalue()
                        zf.writestr(f"ГРАФИК Полуфабрикаты_{new_month}_{new_year}.xlsx", graph_bytes)
                        st.session_state.graph_bytes = graph_bytes
                        st.session_state.graph_name = f"ГРАФИК Полуфабрикаты_{new_month}_{new_year}.xlsx"

                        # 4. Шашина
                        wb_sh = load_workbook(io.BytesIO(open(TPL_SH, "rb").read()))
                        for sname in wb_sh.sheetnames:
                            ws_s = wb_sh[sname]
                            for r in range(4, min(150, ws_s.max_row + 1)):
                                for c in range(2, 11):
                                    cell = ws_s.cell(row=r, column=c)
                                    if cell.value is not None and not str(cell.value).startswith('='): cell.value = None
                        buf_sh = io.BytesIO(); wb_sh.save(buf_sh)
                        sh_bytes = buf_sh.getvalue()
                        zf.writestr(f"ПОДРАБОТКИ ДЛЯ 1С (Шашина И.В.)_{new_month}_{new_year}.xlsx", sh_bytes)
                        st.session_state.sh_bytes = sh_bytes
                        st.session_state.sh_name = f"ПОДРАБОТКИ ДЛЯ 1С (Шашина И.В.)_{new_month}_{new_year}.xlsx"

                    st.success(f"✅ Файлы на {new_month} {new_year} сформированы и **автоматически загружены в приложение!** Можете переходить на другие вкладки. Скачать ZIP-архив можно кнопкой ниже.")
                    st.download_button(label="⬇️ Скачать ZIP-архив со всеми файлами", data=zip_buf.getvalue(), file_name=f"Архив_{new_month}_{new_year}.zip", mime="application/zip")
                except Exception as e:
                    st.error(f"Ошибка при генерации: {e}")


# ================= TAB 1: ТАБЕЛЬ СМЕН =================
with tab1:
    if st.session_state.pf_bytes is None:
        st.info("👈 Сформируйте месяц на первой вкладке или загрузите рабочие файлы в боковой панели, чтобы начать работу.")
    else:
        try:
            wb_pf = load_wb(st.session_state.pf_bytes)
            wb_graph = load_workbook(io.BytesIO(st.session_state.graph_bytes)) if st.session_state.graph_bytes else None
        except Exception as e:
            st.error(f"Ошибка чтения файлов: {e}")
            st.stop()
            
        staff = read_staff(wb_pf)
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
                
            # --- СИНХРОНИЗАЦИЯ С ГРАФИКОМ ---
            if wb_graph:
                ws_gr = wb_graph.active
                day_col = None
                for c_idx in range(4, 36):
                    val = ws_gr.cell(row=3, column=c_idx).value
                    if str(val).strip() == str(day): day_col = c_idx; break
                
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
            st.success(f"Отмечено {n} смен на {shift_date.strftime('%d.%m.%Y')} в Табеле и Графике.")
            st.rerun()

# ================= TAB 2: ПОДРАБОТЧИКИ =================
with tab2:
    st.subheader("Внесение подработки")
    if st.session_state.sub_bytes is None or st.session_state.sh_bytes is None:
        st.info("👈 Сформируйте месяц на первой вкладке или загрузите Подработчиков и Ведомость Шашиной в боковой панели.")
    else:
        try:
            wb_sub = load_wb(st.session_state.sub_bytes)
            wb_sh = load_workbook(io.BytesIO(st.session_state.sh_bytes))
        except Exception as e:
            st.error(f"Ошибка чтения файлов: {e}")
            st.stop()
            
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

                st.success(f"Подработка {fio_clean} внесена в оба файла. Скачайте их в боковой панели.")
                st.rerun()

# ================= TAB 3: СТАВКИ И РАЗРЯДЫ =================
with tab3:
    st.subheader("Справочник ставок и разрядов")
    df = pd.DataFrame(st.session_state.rates)
    df.columns = ["Должность", "Разряд", "Ставка"]
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="rates_editor", column_config={"Должность": st.column_config.TextColumn(required=True), "Разряд": st.column_config.TextColumn(), "Ставка": st.column_config.NumberColumn(min_value=0, step=100, format="%d ₽")})

    if st.button("💾 Применить к табелю", type="primary"):
        if st.session_state.pf_bytes is None:
            st.error("Сначала сформируйте или загрузите Табель.")
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
            st.success(f"Ставки обновлены в Табеле. Обновлено листов: {updated}.")
            st.rerun()

# ================= TAB 4: КАДРЫ =================
with tab4:
    st.subheader("Приём сотрудника")
    if st.session_state.pf_bytes is None:
        st.info("👈 Сначала сформируйте или загрузите Табель.")
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
                    st.success(f"Сотрудник {fio_clean} принят.")
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
                    st.success(f"{fire_fio} уволен.")
                    st.rerun()
        else:
            st.caption("Нет активных сотрудников")
