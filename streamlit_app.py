import streamlit as st
import pandas as pd
import io
import datetime
from openpyxl import load_workbook

st.set_page_config(page_title="Учёт ПФ", page_icon="📊", layout="wide")

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
    "pf_bytes": None,
    "sub_bytes": None,
    "pf_name": "Полуфабрикаты.xlsm",
    "sub_name": "Подработки_Шашина.xlsx",
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ---------- HELPERS ----------
def get_rate(post, grade):
    for r in st.session_state.rates:
        if r["post"] == post and r["grade"] == grade:
            return r["rate"]
    return None


def get_positions():
    return sorted(set(r["post"] for r in st.session_state.rates))


def get_grades(post):
    return [r["grade"] for r in st.session_state.rates if r["post"] == post]


def infer_post_grade(sname, sheet_rate):
    low = sname.lower()
    if "рузиева" in low and "шазие" in low:
        return ("Весовщик", "")
    if "насруддинов" in low or ("рузиев" in low and "ислом" in low):
        return ("Грузчик", "")
    if sheet_rate == 3700 or "никулин" in low or "ермакова" in low:
        return ("Изготовитель мясных п/ф", "1 разряд")
    return ("Изготовитель мясных п/ф", "2 разряд")


def load_wb(data):
    return load_workbook(io.BytesIO(data), keep_vba=True)


def wb_to_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def read_staff(wb):
    """Список активных сотрудников с должностью, разрядом, ставкой."""
    result = []
    for sname in wb.sheetnames:
        if sname in ("Итог", "рубкг"):
            continue
        ws = wb[sname]
        sheet_rate = None
        for r in range(2, 33):
            v = ws.cell(row=r, column=3).value
            if isinstance(v, (int, float)) and v > 0:
                sheet_rate = int(v)
                break
        meta = st.session_state.staff_meta.get(sname)
        if not meta:
            post, grade = infer_post_grade(sname, sheet_rate)
            meta = {"post": post, "grade": grade}
            st.session_state.staff_meta[sname] = meta
        post, grade = meta["post"], meta.get("grade", "")
        rate = get_rate(post, grade) or sheet_rate or 4500
        if not sname.startswith("[Уволен]") and "вакансия" not in sname.lower():
            result.append(
                {"sheet": sname, "fio": sname, "post": post, "grade": grade, "rate": rate}
            )
    return result


# ---------- SIDEBAR: ФАЙЛЫ ----------
st.sidebar.header("📁 Файлы")
st.sidebar.caption("Загрузите файлы в начале работы. После изменений — скачайте обновлённые версии.")

up_pf = st.sidebar.file_uploader(
    "Табель Полуфабрикаты (.xlsm)", type=["xlsm"], key="up_pf"
)
if up_pf is not None and st.sidebar.button("↪ Загрузить табель в сессию"):
    st.session_state.pf_bytes = up_pf.read()
    st.session_state.pf_name = up_pf.name
    st.sidebar.success("Табель загружен")

up_sub = st.sidebar.file_uploader(
    "Ведомость подработок (.xlsx / .xlsm)", type=["xlsx", "xlsm"], key="up_sub"
)
if up_sub is not None and st.sidebar.button("↪ Загрузить подработки в сессию"):
    st.session_state.sub_bytes = up_sub.read()
    st.session_state.sub_name = up_sub.name
    st.sidebar.success("Ведомость загружена")

st.sidebar.markdown("---")
if st.session_state.pf_bytes:
    st.sidebar.download_button(
        "⬇️ Скачать табель",
        data=st.session_state.pf_bytes,
        file_name=st.session_state.pf_name,
        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
        use_container_width=True,
    )
if st.session_state.sub_bytes:
    st.sidebar.download_button(
        "⬇️ Скачать подработки",
        data=st.session_state.sub_bytes,
        file_name=st.session_state.sub_name,
        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
        use_container_width=True,
    )

st.sidebar.markdown("---")
st.sidebar.warning(
    "⚠️ Данные хранятся в сессии браузера. "
    "**Обязательно скачайте файлы перед закрытием вкладки** — иначе изменения пропадут."
)


# ---------- MAIN ----------
st.title("📊 Учёт выработки — Отделение Полуфабрикатов")

if st.session_state.pf_bytes is None:
    st.info("👈 Загрузите файл табеля в боковой панели, чтобы начать работу.")
    st.stop()

try:
    wb_pf = load_wb(st.session_state.pf_bytes)
except Exception as e:
    st.error(f"Не удалось открыть табель: {e}")
    st.stop()

staff = read_staff(wb_pf)

tab1, tab2, tab3, tab4 = st.tabs(
    ["📅 Табель смен", "👷 Подработчики", "💰 Ставки и разряды", "👥 Кадры"]
)


# ================= TAB 1: ТАБЕЛЬ СМЕН =================
with tab1:
    st.subheader("Отметка смен")
    c1, c2 = st.columns(2)
    with c1:
        shift_date = st.date_input("Дата смены", datetime.date.today(), key="shift_date")
    with c2:
        tonnage = st.number_input(
            "Тоннаж цеха (кг)", min_value=0.0, value=0.0, step=1.0, key="tonnage"
        )

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
        st.session_state.pf_bytes = wb_to_bytes(wb_pf)
        st.success(
            f"Отмечено {n} смен на {shift_date.strftime('%d.%m.%Y')}. "
            "Скачайте файл в боковой панели."
        )
        st.rerun()


# ================= TAB 2: ПОДРАБОТЧИКИ =================
with tab2:
    st.subheader("Внесение подработки")
    if st.session_state.sub_bytes is None:
        st.info("👈 Загрузите файл подработок в боковой панели.")
    else:
        try:
            wb_sub = (
                load_workbook(io.BytesIO(st.session_state.sub_bytes))
                if st.session_state.sub_name.endswith(".xlsx")
                else load_wb(st.session_state.sub_bytes)
            )
        except Exception as e:
            st.error(f"Не удалось открыть файл: {e}")
            wb_sub = None

        if wb_sub:
            c1, c2 = st.columns(2)
            with c1:
                fio = st.text_input("ФИО", key="sub_fio")
                post = st.selectbox("Должность", get_positions(), key="sub_post")
                grade_opts = get_grades(post)
                grade = (
                    st.selectbox("Разряд", grade_opts, key="sub_grade")
                    if grade_opts
                    else ""
                )
            with c2:
                sub_date = st.date_input(
                    "Дата", datetime.date.today(), key="sub_date"
                )
                amount = st.number_input(
                    "Сумма (руб)", min_value=0.0, value=4500.0, step=100.0, key="sub_amount"
                )

            st.markdown("**Текущие записи (лист «ПФ»)**")
            if "ПФ" in wb_sub.sheetnames:
                ws = wb_sub["ПФ"]
                rows = []
                for r in range(4, ws.max_row + 1):
                    name = ws.cell(row=r, column=3).value
                    if name and str(name).strip().lower() != "итого":
                        rows.append(
                            {
                                "№": ws.cell(row=r, column=2).value,
                                "ФИО": name,
                                "Должность": ws.cell(row=r, column=4).value,
                                "Разряд": ws.cell(row=r, column=5).value,
                                "Часы": ws.cell(row=r, column=6).value,
                                "Дата": ws.cell(row=r, column=8).value,
                                "Сумма": ws.cell(row=r, column=9).value,
                            }
                        )
                st.dataframe(rows, use_container_width=True) if rows else st.caption(
                    "Записей пока нет"
                )

                if st.button("💾 Сохранить подработку", type="primary"):
                    if not fio.strip():
                        st.error("Укажите ФИО")
                    else:
                        r = 4
                        while (
                            ws.cell(row=r, column=3).value is not None
                            and str(ws.cell(row=r, column=3).value).strip().lower()
                            != "итого"
                        ):
                            r += 1
                        ws.cell(row=r, column=2, value=r - 3)
                        ws.cell(row=r, column=3, value=fio.strip())
                        ws.cell(row=r, column=4, value=post)
                        ws.cell(row=r, column=5, value=grade)
                        ws.cell(row=r, column=6, value=11)
                        ws.cell(row=r, column=7, value="д")
                        ws.cell(row=r, column=8, value=sub_date.strftime("%d.%m.%Y"))
                        ws.cell(row=r, column=9, value=float(amount))
                        buf = io.BytesIO()
                        wb_sub.save(buf)
                        st.session_state.sub_bytes = buf.getvalue()
                        st.success(
                            f"Подработка {fio} на {sub_date.strftime('%d.%m.%Y')} сохранена. "
                            "Скачайте файл в боковой панели."
                        )
                        st.rerun()
            else:
                st.error("В файле нет листа «ПФ»")


# ================= TAB 3: СТАВКИ И РАЗРЯДЫ =================
with tab3:
    st.subheader("Справочник ставок и разрядов")
    st.caption(
        "Измените ставки и нажмите «Применить к табелю» — обновятся листы всех "
        "сотрудников с этой должностью и разрядом."
    )

    df = pd.DataFrame(st.session_state.rates)
    df.columns = ["Должность", "Разряд", "Ставка"]

    edited = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        key="rates_editor",
        column_config={
            "Должность": st.column_config.TextColumn(required=True, width="medium"),
            "Разряд": st.column_config.TextColumn(width="small"),
            "Ставка": st.column_config.NumberColumn(
                min_value=0, step=100, format="%d ₽"
            ),
        },
    )

    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button("💾 Применить к табелю", type="primary"):
            new_rates = []
            for row in edited.to_dict("records"):
                if not row.get("Должность"):
                    continue
                new_rates.append(
                    {
                        "post": str(row["Должность"]).strip(),
                        "grade": (str(row["Разряд"]).strip() if row.get("Разряд") else ""),
                        "rate": int(row["Ставка"] or 0),
                    }
                )
            st.session_state.rates = new_rates

            updated = 0
            for sname, meta in st.session_state.staff_meta.items():
                if sname not in wb_pf.sheetnames:
                    continue
                rt = get_rate(meta.get("post", ""), meta.get("grade", ""))
                if rt:
                    ws = wb_pf[sname]
                    for r in range(2, 33):
                        ws.cell(row=r, column=3, value=rt)
                    updated += 1

            st.session_state.pf_bytes = wb_to_bytes(wb_pf)
            st.success(
                f"Ставки обновлены. Обновлено листов: {updated}. "
                "Скачайте табель в боковой панели."
            )
            st.rerun()
    with c2:
        import json

        st.download_button(
            "⬇️ Скачать справочник (JSON)",
            data=json.dumps(st.session_state.rates, ensure_ascii=False, indent=2),
            file_name="rates.json",
            mime="application/json",
        )


# ================= TAB 4: КАДРЫ =================
with tab4:
    st.subheader("Приём сотрудника")
    c1, c2, c3 = st.columns(3)
    with c1:
        hire_fio = st.text_input("ФИО нового сотрудника", key="hire_fio")
    with c2:
        hire_post = st.selectbox("Должность", get_positions(), key="hire_post")
    with c3:
        hire_grades = get_grades(hire_post)
        hire_grade = (
            st.selectbox("Разряд", hire_grades, key="hire_grade")
            if hire_grades
            else ""
        )

    hire_rate = get_rate(hire_post, hire_grade)
    st.info(f"Ставка: {hire_rate} ₽" if hire_rate else "⚠ Для позиции не задана ставка")

    if st.button("✅ Принять", type="primary", disabled=not hire_rate):
        fio_clean = hire_fio.strip()
        title = fio_clean[:31]
        if not fio_clean:
            st.error("Укажите ФИО")
        elif title in wb_pf.sheetnames:
            st.error("Сотрудник с таким ФИО уже есть")
        else:
            tpl = None
            for s in wb_pf.sheetnames:
                if (
                    s not in ("Итог", "рубкг")
                    and not s.startswith("[Уволен]")
                    and "вакансия" not in s.lower()
                ):
                    tpl = s
                    break
            if tpl is None:
                st.error("Не найден шаблонный лист для копирования")
            else:
                new_ws = wb_pf.copy_worksheet(wb_pf[tpl])
                new_ws.title = title
                for r in range(2, 33):
                    new_ws.cell(row=r, column=2, value=None)
                    new_ws.cell(row=r, column=3, value=hire_rate)
                st.session_state.staff_meta[title] = {
                    "post": hire_post,
                    "grade": hire_grade,
                }
                st.session_state.pf_bytes = wb_to_bytes(wb_pf)
                st.success(
                    f"Сотрудник {fio_clean} принят со ставкой {hire_rate} ₽. "
                    "Скачайте табель в боковой панели."
                )
                st.rerun()

    st.markdown("---")
    st.subheader("Увольнение")
    active_fio = [e["fio"] for e in staff]
    if active_fio:
        fire_fio = st.selectbox("Кого уволить", active_fio, key="fire_fio")
        if st.button("🚫 Уволить"):
            if fire_fio in wb_pf.sheetnames:
                new_title = f"[Уволен] {fire_fio}"[:31]
                wb_pf[fire_fio].title = new_title
                if fire_fio in st.session_state.staff_meta:
                    st.session_state.staff_meta[new_title] = st.session_state.staff_meta.pop(
                        fire_fio
                    )
                st.session_state.pf_bytes = wb_to_bytes(wb_pf)
                st.success(f"{fire_fio} переведён в архив.")
                st.rerun()
    else:
        st.caption("Нет активных сотрудников")