"""Streamlit-приложение: анализ PDF/DOCX на наличие ГОСТ/ТУ/СП и выгрузка
перечня НТД в XLSX на основе шаблона «Перечень_НТД.xlsx».

Запуск:
    streamlit run app.py

Использует то же ядро (libs/pipeline, libs/gost_parser, libs/xlsx_writer),
что и консольная версия main.py.
"""

from datetime import datetime, timedelta, timezone
from pathlib import Path

import streamlit as st

from libs.pipeline import build_rows, process_bytes
from libs.xlsx_writer import analysis_to_xlsx_bytes

DEFAULT_TEMPLATE = Path(__file__).parent / "Перечень_НТД.xlsx"


def output_filename() -> str:
    """Имя выходного файла вида YY-MM-DD_HH-MM_Перечень_НТД.xlsx."""
    timestamp = datetime.now(timezone(timedelta(hours=3))).strftime('%Y-%m-%d_%H-%M')
    return f"{timestamp}_Перечень_НТД.xlsx"


# ── Настройки страницы ──────────────────────────────────────────────
st.set_page_config(
    page_title="PDF/DOCX → Перечень НТД",
    page_icon="📜",
    layout="centered",
)

st.title("📜 PDF/DOCX → Перечень НТД")
#st.markdown(
#    "Выберите один или несколько PDF/DOCX файлов, нажав Upload, или перетащите их сюда мышкой из Проводника"
#)

# ── Загрузка файлов ─────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Выберите один или несколько PDF/DOCX файлов, нажав Upload, или перетащите их сюда мышкой из Проводника",
    type=["pdf", "docx"],
    accept_multiple_files=True,
)

#if not uploaded_files:
#    st.warning("Загрузите один или несколько PDF-файлов для начала работы.")
#    st.stop()



# Показываем имена загруженных файлов
#for f in uploaded_files:
#    st.caption(f"• {f.name} — {f.size / 1024:.1f} КБ")
if  uploaded_files:
    st.info(f"Загружено файлов: **{len(uploaded_files)}**")
# ── Обработка ───────────────────────────────────────────────────────
    if st.button("🚀 Сформировать перечень НТД", type="primary"):
        # Обрабатываем каждый файл отдельно; ошибки не прерывают остальные
        results = []
        progress = st.progress(0.0)
        status = st.empty()
        for i, f in enumerate(uploaded_files):
            status.info(f"Обработка [{i + 1}/{len(uploaded_files)}]: {f.name} ...")
            results.append(process_bytes(f.getvalue(), f.name))
            progress.progress((i + 1) / len(uploaded_files))

        status.empty()

        # ── Сводка по файлам ────────────────────────────────────────────
        ok_count = sum(1 for r in results if r.is_ok)
        err_count = len(results) - ok_count

        st.subheader("Результаты обработки")
        st.write(f"Успешно обработано: **{ok_count}**, с ошибками: **{err_count}**")

        if err_count:
            st.error("Ошибки при обработке следующих файлов:")
            for r in results:
                if not r.is_ok:
                    st.error(f"❌ {r.file_name}: {r.error}")

        # Таблица-сводка по файлам: имя, упоминаний, уникальных ГОСТ
        summary = []
        for r in results:
            if r.is_ok:
                total = sum(e["count"] for e in r.gosts.values())
                summary.append({
                    "Файл": r.file_name,
                    "Упоминаний ГОСТ": total,
                    "Уникальных ГОСТ": len(r.gosts),
                })
        if summary:
            st.dataframe(summary, use_container_width=True)

        # ── Формирование XLSX ───────────────────────────────────────────
        rows = build_rows(results)
        if not rows:
            st.warning("В загруженных файлах не найдено упоминаний ГОСТ/ТУ/СП.")
            st.stop()

        try:
            xlsx_bytes = analysis_to_xlsx_bytes(rows, DEFAULT_TEMPLATE)
        except FileNotFoundError as e:
            st.error(f"Не найден шаблон перечня НТД: {e}")
            st.stop()

        st.success(f"✅ Готово! Строк в перечне: **{len(rows)}**")
        st.download_button(
            label="📥 Скачать XLSX файл",
            data=xlsx_bytes,
            file_name=output_filename(),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    # Предпросмотр первых строк перечня
    #st.subheader("Предпросмотр")
    #preview = [
    #    {
    #        "№": row[0],
    #        "НТД": row[1],
    #        "Кол-во упоминаний": row[2],
    #        "Номера листов": row[3],
    #        "Название файла": row[4],
    #    }
    #    for row in rows[:20]
    #]
    #st.dataframe(preview, use_container_width=True)

st.markdown(
    """
    <hr>
    <p style="text-align: left; color: gray;">
    <small>
    2026, С.В. Медведев, engpython@yandex.ru
    </small>
    </p>
    """,
    unsafe_allow_html=True,
)
