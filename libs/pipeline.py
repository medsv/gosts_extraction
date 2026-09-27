"""Общий конвейер обработки PDF: используется и CLI (main.py), и Streamlit (app.py).

Содержит:
    process_pdf_file  — анализ одного PDF по пути на диске;
    process_pdf_bytes — анализ одного PDF из байтов (UploadedFile в Streamlit);
    build_rows        — формирование строк для XLSX-перечня НТД.
"""

from pathlib import Path

from libs.gost_parser import analyze_pdf, analyze_pdf_bytes
from libs.models import FileResult


def process_pdf_file(pdf_path: Path | str) -> FileResult:
    """Анализирует один PDF-файл; при ошибке возвращает FileResult с error."""
    pdf_path = Path(pdf_path)
    try:
        gosts = analyze_pdf(str(pdf_path))
        return FileResult(
            file_path=str(pdf_path),
            file_name=pdf_path.name,
            gosts=gosts,
        )
    except Exception as e:  # битый/нечитаемый файл не должен останавливать остальные
        return FileResult(
            file_path=str(pdf_path),
            file_name=pdf_path.name,
            gosts={},
            error=str(e),
        )


def process_pdf_bytes(pdf_bytes: bytes, file_name: str) -> FileResult:
    """Анализирует PDF из байтов (например, загруженный через Streamlit).

    При ошибке возвращает FileResult с error, не прерывая обработку остальных.
    """
    try:
        gosts = analyze_pdf_bytes(pdf_bytes)
        return FileResult(
            file_path="",
            file_name=file_name,
            gosts=gosts,
        )
    except Exception as e:
        return FileResult(
            file_path="",
            file_name=file_name,
            gosts={},
            error=str(e),
        )


def build_rows(results: list[FileResult]) -> list[tuple]:
    """Формирует строки для XLSX: каждый ГОСТ каждого файла — отдельная строка.

    Нумерация № сквозная по всему файлу. Порядок: файлы в порядке входных
    аргументов, внутри файла — сортировка ГОСТов (алфавитная, как возвращает
    analyze_pdf/analyze_pdf_bytes).
    """
    rows: list[tuple] = []
    for result in results:
        if not result.is_ok:
            continue
        for gost, entry in result.gosts.items():
            pages_str = ", ".join(map(str, entry["pages"]))
            rows.append((
                len(rows) + 1,        # №
                gost,                 # НТД
                entry["count"],       # Кол-во упоминаний
                pages_str,            # Номера листов
                result.file_name,     # Название файла
            ))
    return rows