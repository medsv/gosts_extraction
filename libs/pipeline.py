"""Общий конвейер обработки документов (PDF и DOCX): используется
и CLI (main.py), и Streamlit (app.py).

Содержит:
    process_file      — анализ одного файла по пути на диске (PDF или DOCX);
    process_bytes     — анализ одного файла из байтов (UploadedFile в Streamlit);
    process_pdf_file  — анализ одного PDF по пути (обёртка над process_file);
    process_pdf_bytes — анализ одного PDF из байтов (обёртка над process_bytes);
    process_docx_file — анализ одного DOCX по пути;
    process_docx_bytes— анализ одного DOCX из байтов;
    build_rows        — формирование строк для XLSX-перечня НТД.
"""

from pathlib import Path

from libs.gost_parser import (
    analyze_docx,
    analyze_docx_bytes,
    analyze_pdf,
    analyze_pdf_bytes,
)
from libs.models import SUPPORTED_EXTENSIONS, FileResult

# Для DOCX разбивка на листы отсутствует (страницы — результат вёрстки Word),
# поэтому в колонке «Номера листов» ставится прочерк.
DASH = "—"


def process_pdf_file(pdf_path: Path | str) -> FileResult:
    """Анализирует один PDF-файл; при ошибке возвращает FileResult с error."""
    return _process_by_path(pdf_path, "pdf")


def process_pdf_bytes(pdf_bytes: bytes, file_name: str) -> FileResult:
    """Анализирует PDF из байтов (например, загруженный через Streamlit)."""
    return _process_by_bytes(pdf_bytes, file_name)


def process_docx_file(docx_path: Path | str) -> FileResult:
    """Анализирует один DOCX-файл; при ошибке возвращает FileResult с error."""
    return _process_by_path(docx_path, "docx")


def process_docx_bytes(docx_bytes: bytes, file_name: str) -> FileResult:
    """Анализирует DOCX из байтов (например, загруженный через Streamlit)."""
    return _process_by_bytes(docx_bytes, file_name)


def _process_by_path(path: Path | str, kind: str) -> FileResult:
    """Анализирует файл по пути для заданного формата; ошибки не прерывают."""
    path = Path(path)
    try:
        gosts = analyze_pdf(str(path)) if kind == "pdf" else analyze_docx(str(path))
        return FileResult(
            file_path=str(path),
            file_name=path.name,
            gosts=gosts,
            kind=kind,
        )
    except Exception as e:  # битый/нечитаемый файл не должен останавливать остальные
        return FileResult(
            file_path=str(path),
            file_name=path.name,
            gosts={},
            kind=kind,
            error=str(e),
        )


def _process_by_bytes(data: bytes, file_name: str) -> FileResult:
    """Анализирует файл из байтов для заданного формата; ошибки не прерывают."""
    kind = Path(file_name).suffix.lower().lstrip(".")
    try:
        if kind == "pdf":
            gosts = analyze_pdf_bytes(data)
        elif kind == "docx":
            gosts = analyze_docx_bytes(data)
        else:
            return FileResult(
                file_path="",
                file_name=file_name,
                gosts={},
                kind=kind,
                error=f"Неподдерживаемый формат: {kind}",
            )
        return FileResult(
            file_path="",
            file_name=file_name,
            gosts=gosts,
            kind=kind,
        )
    except Exception as e:
        return FileResult(
            file_path="",
            file_name=file_name,
            gosts={},
            kind=kind,
            error=str(e),
        )


def process_file(path: Path | str) -> FileResult:
    """Анализирует один файл (PDF или DOCX) по пути, определяя формат
    по расширению. При ошибке возвращает FileResult с error, не прерывая
    обработку остальных файлов.
    """
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        kind = suffix.lstrip(".")
        return FileResult(
            file_path=str(path),
            file_name=path.name,
            gosts={},
            kind=kind,
            error=f"Неподдерживаемый формат: {kind}",
        )
    return _process_by_path(path, suffix.lstrip("."))


def process_bytes(data: bytes, file_name: str) -> FileResult:
    """Анализирует файл из байтов (например, загруженный через Streamlit),
    определяя формат по расширению имени. При ошибке возвращает FileResult
    с error, не прерывая обработку остальных.
    """
    return _process_by_bytes(data, file_name)


def build_rows(results: list[FileResult]) -> list[tuple]:
    """Формирует строки для XLSX: каждый ГОСТ каждого файла — отдельная строка.

    Нумерация № сквозная по всему файлу. Порядок: файлы в порядке входных
    аргументов, внутри файла — сортировка ГОСТов (алфавитная, как возвращает
    analyze_pdf/analyze_docx).
    Для DOCX в колонке «Номера листов» ставится прочерк, т.к. в DOCX нет
    физических страниц.
    """
    rows: list[tuple] = []
    for result in results:
        if not result.is_ok:
            continue
        is_docx = result.kind == "docx"
        for gost, entry in result.gosts.items():
            if is_docx:
                pages_str = DASH
            else:
                pages_str = ", ".join(map(str, entry["pages"]))
            rows.append((
                len(rows) + 1,        # №
                gost,                 # НТД
                entry["count"],       # Кол-во упоминаний
                pages_str,            # Номера листов
                result.file_name,     # Название файла
            ))
    return rows