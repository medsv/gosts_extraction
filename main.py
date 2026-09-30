"""CLI-приложение: анализ PDF/DOCX-файлов на наличие ГОСТ/ТУ/СП и формирование
файла «Перечень_НТД.xlsx» на основе шаблона.

Запуск из консоли:
    python main.py <файл.pdf|файл.docx|папка> [<файл2|папка2> ...] [--template ШАБЛОН]
                  [--out ПАПКА] [--recursive]

Каждый файл обрабатывается отдельно, по каждому файлу формируется свой
перечень ГОСТов. Результат сводится в один XLSX-файл с именем вида
YY-MM-DD_HH-MM_Перечень_НТД.xlsx.

Ядро вынесено в пакет libs (gost_parser, xlsx_writer, models) — те же функции
будут использоваться Streamlit-приложением (app.py).
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

from libs.models import SUPPORTED_EXTENSIONS, FileResult
from libs.pipeline import build_rows, process_file
from libs.xlsx_writer import analysis_to_xlsx_bytes

# Колонки шаблона «Перечень_НТД.xlsx»
COL_NO = 0      # A: №
COL_NTD = 1     # B: НТД
COL_COUNT = 2   # C: Кол-во упоминаний
COL_PAGES = 3   # D: Номера листов файла pdf
COL_FILENAME = 4  # E: Название файла

DEFAULT_TEMPLATE = Path(__file__).parent / "Перечень_НТД.xlsx"

# Расширения файлов, которые ищем в папках (для glob-паттернов)
FOLDER_PATTERNS = tuple(f"*{ext}" for ext in SUPPORTED_EXTENSIONS)


def collect_doc_files(paths: list[str], recursive: bool) -> list[Path]:
    """Собирает список PDF/DOCX-файлов из переданных путей.

    Файлы добавляются напрямую, папки — поиском *.pdf и *.docx (рекурсивно,
    если recursive=True). Отсутствующие пути и файлы других форматов
    пропускаются с предупреждением. Дубликаты устраняются (по resolve()).
    """
    doc_files: dict[Path, None] = {}
    for raw in paths:
        p = Path(raw)
        if not p.exists():
            print(f"ПРЕДУПРЕЖДЕНИЕ: путь не найден, пропуск: {raw}", file=sys.stderr)
            continue
        if p.is_file():
            if p.suffix.lower() in SUPPORTED_EXTENSIONS:
                doc_files.setdefault(p.resolve(), None)
            else:
                print(
                    f"ПРЕДУПРЕЖДЕНИЕ: не поддерживаемый формат, пропуск: {raw}",
                    file=sys.stderr,
                )
        elif p.is_dir():
            found: list[Path] = []
            for pattern in FOLDER_PATTERNS:
                glob_pattern = f"**/{pattern}" if recursive else pattern
                found.extend(p.glob(glob_pattern))
            found = sorted(set(found))
            if not found:
                print(
                    f"ПРЕДУПРЕЖДЕНИЕ: в папке нет PDF/DOCX-файлов: {p}",
                    file=sys.stderr,
                )
            for f in found:
                if f.is_file():
                    doc_files.setdefault(f.resolve(), None)
    return list(doc_files.keys())


def build_output_path(out_dir: str | Path | None, now: datetime | None = None) -> Path:
    """Формирует путь выходного файла YY-MM-DD_HH-MM_Перечень_НТД.xlsx."""
    now = now or datetime.now()
    stamp = now.strftime("%y-%m-%d_%H-%M")
    out_dir = Path(out_dir) if out_dir else Path.cwd()
    return out_dir / f"{stamp}_Перечень_НТД.xlsx"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="main.py",
        description=(
            "Анализ PDF/DOCX-файлов на наличие ГОСТ/ТУ/СП и формирование "
            "перечня НТД в XLSX на основе шаблона «Перечень_НТД.xlsx»."
        ),
        epilog=(
            "Примеры:\n"
            "  python main.py doc.pdf\n"
            "  python main.py report.docx dir1 dir2/file.pdf --recursive\n"
            "  python main.py dir --template template.xlsx --out results\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "paths",
        nargs="+",
        help="PDF/DOCX-файлы и/или папки с такими файлами",
    )
    parser.add_argument(
        "--template",
        default=str(DEFAULT_TEMPLATE),
        help=f"Путь к шаблону XLSX (по умолчанию: {DEFAULT_TEMPLATE.name})",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="Папка для сохранения результата (по умолчанию: текущая)",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Рекурсивно искать PDF/DOCX-файлы в переданных папках",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    print("Анализ PDF/DOCX на наличие ГОСТ/ТУ/СП")

    doc_files = collect_doc_files(args.paths, args.recursive)
    if not doc_files:
        print("Не найдено ни одного PDF/DOCX-файла для обработки.", file=sys.stderr)
        return 1

    print(f"Найдено файлов: {len(doc_files)}\n")

    results: list[FileResult] = []
    for i, doc_path in enumerate(doc_files, start=1):
        print(f"[{i}/{len(doc_files)}] Обработка: {doc_path.name} ...")
        result = process_file(doc_path)
        if result.is_ok:
            total = sum(e["count"] for e in result.gosts.values())
            print(f"    Упоминаний ГОСТ: {total}, уникальных: {len(result.gosts)}")
        else:
            print(f"    ОШИБКА: {result.error}", file=sys.stderr)
        results.append(result)

    ok_count = sum(1 for r in results if r.is_ok)
    err_count = len(results) - ok_count
    print(f"\nУспешно обработано: {ok_count}, с ошибками: {err_count}")

    rows = build_rows(results)
    if not rows:
        print("Нет данных для записи в XLSX (в обработанных файлах ГОСТы не найдены).")
        return 1

    try:
        xlsx_bytes = analysis_to_xlsx_bytes(rows, args.template)
    except FileNotFoundError as e:
        print(f"ОШИБКА: {e}", file=sys.stderr)
        return 1

    out_path = build_output_path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(xlsx_bytes)
    print(f"\nГотово! Перечень сохранён: {out_path}")
    print(f"Всего строк в перечне: {len(rows)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
