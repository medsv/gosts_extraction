"""Парсер ГОСТ/ТУ/СП из текста документов (PDF и DOCX).

Содержит регулярное выражение для поиска обозначений нормативно-технической
документации (ГОСТ, ТУ, СП), функции нормализации найденных обозначений
и постраничного анализа PDF-файла через PyMuPDF, а также анализа
DOCX-файлов через python-docx.
"""

import re
from io import BytesIO

import docx
import pymupdf

# Регулярное выражение для поиска ГОСТов (ГОСТ, ГОСТ Р, ГОСТ ЕН и т.п.),
# ТУ (технических условий) и СП (сводов правил).
#
# ГОСТ: с номером вида 21.110-2013, 2.104-2006, EN 1330-1:2017 или 12345-67.
# Поддерживаются одиночные и составные префиксы: Р, ЕН, EN, ISO, ИСО,
# а также их сочетания вида «Р ИСО», «Р ISO», «Р ЕН», «Р EN».
#
# ТУ (материалы и вещества): ТУ <код_группы>-<рег.номер>-<ОКПО>-<год>,
# например: ТУ 3449-018-40064547-01, ТУ 28.14.16-017-38576343-2013.
# После «ТУ» может стоять дефис; код группы может содержать точки (ОКПД2);
# год — 2 или 4 цифры (до 2000 г. — 2, с 2000 г. — 4).
#
# ТУ (машиностроение, ГОСТ 2.201): <код_орг>.<6 цифр>.<3 цифры> ТУ,
# например: АБВГ.303121.001 ТУ, ШРПИ.041221.002ТУ.
#
# СП (своды правил, ГОСТ Р 1.19-2023): СП <номер>.<рег.номер>.<год>,
# например: СП 36.13330.2012, СП 255.1325800.2016, СП 444.1326000.2019.
# После «СП» может не быть пробела (встречается в каталогах).
# Учитываются возможные переносы
# https://alice.yandex.ru/chat/01a0685f-b491-4197-a509-245c293ba1dd/

_GOST_RE = re.compile(
    # ── ГОСТ ──────────────────────────────────────────────────────
    r'\bГОСТ'
    r'\s*(?:Р)?'                             # «Р» (опционально)
    r'\s*(?:ЕН|EN|ISO|ИСО)?'                 # стандарт (опционально)
    r'\s*\d+(?:\s*[.-]\d+)*'                 # номер; \s* только перед разделителем
    r'(?:\s*[-:]\s*\d{2,4}(?!\d))?'          # год
    # ── ТУ для материалов и веществ ───────────────────────────────
    r'|'
    r'\bТУ'
    r'\s*-?\s*'
    r'\d+(?:\.\d+)*'
    r'\s*-\s*\d{3}'
    r'\s*-\s*\d{7,8}'
    r'\s*-\s*(?:\d{4}|\d{2})(?!\d)'
    # ── ТУ для машиностроения по ГОСТ 2.201 ──────────────────────
    r'|'
    r'\b[A-ZА-Я]{2,4}\.\d{6}\.\d{3}\s*ТУ'
    # ── СП (своды правил) ────────────────────────────────────────
    r'|'
    r'\bСП'
    r'\s*\d+(?:\s*[.-]\d+)*'                 # номер; \s* только перед разделителем
    r'(?:\s*[-:]\s*\d{2,4}(?!\d))?'          # год
)

def normalize_gost(raw: str) -> str:
    """Нормализует найденное совпадение регэкспа.

    Схлопывает пробелы, гарантирует пробел между буквенным префиксом
    и номером («ГОСТ14918-80» -> «ГОСТ 14918-80»), удаляет завершающие
    знаки препинания и приводит все буквы к верхнему регистру.
    """
    normalized = re.sub(r'\s+', ' ', raw.strip())
    # Гарантированный пробел между буквенным префиксом и номером:
    # «ГОСТ14918-80» -> «ГОСТ 14918-80»
    normalized = re.sub(r'(?<=[A-ZА-ЯЁ])(?=\d)', ' ', normalized)
    normalized = normalized.rstrip('.,;:')
    return normalized.upper()


def iter_gost_matches(text: str):
    """Возвращает нормализованные упоминания ГОСТ/ТУ/СП с дубликатами.

    В отличие от extract_gosts(), сохраняет повторные вхождения —
    это нужно для корректного подсчёта общего числа упоминаний.
    """
    if not isinstance(text, str):
        return
    for m in _GOST_RE.finditer(text):
        yield normalize_gost(m.group(0))


def extract_gosts(text: str) -> list[str]:
    """Извлекает из строки все упоминания ГОСТов в виде уникального списка.

    Каждый найденный ГОСТ нормализуется: схлопываются пробелы, удаляется
    завершающая точка, все буквы приводятся к верхнему регистру.
    Порядок элементов соответствует порядку первого вхождения в тексте.
    """
    found: dict[str, None] = {}
    for key in iter_gost_matches(text):
        if key and key not in found:
            found[key] = None
    return list(found.keys())


def _analyze_units(units: list[str]) -> dict[str, dict]:
    """Анализирует список текстовых единиц и собирает статистику по ГОСТ/ТУ/СП.

    Каждая единица (страница PDF или абзац DOCX) обрабатывается отдельно,
    чтобы знать «локацию» каждого упоминания. Пустые единицы пропускаются.

    Возвращает словарь вида:
        {
            "ГОСТ 21.110-2013": {"count": 3, "pages": [1, 4, 7]},
            ...
        }
    где count — общее число упоминаний во всём документе, а pages —
    отсортированный список номеров единиц (страниц/абзацев, с 1),
    на которых ГОСТ упомянут хотя бы раз.
    """
    result: dict[str, dict] = {}
    for idx, text in enumerate(units, start=1):
        if not text or not text.strip():
            continue
        seen_on_unit: set[str] = set()
        for key in iter_gost_matches(text):
            entry = result.setdefault(key, {"count": 0, "pages": []})
            entry["count"] += 1
            if key not in seen_on_unit:
                seen_on_unit.add(key)
                entry["pages"].append(idx)
    # Сортировка номеров единиц и итоговая сортировка по названию ГОСТ
    for entry in result.values():
        entry["pages"].sort()
    return dict(sorted(result.items()))


def _iter_docx_text(doc) -> list[str]:
    """Собирает текст DOCX-документа в виде списка единиц (абзацев).

    Учитываются абзацы основного документа, всех таблиц (включая вложенные)
    и колонтитулы секций — ГОСТы часто встречаются именно в таблицах
    (например, ведомости и перечни).
    """
    units: list[str] = [p.text for p in doc.paragraphs]

    def walk_table(table) -> None:
        for row in table.rows:
            for cell in row.cells:
                units.extend(p.text for p in cell.paragraphs)
                for nested in cell.tables:
                    walk_table(nested)

    for table in doc.tables:
        walk_table(table)

    for section in doc.sections:
        for header_footer in (section.header, section.footer):
            if header_footer is not None:
                units.extend(p.text for p in header_footer.paragraphs)

    return units


def _analyze_document(doc) -> dict[str, dict]:
    """Постранично анализирует открытый PyMuPDF-документ.

    Страницы без текста (чертежи, отсканированные растровые листы)
    пропускаются без анализа для ускорения работы.
    """
    units: list[str] = []
    # Итерация по индексам через load_page() даёт явный тип Page
    # (итератор Document в стабах PyMuPDF не типизирован).
    for i in range(doc.page_count):
        page = doc.load_page(i)
        # Режим "text" возвращает строку; явное приведение к str
        # снимает неоднозначность типов из стабов PyMuPDF.
        text = str(page.get_text("text"))
        # Оптимизация: если на листе нет текста — это чертёж/растр,
        # анализировать нечего, пропускаем страницу.
        if not text.strip():
            continue
        units.append(text)
    return _analyze_units(units)


def analyze_pdf(pdf_path: str) -> dict[str, dict]:
    """Постранично анализирует PDF-файл и собирает статистику по ГОСТ/ТУ/СП.

    Возвращает словарь вида:
        {
            "ГОСТ 21.110-2013": {"count": 3, "pages": [1, 4, 7]},
            ...
        }
    где count — общее число упоминаний во всём документе,
    а pages — отсортированный список номеров листов (порядковых номеров
    страниц в PDF, начиная с 1), на которых ГОСТ упомянут хотя бы раз.
    Лист, на котором ГОСТ встречается несколько раз, фиксируется один раз.
    """
    doc = pymupdf.open(pdf_path)
    try:
        return _analyze_document(doc)
    finally:
        doc.close()


def analyze_pdf_bytes(pdf_bytes: bytes) -> dict[str, dict]:
    """Анализирует PDF из байтов (например, загруженный через Streamlit).

    Возвращает ту же структуру, что и analyze_pdf().
    """
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        return _analyze_document(doc)
    finally:
        doc.close()


def analyze_docx(docx_path: str) -> dict[str, dict]:
    """Поабзацно анализирует DOCX-файл и собирает статистику по ГОСТ/ТУ/СП.

    Единицей анализа служит абзац документа (а также текст из таблиц
    и колонтитулов): в «pages» попадают порядковые номера абзацев.
    Возвращает ту же структуру, что и analyze_pdf().
    """
    doc = docx.Document(docx_path)
    return _analyze_units(_iter_docx_text(doc))


def analyze_docx_bytes(docx_bytes: bytes) -> dict[str, dict]:
    """Анализирует DOCX из байтов (например, загруженный через Streamlit).

    Возвращает ту же структуру, что и analyze_docx().
    """
    doc = docx.Document(BytesIO(docx_bytes))
    return _analyze_units(_iter_docx_text(doc))