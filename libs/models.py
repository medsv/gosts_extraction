"""Общие типы данных для CLI (main.py) и Streamlit-приложения (app.py)."""

from dataclasses import dataclass, field
from pathlib import Path

# Допустимые форматы исходных документов
SUPPORTED_EXTENSIONS = (".pdf", ".docx")


@dataclass
class FileResult:
    """Результат анализа одного файла (PDF или DOCX).

    Атрибуты:
        file_path — путь к исходному файлу;
        file_name — имя файла (для колонки «Название файла»);
        gosts — словарь вида {"ГОСТ ...": {"count": N, "pages": [...]}},
                где count — число упоминаний во всём документе,
                pages — отсортированный список номеров листов без дубликатов;
                для DOCX список пуст (страницы в файле отсутствуют);
        kind — формат файла: "pdf" или "docx" (по расширению file_name);
        error — текст ошибки, если файл не удалось обработать (иначе None).
    """

    file_path: str
    file_name: str
    gosts: dict[str, dict] = field(default_factory=dict)
    kind: str = ""
    error: str | None = None

    def __post_init__(self) -> None:
        """Вычисляет kind по расширению имени файла, если он не задан."""
        if not self.kind:
            self.kind = Path(self.file_name).suffix.lower().lstrip(".") \
                if self.file_name else ""

    @property
    def is_ok(self) -> bool:
        """True, если файл успешно проанализирован."""
        return self.error is None