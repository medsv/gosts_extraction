"""Общие типы данных для CLI (main.py) и Streamlit-приложения (app.py)."""

from dataclasses import dataclass, field


@dataclass
class FileResult:
    """Результат анализа одного PDF-файла.

    Атрибуты:
        file_path — путь к исходному PDF-файлу;
        file_name — имя файла (для колонки «Название файла»);
        gosts — словарь вида {"ГОСТ ...": {"count": N, "pages": [...]}},
                где count — число упоминаний во всём документе,
                pages — отсортированный список номеров листов без дубликатов;
        error — текст ошибки, если файл не удалось обработать (иначе None).
    """

    file_path: str
    file_name: str
    gosts: dict[str, dict] = field(default_factory=dict)
    error: str | None = None

    @property
    def is_ok(self) -> bool:
        """True, если файл успешно проанализирован."""
        return self.error is None