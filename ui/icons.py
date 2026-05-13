"""
Реестр иконок и ленивая загрузка PNG для UI.

Чтобы заменить любую иконку — положите PNG по пути, указанному в реестре
ICONS (или измените значение в словаре). Размеры подгоняются автоматически
методом Image.thumbnail при первом запросе get_icon(key, size).

Если Pillow не установлена или файл отсутствует — get_icon вернёт None,
и виджет покажет текстовый fallback из ICON_FALLBACK_TEXT.
"""

import logging
import os
import tkinter as tk
from typing import Optional

logger = logging.getLogger("ui.icons")

# Стандартные размеры
ICON_LG = 64  # Большие кнопки: Найти, Назад, Проверить всё/выделенное, Исправления
ICON_MD = 20  # Средние: Скрыть, Применить всё, Ревизии
ICON_SM = 16  # Компактные toggle: Шрифт/Размер/Стиль, Аудитор, Таблицы

# Базовая директория с иконками
ICON_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")

# Реестр путей: ключ → имя файла внутри ICON_DIR.
# Чтобы заменить иконку — положите свой PNG по тому же имени.
ICONS: dict[str, str] = {
    "find": "find.png",
    "back": "back.png",
    "check_all": "check_all.png",
    "check_selection": "check_selection.png",
    "format_all": "format_all.png",
    "format_font": "format_font.png",
    "format_size": "format_size.png",
    "format_style": "format_style.png",
    "auditor": "auditor.png",
    "skip_tables": "skip_tables.png",
    "hide_clean": "hide_clean.png",
    "errors_mode": "errors_mode.png",
    "revisions": "revisions.png",
}

# Текстовые fallback'и (используются, если Pillow нет или PNG отсутствует)
ICON_FALLBACK_TEXT: dict[str, str] = {
    "find": "⟳",  # ⟳
    "back": "◂",  # ◂
    "check_all": "✓",  # ✓
    "check_selection": "✂",  # ✂
    "format_all": "✦",  # ✦
    "format_font": "Ff",
    "format_size": "↕",  # ↕
    "format_style": "B",
    "auditor": "A",
    "skip_tables": "▦",  # ▦
    "hide_clean": "\U0001f441",  # 👁
    "errors_mode": "⚠",  # ⚠
    "revisions": "R",
}

_CACHE: dict[tuple[str, int], "tk.PhotoImage"] = {}
_PIL_AVAILABLE: Optional[bool] = None  # None = не пробовали, True/False = знаем


def _try_import_pil():
    """Однократная попытка импортировать Pillow. Кэширует результат."""
    global _PIL_AVAILABLE
    if _PIL_AVAILABLE is None:
        try:
            from PIL import Image, ImageTk  # noqa: F401

            _PIL_AVAILABLE = True
        except ImportError:
            logger.warning("Pillow не установлена — иконки будут показаны как текст")
            _PIL_AVAILABLE = False
    return _PIL_AVAILABLE


def get_icon(key: str, size: int = ICON_LG) -> Optional["tk.PhotoImage"]:
    """Лениво загрузить иконку по ключу и размеру.

    Возвращает tk.PhotoImage, готовый к использованию в Label(image=...),
    либо None — если Pillow недоступна или PNG-файла нет. Кэширует
    результаты по (key, size). Модуль удерживает ссылки на PhotoImage —
    не пытайтесь сохранять их сами, ссылки нужны только если объект
    создан вне кэша.
    """
    cache_key = (key, size)
    if cache_key in _CACHE:
        return _CACHE[cache_key]

    if not _try_import_pil():
        return None

    rel = ICONS.get(key)
    if not rel:
        logger.warning("Неизвестный ключ иконки: %s", key)
        return None

    path = os.path.join(ICON_DIR, rel)
    if not os.path.isfile(path):
        return None

    try:
        from PIL import Image, ImageTk

        img = Image.open(path).convert("RGBA")
        img.thumbnail((size, size), Image.LANCZOS)
        photo = ImageTk.PhotoImage(img)
        _CACHE[cache_key] = photo
        return photo
    except Exception as exc:
        logger.exception("Не удалось загрузить иконку %s (%s): %s", key, path, exc)
        return None


def get_fallback_text(key: str) -> str:
    """Текстовый символ для иконки, если изображение недоступно."""
    return ICON_FALLBACK_TEXT.get(key, "?")
