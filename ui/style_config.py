"""
Загрузка и хранение настроек визуального стиля (цвета, отступы, шрифты).
"""

import json
import logging
import os

from core import paths

logger = logging.getLogger("ui.style_config")

# Стиль хранится в userdata/ (вне кода) — переживает обновления приложения.
_STYLE_CONFIG_PATH = paths.STYLE_PATH

_DEFAULT_STYLE = {
    "colors": {
        "ribbon_bg": "#ffffff",
        "ribbon_normal_fg": "#1a1a1a",
        "ribbon_normal_border": "#ffffff",
        "ribbon_hover_bg": "#f3f3f3",
        "ribbon_hover_border": "#c7c7c7",
        "ribbon_pressed_bg": "#e1ecf7",
        "ribbon_pressed_border": "#a0c4ea",
        "ribbon_danger_fg": "#c0392b",
        "ribbon_danger_hover_bg": "#fdecea",
        "ribbon_danger_hover_border": "#e8a8a0",
        "ribbon_danger_pressed_bg": "#fad4d0",
        "ribbon_danger_pressed_border": "#d77a70",
        "ribbon_toggle_on_bg": "#cfe4f9",
        "ribbon_toggle_on_hover": "#b9d5f3",
        "ribbon_toggle_on_border": "#2e75b5",
        "ribbon_toggle_on_fg": "#1a1a1a",
        "ribbon_toggle_mixed_bg": "#e8f0fb",
        "ribbon_toggle_mixed_hover": "#d6e4f5",
        "ribbon_toggle_mixed_border": "#7aa6d7",
        "ribbon_toggle_mixed_fg": "#1a1a1a",
        "ribbon_disabled_fg": "#a8a8a8",
        "diff_added_bg": "#d4edda",
        "diff_added_fg": "#155724",
        "diff_removed_bg": "#f8d7da",
        "diff_removed_fg": "#721c24"
    },
    "fonts": {
        "ribbon_font_family": "Segoe UI",
        "ribbon_font_size_lg": 10,
        "ribbon_font_size_md": 9,
        "ribbon_font_size_sm": 7
    },
    "ribbon": {
        "compact": {
            "outer_padx": 3,
            "outer_pady": 2,
            "gap": 1,
            "min_height": 40,
            "fixed_width": 56,
            "fixed_height": 44,
            "icon_box": 18
        },
        "medium": {
            "outer_padx": 5,
            "outer_pady": 4,
            "gap": 2,
            "min_height": 62,
            "fixed_width": 80,
            "fixed_height": 66,
            "icon_box": 26
        },
        "normal": {
            "outer_padx": 10,
            "outer_pady": 6,
            "gap": 10,
            "min_height": 110,
            "fixed_width": 120,
            "fixed_height": 108,
            "icon_box": 64
        },
        "border_radius": 4
    },
    "tile": {
        "border_radius": 6,
        "padding": 8
    }
}


class StyleConfig:
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(StyleConfig, cls).__new__(cls)
            cls._instance._load()
        return cls._instance

    def _load(self):
        self._style = _DEFAULT_STYLE.copy()
        if os.path.exists(_STYLE_CONFIG_PATH):
            try:
                with open(_STYLE_CONFIG_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self._update_recursive(self._style, data)
            except Exception as e:
                logger.error("Failed to load ui_style.json: %s", e)

    def _update_recursive(self, base, update):
        for k, v in update.items():
            if isinstance(v, dict) and k in base and isinstance(base[k], dict):
                self._update_recursive(base[k], v)
            else:
                base[k] = v

    def get(self, *keys, default=None):
        """Получить значение из конфига по цепочке ключей."""
        curr = self._style
        for k in keys:
            if isinstance(curr, dict) and k in curr:
                curr = curr[k]
            else:
                return default
        return curr


# Глобальный объект для удобного доступа
style = StyleConfig()
