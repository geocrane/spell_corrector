"""
Загрузка и хранение настроек визуального стиля (цвета, отступы, шрифты).
"""

import json
import logging
import os

logger = logging.getLogger("ui.style_config")

_THIS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_STYLE_CONFIG_PATH = os.path.join(_THIS_DIR, "ui_style.json")

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
        "ribbon_font_size_lg": 9,
        "ribbon_font_size_sm": 8
    },
    "ribbon": {
        "compact": {
            "outer_padx": 4,
            "outer_pady": 3,
            "gap": 2,
            "min_height": 50
        },
        "normal": {
            "outer_padx": 8,
            "outer_pady": 4,
            "gap": 3,
            "min_height": 100
        },
        "border_radius": 0
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
