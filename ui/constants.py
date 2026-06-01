"""
Константы для UI: цвета diff, пунктуация, анимация спиннера, палитра Ribbon.
Все визуальные параметры теперь загружаются из ui_style.json через style_config.
"""

import os
from ui.style_config import style

# Цвета для diff
DIFF_ADDED_BG = style.get("colors", "diff_added_bg")
DIFF_ADDED_FG = style.get("colors", "diff_added_fg")
DIFF_REMOVED_BG = style.get("colors", "diff_removed_bg")
DIFF_REMOVED_FG = style.get("colors", "diff_removed_fg")

# Пунктуация (без зачёркивания при удалении)
PUNCTUATION = set('.,;:!?-—–()[]{}«»"\'""' "…/ ")

# Анимация спиннера
SPINNER_FRAMES = ["◐", "◓", "◑", "◒"]
SPINNER_DELAY = 150  # мс
WAITING_SYMBOL = "○"

# ─── Office Ribbon palette ──────────────────────────────────────────────
# Базовый цвет фона панели/toolbar
RIBBON_BG = style.get("colors", "ribbon_bg")

# Normal — без заливки, текст тёмный
RIBBON_NORMAL_FG       = style.get("colors", "ribbon_normal_fg")
RIBBON_NORMAL_BORDER   = style.get("colors", "ribbon_normal_border")

# Hover
RIBBON_HOVER_BG        = style.get("colors", "ribbon_hover_bg")
RIBBON_HOVER_BORDER    = style.get("colors", "ribbon_hover_border")

# Pressed
RIBBON_PRESSED_BG      = style.get("colors", "ribbon_pressed_bg")
RIBBON_PRESSED_BORDER  = style.get("colors", "ribbon_pressed_border")

# Danger accent
RIBBON_DANGER_FG       = style.get("colors", "ribbon_danger_fg")
RIBBON_DANGER_HOVER_BG = style.get("colors", "ribbon_danger_hover_bg")
RIBBON_DANGER_HOVER_BORDER = style.get("colors", "ribbon_danger_hover_border")
RIBBON_DANGER_PRESSED_BG   = style.get("colors", "ribbon_danger_pressed_bg")
RIBBON_DANGER_PRESSED_BORDER = style.get("colors", "ribbon_danger_pressed_border")

# Toggle on
RIBBON_TOGGLE_ON_BG      = style.get("colors", "ribbon_toggle_on_bg")
RIBBON_TOGGLE_ON_HOVER   = style.get("colors", "ribbon_toggle_on_hover")
RIBBON_TOGGLE_ON_BORDER  = style.get("colors", "ribbon_toggle_on_border")
RIBBON_TOGGLE_ON_FG      = style.get("colors", "ribbon_toggle_on_fg")

# Toggle mixed
RIBBON_TOGGLE_MIXED_BG    = style.get("colors", "ribbon_toggle_mixed_bg")
RIBBON_TOGGLE_MIXED_HOVER = style.get("colors", "ribbon_toggle_mixed_hover")
RIBBON_TOGGLE_MIXED_BORDER = style.get("colors", "ribbon_toggle_mixed_border")
RIBBON_TOGGLE_MIXED_FG    = style.get("colors", "ribbon_toggle_mixed_fg")

# Disabled
RIBBON_DISABLED_FG = style.get("colors", "ribbon_disabled_fg")

# Шрифты
RIBBON_FONT_FAMILY = style.get("fonts", "ribbon_font_family")
RIBBON_FONT_LG = (RIBBON_FONT_FAMILY, style.get("fonts", "ribbon_font_size_lg"))
RIBBON_FONT_MD = (RIBBON_FONT_FAMILY, style.get("fonts", "ribbon_font_size_md"))
RIBBON_FONT_SM = (RIBBON_FONT_FAMILY, style.get("fonts", "ribbon_font_size_sm"))

# Пути
UI_DIR    = os.path.dirname(os.path.abspath(__file__))
ICONS_DIR = os.path.join(UI_DIR, "icons")
