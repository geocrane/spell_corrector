"""
Константы для UI: цвета diff, пунктуация, анимация спиннера, палитра Ribbon.
"""

import os

# Цвета для diff
DIFF_ADDED_BG = "#d4edda"
DIFF_ADDED_FG = "#155724"
DIFF_REMOVED_BG = "#f8d7da"
DIFF_REMOVED_FG = "#721c24"

# Пунктуация (без зачёркивания при удалении)
PUNCTUATION = set('.,;:!?-—–()[]{}«»"\'""' "…/ ")

# Анимация спиннера
SPINNER_FRAMES = ["◐", "◓", "◑", "◒"]
SPINNER_DELAY = 150  # мс
WAITING_SYMBOL = "○"

# ─── Office Ribbon palette ──────────────────────────────────────────────
# Плоский дизайн: нет цветных заливок по умолчанию. Hover — лёгкая
# серая подсветка с тонкой рамкой; pressed — чуть темнее; toggle on —
# светло-голубой фон с синей рамкой; danger — лёгкий красноватый акцент
# в hover; disabled — приглушённый текст.

# Базовый цвет фона панели/toolbar — берётся из системного по умолчанию,
# но удобно иметь явный для перекраски кнопок:
RIBBON_BG = "#ffffff"

# Normal — без заливки, текст тёмный
RIBBON_NORMAL_FG       = "#1a1a1a"
RIBBON_NORMAL_BORDER   = RIBBON_BG  # рамка цвета фона = невидимая

# Hover (для обычных кнопок и primary)
RIBBON_HOVER_BG        = "#f3f3f3"
RIBBON_HOVER_BORDER    = "#c7c7c7"

# Pressed (нажатие)
RIBBON_PRESSED_BG      = "#e1ecf7"
RIBBON_PRESSED_BORDER  = "#a0c4ea"

# Danger accent (для «Исправления» — на hover красноватый акцент)
RIBBON_DANGER_FG       = "#c0392b"
RIBBON_DANGER_HOVER_BG = "#fdecea"
RIBBON_DANGER_HOVER_BORDER = "#e8a8a0"
RIBBON_DANGER_PRESSED_BG   = "#fad4d0"
RIBBON_DANGER_PRESSED_BORDER = "#d77a70"

# Toggle on (нажат)
RIBBON_TOGGLE_ON_BG      = "#cfe4f9"
RIBBON_TOGGLE_ON_HOVER   = "#b9d5f3"
RIBBON_TOGGLE_ON_BORDER  = "#2e75b5"
RIBBON_TOGGLE_ON_FG      = "#1a1a1a"

# Toggle mixed (часть on)
RIBBON_TOGGLE_MIXED_BG    = "#e8f0fb"
RIBBON_TOGGLE_MIXED_HOVER = "#d6e4f5"
RIBBON_TOGGLE_MIXED_BORDER = "#7aa6d7"
RIBBON_TOGGLE_MIXED_FG    = "#1a1a1a"

# Disabled
RIBBON_DISABLED_FG = "#a8a8a8"

# Шрифты
RIBBON_FONT_LG = ("Segoe UI", 9)
RIBBON_FONT_SM = ("Segoe UI", 8)

# Пути
UI_DIR    = os.path.dirname(os.path.abspath(__file__))
ICONS_DIR = os.path.join(UI_DIR, "icons")
