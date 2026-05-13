"""
Виджеты-кнопки в стиле Office Ribbon (плоский / flat-дизайн).

RibbonButton  — обычная icon+text-кнопка (Найти, Назад, Проверить,
                Исправления). Иконка сверху, подпись снизу.
RibbonToggle  — icon+text-toggle с состояниями off/on/mixed/disabled.

В нормальном состоянии — НЕТ заливки, прозрачный фон (как у родителя),
тёмный текст, иконка PNG. На hover — лёгкая серая подсветка с тонкой
рамкой. На pressed — голубоватая подсветка. У toggle on — голубой фон
с синей рамкой. У стиля «danger» (Исправления) акцент в красных тонах.
"""

import tkinter as tk
from tkinter import ttk
from typing import Callable, Optional

from ui import icons
from ui.style_config import style
from ui.constants import (
    RIBBON_BG,
    RIBBON_NORMAL_FG,
    RIBBON_NORMAL_BORDER,
    RIBBON_HOVER_BG,
    RIBBON_HOVER_BORDER,
    RIBBON_PRESSED_BG,
    RIBBON_PRESSED_BORDER,
    RIBBON_DANGER_FG,
    RIBBON_DANGER_HOVER_BG,
    RIBBON_DANGER_HOVER_BORDER,
    RIBBON_DANGER_PRESSED_BG,
    RIBBON_DANGER_PRESSED_BORDER,
    RIBBON_TOGGLE_ON_BG,
    RIBBON_TOGGLE_ON_HOVER,
    RIBBON_TOGGLE_ON_BORDER,
    RIBBON_TOGGLE_ON_FG,
    RIBBON_TOGGLE_MIXED_BG,
    RIBBON_TOGGLE_MIXED_HOVER,
    RIBBON_TOGGLE_MIXED_BORDER,
    RIBBON_TOGGLE_MIXED_FG,
    RIBBON_DISABLED_FG,
    RIBBON_FONT_LG,
    RIBBON_FONT_SM,
    RIBBON_FONT_FAMILY,
)


def _resolve_parent_bg(parent) -> str:
    """Определить фон родителя — нужен для прозрачной отрисовки.

    Возвращает hex-цвет; если ничего не удалось, возвращает RIBBON_BG.
    """
    # tk.Frame / Canvas / etc — есть атрибут bg
    try:
        bg = parent.cget("background")
        if bg:
            return bg
    except tk.TclError:
        pass
    # ttk.Frame — берём из стиля
    try:
        style = ttk.Style(parent)
        bg = style.lookup("TFrame", "background")
        if bg:
            return bg
    except Exception:
        pass
    return RIBBON_BG


class _RibbonBase(tk.Frame):
    """Общая основа для RibbonButton и RibbonToggle."""

    def __init__(
        self,
        parent,
        icon_key: str,
        text: str,
        command: Optional[Callable[[], None]] = None,
        *,
        icon_size: int = icons.ICON_LG,
        orient: str = "vertical",
        compact: bool = False,
    ):
        self._parent_bg = _resolve_parent_bg(parent)
        super().__init__(
            parent,
            bd=0,
            highlightthickness=1,
            highlightbackground=self._parent_bg,
            highlightcolor=self._parent_bg,
            bg=self._parent_bg,
        )
        self._icon_key = icon_key
        self._icon_size = icon_size
        self._text = text
        self._command = command
        self._orient = orient
        self._compact = compact

        # Размеры/отступы — из конфига
        config_key = "compact" if compact else "normal"
        self._outer_padx = style.get("ribbon", config_key, "outer_padx")
        self._outer_pady = style.get("ribbon", config_key, "outer_pady")
        self._gap = style.get("ribbon", config_key, "gap")

        if compact:
            self._font = RIBBON_FONT_SM
            self._fallback_font = (RIBBON_FONT_FAMILY, max(icon_size - 2, 10))
        else:
            self._font = RIBBON_FONT_LG
            self._fallback_font = (RIBBON_FONT_FAMILY, max(icon_size - 4, 12), "bold")

        self._hovering = False
        self._pressed = False
        self._photo = None  # удерживаем ссылку, чтобы GC не съел

        self._build_ui()
        self._bind_events()
        self._apply_visual()  # переопределяется в наследниках

    # ─── UI ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Внутренняя обёртка для padding
        self._inner = tk.Frame(self, bd=0, bg=self._parent_bg)
        self._inner.pack(padx=self._outer_padx, pady=self._outer_pady)

        # Иконка
        photo = icons.get_icon(self._icon_key, self._icon_size)
        self._photo = photo

        if photo is not None:
            self._icon_label = tk.Label(
                self._inner,
                image=photo,
                bd=0,
                bg=self._parent_bg,
            )
        else:
            # Текстовый fallback
            self._icon_label = tk.Label(
                self._inner,
                text=icons.get_fallback_text(self._icon_key),
                font=self._fallback_font,
                bd=0,
                bg=self._parent_bg,
            )

        # Текст
        self._text_label = tk.Label(
            self._inner,
            text=self._text,
            font=self._font,
            bd=0,
            bg=self._parent_bg,
        )

        # Раскладка
        if self._orient == "vertical":
            self._icon_label.pack(side=tk.TOP)
            self._text_label.pack(side=tk.TOP, pady=(self._gap, 0))
        else:
            self._icon_label.pack(side=tk.LEFT)
            self._text_label.pack(side=tk.LEFT, padx=(self._gap * 2, 0))

    def _all_subwidgets(self):
        return (self, self._inner, self._icon_label, self._text_label)

    def _bind_events(self):
        for w in self._all_subwidgets():
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            w.bind("<Button-1>", self._on_press)
            w.bind("<ButtonRelease-1>", self._on_release)

    # ─── Перекраска ────────────────────────────────────────────────────

    def _paint(self, bg: str, fg: str, border: str):
        for w in self._all_subwidgets():
            try:
                w.config(bg=bg)
            except tk.TclError:
                pass
        try:
            self._icon_label.config(fg=fg)
        except tk.TclError:
            pass
        try:
            self._text_label.config(fg=fg)
        except tk.TclError:
            pass
        try:
            self.config(highlightbackground=border, highlightcolor=border)
        except tk.TclError:
            pass

    def _apply_visual(self):
        """Реализуется в наследниках — выбирает текущую палитру по состоянию."""
        raise NotImplementedError

    # ─── События мыши ──────────────────────────────────────────────────

    def _interactive(self) -> bool:
        """Реализуется в наследниках — реагирует ли виджет на мышь."""
        return True

    def _on_enter(self, _event=None):
        if not self._interactive():
            return
        self._hovering = True
        self._apply_visual()

    def _on_leave(self, _event=None):
        if not self._interactive():
            return
        self._hovering = False
        self._pressed = False
        self._apply_visual()

    def _on_press(self, _event=None):
        if not self._interactive():
            return
        self._pressed = True
        self._apply_visual()

    def _on_release(self, _event=None):
        if not self._interactive():
            return
        was_pressed = self._pressed
        self._pressed = False
        self._apply_visual()
        if was_pressed and self._command is not None:
            try:
                self._command()
            except Exception:
                import logging

                logging.getLogger("ui.ribbon").exception("Ribbon command failed")

    # ─── Публичный API ─────────────────────────────────────────────────

    def set_command(self, command: Optional[Callable[[], None]]):
        self._command = command

    def set_text(self, text: str):
        self._text = text
        try:
            self._text_label.config(text=text)
        except tk.TclError:
            pass

    def set_icon(self, icon_key: str, state: str = "normal"):
        self._icon_key = icon_key
        photo = icons.get_icon(icon_key, self._icon_size, state=state)
        self._photo = photo
        if photo is not None:
            self._icon_label.config(image=photo, text="")
        else:
            self._icon_label.config(image="", text=icons.get_fallback_text(icon_key))


class RibbonButton(_RibbonBase):
    """Icon+text-кнопка в стиле Ribbon (flat).

    style:
      - "default" — обычная (Найти, Назад, Проверить...).
      - "danger"  — акцент в красных тонах (Исправления).
    """

    VALID_STYLES = ("default", "danger")

    def __init__(
        self,
        parent,
        icon_key: str,
        text: str,
        command: Optional[Callable[[], None]] = None,
        *,
        icon_size: int = icons.ICON_LG,
        orient: str = "vertical",
        style: str = "default",
        compact: bool = False,
    ):
        # Поддерживаем старые названия стилей, чтобы не ломать main_window.py
        if style in ("primary", "neutral"):
            style = "default"
        if style not in self.VALID_STYLES:
            style = "default"
        self._style = style
        self._enabled = True
        super().__init__(
            parent,
            icon_key,
            text,
            command,
            icon_size=icon_size,
            orient=orient,
            compact=compact,
        )

    def _interactive(self) -> bool:
        return self._enabled and self._command is not None

    def _apply_visual(self):
        if not self._enabled:
            self._paint(self._parent_bg, RIBBON_DISABLED_FG, border=self._parent_bg)
            self.set_icon(self._icon_key, state="disabled")
            self.config(cursor="arrow")
            return

        self.set_icon(self._icon_key, state="normal")
        if self._style == "danger":
            normal_fg = RIBBON_DANGER_FG
            hover_bg, hover_border = RIBBON_DANGER_HOVER_BG, RIBBON_DANGER_HOVER_BORDER
            pressed_bg, pressed_border = (
                RIBBON_DANGER_PRESSED_BG,
                RIBBON_DANGER_PRESSED_BORDER,
            )
        else:
            normal_fg = RIBBON_NORMAL_FG
            hover_bg, hover_border = RIBBON_HOVER_BG, RIBBON_HOVER_BORDER
            pressed_bg, pressed_border = RIBBON_PRESSED_BG, RIBBON_PRESSED_BORDER

        if self._pressed:
            self._paint(pressed_bg, normal_fg, border=pressed_border)
        elif self._hovering:
            self._paint(hover_bg, normal_fg, border=hover_border)
        else:
            self._paint(self._parent_bg, normal_fg, border=RIBBON_NORMAL_BORDER)
        self.config(cursor="hand2" if self._command else "arrow")

    def set_style(self, style: str):
        if style in ("primary", "neutral"):
            style = "default"
        if style in self.VALID_STYLES:
            self._style = style
            self._apply_visual()

    def set_enabled(self, flag: bool):
        self._enabled = bool(flag)
        if not self._enabled:
            self._hovering = False
            self._pressed = False
        self._apply_visual()


class RibbonToggle(_RibbonBase):
    """Icon+text-toggle (flat) с состояниями off / on / mixed / disabled.

    Сам по клику состояние НЕ переключает — только вызывает command.
    Внешний код после реакции движка устанавливает state через set_state(...).
    """

    VALID_STATES = ("off", "on", "mixed", "disabled")

    def __init__(
        self,
        parent,
        icon_key: str,
        text: str,
        command: Optional[Callable[[], None]] = None,
        *,
        icon_size: int = icons.ICON_SM,
        orient: str = "vertical",
        initial_state: str = "off",
        compact: bool = True,
    ):
        if initial_state not in self.VALID_STATES:
            initial_state = "off"
        self._state = initial_state
        super().__init__(
            parent,
            icon_key,
            text,
            command,
            icon_size=icon_size,
            orient=orient,
            compact=compact,
        )

    def _interactive(self) -> bool:
        return self._state != "disabled" and self._command is not None

    def _apply_visual(self):
        if self._state == "disabled":
            self._paint(self._parent_bg, RIBBON_DISABLED_FG, border=self._parent_bg)
            self.set_icon(self._icon_key, state="disabled")
            self.config(cursor="arrow")
            return

        self.set_icon(self._icon_key, state="normal")
        if self._state == "on":
            base_bg, hover_bg = RIBBON_TOGGLE_ON_BG, RIBBON_TOGGLE_ON_HOVER
            border = RIBBON_TOGGLE_ON_BORDER
            fg = RIBBON_TOGGLE_ON_FG
        elif self._state == "mixed":
            base_bg, hover_bg = RIBBON_TOGGLE_MIXED_BG, RIBBON_TOGGLE_MIXED_HOVER
            border = RIBBON_TOGGLE_MIXED_BORDER
            fg = RIBBON_TOGGLE_MIXED_FG
        else:  # off
            base_bg = self._parent_bg
            hover_bg = RIBBON_HOVER_BG
            border = RIBBON_NORMAL_BORDER  # невидимая
            fg = RIBBON_NORMAL_FG

        if self._hovering or self._pressed:
            self._paint(
                hover_bg,
                fg,
                border=RIBBON_HOVER_BORDER if self._state == "off" else border,
            )
        else:
            self._paint(base_bg, fg, border=border)

        self.config(cursor="hand2" if self._command else "arrow")

    def set_state(self, state: str):
        if state not in self.VALID_STATES:
            return
        self._state = state
        if state == "disabled":
            self._hovering = False
            self._pressed = False
        self._apply_visual()

    def get_state(self) -> str:
        return self._state


# ─── Self-test ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Ribbon self-test (flat)")
    root.geometry("520x300")
    root.configure(bg=RIBBON_BG)

    row1 = tk.Frame(root, bg=RIBBON_BG)
    row1.pack(fill=tk.X, padx=10, pady=8)
    RibbonButton(row1, "find", "Поиск", command=lambda: print("find")).pack(
        side=tk.LEFT, padx=4
    )
    ttk.Separator(row1, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)
    RibbonButton(row1, "check_all", "Весь текст", command=lambda: print("all")).pack(
        side=tk.LEFT, padx=4
    )
    RibbonButton(
        row1, "check_selection", "Фрагмент", command=lambda: print("sel")
    ).pack(side=tk.LEFT, padx=4)
    RibbonButton(
        row1,
        "errors_mode",
        "Исправления",
        command=lambda: print("errors"),
        style="danger",
        icon_size=icons.ICON_MD,
    ).pack(side=tk.LEFT, padx=4)

    row2 = tk.Frame(root, bg=RIBBON_BG)
    row2.pack(fill=tk.X, padx=10, pady=8)
    RibbonToggle(
        row2, "format_font", "Шрифт", command=lambda: print("font"), initial_state="off"
    ).pack(side=tk.LEFT, padx=2)
    RibbonToggle(
        row2, "format_size", "Размер", command=lambda: print("size"), initial_state="on"
    ).pack(side=tk.LEFT, padx=2)
    RibbonToggle(
        row2, "format_all", "Всё", command=lambda: print("all"), initial_state="mixed"
    ).pack(side=tk.LEFT, padx=2)
    RibbonToggle(
        row2,
        "format_style",
        "Стиль",
        command=lambda: print("style"),
        initial_state="disabled",
    ).pack(side=tk.LEFT, padx=2)
    RibbonToggle(
        row2, "auditor", "Аудитор", command=lambda: print("aud"), initial_state="on"
    ).pack(side=tk.LEFT, padx=8)
    RibbonToggle(
        row2,
        "skip_tables",
        "Без таблиц",
        command=lambda: print("tbl"),
        initial_state="off",
    ).pack(side=tk.LEFT, padx=2)

    root.mainloop()
