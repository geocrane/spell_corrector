"""
Функции создания и обновления UI-плиток документов и предложений.

Плитки оформлены в ribbon-стиле: Canvas со скруглённым прямоугольником-фоном,
тонкая рамка, мягкая палитра. Состояния normal / selected обновляются через
RibbonTile.set_selected(...).
"""

import difflib
import tkinter as tk
from tkinter import ttk

from core.providers.registry import get_provider
from ui.constants import (
    DIFF_ADDED_BG, DIFF_ADDED_FG,
    DIFF_REMOVED_BG, DIFF_REMOVED_FG,
    PUNCTUATION,
    WAITING_SYMBOL,
    RIBBON_BG,
    RIBBON_HOVER_BG,
    RIBBON_HOVER_BORDER,
    RIBBON_TOGGLE_ON_BG,
    RIBBON_TOGGLE_ON_BORDER,
    RIBBON_NORMAL_FG,
    RIBBON_DANGER_FG,
    RIBBON_DANGER_HOVER_BG,
)
from ui.ribbon import draw_rounded_rect, _resolve_parent_bg
from ui.style_config import style


# ─── RibbonTile — базовая плитка в ribbon-стиле ──────────────────────────


class RibbonTile(tk.Canvas):
    """Плитка с rounded-rect фоном в ribbon-стиле.

    Содержимое складывать в .body (tk.Frame). Состояния:
      - normal: белый фон, тонкая серая рамка
      - selected: голубой фон toggle_on, синяя рамка

    Канвас сам подгоняет высоту под inner.winfo_reqheight() и ширину
    под parent через <Configure> родителя.
    """

    def __init__(self, parent, *, padding: int = None, radius: int = None):
        self._parent_bg = _resolve_parent_bg(parent)
        super().__init__(
            parent, bd=0, highlightthickness=0, bg=self._parent_bg,
        )

        self._padding = padding if padding is not None else style.get(
            "tile", "padding", default=8,
        )
        self._radius = radius if radius is not None else style.get(
            "tile", "border_radius", default=6,
        )

        self._is_selected = False
        self._body = tk.Frame(self, bg=RIBBON_BG, bd=0)
        self._body_window = self.create_window(
            self._padding, self._padding, window=self._body, anchor="nw",
        )

        self.bind("<Configure>", self._on_canvas_configure)
        self._body.bind("<Configure>", self._on_body_configure)

        self._redraw()

    @property
    def body(self) -> tk.Frame:
        return self._body

    def set_selected(self, selected: bool):
        if self._is_selected == bool(selected):
            return
        self._is_selected = bool(selected)
        self._sync_subwidget_bg()
        self._redraw()

    def is_selected(self) -> bool:
        return self._is_selected

    def _current_palette(self):
        if self._is_selected:
            return RIBBON_TOGGLE_ON_BG, RIBBON_TOGGLE_ON_BORDER
        return RIBBON_BG, RIBBON_HOVER_BORDER

    def _on_canvas_configure(self, _event=None):
        w = self.winfo_width()
        body_w = max(1, w - self._padding * 2)
        try:
            self.itemconfig(self._body_window, width=body_w)
        except tk.TclError:
            pass
        self._redraw()

    def _on_body_configure(self, _event=None):
        body_h = self._body.winfo_reqheight()
        h = body_h + self._padding * 2
        if int(self.cget("height")) != h:
            self.config(height=h)
        # winfo_height() ещё не обновился — рисуем с явной высотой
        self._redraw(target_h=h)

    def _redraw(self, target_h: int = None):
        bg, border = self._current_palette()
        self.delete("tile_bg")
        w = self.winfo_width()
        h = target_h if target_h is not None else self.winfo_height()
        if w < 2 or h < 2:
            return
        draw_rounded_rect(
            self, 1, 1, w - 1, h - 1,
            radius=self._radius,
            fill=bg, outline=border, width=1,
            tags="tile_bg",
        )
        self.tag_lower("tile_bg")

    def _sync_subwidget_bg(self):
        """Синхронизировать bg всех вложенных Frame/Label с цветом плитки."""
        bg, _ = self._current_palette()
        self._body.config(bg=bg)
        _recolor_descendants(self._body, bg)

    def bind_click(self, callback):
        """Привязать обработчик клика к Canvas и всем потомкам body."""
        self.bind("<Button-1>", lambda e: callback())
        self.config(cursor="hand2")
        _bind_click_recursive(self._body, callback)


def _recolor_descendants(widget, bg):
    """Рекурсивно красить только «нейтральные» Label/Frame.

    Не трогаем виджеты с защищённым фоном (бейджи индекса с яркими цветами,
    diff-Text-widget и т.п.) — у них setattr('_keep_bg', True).
    """
    for child in widget.winfo_children():
        if getattr(child, "_keep_bg", False):
            continue
        if isinstance(child, (tk.Frame, tk.Label)):
            try:
                child.config(bg=bg)
            except tk.TclError:
                pass
        if isinstance(child, tk.Frame):
            _recolor_descendants(child, bg)


def _bind_click_recursive(widget, callback):
    """Привязать <Button-1> ко всем потомкам — клик в любом месте плитки."""
    for child in widget.winfo_children():
        # Не перехватываем клики у нажимаемых кнопок и Text-виджетов с курсором
        if isinstance(child, tk.Button):
            continue
        try:
            child.bind("<Button-1>", lambda e: callback(), add="+")
            child.configure(cursor="hand2")
        except tk.TclError:
            pass
        if isinstance(child, (tk.Frame, tk.Label)):
            _bind_click_recursive(child, callback)


# ─── Плитки документов ───────────────────────────────────────────────────


def create_document_tile(parent, doc, on_click, is_selected=False,
                         cached_errors=None, is_active_check=False):
    """Создать плитку для документа в ribbon-стиле.

    Returns:
        tuple: (RibbonTile, status_label_or_None).
    """
    tile = RibbonTile(parent)
    tile.pack(fill=tk.X, pady=3, padx=2)

    body = tile.body

    provider = get_provider(doc.get("type", ""))
    icon_text, icon_color = provider.get_icon() if provider else ("?", "#888888")

    icon_label = tk.Label(
        body, text=icon_text, font=("Arial", 14, "bold"),
        fg="white", bg=icon_color, width=2, height=1,
    )
    icon_label._keep_bg = True  # бейдж провайдера остаётся ярким
    icon_label.pack(side=tk.LEFT, padx=(0, 10), pady=2)

    name_label = tk.Label(
        body, text=doc["name"], bg=RIBBON_BG,
        fg=RIBBON_NORMAL_FG, anchor="w", justify="left",
    )
    name_label.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=2)
    name_label.bind("<Configure>", lambda e, lbl=name_label: lbl.config(wraplength=max(1, e.width)))

    status_label = None
    if is_selected and (is_active_check or cached_errors is not None):
        status_label = tk.Label(body, text="", font=("Arial", 10), bg=RIBBON_BG)
        status_label.pack(side=tk.RIGHT, padx=(0, 0), pady=2)
    elif cached_errors is not None:
        errors = sum(1 for r in cached_errors.values() if r.get("has_error"))
        text = f"✗ {errors}" if errors > 0 else "✓"
        fg = "#dc3545" if errors > 0 else "#28a745"
        status_label = tk.Label(
            body, text=text, font=("Arial", 10), fg=fg, bg=RIBBON_BG,
        )
        status_label.pack(side=tk.RIGHT, padx=(0, 0), pady=2)

    tile.set_selected(bool(is_selected))
    tile.bind_click(lambda d=doc: on_click(d))

    return tile, status_label


def highlight_selected_tile(tile_frames, selected_doc):
    """Подсветить выбранную плитку документа.

    Args:
        tile_frames: dict {id(doc): RibbonTile}.
        selected_doc: Выбранный документ или None.
    """
    selected_id = id(selected_doc) if selected_doc is not None else None
    for doc_id, tile in tile_frames.items():
        if not isinstance(tile, RibbonTile) or not tile.winfo_exists():
            continue
        tile.set_selected(doc_id == selected_id)


def highlight_selected_sentence_tile(tile_frames, selected_index):
    """Подсветить выбранную плитку предложения."""
    for idx, tile in tile_frames.items():
        if not isinstance(tile, RibbonTile) or not tile.winfo_exists():
            continue
        tile.set_selected(selected_index is not None and idx == selected_index)


# ─── Плитки предложений ──────────────────────────────────────────────────


def create_sentence_tile_checking(parent, sentence, on_click):
    """Создать плитку со спиннером во время проверки.

    Returns:
        tuple: (RibbonTile, status_label).
    """
    tile = RibbonTile(parent)
    tile.pack(fill=tk.X, pady=3, padx=2)

    body = tile.body
    index = sentence["index"]

    index_label = tk.Label(
        body, text=str(index + 1), font=("Arial", 10),
        fg="white", bg="#666666", width=3, height=1,
    )
    index_label._keep_bg = True
    index_label.pack(side=tk.LEFT, padx=(0, 5), pady=2)

    status_label = tk.Label(
        body, text=WAITING_SYMBOL, bg=RIBBON_BG, fg="#999999",
        font=("Arial", 12),
    )
    status_label.pack(side=tk.LEFT, padx=(0, 5), pady=2)

    text_label = tk.Label(
        body, text=sentence["text"],
        bg=RIBBON_BG, fg=RIBBON_NORMAL_FG, anchor="w", justify="left",
    )
    text_label.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=2, padx=(0, 5))
    text_label.bind("<Configure>", lambda e, lbl=text_label: lbl.config(wraplength=max(1, e.width)))

    tile.bind_click(lambda s=sentence: on_click(s))

    return tile, status_label


def create_checked_sentence_tile(parent, sentence, result, on_click,
                                 on_apply=None, on_skip=None):
    """Создать плитку для проверенного предложения.

    Returns:
        dict: {"tile": RibbonTile, "buttons": {...} | None}.
    """
    index = sentence["index"]
    original = result["original"]
    corrected = result["corrected"]
    has_error = result["has_error"]
    state = result.get("state", "pending")

    tile = RibbonTile(parent)
    tile.pack(fill=tk.X, pady=3, padx=2)

    info = _populate_checked_sentence_body(
        tile, index, original, corrected, has_error, state, on_apply, on_skip,
    )

    tile.bind_click(lambda s=sentence: on_click(s))

    return {"tile": tile, "buttons": info}


def update_sentence_tile(tile, index, original, corrected, has_error, sentence,
                         on_click, on_apply=None, on_skip=None):
    """Обновить существующую плитку после проверки.

    Returns:
        dict: {"buttons": {...} | None}.
    """
    if isinstance(tile, RibbonTile):
        for widget in tile.body.winfo_children():
            widget.destroy()
        info = _populate_checked_sentence_body(
            tile, index, original, corrected, has_error,
            state="pending", on_apply=on_apply, on_skip=on_skip,
        )
        tile.bind_click(lambda s=sentence: on_click(s))
        return {"buttons": info}

    # Fallback на случай старого tk.Frame (на всякий случай)
    for widget in tile.winfo_children():
        widget.destroy()
    return {"buttons": None}


def _populate_checked_sentence_body(tile, index, original, corrected, has_error,
                                    state, on_apply, on_skip):
    """Заполнить body плитки проверенного предложения. Возвращает buttons-info."""
    body = tile.body

    header_frame = tk.Frame(body, bg=RIBBON_BG)
    header_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 0))

    if state in ("applied", "skipped"):
        bg_color = "#28a745"
    elif has_error:
        bg_color = "#666666"
    else:
        bg_color = "#28a745"

    index_label = tk.Label(
        header_frame, text=str(index + 1), font=("Arial", 10),
        fg="white", bg=bg_color, width=3, height=1,
    )
    index_label._keep_bg = True
    index_label.pack(side=tk.LEFT)

    result_label = tk.Label(
        header_frame, text="✓", bg=RIBBON_BG,
        fg="#28a745" if not has_error else "#dc3545", font=("Arial", 12),
    )
    result_label.pack(side=tk.LEFT, padx=(5, 0))

    buttons_info = None

    if has_error:
        text_widget = create_diff_widget(body, original, corrected)
        text_widget._keep_bg = True  # фон Text — пусть остаётся как был

        buttons_frame = tk.Frame(body, bg=RIBBON_BG)
        buttons_frame.pack(side=tk.TOP, fill=tk.X, pady=(4, 0))

        if state == "applied":
            apply_text, apply_bg = "Отменить", "#28a745"
        else:
            apply_text, apply_bg = "Применить", "#dc3545"

        apply_btn = tk.Button(
            buttons_frame, text=apply_text, fg="white", bg=apply_bg, width=10,
            relief="flat", bd=0,
            activebackground=apply_bg, activeforeground="white",
            command=lambda idx=index: on_apply(idx) if on_apply else None,
        )
        apply_btn._keep_bg = True
        apply_btn.pack(side=tk.LEFT, padx=(0, 5))
        if state == "skipped":
            apply_btn.config(state="disabled", bg="#6c757d")

        if state == "skipped":
            skip_text, skip_bg = "Отменить", "#28a745"
        else:
            skip_text, skip_bg = "Пропустить", "#6c757d"

        skip_btn = tk.Button(
            buttons_frame, text=skip_text, fg="white", bg=skip_bg, width=10,
            relief="flat", bd=0,
            activebackground=skip_bg, activeforeground="white",
            command=lambda idx=index: on_skip(idx) if on_skip else None,
        )
        skip_btn._keep_bg = True
        skip_btn.pack(side=tk.LEFT)
        if state == "applied":
            skip_btn.config(state="disabled")

        buttons_info = {
            "apply": apply_btn, "skip": skip_btn, "index_label": index_label,
        }
    else:
        text_label = tk.Label(
            body, text=original, bg=RIBBON_BG,
            fg=RIBBON_NORMAL_FG, anchor="w", justify="left",
        )
        text_label.pack(side=tk.TOP, fill=tk.X, pady=(2, 0))
        text_label.bind("<Configure>", lambda e, lbl=text_label: lbl.config(wraplength=max(1, e.width)))

    return buttons_info


# ─── Diff-виджет ─────────────────────────────────────────────────────────


def create_diff_widget(parent, original, corrected, after_callback=None):
    """Создать Text-виджет с цветным diff."""
    text_widget = tk.Text(
        parent, wrap="word", height=1, borderwidth=0,
        highlightthickness=0, bg=RIBBON_BG, cursor="hand2",
    )
    text_widget.pack(side=tk.TOP, fill=tk.X, pady=(2, 0))

    text_widget.tag_configure(
        "added", background=DIFF_ADDED_BG, foreground=DIFF_ADDED_FG,
    )
    text_widget.tag_configure(
        "removed", background=DIFF_REMOVED_BG, foreground=DIFF_REMOVED_FG,
        overstrike=True,
    )
    text_widget.tag_configure(
        "removed_punct", background=DIFF_REMOVED_BG, foreground=DIFF_REMOVED_FG,
    )
    text_widget.tag_configure("normal")

    matcher = difflib.SequenceMatcher(None, original, corrected)

    for opcode, i1, i2, j1, j2 in matcher.get_opcodes():
        if opcode == "equal":
            text_widget.insert("end", original[i1:i2], "normal")
        elif opcode == "replace":
            insert_deleted_text(text_widget, original[i1:i2])
            text_widget.insert("end", corrected[j1:j2], "added")
        elif opcode == "delete":
            insert_deleted_text(text_widget, original[i1:i2])
        elif opcode == "insert":
            text_widget.insert("end", corrected[j1:j2], "added")

    text_widget.config(state="disabled")

    def adjust_height(event=None):
        if not text_widget.winfo_exists():
            return
        # Убираем update_idletasks() отсюда — при массовом создании плиток
        # это вызывало "шторм" перерисовок. Одной синхронизации в конце
        # в main_window.py будет достаточно.
        result = text_widget.count("1.0", "end", "displaylines")
        if result:
            line_count = result[0]
            try:
                if int(text_widget.cget("height")) != line_count:
                    text_widget.config(height=max(1, line_count))
            except (tk.TclError, ValueError):
                pass

    text_widget.bind("<Configure>", adjust_height, add="+")
    # Вызываем один раз сразу, чтобы запустить расчет
    adjust_height()
    
    # Уменьшаем задержку с 50 до 10мс, чтобы "прогрев" под заглушкой шел быстрее
    if after_callback:
        parent.after(10, after_callback)
    else:
        parent.after(10, adjust_height)

    return text_widget


# ─── Плитка пословной правки (режим ошибок) ──────────────────────────────


def create_word_correction_tile(parent, correction, on_click):
    """Создать плитку для одной пословной правки."""
    tile = RibbonTile(parent)
    tile.pack(fill=tk.X, pady=3, padx=2)

    body = tile.body

    header_frame = tk.Frame(body, bg=RIBBON_BG)
    header_frame.pack(side=tk.TOP, fill=tk.X)

    sentence_index = correction.get("sentence_index", 0)
    index_label = tk.Label(
        header_frame, text=str(sentence_index + 1), font=("Arial", 10),
        fg="white", bg="#28a745", width=3, height=1,
    )
    index_label._keep_bg = True
    index_label.pack(side=tk.LEFT)

    original_chunk = correction.get("original_chunk", "")
    corrected_chunk = correction.get("corrected_chunk", "")

    text_widget = create_diff_widget(body, original_chunk, corrected_chunk)
    text_widget._keep_bg = True

    tile.bind_click(lambda c=correction: on_click(c))


# ─── Плитка сравнения документов ─────────────────────────────────────────────


def create_comparison_tile(parent, diff_item: dict, on_click):
    """Создать плитку для одного блока изменений в режиме сравнения.

    Типы (diff_item["tag"]):
      replace — diff A→B с красным/зелёным подсветом (через create_diff_widget)
      insert  — только зелёный: абзац добавлен в текущем документе
      delete  — только красный: абзац удалён (был в снимке, нет в текущем)
    """
    tile = RibbonTile(parent)
    tile.pack(fill=tk.X, pady=3, padx=2)

    body = tile.body
    tag = diff_item["tag"]

    header_frame = tk.Frame(body, bg=RIBBON_BG)
    header_frame.pack(side=tk.TOP, fill=tk.X)

    badge_cfg = {
        "replace": ("~", "#fd7e14", "Изменён абзац"),
        "insert":  ("+", "#28a745", "Добавлен абзац"),
        "delete":  ("−", "#dc3545", "Удалён абзац"),
    }
    badge_text, badge_color, header_text = badge_cfg.get(tag, ("?", "#6c757d", ""))

    badge = tk.Label(
        header_frame, text=badge_text, font=("Arial", 10, "bold"),
        fg="white", bg=badge_color, width=3, height=1,
    )
    badge._keep_bg = True
    badge.pack(side=tk.LEFT)

    hdr_lbl = tk.Label(
        header_frame, text=header_text,
        font=("Segoe UI", 8, "bold"),
        fg=badge_color, bg=RIBBON_BG,
    )
    hdr_lbl._keep_bg = True
    hdr_lbl.pack(side=tk.LEFT, padx=(4, 0))

    if tag == "replace":
        diff_w = create_diff_widget(body, diff_item["text_a"], diff_item["text_b"])
        diff_w._keep_bg = True
    elif tag == "insert":
        lbl = tk.Label(
            body, text=diff_item["text_b"], bg=DIFF_ADDED_BG, fg=DIFF_ADDED_FG,
            anchor="w", justify="left", wraplength=1,
        )
        lbl._keep_bg = True
        lbl.pack(side=tk.TOP, fill=tk.X, pady=(2, 0))
        lbl.bind("<Configure>", lambda e, l=lbl: l.config(wraplength=max(1, e.width)))
    else:  # delete — показываем полный текст удалённого абзаца
        lbl = tk.Label(
            body, text=diff_item["text_a"], bg=DIFF_REMOVED_BG, fg=DIFF_REMOVED_FG,
            anchor="w", justify="left", wraplength=1,
            font=("TkDefaultFont", 10, "overstrike"),
        )
        lbl._keep_bg = True
        lbl.pack(side=tk.TOP, fill=tk.X, pady=(2, 0))
        lbl.bind("<Configure>", lambda e, l=lbl: l.config(wraplength=max(1, e.width)))

        hint = tk.Label(
            body,
            text="↕ Клик — перейти к месту удаления в документе",
            font=("Segoe UI", 7), fg="#888", bg=RIBBON_BG,
            anchor="w",
        )
        hint._keep_bg = True
        hint.pack(side=tk.TOP, fill=tk.X, pady=(1, 0))

    tile.bind_click(lambda i=diff_item: on_click(i))

    return tile


# ─── Прочее ──────────────────────────────────────────────────────────────


def set_toggle_button_state(btn: tk.Button, is_active: bool) -> None:
    """Покрасить кнопку как toggle: зелёный = активна, серый = неактивна."""
    if is_active:
        btn.config(
            bg="#28a745", fg="white",
            activebackground="#1e7e34", activeforeground="white",
        )
    else:
        btn.config(
            bg="#f0f0f0", fg="#333333",
            activebackground="#d0d0d0", activeforeground="#333333",
        )


def set_toggle_button_mixed(btn: tk.Button) -> None:
    """Промежуточное состояние toggle-кнопки."""
    btn.config(
        bg="#3a82f7", fg="white",
        activebackground="#1f5fc9", activeforeground="white",
    )


def insert_deleted_text(text_widget, text):
    """Вставить удалённый текст: буквы с overstrike, пунктуация без."""
    for char in text:
        if char in PUNCTUATION:
            text_widget.insert("end", char, "removed_punct")
        else:
            text_widget.insert("end", char, "removed")
