"""
Функции создания и обновления UI-плиток документов и предложений.

Все функции принимают параметры явно, не обращаются к self.*.
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
)


def create_document_tile(parent, doc, on_click, is_selected=False, cached_errors=None, is_active_check=False):
    """Создать плитку для документа. Иконка берётся из провайдера.

    Args:
        parent: Родительский виджет.
        doc: Словарь документа с ключами name, type.
        on_click: Callback(doc) при клике.
        is_selected: True если документ выбран.
        cached_errors: dict с результатами проверки из кэша или None.
        is_active_check: True если на этом документе сейчас идёт проверка.

    Returns:
        tuple: (tile, status_label_or_None).
    """
    tile = tk.Frame(parent, relief="raised", borderwidth=1, bg="#f0f0f0")
    tile.pack(fill=tk.X, pady=2, padx=2)

    # Иконка и цвет из провайдера
    provider = get_provider(doc.get("type", ""))
    icon_text, icon_color = provider.get_icon() if provider else ("?", "#888888")

    icon_label = tk.Label(
        tile, text=icon_text, font=("Arial", 14, "bold"),
        fg="white", bg=icon_color, width=2, height=1,
    )
    icon_label.pack(side=tk.LEFT, padx=(5, 10), pady=5)

    name_label = tk.Label(
        tile, text=doc["name"], wraplength=250, bg="#f0f0f0", anchor="w"
    )
    name_label.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=5)

    click_widgets = [tile, icon_label, name_label]
    status_label = None

    if is_selected and (is_active_check or cached_errors is not None):
        status_label = tk.Label(tile, text="", font=("Arial", 10), bg="#f0f0f0")
        status_label.pack(side=tk.RIGHT, padx=(0, 10), pady=5)
        click_widgets.append(status_label)
    elif cached_errors is not None:
        errors = sum(1 for r in cached_errors.values() if r.get("has_error"))
        text = f"✗ {errors}" if errors > 0 else "✓"
        fg = "#dc3545" if errors > 0 else "#28a745"
        status_label = tk.Label(tile, text=text, font=("Arial", 10), fg=fg, bg="#f0f0f0")
        status_label.pack(side=tk.RIGHT, padx=(0, 10), pady=5)
        click_widgets.append(status_label)

    for widget in click_widgets:
        widget.bind("<Button-1>", lambda e, d=doc: on_click(d))
        widget.configure(cursor="hand2")

    return tile, status_label


def highlight_selected_tile(tile_frames, selected_doc):
    """Подсветить выбранную плитку документа, остальные вернуть к обычному виду.

    Args:
        tile_frames: dict {id(doc): tile_frame}.
        selected_doc: Выбранный документ или None.
    """
    for doc_id, tile in tile_frames.items():
        if selected_doc and doc_id == id(selected_doc):
            tile.config(bg="#cce5ff")
            for child in tile.winfo_children():
                if isinstance(child, tk.Label) and child.cget("text") not in ("W", "O"):
                    child.config(bg="#cce5ff")
        else:
            tile.config(bg="#f0f0f0")
            for child in tile.winfo_children():
                if isinstance(child, tk.Label) and child.cget("text") not in ("W", "O"):
                    child.config(bg="#f0f0f0")


def highlight_selected_sentence_tile(tile_frames, selected_index):
    """Подсветить выбранную плитку предложения.

    Args:
        tile_frames: dict {index: tile_frame}.
        selected_index: Индекс выбранного предложения или None.
    """
    for idx, tile in tile_frames.items():
        if not tile.winfo_exists():
            continue
        if selected_index is not None and idx == selected_index:
            tile.config(bg="#e8f4fd", relief="solid", borderwidth=2)
            for child in tile.winfo_children():
                if isinstance(child, (tk.Frame, tk.Label)):
                    bg = child.cget("bg")
                    if bg and bg not in ("#f0f0f0", "", "systembuttonface"):
                        # Не менять фон кнопок и diff-виджетов
                        pass
                    elif isinstance(child, tk.Label):
                        child.config(bg="#e8f4fd")
        else:
            tile.config(bg="#f0f0f0", relief="raised", borderwidth=1)
            for child in tile.winfo_children():
                if isinstance(child, tk.Label):
                    current_bg = child.cget("bg")
                    # Восстановить оригинальные фоны
                    if current_bg in ("#e8f4fd",):
                        child.config(bg="#f0f0f0")


def create_sentence_tile_checking(parent, sentence, on_click):
    """Создать плитку со спиннером во время проверки.

    Args:
        parent: Родительский виджет.
        sentence: Словарь предложения с index, text.
        on_click: Callback(sentence) при клике.

    Returns:
        tuple: (tile, status_label).
    """
    tile = tk.Frame(parent, relief="raised", borderwidth=1, bg="#f0f0f0")
    tile.pack(fill=tk.X, pady=2, padx=2)

    index = sentence["index"]

    index_label = tk.Label(
        tile, text=str(index + 1), font=("Arial", 10),
        fg="white", bg="#666666", width=3, height=1,
    )
    index_label.pack(side=tk.LEFT, padx=(5, 5), pady=5)

    status_label = tk.Label(
        tile, text=WAITING_SYMBOL, bg="#f0f0f0", fg="#999999", font=("Arial", 12)
    )
    status_label.pack(side=tk.LEFT, padx=(0, 5), pady=5)

    text_label = tk.Label(
        tile, text=sentence["text"], wraplength=250,
        bg="#f0f0f0", anchor="w", justify="left",
    )
    text_label.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=5, padx=(0, 5))

    for widget in [tile, index_label, status_label, text_label]:
        widget.bind("<Button-1>", lambda e, s=sentence: on_click(s))
        widget.configure(cursor="hand2")

    return tile, status_label


def create_checked_sentence_tile(parent, sentence, result, on_click, on_apply=None, on_skip=None):
    """Создать плитку для проверенного предложения.

    Args:
        parent: Родительский виджет.
        sentence: Словарь предложения с index, text.
        result: Словарь результата с original, corrected, has_error, state.
        on_click: Callback(sentence) при клике на текст.
        on_apply: Callback(index) при нажатии "Применить"/"Отменить".
        on_skip: Callback(index) при нажатии "Пропустить"/"Отменить".

    Returns:
        dict: {"tile": tile, "buttons": {...} или None}.
    """
    index = sentence["index"]
    original = result["original"]
    corrected = result["corrected"]
    has_error = result["has_error"]
    state = result.get("state", "pending")

    tile = tk.Frame(parent, relief="raised", borderwidth=1, bg="#f0f0f0")
    tile.pack(fill=tk.X, pady=2, padx=2)

    header_frame = tk.Frame(tile, bg="#f0f0f0")
    header_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(5, 0))

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
    index_label.pack(side=tk.LEFT)

    result_label = tk.Label(
        header_frame, text="✓", bg="#f0f0f0",
        fg="#28a745" if not has_error else "#dc3545", font=("Arial", 12),
    )
    result_label.pack(side=tk.LEFT, padx=(5, 0))

    click_widgets = [tile, header_frame, index_label, result_label]
    buttons_info = None

    if has_error:
        text_widget = create_diff_widget(tile, original, corrected)
        click_widgets.append(text_widget)

        buttons_frame = tk.Frame(tile, bg="#f0f0f0")
        buttons_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(0, 5))

        if state == "applied":
            apply_text, apply_bg = "Отменить", "#28a745"
        else:
            apply_text, apply_bg = "Применить", "#dc3545"

        apply_btn = tk.Button(
            buttons_frame, text=apply_text, fg="white", bg=apply_bg, width=10,
            command=lambda idx=index: on_apply(idx) if on_apply else None,
        )
        apply_btn.pack(side=tk.LEFT, padx=(0, 5))
        if state == "skipped":
            apply_btn.config(state="disabled", bg="#6c757d")

        if state == "skipped":
            skip_text, skip_bg = "Отменить", "#28a745"
        else:
            skip_text, skip_bg = "Пропустить", "#6c757d"

        skip_btn = tk.Button(
            buttons_frame, text=skip_text, fg="white", bg=skip_bg, width=10,
            command=lambda idx=index: on_skip(idx) if on_skip else None,
        )
        skip_btn.pack(side=tk.LEFT)
        if state == "applied":
            skip_btn.config(state="disabled")

        buttons_info = {"apply": apply_btn, "skip": skip_btn, "index_label": index_label}
    else:
        text_label = tk.Label(
            tile, text=original, wraplength=300, bg="#f0f0f0", anchor="w", justify="left",
        )
        text_label.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(0, 5))
        click_widgets.append(text_label)

    for widget in click_widgets:
        widget.bind("<Button-1>", lambda e, s=sentence: on_click(s))
        widget.configure(cursor="hand2")

    return {"tile": tile, "buttons": buttons_info}


def update_sentence_tile(tile, index, original, corrected, has_error, sentence, on_click, on_apply=None, on_skip=None):
    """Обновить существующую плитку после проверки.

    Args:
        tile: Существующая плитка (tk.Frame).
        index: Индекс предложения.
        original: Исходный текст.
        corrected: Исправленный текст.
        has_error: True если есть изменения.
        sentence: Полный объект предложения с text, range_start, range_end.
        on_click: Callback(sentence) при клике.
        on_apply: Callback(index) для применить/отменить.
        on_skip: Callback(index) для пропустить.

    Returns:
        dict: {"buttons": {...} или None}.
    """
    for widget in tile.winfo_children():
        widget.destroy()

    header_frame = tk.Frame(tile, bg="#f0f0f0")
    header_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(5, 0))

    bg_color = "#666666" if has_error else "#28a745"
    index_label = tk.Label(
        header_frame, text=str(index + 1), font=("Arial", 10),
        fg="white", bg=bg_color, width=3, height=1,
    )
    index_label.pack(side=tk.LEFT)

    result_label = tk.Label(
        header_frame, text="✓", bg="#f0f0f0",
        fg="#28a745" if not has_error else "#dc3545", font=("Arial", 12),
    )
    result_label.pack(side=tk.LEFT, padx=(5, 0))

    click_widgets = [tile, header_frame, index_label, result_label]
    buttons_info = None

    if has_error:
        text_widget = create_diff_widget(tile, original, corrected)
        click_widgets.append(text_widget)

        buttons_frame = tk.Frame(tile, bg="#f0f0f0")
        buttons_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(0, 5))

        apply_btn = tk.Button(
            buttons_frame, text="Применить", fg="white", bg="#dc3545", width=10,
            command=lambda idx=index: on_apply(idx) if on_apply else None,
        )
        apply_btn.pack(side=tk.LEFT, padx=(0, 5))

        skip_btn = tk.Button(
            buttons_frame, text="Пропустить", fg="white", bg="#6c757d", width=10,
            command=lambda idx=index: on_skip(idx) if on_skip else None,
        )
        skip_btn.pack(side=tk.LEFT)

        buttons_info = {"apply": apply_btn, "skip": skip_btn, "index_label": index_label}
    else:
        text_label = tk.Label(
            tile, text=original, wraplength=300, bg="#f0f0f0", anchor="w", justify="left",
        )
        text_label.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(0, 5))
        click_widgets.append(text_label)

    for widget in click_widgets:
        widget.bind("<Button-1>", lambda e, s=sentence: on_click(s))
        widget.configure(cursor="hand2")

    return {"buttons": buttons_info}


def create_diff_widget(parent, original, corrected, after_callback=None):
    """Создать Text-виджет с цветным diff.

    Args:
        parent: Родительский виджет.
        original: Исходный текст.
        corrected: Исправленный текст.
        after_callback: Callback() для вызова после отрисовки.

    Returns:
        tk.Text: Виджет с diff.
    """
    text_widget = tk.Text(
        parent, wrap="word", height=1, borderwidth=0,
        highlightthickness=0, bg="#f0f0f0", cursor="hand2",
    )
    text_widget.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(0, 5))

    text_widget.tag_configure("added", background=DIFF_ADDED_BG, foreground=DIFF_ADDED_FG)
    text_widget.tag_configure("removed", background=DIFF_REMOVED_BG, foreground=DIFF_REMOVED_FG, overstrike=True)
    text_widget.tag_configure("removed_punct", background=DIFF_REMOVED_BG, foreground=DIFF_REMOVED_FG)
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
        text_widget.update_idletasks()
        result = text_widget.count("1.0", "end", "displaylines")
        if result:
            line_count = result[0]
            text_widget.config(height=max(1, line_count))

    text_widget.bind("<Configure>", adjust_height, add="+")
    if after_callback:
        parent.after(50, after_callback)
    else:
        parent.after(50, adjust_height)

    return text_widget


def insert_deleted_text(text_widget, text):
    """Вставить удалённый текст: буквы с overstrike, пунктуация без.

    Args:
        text_widget: Виджет Text.
        text: Текст для вставки.
    """
    for char in text:
        if char in PUNCTUATION:
            text_widget.insert("end", char, "removed_punct")
        else:
            text_widget.insert("end", char, "removed")
