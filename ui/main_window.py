"""
MainWindow — главное окно приложения (UI слой).

Подписывается на события Engine, отображает документы и результаты проверки.
Все вызовы бизнес-логики делегирует в Engine.
"""

import ctypes
import ctypes.wintypes
import logging
import os
import threading
import tkinter as tk
from tkinter import ttk

import office_finder
from core.providers import get_all_providers
from ui.constants import SPINNER_FRAMES, SPINNER_DELAY
from ui.tiles import (
    create_document_tile,
    highlight_selected_tile,
    highlight_selected_sentence_tile,
    create_sentence_tile_checking,
    create_checked_sentence_tile,
    update_sentence_tile,
)

logger = logging.getLogger("ui.main_window")


def get_active_monitor_workarea():
    """Определить рабочую область монитора, на котором находится курсор.

    Returns:
        dict: {"x": int, "y": int, "width": int, "height": int}.
    """
    try:
        import win32api
        import win32con

        cursor_x, cursor_y = win32api.GetCursorPos()
        monitor = win32api.MonitorFromPoint(
            (cursor_x, cursor_y), win32con.MONITOR_DEFAULTTONEAREST
        )
        monitor_info = win32api.GetMonitorInfo(monitor)
        work_area = monitor_info["Work"]

        return {
            "x": work_area[0],
            "y": work_area[1],
            "width": work_area[2] - work_area[0],
            "height": work_area[3] - work_area[1],
        }
    except Exception:
        return {"x": 0, "y": 0, "width": 1920, "height": 1080}


class MainWindow(tk.Tk):
    """Главное окно приложения. Управляет UI, все вызовы бизнес-логики делегирует в Engine."""

    def __init__(self, engine):
        super().__init__()

        self.engine = engine
        self.title("Корректор орфографии")
        self.geometry("350x500")
        self.resizable(True, True)

        self._position_on_active_monitor()

        # Состояние UI
        self.current_view = "documents"
        self.selected_sentence_index = None  # Индекс выбранного предложения
        self.tile_frames = {}
        self.sentence_tiles = {}
        self.tile_buttons = {}
        self.spinner_index = 0
        self.current_checking_index = -1
        self.spinner_labels = {}
        self.spinner_job = None
        self.doc_status_label = None
        self._extraction_in_progress = False  # блокировка «Назад» во время извлечения

        # Конфигурация
        self._config = engine.get_config()
        self.available_adapters = engine.get_available_adapters()

        self._create_ui()
        self._subscribe_to_engine()

        # Предзагрузка модели в фоне — через 200мс после появления окна
        self.after(200, self._preload_model_in_background)

    # ─── Подписка на события Engine ─────────────────────────────────────

    def _subscribe_to_engine(self):
        """Подписаться на события Engine."""
        ev = self.engine.events
        ev.subscribe("documents_found", self._on_documents_found)
        ev.subscribe("documents_not_found", self._on_documents_not_found)
        ev.subscribe("extraction_started", self._on_extraction_started)
        ev.subscribe("check_started", self._on_check_started)
        ev.subscribe("sentence_start", self._on_sentence_start)
        ev.subscribe("sentence_checked", self._on_sentence_checked)
        ev.subscribe("check_complete", self._on_check_complete)
        ev.subscribe("check_error", self._on_check_error)

    # ─── Предзагрузка модели ────────────────────────────────────────────

    def _preload_model_in_background(self):
        """Загрузить ML-модель в фоновом потоке (не блокирует UI)."""
        def _load():
            import spell_checker
            # Пустой список — модель загрузится, но проверять нечего
            spell_checker.check_sentences_async(
                sentences=[],
                on_progress=lambda *a: None,
                on_complete=lambda: None,
            )
            logger.info("Model preloaded in background")

        threading.Thread(target=_load, daemon=True).start()

    # ─── Обработчики событий ────────────────────────────────────────────

    def _on_extraction_started(self):
        """Мгновенно переключить на экран предложений, заблокировать «Назад»."""
        self.sentences = []
        self.check_results = {}
        self.selected_sentence_index = None
        self.is_checking = True
        self._extraction_in_progress = True
        self._clear_tiles()
        self.current_view = "sentences"
        self._update_button_state()
        self.status_label.config(text="⏳ Извлечение предложений...")
        self.update()  # полная перерисовка окна до COM-вызова

    def _on_documents_found(self, documents):
        self.after(0, lambda: self._render_documents(documents))

    def _on_documents_not_found(self):
        self.after(0, lambda: self.status_label.config(text="Документы не найдены"))

    def _on_check_started(self, total):
        def _update():
            import spell_checker

            self.sentences = self.engine.sentences
            self.check_results = self.engine.check_results
            self.selected_sentence_index = None
            self.is_checking = True
            self._extraction_in_progress = False  # извлечение завершено
            self._refresh_check_buttons()
            self._clear_tiles()
            self.current_view = "sentences"
            self._update_button_state()

            # Если модель ещё не загружена — показать индикатор
            if not spell_checker.is_model_loaded():
                self.status_label.config(text="⏳ Загрузка модели...")
                # Обновить UI немедленно перед тяжёлой операцией
                self.update_idletasks()

            for sentence in self.sentences:
                self._create_sentence_tile_checking_ui(sentence)

            self.status_label.config(text=f"Проверка: 0/{total}")

        self.after(0, _update)

    def _on_sentence_start(self, index):
        def _update():
            if self.spinner_job:
                self.after_cancel(self.spinner_job)

            self.current_checking_index = index
            self.spinner_index = 0

            if index in self.spinner_labels:
                label = self.spinner_labels[index]
                if label.winfo_exists():
                    label.config(text=SPINNER_FRAMES[0], fg="#007bff")

            self.spinner_job = self.after(SPINNER_DELAY, self._animate_spinner)

        self.after(0, _update)

    def _on_sentence_checked(self, index, original, corrected, has_error):
        def _update():
            self.check_results = self.engine.check_results
            self._update_sentence_tile_ui(index, original, corrected, has_error)

            if self.hide_clean_var.get() and self._should_hide_tile(index):
                tile = self.sentence_tiles.get(index)
                if tile and tile.winfo_exists():
                    tile.pack_forget()

            self._update_status_text()

        self.after(0, _update)

    def _on_check_complete(self):
        def _update():
            self._stop_spinner()
            self.is_checking = False
            self._refresh_check_buttons()

            if self.available_adapters:
                self.adapter_combo.config(state="readonly")

            if self.current_view == "documents":
                self._update_doc_status_indicator()

            self._update_status_text()

        self.after(0, _update)

    def _on_check_error(self, error):
        def _update():
            logger.error("Check error: %s", error)
            self._stop_spinner()
            self.is_checking = False
            self._extraction_in_progress = False  # сброс при ошибке
            self._refresh_check_buttons()
            self._update_button_state()
            self.status_label.config(text=f"Ошибка: {error}")

            if self.available_adapters:
                self.adapter_combo.config(state="readonly")

        self.after(0, _update)

    # ─── Позиционирование ───────────────────────────────────────────────

    def _position_on_active_monitor(self):
        """Разместить окно справа на активном мониторе."""
        self.monitor = get_active_monitor_workarea()
        win_width = 350
        win_height = self.monitor["height"] - 40
        x = self.monitor["x"] + self.monitor["width"] - win_width
        y = self.monitor["y"]
        self.geometry(f"{win_width}x{win_height}+{x}+{y}")

    # ─── Создание UI ────────────────────────────────────────────────────

    def _create_ui(self):
        # --- Панель инструментов ---
        self.toolbar = ttk.Frame(self)
        self.toolbar.pack(fill=tk.X, padx=10, pady=(10, 5))

        self.find_button = tk.Button(
            self.toolbar, text="  \u27f3  Найти документы  ",
            command=self.find_documents, bg="#2B579A", fg="white",
            activebackground="#1E3F70", activeforeground="white",
            relief="flat", borderwidth=0, padx=12, pady=5,
            font=("Segoe UI", 9), cursor="hand2",
        )
        self.find_button.pack(fill=tk.X)
        self._set_button_hover(self.find_button, "#2B579A", "#1E3F70")

        self.toolbar_separator = ttk.Separator(self.toolbar, orient="horizontal")
        self.toolbar_separator.pack(fill=tk.X, pady=5)

        self.check_frame = ttk.Frame(self.toolbar)
        self.check_frame.pack(fill=tk.X)

        self.check_button = tk.Button(
            self.check_frame, text="  \u2713  Проверить всё  ",
            command=self.check_selected_document, bg="#aaaaaa", fg="#666666",
            activebackground="#aaaaaa", activeforeground="#666666",
            relief="flat", borderwidth=0, padx=8, pady=4,
            font=("Segoe UI", 9), cursor="arrow",
        )
        self.check_button.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.check_selection_button = tk.Button(
            self.check_frame, text="  \u2702  Проверить выделенное  ",
            command=self.check_selected_fragment, bg="#aaaaaa", fg="#666666",
            activebackground="#aaaaaa", activeforeground="#666666",
            relief="flat", borderwidth=0, padx=8, pady=4,
            font=("Segoe UI", 9), cursor="arrow",
        )
        self.check_selection_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))

        # Опции на экране документов
        self.doc_options_frame = ttk.Frame(self.toolbar)
        self.doc_options_frame.pack(fill=tk.X, pady=(5, 0))

        self.strict_protect_var = tk.BooleanVar(
            value=self._config.get("strict_protection", False)
        )
        self.strict_protect_check = ttk.Checkbutton(
            self.doc_options_frame, text="",
            variable=self.strict_protect_var, command=self._on_toggle_strict_protection,
        )

        self.auditor_format_var = tk.BooleanVar(
            value=self._config.get("auditor_format", False)
        )
        self.auditor_format_check = ttk.Checkbutton(
            self.doc_options_frame, text="Аудиторский формат",
            variable=self.auditor_format_var, command=self._on_toggle_auditor_format,
        )
        self.auditor_format_check.pack(side=tk.LEFT, padx=(10, 0))

        self.skip_tables_var = tk.BooleanVar(
            value=self._config.get("skip_tables", False)
        )
        self.skip_tables_check = ttk.Checkbutton(
            self.doc_options_frame, text="Пропускать таблицы",
            variable=self.skip_tables_var, command=self._on_toggle_skip_tables,
        )
        self.skip_tables_check.pack(side=tk.LEFT, padx=(10, 0))

        # Фрейм скрытия чистых предложений
        self.toggle_frame = ttk.Frame(self.toolbar)

        self.hide_clean_var = tk.BooleanVar(
            value=self._config.get("hide_clean_sentences", True)
        )
        self.hide_clean_check = ttk.Checkbutton(
            self.toggle_frame, text="Скрыть предложения без ошибок",
            variable=self.hide_clean_var, command=self._on_toggle_hide_clean,
        )
        self.hide_clean_check.pack(side=tk.LEFT)

        # Фрейм выбора адаптера
        self.adapter_frame = ttk.Frame(self)
        ttk.Label(self.adapter_frame, text="Адаптер:").pack(side=tk.LEFT, padx=(0, 5))

        default_adapter = self._config.get("default_adapter", "")
        if self.available_adapters:
            if default_adapter not in self.available_adapters:
                default_adapter = self.available_adapters[0]
        else:
            default_adapter = ""

        self.selected_adapter = tk.StringVar(value=default_adapter)
        self.adapter_combo = ttk.Combobox(
            self.adapter_frame, textvariable=self.selected_adapter,
            values=self.available_adapters,
            state="readonly" if self.available_adapters else "disabled",
        )
        self.adapter_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.adapter_combo.bind("<<ComboboxSelected>>", self._on_adapter_changed)

        # Фрейм со скроллом
        container = ttk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.docs_frame = ttk.Frame(self.canvas)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.docs_frame, anchor="nw"
        )

        self.docs_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.status_label = ttk.Label(self, text="Нажмите 'Найти документы'")
        self.status_label.pack(pady=(0, 10))

    # ─── Прокрутка ──────────────────────────────────────────────────────

    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ─── Кнопки ─────────────────────────────────────────────────────────

    def _set_button_hover(self, btn, bg, hover_bg):
        btn.bind("<Enter>", lambda e: btn.config(bg=hover_bg))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg))

    def _update_button_state(self):
        """Обновить видимость элементов toolbar в зависимости от текущего вида."""
        self.toolbar_separator.pack_forget()
        self.check_frame.pack_forget()
        self.doc_options_frame.pack_forget()
        self.toggle_frame.pack_forget()

        if self.current_view == "documents":
            self.find_button.config(
                text="  \u27f3  Найти документы  ",
                command=self.find_documents, bg="#2B579A",
            )
            self._set_button_hover(self.find_button, "#2B579A", "#1E3F70")
            self.toolbar_separator.pack(fill=tk.X, pady=5)
            self.check_frame.pack(fill=tk.X)
            self.doc_options_frame.pack(fill=tk.X, pady=(5, 0))
            self._refresh_check_buttons()
        else:
            extracting = self._extraction_in_progress
            bg = "#5B6770" if not extracting else "#888888"
            fg = "white" if not extracting else "#666666"
            cur = "hand2" if not extracting else "arrow"
            self.find_button.config(
                text="  \u25c2  Назад  ",
                command=self.show_documents_view if not extracting else None,
                bg=bg, fg=fg, activebackground=bg, activeforeground=fg, cursor=cur,
            )
            if not extracting:
                self._set_button_hover(self.find_button, "#5B6770", "#4A545C")
            self.toggle_frame.pack(fill=tk.X, pady=(5, 0))

    def _refresh_check_buttons(self):
        """Обновить состояние кнопок проверки (активна/неактивна)."""
        has_doc = self.engine.selected_doc is not None and not self.engine.is_checking
        bg = "#4472C4" if has_doc else "#aaaaaa"
        fg = "white" if has_doc else "#666666"
        cur = "hand2" if has_doc else "arrow"
        self.check_button.config(
            bg=bg, fg=fg, activebackground=bg, activeforeground=fg, cursor=cur,
            text="  \u2713  Проверяю...  " if self.engine.is_checking else "  \u2713  Проверить всё  ",
        )
        self.check_selection_button.config(
            bg=bg, fg=fg, activebackground=bg, activeforeground=fg, cursor=cur,
            text="  \u2702  Проверяю...  " if self.engine.is_checking else "  \u2702  Проверить выделенное  ",
        )

    # ─── Callback'и настроек ────────────────────────────────────────────

    def _on_adapter_changed(self, event=None):
        self.engine.set_default_adapter(self.selected_adapter.get())

    def _on_toggle_strict_protection(self):
        self.engine.set_config("strict_protection", self.strict_protect_var.get())

    def _on_toggle_auditor_format(self):
        self.engine.set_config("auditor_format", self.auditor_format_var.get())

    def _on_toggle_skip_tables(self):
        self.engine.set_config("skip_tables", self.skip_tables_var.get())

    def _on_toggle_hide_clean(self):
        self.engine.set_config("hide_clean_sentences", self.hide_clean_var.get())
        if self.current_view == "sentences":
            self._apply_sentence_filter()

    # ─── Действия пользователя ──────────────────────────────────────────

    def find_documents(self):
        self.engine.find_documents()

    def check_selected_document(self):
        self.engine.check_document()

    def check_selected_fragment(self):
        self.engine.check_fragment()

    def show_documents_view(self):
        """Вернуться к списку документов."""
        self._stop_spinner()
        self._extraction_in_progress = False
        self.current_view = "documents"
        self._update_button_state()

        for widget in self.docs_frame.winfo_children():
            widget.destroy()
        self.tile_frames = {}
        self.doc_status_label = None
        self.canvas.yview_moveto(0)

        docs = self.engine.documents
        for doc in docs:
            self._create_document_tile_ui(doc)

        self._highlight_selected_tile_ui()

        if self.engine.selected_doc and (self.engine.is_checking or self.engine.check_results):
            self._update_doc_status_indicator()
            if self.engine.is_checking and self.current_checking_index >= 0:
                self.spinner_job = self.after(SPINNER_DELAY, self._animate_spinner)

        self._update_status_text()

    # ─── Рендер документов ──────────────────────────────────────────────

    def _render_documents(self, documents):
        """Отрисовать список документов."""
        self._stop_spinner()
        self._cancel_worker_ui()
        self.engine.doc_state.next_generation()
        self.engine.doc_state.clear()
        self.check_results = {}
        self.sentences = []
        self.engine.is_checking = False
        self.engine.selected_doc = None
        self._refresh_check_buttons()
        self.tile_frames = {}

        for widget in self.docs_frame.winfo_children():
            widget.destroy()
        self.canvas.yview_moveto(0)

        for doc in documents:
            self._create_document_tile_ui(doc)

        # Динамический статус-бар по зарегистрированным провайдерам
        counts = {}
        for provider in get_all_providers():
            counts[provider.doc_type] = sum(
                1 for d in documents if d.get("type") == provider.doc_type
            )
        status_parts = [f"{t}: {c}" for t, c in counts.items() if c > 0]
        self.status_label.config(text=", ".join(status_parts) if status_parts else "Документы не найдены")

    def _create_document_tile_ui(self, doc):
        """Создать плитку документа."""
        is_selected = (
            self.engine.selected_doc
            and id(doc) == id(self.engine.selected_doc)
        )
        is_active = is_selected and (
            self.engine.is_checking or self.engine.check_results
        )
        cached = self.engine.doc_state.load(id(doc))
        cached_errors = cached["check_results"] if cached else None

        tile, status_label = create_document_tile(
            self.docs_frame, doc, self._on_tile_click,
            is_selected=is_selected, cached_errors=cached_errors,
            is_active_check=is_active,
        )
        self.tile_frames[id(doc)] = tile
        if status_label:
            self.doc_status_label = status_label

    def _on_tile_click(self, doc):
        if (
            self.engine.selected_doc
            and id(doc) == id(self.engine.selected_doc)
            and (self.engine.is_checking or self.engine.check_results)
        ):
            self._show_sentences_from_cache()
            return

        self.engine.select_document(doc)
        self._refresh_check_buttons()
        self._highlight_selected_tile_ui()

        app_width = 350
        target_rect = (
            self.monitor["x"],
            self.monitor["y"],
            self.monitor["width"] - app_width,
            self.monitor["height"],
        )
        office_finder.activate_document(doc, target_rect)
        self.after(100, self._raise_window)

        # Если у документа есть кэшированные результаты — показать предложения
        if self.engine.sentences:
            self._show_sentences_from_cache()

    def _highlight_selected_tile_ui(self):
        highlight_selected_tile(self.tile_frames, self.engine.selected_doc)

    # ─── Спиннер ────────────────────────────────────────────────────────

    def _animate_spinner(self):
        """Анимировать спиннер проверки."""
        if self.current_checking_index < 0:
            return

        self.spinner_index = (self.spinner_index + 1) % len(SPINNER_FRAMES)

        if self.current_view == "sentences":
            if self.current_checking_index in self.spinner_labels:
                label = self.spinner_labels[self.current_checking_index]
                if label.winfo_exists():
                    label.config(text=SPINNER_FRAMES[self.spinner_index])
        else:
            self._update_doc_status_indicator()

        self.spinner_job = self.after(SPINNER_DELAY, self._animate_spinner)

    def _stop_spinner(self):
        """Остановить спиннер."""
        if self.spinner_job:
            self.after_cancel(self.spinner_job)
            self.spinner_job = None
        self.current_checking_index = -1

    def _cancel_worker_ui(self):
        """Отменить текущий worker проверки."""
        self.engine.doc_state.cancel()

    def _update_doc_status_indicator(self):
        """Обновить индикатор статуса на плитке документа."""
        if not self.doc_status_label or not self.doc_status_label.winfo_exists():
            return

        total = len(self.engine.sentences)
        checked = len(self.engine.check_results)

        if self.engine.is_checking:
            frame = SPINNER_FRAMES[self.spinner_index]
            self.doc_status_label.config(text=f"{frame} {checked}/{total}", fg="#007bff")
        else:
            errors = sum(1 for r in self.engine.check_results.values() if r["has_error"])
            if errors > 0:
                self.doc_status_label.config(text=f"✗ {errors}", fg="#dc3545")
            else:
                self.doc_status_label.config(text="✓", fg="#28a745")

    # ─── Предложения ────────────────────────────────────────────────────

    def _show_sentences_from_cache(self):
        """Показать предложения из кэша (при повторном входе в документ)."""
        self.current_view = "sentences"
        self.selected_sentence_index = None
        self._update_button_state()

        for widget in self.docs_frame.winfo_children():
            widget.destroy()
        self.canvas.yview_moveto(0)

        self.sentences = self.engine.sentences
        self.check_results = self.engine.check_results
        self.sentence_tiles = {}
        self.spinner_labels = {}

        for sentence in self.sentences:
            index = sentence["index"]
            if index in self.check_results:
                result = self.check_results[index]
                self._create_checked_sentence_tile_ui(sentence, result)
            else:
                self._create_sentence_tile_checking_ui(sentence)

        if self.engine.is_checking and self.current_checking_index >= 0:
            self.spinner_job = self.after(SPINNER_DELAY, self._animate_spinner)

        if self.hide_clean_var.get():
            self._apply_sentence_filter()

        self._update_status_text()

    def _create_sentence_tile_checking_ui(self, sentence):
        tile, status_label = create_sentence_tile_checking(
            self.docs_frame, sentence, self._on_sentence_click_ui,
        )
        index = sentence["index"]
        self.sentence_tiles[index] = tile
        self.spinner_labels[index] = status_label

    def _create_checked_sentence_tile_ui(self, sentence, result):
        info = create_checked_sentence_tile(
            self.docs_frame, sentence, result,
            on_click=self._on_sentence_click_ui,
            on_apply=self._toggle_apply,
            on_skip=self._toggle_skip,
        )
        index = sentence["index"]
        self.sentence_tiles[index] = info["tile"]
        if info["buttons"]:
            self.tile_buttons[index] = info["buttons"]

    def _update_sentence_tile_ui(self, index, original, corrected, has_error):
        """Обновить плитку предложения после проверки."""
        if index not in self.sentence_tiles:
            return

        tile = self.sentence_tiles[index]
        if not tile.winfo_exists():
            return

        if index in self.spinner_labels:
            del self.spinner_labels[index]

        sentence = self.sentences[index] if index < len(self.sentences) else {"index": index}
        info = update_sentence_tile(
            tile, index, original, corrected, has_error, sentence,
            on_click=lambda s=sentence: self._on_sentence_click_ui(s),
            on_apply=self._toggle_apply,
            on_skip=self._toggle_skip,
        )
        if info["buttons"]:
            self.tile_buttons[index] = info["buttons"]

    def _on_sentence_click_ui(self, sentence):
        """Клик по плитке предложения — навигация + подсветка."""
        index = sentence.get("index")
        self._select_sentence_tile(index)

        success = self.engine.navigate_to_sentence(sentence)
        if not success:
            self.status_label.config(
                text=f"⚠ Не удалось найти предложение #{index + 1}"
            )

        app_width = 350
        target_rect = (
            self.monitor["x"],
            self.monitor["y"],
            self.monitor["width"] - app_width,
            self.monitor["height"],
        )
        office_finder.activate_document(self.engine.selected_doc, target_rect)

    def _select_sentence_tile(self, index):
        """Подсветить плитку выбранного предложения, остальные сбросить."""
        self.selected_sentence_index = index
        highlight_selected_sentence_tile(self.sentence_tiles, index)

    # ─── Фильтр ─────────────────────────────────────────────────────────

    def _should_hide_tile(self, index):
        """Определить, нужно ли скрыть плитку (предложение без ошибок)."""
        if index not in self.engine.check_results:
            return False
        result = self.engine.check_results[index]
        if result["has_error"]:
            return False
        return True

    def _apply_sentence_filter(self):
        """Применить фильтр скрытия чистых предложений."""
        hide = self.hide_clean_var.get()

        for index in sorted(self.sentence_tiles.keys()):
            tile = self.sentence_tiles[index]
            if tile.winfo_exists():
                tile.pack_forget()

        for index in sorted(self.sentence_tiles.keys()):
            tile = self.sentence_tiles[index]
            if not tile.winfo_exists():
                continue
            if hide and self._should_hide_tile(index):
                continue
            tile.pack(fill=tk.X, pady=2, padx=2)

        self._update_status_text()

    # ─── Статус ─────────────────────────────────────────────────────────

    def _update_status_text(self):
        """Обновить текст в строке статуса."""
        if self.current_view == "documents":
            docs = self.engine.documents
            counts = {}
            for provider in get_all_providers():
                counts[provider.doc_type] = sum(
                    1 for d in docs if d.get("type") == provider.doc_type
                )
            status_parts = [f"{t}: {c}" for t, c in counts.items() if c > 0]
            self.status_label.config(
                text=", ".join(status_parts) if status_parts else "Документы не найдены"
            )
        elif self.engine.is_checking:
            hidden = 0
            if self.hide_clean_var.get():
                hidden = sum(
                    1 for idx in self.engine.check_results if self._should_hide_tile(idx)
                )
            suffix = f" (скрыто: {hidden})" if hidden else ""
            self.status_label.config(
                text=f"Проверка: {len(self.engine.check_results)}/{len(self.engine.sentences)}{suffix}"
            )
        elif self.engine.check_results:
            errors = sum(1 for r in self.engine.check_results.values() if r["has_error"])
            hidden = 0
            if self.hide_clean_var.get():
                hidden = sum(
                    1 for idx in self.engine.check_results if self._should_hide_tile(idx)
                )
            suffix = f" (скрыто: {hidden})" if hidden else ""
            self.status_label.config(
                text=f"Готово. Ошибок: {errors}/{len(self.engine.sentences)}{suffix}"
            )

    # ─── Очистка плиток ─────────────────────────────────────────────────

    def _clear_tiles(self):
        """Удалить все плитки из canvas."""
        for widget in self.docs_frame.winfo_children():
            widget.destroy()
        self.canvas.yview_moveto(0)

    # ─── Применить/Пропустить ───────────────────────────────────────────

    def _toggle_apply(self, index):
        """Переключить применение исправления (применить ↔ отменить)."""
        result = self.engine.check_results.get(index)
        if not result:
            return

        self._select_sentence_tile(index)

        current_state = result.get("state", "pending")
        buttons = self.tile_buttons.get(index)
        if not buttons:
            return

        if current_state == "pending":
            success = self.engine.apply_correction(index)
            if success:
                result["state"] = "applied"
                buttons["apply"].config(text="Отменить", bg="#28a745")
                buttons["skip"].config(state="disabled")
                if "index_label" in buttons:
                    buttons["index_label"].config(bg="#28a745")

        elif current_state == "applied":
            success = self.engine.revert_correction(index)
            if success:
                result["state"] = "pending"
                buttons["apply"].config(text="Применить", bg="#dc3545")
                buttons["skip"].config(state="normal")
                if "index_label" in buttons:
                    buttons["index_label"].config(bg="#666666")

    def _toggle_skip(self, index):
        """Переключить пропуск предложения."""
        result = self.engine.check_results.get(index)
        if not result:
            return

        self._select_sentence_tile(index)

        current_state = result.get("state", "pending")
        buttons = self.tile_buttons.get(index)
        if not buttons:
            return

        new_state = self.engine.toggle_skip(index)
        if new_state == "skipped":
            buttons["apply"].config(state="disabled", bg="#6c757d")
            buttons["skip"].config(text="Отменить", bg="#28a745")
            if "index_label" in buttons:
                buttons["index_label"].config(bg="#28a745")
        else:
            buttons["apply"].config(state="normal", bg="#dc3545")
            buttons["skip"].config(text="Пропустить", bg="#6c757d")
            if "index_label" in buttons:
                buttons["index_label"].config(bg="#666666")

    # ─── Поднять окно ───────────────────────────────────────────────────

    def _raise_window(self):
        """Поднять окно приложения на передний план."""
        self.lift()
        self.focus_force()
