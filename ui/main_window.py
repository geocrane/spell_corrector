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
from ui import icons
from ui.constants import (
    SPINNER_FRAMES,
    SPINNER_DELAY,
    RIBBON_BG,
    RIBBON_NORMAL_FG,
)
from ui.overlay import BusyOverlay
from ui.ribbon import RibbonButton, RibbonToggle
from ui.tiles import (
    create_document_tile,
    highlight_selected_tile,
    highlight_selected_sentence_tile,
    create_sentence_tile_checking,
    create_checked_sentence_tile,
    create_word_correction_tile,
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


APP_WIDTH = 470  # ширина бокового окна; Word открывается на monitor_width - APP_WIDTH


class MainWindow(tk.Tk):
    """Главное окно приложения. Управляет UI, все вызовы бизнес-логики делегирует в Engine."""

    def __init__(self, engine):
        super().__init__()

        self.engine = engine
        self.title("Корректор орфографии")
        self.geometry(f"{APP_WIDTH}x500")
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
        self._tile_click_in_progress = False  # анти-дребезг быстрых кликов по плиткам
        self.errors_tiles = []  # [(correction, tile_frame), ...] для режима ошибок
        self.revisions_mode_var = tk.BooleanVar(value=False)

        # Универсальный overlay-индикатор долгих операций (создаётся в _create_ui).
        self.overlay: BusyOverlay | None = None
        self._all_toolbar_buttons = []

        # Конфигурация
        self._config = engine.get_config()
        self.available_adapters = engine.get_available_adapters()

        self._create_ui()
        self._subscribe_to_engine()
        self._update_button_state()

        # Предзагрузка модели в фоне — через 200мс после появления окна
        self.after(200, self._preload_model_in_background)

    # ─── Подписка на события Engine ─────────────────────────────────────

    def _subscribe_to_engine(self):
        """Подписаться на события Engine."""
        ev = self.engine.events
        ev.subscribe("documents_found", self._on_documents_found)
        ev.subscribe("documents_not_found", self._on_documents_not_found)
        ev.subscribe("extraction_started", self._on_extraction_started)
        ev.subscribe("extraction_progress", self._on_extraction_progress)
        ev.subscribe("check_started", self._on_check_started)
        ev.subscribe("sentence_start", self._on_sentence_start)
        ev.subscribe("sentence_checked", self._on_sentence_checked)
        ev.subscribe("check_complete", self._on_check_complete)
        ev.subscribe("check_error", self._on_check_error)
        ev.subscribe("word_corrections_changed", self._on_word_corrections_changed)
        ev.subscribe("revisions_mode_changed", self._on_revisions_mode_changed)
        ev.subscribe("navigation_blocked", self._on_navigation_blocked)
        ev.subscribe("format_state_changed", self._on_format_state_changed)
        ev.subscribe("unify_progress", self._on_unify_progress)

    # ─── Предзагрузка модели ────────────────────────────────────────────

    def _preload_model_in_background(self):
        """Загрузить ML-модель в фоновом потоке (не блокирует UI).

        Показывает overlay «Загрузка модели…» в центральном поле, чтобы
        пользователь видел, что приложение готовится, а не «зависло».
        """
        import spell_checker

        if spell_checker.is_model_loaded():
            return

        self.overlay.show("Загрузка модели, подождите…", cancelable=False)

        def _load():
            try:
                spell_checker.SpellChecker.get_instance().load_model()
                logger.info("Model preloaded in background")
            except Exception:
                logger.exception("Model preload failed")
            finally:
                self.after(0, self.overlay.hide)

        threading.Thread(target=_load, daemon=True).start()

    # ─── Обработчики событий ────────────────────────────────────────────

    def _on_extraction_started(self):
        """Мгновенно переключить на экран предложений, заблокировать «Назад».

        Overlay показывается без кнопки «Остановить»: Word-провайдер
        извлекает предложения одним монолитным COM-вызовом и не имеет
        точек реальной отмены, а синхронный update() во время этого
        вызова вызывал артефакты в тексте документа.
        """
        self.sentences = []
        self.check_results = {}
        self.selected_sentence_index = None
        self.is_checking = True
        self._extraction_in_progress = True
        self._clear_tiles()
        self.current_view = "sentences"
        self._update_button_state()
        self.status_label.config(text="⏳ Извлечение предложений...")
        self.overlay.show("Извлечение предложений…", cancelable=False)
        self.update()  # полная перерисовка окна до COM-вызова

    def _on_extraction_progress(self, extracted, processed):
        """Обновить текст overlay (без прокачки Tk-очереди).

        Не вызываем self.update() синхронно — это могло приводить к
        re-entrancy во время COM-вызова extract и портить состояние Word.
        Обновление сообщения через self.after(0, ...) выполнится, когда
        Tk-цикл вернёт управление.
        """
        def _update():
            try:
                self.overlay.update_message(f"Извлечено {extracted} предложений…")
            except tk.TclError:
                pass
        self.after(0, _update)

    def _on_documents_found(self, documents):
        self.after(0, lambda: self._render_documents(documents))

    def _on_documents_not_found(self):
        # Полная перерисовка с пустым списком: убирает зависшие плитки
        # от предыдущего поиска (например, если документы были закрыты).
        self.after(0, lambda: self._render_documents([]))

    def _on_check_started(self, total):
        def _update():
            # Извлечение завершилось — overlay прогресса скрываем; во время
            # самой проверки заглушка не нужна (плитки сами по себе наглядны).
            self.overlay.hide()

            self.sentences = self.engine.sentences
            self.check_results = self.engine.check_results
            self.selected_sentence_index = None
            self.is_checking = True
            self._extraction_in_progress = False  # извлечение завершено
            self._refresh_check_buttons()
            self._clear_tiles()
            self.current_view = "sentences"
            self._update_button_state()

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
            self.overlay.hide()
            self._stop_spinner()
            self.is_checking = False
            self._extraction_in_progress = False  # сброс при ошибке
            self._refresh_check_buttons()
            self._update_button_state()
            self.status_label.config(text=f"Ошибка: {error}")

            if self.available_adapters:
                self.adapter_combo.config(state="readonly")

        self.after(0, _update)

    def _on_navigation_blocked(self, message):
        self.after(0, lambda: self.status_label.config(text=message))

    # ─── Позиционирование ───────────────────────────────────────────────

    def _position_on_active_monitor(self):
        """Разместить окно справа на активном мониторе."""
        self.monitor = get_active_monitor_workarea()
        win_width = APP_WIDTH
        win_height = self.monitor["height"] - 40
        x = self.monitor["x"] + self.monitor["width"] - win_width
        y = self.monitor["y"]
        self.geometry(f"{win_width}x{win_height}+{x}+{y}")

    # ─── Создание UI ────────────────────────────────────────────────────

    def _create_ui(self):
        # --- Панель инструментов ---
        self.toolbar = ttk.Frame(self)
        self.toolbar.pack(fill=tk.X, padx=10, pady=(8, 5))

        # Ряд 1: главная кнопка (find или back, взаимоисключающие).
        # В режиме sentences сюда же упаковываются hide_clean_toggle и errors_mode_button.
        self.primary_row = ttk.Frame(self.toolbar)
        self.primary_row.pack(fill=tk.X)

        self.find_button = RibbonButton(
            self.primary_row,
            icon_key="find",
            text="Найти",
            command=self.find_documents,
            icon_size=icons.ICON_LG,
            orient="vertical",
            style="default",
        )

        self.back_button = RibbonButton(
            self.primary_row,
            icon_key="back",
            text="Назад",
            command=self.show_documents_view,
            icon_size=icons.ICON_SM,
            orient="vertical",
            style="default",
            compact=True,
        )

        # Вертикальный сепаратор между «Найти» и группой проверки (только documents)
        self.primary_sep = ttk.Separator(self.primary_row, orient="vertical")

        self.check_button = RibbonButton(
            self.primary_row,
            icon_key="check_all",
            text="Весь текст",
            command=self.check_selected_document,
            icon_size=icons.ICON_LG,
            orient="vertical",
            style="default",
        )

        self.check_selection_button = RibbonButton(
            self.primary_row,
            icon_key="check_selection",
            text="Фрагмент",
            command=self.check_selected_fragment,
            icon_size=icons.ICON_LG,
            orient="vertical",
            style="default",
        )

        self.toolbar_separator = ttk.Separator(self.toolbar, orient="horizontal")

        # Ряд 3: одна строка — форматирование (LEFT) + опции документа (RIGHT)
        self.combo_row = ttk.Frame(self.toolbar)

        self.format_unify_frame = ttk.Frame(self.combo_row)
        self.format_unify_frame.pack(side=tk.LEFT)

        # «Применить всё» — главная кнопка группы (слева, с отступом-сепаратором)
        self.format_all_btn = RibbonToggle(
            self.format_unify_frame,
            icon_key="format_all",
            text="Всё",
            command=self._on_format_toggle_all,
            icon_size=icons.ICON_SM,
            compact=True,
            initial_state="disabled",
        )
        self.format_all_btn.pack(side=tk.LEFT, padx=(0, 2))

        self.format_group_sep = ttk.Separator(
            self.format_unify_frame, orient="vertical"
        )
        self.format_group_sep.pack(side=tk.LEFT, fill=tk.Y, padx=4)

        self.format_buttons: dict[str, RibbonToggle] = {}
        for icon_key, label, attr in (
            ("format_font", "Шрифт", "font_name"),
            ("format_size", "Размер", "font_size"),
            ("format_style", "Стиль", "style"),
        ):
            tg = RibbonToggle(
                self.format_unify_frame,
                icon_key=icon_key,
                text=label,
                command=lambda a=attr: self._on_format_toggle(a),
                icon_size=icons.ICON_SM,
                compact=True,
                initial_state="disabled",
            )
            tg.pack(side=tk.LEFT, padx=1)
            self.format_buttons[attr] = tg

        # Правая группа: Аудитор + Пропуск таблиц
        self.doc_options_frame = ttk.Frame(self.combo_row)

        self.auditor_format_var = tk.BooleanVar(
            value=self._config.get("auditor_format", False)
        )
        self.auditor_toggle = RibbonToggle(
            self.doc_options_frame,
            icon_key="auditor",
            text="Аудитор",
            command=self._on_toggle_auditor_format,
            icon_size=icons.ICON_SM,
            compact=True,
            initial_state="on" if self.auditor_format_var.get() else "off",
        )
        self.auditor_toggle.pack(side=tk.LEFT, padx=(6, 1))

        self.skip_tables_var = tk.BooleanVar(
            value=self._config.get("skip_tables", False)
        )
        self.skip_tables_toggle = RibbonToggle(
            self.doc_options_frame,
            icon_key="skip_tables",
            text="Без таблиц",
            command=self._on_toggle_skip_tables,
            icon_size=icons.ICON_SM,
            compact=True,
            initial_state="on" if self.skip_tables_var.get() else "off",
        )
        self.skip_tables_toggle.pack(side=tk.LEFT, padx=1)

        # Скрытая переменная strict_protection — поведение сохранено.
        self.strict_protect_var = tk.BooleanVar(
            value=self._config.get("strict_protection", False)
        )

        # Скрыть / Исправления (только sentences). Упаковываются в primary_row,
        # чтобы шапка sentences оставалась одной строкой.
        self.hide_clean_var = tk.BooleanVar(
            value=self._config.get("hide_clean_sentences", True)
        )
        self.hide_clean_toggle = RibbonToggle(
            self.primary_row,
            icon_key="hide_clean",
            text="Скрыть",
            command=self._on_toggle_hide_clean,
            icon_size=icons.ICON_SM,
            orient="vertical",
            compact=True,
            initial_state="on" if self.hide_clean_var.get() else "off",
        )

        self.errors_mode_button = RibbonButton(
            self.primary_row,
            icon_key="errors_mode",
            text="Исправления",
            command=self._toggle_errors_view,
            icon_size=icons.ICON_SM,
            orient="vertical",
            style="danger",
            compact=True,
        )

        # Word ревизии — только errors
        self.revisions_toggle = RibbonToggle(
            self.primary_row,
            icon_key="revisions",
            text="Word ревизии",
            command=self._on_toggle_revisions_mode,
            icon_size=icons.ICON_SM,
            orient="vertical",
            compact=True,
            initial_state="on" if self.revisions_mode_var.get() else "off",
        )

        # Перечень всех кнопок toolbar — для массовой блокировки/разблокировки
        # во время долгих операций (например, унификация форматирования).
        self._all_toolbar_buttons = [
            self.find_button, self.back_button,
            self.check_button, self.check_selection_button,
            self.format_all_btn, *self.format_buttons.values(),
            self.auditor_toggle, self.skip_tables_toggle,
            self.hide_clean_toggle, self.errors_mode_button, self.revisions_toggle,
        ]

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
            self.adapter_frame,
            textvariable=self.selected_adapter,
            values=self.available_adapters,
            state="readonly" if self.available_adapters else "disabled",
        )
        self.adapter_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.adapter_combo.bind("<<ComboboxSelected>>", self._on_adapter_changed)

        # Фрейм со скроллом
        container = ttk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(
            container, orient="vertical", command=self.canvas.yview
        )
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

        # Overlay создаём после canvas — он будет накладываться поверх scrollable-области.
        self.overlay = BusyOverlay(self.canvas)

    # ─── Прокрутка ──────────────────────────────────────────────────────

    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ─── Кнопки ─────────────────────────────────────────────────────────

    def _update_button_state(self):
        """Обновить видимость элементов toolbar в зависимости от текущего вида."""
        # Скрыть все потенциально показываемые блоки
        for w in (
            self.find_button,
            self.back_button,
            self.primary_sep,
            self.check_button,
            self.check_selection_button,
            self.toolbar_separator,
            self.combo_row,
            self.hide_clean_toggle,
            self.errors_mode_button,
            self.revisions_toggle,
        ):
            try:
                w.pack_forget()
            except tk.TclError:
                pass

        if self.current_view == "documents":
            self.find_button.set_enabled(True)
            # primary_row: [Найти] | [Весь текст] [Фрагмент]
            self.find_button.pack(in_=self.primary_row, side=tk.LEFT, padx=(0, 40))
            self.primary_sep.pack(in_=self.primary_row, side=tk.LEFT, fill=tk.Y, padx=4)
            self.check_button.pack(in_=self.primary_row, side=tk.LEFT, padx=2)
            self.check_selection_button.pack(in_=self.primary_row, side=tk.LEFT, padx=2)
            self.toolbar_separator.pack(fill=tk.X, pady=4)
            self.combo_row.pack(fill=tk.X, pady=(4, 0))
            self.format_unify_frame.pack(side=tk.LEFT)
            self.doc_options_frame.pack(side=tk.RIGHT)
            self._refresh_check_buttons()
            self._refresh_doc_dependent_options()
            self._refresh_format_unify_panel()
            self._refresh_doc_options_toggles()
        elif self.current_view == "errors":
            self.back_button.set_command(self._exit_errors_view)
            self.back_button.set_enabled(True)
            self.back_button.pack(in_=self.primary_row, side=tk.LEFT, padx=(0, 4))
            self.revisions_toggle.pack(in_=self.primary_row, side=tk.RIGHT, padx=(0, 4))
            self._refresh_revisions_toggle()
        else:  # sentences
            extracting = self._extraction_in_progress
            self.back_button.set_command(
                None if extracting else self.show_documents_view
            )
            self.back_button.set_enabled(not extracting)
            self.back_button.pack(in_=self.primary_row, side=tk.LEFT, padx=(0, 4))
            self.hide_clean_toggle.pack(in_=self.primary_row, side=tk.LEFT, padx=4)
            self.errors_mode_button.pack(in_=self.primary_row, side=tk.RIGHT, padx=(0, 4))
            self._refresh_errors_mode_button()
            self._refresh_hide_clean_toggle()

    def _refresh_check_buttons(self):
        """Обновить состояние кнопок проверки (активна/неактивна).

        Для Excel «Весь текст» disabled: на больших таблицах извлечение всех
        ячеек длится очень долго; пользователь сам должен выделить диапазон.
        """
        doc = self.engine.selected_doc
        is_excel = doc is not None and doc.get("type") == "excel"
        has_doc = doc is not None and not self.engine.is_checking
        is_checking = self.engine.is_checking
        if is_checking:
            self.check_button.set_text("Проверяю...")
            self.check_selection_button.set_text("Проверяю...")
        else:
            self.check_button.set_text("Весь текст")
            self.check_selection_button.set_text("Фрагмент")
        self.check_button.set_enabled(has_doc and not is_excel)
        self.check_selection_button.set_enabled(has_doc)

    def _refresh_format_unify_panel(self):
        """Синхронизировать состояния и доступность кнопок унификации.

        Сама панель видна всегда в режиме «documents» (упаковывается в
        _update_button_state). Здесь только красим toggle-состояние из
        engine.format_state и enable/disable по наличию выбранного документа.

        Для Excel унификация не поддерживается — все кнопки disabled.
        """
        doc = self.engine.selected_doc
        has_doc = doc is not None
        is_excel = has_doc and doc.get("type") == "excel"
        for attr, tg in self.format_buttons.items():
            st = self.engine.format_state.get(attr, {})
            is_active = bool(st.get("is_active"))
            if not has_doc or is_excel:
                tg.set_state("disabled")
            else:
                tg.set_state("on" if is_active else "off")
        self._refresh_format_all_btn(has_doc and not is_excel)

    def _refresh_format_all_btn(self, has_doc=None):
        """Обновить вид общей кнопки «Применить всё»: on / off / mixed / disabled."""
        if has_doc is None:
            has_doc = self.engine.selected_doc is not None
        states = [st.get("is_active") for st in self.engine.format_state.values()]
        if not has_doc:
            target = "disabled"
        elif states and all(states):
            target = "on"
        elif not any(states):
            target = "off"
        else:
            target = "mixed"
        self.format_all_btn.set_state(target)

    def _toolbar_set_busy(self, busy: bool):
        """Заблокировать/разблокировать toolbar на время долгой операции.

        При busy=True все кнопки toolbar переводятся в disabled. При busy=False
        состояния восстанавливаются штатными _refresh_* (которые пересчитывают
        их из engine), плюс _update_button_state — чтобы режимы sentences/errors
        вернули свои set_enabled.
        """
        if busy:
            for btn in self._all_toolbar_buttons:
                if isinstance(btn, RibbonToggle):
                    btn.set_state("disabled")
                else:  # RibbonButton
                    btn.set_enabled(False)
            return

        # Восстановление: сначала пересчёт состояний из engine, потом
        # _update_button_state для перепаковки/скрытия по текущему view.
        self._refresh_check_buttons()
        self._refresh_format_unify_panel()
        self._refresh_doc_dependent_options()
        self._refresh_doc_options_toggles()
        self._update_button_state()

    def _on_format_toggle(self, attr: str):
        """Клик по одной из кнопок Шрифт/Размер/Стиль.

        На время COM-вызова блокируем весь toolbar и показываем overlay
        со спиннером — операция блокирующая, и без явного индикатора
        пользователь принимает зависание за «зависший интерфейс».
        """
        if not self.engine.selected_doc:
            return
        method = {
            "font_name": self.engine.toggle_unify_font,
            "font_size": self.engine.toggle_unify_size,
            "style": self.engine.toggle_unify_style,
        }.get(attr)
        if method is None:
            return
        self._toolbar_set_busy(True)
        self.overlay.show(
            "Идёт унификация форматирования…",
            cancelable=True,
            on_cancel=self.engine.cancel_unify,
        )
        self.status_label.config(text="⏳ Унификация форматирования…")
        # Полная перерисовка — overlay и спиннер должны появиться ДО COM-вызова
        self.update()
        try:
            method()
        finally:
            self.overlay.hide()
            self._toolbar_set_busy(False)

    def _on_format_toggle_all(self):
        """Клик по «Применить всё»."""
        if not self.engine.selected_doc:
            return
        self._toolbar_set_busy(True)
        self.overlay.show(
            "Идёт унификация форматирования…",
            cancelable=True,
            on_cancel=self.engine.cancel_unify,
        )
        self.status_label.config(text="⏳ Унификация форматирования…")
        self.update()
        try:
            self.engine.toggle_unify_all()
        finally:
            self.overlay.hide()
            self._toolbar_set_busy(False)

    def _on_unify_progress(self, processed, total):
        """Дать Tk обработать клик «Остановить» во время COM-цикла унификации.

        Текст overlay не обновляем (по согласованному UX — статичное сообщение).
        Цель — только прокачать очередь Tk-событий, чтобы клик по кнопке
        дошёл до on_cancel и engine.cancel_unify установил cancel_event.

        ВАЖНО: self.update() вызываем СИНХРОННО, без self.after(0, ...).
        Callback дёргается из главного потока изнутри блокирующего COM-вызова —
        отложенная задача через after(0) не выполнится, пока COM не вернёт
        управление. Синхронный update прокачивает очередь Tk здесь и сейчас.
        """
        try:
            self.update()
        except tk.TclError:
            pass

    def _on_format_state_changed(self, attr, is_active, invalidated=False):
        def _update():
            tg = self.format_buttons.get(attr)
            if tg:
                tg.set_state("on" if is_active else "off")
            self._refresh_format_all_btn()
            if invalidated:
                self.status_label.config(
                    text="Документ изменился — снимок формата сброшен",
                )
            else:
                self._update_status_text()

        self.after(0, _update)

    def _refresh_doc_dependent_options(self):
        """Заблокировать/разблокировать опции по типу выбранного документа.

        Excel не поддерживает skip_tables (он сам — таблица) и Track Changes.
        """
        doc = self.engine.selected_doc
        is_excel = doc is not None and doc.get("type") == "excel"
        # Пропуск таблиц
        if is_excel:
            self.skip_tables_toggle.set_state("disabled")
        else:
            self.skip_tables_toggle.set_state(
                "on" if self.skip_tables_var.get() else "off"
            )
        # Word ревизии
        if is_excel:
            self.revisions_toggle.set_state("disabled")
        else:
            self.revisions_toggle.set_state(
                "on" if self.revisions_mode_var.get() else "off"
            )

    def _refresh_doc_options_toggles(self):
        """Синхронизировать Аудитор/Таблицы при входе в режим documents."""
        self.auditor_toggle.set_state("on" if self.auditor_format_var.get() else "off")
        self._refresh_doc_dependent_options()

    def _refresh_hide_clean_toggle(self):
        self.hide_clean_toggle.set_state("on" if self.hide_clean_var.get() else "off")

    def _refresh_revisions_toggle(self):
        doc = self.engine.selected_doc or {}
        if doc.get("type") == "excel":
            self.revisions_toggle.set_state("disabled")
        else:
            self.revisions_toggle.set_state(
                "on" if self.revisions_mode_var.get() else "off"
            )

    # ─── Callback'и настроек ────────────────────────────────────────────

    def _on_adapter_changed(self, event=None):
        self.engine.set_default_adapter(self.selected_adapter.get())

    def _on_toggle_strict_protection(self):
        self.engine.set_config("strict_protection", self.strict_protect_var.get())

    def _on_toggle_auditor_format(self):
        new_val = not self.auditor_format_var.get()
        self.auditor_format_var.set(new_val)
        self.engine.set_config("auditor_format", new_val)
        self.auditor_toggle.set_state("on" if new_val else "off")

    def _on_toggle_skip_tables(self):
        doc = self.engine.selected_doc
        if doc and doc.get("type") == "excel":
            return  # disabled
        new_val = not self.skip_tables_var.get()
        self.skip_tables_var.set(new_val)
        self.engine.set_config("skip_tables", new_val)
        self.skip_tables_toggle.set_state("on" if new_val else "off")

    def _on_toggle_hide_clean(self):
        new_val = not self.hide_clean_var.get()
        self.hide_clean_var.set(new_val)
        self.engine.set_config("hide_clean_sentences", new_val)
        self.hide_clean_toggle.set_state("on" if new_val else "off")
        if self.current_view == "sentences":
            self._apply_sentence_filter()

    # ─── Действия пользователя ──────────────────────────────────────────

    def find_documents(self):
        self.overlay.show("Поиск открытых документов…", cancelable=False)
        self.update()  # дать overlay появиться до синхронного COM-вызова
        try:
            self.engine.find_documents()
        finally:
            self.overlay.hide()

    def check_selected_document(self):
        # Для Excel «Весь текст» disabled (см. _refresh_check_buttons), но
        # подстрахуемся на случай вызова через горячую клавишу или race.
        doc = self.engine.selected_doc
        if doc is not None and doc.get("type") == "excel":
            self.status_label.config(
                text="Для Excel доступна только проверка выделенного диапазона",
            )
            return
        self.engine.check_document()

    def check_selected_fragment(self):
        self.engine.check_fragment()

    def show_documents_view(self):
        """Вернуться к списку документов."""
        if self.engine.revisions_mode:
            self.engine.set_word_revisions_mode(False)
            self.revisions_mode_var.set(False)
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

        if self.engine.selected_doc and (
            self.engine.is_checking or self.engine.check_results
        ):
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

        self._refresh_format_unify_panel()

        # Динамический статус-бар по зарегистрированным провайдерам
        counts = {}
        for provider in get_all_providers():
            counts[provider.doc_type] = sum(
                1 for d in documents if d.get("type") == provider.doc_type
            )
        status_parts = [f"{t}: {c}" for t, c in counts.items() if c > 0]
        self.status_label.config(
            text=", ".join(status_parts) if status_parts else "Документы не найдены"
        )

    def _create_document_tile_ui(self, doc):
        """Создать плитку документа."""
        is_selected = self.engine.selected_doc and id(doc) == id(
            self.engine.selected_doc
        )
        is_active = is_selected and (
            self.engine.is_checking or self.engine.check_results
        )
        cached = self.engine.doc_state.load(id(doc))
        cached_errors = cached["check_results"] if cached else None

        tile, status_label = create_document_tile(
            self.docs_frame,
            doc,
            self._on_tile_click,
            is_selected=is_selected,
            cached_errors=cached_errors,
            is_active_check=is_active,
        )
        self.tile_frames[id(doc)] = tile
        if status_label:
            self.doc_status_label = status_label

    def _on_tile_click(self, doc):
        # Анти-дребезг: блокирующий activate_document() длится 300-800мс. Если
        # пользователь кликает повторно, новые клики игнорируем до завершения.
        if self._tile_click_in_progress:
            return
        self._tile_click_in_progress = True

        try:
            if (
                self.engine.selected_doc
                and id(doc) == id(self.engine.selected_doc)
                and (self.engine.is_checking or self.engine.check_results)
            ):
                self._show_cached_sentences_with_overlay()
                return

            self.engine.select_document(doc)
            self._refresh_check_buttons()
            self._refresh_doc_dependent_options()
            self._refresh_format_unify_panel()
            self._highlight_selected_tile_ui()

            # Видимая реакция до COM-вызова: пользователь видит, что клик принят.
            self.status_label.config(text="Переключение…")
            self.update_idletasks()

            app_width = APP_WIDTH
            target_rect = (
                self.monitor["x"],
                self.monitor["y"],
                self.monitor["width"] - app_width,
                self.monitor["height"],
            )
            office_finder.activate_document(doc, target_rect)
            self.after(100, self._raise_window)
            self._update_status_text()

            # Если у документа есть кэшированные результаты — показать предложения
            if self.engine.sentences:
                self._show_cached_sentences_with_overlay()
        finally:
            # Снимаем блокировку после возврата фокуса в наше окно (raise_window
            # планируется через 100мс, плюс запас на завершение СOM-callback'ов).
            self.after(200, self._release_tile_click_lock)

    def _show_cached_sentences_with_overlay(self):
        """Перерисовать плитки из кэша под заглушкой «Перерисовка…».

        Перерисовка большого числа плиток занимает 1-3 сек и визуально лагает.
        Overlay скрывает процесс — пользователь видит уже готовый результат.
        """
        self.overlay.show("Перерисовка интерфейса…", cancelable=False)
        self.update()  # форсировать показ overlay до тяжёлой перерисовки
        try:
            self._show_sentences_from_cache()
        finally:
            self.overlay.hide()

    def _release_tile_click_lock(self):
        self._tile_click_in_progress = False

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
            self.doc_status_label.config(
                text=f"{frame} {checked}/{total}", fg="#007bff"
            )
        else:
            errors = sum(
                1 for r in self.engine.check_results.values() if r["has_error"]
            )
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
            self.docs_frame,
            sentence,
            self._on_sentence_click_ui,
        )
        index = sentence["index"]
        self.sentence_tiles[index] = tile
        self.spinner_labels[index] = status_label

    def _create_checked_sentence_tile_ui(self, sentence, result):
        info = create_checked_sentence_tile(
            self.docs_frame,
            sentence,
            result,
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

        sentence = (
            self.sentences[index] if index < len(self.sentences) else {"index": index}
        )
        info = update_sentence_tile(
            tile,
            index,
            original,
            corrected,
            has_error,
            sentence,
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

        app_width = APP_WIDTH
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
                    1
                    for idx in self.engine.check_results
                    if self._should_hide_tile(idx)
                )
            suffix = f" (скрыто: {hidden})" if hidden else ""
            self.status_label.config(
                text=f"Проверка: {len(self.engine.check_results)}/{len(self.engine.sentences)}{suffix}"
            )
        elif self.engine.check_results:
            errors = sum(
                1 for r in self.engine.check_results.values() if r["has_error"]
            )
            hidden = 0
            if self.hide_clean_var.get():
                hidden = sum(
                    1
                    for idx in self.engine.check_results
                    if self._should_hide_tile(idx)
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

    # ─── Режим ошибок ───────────────────────────────────────────────────

    def _refresh_errors_mode_button(self):
        """Включить/отключить кнопку «Исправления» по наличию правок."""
        has_corrs = len(self.engine.word_corrections) > 0
        self.errors_mode_button.set_enabled(has_corrs)

    def _toggle_errors_view(self):
        """Переключение sentences ↔ errors."""
        if self.current_view == "sentences":
            if not self.engine.word_corrections:
                return
            self._show_errors_view()
        elif self.current_view == "errors":
            self._show_sentences_from_cache()

    def _show_errors_view(self):
        """Отрисовать плитки пословных правок.

        При входе автоматически включается отображение ревизий в Word
        (запись ревизий уже выполняется при каждом Apply).
        """
        self.current_view = "errors"
        self.selected_sentence_index = None

        # Включить отображение ревизий в Word и синхронизировать галочку.
        # Для Excel пропускаем — Track Changes там нет.
        doc = self.engine.selected_doc or {}
        if doc.get("type") != "excel" and not self.engine.revisions_mode:
            self.engine.set_word_revisions_mode(True)
        self.revisions_mode_var.set(self.engine.revisions_mode)
        self._update_button_state()
        self._refresh_doc_dependent_options()

        for widget in self.docs_frame.winfo_children():
            widget.destroy()
        self.canvas.yview_moveto(0)
        self.errors_tiles = []

        corrs = self.engine.word_corrections
        if not corrs:
            placeholder = tk.Label(
                self.docs_frame,
                text="Нет применённых правок",
                fg="#999",
                bg="#f0f0f0",
            )
            placeholder.pack(pady=20)
        else:
            for corr in corrs:
                tile = create_word_correction_tile(
                    self.docs_frame,
                    corr,
                    self._on_word_correction_click,
                )
                self.errors_tiles.append((corr, tile))

        self.status_label.config(text=f"Применено правок: {len(corrs)}")

    def _exit_errors_view(self):
        """Кнопка «Назад» в режиме ошибок — выключить ревизии, вернуться к sentences."""
        if self.engine.revisions_mode:
            self.engine.set_word_revisions_mode(False)
            self.revisions_mode_var.set(False)
        self._show_sentences_from_cache()

    def _on_word_correction_click(self, correction):
        """Клик по плитке правки — выделить диапазон в Word и активировать окно."""
        ok = self.engine.navigate_to_word_correction(correction)
        if not ok:
            self.status_label.config(text="⚠ Не удалось найти правку в документе")
            return

        app_width = APP_WIDTH
        target_rect = (
            self.monitor["x"],
            self.monitor["y"],
            self.monitor["width"] - app_width,
            self.monitor["height"],
        )
        office_finder.activate_document(self.engine.selected_doc, target_rect)

    def _on_toggle_revisions_mode(self):
        doc = self.engine.selected_doc
        if doc and doc.get("type") == "excel":
            return  # disabled
        on = not self.revisions_mode_var.get()
        ok = self.engine.set_word_revisions_mode(on)
        if ok:
            self.revisions_mode_var.set(on)
            self.revisions_toggle.set_state("on" if on else "off")
        else:
            self.status_label.config(text="⚠ Не удалось переключить режим ревизий")

    def _on_word_corrections_changed(self):
        def _update():
            if self.current_view == "errors":
                self._show_errors_view()
            elif self.current_view == "sentences":
                self._refresh_errors_mode_button()

        self.after(0, _update)

    def _on_revisions_mode_changed(self, on):
        def _update():
            self.revisions_mode_var.set(bool(on))
            if self.current_view == "errors":
                self._refresh_revisions_toggle()

        self.after(0, _update)

    # ─── Поднять окно ───────────────────────────────────────────────────

    def _raise_window(self):
        """Поднять окно приложения на передний план."""
        self.lift()
        self.focus_force()
