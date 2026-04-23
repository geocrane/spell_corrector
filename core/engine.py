"""
Engine — бизнес-логика приложения.

НЕ импортирует tkinter. Вся коммуникация с UI через EventBus.

Команды (вызывает UI):
    engine.find_documents()
    engine.select_document(doc)
    engine.check_document()
    engine.check_fragment()
    engine.apply_correction(index)
    engine.revert_correction(index)
    engine.toggle_skip(index)
    engine.set_config(key, value)
    engine.cancel()
    engine.navigate_to_sentence(sentence)

События (генерирует Engine, подписывается UI):
    documents_found(documents)
    documents_not_found()
    check_started(total)
    sentence_start(index)
    sentence_checked(index, original, corrected, has_error)
    check_complete()
    check_error(error)
"""

import logging
import os
import threading

import office_finder
# spell_checker импортируется лениво в _run_check() и get_available_adapters()
from core.config import load_config, save_config
from core.doc_state import DocStateCache
from core.events import EventBus

logger = logging.getLogger("core.engine")

_THIS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class _CheckSession:
    """Хранит состояние одной проверки, заменяет сотни lambda-замыканий."""

    def __init__(self, engine, generation):
        self.engine = engine
        self.generation = generation

    def on_start(self, index):
        if self.generation == self.engine.doc_state.generation:
            self.engine.events.emit("sentence_start", index=index)

    def on_progress(self, index, original, corrected, has_error):
        if self.generation == self.engine.doc_state.generation:
            self.engine.check_results[index] = {
                "original": original,
                "corrected": corrected,
                "has_error": has_error,
                "state": "pending",
            }
            self.engine.events.emit(
                "sentence_checked",
                index=index, original=original, corrected=corrected, has_error=has_error,
            )

    def on_complete(self):
        if self.generation == self.engine.doc_state.generation:
            self.engine.is_checking = False
            self.engine.events.emit("check_complete")

    def on_error(self, error):
        if self.generation == self.engine.doc_state.generation:
            self.engine.is_checking = False
            logger.error("Check error: %s", error)
            self.engine.events.emit("check_error", error=error)


class Engine:
    """Бизнес-логика приложения. Общается с UI исключительно через EventBus."""

    def __init__(self):
        self.events = EventBus()
        self.config = load_config()
        self.doc_state = DocStateCache()
        self.documents = []
        self.selected_doc = None
        self.is_checking = False
        self.sentences = []
        self.check_results = {}

    # ─── Поиск документов ───────────────────────────────────────────────

    def find_documents(self):
        """Найти открытые документы Word/Outlook. Генерирует documents_found или documents_not_found."""
        self.doc_state.cancel()
        self.doc_state.next_generation()
        self.doc_state.clear()
        self.check_results = {}
        self.sentences = []
        self.is_checking = False
        self.selected_doc = None

        self.documents = office_finder.find_all_documents()

        if not self.documents:
            self.events.emit("documents_not_found")
        else:
            self.events.emit("documents_found", documents=self.documents)

    # ─── Выбор документа ────────────────────────────────────────────────

    def select_document(self, doc):
        """Выбрать документ. Сохраняет состояние предыдущего, загружает состояние нового.

        Args:
            doc: Словарь документа.
        """
        if self.selected_doc is not None:
            self.doc_state.save(
                id(self.selected_doc), self.check_results, self.sentences,
            )

        cached = self.doc_state.load(id(doc))
        if cached:
            self.check_results = cached["check_results"].copy()
            self.sentences = cached["sentences"].copy()
        else:
            self.check_results = {}
            self.sentences = []

        self.is_checking = False
        self.selected_doc = doc

    # ─── Проверка документа ─────────────────────────────────────────────

    def _start_check(self, extract_func, no_sentences_msg, selection_required_msg=None):
        """Общая логика запуска проверки.

        Сначала переключает UI (событие extraction_started), затем извлекает
        предложения синхронно. Пользователь видит пустой экран со статусом
        «Извлечение...» — UI может зависнуть на 0.5-2 сек (COM-операция),
        но пользователю нечего делать в этот момент.

        Args:
            extract_func: Функция извлечения предложений.
            no_sentences_msg: Сообщение если предложений нет.
            selection_required_msg: Сообщение если выделение отсутствует.
        """
        if not self.selected_doc or self.is_checking:
            return

        self.is_checking = True
        # Мгновенно переключаем UI
        self.events.emit("extraction_started")

        sentences = extract_func(self.selected_doc)

        if sentences is None:
            self.is_checking = False
            if selection_required_msg:
                self.events.emit("check_error", error=selection_required_msg)
            return

        if not sentences:
            self.is_checking = False
            self.events.emit("check_error", error=no_sentences_msg)
            return

        self._run_check(sentences)

    def check_document(self):
        """Запустить проверку всего документа."""
        self._start_check(office_finder.extract_sentences, "Предложения не найдены")

    def check_fragment(self):
        """Запустить проверку выделенного фрагмента."""
        self._start_check(
            office_finder.extract_selected_sentences,
            "Предложения не найдены в выделении",
            "Выделите фрагмент текста в документе",
        )

    def _run_check(self, sentences):
        """Запустить асинхронную проверку предложений.

        Args:
            sentences: Список предложений.
        """
        for i, s in enumerate(sentences):
            s["index"] = i

        if self.config.get("skip_tables", False):
            sentences = [s for s in sentences if not s.get("in_table", False)]
            for i, s in enumerate(sentences):
                s["index"] = i

        self.sentences = sentences
        self.doc_state.cancel()
        self.check_results = {}

        gen = self.doc_state.next_generation()
        total = len(sentences)

        self.events.emit("check_started", total=total)

        import spell_checker  # ленивый импорт

        session = _CheckSession(self, gen)
        self.doc_state.cancel_event = spell_checker.check_sentences_async(
            sentences,
            on_start=session.on_start,
            on_progress=session.on_progress,
            on_complete=session.on_complete,
            on_error=session.on_error,
            adapter_name=self.config.get("default_adapter", ""),
            strict=self.config.get("strict_protection", False),
            auditor_format=self.config.get("auditor_format", False),
            word_blocklist=self.config.get("word_blocklist", []),
        )

    # ─── Применение/откат исправлений ───────────────────────────────────

    def apply_correction(self, index):
        """Применить исправление к документу.

        Args:
            index: Индекс предложения.

        Returns:
            bool: True если успешно.
        """
        return self._apply_correction_to_doc(index, apply=True)

    def revert_correction(self, index):
        """Откатить исправление.

        Args:
            index: Индекс предложения.

        Returns:
            bool: True если успешно.
        """
        return self._apply_correction_to_doc(index, apply=False)

    def _apply_correction_to_doc(self, index, apply):
        """Заменить текст предложения на исправленный или оригинальный.

        Args:
            index: Индекс предложения.
            apply: True — применить исправление, False — откатить.

        Returns:
            bool: True если замена успешна.
        """
        try:
            result = self.check_results.get(index)
            if not result:
                return False

            sentence = self.sentences[index]
            original = result["original"]
            corrected = result["corrected"]

            if apply:
                new_text, old_text = corrected, original
            else:
                new_text, old_text = original, corrected

            ok = office_finder.replace_sentence_text(
                self.selected_doc, sentence, new_text,
                old_text=old_text, all_sentences=self.sentences,
            )

            if ok:
                result["state"] = "applied" if apply else "pending"

            return ok
        except Exception as e:
            logger.error("Ошибка применения/отката #%d: %s", index, e)
            return False

    def toggle_skip(self, index):
        """Переключить состояние пропуска предложения.

        Args:
            index: Индекс предложения.

        Returns:
            str: Новое состояние ("skipped" или "pending").
        """
        result = self.check_results.get(index)
        if not result:
            return "pending"

        current = result.get("state", "pending")
        if current == "pending":
            result["state"] = "skipped"
            return "skipped"
        else:
            result["state"] = "pending"
            return "pending"

    # ─── Навигация ──────────────────────────────────────────────────────

    def navigate_to_sentence(self, sentence):
        """Перейти к предложению в документе и выделить его.

        Args:
            sentence: Словарь предложения.

        Returns:
            bool: True если выделение успешно.
        """
        if self.selected_doc:
            return office_finder.navigate_to_sentence(self.selected_doc, sentence)
        return False

    # ─── Конфигурация ───────────────────────────────────────────────────

    def set_config(self, key, value):
        """Сохранить настройку в конфигурацию и config.json."""
        self.config[key] = value
        save_config(self.config)

    def get_config(self):
        """Вернуть текущую конфигурацию."""
        return self.config

    def get_available_adapters(self):
        """Вернуть список доступных адаптеров."""
        import spell_checker  # ленивый импорт
        return spell_checker.discover_adapters()

    def get_default_adapter(self):
        """Вернуть адаптер по умолчанию из конфигурации."""
        return self.config.get("default_adapter", "")

    def set_default_adapter(self, name):
        """Установить адаптер по умолчанию."""
        self.config["default_adapter"] = name
        save_config(self.config)

    # ─── Отмена ─────────────────────────────────────────────────────────

    def cancel(self):
        """Отменить текущую проверку."""
        self.doc_state.cancel()
        self.is_checking = False
