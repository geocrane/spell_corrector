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
    engine.navigate_to_word_correction(correction)
    engine.set_word_revisions_mode(on)

События (генерирует Engine, подписывается UI):
    documents_found(documents)
    documents_not_found()
    extraction_started()
    extraction_progress(extracted, processed)
    check_started(total)
    sentence_start(index)
    sentence_checked(index, original, corrected, has_error)
    check_complete()
    check_error(error)
    word_corrections_changed()
    revisions_mode_changed(on)
    navigation_blocked(message)
    format_state_changed(attr, is_active, invalidated=False)
"""

import logging
import os
import threading
from typing import Optional

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
        # Плоский список применённых пословных правок (см. core.word_corrections).
        self.word_corrections: list = []
        # Включён ли режим Track Changes в Word (влияет на способ применения правок).
        self.revisions_mode: bool = False
        # Состояние унификации форматирования для текущего документа.
        # is_active=True означает, что унификация применена и доступен откат.
        self.format_state: dict = self._make_default_format_state()
        # Кэш analyze_format для текущего документа — заполняется лениво при
        # первом toggle_unify_*. Инвалидируется при правках текста.
        self.format_stats_cache: Optional[dict] = None
        # Event для прерывания текущего извлечения (см. cancel_extraction).
        self._extract_cancel_event: Optional[threading.Event] = None
        # Event для прерывания идущей унификации форматирования (см. cancel_unify).
        self._unify_cancel_event: Optional[threading.Event] = None

    @staticmethod
    def _make_default_format_state() -> dict:
        return {
            "font_name": {"is_active": False, "descriptor": None},
            "font_size": {"is_active": False, "descriptor": None},
            "style":     {"is_active": False, "descriptor": None},
        }

    # ─── Поиск документов ───────────────────────────────────────────────

    def find_documents(self):
        """Найти открытые документы Word/Outlook. Генерирует documents_found или documents_not_found."""
        self.doc_state.cancel()
        self.doc_state.next_generation()
        self.doc_state.clear()
        self.check_results = {}
        self.sentences = []
        self.word_corrections = []
        self.revisions_mode = False
        self.is_checking = False
        self.selected_doc = None
        self.format_state = self._make_default_format_state()
        self.format_stats_cache = None

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
                word_corrections=self.word_corrections,
                revisions_mode=self.revisions_mode,
                format_state=self.format_state,
            )

        cached = self.doc_state.load(id(doc))
        if cached:
            self.check_results = cached["check_results"].copy()
            self.sentences = cached["sentences"].copy()
            self.word_corrections = list(cached.get("word_corrections", []))
            self.revisions_mode = bool(cached.get("revisions_mode", False))
            cached_fmt = cached.get("format_state") or {}
            self.format_state = self._merge_format_state(cached_fmt)
        else:
            self.check_results = {}
            self.sentences = []
            self.word_corrections = []
            self.revisions_mode = False
            self.format_state = self._make_default_format_state()

        # stats_cache не персистим: он специфичен документу и пересчитывается дёшево.
        self.format_stats_cache = None
        self.is_checking = False
        self.selected_doc = doc

    def _merge_format_state(self, cached: dict) -> dict:
        """Сшить дефолтную структуру с тем, что лежит в кэше."""
        state = self._make_default_format_state()
        for attr, defaults in state.items():
            entry = cached.get(attr)
            if isinstance(entry, dict):
                defaults["is_active"] = bool(entry.get("is_active", False))
                defaults["descriptor"] = entry.get("descriptor")
        return state

    # ─── Проверка документа ─────────────────────────────────────────────

    def _start_check(self, extract_func, no_sentences_msg, selection_required_msg=None):
        """Общая логика запуска проверки.

        Сначала переключает UI (событие extraction_started), затем извлекает
        предложения синхронно. Провайдер вызывает progress_callback каждые
        ~25 ячеек — это эмитит extraction_progress, UI обновляет overlay и
        форсирует перерисовку, чтобы кнопка «Остановить» оставалась отзывчивой.

        Args:
            extract_func: Функция извлечения предложений.
            no_sentences_msg: Сообщение если предложений нет.
            selection_required_msg: Сообщение если выделение отсутствует.
        """
        if not self.selected_doc or self.is_checking:
            return

        self.is_checking = True
        # Сброс пословных правок перед новой проверкой — старые карточки
        # из прошлого прогона уже неактуальны (позиции могут сдвинуться).
        if self.word_corrections:
            self.word_corrections = []
            self.events.emit("word_corrections_changed")
        # Мгновенно переключаем UI
        self.events.emit("extraction_started")

        self._extract_cancel_event = threading.Event()

        def _on_progress(extracted, processed):
            self.events.emit(
                "extraction_progress",
                extracted=extracted, processed=processed,
            )

        sentences = extract_func(
            self.selected_doc,
            progress_callback=_on_progress,
            cancel_event=self._extract_cancel_event,
        )

        if self._extract_cancel_event.is_set():
            self.is_checking = False
            self.events.emit("check_error", error="Извлечение остановлено")
            return

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

    def cancel_extraction(self):
        """Остановить идущее извлечение предложений (если оно идёт)."""
        ev = getattr(self, "_extract_cancel_event", None)
        if ev is not None:
            ev.set()

    def cancel_unify(self):
        """Остановить идущую унификацию форматирования (если она идёт).

        Уже применённые сегменты сохраняются в snapshot — кнопка унификации
        останется в положении «вкл», повторный клик откатит частичные правки
        через restore_snapshot.
        """
        ev = getattr(self, "_unify_cancel_event", None)
        if ev is not None:
            ev.set()

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
        self.word_corrections = []
        self.events.emit("word_corrections_changed")

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

        Откат, как и применение, выполняется с TrackRevisions=True — в Word
        появляется дополнительная ревизия-об-отмене. История изменений
        сохраняется. Поиск конкретной существующей Revision и её Reject не
        реализован — это хрупко и не требуется.

        Args:
            index: Индекс предложения.

        Returns:
            bool: True если успешно.
        """
        return self._apply_correction_to_doc(index, apply=False)

    def _apply_correction_to_doc(self, index, apply):
        """Заменить текст предложения на исправленный или оригинальный.

        Применение и откат ВСЕГДА выполняются с TrackRevisions=True, чтобы в
        Word накапливалась история изменений. Видимость этих ревизий
        управляется отдельно через set_word_revisions_mode (отображение).

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

            replace_result = office_finder.replace_sentence_text_with_corrections(
                self.selected_doc, sentence, new_text,
                old_text=old_text, all_sentences=self.sentences,
                track_revisions=True,
            )
            ok = replace_result["ok"]

            if ok:
                result["state"] = "applied" if apply else "pending"

                old_end = replace_result.get("old_end", 0)
                delta = replace_result.get("delta", 0)

                scope = None
                if self.selected_doc and self.selected_doc.get("type") == "excel":
                    scope = {
                        "excel_workbook": sentence.get("workbook"),
                        "excel_sheet": sentence.get("sheet"),
                        "excel_cell_address": sentence.get("cell_address"),
                    }

                if apply:
                    self._shift_word_corrections(old_end, delta, scope=scope)
                    new_corrs = replace_result.get("corrections", []) or []
                    for c in new_corrs:
                        c["sentence_index"] = index
                    self.word_corrections.extend(new_corrs)
                else:
                    self.word_corrections = [
                        c for c in self.word_corrections
                        if c.get("sentence_index") != index
                    ]
                    self._shift_word_corrections(old_end, delta, scope=scope)

                self.events.emit("word_corrections_changed")
                # Любое изменение текста инвалидирует все формат-снимки —
                # позиции сегментов в snapshot.ranges больше не валидны.
                self._invalidate_format_snapshots()

            return ok
        except Exception as e:
            logger.error("Ошибка применения/отката #%d: %s", index, e)
            return False

    def _shift_word_corrections(self, old_end: int, delta: int, scope: dict | None = None) -> None:
        """Сдвинуть позиции уже сохранённых пословных правок после замены текста.

        Аналогично _after_replacement, но для записей в self.word_corrections.
        Если правка начиналась после old_end — сдвинуть её на delta.

        Args:
            old_end: Позиция конца изменённого фрагмента (до замены).
            delta:   Разница длин (len(new) - len(old)).
            scope:   Если задан — сдвигаются только правки, у которых все
                     ключи scope совпадают с правкой (нужно для Excel:
                     позиции — внутри ячейки, сдвиг только в той же ячейке).
        """
        if delta == 0:
            return
        for c in self.word_corrections:
            if scope is not None and not all(c.get(k) == v for k, v in scope.items()):
                continue
            start = c.get("word_range_start")
            if start is not None and start >= old_end:
                c["word_range_start"] = start + delta
                c["word_range_end"] = c.get("word_range_end", start) + delta
                if "excel_char_start" in c:
                    c["excel_char_start"] = c["word_range_start"]
                    c["excel_char_end"] = c["word_range_end"]

    def toggle_skip(self, index):
        """Переключить состояние пропуска предложения.

        Если предложение было в состоянии "applied" — карточки его пословных правок
        удаляются из self.word_corrections (текст в документе при этом НЕ откатывается,
        это ответственность вызывающего кода — обычно UI вызывает revert до toggle_skip,
        либо переход "applied → skipped" не предусмотрен в UI).

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
            new_state = "skipped"
        else:
            result["state"] = "pending"
            new_state = "pending"

        # При любом переходе убираем карточки этого предложения из режима ошибок:
        # карточки показывают только state="applied", чего после toggle_skip заведомо нет.
        before = len(self.word_corrections)
        self.word_corrections = [
            c for c in self.word_corrections
            if c.get("sentence_index") != index
        ]
        if len(self.word_corrections) != before:
            self.events.emit("word_corrections_changed")

        return new_state

    # ─── Навигация ──────────────────────────────────────────────────────

    def navigate_to_sentence(self, sentence):
        """Перейти к предложению в документе и выделить его.

        Args:
            sentence: Словарь предложения.

        Returns:
            bool: True если выделение успешно.
        """
        if not self.selected_doc:
            return False
        ok = office_finder.navigate_to_sentence(self.selected_doc, sentence)
        if not ok:
            err = office_finder.get_last_provider_error(self.selected_doc)
            if err:
                self.events.emit("navigation_blocked", message=err)
        return ok

    def navigate_to_word_correction(self, correction: dict) -> bool:
        """Выделить диапазон конкретной пословной правки в документе.

        Args:
            correction: WordCorrection-словарь.

        Returns:
            bool: True если выделение успешно.
        """
        if not self.selected_doc:
            return False
        ok = office_finder.navigate_to_correction(self.selected_doc, correction)
        if not ok:
            err = office_finder.get_last_provider_error(self.selected_doc)
            if err:
                self.events.emit("navigation_blocked", message=err)
        return ok

    def set_word_revisions_mode(self, on: bool) -> bool:
        """Включить/выключить ОТОБРАЖЕНИЕ ревизий в Word.

        Запись ревизий (TrackRevisions) выполняется автоматически при каждом
        apply_correction/revert_correction. Здесь управляется только видимость
        накопленных ревизий: при on=True Word показывает All Markup, при
        on=False — Final (как обычный текст).

        Args:
            on: True — включить отображение, False — выключить.

        Returns:
            bool: True если успешно.
        """
        if not self.selected_doc:
            return False
        ok = office_finder.set_revisions_mode(self.selected_doc, on)
        if ok:
            self.revisions_mode = bool(on)
            self.events.emit("revisions_mode_changed", on=self.revisions_mode)
        return ok

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

    # ─── Унификация форматирования ──────────────────────────────────────

    def toggle_unify_font(self) -> None:
        """Toggle унификации шрифта (font_name)."""
        self._toggle_unify("font_name")

    def toggle_unify_size(self) -> None:
        """Toggle унификации размера (font_size)."""
        self._toggle_unify("font_size")

    def toggle_unify_style(self) -> None:
        """Toggle унификации стиля (bold/italic)."""
        self._toggle_unify("style")

    def toggle_unify_all(self) -> None:
        """«Применить всё» / «Откатить всё».

        Если хоть одна унификация активна — откатываем все три в обратном
        порядке. Иначе — применяем все три подряд (font → size → style).
        Между атрибутами проверяем флаг отмены, чтобы клик «Остановить»
        прервал не только текущий attr, но и весь оставшийся каскад.
        """
        if not self.selected_doc:
            return
        any_active = any(st["is_active"] for st in self.format_state.values())
        if any_active:
            for attr in ("style", "font_size", "font_name"):
                if self.format_state[attr]["is_active"]:
                    self._toggle_unify(attr)
        else:
            for attr in ("font_name", "font_size", "style"):
                ev = self._unify_cancel_event
                if ev is not None and ev.is_set():
                    break
                if not self.format_state[attr]["is_active"]:
                    self._toggle_unify(attr)

    def _toggle_unify(self, attr: str) -> None:
        if not self.selected_doc:
            return
        if attr not in self.format_state:
            return

        st = self.format_state[attr]
        if st["is_active"] and st["descriptor"]:
            ok = office_finder.restore_format(self.selected_doc, st["descriptor"])
            st["is_active"] = False
            st["descriptor"] = None
            # Невалидный snapshot (документ менялся) — это не ошибка для UX:
            # просто сообщим UI, что переход в «выкл» не сопровождался откатом.
            self.events.emit(
                "format_state_changed",
                attr=attr, is_active=False, invalidated=(not ok),
            )
            return

        # Применить унификацию. Cancel-event и progress callback создаём
        # ДО analyze_format — он тоже долгий и должен прерываться по «Остановить».
        self._unify_cancel_event = threading.Event()

        def _on_unify_progress(processed, total):
            self.events.emit("unify_progress", processed=processed, total=total)

        if self.format_stats_cache is None:
            try:
                stats = office_finder.analyze_format(
                    self.selected_doc,
                    cancel_event=self._unify_cancel_event,
                    progress_callback=_on_unify_progress,
                )
            except Exception as e:
                logger.error("toggle_unify: analyze_format error: %s", e)
                stats = None
            if self._unify_cancel_event.is_set():
                # analyze был прерван — частичные счётчики не кэшируем,
                # состояние кнопки не меняем (остаётся выкл).
                self.events.emit("format_state_changed", attr=attr, is_active=False)
                return
            self.format_stats_cache = stats
        if not self.format_stats_cache:
            self.events.emit("format_state_changed", attr=attr, is_active=False)
            return

        target_key = "dominant_style" if attr == "style" else f"dominant_{attr}"
        target = self.format_stats_cache.get(target_key)
        if target is None:
            self.events.emit("format_state_changed", attr=attr, is_active=False)
            return

        try:
            descriptor = office_finder.apply_format_uniform(
                self.selected_doc, attr, target,
                cancel_event=self._unify_cancel_event,
                progress_callback=_on_unify_progress,
            )
        except Exception as e:
            logger.error("toggle_unify: apply_format_uniform error: %s", e)
            descriptor = None

        if descriptor:
            st["is_active"] = True
            st["descriptor"] = descriptor
        self.events.emit("format_state_changed", attr=attr, is_active=st["is_active"])

    def _invalidate_format_snapshots(self) -> None:
        """Сбросить все формат-снимки (после правки текста документа).

        Снапшоты хранят (start, end)-позиции, которые сместились — попытка
        восстановить такой snapshot повредила бы текст. Кнопки в UI переходят
        в «выкл» с пометкой invalidated.
        """
        self.format_stats_cache = None
        for attr, st in self.format_state.items():
            if st["is_active"] or st["descriptor"]:
                st["is_active"] = False
                st["descriptor"] = None
                self.events.emit(
                    "format_state_changed",
                    attr=attr, is_active=False, invalidated=True,
                )

    # ─── Отмена ─────────────────────────────────────────────────────────

    def cancel(self):
        """Отменить текущую проверку."""
        self.doc_state.cancel()
        self.is_checking = False
