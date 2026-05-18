"""
Базовый абстрактный класс провайдера документа.

Каждый тип документа (Word, Outlook, Excel, ...) реализует этот интерфейс.
"""

from abc import ABC, abstractmethod
from typing import Any, Optional


class DocumentProvider(ABC):
    """Интерфейс провайдера для одного типа документа.

    Провайдер отвечает за:
    - Поиск открытых документов своего типа
    - Получение COM-объекта для работы с содержимым
    - Активацию и позиционирование окна
    - Извлечение текста (предложений/ячеек)
    - Навигацию к фрагменту и выделение
    - Замену текста с сохранением форматирования
    """

    @property
    @abstractmethod
    def doc_type(self) -> str:
        """Уникальный идентификатор типа: 'word', 'outlook', 'excel', ..."""
        ...

    @abstractmethod
    def find_documents(self) -> list[dict]:
        """Найти все открытые документы этого типа.

        Returns:
            List[dict]: Список документов. Каждый dict содержит минимум:
                - "name": str — имя документа
                - "type": str — тип (должен совпадать с self.doc_type)
                - Дополнительные ключи, специфичные для провайдера
        """
        ...

    @abstractmethod
    def get_doc_com(self, doc: dict) -> Any:
        """Вернуть COM-объект для работы с содержимым документа.

        Для Word — это doc["com_object"] (Word.Document).
        Для Outlook — это doc["inspector"].WordEditor.
        Для Excel — это doc["worksheet"] (Excel.Worksheet).

        Args:
            doc: Словарь документа (из find_documents).

        Returns:
            COM-объект или None.
        """
        ...

    @abstractmethod
    def activate(self, doc: dict, target_rect: Optional[tuple]) -> None:
        """Активировать окно документа и разместить его в указанной области.

        Args:
            doc: Словарь документа.
            target_rect: (x, y, width, height) или None.
        """
        ...

    @abstractmethod
    def extract_sentences(
        self,
        doc: dict,
        *,
        progress_callback=None,
        cancel_event=None,
    ) -> list[dict]:
        """Извлечь предложения/фрагменты из документа.

        Args:
            doc: Словарь документа.
            progress_callback: Опциональный колбэк (extracted: int, processed: int | None)
                — провайдер вызывает периодически. Используется для индикатора прогресса.
            cancel_event: threading.Event. Если установлен — провайдер должен прервать
                обход и вернуть уже накопленные результаты (или пустой список).

        Returns:
            List[dict]: Список фрагментов с ключами:
                - "index": int
                - "text": str
                - "range_start": int (позиция для навигации)
                - "range_end": int
                - "in_table": bool (опционально)
                - Дополнительные ключи, специфичные для провайдера
        """
        ...

    @abstractmethod
    def extract_selected_sentences(
        self,
        doc: dict,
        *,
        progress_callback=None,
        cancel_event=None,
    ) -> Optional[list[dict]]:
        """Извлечь предложения из выделенного фрагмента.

        Args:
            doc: Словарь документа.
            progress_callback: См. extract_sentences.
            cancel_event: См. extract_sentences.

        Returns:
            List[dict] или None если нет выделения.
        """
        ...

    @abstractmethod
    def navigate_to_sentence(self, doc: dict, sentence: dict) -> bool:
        """Перейти к фрагменту и выделить его в документе.

        Args:
            doc: Словарь документа.
            sentence: Словарь фрагмента с range_start, range_end, text.

        Returns:
            bool: True если выделение успешно, False если не найдено.
        """
        ...

    @abstractmethod
    def replace_sentence_text(
        self,
        doc: dict,
        sentence: dict,
        new_text: str,
        old_text: Optional[str] = None,
        all_sentences: Optional[list] = None,
    ) -> bool:
        """Заменить текст фрагмента в документе.

        Args:
            doc: Словарь документа.
            sentence: Словарь фрагмента.
            new_text: Новый текст.
            old_text: Исходный текст для diff (опционально).
            all_sentences: Полный список для сдвига позиций.

        Returns:
            bool: True если успешно.
        """
        ...

    def navigate_to_range(self, doc: dict, start: int, end: int) -> bool:
        """Выделить произвольный диапазон позиций в документе.

        По умолчанию не поддерживается. Используется UI режима ошибок для
        навигации к конкретной пословной правке.
        """
        return False

    def set_revisions_mode(self, doc: dict, on: bool) -> bool:
        """Включить/выключить режим Track Changes в документе.

        По умолчанию не поддерживается (актуально для Word).
        """
        return False

    def replace_sentence_text_with_corrections(
        self,
        doc: dict,
        sentence: dict,
        new_text: str,
        old_text: Optional[str] = None,
        all_sentences: Optional[list] = None,
        track_revisions: bool = False,
    ) -> dict:
        """Заменить текст фрагмента и вернуть пословные правки + сдвиг.

        По умолчанию делегирует в replace_sentence_text без granular-corrections.
        Провайдеры, поддерживающие пословный diff (Word), переопределяют.

        Returns:
            dict с ключами:
                ok (bool), corrections (list[dict]),
                delta (int), old_end (int), rng_start (int).
        """
        ok = self.replace_sentence_text(doc, sentence, new_text, old_text, all_sentences)
        return {"ok": ok, "corrections": [], "delta": 0, "old_end": 0, "rng_start": 0}

    def navigate_to_correction(self, doc: dict, correction: dict) -> bool:
        """Выделить диапазон конкретной пословной правки.

        Default-реализация читает word_range_start/end и вызывает
        navigate_to_range. Провайдеры со своей идентификацией (Excel:
        ячейка) переопределяют.
        """
        start = correction.get("word_range_start")
        end = correction.get("word_range_end")
        if start is None or end is None:
            return False
        return self.navigate_to_range(doc, start, end)

    def get_last_error(self) -> Optional[str]:
        """Вернуть текст последней ошибки (если провайдер её сохраняет).

        Используется UI для показа уведомлений типа «Excel заблокирован».
        По умолчанию возвращает None.
        """
        return None

    def get_icon(self) -> tuple[str, str]:
        """Вернуть иконку и цвет для отображения в UI.

        Returns:
            tuple: (символ, цвет) — например ("W", "#2B579A").
        """
        return ("?", "#888888")

    # ─── Унификация форматирования (опционально) ────────────────────────

    def analyze_format(
        self, doc: dict, *, cancel_event=None, progress_callback=None,
    ) -> Optional[dict]:
        """Проанализировать форматирование и вернуть преобладающие значения.

        Args:
            cancel_event: threading.Event. Если .set() — провайдер должен
                прервать обход документа. Накопленные счётчики возвращаются
                как есть; caller сам решит, использовать частичный результат
                или отбросить (по проверке cancel_event.is_set() после вызова).
            progress_callback: callable(processed, total). Дёргается из цикла
                анализа на каждом шаге, чтобы UI смог обработать клик
                «Остановить» во время синхронного COM-вызова.

        Returns:
            dict с ключами dominant_font_name, dominant_font_size,
            dominant_style, coverage, total_chars — или None если провайдер
            не поддерживает анализ форматирования.
        """
        return None

    def apply_format_uniform(
        self,
        doc: dict,
        attr: str,
        target_value: Any = None,
        *,
        cancel_event=None,
        progress_callback=None,
    ) -> Optional[dict]:
        """Применить преобладающее значение атрибута ко всему основному тексту.

        Args:
            doc: Словарь документа.
            attr: "font_name" | "font_size" | "style".
            target_value: Конкретное значение или None (тогда провайдер сам
                          считает преобладающее через analyze_format).
            cancel_event: threading.Event. Если .set() — провайдер должен
                прервать главный цикл применения на ближайшей точке проверки.
                Уже совершённые правки остаются в snapshot и откатываются
                обычным restore_format.
            progress_callback: callable(processed, total). Дёргается из
                цикла применения каждые ~N единиц, чтобы UI смог обработать
                клик «Остановить» во время синхронного COM-вызова.

        Returns:
            Лёгкий дескриптор snapshot для последующего restore_format, либо
            None если операция не поддерживается / не удалась. Дескриптор
            должен содержать как минимум {"snap_id": str, "attr": str}.
        """
        return None

    def restore_format(self, doc: dict, snapshot_descriptor: dict) -> bool:
        """Восстановить исходный формат по дескриптору snapshot.

        Returns:
            True если откат успешен; False если snapshot не найден или
            документ был отредактирован между apply и restore.
        """
        return False
