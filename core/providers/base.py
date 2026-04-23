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
    def extract_sentences(self, doc: dict) -> list[dict]:
        """Извлечь предложения/фрагменты из документа.

        Args:
            doc: Словарь документа.

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
    def extract_selected_sentences(self, doc: dict) -> Optional[list[dict]]:
        """Извлечь предложения из выделенного фрагмента.

        Args:
            doc: Словарь документа.

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

    def get_icon(self) -> tuple[str, str]:
        """Вернуть иконку и цвет для отображения в UI.

        Returns:
            tuple: (символ, цвет) — например ("W", "#2B579A").
        """
        return ("?", "#888888")
