"""
Кэш состояний документов.

Хранит результаты проверки и предложения для каждого документа,
чтобы быстро восстанавливать состояние при переключении.
"""

import threading
from typing import Any, Dict, Optional


class DocStateCache:
    """Кэш состояний документов с generation counter и cancel event.

    Атрибуты:
        generation: Счётчик генерации для защиты от устаревших callback'ов.
        cancel_event: Event для отмены текущего worker'а.
    """

    def __init__(self):
        self._cache: Dict[int, Dict[str, Any]] = {}
        self.generation: int = 0
        self.cancel_event: Optional[threading.Event] = None

    def save(self, doc_id: int, check_results: dict, sentences: list) -> None:
        """Сохранить состояние проверки документа в кэш."""
        self._cache[doc_id] = {
            "check_results": check_results.copy(),
            "sentences": sentences.copy(),
        }

    def load(self, doc_id: int) -> Optional[Dict[str, Any]]:
        """Загрузить состояние проверки документа из кэша.

        Returns:
            dict с ключами "check_results" и "sentences", или None.
        """
        return self._cache.get(doc_id)

    def clear(self) -> None:
        """Очистить весь кэш."""
        self._cache.clear()

    def next_generation(self) -> int:
        """Инкрементировать и вернуть новый generation counter."""
        self.generation += 1
        return self.generation

    def cancel(self) -> None:
        """Отменить текущий worker проверки."""
        if self.cancel_event is not None:
            self.cancel_event.set()
            self.cancel_event = None

    def has(self, doc_id: int) -> bool:
        """Проверить, есть ли состояние для документа в кэше."""
        return doc_id in self._cache
