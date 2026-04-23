"""
EventBus — шина событий для разделения UI и бизнес-логики.

Чистый Python, 0 внешних зависимостей.

Использование:
    bus = EventBus()
    bus.subscribe("check_started", callback)
    bus.emit("check_started", total=20)
    bus.unsubscribe("check_started", callback)
"""

import logging
from collections import defaultdict
from typing import Any, Callable, Dict, List

logger = logging.getLogger("core.events")


class EventBus:
    """Простая шина событий (publish/subscribe).

    Потокобезопасность НЕ гарантируется. Вызов emit() должен происходить
    в том же потоке, где работают подписчики.
    """

    def __init__(self):
        self._subscribers: Dict[str, List[Callable]] = defaultdict(list)

    def subscribe(self, event_name: str, callback: Callable) -> None:
        """Подписать callback на событие. Повторная подписка игнорируется."""
        if callback not in self._subscribers[event_name]:
            self._subscribers[event_name].append(callback)
            logger.debug("Subscribed to '%s': %s", event_name, callback.__qualname__)

    def unsubscribe(self, event_name: str, callback: Callable) -> None:
        """Отписать callback от события."""
        try:
            self._subscribers[event_name].remove(callback)
            logger.debug("Unsubscribed from '%s': %s", event_name, callback.__qualname__)
        except (ValueError, KeyError):
            pass

    def emit(self, event_name: str, **data: Any) -> None:
        """Сгенерировать событие. Исключения в callback'ах логируются и не прерывают остальных."""
        subscribers = self._subscribers.get(event_name, [])
        if not subscribers:
            logger.debug("Event '%s' emitted with no subscribers", event_name)
            return

        logger.debug("Emitting '%s' to %d subscriber(s)", event_name, len(subscribers))
        for callback in subscribers:
            try:
                callback(**data)
            except Exception:
                logger.exception(
                    "Error in subscriber '%s' for event '%s'",
                    callback.__qualname__,
                    event_name,
                )

    def clear(self, event_name: str = None) -> None:
        """Очистить подписки. Если event_name=None — очистить все."""
        if event_name is not None:
            self._subscribers.pop(event_name, None)
        else:
            self._subscribers.clear()

    def has_subscribers(self, event_name: str) -> bool:
        """Проверить, есть ли подписчики на событие."""
        return bool(self._subscribers.get(event_name))
