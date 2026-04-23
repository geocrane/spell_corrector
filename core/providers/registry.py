"""
Реестр провайдеров документов.

Позволяет регистрировать новые типы документов и получать провайдер по типу.
"""

from typing import Optional

from core.providers.base import DocumentProvider

_registry: dict[str, DocumentProvider] = {}


def register_provider(provider: DocumentProvider) -> None:
    """Зарегистрировать провайдер документа.

    Args:
        provider: Экземпляр DocumentProvider.
    """
    _registry[provider.doc_type] = provider


def get_provider(doc_type: str) -> Optional[DocumentProvider]:
    """Получить провайдер по типу документа.

    Args:
        doc_type: Строка типа ("word", "outlook", "excel", ...).

    Returns:
        DocumentProvider или None если не зарегистрирован.
    """
    return _registry.get(doc_type)


def get_all_providers() -> list[DocumentProvider]:
    """Вернуть все зарегистрированные провайдеры.

    Returns:
        List[DocumentProvider]: Все провайдеры в порядке регистрации.
    """
    return list(_registry.values())


def find_all_documents() -> list[dict]:
    """Найти все открытые документы всех зарегистрированных типов.

    Returns:
        List[dict]: Объединённый список документов.
    """
    docs = []
    for provider in _registry.values():
        docs.extend(provider.find_documents())
    return docs
