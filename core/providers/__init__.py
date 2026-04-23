"""
Пакет провайдеров документов.

Регистрация провайдеров отложена до первого вызова — не замедляет импорт.
"""

from core.providers.base import DocumentProvider
from core.providers.registry import (
    register_provider,
    get_provider,
    get_all_providers,
    find_all_documents,
)

# Ленивая регистрация — флаг и lock
_registered = False

__all__ = [
    "DocumentProvider",
    "register_provider",
    "get_provider",
    "get_all_providers",
    "find_all_documents",
    "WordProvider",
    "OutlookProvider",
    "ExcelProvider",
]


def _ensure_registered():
    """Зарегистрировать провайдеры при первом вызове (thread-safe)."""
    global _registered
    if _registered:
        return
    # Импорты только когда реально нужны
    from core.providers.word_provider import WordProvider
    from core.providers.outlook_provider import OutlookProvider
    from core.providers.excel_provider import ExcelProvider

    register_provider(WordProvider())
    register_provider(OutlookProvider())
    register_provider(ExcelProvider())
    _registered = True


# Переопределяем функции из registry, чтобы они вызывали _ensure_registered
_original_get_provider = get_provider
_original_get_all_providers = get_all_providers
_original_find_all_documents = find_all_documents


def get_provider(doc_type):
    """Получить провайдер по типу документа (с ленивой регистрацией)."""
    _ensure_registered()
    return _original_get_provider(doc_type)


def get_all_providers():
    """Вернуть все зарегистрированные провайдеры (с ленивой регистрацией)."""
    _ensure_registered()
    return _original_get_all_providers()


def find_all_documents():
    """Найти все открытые документы всех зарегистрированных типов (с ленивой регистрацией)."""
    _ensure_registered()
    return _original_find_all_documents()
