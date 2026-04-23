"""
Фасад для обратной совместимости.

Все функции делегируют в зарегистрированные провайдеры из core.providers.
Новый код должен импортировать напрямую из core.providers.
"""

from typing import Optional

from core.providers import (
    get_provider,
    find_all_documents as _providers_find_all,
)

# Re-export для совместимости
find_all_documents = _providers_find_all


def find_word_documents():
    """Найти открытые документы Word. Обратная совместимость."""
    provider = get_provider("word")
    return provider.find_documents() if provider else []


def find_outlook_emails():
    """Найти открытые письма Outlook. Обратная совместимость."""
    provider = get_provider("outlook")
    return provider.find_documents() if provider else []


def activate_document(doc, target_rect=None):
    """Активировать документ. Обратная совместимость."""
    provider = get_provider(doc.get("type", ""))
    if provider:
        provider.activate(doc, target_rect)


def extract_sentences(doc):
    """Извлечь предложения из документа. Обратная совместимость."""
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.extract_sentences(doc)
    return []


def extract_selected_sentences(doc):
    """Извлечь предложения из выделения. Обратная совместимость."""
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.extract_selected_sentences(doc)
    return None


def navigate_to_sentence(doc, sentence):
    """Перейти к предложению. Обратная совместимость.

    Returns:
        bool: True если выделение успешно.
    """
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.navigate_to_sentence(doc, sentence)
    return False


def replace_sentence_text(doc, sentence, new_text, old_text=None, all_sentences=None):
    """Заменить текст предложения. Обратная совместимость."""
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.replace_sentence_text(doc, sentence, new_text, old_text, all_sentences)
    return False
