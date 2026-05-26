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


def extract_sentences(doc, *, progress_callback=None, cancel_event=None):
    """Извлечь предложения из документа. Обратная совместимость."""
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.extract_sentences(
            doc, progress_callback=progress_callback, cancel_event=cancel_event,
        )
    return []


def extract_selected_sentences(doc, *, progress_callback=None, cancel_event=None):
    """Извлечь предложения из выделения. Обратная совместимость."""
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.extract_selected_sentences(
            doc, progress_callback=progress_callback, cancel_event=cancel_event,
        )
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


def replace_sentence_text_with_corrections(
    doc, sentence, new_text,
    old_text=None, all_sentences=None, track_revisions=False,
):
    """Заменить текст предложения с возвратом пословных правок и сдвига позиций."""
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.replace_sentence_text_with_corrections(
            doc, sentence, new_text, old_text, all_sentences, track_revisions,
        )
    return {"ok": False, "corrections": [], "delta": 0, "old_end": 0, "rng_start": 0}


def reject_sentence_revisions(doc, sentence, marker, original_text, all_sentences=None):
    """Откатить ревизии предложения через Revision.Reject() (Word).

    Используется для надёжной отмены применённых исправлений с TrackRevisions.
    Для не-Word провайдеров возвращает reason='unsupported'.
    """
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.reject_sentence_revisions(
            doc, sentence, marker, original_text, all_sentences,
        )
    return {"ok": False, "reason": "no_provider",
            "delta": 0, "old_end": 0, "rng_start": 0}


def navigate_to_range(doc, start, end):
    """Выделить произвольный диапазон в документе."""
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.navigate_to_range(doc, start, end)
    return False


def set_revisions_mode(doc, on):
    """Включить/выключить режим Track Changes в документе."""
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.set_revisions_mode(doc, on)
    return False


def navigate_to_correction(doc, correction):
    """Выделить диапазон конкретной пословной правки (режим ошибок)."""
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.navigate_to_correction(doc, correction)
    return False


def get_last_provider_error(doc):
    """Вернуть последнюю ошибку провайдера документа (или None)."""
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.get_last_error()
    return None


def analyze_format(doc, *, cancel_event=None, progress_callback=None):
    """Проанализировать форматирование документа (преобладающий шрифт/размер/стиль).

    Args:
        cancel_event: threading.Event для прерывания обхода документа.
        progress_callback: callable(processed, total) — точка прокачки Tk-очереди.
    """
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.analyze_format(
            doc, cancel_event=cancel_event, progress_callback=progress_callback,
        )
    return None


def apply_format_uniform(doc, attr, target_value=None, *,
                         cancel_event=None, progress_callback=None):
    """Применить унификацию форматирования по одному атрибуту.

    Args:
        cancel_event: threading.Event для прерывания главного цикла применения.
        progress_callback: callable(processed, total) — дёргается из COM-цикла
            чтобы UI смог обработать клик «Остановить».

    Returns:
        dict-дескриптор snapshot или None.
    """
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.apply_format_uniform(
            doc, attr, target_value,
            cancel_event=cancel_event, progress_callback=progress_callback,
        )
    return None


def restore_format(doc, descriptor):
    """Откатить унификацию по ранее сохранённому дескриптору snapshot."""
    provider = get_provider(doc.get("type", ""))
    if provider:
        return provider.restore_format(doc, descriptor)
    return False
