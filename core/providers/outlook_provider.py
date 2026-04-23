"""
Провайдер для Microsoft Outlook.

Outlook использует WordEditor (тот же Word API), поэтому делегирует
большую часть логики WordProvider.
"""

import logging
from typing import Any, Optional

import win32com.client
import win32gui

from core.providers.base import DocumentProvider
from core.providers.word_provider import (
    WordProvider,
    _extract_sentences_from_doc,
    _find_sentence_range,
)

logger = logging.getLogger("core.providers.outlook")


class OutlookProvider(DocumentProvider):
    """Провайдер для Microsoft Outlook (письма в отдельных окнах).

    Outlook использует WordEditor — тот же движок что и Word,
    поэтому извлечение/замена текста делегируется WordProvider.
    """

    def __init__(self):
        self._word = WordProvider()

    @property
    def doc_type(self) -> str:
        return "outlook"

    def find_documents(self) -> list[dict]:
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            emails = []
            for inspector in outlook.Inspectors:
                try:
                    item = inspector.CurrentItem
                    if hasattr(item, 'Subject'):
                        emails.append({
                            "name": item.Subject or "(Без темы)",
                            "inspector": inspector,
                            "type": "outlook",
                            "mail_item": item,
                        })
                except Exception:
                    continue
            return emails
        except Exception:
            return []

    def get_doc_com(self, doc: dict) -> Any:
        try:
            return doc["inspector"].WordEditor
        except Exception:
            return None

    def activate(self, doc: dict, target_rect: Optional[tuple]) -> None:
        try:
            inspector = doc["inspector"]
            if target_rect:
                x, y, width, height = target_rect
                inspector.Left = x
                inspector.Top = y
                inspector.Width = width
                inspector.Height = height
            inspector.Activate()
        except Exception:
            pass

    def extract_sentences(self, doc: dict) -> list[dict]:
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return []
        # WordEditor — тот же Word API, но без _find_body_start
        sentences = _extract_sentences_from_doc(doc_com)
        return sentences

    def extract_selected_sentences(self, doc: dict) -> Optional[list[dict]]:
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return None

        try:
            sel = doc_com.ActiveWindow.Selection
            sel_start = sel.Start
            sel_end = sel.End

            if sel_start == sel_end:
                return None

            all_sentences = _extract_sentences_from_doc(doc_com)
            if not all_sentences:
                return None

            selected = [
                s for s in all_sentences
                if s["range_start"] < sel_end and s["range_end"] > sel_start
            ]

            if not selected:
                return None

            for i, s in enumerate(selected):
                s["index"] = i

            return selected
        except Exception as e:
            logger.warning("extract_selected_sentences failed: %s", e)
            return None

    def navigate_to_sentence(self, doc: dict, sentence: dict) -> bool:
        """Перейти к предложению и выделить его.

        Returns:
            bool: True если выделение успешно.
        """
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return False
        try:
            rng = _find_sentence_range(doc_com, sentence)
            if rng:
                rng.Select()
                return True
            logger.warning(
                "navigate_to_sentence: не найдено sentence index=%s",
                sentence.get("index"),
            )
            return False
        except Exception as e:
            logger.error("navigate_to_sentence error: %s", e)
            return False

    def replace_sentence_text(
        self,
        doc: dict,
        sentence: dict,
        new_text: str,
        old_text: Optional[str] = None,
        all_sentences: Optional[list] = None,
    ) -> bool:
        # Делегируем WordProvider — WordEditor использует тот же API
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return False

        # Создаём временный doc-словарь для делегирования
        virtual_doc = {"type": "word", "com_object": doc_com}
        return self._word.replace_sentence_text(
            virtual_doc, sentence, new_text, old_text, all_sentences,
        )

    def get_icon(self) -> tuple[str, str]:
        return ("O", "#0078D4")
