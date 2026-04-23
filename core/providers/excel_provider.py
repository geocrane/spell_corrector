"""
Провайдер для Microsoft Excel (заглушка).

Excel не использует Word API — работа идёт через Range ячеек.
Этот провайдер готов для будущей реализации.
"""

import logging
from typing import Any, Optional

import win32com.client
import win32gui

from core.providers.base import DocumentProvider

logger = logging.getLogger("core.providers.excel")


class ExcelProvider(DocumentProvider):
    """Провайдер для Microsoft Excel.

    TODO: реализовать при необходимости.
    Excel использует другой API: Workbook → Worksheet → Range/Cells,
    а не Word.Sentences.
    """

    @property
    def doc_type(self) -> str:
        return "excel"

    def find_documents(self) -> list[dict]:
        # TODO: реализовать поиск открытых книг Excel
        # try:
        #     excel = win32com.client.GetActiveObject("Excel.Application")
        #     docs = []
        #     for wb in excel.Workbooks:
        #         docs.append({
        #             "name": wb.Name,
        #             "hwnd": excel.Hwnd,
        #             "type": "excel",
        #             "workbook": wb,
        #             "application": excel,
        #         })
        #     return docs
        # except Exception:
        #     return []
        return []

    def get_doc_com(self, doc: dict) -> Any:
        # TODO: вернуть активный worksheet
        return None

    def activate(self, doc: dict, target_rect: Optional[tuple]) -> None:
        # TODO: активировать окно Excel
        pass

    def extract_sentences(self, doc: dict) -> list[dict]:
        # TODO: извлечь текст из ячеек worksheet
        return []

    def extract_selected_sentences(self, doc: dict) -> Optional[list[dict]]:
        # TODO: извлечь текст из выделенных ячеек
        return None

    def navigate_to_sentence(self, doc: dict, sentence: dict) -> bool:
        # TODO: выделить ячейку/диапазон
        return False

    def replace_sentence_text(
        self,
        doc: dict,
        sentence: dict,
        new_text: str,
        old_text: Optional[str] = None,
        all_sentences: Optional[list] = None,
    ) -> bool:
        # TODO: заменить текст в ячейке
        return False

    def get_icon(self) -> tuple[str, str]:
        return ("X", "#217346")
