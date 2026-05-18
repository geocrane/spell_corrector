"""
Провайдер для Microsoft Excel.

Реализует DocumentProvider для работы с Excel.Workbook через COM.

Особенности vs Word:
- Текст хранится в ячейках, а не в потоке Range.Sentences.
- range_start/range_end в sentence-dict — это позиции **внутри ячейки**,
  а не глобальные. Сдвиг при замене затрагивает только sentence'ы той же
  ячейки.
- Track Changes отсутствует — track_revisions флаг игнорируется.
- Skip-tables не имеет смысла (Excel сам — таблица). Опция блокируется в UI.
- Rich text: формат символов сохраняется через snapshot-and-restore
  (см. _snapshot_char_formats и _restore_char_formats).
"""

import difflib
import logging
import re
import time
import uuid
from collections import Counter
from typing import Any, Optional

import pywintypes
import win32api
import win32com.client
import win32con
import win32gui

from core.providers.base import DocumentProvider
from core.sentence_split import has_letters, split_into_sentences
from core.word_corrections import build_word_corrections

logger = logging.getLogger("core.providers.excel")

XL_SHEET_VISIBLE = -1
XL_CELL_TYPE_VISIBLE = 12

_VK_ESCAPE = 0x1B


def _has_letters(text: str) -> bool:
    return has_letters(text)


def _is_numeric_value(value) -> bool:
    """True если значение — число/дата (а не строка)."""
    if value is None:
        return False
    if isinstance(value, bool):
        return True
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, pywintypes.TimeType):
        return True
    return False


def _snapshot_font(font) -> dict:
    """Снять формат COM-объекта Characters.Font в простой dict.

    Возвращает None для атрибутов с неоднородным значением (но мы вызываем
    snapshot для отдельного символа — там значения всегда конкретные).
    """
    try:
        return {
            "name": font.Name,
            "size": font.Size,
            "bold": font.Bold,
            "italic": font.Italic,
            "underline": font.Underline,
            "color": font.Color,
            "strike": font.Strikethrough,
        }
    except pywintypes.com_error:
        return {}


def _apply_font(font, fmt: dict) -> None:
    """Применить dict с атрибутами к Characters.Font.

    Каждое присвоение в try/except — некоторые атрибуты могут падать
    на специфических версиях Excel.
    """
    for key, value in fmt.items():
        if value is None:
            continue
        try:
            if key == "name":
                font.Name = value
            elif key == "size":
                font.Size = value
            elif key == "bold":
                font.Bold = value
            elif key == "italic":
                font.Italic = value
            elif key == "underline":
                font.Underline = value
            elif key == "color":
                font.Color = value
            elif key == "strike":
                font.Strikethrough = value
        except pywintypes.com_error:
            continue


def _fonts_equal(a: dict, b: dict) -> bool:
    """Проверка эквивалентности двух snapshot-форматов."""
    if not a or not b:
        return a == b
    return all(a.get(k) == b.get(k) for k in ("name", "size", "bold", "italic", "underline", "color", "strike"))


class ExcelProvider(DocumentProvider):
    """Провайдер для Microsoft Excel."""

    def __init__(self):
        self._last_error: Optional[str] = None
        # snap_id → dict с ranges/total_len_signature для откатов унификации.
        self._snapshots: dict[str, dict] = {}

    @property
    def doc_type(self) -> str:
        return "excel"

    def get_icon(self) -> tuple[str, str]:
        return ("X", "#217346")

    def get_last_error(self) -> Optional[str]:
        return self._last_error

    # ─── Поиск документов ──────────────────────────────────────────────

    def find_documents(self) -> list[dict]:
        """Найти все открытые книги Excel."""
        try:
            app = win32com.client.GetActiveObject("Excel.Application")
        except pywintypes.com_error:
            return []
        except Exception as e:
            logger.warning("Excel.GetActiveObject failed: %s", e)
            return []

        docs: list[dict] = []
        try:
            hwnd = app.Hwnd
        except pywintypes.com_error:
            hwnd = 0

        try:
            for wb in app.Workbooks:
                try:
                    name = wb.Name
                except pywintypes.com_error:
                    continue
                docs.append({
                    "name": name,
                    "hwnd": hwnd,
                    "type": "excel",
                    "workbook": wb,
                    "application": app,
                })
        except pywintypes.com_error as e:
            logger.warning("Не удалось перебрать Workbooks: %s", e)

        return docs

    def get_doc_com(self, doc: dict) -> Any:
        return doc.get("workbook")

    def activate(self, doc: dict, target_rect: Optional[tuple] = None) -> None:
        """Активировать окно Excel и нужную книгу."""
        try:
            wb = doc.get("workbook")
            if wb is not None:
                wb.Activate()
        except pywintypes.com_error as e:
            logger.warning("Не удалось активировать книгу: %s", e)

        hwnd = doc.get("hwnd") or 0
        if not hwnd:
            return
        try:
            win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
            if target_rect:
                x, y, w, h = target_rect
                win32gui.MoveWindow(hwnd, int(x), int(y), int(w), int(h), True)
            win32gui.SetForegroundWindow(hwnd)
        except Exception as e:
            logger.debug("activate window: %s", e)

    # ─── Извлечение предложений ────────────────────────────────────────

    def extract_sentences(
        self,
        doc: dict,
        *,
        progress_callback=None,
        cancel_event=None,
    ) -> list[dict]:
        """Извлечь предложения из всех видимых текстовых ячеек книги."""
        wb = doc.get("workbook")
        if wb is None:
            return []

        sentences: list[dict] = []
        try:
            workbook_name = wb.Name
        except pywintypes.com_error:
            return []

        try:
            sheets = list(wb.Worksheets)
        except pywintypes.com_error as e:
            logger.warning("Не удалось получить Worksheets: %s", e)
            return []

        for sheet in sheets:
            if cancel_event is not None and cancel_event.is_set():
                break
            try:
                if sheet.Visible != XL_SHEET_VISIBLE:
                    continue
            except pywintypes.com_error:
                continue
            try:
                sheet_name = sheet.Name
            except pywintypes.com_error:
                continue
            self._extract_from_sheet(
                workbook_name, sheet, sheet_name, sentences,
                progress_callback=progress_callback, cancel_event=cancel_event,
            )

        for i, s in enumerate(sentences):
            s["index"] = i
        return sentences

    def extract_selected_sentences(
        self,
        doc: dict,
        *,
        progress_callback=None,
        cancel_event=None,
    ) -> Optional[list[dict]]:
        """Извлечь предложения из выделенных ячеек."""
        wb = doc.get("workbook")
        app = doc.get("application")
        if wb is None or app is None:
            return None

        try:
            sel = app.Selection
            if sel is None:
                return None
            workbook_name = wb.Name
            areas_count = sel.Areas.Count
        except pywintypes.com_error as e:
            logger.warning("Selection недоступен: %s", e)
            return None

        sentences: list[dict] = []
        for i in range(1, areas_count + 1):
            if cancel_event is not None and cancel_event.is_set():
                break
            try:
                area = sel.Areas(i)
                sheet = area.Worksheet
                sheet_name = sheet.Name
            except pywintypes.com_error:
                continue
            self._extract_from_range(
                workbook_name, sheet, sheet_name, area, sentences,
                progress_callback=progress_callback, cancel_event=cancel_event,
            )

        for i, s in enumerate(sentences):
            s["index"] = i
        return sentences

    def _extract_from_sheet(
        self, workbook_name, sheet, sheet_name, sentences_out,
        *, progress_callback=None, cancel_event=None,
    ):
        """Обойти UsedRange листа и накопить предложения."""
        try:
            used = sheet.UsedRange
        except pywintypes.com_error:
            return
        self._extract_from_range(
            workbook_name, sheet, sheet_name, used, sentences_out,
            progress_callback=progress_callback, cancel_event=cancel_event,
        )

    def _extract_from_range(
        self, workbook_name, sheet, sheet_name, rng, sentences_out,
        *, progress_callback=None, cancel_event=None,
    ):
        """Обойти ячейки в произвольном Range и накопить предложения.

        Каждые PROGRESS_INTERVAL ячеек вызывает progress_callback и проверяет
        cancel_event — это даёт UI шанс перерисоваться и обработать клик «Остановить».
        """
        try:
            cells_count = rng.Cells.Count
            if cells_count == 0:
                return
        except pywintypes.com_error:
            return

        try:
            sheet_protected = bool(sheet.ProtectContents)
        except pywintypes.com_error:
            sheet_protected = False

        try:
            cells_iter = rng.Cells
        except pywintypes.com_error:
            return

        PROGRESS_INTERVAL = 25
        processed = 0

        try:
            for cell in cells_iter:
                if cancel_event is not None and cancel_event.is_set():
                    return
                self._extract_from_cell(workbook_name, sheet_name, cell, sheet_protected, sentences_out)
                processed += 1
                if progress_callback is not None and processed % PROGRESS_INTERVAL == 0:
                    try:
                        progress_callback(len(sentences_out), processed)
                    except Exception:
                        logger.exception("progress_callback failed")
        except pywintypes.com_error as e:
            logger.warning("Ошибка обхода ячеек: %s", e)
        finally:
            if progress_callback is not None and processed > 0:
                try:
                    progress_callback(len(sentences_out), processed)
                except Exception:
                    pass

    def _extract_from_cell(self, workbook_name, sheet_name, cell, sheet_protected, sentences_out):
        """Извлечь предложения из одной ячейки (если она проходит фильтры)."""
        try:
            row = cell.Row
            col = cell.Column
            address = cell.Address
        except pywintypes.com_error:
            return

        try:
            if cell.EntireRow.Hidden or cell.EntireColumn.Hidden:
                return
        except pywintypes.com_error:
            pass

        try:
            if cell.HasFormula:
                return
        except pywintypes.com_error:
            return

        try:
            if cell.MergeCells:
                primary_addr = cell.MergeArea.Cells(1).Address
                if primary_addr != address:
                    return
        except pywintypes.com_error:
            pass

        if sheet_protected:
            try:
                if cell.Locked:
                    return
            except pywintypes.com_error:
                return

        try:
            value = cell.Value2
        except pywintypes.com_error:
            return

        if value is None or value == "":
            return
        if _is_numeric_value(value):
            return
        if not isinstance(value, str):
            return

        try:
            text = cell.Text or ""
        except pywintypes.com_error:
            text = value

        if not text or not _has_letters(text):
            return

        fragments = split_into_sentences(text)
        for char_start, char_end, sub_text in fragments:
            if not _has_letters(sub_text):
                continue
            sentences_out.append({
                "index": len(sentences_out),
                "workbook": workbook_name,
                "sheet": sheet_name,
                "cell_address": address,
                "row": row,
                "col": col,
                "char_offset": char_start,
                "length": char_end - char_start,
                "text": sub_text,
                "range_start": char_start,
                "range_end": char_end,
                "in_table": False,
                "cell_full_text": text,
            })

    # ─── Навигация ──────────────────────────────────────────────────────

    def navigate_to_sentence(self, doc: dict, sentence: dict) -> bool:
        """Активировать книгу/лист/ячейку с предложением.

        При ошибке (Excel в режиме редактирования) пытается снять блокировку
        отправкой VK_ESCAPE в hwnd окна, затем повторяет попытку.
        """
        self._last_error = None
        try:
            return self._do_navigate(doc, sentence["sheet"], sentence["cell_address"])
        except pywintypes.com_error:
            pass
        self._try_unblock_excel(doc)
        try:
            return self._do_navigate(doc, sentence["sheet"], sentence["cell_address"])
        except pywintypes.com_error as e:
            logger.warning("Excel заблокирован: %s", e)
            self._last_error = "Excel заблокирован пользователем"
            return False

    def navigate_to_correction(self, doc: dict, correction: dict) -> bool:
        """Активировать ячейку конкретной пословной правки (режим ошибок)."""
        sheet_name = correction.get("excel_sheet")
        cell_address = correction.get("excel_cell_address")
        if not sheet_name or not cell_address:
            return False

        self._last_error = None
        try:
            return self._do_navigate(doc, sheet_name, cell_address)
        except pywintypes.com_error:
            pass
        self._try_unblock_excel(doc)
        try:
            return self._do_navigate(doc, sheet_name, cell_address)
        except pywintypes.com_error as e:
            logger.warning("Excel заблокирован: %s", e)
            self._last_error = "Excel заблокирован пользователем"
            return False

    def _do_navigate(self, doc: dict, sheet_name: str, cell_address: str) -> bool:
        wb = doc.get("workbook")
        if wb is None:
            return False
        wb.Activate()
        sheet = wb.Worksheets(sheet_name)
        sheet.Activate()
        cell = sheet.Range(cell_address)
        cell.Activate()
        return True

    def _try_unblock_excel(self, doc: dict) -> None:
        """Попытаться снять режим редактирования через PostMessage Esc.

        Отправляет VK_ESCAPE прямо в hwnd главного окна Excel. Это
        безопаснее SendKeys (не зависит от того, какое окно сейчас активно
        в ОС) и достаточно надёжно для типичных кейсов (курсор в Formula
        Bar или редактируемой ячейке).
        """
        hwnd = doc.get("hwnd") or 0
        if not hwnd:
            return
        try:
            win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, _VK_ESCAPE, 0)
            win32api.PostMessage(hwnd, win32con.WM_KEYUP, _VK_ESCAPE, 0)
        except Exception as e:
            logger.debug("PostMessage Esc failed: %s", e)
        time.sleep(0.1)

    def navigate_to_range(self, doc: dict, start: int, end: int) -> bool:
        # Для Excel недоступно: start/end — позиции внутри неизвестной
        # ячейки. Реальная навигация идёт через navigate_to_correction.
        return False

    # ─── Замена с сохранением форматирования ───────────────────────────

    def replace_sentence_text(
        self,
        doc: dict,
        sentence: dict,
        new_text: str,
        old_text: Optional[str] = None,
        all_sentences: Optional[list] = None,
    ) -> bool:
        result = self.replace_sentence_text_with_corrections(
            doc, sentence, new_text, old_text=old_text, all_sentences=all_sentences,
        )
        return result.get("ok", False)

    def replace_sentence_text_with_corrections(
        self,
        doc: dict,
        sentence: dict,
        new_text: str,
        old_text: Optional[str] = None,
        all_sentences: Optional[list] = None,
        track_revisions: bool = False,
    ) -> dict:
        """Заменить текст предложения в ячейке с сохранением rich-text формата.

        track_revisions игнорируется (Excel не имеет TrackRevisions).
        Возвращает контракт {ok, corrections, delta, old_end, rng_start}.
        """
        empty = {"ok": False, "corrections": [], "delta": 0, "old_end": 0, "rng_start": 0}
        wb = doc.get("workbook")
        if wb is None:
            return empty

        try:
            sheet = wb.Worksheets(sentence["sheet"])
            cell = sheet.Range(sentence["cell_address"])
            current_full = cell.Text or ""
        except pywintypes.com_error as e:
            logger.warning("Не удалось прочитать ячейку %s: %s", sentence.get("cell_address"), e)
            return empty

        start, end = self._find_sentence_in_cell(current_full, sentence)
        if start < 0 or end <= start:
            logger.warning("Не нашли предложение %r в ячейке %s",
                           sentence.get("text", "")[:60], sentence.get("cell_address"))
            return empty

        existing_old = current_full[start:end]
        if old_text and existing_old != old_text:
            logger.debug("old_text mismatch in cell: expected %r, found %r",
                         old_text[:60], existing_old[:60])

        corrections = build_word_corrections(
            existing_old, new_text, sentence.get("index", 0), sentence_range_start=start,
        )
        cell_meta = {
            "excel_workbook": sentence.get("workbook"),
            "excel_sheet": sentence.get("sheet"),
            "excel_cell_address": sentence.get("cell_address"),
        }
        for c in corrections:
            c.update(cell_meta)
            c["excel_char_start"] = c.get("word_range_start")
            c["excel_char_end"] = c.get("word_range_end")

        new_full = current_full[:start] + new_text + current_full[end:]
        delta = len(new_text) - len(existing_old)

        app = doc.get("application")
        prev_screen_updating = None
        try:
            if app is not None:
                try:
                    prev_screen_updating = app.ScreenUpdating
                    app.ScreenUpdating = False
                except pywintypes.com_error:
                    prev_screen_updating = None

            if self._is_uniform_format(cell):
                # Формат единый по всей ячейке — Excel сохранит его при
                # записи cell.Value. Per-character snapshot не нужен,
                # экономим ~7×N COM-вызовов.
                cell.Value = new_full
            else:
                char_formats = self._snapshot_char_formats(cell, len(current_full))
                cell.Value = new_full
                self._restore_char_formats(
                    cell, current_full, new_full,
                    existing_old, new_text, start, end, char_formats,
                )
        except pywintypes.com_error as e:
            logger.warning("Ошибка записи ячейки %s: %s", sentence.get("cell_address"), e)
            return empty
        finally:
            if app is not None and prev_screen_updating is not None:
                try:
                    app.ScreenUpdating = prev_screen_updating
                except pywintypes.com_error:
                    pass

        self._after_excel_replacement(sentence, new_text, delta, end, new_full, all_sentences)

        return {
            "ok": True,
            "corrections": corrections,
            "delta": delta,
            "old_end": end,
            "rng_start": start,
        }

    def _find_sentence_in_cell(self, current_full: str, sentence: dict) -> tuple[int, int]:
        """Найти границы предложения в актуальном тексте ячейки.

        Сначала проверяем по сохранённым char_offset/length, потом fallback
        на поиск по тексту. Возвращает (-1, -1) если не нашли.
        """
        offset = sentence.get("char_offset", 0)
        length = sentence.get("length", 0)
        text = sentence.get("text", "")
        if not text:
            return -1, -1

        if 0 <= offset and offset + length <= len(current_full):
            candidate = current_full[offset:offset + length]
            if candidate == text:
                return offset, offset + length

        idx = current_full.find(text, offset)
        if idx < 0:
            idx = current_full.find(text)
        if idx < 0:
            return -1, -1
        return idx, idx + len(text)

    def _is_uniform_format(self, cell) -> bool:
        """True если формат всей ячейки однороден (нет rich text).

        Excel возвращает None из cell.Font.X, если соответствующий атрибут
        смешан между символами ячейки. Если ни один из ключевых атрибутов
        не None — формат единый, snapshot/restore не нужны.
        """
        try:
            font = cell.Font
            for attr in ("Bold", "Italic", "Size", "Name", "Color", "Underline", "Strikethrough"):
                try:
                    if getattr(font, attr) is None:
                        return False
                except pywintypes.com_error:
                    return False
            return True
        except pywintypes.com_error:
            return False

    def _snapshot_char_formats(self, cell, length: int) -> list[dict]:
        """Снять формат каждого символа ячейки до изменения."""
        formats: list[dict] = []
        for k in range(1, length + 1):
            try:
                ch = cell.Characters(k, 1)
                formats.append(_snapshot_font(ch.Font))
            except pywintypes.com_error:
                formats.append({})
        return formats

    def _restore_char_formats(
        self, cell, current_full, new_full, old_text, new_text,
        sentence_start, sentence_end, char_formats,
    ) -> None:
        """Восстановить формат символов после замены через cell.Value.

        Маппинг символов:
        - Префикс current_full[:sentence_start] → 1-в-1 из char_formats.
        - Внутри предложения — по diff old_text vs new_text:
          * equal: формат из соответствующего исходного символа.
          * insert/replace: формат первого затронутого символа исходного
            диапазона (или предыдущего, если затронут начало).
        - Суффикс current_full[sentence_end:] → 1-в-1 из char_formats.

        После составления полного списка new_formats группируем соседние
        символы с одинаковым форматом в один COM-вызов на run.
        """
        new_len = len(new_full)
        new_formats: list[dict] = [{}] * new_len

        new_formats[:sentence_start] = char_formats[:sentence_start]

        suffix_start_new = sentence_start + len(new_text)
        suffix_start_old = sentence_end
        suffix_len = len(current_full) - suffix_start_old
        for i in range(suffix_len):
            new_formats[suffix_start_new + i] = char_formats[suffix_start_old + i]

        matcher = difflib.SequenceMatcher(None, old_text, new_text)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                for k in range(i2 - i1):
                    src = sentence_start + i1 + k
                    dst = sentence_start + j1 + k
                    if 0 <= src < len(char_formats) and 0 <= dst < new_len:
                        new_formats[dst] = char_formats[src]
            elif tag in ("insert", "replace"):
                src_idx = sentence_start + (i1 if i1 < len(old_text) else max(i1 - 1, 0))
                if src_idx >= len(char_formats):
                    src_idx = len(char_formats) - 1
                template = char_formats[src_idx] if 0 <= src_idx < len(char_formats) else {}
                for k in range(j2 - j1):
                    dst = sentence_start + j1 + k
                    if 0 <= dst < new_len:
                        new_formats[dst] = template

        run_start = 0
        while run_start < new_len:
            fmt = new_formats[run_start]
            run_end = run_start + 1
            while run_end < new_len and _fonts_equal(new_formats[run_end], fmt):
                run_end += 1
            if fmt:
                try:
                    chars = cell.Characters(run_start + 1, run_end - run_start)
                    _apply_font(chars.Font, fmt)
                except pywintypes.com_error as e:
                    logger.debug("restore font run [%d:%d] failed: %s", run_start, run_end, e)
            run_start = run_end

    def _after_excel_replacement(
        self, sentence, new_text, delta, old_end, new_full, all_sentences,
    ) -> None:
        """Обновить sentence и сдвинуть позиции остальных предложений той же ячейки."""
        sentence["text"] = new_text
        sentence["length"] = sentence["length"] + delta
        sentence["range_end"] = sentence["range_start"] + sentence["length"]
        sentence["cell_full_text"] = new_full

        if not all_sentences or delta == 0:
            return

        wb = sentence.get("workbook")
        sh = sentence.get("sheet")
        addr = sentence.get("cell_address")

        for s in all_sentences:
            if s is sentence:
                continue
            if s.get("workbook") != wb or s.get("sheet") != sh or s.get("cell_address") != addr:
                continue
            s_offset = s.get("char_offset", 0)
            if s_offset >= old_end:
                s["char_offset"] = s_offset + delta
                s["range_start"] = s.get("range_start", s_offset) + delta
                s["range_end"] = s.get("range_end", s_offset) + delta
            s["cell_full_text"] = new_full

    # ─── Унификация форматирования ────────────────────────────────────────

    def _iter_text_cells(self, wb):
        """Yield текстовые ячейки книги (без формул, без скрытых).

        Возвращает (sheet, sheet_name, cell, text_len). Используем тот же
        фильтр, что и при extract_sentences, но без разбиения на предложения.
        """
        try:
            workbook_name = wb.Name
        except pywintypes.com_error:
            return
        try:
            sheets = list(wb.Worksheets)
        except pywintypes.com_error as e:
            logger.warning("iter_text_cells: Worksheets error: %s", e)
            return

        for sheet in sheets:
            try:
                if sheet.Visible != XL_SHEET_VISIBLE:
                    continue
                sheet_name = sheet.Name
                used = sheet.UsedRange
            except pywintypes.com_error:
                continue

            try:
                cells_iter = used.Cells
            except pywintypes.com_error:
                continue

            try:
                for cell in cells_iter:
                    try:
                        if cell.HasFormula:
                            continue
                        if cell.EntireRow.Hidden or cell.EntireColumn.Hidden:
                            continue
                        value = cell.Value2
                        if value is None or value == "":
                            continue
                        if _is_numeric_value(value):
                            continue
                        if not isinstance(value, str):
                            continue
                        text = cell.Text or ""
                    except pywintypes.com_error:
                        continue
                    text_len = len(text)
                    if text_len == 0:
                        continue
                    yield sheet, sheet_name, cell, text_len
            except pywintypes.com_error as e:
                logger.warning("iter_text_cells: ошибка обхода UsedRange: %s", e)

    def _cell_font_value(self, cell, attr: str):
        """Прочитать атрибут Font ячейки. Возвращает None для смешанного."""
        try:
            font = cell.Font
        except pywintypes.com_error:
            return None
        try:
            if attr == "font_name":
                return font.Name
            if attr == "font_size":
                return font.Size
            if attr == "style":
                bold = font.Bold
                italic = font.Italic
                if bold is None or italic is None:
                    return None
                return _excel_style_key(bold, italic)
        except pywintypes.com_error:
            return None
        return None

    def _char_font_value(self, cell, char_idx: int, attr: str):
        """Прочитать значение атрибута для одного символа ячейки (1-based)."""
        try:
            font = cell.Characters(char_idx, 1).Font
            if attr == "font_name":
                return font.Name
            if attr == "font_size":
                return font.Size
            if attr == "style":
                return _excel_style_key(font.Bold, font.Italic)
        except pywintypes.com_error:
            return None
        return None

    def analyze_format(
        self, doc: dict, *, cancel_event=None, progress_callback=None,
    ) -> Optional[dict]:
        wb = doc.get("workbook")
        if wb is None:
            return None

        font_counter: Counter = Counter()
        size_counter: Counter = Counter()
        style_counter: Counter = Counter()
        total_chars = 0

        for _sheet, _sheet_name, cell, text_len in self._iter_text_cells(wb):
            name_val = self._cell_font_value(cell, "font_name")
            size_val = self._cell_font_value(cell, "font_size")
            style_val = self._cell_font_value(cell, "style")
            uniform = name_val is not None and size_val is not None and style_val is not None

            if uniform:
                total_chars += text_len
                font_counter[name_val] += text_len
                try:
                    size_counter[float(size_val)] += text_len
                except (TypeError, ValueError):
                    pass
                style_counter[style_val] += text_len
                continue

            # Rich text — посимвольный обход
            for k in range(1, text_len + 1):
                nv = self._char_font_value(cell, k, "font_name")
                sv = self._char_font_value(cell, k, "font_size")
                stv = self._char_font_value(cell, k, "style")
                if nv is None and sv is None and stv is None:
                    continue
                total_chars += 1
                if nv is not None:
                    font_counter[nv] += 1
                if sv is not None:
                    try:
                        size_counter[float(sv)] += 1
                    except (TypeError, ValueError):
                        pass
                if stv is not None:
                    style_counter[stv] += 1

        result = {
            "dominant_font_name": None,
            "dominant_font_size": None,
            "dominant_style": "normal",
            "coverage": {},
            "total_chars": total_chars,
        }
        if font_counter:
            top_name, top_n = font_counter.most_common(1)[0]
            result["dominant_font_name"] = top_name
            result["coverage"]["font_name"] = top_n / max(1, sum(font_counter.values()))
        if size_counter:
            top_size, top_s = size_counter.most_common(1)[0]
            result["dominant_font_size"] = top_size
            result["coverage"]["font_size"] = top_s / max(1, sum(size_counter.values()))
        if style_counter:
            top_style, top_st = style_counter.most_common(1)[0]
            result["dominant_style"] = top_style
            result["coverage"]["style"] = top_st / max(1, sum(style_counter.values()))
        return result

    def apply_format_uniform(
        self,
        doc: dict,
        attr: str,
        target_value: Any = None,
        *,
        cancel_event=None,
        progress_callback=None,
    ) -> Optional[dict]:
        if attr not in ("font_name", "font_size", "style"):
            return None
        wb = doc.get("workbook")
        if wb is None:
            return None

        if target_value is None:
            stats = self.analyze_format(doc)
            if not stats:
                return None
            if attr == "font_name":
                target_value = stats.get("dominant_font_name")
            elif attr == "font_size":
                target_value = stats.get("dominant_font_size")
            elif attr == "style":
                target_value = stats.get("dominant_style")
        if target_value is None:
            return None

        # Снижение мерцания на время массовой записи.
        app = doc.get("application")
        prev_screen_updating = None
        if app is not None:
            try:
                prev_screen_updating = app.ScreenUpdating
                app.ScreenUpdating = False
            except pywintypes.com_error:
                prev_screen_updating = None

        ranges: list[dict] = []
        try:
            for _sheet, sheet_name, cell, text_len in self._iter_text_cells(wb):
                try:
                    cell_address = cell.Address
                except pywintypes.com_error:
                    continue

                current = self._cell_font_value(cell, attr)
                if current is not None:
                    # Однородная ячейка.
                    if _excel_attr_equal(attr, current, target_value):
                        continue
                    ranges.append({
                        "sheet": sheet_name,
                        "cell": cell_address,
                        "text_len": text_len,
                        "mode": "cell",
                        "original": current,
                    })
                    _excel_apply_font_attr(cell.Font, attr, target_value)
                else:
                    # Rich-text — посимвольно.
                    char_records: list[tuple[int, Any]] = []
                    for k in range(1, text_len + 1):
                        cv = self._char_font_value(cell, k, attr)
                        if cv is None or _excel_attr_equal(attr, cv, target_value):
                            continue
                        char_records.append((k, cv))
                    if not char_records:
                        continue
                    # Сгруппировать соседние char_idx с одинаковым cv для эффективности.
                    for run_start, run_len in _group_char_runs(char_records):
                        try:
                            chars = cell.Characters(run_start, run_len).Font
                            _excel_apply_font_attr(chars, attr, target_value)
                        except pywintypes.com_error as e:
                            logger.debug("apply chars [%s:%s] %s: %s",
                                         run_start, run_len, cell_address, e)
                    ranges.append({
                        "sheet": sheet_name,
                        "cell": cell_address,
                        "text_len": text_len,
                        "mode": "chars",
                        "chars": char_records,
                    })
        finally:
            if app is not None and prev_screen_updating is not None:
                try:
                    app.ScreenUpdating = prev_screen_updating
                except pywintypes.com_error:
                    pass

        snap_id = f"{id(wb)}:{attr}:{uuid.uuid4().hex[:8]}"
        self._snapshots[snap_id] = {
            "attr": attr,
            "target": target_value,
            "ranges": ranges,
        }
        logger.info("excel apply_format_uniform: attr=%s target=%r cells=%d",
                    attr, target_value, len(ranges))
        return {
            "snap_id": snap_id,
            "attr": attr,
            "target": target_value,
            "segments": len(ranges),
        }

    def restore_format(self, doc: dict, snapshot_descriptor: dict) -> bool:
        wb = doc.get("workbook")
        if wb is None:
            return False
        snap_id = snapshot_descriptor.get("snap_id") if snapshot_descriptor else None
        if not snap_id:
            return False
        snap = self._snapshots.pop(snap_id, None)
        if snap is None:
            return False

        attr = snap.get("attr")
        ranges = snap.get("ranges", [])
        if attr not in ("font_name", "font_size", "style"):
            return False

        app = doc.get("application")
        prev_screen_updating = None
        if app is not None:
            try:
                prev_screen_updating = app.ScreenUpdating
                app.ScreenUpdating = False
            except pywintypes.com_error:
                prev_screen_updating = None

        ok_all = True
        try:
            for rec in ranges:
                sheet_name = rec.get("sheet")
                cell_addr = rec.get("cell")
                expected_len = rec.get("text_len", 0)
                try:
                    sheet = wb.Worksheets(sheet_name)
                    cell = sheet.Range(cell_addr)
                    text = cell.Text or ""
                except pywintypes.com_error as e:
                    logger.debug("restore: cell %s!%s недоступна: %s",
                                 sheet_name, cell_addr, e)
                    ok_all = False
                    continue
                if len(text) != expected_len:
                    logger.info("restore: cell %s!%s изменилась (%d → %d) — пропуск",
                                sheet_name, cell_addr, expected_len, len(text))
                    ok_all = False
                    continue

                if rec.get("mode") == "cell":
                    _excel_apply_font_attr(cell.Font, attr, rec.get("original"))
                elif rec.get("mode") == "chars":
                    for k, orig_val in rec.get("chars", []):
                        try:
                            font = cell.Characters(k, 1).Font
                            _excel_apply_font_attr(font, attr, orig_val)
                        except pywintypes.com_error as e:
                            logger.debug("restore char %s!%s[%d]: %s",
                                         sheet_name, cell_addr, k, e)
                            ok_all = False
        finally:
            if app is not None and prev_screen_updating is not None:
                try:
                    app.ScreenUpdating = prev_screen_updating
                except pywintypes.com_error:
                    pass

        return ok_all


def _excel_style_key(bold_val, italic_val) -> str:
    b = bold_val is True or bold_val == -1
    i = italic_val is True or italic_val == -1
    if b and i:
        return "bold_italic"
    if b:
        return "bold"
    if i:
        return "italic"
    return "normal"


def _excel_attr_equal(attr: str, a, b) -> bool:
    if attr == "font_size":
        try:
            return abs(float(a) - float(b)) < 1e-6
        except (TypeError, ValueError):
            return False
    return a == b


def _excel_apply_font_attr(font, attr: str, value) -> None:
    """Применить одно значение атрибута к Font."""
    if value is None:
        return
    try:
        if attr == "font_name":
            font.Name = str(value)
        elif attr == "font_size":
            font.Size = float(value)
        elif attr == "style":
            bold = value in ("bold", "bold_italic")
            italic = value in ("italic", "bold_italic")
            font.Bold = bold
            font.Italic = italic
    except pywintypes.com_error as e:
        logger.debug("apply font %s=%r failed: %s", attr, value, e)


def _group_char_runs(char_records: list[tuple[int, Any]]) -> list[tuple[int, int]]:
    """Сгруппировать соседние char-индексы в run'ы (start, length).

    Возвращает диапазоны для COM-вызова cell.Characters(start, length).
    """
    if not char_records:
        return []
    char_records = sorted(char_records, key=lambda x: x[0])
    runs: list[tuple[int, int]] = []
    run_start = char_records[0][0]
    run_len = 1
    for prev, curr in zip(char_records, char_records[1:]):
        if curr[0] == prev[0] + 1:
            run_len += 1
        else:
            runs.append((run_start, run_len))
            run_start = curr[0]
            run_len = 1
    runs.append((run_start, run_len))
    return runs
