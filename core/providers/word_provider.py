"""
Провайдер для Microsoft Word.

Реализует DocumentProvider для работы с Word.Document через COM.
"""

import difflib
import logging
import re
from typing import Any, Optional

import win32com.client
import win32gui

from core.providers.base import DocumentProvider

logger = logging.getLogger("core.providers.word")

# ─── Утилиты ────────────────────────────────────────────────────────────

_ABBREV_PATTERN = re.compile(
    r'(?:'
    r'г\.г|гг|тыс|млн|млрд|трлн|руб|коп|ед|шт'
    r'|т\.д|т\.п|т\.е|т\.к|пр|др'
    r'|ул|корп|кв|стр|пер|обл'
    r'|и\.о|врио|проф|доц|акад'
    r'|рис|табл|гл|пп|разд|см|ср|напр|прим'
    r'|(?<![а-яА-Яa-zA-ZёЁ])(?:г|д|п|ч|ст|им|[А-ЯA-Z])'
    r')\.$'
)

_WORD_SPECIAL_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')

_BODY_START_PATTERN = re.compile(
    r'^\s*\d+[\.\)]\s'
    r'|^\s*(статья|раздел|глава|часть)\s+\d'
    r'|^\s*(уважаем|настоящим\s+сообщ|сообщаем)',
    re.IGNORECASE,
)

_BULLET_PREFIX_RE = re.compile(
    r'^(\s*(?:\d+[.)]\s+|[•\-\u2022\u2023\u2043\u25e6\u25aa\u25ab\u2027\u2013\u2014\u2212*·]\s+))'
)


def _starts_with_lower(text):
    for ch in text:
        if ch.isalpha():
            return ch.islower()
        if not ch.isspace():
            return False
    return False


def _normalize_ws(text):
    return re.sub(r'\s+', ' ', text).strip()


def _strip_word_special(text):
    return _WORD_SPECIAL_RE.sub('', text)


def _merge_false_splits(sentences, doc_com):
    """Объединить ошибочно разделённые Word фрагменты предложений."""
    if not sentences:
        return sentences

    merged = [sentences[0]]

    for nxt in sentences[1:]:
        cur = merged[-1]
        gap = nxt["range_start"] - cur["range_end"]

        should_merge = False
        if gap <= 2:
            try:
                gap_text = doc_com.Range(cur["range_end"], nxt["range_start"]).Text or ""
            except Exception:
                gap_text = ""
            if '\x07' not in gap_text:
                if _starts_with_lower(nxt["text"]):
                    should_merge = True
                elif _ABBREV_PATTERN.search(cur["text"]):
                    should_merge = True

        if should_merge:
            new_start = cur["range_start"]
            new_end = nxt["range_end"]
            try:
                raw = doc_com.Range(new_start, new_end).Text or ""
                new_text = _strip_word_special(raw).strip()
            except Exception:
                new_text = cur["text"] + " " + nxt["text"]
            merged[-1] = {
                "index": cur["index"],
                "word_sentence_index": cur["word_sentence_index"],
                "range_start": new_start,
                "range_end": new_end,
                "text": new_text,
                "in_table": cur.get("in_table", False),
            }
        else:
            merged.append(nxt)

    for i, s in enumerate(merged):
        s["index"] = i

    return merged


def _find_body_start(doc_com) -> int:
    """Вернуть позицию начала основного текста документа."""
    try:
        paras = doc_com.Paragraphs
        for i in range(1, paras.Count + 1):
            try:
                para = paras.Item(i)
                rng = para.Range

                try:
                    if rng.Information(12):
                        continue
                except Exception:
                    pass

                text = rng.Text.strip()
                if not text:
                    continue

                if _BODY_START_PATTERN.search(text):
                    return rng.Start

                if len(text) >= 100 and text != text.upper():
                    return rng.Start

            except Exception:
                continue
    except Exception:
        pass

    return 0


def _find_sentence_range(doc_com, sentence, expected_text=None):
    """Найти Range предложения с 3-ступенчатым поиском."""
    if expected_text is None:
        expected_text = sentence.get("text", "")

    dbg_prefix = f"[FSR] sent={repr(expected_text[:40])}"

    # 1. По сохранённым позициям
    r_start = sentence.get("range_start")
    r_end = sentence.get("range_end")
    if r_start is not None and r_end is not None:
        try:
            rng = doc_com.Range(r_start, r_end)
            if rng.Text and _normalize_ws(_strip_word_special(rng.Text)) == _normalize_ws(_strip_word_special(expected_text)):
                logger.debug("%s → ступень 1 OK start=%s", dbg_prefix, r_start)
                return rng
        except Exception as e:
            logger.debug("%s ступень 1 EXCEPTION: %s", dbg_prefix, e)

    # 2. По word_sentence_index
    word_idx = sentence.get("word_sentence_index")
    if word_idx:
        try:
            rng = doc_com.Sentences.Item(word_idx)
            if rng.Text and _normalize_ws(_strip_word_special(rng.Text)) == _normalize_ws(_strip_word_special(expected_text)):
                logger.debug("%s → ступень 2 OK", dbg_prefix)
                return rng
        except Exception as e:
            logger.debug("%s ступень 2 EXCEPTION: %s", dbg_prefix, e)

    # 3. Через Range.Find
    try:
        search_text = expected_text[:255]
        found_rng = None
        if r_start is not None:
            try:
                rng = doc_com.Range(r_start, r_start)
                rng.Find.ClearFormatting()
                rng.Find.Text = search_text
                rng.Find.Forward = True
                rng.Find.Wrap = 0
                if rng.Find.Execute():
                    found_rng = rng
                    logger.debug("%s → ступень 3a OK", dbg_prefix)
            except Exception:
                pass
        if found_rng is None:
            rng = doc_com.Range(0, 0)
            rng.Find.ClearFormatting()
            rng.Find.Text = search_text
            rng.Find.Forward = True
            rng.Find.Wrap = 0
            if rng.Find.Execute():
                found_rng = rng
                logger.debug("%s → ступень 3b OK", dbg_prefix)
        if found_rng is not None:
            return found_rng
    except Exception as e:
        logger.debug("%s ступень 3 EXCEPTION: %s", dbg_prefix, e)

    return None


def _apply_diff_to_range(doc_com, rng, old_text, new_text):
    """Применить посимвольный diff к Range, сохраняя форматирование."""
    full_text = rng.Text
    if not full_text:
        return False

    leading_offset = len(full_text) - len(full_text.lstrip())
    content_text = full_text.strip()

    if _strip_word_special(content_text) != old_text:
        return False

    range_start = rng.Start + leading_offset
    opcodes = difflib.SequenceMatcher(None, old_text, new_text).get_opcodes()

    pos_map = []
    for idx, ch in enumerate(content_text):
        if not _WORD_SPECIAL_RE.search(ch):
            pos_map.append(idx)
    pos_map.append(len(content_text))

    applied_count = 0
    try:
        for tag, i1, i2, j1, j2 in reversed(opcodes):
            if tag == "equal":
                continue
            abs_start = range_start + pos_map[i1]
            abs_end = range_start + pos_map[i2]
            sub_rng = doc_com.Range(abs_start, abs_end)
            sub_rng.Text = new_text[j1:j2]
            applied_count += 1
    except Exception:
        for _ in range(applied_count):
            try:
                doc_com.Undo()
            except Exception:
                break
        return False

    return True


def _preserve_bullet_prefix(old_text, new_text):
    """Восстановить буллит/номер в начале new_text."""
    if not old_text or not new_text:
        return new_text
    m = _BULLET_PREFIX_RE.match(old_text)
    if m:
        prefix = m.group(1)
        if not new_text.startswith(prefix):
            new_text = prefix + new_text.lstrip()
    return new_text


def _after_replacement(sentence, new_text, rng_start, old_end, delta, all_sentences):
    """Обновить позиции после замены."""
    sentence["range_start"] = rng_start
    sentence["range_end"] = old_end + delta
    sentence["text"] = new_text.strip()

    if all_sentences and delta != 0:
        for s in all_sentences:
            if s is sentence:
                continue
            s_start = s.get("range_start")
            if s_start is not None and s_start >= old_end:
                s["range_start"] += delta
                s["range_end"] += delta


def _extract_sentences_from_doc(doc_com):
    """Извлечь предложения из COM-объекта Word.Document или WordEditor."""
    sentences = []
    try:
        word_sentences = doc_com.Sentences
        count = word_sentences.Count

        for i in range(1, count + 1):
            try:
                sent_range = word_sentences.Item(i)
                raw = sent_range.Text or ""
                text = _strip_word_special(raw).strip()

                if not text:
                    continue

                trailing = len(raw) - len(raw.rstrip('\x07'))
                try:
                    in_table = bool(sent_range.Information(12))
                except Exception:
                    in_table = False
                sentences.append({
                    "index": len(sentences),
                    "word_sentence_index": i,
                    "range_start": sent_range.Start,
                    "range_end": sent_range.End - trailing,
                    "text": text,
                    "in_table": in_table,
                })
            except Exception:
                continue
    except Exception:
        pass

    sentences = _merge_false_splits(sentences, doc_com)
    sentences = [s for s in sentences
                 if len(s["text"]) >= 3 and re.search(r'[а-яА-Яa-zA-ZёЁ]', s["text"])]
    for i, s in enumerate(sentences):
        s["index"] = i
    return sentences


# ─── WordProvider ────────────────────────────────────────────────────────

class WordProvider(DocumentProvider):
    """Провайдер для Microsoft Word."""

    @property
    def doc_type(self) -> str:
        return "word"

    def find_documents(self) -> list[dict]:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            docs = []
            for doc in word.Documents:
                try:
                    hwnd = doc.ActiveWindow.Hwnd
                    docs.append({
                        "name": doc.Name,
                        "hwnd": hwnd,
                        "type": "word",
                        "com_object": doc,
                    })
                except Exception:
                    continue
            return docs
        except Exception:
            return []

    def get_doc_com(self, doc: dict):
        return doc.get("com_object")

    def activate(self, doc: dict, target_rect: Optional[tuple]) -> None:
        try:
            hwnd = doc["hwnd"]
            win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
            if target_rect:
                x, y, width, height = target_rect
                win32gui.MoveWindow(hwnd, x, y, width, height, True)
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass

    def extract_sentences(self, doc: dict) -> list[dict]:
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return []
        sentences = _extract_sentences_from_doc(doc_com)
        body_start = _find_body_start(doc_com)
        if body_start > 0:
            sentences = [s for s in sentences if s["range_start"] >= body_start]
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
                "navigate_to_sentence: не найдено sentence index=%s text=%r",
                sentence.get("index"), sentence.get("text", "")[:60],
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
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return False

        try:
            expected = old_text if old_text is not None else sentence.get("text", "")
            rng = _find_sentence_range(doc_com, sentence, expected_text=expected)
            if rng is None:
                return False

            ref_text = old_text if old_text is not None else sentence.get("text", "")
            new_text = _preserve_bullet_prefix(ref_text, new_text)

            rng_start = rng.Start
            old_end = rng.End

            if old_text is not None:
                if _apply_diff_to_range(doc_com, rng, old_text, new_text):
                    delta = len(new_text) - len(old_text)
                    _after_replacement(sentence, new_text, rng_start, old_end, delta, all_sentences)
                    return True
                rng = _find_sentence_range(doc_com, sentence, expected_text=expected)
                if rng is None:
                    return False
                rng_start = rng.Start
                old_end = rng.End

            # Fallback
            original_text = rng.Text or ""
            stripped_original = original_text.rstrip('\r\n')
            trailing_cr_count = len(original_text) - len(stripped_original)
            new_text_stripped = new_text.rstrip()

            if trailing_cr_count > 0:
                content_rng = doc_com.Range(rng_start, old_end - trailing_cr_count)
                content_rng.Text = new_text_stripped
            else:
                rng.Text = new_text_stripped

            delta = len(new_text_stripped) - len(stripped_original)
            _after_replacement(sentence, new_text, rng_start, old_end, delta, all_sentences)
            return True
        except Exception as e:
            logger.error("replace_sentence_text error: %s", e)
            return False

    def get_icon(self) -> tuple[str, str]:
        return ("W", "#2B579A")
