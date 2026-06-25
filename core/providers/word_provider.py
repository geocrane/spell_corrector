"""
Провайдер для Microsoft Word.

Реализует DocumentProvider для работы с Word.Document через COM.
"""

import difflib
import logging
import re
import uuid
from dataclasses import asdict
from typing import Any, Optional

import win32com.client
import win32gui

from core import format_unifier
from core.providers.base import DocumentProvider
from core.sentence_split import ABBREV_PATTERN as _ABBREV_PATTERN, starts_with_lower as _starts_with_lower
from core.word_corrections import build_word_corrections

logger = logging.getLogger("core.providers.word")

# ─── Утилиты ────────────────────────────────────────────────────────────

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
        # ends_para: cur завершает абзац (\r) или ячейку таблицы (\x07) —
        # через границу абзаца склейка запрещена (каждый пункт списка / абзац
        # остаётся отдельным предложением).
        if gap <= 2 and not cur.get("ends_para", False):
            try:
                gap_text = doc_com.Range(cur["range_end"], nxt["range_start"]).Text or ""
            except Exception:
                gap_text = ""
            # \x07 — разделитель ячеек таблицы; \r — конец абзаца (буллит/параграф).
            # Через оба символа склейка запрещена — каждый абзац отдельное предложение.
            if '\x07' not in gap_text and '\r' not in gap_text:
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
                # объединённое предложение заканчивается там же, где nxt
                "ends_para": nxt.get("ends_para", False),
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


def _apply_diff_to_range(doc_com, rng, old_text, new_text, skip_post_validation=False):
    """Применить посимвольный diff к Range, сохраняя форматирование.

    При skip_post_validation=True пост-проверка пропускается. Это нужно при
    включённом TrackRevisions=True: Word держит и зачёркнутый оригинал, и
    новые вставки в одном диапазоне, и обычное сравнение текста не работает.
    Также при skip_post_validation=True мы НЕ делаем Undo при ошибке, чтобы
    не разрушить уже записанные ревизии.
    """
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
    except Exception as e:
        logger.debug("_apply_diff_to_range exception, applied=%d, skip_post=%s: %s",
                     applied_count, skip_post_validation, e)
        if not skip_post_validation:
            for _ in range(applied_count):
                try:
                    doc_com.Undo()
                except Exception:
                    break
        return False

    if skip_post_validation:
        return True

    try:
        delta = len(new_text) - len(old_text)
        check_rng = doc_com.Range(range_start, rng.End + delta)
        actual = _strip_word_special(check_rng.Text or "").rstrip('\r\n')
        if actual != new_text:
            logger.warning(
                "_apply_diff_to_range пост-проверка не прошла: ожидали %r, получили %r — откатываемся",
                new_text[:80], actual[:80],
            )
            for _ in range(applied_count):
                try:
                    doc_com.Undo()
                except Exception:
                    break
            return False
    except Exception as e:
        logger.debug("_apply_diff_to_range пост-проверка EXCEPTION: %s", e)

    return True


def _apply_diff_to_range_chunked(doc_com, rng, old_text, new_text, corrections,
                                  skip_post_validation=False):
    """Применить пословные правки к Range отдельными мини-заменами.

    При skip_post_validation=True (включённый TrackRevisions) пост-проверка
    и Undo-rollback пропускаются: Word держит «оригинал + вставка» в одном
    диапазоне, и сравнение текста не работает; а Undo откатил бы записанную
    ревизию и испортил историю изменений.
    """
    full_text = rng.Text
    if not full_text:
        return False

    leading_offset = len(full_text) - len(full_text.lstrip())
    content_text = full_text.strip()

    if _strip_word_special(content_text) != old_text:
        return False

    range_start = rng.Start + leading_offset

    pos_map = []
    for idx, ch in enumerate(content_text):
        if not _WORD_SPECIAL_RE.search(ch):
            pos_map.append(idx)
    pos_map.append(len(content_text))

    applied_count = 0
    try:
        # Идём справа налево, чтобы позиции более ранних правок не сдвигались.
        for corr in reversed(corrections):
            i1 = corr["i1_in_original"]
            i2 = corr["i2_in_original"]
            j1 = corr["j1_in_corrected"]
            j2 = corr["j2_in_corrected"]

            if i1 == i2 and j1 == j2:
                continue

            abs_start = range_start + pos_map[i1]
            abs_end = range_start + pos_map[i2]
            sub_rng = doc_com.Range(abs_start, abs_end)
            sub_rng.Text = new_text[j1:j2]
            applied_count += 1
    except Exception as e:
        logger.debug("_apply_diff_to_range_chunked exception applied=%d skip_post=%s: %s",
                     applied_count, skip_post_validation, e)
        if not skip_post_validation:
            for _ in range(applied_count):
                try:
                    doc_com.Undo()
                except Exception:
                    break
        return False

    if skip_post_validation:
        return True

    try:
        delta = len(new_text) - len(old_text)
        check_rng = doc_com.Range(range_start, rng.End + delta)
        actual = _strip_word_special(check_rng.Text or "").rstrip('\r\n')
        if actual != new_text:
            logger.warning(
                "_apply_diff_to_range_chunked пост-проверка не прошла: ожидали %r, получили %r — откатываемся",
                new_text[:80], actual[:80],
            )
            for _ in range(applied_count):
                try:
                    doc_com.Undo()
                except Exception:
                    break
            return False
    except Exception as e:
        logger.debug("_apply_diff_to_range_chunked пост-проверка EXCEPTION: %s", e)

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


_SENTENCE_FINAL = frozenset('.!?…')


def _preserve_bullet_ending(old_text, new_text):
    """Восстановить концевую пунктуацию буллита, если T5 изменил её.

    Буллиты часто заканчиваются на ; , : или вообще без знака.
    T5 нередко добавляет точку — это нежелательно для элементов списка.
    Если оригинал НЕ заканчивался на . ! ? … — восстанавливаем исходное окончание.
    """
    if not old_text or not new_text:
        return new_text
    old_stripped = old_text.rstrip()
    if not old_stripped or old_stripped[-1] in _SENTENCE_FINAL:
        return new_text  # обычное предложение — не трогаем
    old_last = old_stripped[-1]
    new_stripped = new_text.rstrip()
    if not new_stripped:
        return new_text
    if new_stripped[-1] == '.' and old_last != '.':
        trailing = new_text[len(new_stripped):]
        return new_stripped[:-1] + old_last + trailing
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
                    # \r (конец абзаца) и \x07 (ячейка таблицы) Word включает в
                    # конец диапазона предложения. Запоминаем это здесь: в
                    # _merge_false_splits «зазор» между предложениями пуст и
                    # детектировать границу абзаца по нему нельзя.
                    "ends_para": ('\r' in raw) or ('\x07' in raw),
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

    def __init__(self):
        # snap_id → format_unifier.Snapshot. Хранятся в провайдере, наружу
        # отдаётся только лёгкий дескриптор {snap_id, attr, target, coverage}.
        self._snapshots: dict[str, format_unifier.Snapshot] = {}

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

    def activate_two(
        self,
        doc_a: dict,
        doc_b: dict,
        monitor: dict,
        app_width: int,
        layout: str = "vertical",
    ) -> None:
        """Разместить два Word-окна слева от панели приложения.

        layout="vertical"   — стопкой (doc_a сверху, doc_b снизу)
        layout="horizontal" — рядом (doc_a слева, doc_b правее)
        """
        try:
            left_total = monitor["width"] - app_width
            x0, y0 = monitor["x"], monitor["y"]
            if layout == "horizontal":
                half_w = left_total // 2
                positions = [
                    (doc_a["hwnd"], x0,           y0, half_w, monitor["height"]),
                    (doc_b["hwnd"], x0 + half_w,  y0, half_w, monitor["height"]),
                ]
            else:  # vertical
                half_h = monitor["height"] // 2
                positions = [
                    (doc_a["hwnd"], x0, y0,          left_total, half_h),
                    (doc_b["hwnd"], x0, y0 + half_h, left_total, half_h),
                ]
            for hwnd, x, y, w, h in positions:
                win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
                win32gui.MoveWindow(hwnd, x, y, w, h, True)
            win32gui.SetForegroundWindow(doc_a["hwnd"])
        except Exception:
            pass

    def extract_sentences(
        self,
        doc: dict,
        *,
        progress_callback=None,
        cancel_event=None,
    ) -> list[dict]:
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return []
        sentences = _extract_sentences_from_doc(doc_com)
        body_start = _find_body_start(doc_com)
        if body_start > 0:
            sentences = [s for s in sentences if s["range_start"] >= body_start]
        if progress_callback is not None:
            try:
                progress_callback(len(sentences), None)
            except Exception:
                pass
        return sentences

    def extract_selected_sentences(
        self,
        doc: dict,
        *,
        progress_callback=None,
        cancel_event=None,
    ) -> Optional[list[dict]]:
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

            if progress_callback is not None:
                try:
                    progress_callback(len(selected), None)
                except Exception:
                    pass
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
        """Тонкая обёртка, оставленная для обратной совместимости (Outlook/Excel)."""
        result = self.replace_sentence_text_with_corrections(
            doc, sentence, new_text, old_text, all_sentences, track_revisions=False,
        )
        return result["ok"]

    def replace_sentence_text_with_corrections(
        self,
        doc: dict,
        sentence: dict,
        new_text: str,
        old_text: Optional[str] = None,
        all_sentences: Optional[list] = None,
        track_revisions: bool = False,
    ) -> dict:
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return {"ok": False, "corrections": [], "delta": 0, "old_end": 0, "rng_start": 0}

        prev_track = None
        try:
            try:
                prev_track = bool(doc_com.TrackRevisions)
            except Exception:
                prev_track = None

            if track_revisions:
                try:
                    doc_com.TrackRevisions = True
                    logger.info("TrackRevisions включён для index=%s (был=%s)",
                                sentence.get("index"), prev_track)
                except Exception as e:
                    logger.warning("Не удалось включить TrackRevisions: %s", e)

            # Сохранить общую длину документа и число ревизий ДО операции —
            # нужно, чтобы корректно вычислить реальный сдвиг позиций.
            doc_len_before = None
            revisions_before = None
            try:
                doc_len_before = int(doc_com.Range().End)
            except Exception:
                pass
            try:
                revisions_before = int(doc_com.Revisions.Count)
            except Exception:
                pass

            expected = old_text if old_text is not None else sentence.get("text", "")
            rng = _find_sentence_range(doc_com, sentence, expected_text=expected)
            if rng is None:
                return {"ok": False, "corrections": [], "delta": 0, "old_end": 0, "rng_start": 0}

            ref_text = old_text if old_text is not None else sentence.get("text", "")
            new_text = _preserve_bullet_prefix(ref_text, new_text)
            new_text = _preserve_bullet_ending(ref_text, new_text)

            rng_start = rng.Start
            old_end = rng.End

            def _real_delta(default):
                """Вычислить актуальный сдвиг длины документа после операции.

                При TrackRevisions «удалённый» текст не уходит из документа —
                он остаётся как зачёркнутый, поэтому len(new)-len(old) неверен.
                """
                if doc_len_before is None:
                    return default
                try:
                    return int(doc_com.Range().End) - doc_len_before
                except Exception:
                    return default

            def _log_revisions_after():
                if not track_revisions:
                    return
                try:
                    after = int(doc_com.Revisions.Count)
                    delta_revs = (after - revisions_before) if revisions_before is not None else after
                    logger.info(
                        "ревизий стало %s (+%s) после apply index=%s",
                        after, delta_revs, sentence.get("index"),
                    )
                except Exception as e:
                    logger.debug("не удалось прочитать Revisions.Count: %s", e)

            def _make_marker(initial_rng_start, initial_old_end, current_delta):
                """Снапшот ревизий, созданных операцией, для будущего Reject."""
                if not track_revisions or revisions_before is None:
                    return None
                try:
                    count_after = int(doc_com.Revisions.Count)
                except Exception:
                    return None
                return {
                    "count_before": revisions_before,
                    "count_after": count_after,
                    "range_start": initial_rng_start,
                    "range_end_after": initial_old_end + current_delta,
                }

            corrections = []
            if old_text is not None:
                corrections = build_word_corrections(
                    old_text, new_text,
                    sentence_index=sentence.get("index", 0),
                    sentence_range_start=rng_start,
                )

                if corrections and _apply_diff_to_range_chunked(
                    doc_com, rng, old_text, new_text, corrections,
                    skip_post_validation=track_revisions,
                ):
                    delta = _real_delta(len(new_text) - len(old_text))
                    _after_replacement(sentence, new_text, rng_start, old_end, delta, all_sentences)
                    _log_revisions_after()
                    return {
                        "ok": True, "corrections": corrections,
                        "delta": delta, "old_end": old_end, "rng_start": rng_start,
                        "revisions_marker": _make_marker(rng_start, old_end, delta),
                    }

                # Перепоиск range на случай частичного применения и отката
                rng = _find_sentence_range(doc_com, sentence, expected_text=expected)
                if rng is None:
                    return {"ok": False, "corrections": [], "delta": 0, "old_end": 0, "rng_start": 0}
                rng_start = rng.Start
                old_end = rng.End

                if _apply_diff_to_range(
                    doc_com, rng, old_text, new_text,
                    skip_post_validation=track_revisions,
                ):
                    delta = _real_delta(len(new_text) - len(old_text))
                    _after_replacement(sentence, new_text, rng_start, old_end, delta, all_sentences)
                    _log_revisions_after()
                    return {
                        "ok": True, "corrections": corrections,
                        "delta": delta, "old_end": old_end, "rng_start": rng_start,
                        "revisions_marker": _make_marker(rng_start, old_end, delta),
                    }

                rng = _find_sentence_range(doc_com, sentence, expected_text=expected)
                if rng is None:
                    return {"ok": False, "corrections": [], "delta": 0, "old_end": 0, "rng_start": 0}
                rng_start = rng.Start
                old_end = rng.End

            # Fallback: замена всего предложения. При track_revisions=True это даст
            # одну большую ревизию для этого предложения — но это лучше, чем потерять
            # запись (раньше мы тут выключали TrackRevisions и теряли подсветку).
            if track_revisions:
                logger.warning(
                    "fallback на полную замену предложения index=%s — будет одна большая ревизия",
                    sentence.get("index"),
                )

            original_text = rng.Text or ""
            stripped_original = original_text.rstrip('\r\n')
            trailing_cr_count = len(original_text) - len(stripped_original)
            new_text_stripped = new_text.rstrip()

            if trailing_cr_count > 0:
                content_rng = doc_com.Range(rng_start, old_end - trailing_cr_count)
                content_rng.Text = new_text_stripped
            else:
                rng.Text = new_text_stripped

            delta = _real_delta(len(new_text_stripped) - len(stripped_original))
            _after_replacement(sentence, new_text, rng_start, old_end, delta, all_sentences)
            _log_revisions_after()
            return {
                "ok": True, "corrections": corrections,
                "delta": delta, "old_end": old_end, "rng_start": rng_start,
                "revisions_marker": _make_marker(rng_start, old_end, delta),
            }
        except Exception as e:
            logger.error("replace_sentence_text_with_corrections error: %s", e)
            return {"ok": False, "corrections": [], "delta": 0, "old_end": 0, "rng_start": 0}
        finally:
            if track_revisions and prev_track is not None:
                try:
                    doc_com.TrackRevisions = prev_track
                except Exception:
                    pass

    def reject_sentence_revisions(
        self,
        doc: dict,
        sentence: dict,
        marker: dict,
        original_text: str,
        all_sentences: Optional[list] = None,
    ) -> dict:
        """Откатить ревизии предложения через Revision.Reject().

        Word сам корректно восстановит исходный текст: для insertion-ревизии
        удалит вставленные символы, для deletion-ревизии вернёт зачёркнутые.
        Это надёжнее, чем повторное вычисление обратного diff поверх документа
        в состоянии TrackRevisions.

        Args:
            doc: дескриптор документа.
            sentence: предложение для пост-обновления позиций.
            marker: revisions_marker, сохранённый при apply.
            original_text: исходный текст предложения (до Apply) — записывается
                в sentence["text"] после успешного Reject.
            all_sentences: для сдвига позиций других предложений.

        Returns:
            dict с ключами:
                ok (bool), reason (str, при ok=False),
                delta (int), old_end (int), rng_start (int).
        """
        empty_fail = {"ok": False, "delta": 0, "old_end": 0, "rng_start": 0}
        if not marker:
            return {**empty_fail, "reason": "no_marker"}

        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return {**empty_fail, "reason": "no_doc"}

        try:
            range_start = int(marker.get("range_start", 0))
            range_end_after = int(marker.get("range_end_after", 0))
            expected_count = int(marker.get("count_after", 0)) - int(marker.get("count_before", 0))

            if expected_count <= 0 or range_end_after <= range_start:
                return {**empty_fail, "reason": "invalid_marker"}

            try:
                doc_len_before = int(doc_com.Range().End)
            except Exception:
                doc_len_before = None

            # Запрашиваем только ревизии целевого диапазона — O(M) вместо O(N)
            # по всему документу. Каждый COM-вызов пересекает границу процессов,
            # поэтому разница заметна даже при десятке ревизий.
            to_reject = []
            try:
                range_rng = doc_com.Range(range_start, range_end_after)
                range_revisions = range_rng.Revisions
                count_in_range = int(range_revisions.Count)
                for i in range(1, count_in_range + 1):
                    rev = range_revisions.Item(i)
                    rev_start = int(rev.Range.Start)
                    to_reject.append((rev_start, rev))
            except Exception:
                return {**empty_fail, "reason": "no_revisions_api"}

            if not to_reject:
                return {**empty_fail, "reason": "not_found"}

            if len(to_reject) != expected_count:
                logger.warning(
                    "reject_sentence_revisions: ожидали %s ревизий, нашли %s "
                    "(документ изменён вручную?)",
                    expected_count, len(to_reject),
                )
                return {**empty_fail, "reason": "count_mismatch"}

            # Reject в обратном порядке — чтобы позиции ещё-не-отвергнутых
            # ревизий не сдвигались.
            to_reject.sort(key=lambda x: x[0], reverse=True)
            for _, rev in to_reject:
                try:
                    rev.Reject()
                except Exception as e:
                    logger.error("Revision.Reject() failed: %s", e)
                    return {**empty_fail, "reason": "reject_exception"}

            if doc_len_before is not None:
                try:
                    delta = int(doc_com.Range().End) - doc_len_before
                except Exception:
                    delta = 0
            else:
                delta = 0

            _after_replacement(
                sentence, original_text, range_start, range_end_after, delta, all_sentences,
            )

            return {
                "ok": True, "delta": delta,
                "old_end": range_end_after, "rng_start": range_start,
            }
        except Exception as e:
            logger.error("reject_sentence_revisions error: %s", e)
            return {**empty_fail, "reason": "exception"}

    def navigate_to_range(self, doc: dict, start: int, end: int) -> bool:
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return False
        try:
            rng = doc_com.Range(start, end)
            rng.Select()
            # Передать фокус окну Word, чтобы выделение было видно
            hwnd = doc.get("hwnd")
            if hwnd:
                win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
                win32gui.SetForegroundWindow(hwnd)
            return True
        except Exception as e:
            logger.error("navigate_to_range error (%s, %s): %s", start, end, e)
            return False

    def set_revisions_mode(self, doc: dict, on: bool) -> bool:
        """Включить/выключить ОТОБРАЖЕНИЕ ревизий в Word.

        TrackRevisions (запись изменений) НЕ трогаем — она управляется временно
        при каждом apply_correction_to_doc. Здесь меняется только видимость.
        """
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            logger.warning("set_revisions_mode: нет com_object")
            return False

        on = bool(on)
        any_ok = False

        # Попытка 1: новое API RevisionsFilter (Word 2013+).
        # wdRevisionsViewFinal = 0 (показывать после правок), wdRevisionsViewOriginal = 1.
        # wdRevisionsMarkupAll = 1, wdRevisionsMarkupNone = 0,
        # wdRevisionsMarkupSimple = 2.
        try:
            rf = doc_com.ActiveWindow.View.RevisionsFilter
            rf.View = 0
            rf.Markup = 1 if on else 0
            any_ok = True
            logger.info("set_revisions_mode: RevisionsFilter.Markup=%s OK", rf.Markup)
        except Exception as e:
            logger.warning("set_revisions_mode: RevisionsFilter недоступен: %s", e)

        # Попытка 2: старое API ShowRevisionsAndComments.
        try:
            view = doc_com.ActiveWindow.View
            view.ShowRevisionsAndComments = on
            view.RevisionsView = 0  # wdRevisionsViewFinal
            any_ok = True
            logger.info("set_revisions_mode: ShowRevisionsAndComments=%s OK", on)
        except Exception as e:
            logger.warning("set_revisions_mode: ShowRevisionsAndComments недоступен: %s", e)

        # Дополнительно: типы ревизий, чтобы при on=True гарантированно показывались
        # вставки и удаления (иначе помечен только формат).
        if on:
            for attr in ("ShowInsertionsAndDeletions", "ShowFormatChanges"):
                try:
                    setattr(doc_com.ActiveWindow.View, attr, True)
                except Exception as e:
                    logger.debug("set_revisions_mode: %s недоступен: %s", attr, e)

        if not any_ok:
            logger.error("set_revisions_mode: ни одно API отображения не сработало")
            return False
        return True

    def highlight_range(self, doc: dict, start: int, end: int, color_index: int) -> bool:
        """Применить/снять заливку-выделение на диапазоне.

        color_index: 7 = wdYellow, -1 = wdNoHighlight.
        """
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return False
        try:
            rng = doc_com.Range(start, end)
            rng.HighlightColorIndex = color_index
            return True
        except Exception as e:
            logger.error("highlight_range error (%s, %s): %s", start, end, e)
            return False

    def highlight_text_fragment(
        self, doc: dict,
        search_start: int, search_end: int,
        text: str,
    ) -> tuple[int, int] | None:
        """Найти текст в диапазоне и применить жёлтую заливку.

        Использует Word Find — не зависит от позиционного смещения строки.

        Returns: (start, end) найденного диапазона, или None если не найдено.
        """
        doc_com = self.get_doc_com(doc)
        if not doc_com or not text or not text.strip():
            return None
        try:
            rng = doc_com.Range(search_start, search_end)
            find = rng.Find
            find.ClearFormatting()
            find.Text = text
            find.Forward = True
            find.Wrap = 0          # wdFindStop
            find.MatchCase = True
            find.MatchWholeWord = False
            if find.Execute():
                rng.HighlightColorIndex = 7  # wdYellow
                return (rng.Start, rng.End)
            return None
        except Exception as e:
            logger.error("highlight_text_fragment error (%s..%s, %r): %s",
                         search_start, search_end, text[:40], e)
            return None

    def get_icon(self) -> tuple[str, str]:
        return ("W", "#2B579A")

    # ─── Унификация форматирования ────────────────────────────────────────

    def analyze_format(
        self, doc: dict, *, cancel_event=None, progress_callback=None,
    ) -> Optional[dict]:
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return None
        try:
            stats = format_unifier.analyze_format(
                doc_com,
                skip_heading_styles=True,
                cancel_event=cancel_event,
                progress_callback=progress_callback,
            )
        except Exception as e:
            logger.error("analyze_format error: %s", e)
            return None
        return asdict(stats)

    def apply_format_uniform(
        self,
        doc: dict,
        attr: str,
        target_value: Any = None,
        *,
        cancel_event=None,
        progress_callback=None,
    ) -> Optional[dict]:
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return None

        coverage = None
        if target_value is None:
            try:
                stats = format_unifier.analyze_format(
                    doc_com,
                    skip_heading_styles=True,
                    cancel_event=cancel_event,
                    progress_callback=progress_callback,
                )
            except Exception as e:
                logger.error("apply_format_uniform: analyze failed: %s", e)
                return None
            if cancel_event is not None and cancel_event.is_set():
                return None
            if attr == "font_name":
                target_value = stats.dominant_font_name
            elif attr == "font_size":
                target_value = stats.dominant_font_size
            elif attr == "style":
                target_value = stats.dominant_style
            coverage = stats.coverage.get(attr)

        if target_value is None:
            logger.info("apply_format_uniform: target_value is None для attr=%s, skip", attr)
            return None

        try:
            snap = format_unifier.apply_unified_attribute(
                doc_com, attr, target_value,
                skip_heading_styles=True,
                cancel_event=cancel_event,
                progress_callback=progress_callback,
            )
        except Exception as e:
            logger.error("apply_format_uniform error: %s", e)
            return None

        snap_id = f"{id(doc_com)}:{attr}:{uuid.uuid4().hex[:8]}"
        self._snapshots[snap_id] = snap
        return {
            "snap_id": snap_id,
            "attr": attr,
            "target": target_value,
            "coverage": coverage,
            "segments": len(snap.ranges),
        }

    def restore_format(self, doc: dict, snapshot_descriptor: dict) -> bool:
        doc_com = self.get_doc_com(doc)
        if doc_com is None:
            return False
        snap_id = snapshot_descriptor.get("snap_id") if snapshot_descriptor else None
        if not snap_id:
            return False
        snap = self._snapshots.pop(snap_id, None)
        if snap is None:
            return False
        try:
            return format_unifier.restore_snapshot(doc_com, snap)
        except Exception as e:
            logger.error("restore_format error: %s", e)
            return False
