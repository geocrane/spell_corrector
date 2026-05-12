"""
Унификация форматирования документа Word.

Анализирует преобладающий шрифт / размер / стиль (bold+italic) основного
текста, применяет выбранное значение ко всем non-heading параграфам и
позволяет откатить изменение через in-memory snapshot.

Заголовки (OutlineLevel < 10 или Style.NameLocal с известным префиксом)
пропускаются — иерархия документа неприкосновенна.

Модуль не знает об Engine/UI: работает только с COM-объектом Word.Document.
"""

import contextlib
import logging
from collections import Counter
from dataclasses import dataclass, field
from typing import Any, Callable, Iterator, Optional

logger = logging.getLogger("core.format_unifier")


# Word-константы. win32com.client.constants не используем — заполняется только
# после gencache.EnsureModule, в проекте gencache не вызывается.
WD_STYLE_TYPE_PARAGRAPH = 1
WD_UNDEFINED = 9999999           # возвращается Word для смешанных значений
WD_OUTLINE_BODY = 10             # 1..9 — Heading 1..9, 10 — Body Text

HEADING_NAMES_PREFIXES = (
    "Заголов", "Загл", "Назван", "Title", "Heading", "TOC",
)

STYLE_NORMAL = "normal"
STYLE_BOLD = "bold"
STYLE_ITALIC = "italic"
STYLE_BOLD_ITALIC = "bold_italic"

# Защита от патологий: на каком max-уровне рекурсии прекращаем split.
_MAX_SPLIT_DEPTH = 24


# ─── Низкоуровневые хелперы ────────────────────────────────────────────────

def _read_attr(font, attr_name: str) -> Any:
    """Читать атрибут Font.<attr_name>, проглатывая COM-ошибки."""
    try:
        return getattr(font, attr_name)
    except Exception:
        return None


def _is_undefined(val) -> bool:
    """True если значение шрифта «смешанное» / неопределённое."""
    if val is None:
        return True
    if isinstance(val, str):
        return val == ""
    return val == WD_UNDEFINED


def _norm_bool(val) -> bool:
    """Word возвращает -1/True для активного флага, 0/False иначе."""
    return val is True or val == -1


def _style_key(bold_val, italic_val) -> str:
    """Свернуть пару (bold, italic) в строковый ключ."""
    b = _norm_bool(bold_val)
    i = _norm_bool(italic_val)
    if b and i:
        return STYLE_BOLD_ITALIC
    if b:
        return STYLE_BOLD
    if i:
        return STYLE_ITALIC
    return STYLE_NORMAL


def _segments_match(attr: str, current_val, target_val) -> bool:
    if attr == "font_size":
        try:
            return abs(float(current_val) - float(target_val)) < 1e-6
        except (TypeError, ValueError):
            return False
    return current_val == target_val


# ─── Структуры данных ──────────────────────────────────────────────────────

@dataclass
class FormatStats:
    """Результат analyze_format: преобладающие значения и покрытие."""
    dominant_font_name: Optional[str] = None
    dominant_font_size: Optional[float] = None
    dominant_style: str = STYLE_NORMAL
    coverage: dict = field(default_factory=dict)
    total_chars: int = 0


@dataclass
class Snapshot:
    """In-memory снимок исходного форматирования для одного атрибута."""
    doc_id: int
    attr: str                              # "font_name" | "font_size" | "style"
    target_value: Any
    doc_length: int                        # doc.Range().End на момент apply
    chars_count: int                       # doc.Characters.Count
    ranges: list = field(default_factory=list)  # [(start, end, orig_fmt_dict)]
    track_revisions_was: Optional[bool] = None


# ─── Heading-detection ─────────────────────────────────────────────────────

def is_heading_paragraph(paragraph) -> bool:
    """True для Heading 1..9 / Title / TOC.

    Иерархия проверок (от самого надёжного):
    1. paragraph.OutlineLevel < 10 — числовое, не локализовано.
    2. Style.NameLocal начинается с известного префикса (страховка для документов
       без выставленного OutlineLevel).
    """
    try:
        outline = int(paragraph.OutlineLevel)
        if outline < WD_OUTLINE_BODY:
            return True
    except Exception:
        pass
    try:
        name = paragraph.Style.NameLocal or ""
        for prefix in HEADING_NAMES_PREFIXES:
            if name.startswith(prefix):
                return True
    except Exception:
        pass
    return False


# ─── TrackRevisions ────────────────────────────────────────────────────────

@contextlib.contextmanager
def disable_track_revisions(doc_com):
    """Временно выключить TrackRevisions, восстановить в finally.

    Если у документа нет атрибута TrackRevisions (теоретически возможно для
    необычных COM-объектов) — yield с None, ничего не меняем.
    """
    prev = None
    try:
        try:
            prev = bool(doc_com.TrackRevisions)
        except Exception:
            prev = None
        if prev is True:
            try:
                doc_com.TrackRevisions = False
            except Exception as e:
                logger.debug("disable_track_revisions: не удалось выключить: %s", e)
        yield prev
    finally:
        if prev is True:
            try:
                doc_com.TrackRevisions = True
            except Exception as e:
                logger.debug("disable_track_revisions: не удалось восстановить: %s", e)


# ─── Чтение атрибутов и сегментация ────────────────────────────────────────

def _read_font_name(font) -> Any:
    return _read_attr(font, "Name")


def _read_font_size(font) -> Any:
    return _read_attr(font, "Size")


def _read_style(font) -> Optional[str]:
    """Вернуть строковый ключ стиля или None, если bold/italic смешаны."""
    b = _read_attr(font, "Bold")
    i = _read_attr(font, "Italic")
    if _is_undefined(b) or _is_undefined(i):
        return None
    return _style_key(b, i)


def _read_combined(font) -> Optional[dict]:
    """Прочитать сразу name/size/style; None если хоть один атрибут смешан."""
    name = _read_attr(font, "Name")
    size = _read_attr(font, "Size")
    bold = _read_attr(font, "Bold")
    italic = _read_attr(font, "Italic")
    if _is_undefined(name) or _is_undefined(size) or _is_undefined(bold) or _is_undefined(italic):
        return None
    try:
        size_f = float(size)
    except (TypeError, ValueError):
        size_f = None
    return {
        "name": name,
        "size": size_f,
        "style": _style_key(bold, italic),
    }


def _split_segment(
    doc_com, start: int, end: int, read_value: Callable[[Any], Any], depth: int = 0,
) -> Iterator[tuple]:
    """Рекурсивно разбить [start, end] на однородные по read_value сегменты.

    Yield (start, end, value). Если значение «смешанное» (read_value вернул
    None или WD_UNDEFINED) — делим пополам, иначе отдаём как есть.
    """
    if end <= start:
        return
    try:
        sub_rng = doc_com.Range(start, end)
        val = read_value(sub_rng.Font)
    except Exception as e:
        logger.debug("_split_segment [%d:%d] read error: %s", start, end, e)
        return

    is_undef = val is None or _is_undefined(val)
    if not is_undef or end - start <= 1 or depth >= _MAX_SPLIT_DEPTH:
        yield (start, end, val)
        return

    mid = (start + end) // 2
    if mid <= start or mid >= end:
        yield (start, end, val)
        return
    yield from _split_segment(doc_com, start, mid, read_value, depth + 1)
    yield from _split_segment(doc_com, mid, end, read_value, depth + 1)


# ─── Анализ форматирования ─────────────────────────────────────────────────

def analyze_format(doc_com, *, skip_heading_styles: bool = True) -> FormatStats:
    """Посчитать преобладающие шрифт/размер/стиль в основном тексте.

    Обходит doc.Paragraphs, пропускает heading-параграфы, аккумулирует Counter
    по числу символов в однородных сегментах. Один проход через combined-reader
    минимизирует число COM-вызовов.
    """
    font_counter: Counter = Counter()
    size_counter: Counter = Counter()
    style_counter: Counter = Counter()
    total_chars = 0

    try:
        paragraphs = doc_com.Paragraphs
        para_count = int(paragraphs.Count)
    except Exception as e:
        logger.warning("analyze_format: Paragraphs error: %s", e)
        return FormatStats()

    for i in range(1, para_count + 1):
        try:
            p = paragraphs.Item(i)
        except Exception:
            continue
        if skip_heading_styles and is_heading_paragraph(p):
            continue
        try:
            rng = p.Range
            p_start = int(rng.Start)
            p_end = int(rng.End)
        except Exception:
            continue
        if p_end <= p_start:
            continue

        for (s, e, val) in _split_segment(doc_com, p_start, p_end, _read_combined):
            if val is None:
                continue
            weight = e - s
            if weight <= 0:
                continue
            total_chars += weight
            font_counter[val["name"]] += weight
            if val["size"] is not None:
                size_counter[val["size"]] += weight
            style_counter[val["style"]] += weight

    stats = FormatStats(total_chars=total_chars)
    if font_counter:
        top_name, top_n = font_counter.most_common(1)[0]
        stats.dominant_font_name = top_name
        stats.coverage["font_name"] = top_n / max(1, sum(font_counter.values()))
    if size_counter:
        top_size, top_s = size_counter.most_common(1)[0]
        stats.dominant_font_size = top_size
        stats.coverage["font_size"] = top_s / max(1, sum(size_counter.values()))
    if style_counter:
        top_style, top_st = style_counter.most_common(1)[0]
        stats.dominant_style = top_style
        stats.coverage["style"] = top_st / max(1, sum(style_counter.values()))
    return stats


# ─── Применение и откат ────────────────────────────────────────────────────

def _attr_reader(attr: str) -> Callable[[Any], Any]:
    if attr == "font_name":
        return _read_font_name
    if attr == "font_size":
        return _read_font_size
    if attr == "style":
        return _read_style
    raise ValueError(f"unknown attr: {attr}")


def _snapshot_segment_fmt(doc_com, start: int, end: int, attr: str) -> dict:
    try:
        rng = doc_com.Range(start, end)
        font = rng.Font
        if attr == "font_name":
            return {"name": _read_attr(font, "Name")}
        if attr == "font_size":
            return {"size": _read_attr(font, "Size")}
        if attr == "style":
            return {
                "bold": _norm_bool(_read_attr(font, "Bold")),
                "italic": _norm_bool(_read_attr(font, "Italic")),
            }
    except Exception as e:
        logger.debug("_snapshot_segment_fmt [%d:%d] %s error: %s", start, end, attr, e)
    return {}


def _apply_segment_attr(doc_com, start: int, end: int, attr: str, target) -> None:
    try:
        rng = doc_com.Range(start, end)
        font = rng.Font
        if attr == "font_name":
            font.Name = str(target)
        elif attr == "font_size":
            font.Size = float(target)
        elif attr == "style":
            bold = target in (STYLE_BOLD, STYLE_BOLD_ITALIC)
            italic = target in (STYLE_ITALIC, STYLE_BOLD_ITALIC)
            font.Bold = -1 if bold else 0
            font.Italic = -1 if italic else 0
    except Exception as e:
        logger.debug("_apply_segment_attr [%d:%d] %s=%r error: %s", start, end, attr, target, e)


def apply_unified_attribute(
    doc_com, attr: str, target_value: Any, *, skip_heading_styles: bool = True,
) -> Snapshot:
    """Применить target_value к атрибуту attr во всех non-heading параграфах.

    Возвращает Snapshot для последующего restore_snapshot. Внутри
    disable_track_revisions, чтобы не засорять историю ревизиями
    «изменён шрифт».
    """
    if attr not in ("font_name", "font_size", "style"):
        raise ValueError(f"unknown attr: {attr}")
    if target_value is None:
        raise ValueError("target_value is None")

    try:
        doc_length = int(doc_com.Range().End)
    except Exception:
        doc_length = 0
    try:
        chars_count = int(doc_com.Characters.Count)
    except Exception:
        chars_count = 0

    snap = Snapshot(
        doc_id=id(doc_com),
        attr=attr,
        target_value=target_value,
        doc_length=doc_length,
        chars_count=chars_count,
    )

    reader = _attr_reader(attr)

    with disable_track_revisions(doc_com) as prev_track:
        snap.track_revisions_was = prev_track
        try:
            paragraphs = doc_com.Paragraphs
            para_count = int(paragraphs.Count)
        except Exception as e:
            logger.warning("apply_unified_attribute: Paragraphs error: %s", e)
            return snap

        for i in range(1, para_count + 1):
            try:
                p = paragraphs.Item(i)
            except Exception:
                continue
            if skip_heading_styles and is_heading_paragraph(p):
                continue
            try:
                rng = p.Range
                p_start = int(rng.Start)
                p_end = int(rng.End)
            except Exception:
                continue
            if p_end <= p_start:
                continue

            for (s, e, val) in _split_segment(doc_com, p_start, p_end, reader):
                if val is None:
                    continue
                if _segments_match(attr, val, target_value):
                    continue
                orig_fmt = _snapshot_segment_fmt(doc_com, s, e, attr)
                snap.ranges.append((s, e, orig_fmt))
                _apply_segment_attr(doc_com, s, e, attr, target_value)

    logger.info(
        "apply_unified_attribute: attr=%s target=%r segments=%d",
        attr, target_value, len(snap.ranges),
    )
    return snap


def restore_snapshot(doc_com, snap: Snapshot) -> bool:
    """Восстановить исходные значения из snapshot.

    Возвращает False, если документ был отредактирован после apply (длина
    Range().End изменилась) — снимок невалиден, ничего не трогаем.
    """
    try:
        cur_length = int(doc_com.Range().End)
    except Exception:
        return False
    if cur_length != snap.doc_length:
        logger.info(
            "restore_snapshot: документ изменился (%s → %s), snapshot инвалидирован",
            snap.doc_length, cur_length,
        )
        return False

    with disable_track_revisions(doc_com):
        for (start, end, orig_fmt) in snap.ranges:
            try:
                rng = doc_com.Range(start, end)
                font = rng.Font
                if snap.attr == "font_name":
                    name = orig_fmt.get("name")
                    if name and not _is_undefined(name):
                        font.Name = str(name)
                elif snap.attr == "font_size":
                    size = orig_fmt.get("size")
                    if size is not None and not _is_undefined(size):
                        try:
                            font.Size = float(size)
                        except (TypeError, ValueError):
                            pass
                elif snap.attr == "style":
                    bold = bool(orig_fmt.get("bold", False))
                    italic = bool(orig_fmt.get("italic", False))
                    font.Bold = -1 if bold else 0
                    font.Italic = -1 if italic else 0
            except Exception as e:
                logger.debug("restore_snapshot [%d:%d] error: %s", start, end, e)
                continue
    return True
