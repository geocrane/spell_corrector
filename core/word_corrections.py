"""
Разбиение пары (original, corrected) на пословные правки.

Используется для:
- генерации списка карточек в «Режиме ошибок» UI;
- пословного применения diff в Word (для маленьких ревизий при TrackRevisions=True).
"""

import difflib
import re

_WORD_CHAR_RE = re.compile(r'\w', re.UNICODE)

# Максимальная длина chunk-а карточки. Длиннее не расширяем — даже если правка
# покрывает несколько слов через пунктуацию.
_MAX_CHUNK_LEN = 80


def _is_word_char(ch: str) -> bool:
    return bool(_WORD_CHAR_RE.match(ch))


def _is_whitespace(ch: str) -> bool:
    return ch.isspace()


def _expand_left_to_token(text: str, pos: int) -> int:
    """Расширить позицию pos влево до начала текущего токена.

    Токен = последовательность \\w символов. Если pos указывает в середину
    слова — двигаем влево до пробела/пунктуации. Если pos на пробеле/пунктуации
    — не двигаем.
    """
    while pos > 0 and _is_word_char(text[pos - 1]):
        pos -= 1
    return pos


def _expand_right_to_token(text: str, pos: int) -> int:
    """Расширить позицию pos вправо до конца текущего токена."""
    n = len(text)
    while pos < n and _is_word_char(text[pos]):
        pos += 1
    return pos


def _expand_left_with_neighbour(text: str, pos: int) -> int:
    """Захватить соседнее слово слева целиком (через любые не-\\w символы)."""
    # Пропустить любые не-word символы до соседнего слова
    while pos > 0 and not _is_word_char(text[pos - 1]):
        pos -= 1
    # Захватить само слово
    while pos > 0 and _is_word_char(text[pos - 1]):
        pos -= 1
    return pos


def _expand_right_with_neighbour(text: str, pos: int) -> int:
    """Захватить соседнее слово справа целиком."""
    n = len(text)
    while pos < n and not _is_word_char(text[pos]):
        pos += 1
    while pos < n and _is_word_char(text[pos]):
        pos += 1
    return pos


def _change_is_only_non_word(text_a: str, text_b: str) -> bool:
    """True если в обоих текстах нет ни одного \\w символа (только пунктуация/пробел)."""
    return not _WORD_CHAR_RE.search(text_a) and not _WORD_CHAR_RE.search(text_b)


def _group_opcodes(opcodes, original, corrected):
    """Объединить подряд идущие не-equal opcodes, если разделяющий equal-фрагмент
    короткий (≤ 2 символов) и состоит только из пробелов/пунктуации.

    Возвращает список кортежей (i1, i2, j1, j2) — координат каждой группы.
    """
    groups = []
    current = None  # (i1, i2, j1, j2)

    for idx, (tag, i1, i2, j1, j2) in enumerate(opcodes):
        if tag == "equal":
            if current is None:
                continue
            eq_text = original[i1:i2]
            # Решение: продолжать ли расширять текущую группу через этот equal-разрыв
            short_separator = len(eq_text) <= 2 and not _WORD_CHAR_RE.search(eq_text)
            if short_separator:
                # Не закрываем группу — она расширится на следующий не-equal opcode.
                continue
            # Закрываем текущую группу.
            groups.append(current)
            current = None
        else:
            if current is None:
                current = [i1, i2, j1, j2]
            else:
                current[1] = i2
                current[3] = j2

    if current is not None:
        groups.append(current)

    return [tuple(g) for g in groups]


def _build_chunk(original, corrected, i1, i2, j1, j2):
    """Расширить (i1, i2) и (j1, j2) до границ соседних токенов.

    Если изменение касается только не-\\w символа (пунктуация/пробел) — захватить
    оба соседних слова в chunk, чтобы карточка была информативной.

    Возвращает (chunk_i1, chunk_i2, chunk_j1, chunk_j2) — границы chunk.
    """
    diff_orig = original[i1:i2]
    diff_corr = corrected[j1:j2]

    only_non_word = _change_is_only_non_word(diff_orig, diff_corr)

    if only_non_word:
        # Захватить полные соседние слова с обеих сторон
        chunk_i1 = _expand_left_with_neighbour(original, i1)
        chunk_i2 = _expand_right_with_neighbour(original, i2)
        chunk_j1 = _expand_left_with_neighbour(corrected, j1)
        chunk_j2 = _expand_right_with_neighbour(corrected, j2)
    else:
        # Расширить до границ текущего слова
        chunk_i1 = _expand_left_to_token(original, i1)
        chunk_i2 = _expand_right_to_token(original, i2)
        chunk_j1 = _expand_left_to_token(corrected, j1)
        chunk_j2 = _expand_right_to_token(corrected, j2)

    # Ограничить длину chunk
    if chunk_i2 - chunk_i1 > _MAX_CHUNK_LEN:
        chunk_i1 = max(chunk_i1, i1 - _MAX_CHUNK_LEN // 2)
        chunk_i2 = min(chunk_i2, chunk_i1 + _MAX_CHUNK_LEN)
    if chunk_j2 - chunk_j1 > _MAX_CHUNK_LEN:
        chunk_j1 = max(chunk_j1, j1 - _MAX_CHUNK_LEN // 2)
        chunk_j2 = min(chunk_j2, chunk_j1 + _MAX_CHUNK_LEN)

    return chunk_i1, chunk_i2, chunk_j1, chunk_j2


def _merge_overlapping_chunks(chunks):
    """Слить пересекающиеся chunk-и (если расширение двух соседних правок их перекрыло)."""
    if not chunks:
        return chunks

    merged = [chunks[0]]
    for nxt in chunks[1:]:
        prev = merged[-1]
        # prev: (i1, i2, j1, j2, chunk_i1, chunk_i2, chunk_j1, chunk_j2)
        if nxt[4] <= prev[5] or nxt[6] <= prev[7]:
            # Пересечение — объединяем
            merged[-1] = (
                prev[0], nxt[1], prev[2], nxt[3],
                min(prev[4], nxt[4]), max(prev[5], nxt[5]),
                min(prev[6], nxt[6]), max(prev[7], nxt[7]),
            )
        else:
            merged.append(nxt)

    return merged


def build_word_corrections(
    original: str,
    corrected: str,
    sentence_index: int,
    sentence_range_start: int,
) -> list[dict]:
    """Сгенерировать список пословных правок из пары (original, corrected).

    Args:
        original: Исходный текст предложения.
        corrected: Исправленный текст предложения.
        sentence_index: Индекс предложения в списке (для UI и привязки).
        sentence_range_start: Абсолютная позиция начала предложения в Word.Document.Range.
            Используется для вычисления word_range_start/end после применения правки.

    Returns:
        Список dict с полями:
            id (str), sentence_index (int), sub_index (int),
            original_chunk (str), corrected_chunk (str),
            i1_in_original (int), i2_in_original (int),
            j1_in_corrected (int), j2_in_corrected (int),
            chunk_offset_in_original (int), chunk_offset_in_corrected (int),
            char_range_in_original (tuple), char_range_in_corrected (tuple),
            word_range_start (int), word_range_end (int).
    """
    if original == corrected:
        return []

    opcodes = difflib.SequenceMatcher(None, original, corrected).get_opcodes()
    groups = _group_opcodes(opcodes, original, corrected)

    expanded = []
    for i1, i2, j1, j2 in groups:
        ci1, ci2, cj1, cj2 = _build_chunk(original, corrected, i1, i2, j1, j2)
        expanded.append((i1, i2, j1, j2, ci1, ci2, cj1, cj2))

    expanded = _merge_overlapping_chunks(expanded)

    corrections = []
    for sub_idx, (i1, i2, j1, j2, ci1, ci2, cj1, cj2) in enumerate(expanded):
        original_chunk = original[ci1:ci2]
        corrected_chunk = corrected[cj1:cj2]

        char_range_in_original = (i1 - ci1, i2 - ci1)
        char_range_in_corrected = (j1 - cj1, j2 - cj1)

        word_range_start = sentence_range_start + cj1
        word_range_end = sentence_range_start + cj2

        corrections.append({
            "id": f"{sentence_index}:{sub_idx}",
            "sentence_index": sentence_index,
            "sub_index": sub_idx,
            "original_chunk": original_chunk,
            "corrected_chunk": corrected_chunk,
            "i1_in_original": i1,
            "i2_in_original": i2,
            "j1_in_corrected": j1,
            "j2_in_corrected": j2,
            "chunk_offset_in_original": ci1,
            "chunk_offset_in_corrected": cj1,
            "char_range_in_original": char_range_in_original,
            "char_range_in_corrected": char_range_in_corrected,
            "word_range_start": word_range_start,
            "word_range_end": word_range_end,
        })

    return corrections
