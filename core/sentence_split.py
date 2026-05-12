"""
Общий sentence-splitter с поддержкой русских сокращений.

Используется и WordProvider (импортирует ABBREV_PATTERN, starts_with_lower),
и ExcelProvider (использует split_into_sentences для разбиения текста ячейки).

Word имеет встроенное API doc.Sentences, поэтому split_into_sentences()
там не нужен. Excel такого API не имеет — текст ячейки нужно разбивать
вручную, учитывая русские сокращения (т.д., т.е., г., руб. и т.п.).
"""

import re

ABBREV_PATTERN = re.compile(
    r'(?:'
    r'г\.г|гг|тыс|млн|млрд|трлн|руб|коп|ед|шт'
    r'|т\.д|т\.п|т\.е|т\.к|пр|др'
    r'|ул|корп|кв|стр|пер|обл'
    r'|и\.о|врио|проф|доц|акад'
    r'|рис|табл|гл|пп|разд|см|ср|напр|прим'
    r'|(?<![а-яА-Яa-zA-ZёЁ])(?:г|д|п|ч|ст|им|[А-ЯA-Z])'
    r')\.$'
)

_LETTER_RE = re.compile(r'[A-Za-zА-Яа-яЁё]')
_SENT_END_RE = re.compile(r'[.!?…]+')


def starts_with_lower(text: str) -> bool:
    """True если первый буквенный символ — строчная буква."""
    for ch in text:
        if ch.isalpha():
            return ch.islower()
        if not ch.isspace():
            return False
    return False


def has_letters(text: str) -> bool:
    """True если в строке есть хотя бы один буквенный символ."""
    return bool(_LETTER_RE.search(text))


def split_into_sentences(text: str) -> list[tuple[int, int, str]]:
    """Разбить text на предложения с учётом русских сокращений.

    Перевод строки (\\n) — жёсткий разделитель: между фрагментами,
    разделёнными переносом, склейка не делается.

    Args:
        text: Исходный текст (например, содержимое одной ячейки Excel).

    Returns:
        Список кортежей (start_offset, end_offset, sub_text) — позиции
        и текст каждого предложения в исходной строке. sub_text — без
        внешних пробелов; (start, end) — границы в text, включая
        концевую пунктуацию (после .stop()), но без trailing-пробелов.
    """
    if not text:
        return []

    raw_fragments: list[tuple[int, int]] = []
    for line_match in re.finditer(r'[^\n]+', text):
        line_start = line_match.start()
        line_text = line_match.group(0)

        last = 0
        for m in _SENT_END_RE.finditer(line_text):
            end = m.end()
            if end < len(line_text) and not line_text[end].isspace():
                continue
            raw_fragments.append((line_start + last, line_start + end))
            last = end
        if last < len(line_text):
            raw_fragments.append((line_start + last, line_start + len(line_text)))

    if not raw_fragments:
        return []

    merged: list[tuple[int, int]] = [raw_fragments[0]]
    for nxt_start, nxt_end in raw_fragments[1:]:
        cur_start, cur_end = merged[-1]
        if '\n' in text[cur_end:nxt_start]:
            merged.append((nxt_start, nxt_end))
            continue
        cur_text = text[cur_start:cur_end]
        nxt_text = text[nxt_start:nxt_end]
        if starts_with_lower(nxt_text) or ABBREV_PATTERN.search(cur_text.rstrip()):
            merged[-1] = (cur_start, nxt_end)
        else:
            merged.append((nxt_start, nxt_end))

    result: list[tuple[int, int, str]] = []
    for start, end in merged:
        sub = text[start:end]
        lstrip_len = len(sub) - len(sub.lstrip())
        rstrip_len = len(sub) - len(sub.rstrip())
        start2 = start + lstrip_len
        end2 = end - rstrip_len
        if start2 >= end2:
            continue
        sub_text = text[start2:end2]
        if not has_letters(sub_text):
            continue
        result.append((start2, end2, sub_text))

    return result
