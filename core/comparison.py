"""
Сравнение двух Word-документов на уровне абзацев.
"""

import difflib
import re

_WORD_SPECIAL_RE = re.compile(r"[\x07\x0b\x0c\x1e\x1f\x01-\x06\x08\x0e-\x1d]")

# Ведущая нумерация / маркеры списка, которые Word авто-обновляет при
# добавлении/удалении пунктов. Убираем их при построении ключа сравнения,
# чтобы SequenceMatcher видел "delete" там, где сдвинулись номера.
_LIST_PREFIX_RE = re.compile(
    r"^(?:\d+[.)]\s+|[a-zA-Zа-яА-Я][.)]\s+|[•·–—\-\*–—▪●•]\s*)",
    re.UNICODE,
)


def _strip_special(text: str) -> str:
    return _WORD_SPECIAL_RE.sub("", text).strip()


def _comparison_key(text: str) -> str:
    """Убирает ведущую нумерацию/маркер списка для точного сопоставления."""
    return _LIST_PREFIX_RE.sub("", text).strip()


def extract_paragraphs(doc_com) -> list[dict]:
    """Извлечь абзацы из COM-объекта Word.Document.

    Returns:
        list of {text, range_start, range_end, comparison_key}
    """
    result = []
    try:
        count = doc_com.Paragraphs.Count
        for i in range(1, count + 1):
            try:
                para = doc_com.Paragraphs.Item(i)
                text = _strip_special(para.Range.Text or "")
                if not text:
                    continue
                result.append({
                    "text": text,
                    "range_start": para.Range.Start,
                    "range_end": para.Range.End,
                    "comparison_key": _comparison_key(text),
                })
            except Exception:
                continue
    except Exception:
        pass
    return result


def compute_diff(paras_a: list[dict], paras_b: list[dict]) -> list[dict]:
    """Сравнить два списка абзацев, вернуть только изменённые блоки.

    Для сопоставления использует `comparison_key` (без нумерации), что позволяет
    корректно определять удалённые/добавленные пункты в нумерованных списках.

    Returns:
        list of {
            tag: "replace"|"insert"|"delete",
            text_a: str,
            text_b: str,
            range_start_a: int|None,
            range_end_a:   int|None,
            range_start_b: int|None,
            range_end_b:   int|None,
        }
    """
    keys_a = [p["comparison_key"] for p in paras_a]
    keys_b = [p["comparison_key"] for p in paras_b]

    matcher = difflib.SequenceMatcher(None, keys_a, keys_b, autojunk=False)
    result = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue

        block_a = paras_a[i1:i2]
        block_b = paras_b[j1:j2]

        text_a = "\n".join(p["text"] for p in block_a)
        text_b = "\n".join(p["text"] for p in block_b)

        range_start_a = block_a[0]["range_start"] if block_a else None
        range_end_a = block_a[-1]["range_end"] if block_a else None
        range_start_b = block_b[0]["range_start"] if block_b else None
        range_end_b = block_b[-1]["range_end"] if block_b else None

        result.append({
            "tag": tag,
            "text_a": text_a,
            "text_b": text_b,
            "range_start_a": range_start_a,
            "range_end_a": range_end_a,
            "range_start_b": range_start_b,
            "range_end_b": range_end_b,
        })

    return result
