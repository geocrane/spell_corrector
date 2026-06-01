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


def extract_paragraphs_from_docx(path: str) -> list[dict]:
    """Извлечь абзацы из .docx-файла снимка через Word COM.

    Использует тот же метод, что и extract_paragraphs(), чтобы тексты
    были идентично форматированы и SequenceMatcher работал корректно.
    Если файл уже открыт в Word — читает его напрямую (без открытия/закрытия).

    Returns:
        list of {text, range_start, range_end, comparison_key}
        range_start/range_end = None (снимок — для навигации не используется).
    """
    import os  # noqa: PLC0415
    import win32com.client  # noqa: PLC0415

    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception:
        return []

    norm_path = os.path.normcase(os.path.abspath(path))

    # Если файл уже открыт в Word — читаем без открытия/закрытия
    for d in word.Documents:
        try:
            if os.path.normcase(os.path.abspath(d.FullName)) == norm_path:
                paras = extract_paragraphs(d)
                # Обнуляем позиции — снимок не используется для навигации
                for p in paras:
                    p["range_start"] = None
                    p["range_end"] = None
                return paras
        except Exception:
            continue

    # Открываем скрыто, читаем, закрываем
    doc = None
    try:
        doc = word.Documents.Open(
            path,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        paras = extract_paragraphs(doc)
        for p in paras:
            p["range_start"] = None
            p["range_end"] = None
        return paras
    except Exception:
        return []
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass


def compute_similarity_ratio(paras_a: list[dict], paras_b: list[dict]) -> float:
    """Вычислить коэффициент схожести двух наборов абзацев (0.0 – 1.0).

    Использует пословное сравнение — менее чувствительно к форматированию,
    лучше отделяет «много правок» от «совсем другой текст».
    """
    if not paras_a or not paras_b:
        return 0.0
    words_a = " ".join(p["comparison_key"] for p in paras_a).split()
    words_b = " ".join(p["comparison_key"] for p in paras_b).split()
    if not words_a or not words_b:
        return 0.0
    return difflib.SequenceMatcher(None, words_a, words_b, autojunk=False).ratio()


def compute_diff(paras_a: list[dict], paras_b: list[dict]) -> list[dict]:
    """Сравнить два списка абзацев, вернуть только изменённые блоки.

    Для сопоставления использует `comparison_key` (без нумерации), что позволяет
    корректно определять удалённые/добавленные пункты в нумерованных списках.

    Для блоков типа "delete" добавляет anchor_range_start_b / anchor_range_end_b —
    позицию ближайшего соседнего абзаца в текущем (B) документе, что позволяет
    прокрутить текущий документ к месту, где был удалённый абзац.

    Returns:
        list of {
            tag: "replace"|"insert"|"delete",
            text_a: str,
            text_b: str,
            range_start_a: int|None,
            range_end_a:   int|None,
            range_start_b: int|None,
            range_end_b:   int|None,
            anchor_range_start_b: int|None,  # только для delete
            anchor_range_end_b:   int|None,  # только для delete
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

        # Для удалённых блоков: ближайший якорь в текущем (B) документе
        anchor_start_b = None
        anchor_end_b = None
        if tag == "delete":
            # j1 == j2, удаление произошло между j1-1 и j1 в B
            if j1 > 0:
                anchor_start_b = paras_b[j1 - 1]["range_start"]
                anchor_end_b = paras_b[j1 - 1]["range_end"]
            elif paras_b:
                anchor_start_b = paras_b[0]["range_start"]
                anchor_end_b = paras_b[0]["range_end"]

        result.append({
            "tag": tag,
            "text_a": text_a,
            "text_b": text_b,
            "range_start_a": range_start_a,
            "range_end_a": range_end_a,
            "range_start_b": range_start_b,
            "range_end_b": range_end_b,
            "anchor_range_start_b": anchor_start_b,
            "anchor_range_end_b": anchor_end_b,
            # Исходные абзацы для точного расчёта позиций при подсветке
            "paras_a": block_a,
            "paras_b": block_b,
        })

    return result
