"""
Unit-тесты Apply/Undo через Mock-Word с TrackRevisions.

Проверяют, что:
- Apply корректно создаёт ревизии и возвращает revisions_marker.
- Reject корректно восстанавливает оригинальный текст для любого числа
  пословных правок (включая тот самый кейс с 3+ опечатками).
- Reject устойчив к ручному вмешательству пользователя (count_mismatch / not_found).
- Множественный Apply + Undo одного из предложений не ломает остальные.
- Маркеры ревизий корректно сдвигаются при последующих изменениях документа.

win32com не доступен на macOS — мокаем перед импортом word_provider.
"""
import os
import sys
import types
from unittest import mock

import pytest

# Корень проекта в sys.path
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

# Мокаем pywin32 ДО первого импорта word_provider
sys.modules.setdefault("win32com", types.ModuleType("win32com"))
sys.modules.setdefault("win32com.client", types.ModuleType("win32com.client"))
sys.modules.setdefault("win32gui", types.ModuleType("win32gui"))

# Теперь можно безопасно импортировать
from core.providers.word_provider import WordProvider  # noqa: E402


# ────────────────────────────────────────────────────────────────────────
# Mock-модель Word с TrackRevisions
# ────────────────────────────────────────────────────────────────────────


class MockRevision:
    """Mock Word Revision.

    kind = 'insertion' | 'deletion'.
    Хранит индексы в parent.chars (нельзя кешировать — после Reject позиции
    меняются). Идентифицируется по id-объекту (set membership).
    """

    def __init__(self, doc, kind, start, end):
        self._doc = doc
        self._kind = kind
        self._start_marker = object()  # стабильный идентификатор
        self._chars = doc.chars[start:end]
        # Каждому символу указываем владельца — для удаления.
        for c in self._chars:
            c.revision = self
            c.kind = "inserted" if kind == "insertion" else "deleted"

    @property
    def Range(self):
        # Найти текущие индексы наших символов в документе
        indices = [i for i, c in enumerate(self._doc.chars) if c.revision is self]
        if not indices:
            return MockRange(self._doc, 0, 0)
        return MockRange(self._doc, indices[0], indices[-1] + 1)

    def Reject(self):
        if self._kind == "insertion":
            # Удалить вставленные символы из документа.
            self._doc.chars = [c for c in self._doc.chars if c.revision is not self]
        else:  # deletion
            # Восстановить удалённые символы как обычные.
            for c in self._doc.chars:
                if c.revision is self:
                    c.kind = "normal"
                    c.revision = None
        # Снять регистрацию ревизии.
        if self in self._doc._revisions:
            self._doc._revisions.remove(self)


class MockChar:
    """Один символ в документе."""

    __slots__ = ("ch", "kind", "revision")

    def __init__(self, ch, kind="normal", revision=None):
        self.ch = ch
        self.kind = kind
        self.revision = revision


class MockRange:
    def __init__(self, doc, start, end):
        self._doc = doc
        self._start = start
        self._end = end

    @property
    def Start(self):
        return self._start

    @property
    def End(self):
        return self._end

    @property
    def Text(self):
        # Возвращает физический текст диапазона (включая зачёркнутые/вставленные).
        return "".join(c.ch for c in self._doc.chars[self._start:self._end])

    @Text.setter
    def Text(self, value):
        if not self._doc._track_revisions:
            # Обычная замена без TrackRevisions.
            del self._doc.chars[self._start:self._end]
            for i, ch in enumerate(value):
                self._doc.chars.insert(self._start + i, MockChar(ch))
            return

        # С TrackRevisions: помечаем существующие символы как deletion,
        # вставляем новые как insertion ПЕРЕД ними.
        old_start = self._start
        old_end = self._end

        new_chars = [MockChar(ch) for ch in value]
        ins_rev = MockRevision.__new__(MockRevision)
        ins_rev._doc = self._doc
        ins_rev._kind = "insertion"
        for c in new_chars:
            c.kind = "inserted"
            c.revision = ins_rev

        # Вставляем новые ПЕРЕД старыми.
        self._doc.chars[old_start:old_start] = new_chars

        # Помечаем старые (теперь сдвинутые) как deletion.
        del_rev = MockRevision.__new__(MockRevision)
        del_rev._doc = self._doc
        del_rev._kind = "deletion"
        for c in self._doc.chars[old_start + len(new_chars):old_end + len(new_chars)]:
            c.kind = "deleted"
            c.revision = del_rev

        self._doc._revisions.append(ins_rev)
        self._doc._revisions.append(del_rev)


class MockRevisionsCollection:
    def __init__(self, revs):
        self._revs = revs

    @property
    def Count(self):
        return len(self._revs)

    def Item(self, i):  # 1-based
        return self._revs[i - 1]


class MockDoc:
    """Mock Word.Document."""

    def __init__(self, text=""):
        self.chars = [MockChar(ch) for ch in text]
        self._track_revisions = False
        self._revisions: list[MockRevision] = []

    @property
    def TrackRevisions(self):
        return self._track_revisions

    @TrackRevisions.setter
    def TrackRevisions(self, v):
        self._track_revisions = bool(v)

    @property
    def Revisions(self):
        return MockRevisionsCollection(self._revisions)

    def Range(self, start=None, end=None):
        if start is None:
            return MockRange(self, 0, len(self.chars))
        if end is None:
            end = start
        return MockRange(self, start, end)

    def get_visible_text(self) -> str:
        """Финальный текст без deletion-ревизий — то, что увидит пользователь."""
        return "".join(c.ch for c in self.chars if c.kind != "deleted")

    def get_original_text(self) -> str:
        """Исходный текст: вставленные исключаем, удалённые возвращаем как обычные."""
        return "".join(c.ch for c in self.chars if c.kind != "inserted")


# ────────────────────────────────────────────────────────────────────────
# Хелперы
# ────────────────────────────────────────────────────────────────────────


def make_doc_with_sentence(text):
    """Создать MockDoc и dict-описание единственного предложения."""
    doc = MockDoc(text)
    sentence = {
        "index": 0,
        "word_sentence_index": 1,
        "range_start": 0,
        "range_end": len(text),
        "text": text,
        "in_table": False,
    }
    return doc, sentence


def do_apply(provider, doc, sentence, original, corrected):
    doc_dict = {"type": "word", "com_object": doc}
    return provider.replace_sentence_text_with_corrections(
        doc_dict, sentence, corrected, old_text=original,
        all_sentences=[sentence], track_revisions=True,
    )


def do_undo(provider, doc, sentence, marker, original):
    doc_dict = {"type": "word", "com_object": doc}
    return provider.reject_sentence_revisions(
        doc_dict, sentence, marker, original, all_sentences=[sentence],
    )


# ────────────────────────────────────────────────────────────────────────
# Тесты
# ────────────────────────────────────────────────────────────────────────


@pytest.fixture
def provider():
    return WordProvider()


def _assert_apply_undo_roundtrip(provider, doc, sent, original, corrected):
    """Полный roundtrip: Apply показывает corrected, Undo возвращает ровно original."""
    apply_result = do_apply(provider, doc, sent, original, corrected)
    assert apply_result["ok"], "Apply должен пройти"
    assert apply_result["revisions_marker"] is not None
    assert doc.get_visible_text() == corrected, (
        f"После Apply ожидали {corrected!r}, получили {doc.get_visible_text()!r}"
    )

    undo_result = do_undo(provider, doc, sent, apply_result["revisions_marker"], original)
    assert undo_result["ok"], f"Undo failed: {undo_result}"
    assert doc.get_visible_text() == original, (
        f"После Undo ожидали {original!r}, получили {doc.get_visible_text()!r}"
    )
    assert doc.Revisions.Count == 0, "После Undo все ревизии должны быть удалены"


def test_single_correction_same_length(provider):
    """1 правка той же длины: «кодга» → «когда»."""
    text = "Это слово кодга встречается."
    expected = "Это слово когда встречается."
    doc, sent = make_doc_with_sentence(text)
    _assert_apply_undo_roundtrip(provider, doc, sent, text, expected)


def test_single_correction_length_change(provider):
    """1 правка с изменением длины: «прпущены» → «пропущены» (+1)."""
    text = "Буквы могут быть прпущены здесь."
    expected = "Буквы могут быть пропущены здесь."
    doc, sent = make_doc_with_sentence(text)
    _assert_apply_undo_roundtrip(provider, doc, sent, text, expected)


def test_two_corrections_one_changes_length(provider):
    """2 правки в одном предложении (рабочий случай: «когда» правильное)."""
    text = "Буквы могут быть прпущены или изменены метсами."
    expected = "Буквы могут быть пропущены или изменены местами."
    doc, sent = make_doc_with_sentence(text)
    _assert_apply_undo_roundtrip(provider, doc, sent, text, expected)


def test_three_corrections_bug_repro(provider):
    """Исходный кейс бага: 3 правки разной длины в одном предложении.

    До фикса Undo возвращал «...кодга... ропущены... тсстами.» (мусор).
    После фикса должен вернуться РОВНО исходный текст.
    """
    text = (
        "Демонстрируется предложение с опечаткой, кодга при активном наборе "
        "буквы могут быть прпущены или изменены метсами."
    )
    expected = (
        "Демонстрируется предложение с опечаткой, когда при активном наборе "
        "буквы могут быть пропущены или изменены местами."
    )
    doc, sent = make_doc_with_sentence(text)
    _assert_apply_undo_roundtrip(provider, doc, sent, text, expected)


def test_punctuation_correction(provider):
    """Правка касается слова с пунктуацией в конце предложения."""
    text = "Документ заверен метсами."
    expected = "Документ заверен местами."
    doc, sent = make_doc_with_sentence(text)
    _assert_apply_undo_roundtrip(provider, doc, sent, text, expected)


def test_correction_at_start(provider):
    """Правка в самом начале предложения."""
    text = "Кодга мы встретимся, обсудим."
    expected = "Когда мы встретимся, обсудим."
    doc, sent = make_doc_with_sentence(text)
    _assert_apply_undo_roundtrip(provider, doc, sent, text, expected)


def test_correction_at_end(provider):
    """Правка в самом конце предложения."""
    text = "Изменены метсами"
    expected = "Изменены местами"
    doc, sent = make_doc_with_sentence(text)
    _assert_apply_undo_roundtrip(provider, doc, sent, text, expected)


def test_four_corrections_in_one_sentence(provider):
    """4 правки разной длины — стресс-тест."""
    text = "Сегодне мы рассмотрим вапросы, бугалтер закроет переод."
    expected = "Сегодня мы рассмотрим вопросы, бухгалтер закроет период."
    doc, sent = make_doc_with_sentence(text)
    _assert_apply_undo_roundtrip(provider, doc, sent, text, expected)


def test_double_undo_second_fails(provider):
    """Повторный Undo: первый успешен, второй — not_found (ревизии уже сняты)."""
    text = "Слово прпущены здесь."
    expected = "Слово пропущены здесь."
    doc, sent = make_doc_with_sentence(text)
    doc_dict = {"type": "word", "com_object": doc}

    apply_result = provider.replace_sentence_text_with_corrections(
        doc_dict, sent, expected, old_text=text,
        all_sentences=[sent], track_revisions=True,
    )
    marker = apply_result["revisions_marker"]

    first = provider.reject_sentence_revisions(doc_dict, sent, marker, text, [sent])
    assert first["ok"]

    # Второй вызов с тем же маркером — ревизий уже нет.
    second = provider.reject_sentence_revisions(doc_dict, sent, marker, text, [sent])
    assert not second["ok"]
    assert second["reason"] in {"not_found", "count_mismatch"}


def test_manual_accept_between_apply_and_undo(provider):
    """Если пользователь вручную принял часть ревизий — Undo вернёт count_mismatch."""
    text = "Слово прпущены здесь и метсами."
    expected = "Слово пропущены здесь и местами."
    doc, sent = make_doc_with_sentence(text)
    doc_dict = {"type": "word", "com_object": doc}

    apply_result = provider.replace_sentence_text_with_corrections(
        doc_dict, sent, expected, old_text=text,
        all_sentences=[sent], track_revisions=True,
    )
    marker = apply_result["revisions_marker"]

    # Имитируем ручной Accept: убираем одну ревизию из коллекции.
    # Accept для insertion = сделать «normal». Для deletion = удалить символы.
    doc._revisions.pop()  # любая ревизия

    undo_result = provider.reject_sentence_revisions(
        doc_dict, sent, marker, text, [sent],
    )
    assert not undo_result["ok"]
    assert undo_result["reason"] in {"count_mismatch", "not_found"}


def test_multiple_sentences_undo_one(provider):
    """Apply двух разных предложений, потом Undo только первого.
       Второе должно остаться applied (его ревизии не затронуты)."""
    full_text = "Слово прпущены здесь. Также метсами далее."
    sent_a_orig = "Слово прпущены здесь."
    sent_a_fix = "Слово пропущены здесь."
    sent_b_orig = "Также метсами далее."
    sent_b_fix = "Также местами далее."

    doc = MockDoc(full_text)
    sent_a = {
        "index": 0, "word_sentence_index": 1,
        "range_start": 0, "range_end": len(sent_a_orig),
        "text": sent_a_orig, "in_table": False,
    }
    sent_b = {
        "index": 1, "word_sentence_index": 2,
        "range_start": len(sent_a_orig) + 1,
        "range_end": len(full_text),
        "text": sent_b_orig, "in_table": False,
    }
    all_sents = [sent_a, sent_b]
    doc_dict = {"type": "word", "com_object": doc}

    # Apply A
    res_a = provider.replace_sentence_text_with_corrections(
        doc_dict, sent_a, sent_a_fix, old_text=sent_a_orig,
        all_sentences=all_sents, track_revisions=True,
    )
    assert res_a["ok"]
    # Apply B (на сдвинутых позициях — _after_replacement обновил sent_b)
    res_b = provider.replace_sentence_text_with_corrections(
        doc_dict, sent_b, sent_b_fix, old_text=sent_b_orig,
        all_sentences=all_sents, track_revisions=True,
    )
    assert res_b["ok"]

    assert doc.get_visible_text() == sent_a_fix + " " + sent_b_fix

    # Undo только A
    undo_a = provider.reject_sentence_revisions(
        doc_dict, sent_a, res_a["revisions_marker"], sent_a_orig, all_sents,
    )
    assert undo_a["ok"], f"Undo A failed: {undo_a}"
    assert doc.get_visible_text() == sent_a_orig + " " + sent_b_fix, (
        f"После Undo A: {doc.get_visible_text()!r}"
    )

    # B всё ещё применено и его маркер валиден — можно тоже откатить
    undo_b = provider.reject_sentence_revisions(
        doc_dict, sent_b, res_b["revisions_marker"], sent_b_orig, all_sents,
    )
    assert undo_b["ok"]
    assert doc.get_visible_text() == full_text


def test_shift_revisions_markers_in_engine():
    """Engine._shift_revisions_markers сдвигает маркеры check_results
    предложений, которые ПОСЛЕ изменённого участка."""
    # Импорт engine лениво — он тянет провайдеры со своими зависимостями
    from core.engine import Engine

    eng = Engine()
    eng.check_results = {
        0: {
            "original": "a", "corrected": "b", "state": "pending",
            "revisions_marker": {
                "range_start": 100, "range_end_after": 110,
                "count_before": 0, "count_after": 2,
            },
        },
        1: {
            "original": "x", "corrected": "y", "state": "pending",
            "revisions_marker": {
                "range_start": 50, "range_end_after": 60,
                "count_before": 0, "count_after": 1,
            },
        },
    }
    # Изменили что-то в районе old_end=70 на delta=+5 — должен сдвинуться
    # только маркер #0 (start=100 >= 70). Маркер #1 (start=50 < 70) — не должен.
    eng._shift_revisions_markers(exclude_index=2, old_end=70, delta=5)
    assert eng.check_results[0]["revisions_marker"]["range_start"] == 105
    assert eng.check_results[0]["revisions_marker"]["range_end_after"] == 115
    assert eng.check_results[1]["revisions_marker"]["range_start"] == 50
    assert eng.check_results[1]["revisions_marker"]["range_end_after"] == 60


def test_no_marker_returns_failure(provider):
    """reject_sentence_revisions без маркера возвращает no_marker."""
    doc = MockDoc("text")
    sentence = {"text": "text", "range_start": 0, "range_end": 4}
    res = provider.reject_sentence_revisions(
        {"type": "word", "com_object": doc}, sentence, None, "text", [sentence],
    )
    assert not res["ok"]
    assert res["reason"] == "no_marker"
