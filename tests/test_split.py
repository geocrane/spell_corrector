"""Юнит-тесты core.sentence_split.split_into_sentences."""
import pytest

from core.sentence_split import split_into_sentences

from tests.corpus import SPLIT_CASES


@pytest.mark.parametrize("text,expected_count,expected_chunks", SPLIT_CASES,
                         ids=[f"split_{i}" for i in range(len(SPLIT_CASES))])
def test_split(text, expected_count, expected_chunks):
    result = split_into_sentences(text)
    chunks = [c[2] for c in result]
    assert len(chunks) == expected_count, (
        f"Ожидалось {expected_count} предложений, получено {len(chunks)}: {chunks!r}"
    )
    if expected_chunks is not None:
        assert chunks == expected_chunks


def test_split_offsets_consistent():
    """Проверяем, что start/end offset-ы корректно указывают на текст в исходнике."""
    text = "Первое предложение. Второе предложение. Третье."
    result = split_into_sentences(text)
    for start, end, sub in result:
        assert text[start:end] == sub


def test_split_empty():
    assert split_into_sentences("") == []
    assert split_into_sentences("    \n\n  ") == []


def test_split_only_punct():
    """Фрагменты без букв должны отфильтровываться."""
    result = split_into_sentences("... !!! 123 ???")
    chunks = [c[2] for c in result]
    # 123 без букв — должно отфильтроваться
    assert all(any(c.isalpha() for c in chunk) for chunk in chunks)


def test_split_multiline():
    text = "Первое.\nВторое.\nТретье."
    result = split_into_sentences(text)
    assert len(result) == 3
    assert [c[2] for c in result] == ["Первое.", "Второе.", "Третье."]


def test_split_abbrev_pattern():
    """Регрессия: г., тыс., млн., руб., и т.д. — не должны рвать."""
    cases = [
        ("В 2020 г. компания выросла. Дальше всё хорошо.", 2),
        ("Бюджет 5 млн. руб. согласован.", 1),
        ("Это т.д. подобное. Конец.", 2),
        ("Иван г. Москва. Это адрес.", 2),
    ]
    for text, expected in cases:
        result = split_into_sentences(text)
        assert len(result) == expected, (
            f"{text!r}: ожидалось {expected}, получено {len(result)}: {[c[2] for c in result]!r}"
        )


def test_split_initials():
    """Регрессия: А.С. Пушкин не должен рваться на 2 предложения."""
    text = "А.С. Пушкин родился в 1799 году."
    result = split_into_sentences(text)
    assert len(result) == 1
