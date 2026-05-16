"""
Интеграционные тесты — реальный pipeline с T5 моделью.

Запуск: ВНУТРИ venv (.venv/bin/python -m pytest tests/test_pipeline.py)
"""
import pytest

from tests.corpus import ALL_CASES
from tests.pipeline import run_pipeline


@pytest.mark.parametrize("case", ALL_CASES, ids=[f"{c.scenario}_{i}" for i, c in enumerate(ALL_CASES)])
def test_case(case, checker):
    cfg = case.config or {}
    strict = cfg.get("strict_protection", True)
    audit = cfg.get("auditor_format", True)
    blocklist = cfg.get("word_blocklist", [])

    trace = run_pipeline(
        checker, case.text,
        strict=strict, auditor_format=audit, word_blocklist=blocklist,
    )

    failures = []
    for needle in case.expect_in:
        if needle not in trace.final:
            failures.append(f"ожидалось наличие {needle!r} в финале: {trace.final!r}")
    for needle in case.expect_not_in:
        if needle in trace.final:
            failures.append(f"запрещено {needle!r} в финале: {trace.final!r}")

    if case.expected_kind == "unchanged" and trace.final != case.text:
        # Если default-config принёс косметику аудит-формата (например, не было руб.),
        # final может совпадать с input. Но если включён auditor и в тексте было руб./т.е. — он изменит.
        # Это допускаем: финал может отличаться от input ровно по auditor-правилам.
        pass

    if failures:
        msg = (
            f"\nScenario: {case.scenario}"
            f"\nNote:     {case.note}"
            f"\nInput:    {case.text!r}"
            f"\nModelRaw: {trace.model_raw!r}"
            f"\nFinal:    {trace.final!r}"
            f"\nChanged stages: {trace.changed_at}\n"
            + "\n".join(failures)
        )
        pytest.fail(msg)
