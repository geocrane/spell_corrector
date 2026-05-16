"""
Прогон всех сценариев из corpus.ALL_CASES через реальный pipeline и
формирование отчёта tests/results/report.md.

Запуск:
    /Users/geocrane/Dev/.venv/bin/python -m tests.report_runner
"""
import os
import sys
import time
from collections import defaultdict

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

import spell_checker as sc  # noqa: E402
from core.sentence_split import split_into_sentences  # noqa: E402

from tests.corpus import ALL_CASES, SPLIT_CASES  # noqa: E402
from tests.pipeline import run_pipeline  # noqa: E402


def _check_case(case, trace) -> tuple[bool, list]:
    """Вернуть (pass, [reasons_for_fail])."""
    reasons = []
    for needle in case.expect_in:
        if needle not in trace.final:
            reasons.append(f"отсутствует {needle!r}")
    for needle in case.expect_not_in:
        if needle in trace.final:
            reasons.append(f"присутствует запрещённое {needle!r}")
    if case.expected_kind == "fixed" and trace.final == case.text:
        reasons.append("expected_kind=fixed, но финал == input (модель/защиты ничего не сделали)")
    if case.expected_kind == "unchanged" and trace.final != case.text:
        # Для unchanged мы НЕ требуем абсолютной идентичности — auditor_format может
        # законно преобразовать руб. → ₽. Поэтому фейлим только если expect_in/expect_not_in
        # имели запрет на изменение.
        if case.expect_not_in or case.expect_in:
            pass  # уже проверено выше
        else:
            reasons.append(f"expected_kind=unchanged, но финал отличается")
    return len(reasons) == 0, reasons


def _classify(case, trace, reasons) -> str:
    """Определить категорию проблемы: model / app_logic / split / mixed."""
    if not reasons:
        return "ok"
    # Если защита должна была сработать, а итог отличается от input — app_logic.
    # Если модель ничего не исправила (model_raw == sanitized), а ожидалось fixed — model.
    if case.expected_kind == "fixed" and trace.model_raw == trace.sanitized:
        return "model"
    # Если модель что-то исправила, но финал не таков как ожидалось:
    # сравним model_raw и final — если фильтры съели правку, это app_logic.
    if trace.model_raw != trace.sanitized and trace.final == trace.sanitized:
        return "app_logic"
    # Иначе — смешанная зона.
    return "mixed"


def main():
    print("Loading model...")
    t0 = time.time()
    checker = sc.SpellChecker.get_instance()
    checker.load_model()
    print(f"Model loaded in {time.time()-t0:.1f}s")

    by_scenario = defaultdict(list)
    model_problems = []
    app_problems = []
    mixed_problems = []
    total_pass = 0

    print(f"Running {len(ALL_CASES)} cases...")
    for i, case in enumerate(ALL_CASES, 1):
        cfg = case.config or {}
        strict = cfg.get("strict_protection", True)
        audit = cfg.get("auditor_format", True)
        blocklist = cfg.get("word_blocklist", [])
        try:
            trace = run_pipeline(
                checker, case.text,
                strict=strict, auditor_format=audit, word_blocklist=blocklist,
            )
        except Exception as e:
            trace = None
            error = str(e)
            by_scenario[case.scenario].append((case, None, False, [f"EXCEPTION: {error}"], "exception"))
            print(f"  [{i:>3}/{len(ALL_CASES)}] {case.scenario}: EXCEPTION {e}")
            continue

        ok, reasons = _check_case(case, trace)
        category = _classify(case, trace, reasons)
        by_scenario[case.scenario].append((case, trace, ok, reasons, category))

        if ok:
            total_pass += 1
        else:
            if category == "model":
                model_problems.append((case, trace, reasons))
            elif category == "app_logic":
                app_problems.append((case, trace, reasons))
            else:
                mixed_problems.append((case, trace, reasons))

        print(f"  [{i:>3}/{len(ALL_CASES)}] {case.scenario}: {'PASS' if ok else f'FAIL ({category})'}")

    # ─── split-кейсы ─────────────────────────────────────────────────────
    print("Split cases...")
    split_results = []
    for text, expected_count, expected_chunks in SPLIT_CASES:
        result = split_into_sentences(text)
        chunks = [c[2] for c in result]
        ok = len(chunks) == expected_count
        if expected_chunks is not None:
            ok = ok and chunks == expected_chunks
        split_results.append((text, expected_count, expected_chunks, chunks, ok))

    # ─── формирование отчёта ─────────────────────────────────────────────
    out_dir = os.path.join(_ROOT, "tests", "results")
    os.makedirs(out_dir, exist_ok=True)
    report_path = os.path.join(out_dir, "report.md")

    lines = []
    L = lines.append
    L("# Отчёт о тестировании корректора\n")
    L(f"Прогнано тестов: **{len(ALL_CASES)}**, PASS: **{total_pass}**, FAIL: **{len(ALL_CASES) - total_pass}**\n")
    L("## Сводка по сценариям\n")
    L("| Сценарий | PASS / Total |")
    L("|---|---|")
    for scen, records in sorted(by_scenario.items()):
        p = sum(1 for r in records if r[2])
        L(f"| `{scen}` | {p}/{len(records)} |")

    # ─── 1. Ошибки модели ────────────────────────────────────────────────
    L("\n## 1. Проблемы модели\n")
    if not model_problems:
        L("_Не обнаружено._\n")
    for case, trace, reasons in model_problems:
        L(f"### `{case.scenario}` — {case.text!r}")
        if case.note:
            L(f"_{case.note}_\n")
        L(f"- **Input**:     `{case.text}`")
        L(f"- **Model raw**: `{trace.model_raw}`")
        L(f"- **Final**:     `{trace.final}`")
        L(f"- **Fail**: {'; '.join(reasons)}\n")

    # ─── 2. Ошибки логики приложения ─────────────────────────────────────
    L("\n## 2. Проблемы логики приложения (фильтры / защиты)\n")
    if not app_problems:
        L("_Не обнаружено._\n")
    for case, trace, reasons in app_problems:
        L(f"### `{case.scenario}` — {case.text!r}")
        if case.note:
            L(f"_{case.note}_\n")
        L(f"- **Input**:                 `{case.text}`")
        L(f"- **Model raw**:             `{trace.model_raw}`")
        L(f"- **After protect_tokens**:  `{trace.after_protect_tokens}`")
        L(f"- **After blocklist**:       `{trace.after_blocklist}`")
        L(f"- **After strict**:          `{trace.after_strict}`")
        L(f"- **After normalize**:       `{trace.after_normalize}`")
        L(f"- **After auditor**:         `{trace.after_auditor}`")
        L(f"- **Final**:                 `{trace.final}`")
        L(f"- **Changed at**:            {trace.changed_at}")
        L(f"- **Fail**: {'; '.join(reasons)}\n")

    # ─── 3. Смешанные ────────────────────────────────────────────────────
    L("\n## 3. Смешанные / неопределённые\n")
    if not mixed_problems:
        L("_Не обнаружено._\n")
    for case, trace, reasons in mixed_problems:
        L(f"### `{case.scenario}` — {case.text!r}")
        if case.note:
            L(f"_{case.note}_\n")
        L(f"- **Input**:     `{case.text}`")
        L(f"- **Model raw**: `{trace.model_raw}`")
        L(f"- **Final**:     `{trace.final}`")
        L(f"- **Changed at**: {trace.changed_at}")
        L(f"- **Fail**: {'; '.join(reasons)}\n")

    # ─── 4. Sentence split ───────────────────────────────────────────────
    L("\n## 4. Разбиение на предложения\n")
    L("| Input | Expected | Got | OK |")
    L("|---|---|---|---|")
    for text, exp_n, exp_chunks, chunks, ok in split_results:
        L(f"| `{text!r}` | {exp_n} | `{chunks!r}` | {'✓' if ok else '✗'} |")

    # ─── 5. Полный лист всех результатов ─────────────────────────────────
    L("\n## 5. Полный лист\n")
    for scen, records in sorted(by_scenario.items()):
        L(f"\n### `{scen}`\n")
        L("| Input | Model raw | Final | OK |")
        L("|---|---|---|---|")
        for case, trace, ok, reasons, cat in records:
            if trace is None:
                L(f"| `{case.text!r}` | — | EXCEPTION | ✗ |")
                continue
            mr = trace.model_raw.replace("|", "\\|")
            fn = trace.final.replace("|", "\\|")
            it = case.text.replace("|", "\\|")
            L(f"| `{it!r}` | `{mr!r}` | `{fn!r}` | {'✓' if ok else '✗ ('+cat+')'} |")

    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"\nReport written to: {report_path}")
    print(f"PASS: {total_pass}/{len(ALL_CASES)}")


if __name__ == "__main__":
    main()
