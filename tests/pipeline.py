"""
Воспроизведение полного пайплайна из spell_checker.check_sentences_async,
но синхронно — для тестов и репортинга.

Возвращает не только final, но и промежуточные результаты каждого слоя.
"""
from dataclasses import dataclass, field
from typing import Optional

import spell_checker as sc


@dataclass
class StageTrace:
    original: str
    sanitized: str
    model_raw: str
    after_protect_tokens: str
    after_blocklist: str
    after_strict: str
    pre_normalize: str
    after_normalize: str
    after_auditor: str
    final: str
    has_error: bool
    # Какие фильтры изменили строку (для аналитики)
    changed_at: list = field(default_factory=list)


def run_pipeline(
    checker, text: str, *,
    strict: bool = True, auditor_format: bool = True,
    word_blocklist: Optional[list] = None,
) -> StageTrace:
    """Полный пайплайн как в worker(), но с трассировкой."""
    word_blocklist = word_blocklist or []

    sanitized = sc._sanitize_text(text)
    model_raw = checker.check(sanitized).strip()

    after_pt = sc._protect_tokens(sanitized, model_raw)
    after_bl = sc._protect_word_blocklist(sanitized, after_pt, word_blocklist) if word_blocklist else after_pt

    effective_strict = strict or auditor_format
    after_strict = sc._strict_protect(sanitized, after_bl) if effective_strict else after_bl
    pre_norm = after_strict
    after_norm = sc._normalize_corrected(sanitized, after_strict, strict=effective_strict)
    after_audit = sc._apply_auditor_format(after_norm) if auditor_format else after_norm

    final = after_audit
    has_error = final != sanitized
    if not has_error:
        # повторение логики worker'а: вернуть подмену если кол-во запятых разное
        def _mid_commas(t):
            return t.rstrip('.,;:!?…—–- ').count(',')
        if _mid_commas(sanitized) != _mid_commas(pre_norm):
            has_error = True
            fixed = sc._apply_only_comma_changes(sanitized, pre_norm)
            final = fixed if fixed != sanitized else pre_norm

    changed_at = []
    stages = [
        ("sanitize",         text,             sanitized),
        ("model",            sanitized,        model_raw),
        ("protect_tokens",   model_raw,        after_pt),
        ("blocklist",        after_pt,         after_bl),
        ("strict",           after_bl,         after_strict),
        ("normalize",        after_strict,     after_norm),
        ("auditor_format",   after_norm,       after_audit),
    ]
    for name, prev, cur in stages:
        if prev != cur:
            changed_at.append(name)

    return StageTrace(
        original=text, sanitized=sanitized, model_raw=model_raw,
        after_protect_tokens=after_pt, after_blocklist=after_bl,
        after_strict=after_strict, pre_normalize=pre_norm,
        after_normalize=after_norm, after_auditor=after_audit,
        final=final, has_error=has_error, changed_at=changed_at,
    )
