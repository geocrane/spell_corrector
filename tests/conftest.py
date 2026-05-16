"""
Общие фикстуры. Singleton-модель SpellChecker загружается один раз на сессию.
"""
import os
import sys

import pytest

# Корень проекта в sys.path, чтобы импортировать spell_checker как пакет
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)


@pytest.fixture(scope="session")
def checker():
    """Загруженный экземпляр SpellChecker (T5 + tokenizer)."""
    import spell_checker
    inst = spell_checker.SpellChecker.get_instance()
    inst.load_model()
    return inst
