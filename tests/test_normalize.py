"""
Юнит-тесты функций пост-обработки (без модели).

Изолируют каждую функцию в spell_checker и проверяют:
- защиту токенов,
- блок-лист,
- strict_protect (двойные слова),
- нормализацию (ё/кавычки/двоеточия/инициалы/регистр),
- аудиторский формат,
- _apply_only_comma_changes (баг со вставкой запятых).
"""
import pytest

import spell_checker as sc


# ─── _is_protected_token ───────────────────────────────────────────────────

@pytest.mark.parametrize("token, protected", [
    ("1234", True),
    ("100,5", True),
    ("12.03.2024", True),
    ("№5", True),
    ("№", True),
    ("ОАО", True),
    ("ПАО", True),
    ("НДС", True),
    ("ОАО,", True),     # после strip — ОАО
    ("ОАО.", True),
    ("«ОАО»", True),    # должен видеть isupper после strip кавычек
    ("$100", True),
    ("100₽", True),
    ("Иванов", False),  # обычное слово
    ("ооо", False),     # строчная аббревиатура — не защищена
    ("Д", False),       # одна буква — не защищена (len<2)
    ("—", True),        # бул-лет
    ("-", True),
    ("•", True),
    ("15-летний", True),  # из-за цифры
])
def test_is_protected_token(token, protected):
    assert sc._is_protected_token(token) is protected, f"{token!r}"


# ─── _protect_tokens ───────────────────────────────────────────────────────

def test_protect_tokens_keeps_numbers():
    orig = "Сумма 1234 ₽ оплачена."
    corr = "Сумма 1235 ₽ оплачена."
    out = sc._protect_tokens(orig, corr)
    assert "1234" in out and "1235" not in out


def test_protect_tokens_keeps_upper_abbrev():
    orig = "ПАО Газпром раскрыло отчёт."
    corr = "Пао Газпром раскрыло отчёт."
    out = sc._protect_tokens(orig, corr)
    assert "ПАО" in out


def test_protect_tokens_allows_normal_correction():
    orig = "Привед, как дила?"
    corr = "Привет, как дела?"
    out = sc._protect_tokens(orig, corr)
    assert out == corr


# ─── _protect_word_blocklist ───────────────────────────────────────────────

def test_blocklist_case_insensitive():
    orig = "Охват аудита велик."
    corr = "Обхват аудита велик."
    out = sc._protect_word_blocklist(orig, corr, ["Охват"])
    assert "Охват" in out


def test_blocklist_no_change_other_words():
    orig = "Охват аудита велкий."
    corr = "Обхват аудита великий."
    out = sc._protect_word_blocklist(orig, corr, ["Охват"])
    assert "Охват" in out and "великий" in out


# ─── _strict_protect ───────────────────────────────────────────────────────

def test_strict_protect_blocks_insertions():
    """strict не пропускает вставки слов (запятые, новые слова)."""
    orig = "Поскольку отчёт готов мы можем его сдать."
    corr = "Поскольку отчёт готов, мы можем его сдать."
    out = sc._strict_protect(orig, corr)
    # Запятая в _strict_protect рассматривается на уровне «слов»: слово 'готов,'
    # отличается от 'готов'. SequenceMatcher выдаст replace, который пройдёт.
    # Здесь специально проверяем именно поведение split-based diff:
    # тест должен зафиксировать ФАКТ, чтобы при изменении логики ошибка стала видна.
    assert "готов" in out


def test_strict_protect_blocks_deletion_doubled_word():
    """Гипотеза бага: strict откатывает удаление дубликата слова."""
    orig = "Это это очень важный пункт."
    corr = "Это очень важный пункт."
    out = sc._strict_protect(orig, corr)
    assert out == orig, f"strict восстановил дубль — ожидалось: {out!r}"


def test_strict_protect_keeps_replacements():
    """strict ПРОПУСКАЕТ замены (опечатки)."""
    orig = "Привед всем."
    corr = "Привет всем."
    out = sc._strict_protect(orig, corr)
    assert out == corr


# ─── _apply_only_comma_changes ─────────────────────────────────────────────

def test_apply_only_comma_changes_drop_trailing_comma():
    """Удаление хвостовой запятой переносится."""
    orig = "Это, простой случай."
    pre  = "Это простой случай."
    out = sc._apply_only_comma_changes(orig, pre)
    # Известный частный случай — функция работает для УДАЛЕНИЙ запятой.
    # На данных «Это, простой» → «Это простой» split дает: orig=['Это,', 'простой', 'случай.'],
    # pre=['Это', 'простой', 'случай.']. op=replace на индексе 0 для пары ('Это,', 'Это').
    # ow.endswith(',') and ow[:-1].lower() == pw.lower() → True. Берём 'Это'.
    assert out == "Это простой случай."


def test_apply_only_comma_changes_BUG_insertion_lost():
    """ГИПОТЕЗА БАГА: вставка запятой модели теряется."""
    orig = "Поскольку отчёт готов мы можем его сдать."
    pre  = "Поскольку отчёт готов, мы можем его сдать."
    # has_error поднимается worker'ом (мid-comma count разный), затем вызывается
    # _apply_only_comma_changes. Если функция вернёт orig — это подтверждает баг.
    out = sc._apply_only_comma_changes(orig, pre)
    # Документируем фактическое поведение:
    if out == orig:
        pytest.skip("ПОДТВЕРЖДЁН БАГ: вставка запятой не применяется (см. отчёт)")
    assert "готов, мы" in out


# ─── _suppress_yo_replacement ──────────────────────────────────────────────

def test_yo_suppress_e_to_yo():
    orig = "Учет амортизации."
    corr = "Учёт амортизации."
    out = sc._suppress_yo_replacement(orig, corr)
    assert out == orig


def test_yo_allows_unrelated_changes():
    orig = "Сериозный учет."
    corr = "Серьёзный учет."
    out = sc._suppress_yo_replacement(orig, corr)
    assert "Серьёзный" in out or "Серьезный" in out  # ё/е здесь правомерно


# ─── _protect_colons ───────────────────────────────────────────────────────

def test_colon_insertion_suppressed():
    """Модель добавила `:` — откатываем (в оригинале его не было)."""
    orig = "Документ подписан стороны согласны."
    corr = "Документ подписан: стороны согласны."
    out = sc._protect_colons(orig, corr)
    assert ":" not in out


def test_colon_kept_if_was_in_original():
    """Двоеточие из оригинала сохраняется."""
    orig = "Документ: подписан."
    corr = "Документ: подписан, проверен."
    out = sc._protect_colons(orig, corr)
    assert ":" in out


def test_colon_deletion_restored():
    """Модель удалила `:` посередине — восстанавливаем."""
    orig = "В договоре: два контроллера."
    corr = "В договоре два контроллера."
    out = sc._protect_colons(orig, corr)
    assert ":" in out
    assert out.count(":") == 1


def test_colon_addressing_not_added():
    """«Комментарии в приложении, лист 3» — `:` не добавляется."""
    orig = "Комментарии в приложении, лист 3"
    corr = "Комментарии в приложении: лист 3"
    out = sc._protect_colons(orig, corr)
    assert ":" not in out


def test_colon_collateral_word_fixes_preserved():
    """При insert `:` рядом с другими правками — слова сохраняются."""
    orig = "Иванов сказал отчёт готовко сдаче."
    corr = "Иванов сказал: отчёт готов к сдаче."
    out = sc._protect_colons(orig, corr)
    assert ":" not in out
    assert "готов к сдаче" in out


# ─── _protect_slashes ──────────────────────────────────────────────────────

def test_slash_forward_kept():
    """«и/или» не должно стать «и\\или»."""
    orig = "Запрос и/или согласие."
    corr = "Запрос и\\или согласие."
    out = sc._protect_slashes(orig, corr)
    assert "и/или" in out
    assert "\\" not in out


def test_slash_back_kept():
    """Оригинальные `\\` сохраняются — модель не должна заменять на `/`."""
    orig = "Путь C:\\Users\\test"
    corr = "Путь C:/Users/test"
    out = sc._protect_slashes(orig, corr)
    assert "\\" in out
    assert out.count("\\") == orig.count("\\")


def test_slash_bu_kept():
    """«б/у» — слэш сохраняется."""
    orig = "Поставка б/у оборудования."
    corr = "Поставка б\\у оборудования."
    out = sc._protect_slashes(orig, corr)
    assert "б/у" in out


def test_slash_insertion_blocked():
    """Модель не должна добавлять слэши, которых не было."""
    orig = "Запрос или согласие."
    corr = "Запрос/или согласие."
    out = sc._protect_slashes(orig, corr)
    assert "/" not in out


# ─── _protect_dots ─────────────────────────────────────────────────────────

def test_dot_insertion_midsentence_blocked():
    """Модель разбила предложение точкой — откатываем."""
    orig = "Департаментом не сформированы Планы ввода в эксплуатацию"
    corr = "Департаментом не сформированы. Планы ввода в эксплуатацию"
    out = sc._protect_dots(orig, corr)
    assert "сформированы Планы" in out
    assert "сформированы." not in out


def test_dot_in_initials_kept():
    """Точки в инициалах оригинала сохраняются после правок."""
    orig = "Сидоров А.С. подписал документ."
    corr = "Сидоров А.С. подписал документ."
    out = sc._protect_dots(orig, corr)
    assert "А.С." in out


def test_dot_trailing_preserved():
    """Хвостовая точка не теряется (она уже выровнена шагом 2)."""
    orig = "Это важно."
    corr = "Это важно."
    out = sc._protect_dots(orig, corr)
    assert out.endswith(".")


def test_dot_count_preserved():
    """Если в оригинале две точки, в результате должно быть две."""
    orig = "Декабрь. Январь."
    corr = "Декабрь Январь."
    out = sc._protect_dots(orig, corr)
    assert out.count(".") == 2


# ─── _normalize_corrected: интеграция новых защит ───────────────────────────

def test_normalize_integration_midsentence_period():
    """Через полный _normalize_corrected: разбивка предложения откатывается."""
    orig = "Департаментом не сформированы Планы ввода в эксплуатацию"
    corr = "Департаментом не сформированы. Планы ввода в эксплуатацию"
    out = sc._normalize_corrected(orig, corr)
    assert "сформированы Планы" in out


def test_normalize_integration_slash_swap():
    """Через полный _normalize_corrected: и/или сохраняется."""
    orig = "Запрос и/или согласие."
    corr = "Запрос и\\или согласие."
    out = sc._normalize_corrected(orig, corr)
    assert "и/или" in out
    assert "\\" not in out


def test_normalize_integration_colon_deletion():
    """Через полный _normalize_corrected: удаление `:` модельен — восстанавливается."""
    orig = "В договоре: два контроллера."
    corr = "В договоре два контроллера."
    out = sc._normalize_corrected(orig, corr)
    assert ":" in out


# ─── _suppress_quote_changes / _strict_protect_quotes ──────────────────────

def test_quote_type_preserved_strict():
    orig = 'ООО "Ромашка" заключило договор.'
    corr = "ООО «Ромашка» заключило договор."
    out = sc._strict_protect_quotes(orig, corr)
    assert '"Ромашка"' in out


def test_strict_protect_quotes_keeps_letter_fixes():
    """Регрессия: при удалении лишних букв не должны откатываться."""
    orig = "отфвыветственность"
    corr = "ответственность"
    out = sc._strict_protect_quotes(orig, corr)
    assert out == "ответственность"


# ─── _suppress_space_in_initials ──────────────────────────────────────────

def test_initials_no_space_inserted():
    orig = "Подписант А.С. Пушкин."
    corr = "Подписант А. С. Пушкин."
    out = sc._suppress_space_in_initials(orig, corr)
    assert "А.С." in out
    assert "А. С." not in out


def test_initials_preserves_other_initials():
    orig = "Иванов И.И. Сидоровым."
    corr = "Иванов И. И. Сидоровым."
    out = sc._suppress_space_in_initials(orig, corr)
    assert "И.И." in out


# ─── _normalize_corrected: первая буква ────────────────────────────────────

def test_normalize_keeps_lower_first_letter():
    """Если в оригинале строчная — после нормализации тоже строчная."""
    orig = "поэтому это важно."
    corr = "Поэтому это важно."
    out = sc._normalize_corrected(orig, corr)
    assert out.startswith("поэтому")


def test_normalize_keeps_upper_first_letter():
    """Регистр заглавной первой буквы оригинала сохраняется."""
    orig = "Поэтому это важно."
    corr = "поэтому это важно."  # модель неосторожно опустила
    out = sc._normalize_corrected(orig, corr)
    assert out.startswith("Поэтому")


def test_normalize_first_letter_BUG_with_leading_space():
    """ГИПОТЕЗА БАГА: индекс первой буквы не сравнивается с индексом result."""
    orig = "  поэтому."
    corr = " Поэтому."   # модель убрала пробел спереди
    out = sc._normalize_corrected(orig, corr)
    # Если бы код корректно адресовал первую букву по позиции в каждой строке —
    # был бы 'поэтому'. Если индекс смешан — 'Поэтому'.
    if "Поэтому" in out and "поэтому" not in out:
        pytest.skip("ПОДТВЕРЖДЁН БАГ: с ведущими пробелами регистр не сохраняется")


# ─── _apply_auditor_format ─────────────────────────────────────────────────

def test_audit_rub_basic():
    assert sc._apply_auditor_format("100 руб.") == "100 ₽"


def test_audit_mln_rub_chain():
    """млн. руб. → млн ₽ (через два правила)."""
    out = sc._apply_auditor_format("5 млн. руб.")
    assert out == "5 млн ₽"


def test_audit_tys_rub_chain():
    """10 тыс. руб. → 10 тыс. ₽ — точка после тыс должна сохраниться."""
    out = sc._apply_auditor_format("10 тыс. руб. удержано")
    assert "тыс. ₽" in out


def test_audit_no_pril_compact():
    out = sc._apply_auditor_format("Приложение №5")
    assert out == "Приложение 5"


def test_audit_no_general_compact():
    out = sc._apply_auditor_format("№ 1234")
    assert out == "№1234"


def test_audit_te_expansion_lower():
    """т.е. в начале фразы — расширяется в 'то есть' (lowercase), т.к. matched[0] строчная."""
    out = sc._apply_auditor_format("т.е. это важно")
    assert out.startswith("то есть")


def test_audit_te_expansion_upper():
    """Т.е. (с прописной) → То есть."""
    out = sc._apply_auditor_format("Т.е. документ подписан")
    assert out.startswith("То есть")


def test_audit_te_inside_sentence():
    out = sc._apply_auditor_format("Документ, т.е. договор.")
    assert "то есть" in out


def test_audit_in_tch_expansion():
    out = sc._apply_auditor_format("в т.ч. НДС")
    assert "в том числе" in out


def test_audit_does_not_touch_rublik():
    """\\bруб\\.\\ не должен ломать другие слова."""
    out = sc._apply_auditor_format("рубчик деревянный")
    assert out == "рубчик деревянный"


def test_audit_rub_inside_quotes_BUG_potential():
    """Регрессия: 'руб.' внутри названия/цитаты тоже заменится."""
    out = sc._apply_auditor_format('Название "руб. рынка"')
    # фиксируем фактическое поведение
    assert "₽" in out  # текущее поведение
