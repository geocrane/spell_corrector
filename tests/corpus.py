"""
Тестовый корпус по сценариям исправления ошибок.

Каждый Case — независимая единица проверки:
- scenario:      тег группы (orthography_simple, punct_comma_missing, ...)
- text:          исходное предложение
- expected_kind: "fixed" / "unchanged" / "partial"
- expect_in:     подстроки, которые должны присутствовать в исправлении
- expect_not_in: подстроки, которые НЕ должны появиться (защита/неизменность)
- note:          краткий комментарий о сути проверки
- config:        переопределения настроек (strict_protection, auditor_format, word_blocklist)
"""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Case:
    scenario: str
    text: str
    expected_kind: str  # "fixed", "unchanged", "partial"
    expect_in: list = field(default_factory=list)
    expect_not_in: list = field(default_factory=list)
    note: str = ""
    config: dict = field(default_factory=dict)


# Базовая конфигурация для большинства сценариев — пользовательский config.json
_CFG_DEFAULT = {"strict_protection": True, "auditor_format": True, "word_blocklist": ["Охват"]}
_CFG_NO_STRICT = {"strict_protection": False, "auditor_format": False, "word_blocklist": []}
_CFG_BLOCK = {"strict_protection": True, "auditor_format": False, "word_blocklist": ["Охват", "Дельта"]}
_CFG_NO_AUDIT = {"strict_protection": True, "auditor_format": False, "word_blocklist": []}


# ─── A. Простая орфография ────────────────────────────────────────────────
A_simple = [
    Case("orthography_simple", "Превед, как дила?", "fixed",
         expect_in=["Привет", "дела"], note="базовые опечатки"),
    Case("orthography_simple", "Я хочю кофий с молаком.", "fixed",
         expect_in=["хочу", "молоком"]),
    Case("orthography_simple", "Завтре мы пойдём в кено.", "fixed",
         expect_in=["Завтра", "кино"]),
    Case("orthography_simple", "Скуратор прислал зомечяние.", "fixed",
         expect_in=["Куратор", "замечание"]),
    Case("orthography_simple", "Он раздолбает работу за нидилю.", "fixed",
         expect_in=["неделю"]),
    Case("orthography_simple", "Это очинь сильная позиция кампании.", "fixed",
         expect_in=["очень", "компании"]),
    Case("orthography_simple", "Сегодне мы рассмотрим вапросы аудита.", "fixed",
         expect_in=["Сегодня", "вопросы"]),
    Case("orthography_simple", "Бугалтер закроет переод в пятницу.", "fixed",
         expect_in=["Бухгалтер", "период"]),
    Case("orthography_simple", "Пожалусто, проверти запоси в журнале.", "fixed",
         expect_in=["Пожалуйста", "проверьте", "записи"]),
    Case("orthography_simple", "Этат рапорт нужин зафтра.", "fixed",
         expect_in=["Этот", "нужен", "завтра"]),
]


# ─── B. Сложная орфография ────────────────────────────────────────────────
B_complex = [
    Case("orthography_complex", "Компания провела учет на конец преода.", "fixed",
         expect_in=["периода"], note="усечённое слово"),
    Case("orthography_complex", "Аудитор зафиксировал суцщественные ошибки.", "fixed",
         expect_in=["существенные"]),
    Case("orthography_complex", "Поясните метод деректного перехода между активами.", "fixed",
         expect_in=["прямого"]),
    Case("orthography_complex", "Эти данные коминяируют разные источники доходов.", "fixed",
         expect_in=["комбинируют"]),
    Case("orthography_complex", "Контрагент не выпонил договорные обязательства в срок.", "fixed",
         expect_in=["выполнил"]),
    Case("orthography_complex", "Договорной отел согласовал условия с поставщиком.", "fixed",
         expect_in=["отдел"]),
    Case("orthography_complex", "Кампания вкючает 5 ключевых процедур.", "fixed",
         expect_in=["включает", "5"]),
    Case("orthography_complex", "Премия начислятется ежемесчно по графику.", "fixed",
         expect_in=["начисляется", "ежемесячно"]),
    Case("orthography_complex", "Сотруник управления опубликвал отчет об инциденте.", "fixed",
         expect_in=["Сотрудник", "опубликовал"]),
    Case("orthography_complex", "Поэому стороны прекращают сотрудничесво.", "fixed",
         expect_in=["Поэтому", "сотрудничество"]),
]


# ─── C. Пунктуация — отсутствующие запятые ────────────────────────────────
C_comma_missing = [
    Case("punct_comma_missing", "Поскольку отчёт готов мы можем его сдать.", "fixed",
         expect_in=["готов, мы"], note="запятая перед главной частью СПП"),
    Case("punct_comma_missing", "Несмотря на задержку проект завершён вовремя.", "fixed",
         expect_in=["задержку, проект"]),
    Case("punct_comma_missing", "Иванов А.С. который возглавлял отдел уволился.", "fixed",
         expect_in=["А.С., который"]),
    Case("punct_comma_missing", "Если данные верны то их можно отправить.", "fixed",
         expect_in=["верны, то"]),
    Case("punct_comma_missing", "Документ подписан проверен и сдан.", "fixed",
         expect_in=["подписан, проверен"]),
    Case("punct_comma_missing", "По мнению аудитора подходящая методика выборка.", "fixed",
         expect_in=["аудитора,"]),
    Case("punct_comma_missing", "Сергей пришёл сел и начал работать.", "fixed",
         expect_in=["пришёл, сел"]),
    Case("punct_comma_missing", "Когда сделка закроется деньги поступят на счёт.", "fixed",
         expect_in=["закроется, деньги"]),
    Case("punct_comma_missing", "Аудит проведён замечаний нет.", "fixed",
         expect_in=["проведён, замечаний"]),
    Case("punct_comma_missing", "Хотя срок истёк работа продолжается.", "fixed",
         expect_in=["истёк, работа"]),
]


# ─── D. Пунктуация — лишние запятые ───────────────────────────────────────
D_comma_extra = [
    Case("punct_comma_extra", "Договор подписан, и сдан в архив.", "fixed",
         expect_in=["подписан и сдан"], expect_not_in=["подписан, и"]),
    Case("punct_comma_extra", "Все, активы переданы по акту.", "fixed",
         expect_in=["Все активы"]),
    Case("punct_comma_extra", "Сторона, А обязуется заплатить стороне Б.", "fixed",
         expect_in=["Сторона А"]),
    Case("punct_comma_extra", "Документ, и приложение к нему утверждены.", "fixed",
         expect_in=["Документ и приложение"]),
    Case("punct_comma_extra", "Аудитор, провёл проверку отчёта.", "fixed",
         expect_in=["Аудитор провёл"]),
    Case("punct_comma_extra", "Иванов, и Петров согласовали условия.", "fixed",
         expect_in=["Иванов и Петров"]),
    Case("punct_comma_extra", "Это, простой случай.", "fixed",
         expect_in=["Это простой"]),
    Case("punct_comma_extra", "Сумма, в размере 100 ₽ оплачена.", "fixed",
         expect_in=["Сумма в размере"]),
    Case("punct_comma_extra", "По итогам, проверки замечаний нет.", "fixed",
         expect_in=["По итогам проверки"]),
    Case("punct_comma_extra", "В соответствии, с регламентом отчёт сдан.", "fixed",
         expect_in=["В соответствии с"]),
]


# ─── E. Защита чисел / валют ──────────────────────────────────────────────
E_protect_numbers = [
    Case("protect_numbers", "Сумма штрафа состовляет 1234 рубля.", "fixed",
         expect_in=["1234"], expect_not_in=[" 1235 ", " 12 34 "]),
    Case("protect_numbers", "Договор № 12-345/А зарегестрирован.", "fixed",
         expect_in=["12-345/А", "зарегистрирован"]),
    Case("protect_numbers", "Бюджет — 100 500 ₽ на год.", "unchanged",
         expect_in=["100 500 ₽"], note="ничего трогать нельзя"),
    Case("protect_numbers", "Курс доллара 92.45 ₽ по состаянию на сегодня.", "fixed",
         expect_in=["92.45 ₽", "состоянию"]),
    Case("protect_numbers", "Здание построено в 1985 году.", "unchanged",
         expect_in=["1985"]),
    Case("protect_numbers", "Цена 1 шт. — 19 999.99 ₽.", "unchanged",
         expect_in=["19 999.99 ₽", "шт."]),
    Case("protect_numbers", "Доля участия — 12,5 % уставного капитала.", "unchanged",
         expect_in=["12,5 %"]),
    Case("protect_numbers", "Дата отчёта: 12.03.2024 — пятница.", "unchanged",
         expect_in=["12.03.2024"]),
    Case("protect_numbers", "Премия рассчитона как 0,5 % от выручки.", "fixed",
         expect_in=["рассчитана", "0,5 %"]),
    Case("protect_numbers", "Артикул 2024-AC/15-bis в новой колекции.", "fixed",
         expect_in=["2024-AC/15-bis", "коллекции"]),
]


# ─── F. Защита заглавных аббревиатур ──────────────────────────────────────
F_protect_abbrev = [
    Case("protect_abbrev_upper", "ПАО Газпром раскрыло консолидированую отчётность.", "fixed",
         expect_in=["ПАО", "консолидированную"]),
    Case("protect_abbrev_upper", "ООО Ромашка зарегестрировано в 2010 году.", "fixed",
         expect_in=["ООО", "зарегистрировано", "2010"]),
    Case("protect_abbrev_upper", "НДС учтён в стоимасти товара.", "fixed",
         expect_in=["НДС", "стоимости"]),
    Case("protect_abbrev_upper", "УК РФ статья 159 — мошеничество.", "fixed",
         expect_in=["УК РФ", "159", "мошенничество"]),
    Case("protect_abbrev_upper", "ИФНС № 5 провела камеральную поверку.", "fixed",
         expect_in=["ИФНС", "5", "проверку"]),
    Case("protect_abbrev_upper", "МСФО IFRS 16 применяется с 2019 года.", "unchanged",
         expect_in=["МСФО", "IFRS", "16", "2019"]),
    Case("protect_abbrev_upper", "КПП и ИНН указаны на титульной строанице.", "fixed",
         expect_in=["КПП", "ИНН", "странице"]),
    Case("protect_abbrev_upper", "ЦБ РФ снизил ключевую ставку на 0,5 %.", "unchanged",
         expect_in=["ЦБ РФ", "0,5 %"]),
    Case("protect_abbrev_upper", "ФНС России опуликовала разъяснения.", "fixed",
         expect_in=["ФНС", "опубликовала"]),
    Case("protect_abbrev_upper", "КГБ ФСБ МВД — историчиское переименование.", "fixed",
         expect_in=["КГБ", "ФСБ", "МВД", "историческое"]),
]


# ─── G. Блок-лист ─────────────────────────────────────────────────────────
G_blocklist = [
    Case("protect_blocklist", "Охват аудита включает все процессы.", "unchanged",
         expect_in=["Охват"], config=_CFG_BLOCK),
    Case("protect_blocklist", "Параметр охват скоррктирован по периоду.", "fixed",
         expect_in=["скорректирован"], expect_not_in=[], config=_CFG_BLOCK,
         note="блок-лист case-insensitive, охват не должен меняться, но опечатка скоррктирован — исправиться"),
    Case("protect_blocklist", "Дельта между планом и фактом велика.", "unchanged",
         expect_in=["Дельта"], config=_CFG_BLOCK),
    Case("protect_blocklist", "Эта дельта обусловлена сезонностью.", "unchanged",
         expect_in=["дельта"], config=_CFG_BLOCK),
    Case("protect_blocklist", "Охват выборки 25 %.", "unchanged",
         expect_in=["Охват", "25 %"], config=_CFG_BLOCK),
    Case("protect_blocklist", "Расчёт дельта потоков платежей.", "unchanged",
         expect_in=["дельта"], config=_CFG_BLOCK),
    Case("protect_blocklist", "Охват тестирования контроль превышает 80 %.", "partial",
         expect_in=["Охват", "80 %"], config=_CFG_BLOCK,
         note="нужна запятая после тестирования; охват не трогаем"),
    Case("protect_blocklist", "Дельта-хеджирование применено к опционам.", "unchanged",
         expect_in=["Дельта-хеджирование"], config=_CFG_BLOCK),
    Case("protect_blocklist", "На этот период охват аудита расширен.", "unchanged",
         expect_in=["охват"], config=_CFG_BLOCK),
    Case("protect_blocklist", "Дельта запасов соответствует ожиданиям.", "unchanged",
         expect_in=["Дельта"], config=_CFG_BLOCK),
]


# ─── H. Защита инициалов ──────────────────────────────────────────────────
H_initials = [
    Case("protect_initials", "Подписант — Иванов А.С. главный бухгалтер.", "partial",
         expect_in=["А.С."], expect_not_in=["А. С."], config=_CFG_DEFAULT,
         note="пробел в инициалы не вставлять"),
    Case("protect_initials", "Контракт согласован И.И.Сидоровым.", "unchanged",
         expect_in=["И.И.Сидоровым"], expect_not_in=["И. И."], config=_CFG_DEFAULT),
    Case("protect_initials", "Аудитор Петров А.А. провёл проверку.", "unchanged",
         expect_in=["А.А."], expect_not_in=["А. А."], config=_CFG_DEFAULT),
    Case("protect_initials", "Заверил М.К. Жуков по доверенности.", "unchanged",
         expect_in=["М.К."], expect_not_in=["М. К."], config=_CFG_DEFAULT),
    Case("protect_initials", "Куратор проекта — Сидоров П.П.", "unchanged",
         expect_in=["П.П."], expect_not_in=["П. П."], config=_CFG_DEFAULT),
    Case("protect_initials", "Е.А. Ковалёв возглавляет отдел.", "unchanged",
         expect_in=["Е.А."], config=_CFG_DEFAULT),
    Case("protect_initials", "Иванов, А.С., передал документы.", "unchanged",
         expect_in=["А.С."], config=_CFG_DEFAULT),
    Case("protect_initials", "Письмо от Петрова А. А. датировано январём.", "partial",
         expect_in=["Петрова"], config=_CFG_DEFAULT,
         note="модель может слепить А. А. → А.А., в этом сценарии ОК"),
    Case("protect_initials", "Подпись: Кузнецов В.С.; печать налогового органа.", "unchanged",
         expect_in=["В.С."], expect_not_in=["В. С."], config=_CFG_DEFAULT),
    Case("protect_initials", "Решение принято Ивановым А.С. лично.", "unchanged",
         expect_in=["А.С."], expect_not_in=["А. С."], config=_CFG_DEFAULT),
]


# ─── I. Защита кавычек ────────────────────────────────────────────────────
I_quotes = [
    Case("quotes_mix", 'ООО "Ромашка" заключило договор.', "unchanged",
         expect_in=['"Ромашка"'], expect_not_in=["«Ромашка»"], config=_CFG_DEFAULT,
         note="strict_protect_quotes должна сохранить тип"),
    Case("quotes_mix", "Компания «Альфа» расширила деятельность.", "unchanged",
         expect_in=["«Альфа»"], expect_not_in=['"Альфа"'], config=_CFG_DEFAULT),
    Case("quotes_mix", 'Документ — "Положение об учётной политике".', "unchanged",
         expect_in=['"Положение'], config=_CFG_DEFAULT),
    Case("quotes_mix", 'Параграф „Введение" описывает цели.', "unchanged",
         expect_in=["„Введение"], config=_CFG_DEFAULT),
    Case("quotes_mix", "Это «термин» из стандарта IFRS.", "unchanged",
         expect_in=["«термин»", "IFRS"], config=_CFG_DEFAULT),
    Case("quotes_mix", 'Цитата: "будет завершено в срок" — из протокола.', "unchanged",
         expect_in=['"будет', 'срок"'], config=_CFG_DEFAULT),
    Case("quotes_mix", "Использован термин «существенность» в значении …", "unchanged",
         expect_in=["«существенность»"], config=_CFG_DEFAULT),
    Case("quotes_mix", "«Договор-1» зарегистрирован в журнале.", "unchanged",
         expect_in=["«Договор-1»"], config=_CFG_DEFAULT),
    Case("quotes_mix", "‛Сторона A' и ’Сторона Б' заключили соглашение.", "unchanged",
         expect_in=["Сторона"], config=_CFG_DEFAULT),
    Case("quotes_mix", 'На странице 5 встречается слово "ошибка".', "unchanged",
         expect_in=['"ошибка"', "5"], config=_CFG_DEFAULT),
]


# ─── J. Ё/Е ──────────────────────────────────────────────────────────────
J_yo = [
    Case("yo_swap", "Учет осуществляется ежемесячно.", "unchanged",
         expect_in=["Учет"], expect_not_in=["Учёт"], config=_CFG_DEFAULT,
         note="модель часто предлагает Учёт — должно подавляться"),
    Case("yo_swap", "Все объекты прошли инвентаризацию.", "unchanged",
         expect_in=["Все"], expect_not_in=["Всё"], config=_CFG_DEFAULT),
    Case("yo_swap", "Береза посажена в 2015 году.", "unchanged",
         expect_in=["Береза"], expect_not_in=["Берёза"], config=_CFG_DEFAULT),
    Case("yo_swap", "Поезд прибыл во время.", "unchanged",
         expect_in=["Поезд"], expect_not_in=["Поёзд"], config=_CFG_DEFAULT),
    Case("yo_swap", "На елке висят игрушки.", "unchanged",
         expect_in=["елке"], expect_not_in=["ёлке"], config=_CFG_DEFAULT),
    Case("yo_swap", "Сериозный подход к работе.", "fixed",
         expect_in=["Серьёзный", "к работе"], config=_CFG_DEFAULT,
         note="нужна реальная фикс через ь+ё, не должно заглохнуть из-за подавления"),
    Case("yo_swap", "На съезде делегаты обсудили устав.", "unchanged",
         expect_in=["съезде"], config=_CFG_DEFAULT),
    Case("yo_swap", "Ёлка — символ Нового года.", "unchanged",
         expect_in=["Ёлка"], expect_not_in=["Елка"], config=_CFG_DEFAULT),
    Case("yo_swap", "Тёплый чай в холодную пагоду.", "fixed",
         expect_in=["погоду"], config=_CFG_DEFAULT),
    Case("yo_swap", "Учёт амортизации проводится ежемесчно.", "fixed",
         expect_in=["ежемесячно"], config=_CFG_DEFAULT),
]


# ─── K. Регистр первой буквы ──────────────────────────────────────────────
K_case = [
    Case("case_first_letter", "поэтому стороны прекращают сотрудничество.", "unchanged",
         expect_in=["поэтому"], expect_not_in=["Поэтому"], config=_CFG_DEFAULT,
         note="регистр первой буквы оригинала должен сохраняться"),
    Case("case_first_letter", "это простое предложение.", "unchanged",
         expect_in=["это"], expect_not_in=["Это"], config=_CFG_DEFAULT),
    Case("case_first_letter", "В договоре написано: «оплата 100 ₽».", "unchanged",
         expect_in=["«оплата", "100 ₽"], config=_CFG_DEFAULT),
    Case("case_first_letter", "Сумма — 100 ₽; срок — 30 дней.", "unchanged",
         expect_in=["100 ₽", "30 дней"], config=_CFG_DEFAULT),
    Case("case_first_letter", "как мы уже отмечали ранее.", "unchanged",
         expect_in=["как"], expect_not_in=["Как"], config=_CFG_DEFAULT),
    Case("case_first_letter", "ивановА.С.не подписал документ.", "fixed",
         expect_in=["Иванов"], config=_CFG_DEFAULT,
         note="плотно слитный текст — пограничный случай"),
    Case("case_first_letter", "(примечание сослки см. ниже)", "fixed",
         expect_in=["ссылки"], config=_CFG_DEFAULT),
    Case("case_first_letter", "— Здесь начинается прямая речь.", "unchanged",
         expect_in=["— Здесь"], config=_CFG_DEFAULT),
    Case("case_first_letter", "№ 5 от 12.03.2024", "unchanged",
         expect_in=["№ 5", "12.03.2024"], config=_CFG_NO_AUDIT,
         note="без аудит-формата № 5 не сжимается"),
    Case("case_first_letter", "п.1.2.3. договора устанавливает срок.", "unchanged",
         expect_in=["п.1.2.3."], config=_CFG_DEFAULT),
]


# ─── L. Аудит: руб. → ₽ ──────────────────────────────────────────────────
L_rub = [
    Case("auditor_rub", "Сумма штрафа — 100 руб.", "fixed",
         expect_in=["100 ₽"], expect_not_in=["руб."], config=_CFG_DEFAULT),
    Case("auditor_rub", "Бюджет 5 млн. руб.", "fixed",
         expect_in=["млн ₽"], expect_not_in=["млн. руб"], config=_CFG_DEFAULT,
         note="млн. руб → млн ₽ через два правила"),
    Case("auditor_rub", "Стоимость 1500 рублей за единицу.", "fixed",
         expect_in=["1500 ₽"], expect_not_in=["рублей"], config=_CFG_DEFAULT),
    Case("auditor_rub", "Доход 200 рубли по контракту.", "fixed",
         expect_in=["200 ₽"], expect_not_in=["рубли"], config=_CFG_DEFAULT),
    Case("auditor_rub", "Налог 0,5 млрд. руб. за период.", "fixed",
         expect_in=["млрд ₽"], expect_not_in=["млрд."], config=_CFG_DEFAULT),
    Case("auditor_rub", "10 тыс. руб. удержано в виде НДФЛ.", "fixed",
         expect_in=["тыс. ₽"], expect_not_in=["тыс. руб"], config=_CFG_DEFAULT,
         note="должно остаться тыс. ₽ (с точкой)"),
    Case("auditor_rub", "Авансы выданы в размере 50 000 руб.", "fixed",
         expect_in=["50 000 ₽"], expect_not_in=["руб."], config=_CFG_DEFAULT),
    Case("auditor_rub", "По договору № 5 уплачено 30 руб.", "fixed",
         expect_in=["№5", "30 ₽"], expect_not_in=["№ 5", "руб."], config=_CFG_DEFAULT,
         note="№ 5 → №5 + руб → ₽"),
    Case("auditor_rub", "Премия 25 000 рублей.", "fixed",
         expect_in=["25 000 ₽"], config=_CFG_DEFAULT),
    Case("auditor_rub", "Курс: 92,45 руб. за доллар.", "fixed",
         expect_in=["92,45 ₽"], expect_not_in=["руб."], config=_CFG_DEFAULT),
]


# ─── M. Аудит: тыс/млн/млрд ──────────────────────────────────────────────
M_tysm = [
    Case("auditor_tysm", "Бюджет 5 тыс ₽.", "fixed",
         expect_in=["тыс. ₽"], config=_CFG_DEFAULT, note="вставка точки"),
    Case("auditor_tysm", "Стоимость 10 млн. ₽.", "fixed",
         expect_in=["млн ₽"], expect_not_in=["млн."], config=_CFG_DEFAULT),
    Case("auditor_tysm", "Капитал 5 млрд. ₽.", "fixed",
         expect_in=["млрд ₽"], expect_not_in=["млрд."], config=_CFG_DEFAULT),
    Case("auditor_tysm", "Расходы 100 тыс. ₽.", "unchanged",
         expect_in=["тыс. ₽"], config=_CFG_DEFAULT),
    Case("auditor_tysm", "Доход 1 млн ₽.", "unchanged",
         expect_in=["млн ₽"], expect_not_in=["млн."], config=_CFG_DEFAULT),
    Case("auditor_tysm", "Активы 25 млрд ₽.", "unchanged",
         expect_in=["млрд ₽"], expect_not_in=["млрд."], config=_CFG_DEFAULT),
    Case("auditor_tysm", "Премия 50 тыс ₽.", "fixed",
         expect_in=["тыс. ₽"], config=_CFG_DEFAULT),
    Case("auditor_tysm", "Объём 200 млн. ₽.", "fixed",
         expect_in=["млн ₽"], expect_not_in=["млн."], config=_CFG_DEFAULT),
    Case("auditor_tysm", "Затраты 5 млрд. ₽.", "fixed",
         expect_in=["млрд ₽"], expect_not_in=["млрд."], config=_CFG_DEFAULT),
    Case("auditor_tysm", "Резерв 10 тыс. ₽.", "unchanged",
         expect_in=["тыс. ₽"], config=_CFG_DEFAULT),
]


# ─── N. Аудит: Приложение № ──────────────────────────────────────────────
N_pril = [
    Case("auditor_pril_no", "См. Приложение №5 к договору.", "fixed",
         expect_in=["Приложение 5"], expect_not_in=["Приложение №5"], config=_CFG_DEFAULT),
    Case("auditor_pril_no", "Согласно приложения № 12.", "fixed",
         expect_in=["приложения 12"], expect_not_in=["№ 12", "№12"], config=_CFG_DEFAULT),
    Case("auditor_pril_no", "В приложении № 3 указаны параметры.", "fixed",
         expect_in=["приложении 3"], config=_CFG_DEFAULT),
    Case("auditor_pril_no", "Приложение №7 содержит расчёт.", "fixed",
         expect_in=["Приложение 7"], config=_CFG_DEFAULT),
    Case("auditor_pril_no", "К приложениях №2 нет замечаний.", "fixed",
         expect_in=["приложениях 2"], config=_CFG_DEFAULT,
         note="приложениях — грамматическая ошибка, должно быть приложении"),
    Case("auditor_pril_no", "Утверждено в приложении №1.", "fixed",
         expect_in=["приложении 1"], config=_CFG_DEFAULT),
    Case("auditor_pril_no", "Приложением №4 к договору установлено.", "fixed",
         expect_in=["Приложением 4"], config=_CFG_DEFAULT),
    Case("auditor_pril_no", "Приложения №№ 1, 2, 3 подписаны.", "fixed",
         expect_in=["Приложения"], config=_CFG_DEFAULT,
         note="двойное №№ — особый случай"),
    Case("auditor_pril_no", "В приложение №6 включён реестр.", "fixed",
         expect_in=["приложение 6"], config=_CFG_DEFAULT),
    Case("auditor_pril_no", "Согласно приложениям № 8 и 9.", "fixed",
         expect_in=["приложениям 8"], config=_CFG_DEFAULT),
]


# ─── O. Аудит: т.е./т.ч. ─────────────────────────────────────────────────
O_te_tch = [
    Case("auditor_te_tch", "Документ подписан, т.е. вступил в силу.", "fixed",
         expect_in=["то есть"], expect_not_in=["т.е."], config=_CFG_DEFAULT),
    Case("auditor_te_tch", "В т.ч. НДС 20 %.", "fixed",
         expect_in=["в том числе"], expect_not_in=["т.ч."], config=_CFG_DEFAULT),
    Case("auditor_te_tch", "Все участники, т.е. стороны договора, согласны.", "fixed",
         expect_in=["то есть"], config=_CFG_DEFAULT),
    Case("auditor_te_tch", "Расходы, в т.ч. зарплата, выросли.", "fixed",
         expect_in=["в том числе"], config=_CFG_DEFAULT),
    Case("auditor_te_tch", "Платёж, т.е. сумма, превышает лимит.", "fixed",
         expect_in=["то есть"], config=_CFG_DEFAULT),
    Case("auditor_te_tch", "Премия (в т.ч. НДФЛ) выплачена.", "fixed",
         expect_in=["в том числе"], config=_CFG_DEFAULT),
    Case("auditor_te_tch", "Иванов т.е. директор подписал акт.", "fixed",
         expect_in=["то есть"], config=_CFG_DEFAULT),
    Case("auditor_te_tch", "Сторона А (т.е. подрядчик) обязуется.", "fixed",
         expect_in=["то есть"], config=_CFG_DEFAULT),
    Case("auditor_te_tch", "Аудит, в т.ч. налоговой части, завершён.", "fixed",
         expect_in=["в том числе"], config=_CFG_DEFAULT),
    Case("auditor_te_tch", "Проверены документы, т.е. первичка и регистры.", "fixed",
         expect_in=["то есть"], config=_CFG_DEFAULT),
]


# ─── P. Двойные слова ─────────────────────────────────────────────────────
P_double = [
    Case("double_word_strict", "Это это очень важный пункт.", "fixed",
         expect_in=["Это очень важный"], expect_not_in=["Это это"], config=_CFG_NO_STRICT,
         note="без strict двойное должно исправиться"),
    Case("double_word_strict", "Это это очень важный пункт.", "fixed",
         expect_in=["Это очень важный"], expect_not_in=["Это это"], config=_CFG_DEFAULT,
         note="со strict — БАГ: дубль остаётся; этот case упадёт"),
    Case("double_word_strict", "Сторона по по договору обязана.", "fixed",
         expect_in=["по договору"], expect_not_in=["по по"], config=_CFG_NO_STRICT),
    Case("double_word_strict", "Дату дату нужно указать.", "fixed",
         expect_in=["Дату нужно"], expect_not_in=["Дату дату"], config=_CFG_NO_STRICT),
    Case("double_word_strict", "К контракту приложение приложение №1.", "fixed",
         expect_in=["приложение"], expect_not_in=["приложение приложение"], config=_CFG_NO_STRICT),
    Case("double_word_strict", "На на следующий период.", "fixed",
         expect_in=["На следующий"], expect_not_in=["На на"], config=_CFG_NO_STRICT),
    Case("double_word_strict", "Сумма сумма штрафа составляет.", "fixed",
         expect_in=["Сумма штрафа"], expect_not_in=["Сумма сумма"], config=_CFG_NO_STRICT),
    Case("double_word_strict", "В договоре нет нет особых условий.", "fixed",
         expect_in=["нет особых"], expect_not_in=["нет нет"], config=_CFG_NO_STRICT),
    Case("double_word_strict", "По по итогам проверки замечаний нет.", "fixed",
         expect_in=["По итогам"], expect_not_in=["По по"], config=_CFG_NO_STRICT),
    Case("double_word_strict", "Это — — окончательное решение.", "partial",
         expect_in=["окончательное решение"], config=_CFG_NO_STRICT),
]


# ─── Q. Вставка двоеточия ────────────────────────────────────────────────
Q_colon = [
    Case("colon_insertion", "Документ подписан стороны согласны.", "fixed",
         expect_not_in=[":"], config=_CFG_DEFAULT,
         note="модель часто вставляет двоеточие — должно подавляться"),
    Case("colon_insertion", "Иванов сказал отчёт готов к сдаче.", "fixed",
         expect_not_in=[":"], config=_CFG_DEFAULT),
    Case("colon_insertion", "По итогам проверки замечания отсутствуют.", "unchanged",
         expect_not_in=[":"], config=_CFG_DEFAULT),
    Case("colon_insertion", "Аудитор отметил отчёт соответствует МСФО.", "fixed",
         expect_not_in=[":"], config=_CFG_DEFAULT),
    Case("colon_insertion", "Закон гласит налог 20 процентов.", "fixed",
         expect_not_in=[":"], config=_CFG_DEFAULT),
    Case("colon_insertion", "Контракт включает три части предмет, цена, срок.", "fixed",
         expect_not_in=[":"], config=_CFG_DEFAULT),
    Case("colon_insertion", "Условие простое выплата в течение 30 дней.", "fixed",
         expect_not_in=[":"], config=_CFG_DEFAULT),
    Case("colon_insertion", "Замечание устранено отчёт пересдан.", "fixed",
         expect_not_in=[":"], config=_CFG_DEFAULT),
    Case("colon_insertion", "Подсказка читайте раздел 5.", "fixed",
         expect_not_in=[":"], expect_in=["5"], config=_CFG_DEFAULT),
    Case("colon_insertion", "В таблице указано доход и расход.", "fixed",
         expect_not_in=[":"], config=_CFG_DEFAULT),
]


# ─── R. Риск галлюцинаций ────────────────────────────────────────────────
R_hallu = [
    Case("hallucination_risk", "Вклады в банке «АкадемИнвест» приносят 5 % годовых.", "unchanged",
         expect_in=["АкадемИнвест", "5 %"], config=_CFG_DEFAULT),
    Case("hallucination_risk", "Договор заключён между Хёюнсон Лимитед и Иваном.", "unchanged",
         expect_in=["Хёюнсон", "Иваном"], config=_CFG_DEFAULT),
    Case("hallucination_risk", "Контрагент ʻVisionPay LLCʼ перечислил сумму.", "unchanged",
         expect_in=["VisionPay", "LLC"], config=_CFG_DEFAULT),
    Case("hallucination_risk", "Эксперт BDO выпустил заключение по МСФО IFRS 17.", "unchanged",
         expect_in=["BDO", "МСФО", "IFRS", "17"], config=_CFG_DEFAULT),
    Case("hallucination_risk", "Ассоциация ЕАЭС утвердила Стратегию-2030.", "unchanged",
         expect_in=["ЕАЭС", "Стратегию-2030"], config=_CFG_DEFAULT),
    Case("hallucination_risk", "Свидетель Соколов-Микитов А.А. дал показания.", "unchanged",
         expect_in=["Соколов-Микитов", "А.А."], config=_CFG_DEFAULT),
    Case("hallucination_risk", "Сделка SPA-2024/15 закрыта.", "unchanged",
         expect_in=["SPA-2024/15"], config=_CFG_DEFAULT),
    Case("hallucination_risk", "ИП Pavlovsky предоставил услуги.", "unchanged",
         expect_in=["Pavlovsky"], config=_CFG_DEFAULT),
    Case("hallucination_risk", "Контрагент Шольц-Müller GmbH в реестре.", "unchanged",
         expect_in=["Шольц", "Müller", "GmbH"], config=_CFG_DEFAULT),
    Case("hallucination_risk", "Покупатель — концерн «РосТех-Калашников».", "unchanged",
         expect_in=["РосТех-Калашников"], config=_CFG_DEFAULT),
]


ALL_CASES = (
    A_simple + B_complex + C_comma_missing + D_comma_extra +
    E_protect_numbers + F_protect_abbrev + G_blocklist + H_initials +
    I_quotes + J_yo + K_case + L_rub + M_tysm + N_pril + O_te_tch +
    P_double + Q_colon + R_hallu
)


# ─── S. Sentence split — отдельный корпус ─────────────────────────────────
# Список пар (text, expected_count, expected_chunks_optional)
SPLIT_CASES = [
    ("А.С. Пушкин родился в 1799 году.", 1,
     ["А.С. Пушкин родился в 1799 году."]),
    ("См. рис. 5 на странице 12.", 1,
     ["См. рис. 5 на странице 12."]),
    ("Это т.е. означает, что счёт открыт.", 1,
     ["Это т.е. означает, что счёт открыт."]),
    ("Дом 5. Этаж 3. Кв 12.", 3,
     ["Дом 5.", "Этаж 3.", "Кв 12."]),
    ("Он сказал... но не закончил мысль.", 1,
     ["Он сказал... но не закончил мысль."]),
    ("Что?! Не верю!!! Иду.", 3,
     ["Что?!", "Не верю!!!", "Иду."]),
    ("Иванов И.И. подписал. Петров П.П. — нет.", 2,
     ["Иванов И.И. подписал.", "Петров П.П. — нет."]),
    ("Согласно п. 1.2.3. договора.", 1,
     ["Согласно п. 1.2.3. договора."]),
    ("Договор № 5. Срок 30 дней.", 2,
     ["Договор № 5.", "Срок 30 дней."]),
    ("Текст\nНовый параграф.", 2,
     ["Текст", "Новый параграф."]),
    ("Слово.слово2.", 1,
     ["Слово.слово2."]),
    ("«Цитата конец.» Следующее.", 2, None),
    ("Файл config.json в каталоге.", 1,
     ["Файл config.json в каталоге."]),
    ("5.6 случаев в год.", 1,
     ["5.6 случаев в год."]),
    ("ст. 159 УК РФ.", 1,
     ["ст. 159 УК РФ."]),
]
