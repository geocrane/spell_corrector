"""
Модуль проверки орфографии с T5 моделью.

Предоставляет функции для:
- Асинхронной проверки текста через ML-модель (T5 с PEFT-адаптерами)
- Защиты защищённых токенов (числа, аббревиатуры, валюта, буллиты)
- Нормализации исправлений (подавление е→ё, кавычек, двоеточий, инициалов)
- Аудиторского формата (№5 → 5, т.е. → то есть, руб. → ₽)
"""

import difflib
import logging
import os
import re
import sys
import threading

# ВАЖНО: torch, transformers, peft импортируются лениво в load_model()
# чтобы не замедлять запуск приложения.

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
MODEL_PATH = os.path.join(_THIS_DIR, "model", "sage-fredt5-distilled-95m")
ADAPTERS_DIR = os.path.join(_THIS_DIR, "adapters")

_LOG_FILE = os.path.join(_THIS_DIR, "spell_debug.log")
logger = logging.getLogger("spell_checker")
logger.setLevel(logging.DEBUG)
_fh = logging.FileHandler(_LOG_FILE, encoding="utf-8", mode="a")
_fh.setFormatter(logging.Formatter(
    "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
))
logger.addHandler(_fh)


def discover_adapters():
    """Сканирует папку adapters/ и возвращает отсортированный список имён адаптеров.

    Адаптером считается директория, содержащая adapter_config.json на верхнем уровне.

    Returns:
        List[str]: Отсортированный список имён адаптеров.
    """
    if not os.path.isdir(ADAPTERS_DIR):
        logger.warning("Adapters directory not found: %s", ADAPTERS_DIR)
        return []
    adapters = []
    for name in os.listdir(ADAPTERS_DIR):
        path = os.path.join(ADAPTERS_DIR, name)
        if os.path.isdir(path) and os.path.isfile(os.path.join(path, "adapter_config.json")):
            adapters.append(name)
    adapters.sort()
    logger.info("Discovered adapters: %s", adapters)
    return adapters


def is_model_loaded():
    """Проверить, загружена ли модель (для индикатора загрузки).

    Returns:
        bool: True если модель загружена.
    """
    return SpellChecker.get_instance().is_model_loaded()


class SpellChecker:
    """Singleton для проверки орфографии.

    Модель загружается один раз при первом вызове load_model()
    и кэшируется (guard: if self.is_loaded and self.model is not None).

    Атрибуты:
        model: T5 модель для генерации исправлений.
        tokenizer: Токенизатор для модели.
        device: "cuda" или "cpu".
        is_loaded: Флаг загрузки модели.
        current_adapter_name: Имя текущего адаптера.
        _active_peft_name: Внутреннее имя активного PEFT адаптера.
    """

    _instance = None
    _lock = threading.Lock()

    def __init__(self):
        self.model = None
        self.tokenizer = None
        self._device = None  # ленивая инициализация
        self.is_loaded = False
        self.current_adapter_name = None
        self._active_peft_name = None

    @property
    def device(self):
        """Вернуть устройство (ленивая проверка CUDA)."""
        if self._device is None:
            import torch
            self._device = "cuda" if torch.cuda.is_available() else "cpu"
        return self._device

    @classmethod
    def get_instance(cls):
        """Вернуть единственный экземпляр класса (Singleton)."""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def _validate_model_files(self, adapter_path):
        """Заглушка валидации ключевых файлов модели (всегда pass)."""
        pass

    def load_model(self, adapter_name=None):
        """Загрузить модель и токенизатор.

        При повторном вызове (модель уже загружена) — мгновенный возврат.
        Параметр adapter_name игнорируется; переключение адаптеров не реализовано.

        Thread-safe: блокировка предотвращает двойную загрузку при одновременном
        вызове из фонового потока (предзагрузка) и worker'а (проверка).

        Args:
            adapter_name: Имя адаптера (не используется).
        """
        with self._lock:
            if self.is_loaded and self.model is not None:
                return

            import torch
            from transformers import AutoModelForSeq2SeqLM, AutoTokenizer

            # Логирование при первой загрузке (ранее было на уровне модуля)
            logger.info("=" * 50)
            logger.info("spell_checker loaded | Python %s | torch %s | CUDA: %s",
                        sys.version.split()[0], torch.__version__, torch.cuda.is_available())
            logger.info("MODEL_PATH: %s (exists: %s)", MODEL_PATH, os.path.isdir(MODEL_PATH))
            logger.info("ADAPTERS_DIR: %s (exists: %s)", ADAPTERS_DIR, os.path.isdir(ADAPTERS_DIR))

            self.tokenizer = AutoTokenizer.from_pretrained("model/sage-fredt5-distilled-95m")
            self.model = AutoModelForSeq2SeqLM.from_pretrained("model/sage-fredt5-distilled-95m")
            self.is_loaded = True

    def is_model_loaded(self):
        """Проверить, загружена ли модель (для индикатора загрузки).

        Returns:
            bool: True если модель загружена.
        """
        return self.is_loaded and self.model is not None

    def _switch_adapter(self, name):
        """Переключить адаптер без перезагрузки базовой модели.

        Метод существует, но нигде не вызывается.

        Args:
            name: Имя адаптера для переключения.
        """
        from peft import PeftModel  # noqa: F811 — ленивый импорт

        with self._lock:
            adapter_path = os.path.join(ADAPTERS_DIR, name)
            self._validate_model_files(adapter_path)

            old_name = self._active_peft_name
            logger.info("Switching adapter: %s -> %s", old_name, name)

            self.model.load_adapter(adapter_path, adapter_name=name)
            self.model.set_adapter(name)

            if old_name and old_name != name:
                self.model.delete_adapter(old_name)
                logger.info("Deleted old adapter: %s", old_name)

            self.current_adapter_name = name
            self._active_peft_name = name
            logger.info("Adapter switched to: %s", name)

    def check(self, text):
        """Проверить текст и вернуть исправленный вариант.

        Токенизирует входной текст, генерирует исправление через model.generate(),
        декодирует результат.

        Args:
            text: Текст для проверки.

        Returns:
            str: Исправленный текст.
        """
        logger.debug("check() input: %s", text[:80])
        inputs = self.tokenizer(text, max_length=None, padding="longest", truncation=False, return_tensors="pt")
        outputs = self.model.generate(**inputs.to(self.model.device), max_length=inputs["input_ids"].size(1) * 1.5)
        result = self.tokenizer.batch_decode(outputs, skip_special_tokens=True)
        return result[0]


# ─── Очистка текста ────────────────────────────────────────────────────

# Управляющие и невидимые символы, которые нужно удалить
_INVISIBLE_RE = re.compile(
    '['
    '\x00-\x08'     # C0 control chars (включая \x02 — маркер сноски Word)
    '\x0b\x0c'      # вертикальная табуляция, form feed
    '\x0e-\x1f'     # остальные C0
    '\x7f'           # DEL
    '\u00ad'         # soft hyphen
    '\u200b-\u200f'  # zero-width spaces, direction marks
    '\u2028\u2029'   # line/paragraph separator
    '\u202a-\u202e'  # direction embedding
    '\u2060-\u2064'  # invisible operators
    '\ufeff'         # BOM / zero-width no-break space
    '\ufff0-\ufff8'  # specials
    ']+'
)


def _sanitize_text(text):
    """Удалить невидимые символы и нормализовать пробелы.

    Args:
        text: Исходный текст.

    Returns:
        str: Очищенный текст.
    """
    text = _INVISIBLE_RE.sub('', text)
    text = re.sub(r'[\t\u00a0\u202f\u2007\u2008\u2009\u200a\u205f\u3000]+', ' ', text)
    text = re.sub(r' {2,}', ' ', text)
    return text.strip()


# ─── Защита токенов ────────────────────────────────────────────────────

_CURRENCY_SYMBOLS = set('₽$€£¥')
_PROTECTED_RE = re.compile(r'\d')
# Символы-буллиты: дефис-минус, en-dash, em-dash, math minus, bullet, triangular bullet
_BULLET_CHARS = set('-\u2013\u2014\u2212\u2022\u2023\u2043\u25e6\u25aa\u25ab\u2027')


def _is_protected_token(token):
    """Определить, должен ли токен быть защищён от изменений модели.

    Защищённые токены: числа, валюты, номера (№), буллиты, ЗАГЛАВНЫЕ аббревиатуры.

    Args:
        token: Токен для проверки.

    Returns:
        bool: True если токен должен быть защищён.
    """
    if _PROTECTED_RE.search(token):
        return True
    if any(c in _CURRENCY_SYMBOLS for c in token):
        return True
    if '№' in token:
        return True
    if token in _BULLET_CHARS:
        return True
    stripped = token.strip('.,;:!?…—–-()[]{}"\'"«»')
    if len(stripped) >= 2 and stripped.isupper() and stripped.isalpha():
        return True
    return False


def _protect_tokens(original, corrected):
    """Откатить изменения модели для защищённых токенов.

    Использует difflib.SequenceMatcher для выравнивания токенов, затем
    заменяет изменённые защищённые токены на оригинальные.

    Args:
        original: Исходный текст.
        corrected: Исправленный моделью текст.

    Returns:
        str: Текст с восстановленными защищёнными токенами.
    """
    if original == corrected:
        return corrected

    orig_words = original.split()
    corr_words = corrected.split()

    sm = difflib.SequenceMatcher(None, orig_words, corr_words)
    result = []

    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'equal':
            result.extend(orig_words[i1:i2])
        elif op == 'replace':
            if any(_is_protected_token(w) for w in orig_words[i1:i2]):
                result.extend(orig_words[i1:i2])
            else:
                result.extend(corr_words[j1:j2])
        elif op == 'delete':
            if any(_is_protected_token(w) for w in orig_words[i1:i2]):
                result.extend(orig_words[i1:i2])
        elif op == 'insert':
            result.extend(corr_words[j1:j2])

    return ' '.join(result)


def _protect_word_blocklist(original, corrected, blocklist):
    """Откатить изменения модели для слов из пользовательского блоклиста.

    Args:
        original: Исходный текст.
        corrected: Исправленный моделью текст.
        blocklist: Список слов для защиты (case-insensitive).

    Returns:
        str: Текст с восстановленными словами из блоклиста.
    """
    if not blocklist or original == corrected:
        return corrected

    blocklist_lower = {w.lower() for w in blocklist}

    orig_words = original.split()
    corr_words = corrected.split()

    sm = difflib.SequenceMatcher(None, orig_words, corr_words)
    result = []

    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'equal':
            result.extend(orig_words[i1:i2])
        elif op in ('replace', 'delete'):
            if any(w.lower() in blocklist_lower for w in orig_words[i1:i2]):
                result.extend(orig_words[i1:i2])
            else:
                if op == 'replace':
                    result.extend(corr_words[j1:j2])
        elif op == 'insert':
            result.extend(corr_words[j1:j2])

    return ' '.join(result)


# ─── Строгая защита ────────────────────────────────────────────────────

def _strict_protect(original, corrected):
    """Откатить вставки и удаления слов — разрешить только замены и равенство.

    Args:
        original: Исходный текст.
        corrected: Исправленный моделью текст.

    Returns:
        str: Текст без вставок и удалений.
    """
    if original == corrected:
        return corrected
    orig_words = original.split()
    corr_words = corrected.split()
    sm = difflib.SequenceMatcher(None, orig_words, corr_words)
    result = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'equal':
            result.extend(orig_words[i1:i2])
        elif op == 'replace':
            result.extend(corr_words[j1:j2])
        elif op == 'delete':
            result.extend(orig_words[i1:i2])
        elif op == 'insert':
            pass
    return ' '.join(result)


def _apply_only_comma_changes(original, corrected_pre_norm):
    """Перенести из corrected_pre_norm в original ТОЛЬКО удаление хвостовых запятых.

    Все остальные изменения (кавычки, регистр, е→ё) подавляются.

    Args:
        original: Исходный текст.
        corrected_pre_norm: Исправленный текст до нормализации.

    Returns:
        str: Текст с применёнными только изменениями запятых.
    """
    orig_words = original.split()
    pre_words = corrected_pre_norm.split()
    sm = difflib.SequenceMatcher(None, orig_words, pre_words)
    result = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'equal':
            result.extend(orig_words[i1:i2])
        elif op == 'delete':
            result.extend(orig_words[i1:i2])
        elif op == 'insert':
            pass
        elif op == 'replace':
            orig_block = orig_words[i1:i2]
            pre_block = pre_words[j1:j2]
            for ow, pw in zip(orig_block, pre_block):
                if ow.endswith(',') and ow[:-1].lower() == pw.lower():
                    result.append(ow[:-1])
                else:
                    result.append(ow)
            if len(orig_block) > len(pre_block):
                result.extend(orig_block[len(pre_block):])
    return ' '.join(result)


# ─── Подавление конкретных изменений ───────────────────────────────────

_TRAILING_PUNCT = set('.,;:!?…—–-')


def _suppress_colon_insertion(original, corrected):
    """Откатить вставки двоеточия, если в оригинале его не было.

    Args:
        original: Исходный текст.
        corrected: Исправленный текст.

    Returns:
        str: Текст без добавленных двоеточий.
    """
    if ':' not in corrected or ':' in original:
        return corrected
    orig_words = original.split()
    corr_words = corrected.split()
    sm = difflib.SequenceMatcher(None, orig_words, corr_words)
    result = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'equal':
            result.extend(orig_words[i1:i2])
        elif op == 'replace':
            new_chunk = corr_words[j1:j2]
            if any(':' in w for w in new_chunk) and not any(':' in w for w in orig_words[i1:i2]):
                result.extend(orig_words[i1:i2])
            else:
                result.extend(new_chunk)
        elif op == 'delete':
            result.extend(orig_words[i1:i2])
        elif op == 'insert':
            new_chunk = corr_words[j1:j2]
            if any(':' in w for w in new_chunk):
                pass
            else:
                result.extend(new_chunk)
    return ' '.join(result)


_YO_MAP = str.maketrans('ёЁ', 'еЕ')


def _suppress_yo_replacement(original, result):
    """Откатить замену е→ё через посимвольный diff.

    Args:
        original: Исходный текст.
        result: Исправленный текст.

    Returns:
        str: Текст с заменённой ё на е (где модель изменила).
    """
    orig_chars = list(original)
    res_chars = list(result)
    sm = difflib.SequenceMatcher(None, orig_chars, res_chars)
    out = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'equal':
            out.extend(orig_chars[i1:i2])
        elif op == 'replace':
            for o, r in zip(orig_chars[i1:i2], res_chars[j1:j2]):
                if o.translate(_YO_MAP) == r.translate(_YO_MAP):
                    out.append(o)
                else:
                    out.append(r)
            orig_len = i2 - i1
            res_len = j2 - j1
            if res_len > orig_len:
                for r in res_chars[j1 + orig_len:j2]:
                    out.append(r.translate(_YO_MAP) if r in 'ёЁ' else r)
            elif orig_len > res_len:
                pass
        elif op == 'insert':
            for r in res_chars[j1:j2]:
                out.append(r.translate(_YO_MAP) if r in 'ёЁ' else r)
        elif op == 'delete':
            pass
    return ''.join(out)


_ALL_QUOTES = set('«»„\u201c\u201d\u2018\u2019\u2039\u203a"\'')


def _suppress_quote_changes(original, corrected):
    """Откатить замену одного типа кавычек на другой.

    Args:
        original: Исходный текст.
        corrected: Исправленный текст.

    Returns:
        str: Текст с оригинальными кавычками.
    """
    if original == corrected:
        return corrected
    matcher = difflib.SequenceMatcher(None, original, corrected)
    out = []
    for op, i1, i2, j1, j2 in matcher.get_opcodes():
        if op == 'equal':
            out.append(original[i1:i2])
        elif op == 'replace':
            orig_chunk = original[i1:i2]
            corr_chunk = corrected[j1:j2]
            chars = []
            for o, c in zip(orig_chunk, corr_chunk):
                if o in _ALL_QUOTES and c in _ALL_QUOTES:
                    chars.append(o)
                else:
                    chars.append(c)
            out.append(''.join(chars))
            if len(corr_chunk) > len(orig_chunk):
                out.append(corr_chunk[len(orig_chunk):])
        elif op == 'delete':
            out.append(original[i1:i2])
        elif op == 'insert':
            out.append(corrected[j1:j2])
    return ''.join(out)


def _strict_protect_quotes(original, corrected):
    """Полная защита кавычек — все кавычки из оригинала сохраняются.

    Args:
        original: Исходный текст.
        corrected: Исправленный текст.

    Returns:
        str: Текст с сохранёнными оригинальными кавычками.
    """
    if original == corrected:
        return corrected
    orig_chars = list(original)
    corr_chars = list(corrected)
    sm = difflib.SequenceMatcher(None, orig_chars, corr_chars)
    out = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'equal':
            out.extend(orig_chars[i1:i2])
        elif op == 'replace':
            for o, c in zip(orig_chars[i1:i2], corr_chars[j1:j2]):
                if o in _ALL_QUOTES or c in _ALL_QUOTES:
                    out.append(o)
                else:
                    out.append(c)
            orig_len = i2 - i1
            corr_len = j2 - j1
            if corr_len > orig_len:
                for c in corr_chars[j1 + orig_len:j2]:
                    if c not in _ALL_QUOTES:
                        out.append(c)
        elif op == 'insert':
            for c in corr_chars[j1:j2]:
                if c not in _ALL_QUOTES:
                    out.append(c)
        elif op == 'delete':
            out.extend(orig_chars[i1:i2])
    return ''.join(out)


_INITIALS_RE = re.compile(r'(?=([А-Яа-яЁёA-Za-z]\.[А-Яа-яЁёA-Za-z]))')


def _suppress_space_in_initials(original, corrected):
    """Откатить вставку пробелов в инициалах (А.С. → А. С.).

    Args:
        original: Исходный текст.
        corrected: Исправленный текст.

    Returns:
        str: Текст с восстановленными инициалами.
    """
    if original == corrected:
        return corrected
    for m in _INITIALS_RE.finditer(original):
        pattern = m.group(1)
        spaced = pattern[0] + '. ' + pattern[2]
        corrected = corrected.replace(spaced, pattern)
    return corrected


# ─── Общая нормализация ────────────────────────────────────────────────

def _normalize_corrected(original, corrected, strict=False):
    """Нормализовать исправленный текст: подавить косметические изменения.

    Шаги:
    1. Привести регистр первой буквы к оригиналу.
    2. Заменить хвостовую пунктуацию на оригинальную.
    3. Подавить замены е↔ё.
    4. Запретить смену заглавной → строчной.
    5. Подавить вставку двоеточия.
    6. Подавить замену типа кавычек.
    7. Подавить вставку пробелов в инициалах.

    Args:
        original: Исходный текст.
        corrected: Исправленный текст.
        strict: Если True — использовать строгую защиту кавычек.

    Returns:
        str: Нормализованный текст.
    """
    if not original or not corrected:
        return corrected

    result = corrected

    # 1. Привести регистр первой буквы к оригиналу
    oi = next((i for i, c in enumerate(original) if c.isalpha()), -1)
    ri = next((i for i, c in enumerate(result) if c.isalpha()), -1)
    if (oi >= 0 and ri >= 0
            and result[ri].lower() == original[oi].lower()):
        result = result[:ri] + original[oi] + result[ri + 1:]

    # 2. Заменить хвостовую пунктуацию на оригинальную
    i = len(original)
    while i > 0 and original[i - 1] in _TRAILING_PUNCT:
        i -= 1
    orig_trail = original[i:]

    j = len(result)
    while j > 0 and result[j - 1] in _TRAILING_PUNCT:
        j -= 1
    result = result[:j] + orig_trail

    # 3. Подавить замены е↔ё
    if result.translate(_YO_MAP) == original.translate(_YO_MAP):
        return original
    result = _suppress_yo_replacement(original, result)

    # 4. Запретить смену заглавной → строчной
    orig_chars = list(original)
    res_chars = list(result)
    sm = difflib.SequenceMatcher(None, orig_chars, res_chars)
    out = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'equal':
            out.extend(res_chars[j1:j2])
        elif op == 'replace':
            for o, r in zip(orig_chars[i1:i2], res_chars[j1:j2]):
                if o.isupper() and r.islower() and o.lower() == r:
                    out.append(o)
                else:
                    out.append(r)
            res_len = j2 - j1
            orig_len = i2 - i1
            if res_len > orig_len:
                out.extend(res_chars[j1 + orig_len:j2])
        elif op == 'insert':
            out.extend(res_chars[j1:j2])
        elif op == 'delete':
            pass
    result = ''.join(out)

    # 5. Подавить вставку двоеточия
    result = _suppress_colon_insertion(original, result)

    # 6. Подавить замену типа кавычек
    if strict:
        result = _strict_protect_quotes(original, result)
    else:
        result = _suppress_quote_changes(original, result)

    # 7. Подавить вставку пробелов в инициалах
    result = _suppress_space_in_initials(original, result)

    return result


# ─── Аудиторский формат ────────────────────────────────────────────────

def _expand_abbreviation(text, pattern, replacement):
    """Раскрыть сокращение, сохраняя регистр первой буквы.

    Args:
        text: Текст для обработки.
        pattern: Regex-паттерн для поиска сокращения.
        replacement: Строка замены.

    Returns:
        str: Текст с раскрытым сокращением.
    """
    def _replacer(m):
        matched = m.group()
        if matched[0].isupper():
            return replacement[0].upper() + replacement[1:]
        return replacement
    return re.compile(pattern, re.IGNORECASE).sub(_replacer, text)


def _apply_auditor_format(text):
    """Применить правила аудиторского формата.

    Правила:
    1. Приложение №5 → Приложение 5
    2. т.е. → то есть, в т.ч. → в том числе
    3. руб./рубли/рублей → ₽
    4. млн. ₽ → млн ₽, млрд. ₽ → млрд ₽; тыс ₽ → тыс. ₽
    5. № 1234 → №1234

    Args:
        text: Исходный текст.

    Returns:
        str: Текст в аудиторском формате.
    """
    text = re.sub(r'([Пп]риложени[еяюий])\s*№\s*', r'\1 ', text)
    text = _expand_abbreviation(text, r'\bв\s+т\.ч\.', 'в том числе')
    text = _expand_abbreviation(text, r'\bт\.е\.', 'то есть')
    text = re.sub(r'\bруб\.', '₽', text)
    text = re.sub(r'\bрублей\b', '₽', text)
    text = re.sub(r'\bрубли\b', '₽', text)
    text = re.sub(r'\bмлн\.\s*₽', 'млн ₽', text)
    text = re.sub(r'\bмлрд\.\s*₽', 'млрд ₽', text)
    text = re.sub(r'\bтыс\s+₽', 'тыс. ₽', text)
    text = re.sub(r'№\s+(\d)', r'№\1', text)
    return text


# ─── Асинхронная проверка ──────────────────────────────────────────────

def check_sentences_async(sentences, on_progress, on_complete, on_error=None, on_start=None,
                          adapter_name=None, strict=False, auditor_format=False, word_blocklist=None):
    """Асинхронная проверка списка предложений в фоновом потоке.

    Поток обработки для каждого предложения:
    1. _sanitize_text() — очистка от невидимых символов.
    2. checker.check() — генерация исправления моделью.
    3. _protect_tokens() — защита чисел, валют, аббревиатур.
    4. _protect_word_blocklist() — защита слов из блоклиста.
    5. _strict_protect() — запрет вставок/удалений (если strict).
    6. _normalize_corrected() — нормализация (ё→е, кавычки, двоеточия).
    7. _apply_auditor_format() — аудиторский формат (если включён).

    Args:
        sentences: Список предложений с ключами "text" и "index".
        on_progress: Callback(index, original, corrected, has_error).
        on_complete: Callback при завершении.
        on_error: Callback(error_message) при ошибке.
        on_start: Callback(index) перед проверкой каждого предложения.
        adapter_name: Имя адаптера (игнорируется).
        strict: Строгая защита (запрет вставок/удалений).
        auditor_format: Применять аудиторский формат.
        word_blocklist: Список слов для защиты.

    Returns:
        threading.Event: cancel_event — установить для остановки worker'а.
    """
    cancel_event = threading.Event()

    def worker():
        try:
            logger.info("worker started, %d sentences, adapter=%s", len(sentences), adapter_name)
            checker = SpellChecker.get_instance()
            checker.load_model(adapter_name=adapter_name)

            for sentence in sentences:
                if cancel_event.is_set():
                    logger.info("worker cancelled")
                    return

                idx = sentence["index"]
                if on_start:
                    on_start(idx)
                original = _sanitize_text(sentence["text"])
                logger.info("Checking sentence %d/%d", idx + 1, len(sentences))
                logger.debug("after sanitize:     %r", original[:120])
                corrected = checker.check(original).strip()
                logger.debug("raw model output:   %r", corrected[:120])
                corrected = _protect_tokens(original, corrected)
                logger.debug("after protect_tok:  %r", corrected[:120])
                if word_blocklist:
                    corrected = _protect_word_blocklist(original, corrected, word_blocklist)
                    logger.debug("after blocklist:    %r", corrected[:120])
                effective_strict = strict or auditor_format
                if effective_strict:
                    corrected = _strict_protect(original, corrected)
                    logger.debug("after strict_prot:  %r", corrected[:120])
                corrected_pre_norm = corrected
                corrected = _normalize_corrected(original, corrected, strict=effective_strict)
                logger.debug("after normalize:    %r", corrected[:120])
                if auditor_format:
                    corrected = _apply_auditor_format(corrected)
                    logger.debug("after audit_fmt:    %r", corrected[:120])

                if cancel_event.is_set():
                    logger.info("worker cancelled after check")
                    return

                has_error = corrected != original
                if not has_error:
                    def _mid_commas(text):
                        return text.rstrip('.,;:!?…—–- ').count(',')
                    if _mid_commas(original) != _mid_commas(corrected_pre_norm):
                        has_error = True
                        corrected = _apply_only_comma_changes(original, corrected_pre_norm)
                        if corrected == original:
                            logger.debug("_apply_only_comma_changes: no word-level comma diff found, "
                                         "falling back to pre_norm")
                            corrected = corrected_pre_norm
                        else:
                            logger.debug("has_error forced True: mid-comma count changed "
                                         "(%d→%d), applied only-comma fix", _mid_commas(original), _mid_commas(corrected_pre_norm))
                on_progress(idx, original, corrected, has_error)

            on_complete()
            logger.info("worker finished OK")
        except Exception as e:
            logger.exception("worker failed")
            if on_error:
                on_error(str(e))

    threading.Thread(target=worker, daemon=True).start()
    return cancel_event
