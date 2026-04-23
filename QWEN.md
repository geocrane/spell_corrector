# QWEN.md — Corrector 1.4

## Обзор проекта

**Corrector 1.4** — десктопное приложение для автоматизированной проверки орфографии и грамматики в документах Microsoft Word и Outlook. Использует LLM-модель T5 (`sage-fredt5-distilled-95m`) с PEFT-адаптерами для качественного исправления ошибок.

**Платформа:** Только Windows (требуется COM-интеграция с Microsoft Office).

**Технологии:**
- **Python 3.8+**
- **tkinter** — GUI
- **torch, transformers, peft** — ML-модель
- **pywin32** — COM-автоматизация Windows
- **difflib** — визуальное сравнение (diff)

---

## Структура проекта

```
corrector_1.4/
├── main.py                    # Точка входа (~60 строк) — логирование, Engine, MainWindow
├── office_finder.py           # COM-утилиты для работы с Word/Outlook (НЕ менять)
├── spell_checker.py           # ML-проверка орфографии (НЕ менять)
├── config.json                # Конфигурация: адаптеры, блок-лист, strict_protection
├── run.bat                    # BAT-файл для запуска на Windows
├── QWEN.md                    # Этот файл
│
├── core/                      # БИЗНЕС-ЛОГИКА (без tkinter)
│   ├── __init__.py
│   ├── events.py              # EventBus — паттерн событий
│   ├── config.py              # Загрузка/сохранение config.json
│   ├── doc_state.py           # DocStateCache — кэш состояний документов
│   ├── engine.py              # Engine — главный контроллер бизнес-логики
│   └── providers/             # Провайдеры документов (новая архитектура)
│       ├── __init__.py
│       ├── base.py            # Базовый класс провайдера
│       ├── word_provider.py   # Провайдер для Word
│       ├── outlook_provider.py # Провайдер для Outlook
│       ├── excel_provider.py  # Провайдер для Excel
│       └── registry.py        # Реестр провайдеров
│
└── ui/                        # ПРЕДСТАВЛЕНИЕ (tkinter)
    ├── __init__.py
    ├── constants.py           # Цвета diff, спиннеры, пунктуация
    ├── tiles.py               # Функции создания плиток (документы, предложения)
    └── main_window.py         # MainWindow — главное окно, подписка на события Engine
```

---

## Архитектура: Event Bus

Проект прошёл рефакторинг (апрель 2026) — разделены UI и бизнес-логика через паттерн **Event Bus**.

### Направление зависимостей (без циклов)

```
ui/main_window.py  →  core/engine.py  →  core/events.py
ui/tiles.py        →  ui/constants.py     core/config.py
core/engine.py     →  core/doc_state.py   core/config.py
core/engine.py     →  office_finder, spell_checker
```

**Правило:** `engine.py` НЕ импортирует tkinter. Вся коммуникация с UI — через `EventBus.emit()`.

### События (Engine → View)

| Событие | Данные | Описание |
|---------|--------|----------|
| `documents_found` | `documents: list[dict]` | Найдены открытые документы |
| `documents_not_found` | — | Документы не найдены |
| `check_started` | `total: int` | Началась проверка |
| `sentence_start` | `index: int` | Проверка предложения началась |
| `sentence_checked` | `index, original, corrected, has_error` | Предложение проверено |
| `check_complete` | — | Проверка завершена |
| `check_error` | `error: str` | Ошибка при проверке |

### Команды (View → Engine)

| Метод | Описание |
|-------|----------|
| `engine.find_documents()` | Поиск открытых документов |
| `engine.select_document(doc)` | Выбор документа |
| `engine.check_document()` | Проверка всего документа |
| `engine.check_fragment()` | Проверка выделенного фрагмента |
| `engine.apply_correction(index)` | Применение исправления |
| `engine.revert_correction(index)` | Откат исправления |
| `engine.toggle_skip(index)` | Переключение пропуска |
| `engine.set_config(key, value)` | Сохранение настройки |
| `engine.cancel()` | Отмена проверки |
| `engine.navigate_to_sentence(sentence)` | Навигация к предложению |

---

## Запуск и установка

### Установка зависимостей

```bash
pip install torch transformers peft pywin32
```

### Запуск

```bash
python main.py
```

Или через `run.bat` на Windows (путь в BAT-файле может требовать корректировки под конкретную машину).

### Требования к модели

- Модель: `model/sage-fredt5-distilled-95m/`
- Адаптеры: `adapters/lora_adapter/`

---

## Конфигурация (config.json)

```json
{
  "default_adapter": "lora_adapter_v2.1",
  "hide_clean_sentences": true,
  "strict_protection": true,
  "auditor_format": true,
  "skip_tables": false,
  "word_blocklist": ["Охват"]
}
```

| Параметр | Описание |
|----------|----------|
| `default_adapter` | PEFT-адаптер по умолчанию |
| `hide_clean_sentences` | Скрывать предложения без изменений |
| `strict_protection` | Строгая защита чисел, аббревиатур, валют |
| `auditor_format` | Нормализация сокращений (руб. → ₽) |
| `skip_tables` | Пропускать таблицы при проверке |
| `word_blocklist` | Слова, которые модель не должна изменять |

---

## Ключевые особенности

- **Интеграция с MS Office** через COM (Word, Outlook, Excel)
- **Diff-сравнение** с цветовым кодированием (удаления/вставки)
- **Сохранение форматирования** при применении исправлений
- **Кэш состояний документов** (`DocStateCache`) — не проверяет повторно
- **Generation counter** — защита от устаревших callback'ов
- **Cancel event** — возможность отмены проверки
- **Логирование** в `spell_debug.log`

---

## Известные ограничения архитектуры

1. Модель загружается заново каждый раз (spell_checker.py)
2. Адаптеры могут не переключаться (adapter_name игнорируется)
3. Fallback 1920×1080 для мультимониторных конфигураций
4. Lambda-замыкания в памяти при создании плиток

---

## Практики разработки

- **Циклические импорты запрещены** — зависимости идут в одном направлении
- **Engine НЕ импортирует tkinter** — чистая бизнес-логика
- **UI подписывается на события** и оборачивает callback'и через `after()`
- **office_finder.py и spell_checker.py** — НЕ менять без явной необходимости
- **Каждый файл синтаксически валиден** — проверять после изменений

---

## Логирование

Все действия и ошибки записываются в `spell_debug.log`. Формат:
```
%(asctime)s [%(levelname)s] [main] %(message)s
```

Глобальные обработчики исключений (`sys.excepthook`, `threading.excepthook`) также логируют в этот файл.
