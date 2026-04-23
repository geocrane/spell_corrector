# План рефакторинга: Разделение UI / Логика (Event Bus)

> **Дата начала:** 9 апреля 2026
> **Подход:** Event Bus (паттерн событий)
> **Платформа:** Только Windows
> **Принцип:** НИЧЕГО не меняется визуально и функционально

---

## Статус выполнения

| Шаг | Файл(ы) | Статус | Примечание |
|-----|---------|--------|------------|
| 1 | `core/events.py` | ✅ Выполнено | EventBus — чистый Python, все тесты прошли |
| 2 | `core/config.py` | ✅ Выполнено | Вынесено, main.py обновлён, синтаксис OK |
| 3 | `core/doc_state.py` | ✅ Выполнено | Вынесено, main.py обновлён, синтаксис OK |
| 4 | `ui/constants.py` | ✅ Выполнено | Вынесено, main.py обновлён, синтаксис OK |
| 5 | `core/engine.py` | ✅ Выполнено | Создан, синтаксис OK (win32com недоступен на macOS) |
| 6 | `ui/tiles.py` | ✅ Выполнено | Создан, синтаксис OK |
| 7 | `ui/main_window.py` | ✅ Выполнено | Создан, синтаксис OK |
| 8 | `main.py` | ✅ Выполнено | Минимизирован, синтаксис OK |

---

## 1. Текущий анализ архитектуры

### Граф зависимостей (ПРОБЛЕМА)

```
main.py (App — 2112 строки)
    ├── imports: office_finder.py (918 строк)
    ├── imports: spell_checker.py (967 строк)
    │
    └── класс App(tk.Tk) содержит ВСЁ одновременно:
         ├── UI виджеты: tk.Button, tk.Frame, tk.Canvas, ttk.Combobox...
         ├── Состояние: check_results, sentences, documents, _doc_cache...
         ├── Бизнес-логику: check_selected_document(), _run_check(),
         │                    _apply_correction_to_doc(), find_documents()...
         ├── Callback'и: _on_tile_click(), _toggle_apply(), _on_sentence_checked()...
         └── Анимации: _animate_spinner(), _stop_spinner()...
```

### Почему простой вынос функций НЕ работает

1. **Циклические зависимости**: методы UI вызывают бизнес-логику (`self.check_selected_document()`), бизнес-логика вызывает методы UI (`self.status_label.config()`, `self._create_sentence_tile()`, `self.after()`)
2. **Смешанное состояние**: `self.check_results` используется и для UI (отрисовка плиток), и для логики (кэш документов)
3. **`self.after()`** — tkinter-специфичный метод, используется в бизнес-логике для планирования callback'ов
4. **Lambda-замыкания** с `gen` (generation counter) для защиты от устаревших callback'ов — привязаны к `self`

### Точки разрыва циклических импортов

**Направление зависимостей (без циклов):**
```
events.py  ←  engine.py  ←  main_window.py  ←  main.py
                ↑
      (импортирует events, config, doc_state)
      (импортирует office_finder, spell_checker)
```

- `engine.py` **НЕ** импортирует tkinter
- `main_window.py` импортирует engine и events
- `office_finder.py` и `spell_checker.py` **без изменений**

---

## 2. Целевая структура файлов

```
corrector_1.4/
├── main.py                    # Точка входа (минимальный, ~20 строк)
├── office_finder.py           # БЕЗ ИЗМЕНЕНИЙ (COM-утилиты)
├── spell_checker.py           # БЕЗ ИЗМЕНЕНИЙ (ML-проверка)
├── config.json                # БЕЗ ИЗМЕНЕНИЙ
├── spell_debug.log            # БЕЗ ИЗМЕНЕНИЙ
│
├── core/                      # БИЗНЕС-ЛОГИКА (без tkinter!)
│   ├── __init__.py
│   ├── events.py              # EventBus (~60 строк) — НОВЫЙ
│   ├── config.py              # load/save config (~40 строк) — ВЫНЕСЕН
│   ├── doc_state.py           # Кэш состояний (~80 строк) — ВЫНЕСЕН
│   └── engine.py              # Главный контроллер (~900 строк) — ВЫНЕСЕН
│
└── ui/                        # ПРЕДСТАВЛЕНИЕ (tkinter)
    ├── __init__.py
    ├── constants.py           # Цвета, спиннеры (~30 строк) — ВЫНЕСЕНЫ
    ├── tiles.py               # Создание плиток (~300 строк) — ВЫНЕСЕНЫ
    └── main_window.py         # Главное окно (~1200 строк) — ВЫНЕСЕНО
```

---

## 3. EventBus — интерфейс взаимодействия

### События (Engine → View)

Engine генерирует, View подписывается:

| Событие | Данные | Когда генерируется | Что делает View |
|---------|--------|-------------------|-----------------|
| `documents_found` | `documents: list[dict]` | После поиска документов | Перерисовывает плитки документов |
| `documents_not_found` | — | Документы не найдены | Показывает «Документы не найдены» |
| `check_started` | `total: int` | Начало проверки | Блокирует кнопки, создаёт плитки со спиннерами |
| `sentence_start` | `index: int` | Перед проверкой предложения | Запускает спиннер на плитке |
| `sentence_checked` | `index, original, corrected, has_error` | После проверки предложения | Обновляет плитку с diff |
| `check_complete` | — | Проверка завершена | Разблокирует кнопки, обновляет статус |
| `check_error` | `error: str` | Ошибка проверки | Показывает ошибку |

### Команды (View → Engine)

View вызывает методы Engine:

| Метод Engine | Аргументы | Что делает |
|---|---|---|
| `find_documents()` | — | Ищет документы, emit `documents_found` |
| `select_document(doc)` | `doc: dict` | Выбирает документ, обновляет кэш |
| `check_document()` | — | Запускает проверку всего документа |
| `check_fragment()` | — | Запускает проверку выделенного фрагмента |
| `apply_correction(index)` | `index: int` | Применяет исправление к документу |
| `revert_correction(index)` | `index: int` | Откатывает исправление |
| `toggle_skip(index)` | `index: int` | Переключает пропуск |
| `set_config(key, value)` | `key: str, value` | Сохраняет настройку в config.json |
| `cancel()` | — | Отменяет текущую проверку |
| `navigate_to_sentence(sentence)` | `sentence: dict` | Навигация к предложению в документе |
| `get_config()` | — | Возвращает текущую конфигурацию |
| `get_selected_adapter()` | — | Возвращает имя выбранного адаптера |
| `get_available_adapters()` | — | Возвращает список доступных адаптеров |

---

## 4. Детальный план по шагам

### ШАГ 1: Создать `core/events.py` (EventBus)

**Что сделать:**
- Класс `EventBus` с методами:
  - `subscribe(event_name: str, callback: Callable)` — подписка
  - `emit(event_name: str, **data)` — генерация события
  - `unsubscribe(event_name: str, callback: Callable)` — отписка
  - `clear()` — очистка всех подписок
- Поддержка нескольких подписчиков на одно событие
- **0 зависимостей** — чистый Python

**Что НЕ меняется:** Никаких изменений в других файлах на этом шаге.

**Проверка:** `python -c "from core.events import EventBus; eb = EventBus(); print('OK')"`

---

### ШАГ 2: Создать `core/config.py`

**Что вынести из `main.py`:**
- `_DEFAULT_CONFIG` (словарь)
- `_load_config()` → `load_config()`
- `_save_config()` → `save_config()`
- `_CONFIG_PATH` → константа

**Изменения в `main.py`:**
- Заменить `_load_config()` на `from core.config import load_config`
- Заменить `_save_config()` на `from core.config import save_config`
- Удалить `_DEFAULT_CONFIG` и `_CONFIG_PATH`

**Что НЕ меняется:** Логика загрузки/сохранения идентична.

---

### ШАГ 3: Создать `core/doc_state.py`

**Что вынести из `main.py`:**
- `_doc_cache` → класс `DocStateCache`
- `_save_doc_state()` → `DocStateCache.save(doc_id, check_results, sentences)`
- `_load_doc_state()` → `DocStateCache.load(doc_id) -> dict | None`
- `_check_generation` → `DocStateCache.generation` (int)
- `_cancel_event` → `DocStateCache.cancel_event`

**Класс `DocStateCache`:**
```python
class DocStateCache:
    def __init__(self):
        self._cache = {}  # {id(doc): {"check_results": {}, "sentences": []}}
        self.generation = 0
        self.cancel_event = None

    def save(self, doc_id, check_results, sentences)
    def load(self, doc_id) -> dict | None
    def clear(self)
    def next_generation(self) -> int
    def cancel(self)
```

**Изменения в `main.py`:**
- Заменить `self._doc_cache` на экземпляр `DocStateCache`
- Заменить вызовы `_save_doc_state()` / `_load_doc_state()`

---

### ШАГ 4: Создать `ui/constants.py`

**Что вынести из `main.py`:**
- `DIFF_ADDED_BG`, `DIFF_ADDED_FG`, `DIFF_REMOVED_BG`, `DIFF_REMOVED_FG`
- `PUNCTUATION`
- `SPINNER_FRAMES`, `SPINNER_DELAY`, `WAITING_SYMBOL`

**Изменения в `main.py`:**
- Импортировать из `ui.constants` вместо локальных определений

---

### ШАГ 5: Создать `core/engine.py` (КЛЮЧЕВОЙ ШАГ)

**Что вынести из `main.py` (бизнес-логика):**

| Метод в App | Метод в Engine | Примечание |
|---|---|---|
| `find_documents()` | `Engine.find_documents()` | Emit `documents_found` / `documents_not_found` |
| `check_selected_document()` | `Engine.check_document()` | Emit `check_started`, `sentence_start`, ... |
| `check_selected_fragment()` | `Engine.check_fragment()` | Аналогично |
| `_run_check()` | `Engine._run_check()` | Worker-поток с callback'ами через EventBus |
| `_apply_correction_to_doc()` | `Engine.apply_correction()` | Вызывает office_finder.replace_sentence_text() |
| `_revert_correction_to_doc()` | `Engine.revert_correction()` | Аналогично |
| `_toggle_apply()` | `Engine.toggle_apply()` | Переключает pending ↔ applied |
| `_toggle_skip()` | `Engine.toggle_skip()` | Переключает pending ↔ skipped |
| `_save_doc_state()` | `DocStateCache.save()` | Делегирует |
| `_load_doc_state()` | `DocStateCache.load()` | Делегирует |
| `_cancel_worker()` | `DocStateCache.cancel()` | |
| `_on_toggle_*` | `Engine.set_config()` | Сохраняет в config.json |

**Конструктор Engine:**
```python
class Engine:
    def __init__(self):
        self.events = EventBus()
        self.config = load_config()
        self.doc_state = DocStateCache()
        self.documents = []
        self.selected_doc = None
        self.is_checking = False
        # ... остальное состояние БЕЗ tkinter
```

**Ключевой принцип:** Engine **НЕ** импортирует tkinter. Вместо `self.status_label.config()` — `self.events.emit('check_started', total=...)`.

**Изменения в `main.py`:**
- Создать экземпляр `Engine`
- Подписаться на события Engine и делегировать в UI

---

### ШАГ 6: Создать `ui/tiles.py`

**Что вынести из `main.py`:**

| Метод в App | Функция в tiles.py | Примечание |
|---|---|---|
| `_create_document_tile()` | `create_document_tile(parent, doc, on_click, is_selected, cached_errors)` | |
| `_create_sentence_tile_checking()` | `create_sentence_tile_checking(parent, sentence, on_click)` | |
| `_create_checked_sentence_tile()` | `create_checked_sentence_tile(parent, sentence, result, on_click, on_apply, on_skip)` | |
| `_create_diff_widget()` | `create_diff_widget(parent, original, corrected, after_callback)` | |
| `_update_sentence_tile()` | `update_sentence_tile(tile, index, original, corrected, has_error, ...)` | |
| `_highlight_selected_tile()` | `highlight_selected_tile(tile_frames, selected_doc)` | |

**Принцип:** Функции принимают все параметры явно, не обращаются к `self.*`.

---

### ШАГ 7: Создать `ui/main_window.py`

**Что перенести из `main.py` (UI-код):**

| Метод в App | Метод в MainWindow | Примечание |
|---|---|---|
| `__init__()` | `MainWindow.__init__(engine)` | Подписка на события engine |
| `_create_ui()` | `MainWindow._create_ui()` | Без изменений |
| `_update_button_state()` | `MainWindow._update_button_state()` | |
| `_refresh_check_buttons()` | `MainWindow._refresh_check_buttons()` | |
| `_animate_spinner()` | `MainWindow._animate_spinner()` | |
| `_stop_spinner()` | `MainWindow._stop_spinner()` | |
| `_apply_sentence_filter()` | `MainWindow._apply_sentence_filter()` | |
| `_should_hide_tile()` | `MainWindow._should_hide_tile()` | |
| `_update_status_text()` | `MainWindow._update_status_text()` | |
| `_on_tile_click()` | `MainWindow._on_tile_click()` | Вызывает engine.select_document() |
| `_on_sentence_click()` | `MainWindow._on_sentence_click()` | Вызывает engine.navigate_to_sentence() |
| `_raise_window()` | `MainWindow._raise_window()` | |
| `_position_on_active_monitor()` | `MainWindow._position_on_active_monitor()` | |
| `_on_frame_configure()` | `MainWindow._on_frame_configure()` | |
| `_on_canvas_configure()` | `MainWindow._on_canvas_configure()` | |
| `_on_mousewheel()` | `MainWindow._on_mousewheel()` | |
| `_set_button_hover()` | `MainWindow._set_button_hover()` | |

**Подписка на события в `__init__`:**
```python
def __init__(self, engine):
    self.engine = engine
    # ... создание UI ...

    # Подписка на события
    engine.events.subscribe("documents_found", self._on_documents_found)
    engine.events.subscribe("documents_not_found", self._on_documents_not_found)
    engine.events.subscribe("check_started", self._on_check_started)
    engine.events.subscribe("sentence_start", self._on_sentence_start)
    engine.events.subscribe("sentence_checked", self._on_sentence_checked)
    engine.events.subscribe("check_complete", self._on_check_complete)
    engine.events.subscribe("check_error", self._on_check_error)
```

**Обёртка callback'ов через `after()`:**
```python
def _on_documents_found(self, documents):
    self.after(0, lambda: self._render_documents(documents))
```

---

### ШАГ 8: Обновить `main.py`

**Целевой вид:**
```python
"""Точка входа приложения Corrector."""

import logging
import os
import sys
import threading as _threading

# Настройка логирования (остаётся здесь — инициализация до всего)
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_LOG_FILE = os.path.join(_THIS_DIR, "spell_debug.log")
logger = logging.getLogger("main_app")
logger.setLevel(logging.DEBUG)
_fh = logging.FileHandler(_LOG_FILE, encoding="utf-8", mode="a")
_fh.setFormatter(logging.Formatter(
    "%(asctime)s [%(levelname)s] [main] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
))
logger.addHandler(_fh)

# Глобальные обработчики исключений (остаются здесь)
def _global_except_hook(exc_type, exc_value, exc_tb):
    logger.critical("Unhandled exception", exc_info=(exc_type, exc_value, exc_tb))
    sys.__excepthook__(exc_type, exc_value, exc_tb)

def _thread_except_hook(args):
    logger.critical("Unhandled thread exception [%s]",
                     args.thread.name if args.thread else "?",
                     exc_info=(args.exc_type, args.exc_value, args.exc_traceback))

sys.excepthook = _global_except_hook
_threading.excepthook = _thread_except_hook

from core.engine import Engine
from ui.main_window import MainWindow


def main():
    logger.info("Application starting...")
    try:
        engine = Engine()
        view = MainWindow(engine)
        view.mainloop()
    except Exception:
        logger.critical("Fatal error", exc_info=True)
        raise
    finally:
        logger.info("Application exiting.")


if __name__ == "__main__":
    main()
```

---

## 5. Как избежать циклических импортов

**Правило:** Зависимости идут ТОЛЬКО в одном направлении:

```
ui/main_window.py  →  core/engine.py  →  core/events.py
ui/tiles.py        →  ui/constants.py     core/config.py
core/engine.py     →  core/doc_state.py   core/config.py
core/engine.py     →  office_finder, spell_checker (существующие)
```

**ЗАПРЕЩЕНО:**
- `engine.py` НЕ импортирует `ui/`
- `events.py` НЕ импортирует ничего
- `config.py` НЕ импортирует ничего кроме `json`, `os`
- `office_finder.py` и `spell_checker.py` БЕЗ ИЗМЕНЕНИЙ

**Точка сборки** — `main.py`:
```python
engine = Engine()        # создаёт EventBus внутри
view = MainWindow(engine) # подписывается на события engine
```

---

## 6. Что НЕ изменится (ГАРАНТИЯ)

| Аспект | Статус |
|---|---|
| Визуальное отображение | ✅ Те же виджеты, цвета, layout, размеры |
| Функционал | ✅ Те же методы, та же логика, те же callback'и |
| Анимация спиннера | ✅ Те же `SPINNER_FRAMES`, `SPINNER_DELAY` |
| Diff-отображение | ✅ Те же `difflib.SequenceMatcher`, теги, цвета |
| COM-интеграция | ✅ `office_finder.py` БЕЗ ИЗМЕНЕНИЙ |
| ML-проверка | ✅ `spell_checker.py` БЕЗ ИЗМЕНЕНИЙ |
| Конфигурация | ✅ Тот же `config.json`, тот же формат |
| Логирование | ✅ Тот же `spell_debug.log`, тот же формат |
| Позиционирование окна | ✅ То же `get_active_monitor_workarea()` |
| Прокрутка | ✅ Та же логика canvas + mousewheel |
| Кэш документов | ✅ Те же данные, тот же механизм |
| Generation counter | ✅ Та же защита от устаревших callback'ов |
| Cancel event | ✅ Та же отмена worker'а |

---

## 7. Риски и митигация

| Риск | Митигация |
|---|---|
| `self.after()` требует tkinter в бизнес-логике | Engine вызывает `emit()` синхронно, View оборачивает в `after()` при подписке |
| Lambda-замыкания с `gen` (generation counter) | Перенести generation counter в `DocStateCache`, View получает `gen` из события |
| Большие файлы при разбиении | Разбивать постепенно, каждый шаг — working state |
| Случайное изменение UI | Каждый шаг — сверка с оригинальным поведением |
| Потеря контекста | Этот файл — полный контекст для продолжения |

---

## 8. Заметки по ходу работы

> Здесь будут записываться проблемы, решения и наблюдения по ходу рефакторинга.

### Шаг 1 (events.py):
- ✅ Создан `core/events.py` с классом `EventBus`
- ✅ Методы: `subscribe()`, `unsubscribe()`, `emit()`, `clear()`, `has_subscribers()`
- ✅ 0 внешних зависимостей — чистый Python
- ✅ Тесты: подписка, несколько подписчиков, отписка, clear, повторная подписка, обработка исключений
- ✅ Исключения в callback'ах логируются и не прерывают остальных подписчиков

### Шаг 2 (config.py):
- ✅ Создан `core/config.py` с `load_config()` и `save_config()`
- ✅ `_DEFAULT_CONFIG`, `_CONFIG_PATH` перенесены
- ✅ `main.py` обновлён — импортирует из `core.config`
- ✅ Синтаксис main.py валиден

### Шаг 3 (doc_state.py):
- ✅ Создан `core/doc_state.py` с классом `DocStateCache`
- ✅ Методы: `save()`, `load()`, `clear()`, `next_generation()`, `cancel()`, `has()`
- ✅ `main.py` обновлён — `self._doc_cache` → `self.doc_state`, `self._check_generation` → `self.doc_state.generation`, `self._cancel_event` → `self.doc_state.cancel_event`
- ✅ Методы `_save_doc_state()`, `_load_doc_state()`, `_cancel_worker()` делегируют в `doc_state`
- ✅ Синтаксис main.py валиден

### Шаг 4 (constants.py):
- ✅ Создан `ui/__init__.py` и `ui/constants.py`
- ✅ Вынесены: цвета diff, PUNCTUATION, SPINNER_FRAMES, SPINNER_DELAY, WAITING_SYMBOL
- ✅ `main.py` импортирует из `ui.constants`
- ✅ Синтаксис main.py валиден

### Шаг 5 (engine.py):
- ✅ Создан `core/engine.py` с классом `Engine`
- ✅ НЕ импортирует tkinter — чистая бизнес-логика
- ✅ Методы: find_documents(), select_document(), check_document(), check_fragment(), apply_correction(), revert_correction(), toggle_skip(), navigate_to_sentence(), set_config(), cancel()
- ✅ Генерирует события через EventBus: documents_found, documents_not_found, check_started, sentence_start, sentence_checked, check_complete, check_error
- ✅ Использует DocStateCache для кэша, config для настроек
- ⚠️ Не тестируется на macOS (win32com недоступен) — проверка на Windows

### Шаг 6 (tiles.py):
- ✅ Создан `ui/tiles.py` с функциями создания плиток
- ✅ Функции: create_document_tile(), highlight_selected_tile(), create_sentence_tile_checking(), create_checked_sentence_tile(), update_sentence_tile(), create_diff_widget(), insert_deleted_text()
- ✅ Все функции принимают параметры явно, не обращаются к self.*
- ✅ Используют константы из ui.constants
- ⚠️ Не тестируется на macOS (tkinter есть, но полная проверка требует Windows)

### Шаг 7 (main_window.py):
- ✅ Создан `ui/main_window.py` с классом `MainWindow(tk.Tk)`
- ✅ Подписывается на события Engine: documents_found, documents_not_found, check_started, sentence_start, sentence_checked, check_complete, check_error
- ✅ Вызывает методы Engine: find_documents(), check_document(), check_fragment(), select_document(), apply_correction(), revert_correction(), toggle_skip(), navigate_to_sentence()
- ✅ Использует функции из ui.tiles для создания плиток
- ✅ Делегирует бизнес-логику в Engine, не содержит ML/COM кода
- ⚠️ Не тестируется на macOS (win32com, tkinter требуют Windows для полной проверки)

### Шаг 8 (main.py):
- ✅ `main.py` минимизирован до ~60 строк
- ✅ Содержит: логирование, обработчики исключений, создание Engine + MainWindow
- ✅ Все 8 файлов синтаксически валидны
- ✅ Циклических импортов нет — зависимости идут в одном направлении
- ⚠️ Полная проверка на Windows (требуется win32com)

---

## 9. Справочная информация

### Ключевые файлы (текущие)
- `main.py` — 2112 строк, класс `App(tk.Tk)`
- `office_finder.py` — 918 строк, COM-утилиты (НЕ менять)
- `spell_checker.py` — 967 строк, ML-проверка (НЕ менять)
- `config.json` — конфигурация

### Зависимости проекта
- `torch`, `transformers`, `peft` — ML
- `pywin32` — COM-автоматизация Windows
- `tkinter` — GUI (встроенный в Python)

### Известные проблемы архитектуры (НЕ решаются в этом рефакторинге)
1. Модель загружается заново каждый раз (spell_checker.py)
2. Адаптеры не переключаются (adapter_name игнорируется)
3. Дублирование логики check_document/check_fragment
4. Пересоздание UI при переключении видов
5. Fallback 1920×1080 для мультимониторных конфигураций
6. Lambda-замыкания в памяти при создании плиток

---

*Этот файл — полный контекст для продолжения работы. Любой следующий запрос к AI должен начинаться с прочтения этого файла.*
