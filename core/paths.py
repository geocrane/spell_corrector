"""
Единый источник путей приложения и пользовательских данных.

Принцип: КОД и ДАННЫЕ разделены, чтобы обновление (замена файлов кода новым
архивом) не затрагивало настройки и снимки пользователя.

- APP_ROOT — корень приложения (папка, которую пользователь распаковывает).
- userdata/ — все изменяемые данные пользователя (config, ui_style, снимки, логи).
  Лежит ВНУТРИ папки приложения, но РЯДОМ с кодом и НИКОГДА не трогается обновлением.
- backups/ — резервные копии прежних версий кода (для отката).
- version.json — текущая версия (часть кода, входит в релизный архив).

Здесь же — чтение/сравнение версий и одноразовая миграция данных из старых мест
(config.json/ui_style.json в корне кода, снимки в %TEMP%).
"""

import json
import logging
import os
import re
import shutil
import tempfile

logger = logging.getLogger("core.paths")

# ─── Базовые пути ────────────────────────────────────────────────────────

# core/paths.py → APP_ROOT на уровень выше пакета core/
APP_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

USERDATA_DIR = os.path.join(APP_ROOT, "userdata")
CONFIG_PATH = os.path.join(USERDATA_DIR, "config.json")
STYLE_PATH = os.path.join(USERDATA_DIR, "ui_style.json")
SNAPSHOTS_DIR = os.path.join(USERDATA_DIR, "snapshots")
INDEX_FILE = os.path.join(SNAPSHOTS_DIR, "index.json")
LOG_DIR = os.path.join(USERDATA_DIR, "logs")

BACKUPS_DIR = os.path.join(APP_ROOT, "backups")
VERSION_FILE = os.path.join(APP_ROOT, "version.json")

# Старые расположения (для миграции с прежних версий)
_OLD_CONFIG = os.path.join(APP_ROOT, "config.json")
_OLD_STYLE = os.path.join(APP_ROOT, "ui_style.json")
_OLD_SNAPSHOTS_DIR = os.path.join(tempfile.gettempdir(), "spell_corrector_tracking")


def ensure_dirs():
    """Создать все каталоги пользовательских данных (идемпотентно).

    Вызывается при импорте модуля, чтобы логгеры и конфиги могли писать
    в userdata ещё до явного вызова ensure_userdata().
    """
    for d in (USERDATA_DIR, SNAPSHOTS_DIR, LOG_DIR, BACKUPS_DIR):
        try:
            os.makedirs(d, exist_ok=True)
        except OSError as e:
            logger.error("Не удалось создать каталог %s: %s", d, e)


# ─── Версии ──────────────────────────────────────────────────────────────

def get_version() -> str:
    """Текущая версия приложения из version.json (fallback '0.0')."""
    try:
        with open(VERSION_FILE, "r", encoding="utf-8") as f:
            return str(json.load(f).get("version", "0.0"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return "0.0"


def parse_version(s: str) -> tuple:
    """Разобрать строку версии в кортеж целых ('1.5.2' → (1, 5, 2))."""
    parts = []
    for chunk in str(s).split("."):
        m = re.match(r"\d+", chunk.strip())
        parts.append(int(m.group()) if m else 0)
    return tuple(parts) or (0,)


def is_newer(candidate: str, current: str) -> bool:
    """True, если candidate СТРОГО новее current."""
    a, b = parse_version(candidate), parse_version(current)
    n = max(len(a), len(b))
    a += (0,) * (n - len(a))
    b += (0,) * (n - len(b))
    return a > b


# ─── Миграция данных при первом запуске после разделения ────────────────

def ensure_userdata():
    """Создать userdata и перенести данные из старых расположений (идемпотентно)."""
    ensure_dirs()
    _migrate_file(_OLD_CONFIG, CONFIG_PATH, "config.json")
    _migrate_file(_OLD_STYLE, STYLE_PATH, "ui_style.json")
    _migrate_snapshots()


def _migrate_file(old_path: str, new_path: str, label: str):
    """Перенести файл из старого места в новое, если нового ещё нет."""
    if os.path.isfile(new_path) or not os.path.isfile(old_path):
        return
    try:
        shutil.move(old_path, new_path)
        logger.info("Миграция %s → %s", label, new_path)
    except OSError as e:
        logger.error("Не удалось перенести %s: %s", label, e)


def _migrate_snapshots():
    """Перенести снимки из %TEMP%/spell_corrector_tracking в userdata/snapshots.

    Выполняется только если новое хранилище ещё пусто (нет *.docx). Индекс
    переписывается на схему с относительным именем файла (filename), чтобы
    снимки были независимы от пути.
    """
    if not os.path.isdir(_OLD_SNAPSHOTS_DIR):
        return
    has_new = any(
        f.lower().endswith(".docx") for f in os.listdir(SNAPSHOTS_DIR)
    ) if os.path.isdir(SNAPSHOTS_DIR) else False
    if has_new:
        return

    old_index = os.path.join(_OLD_SNAPSHOTS_DIR, "index.json")
    try:
        with open(old_index, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return

    new_entries = []
    for entry in data.get("snapshots", []):
        old_file = entry.get("path") or os.path.join(
            _OLD_SNAPSHOTS_DIR, f"{entry.get('display_name', '')}.docx"
        )
        filename = os.path.basename(old_file)
        if not filename:
            continue
        try:
            if os.path.isfile(old_file):
                shutil.copy2(old_file, os.path.join(SNAPSHOTS_DIR, filename))
        except OSError as e:
            logger.error("Снимок %s не перенесён: %s", filename, e)
            continue
        new_entries.append({
            "original_doc_name": entry.get("original_doc_name", ""),
            "filename": filename,
            "display_name": entry.get(
                "display_name", os.path.splitext(filename)[0]
            ),
            "created_at": entry.get("created_at", ""),
        })

    try:
        with open(INDEX_FILE, "w", encoding="utf-8") as f:
            json.dump({"snapshots": new_entries}, f, ensure_ascii=False, indent=2)
        logger.info("Перенесено снимков: %d", len(new_entries))
    except OSError as e:
        logger.error("Не удалось записать индекс снимков: %s", e)


# Создать каталоги при импорте, чтобы логгеры/конфиги сразу могли писать.
ensure_dirs()
