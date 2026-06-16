"""
Обновление приложения из локального zip-архива.

Сценарий (всё через интерфейс):
  1. validate_archive  — проверить, что это корректный архив обновления, прочитать версию.
  2. create_backup     — заархивировать текущий код (без model/venv/userdata/backups).
  3. apply_update      — распаковать архив и заменить файлы кода.
  4. rollback          — при необходимости восстановить прежнюю версию из backups/.

Пользовательские данные (userdata/) и тяжёлые ресурсы (model/, venv/) НИКОГДА не
затрагиваются: они в списке исключений _EXCLUDE_TOP.

Все длительные операции принимают progress_cb(message: str) и сообщают реальные
этапы с счётчиком файлов, чтобы UI мог показать прогресс.
"""

import json
import logging
import os
import shutil
import tempfile
import zipfile
from datetime import datetime

from core import paths

logger = logging.getLogger("core.updater")

# Папки верхнего уровня, которые НЕ входят ни в бэкап, ни в применение обновления.
_EXCLUDE_TOP = {
    "model", "venv", ".venv", "userdata", "backups",
    ".git", "__pycache__", ".pytest_cache", ".idea", ".vscode", "dist",
}

# Файлы, без которых архив считается некорректным.
_REQUIRED = ("version.json", "main.py", "core/engine.py", "ui/main_window.py")

# Файлы, которые ОБЯЗАНЫ присутствовать в APP_ROOT, чтобы считать его настоящим
# корнем установленного приложения. Если их нет — APP_ROOT определён неверно
# (запуск из чужой копии / PYTHONPATH / вложенной структуры), и трогать диск
# нельзя: иначе бэкап выйдет пустым, а файлы применятся «не туда».
_REQUIRED_LOCAL = ("version.json", "main.py", "core/engine.py", "ui/main_window.py")


def _noop(_msg: str):
    pass


def verify_app_root() -> list[str]:
    """Проверить, что paths.APP_ROOT — настоящий корень установки.

    Returns: список отсутствующих обязательных файлов (пустой → всё в порядке).
    """
    missing = []
    for rel in _REQUIRED_LOCAL:
        if not os.path.isfile(os.path.join(paths.APP_ROOT, *rel.split("/"))):
            missing.append(rel)
    return missing


def _assert_app_root(stage: str):
    """Прервать операцию, если APP_ROOT не похож на корень установки."""
    missing = verify_app_root()
    logger.info("%s: APP_ROOT=%s (missing=%s)", stage, paths.APP_ROOT, missing)
    if missing:
        raise RuntimeError(
            "Обновление прервано: не удалось определить папку приложения "
            f"(APP_ROOT={paths.APP_ROOT}).\n"
            "Отсутствуют ожидаемые файлы: " + ", ".join(missing) + ".\n"
            "Запустите приложение напрямую из его папки (python main.py) и "
            "повторите. Диск не изменён."
        )


# ─── Работа со структурой архива ─────────────────────────────────────────

def _archive_prefix(names: list[str]) -> str:
    """Определить общий верхний каталог архива ('' если файлы лежат в корне).

    Поддерживает архивы вида `code/...` (одна верхняя папка) и архивы, где
    файлы лежат прямо в корне.
    """
    tops = set()
    for n in names:
        n = n.replace("\\", "/")
        if not n or n.startswith("__MACOSX/"):
            continue
        tops.add(n.split("/", 1)[0])
    if len(tops) == 1:
        only = next(iter(tops))
        # Это одна верхняя папка, только если внутри неё есть version.json
        if f"{only}/version.json" in [x.replace("\\", "/") for x in names]:
            return only + "/"
    return ""


# ─── Валидация ───────────────────────────────────────────────────────────

def validate_archive(zip_path: str) -> tuple[bool, str | None, list[str]]:
    """Проверить архив обновления.

    Returns:
        (ok, version, errors). version — строка из version.json архива или None.
    """
    errors: list[str] = []
    if not os.path.isfile(zip_path):
        return False, None, ["Файл не найден."]
    if not zipfile.is_zipfile(zip_path):
        return False, None, ["Файл не является ZIP-архивом."]

    try:
        with zipfile.ZipFile(zip_path) as zf:
            names = zf.namelist()
            prefix = _archive_prefix(names)
            name_set = {n.replace("\\", "/") for n in names}

            for req in _REQUIRED:
                if prefix + req not in name_set:
                    errors.append(f"В архиве отсутствует обязательный файл: {req}")

            version = None
            vkey = prefix + "version.json"
            if vkey in name_set:
                try:
                    version = str(json.loads(zf.read(vkey)).get("version"))
                except (json.JSONDecodeError, ValueError):
                    errors.append("version.json в архиве повреждён.")
    except (zipfile.BadZipFile, OSError) as e:
        return False, None, [f"Не удалось прочитать архив: {e}"]

    if version in (None, "None"):
        errors.append("В архиве не указана версия.")
        version = None
    return (not errors), version, errors


# ─── Перечисление файлов кода ────────────────────────────────────────────

def _is_excluded(rel_path: str) -> bool:
    """Входит ли путь (относительно APP_ROOT) в исключаемую папку верхнего уровня."""
    top = rel_path.replace("\\", "/").split("/", 1)[0]
    return top in _EXCLUDE_TOP


def _iter_code_files():
    """Перечислить все файлы кода в APP_ROOT (с учётом исключений)."""
    for root, dirs, files in os.walk(paths.APP_ROOT):
        rel_root = os.path.relpath(root, paths.APP_ROOT)
        if rel_root == ".":
            dirs[:] = [d for d in dirs if d not in _EXCLUDE_TOP]
        for fname in files:
            if fname.endswith((".pyc", ".pyo")):
                continue
            abs_path = os.path.join(root, fname)
            rel = os.path.relpath(abs_path, paths.APP_ROOT)
            if not _is_excluded(rel):
                yield abs_path, rel


# ─── Резервная копия ─────────────────────────────────────────────────────

def create_backup(progress_cb=None) -> str:
    """Заархивировать текущий код в backups/backup_v{ver}_{timestamp}.zip.

    Исключает model/, venv/, userdata/, backups/ и служебные папки.
    Returns: путь к созданному архиву бэкапа.
    """
    progress_cb = progress_cb or _noop
    _assert_app_root("create_backup")
    os.makedirs(paths.BACKUPS_DIR, exist_ok=True)

    progress_cb("Подготовка резервной копии…")
    items = list(_iter_code_files())
    total = len(items)

    # Защита: пустой бэкап означает, что обходить было нечего (APP_ROOT не там).
    # Не создаём бесполезный архив и не пускаем обновление дальше.
    if total == 0:
        raise RuntimeError(
            "Обновление прервано: резервная копия пуста — в папке приложения "
            f"(APP_ROOT={paths.APP_ROOT}) не найдено файлов кода. Диск не изменён."
        )

    ver = paths.get_version()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(paths.BACKUPS_DIR, f"backup_v{ver}_{ts}.zip")

    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, (abs_path, rel) in enumerate(items, 1):
            zf.write(abs_path, rel)
            if i % 25 == 0 or i == total:
                progress_cb(f"Резервная копия: {i}/{total} файлов…")
    logger.info("Создан бэкап: %s (%d файлов)", dest, total)
    return dest


# ─── Применение обновления ───────────────────────────────────────────────

def apply_update(zip_path: str, progress_cb=None):
    """Распаковать архив и заменить файлы кода в APP_ROOT.

    userdata/, backups/, model/, venv/ не трогаются (их нет в архиве, а копирование
    идёт только по содержимому архива).
    """
    progress_cb = progress_cb or _noop
    _assert_app_root("apply_update")
    root_real = os.path.realpath(paths.APP_ROOT)

    with zipfile.ZipFile(zip_path) as zf:
        names = [n for n in zf.namelist() if not n.endswith("/")]
        prefix = _archive_prefix(zf.namelist())

        progress_cb("Распаковка архива…")
        staging = tempfile.mkdtemp(prefix="spell_update_")
        try:
            zf.extractall(staging)

            total = len(names)
            progress_cb(f"Применение файлов: 0/{total}…")

            copied = 0
            for n in names:
                rel = n.replace("\\", "/")
                if prefix and rel.startswith(prefix):
                    rel = rel[len(prefix):]
                if not rel or rel.startswith("__MACOSX") or _is_excluded(rel):
                    continue
                src = os.path.join(staging, n)
                dst = os.path.join(paths.APP_ROOT, rel)
                # Защита от zip-slip: целевой путь обязан лежать внутри APP_ROOT.
                dst_real = os.path.realpath(dst)
                if dst_real != root_real and not dst_real.startswith(root_real + os.sep):
                    logger.warning("Пропущен файл вне APP_ROOT: %s", rel)
                    continue
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                shutil.copy2(src, dst)
                copied += 1
                if copied % 25 == 0 or copied == total:
                    progress_cb(f"Применение файлов: {copied}/{total}…")
            logger.info("Обновление применено: %d файлов из %s", copied, zip_path)
        finally:
            shutil.rmtree(staging, ignore_errors=True)
    progress_cb("Завершение…")


# ─── Откат ───────────────────────────────────────────────────────────────

def list_backups() -> list[dict]:
    """Список бэкапов (новые сверху): [{path, name, created_at}]."""
    if not os.path.isdir(paths.BACKUPS_DIR):
        return []
    out = []
    for f in os.listdir(paths.BACKUPS_DIR):
        if f.endswith(".zip"):
            p = os.path.join(paths.BACKUPS_DIR, f)
            out.append({
                "path": p,
                "name": f,
                "created_at": datetime.fromtimestamp(
                    os.path.getmtime(p)
                ).isoformat(timespec="seconds"),
            })
    return sorted(out, key=lambda e: e["created_at"], reverse=True)


def rollback(backup_zip: str, progress_cb=None):
    """Восстановить код из бэкапа (распаковать поверх APP_ROOT)."""
    progress_cb = progress_cb or _noop
    if not zipfile.is_zipfile(backup_zip):
        raise ValueError("Файл бэкапа повреждён или не является ZIP-архивом.")
    with zipfile.ZipFile(backup_zip) as zf:
        names = [n for n in zf.namelist() if not n.endswith("/")]
        total = len(names)
        progress_cb(f"Восстановление: 0/{total}…")
        for i, n in enumerate(names, 1):
            rel = n.replace("\\", "/")
            if _is_excluded(rel):
                continue
            dst = os.path.join(paths.APP_ROOT, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            with zf.open(n) as src, open(dst, "wb") as out:
                shutil.copyfileobj(src, out)
            if i % 25 == 0 or i == total:
                progress_cb(f"Восстановление: {i}/{total}…")
    logger.info("Откат выполнен из %s", backup_zip)
