"""
Сборка релизного zip-архива приложения на машине разработчика.

Кладёт в архив только КОД (для запуска и обновления), исключая тяжёлые ресурсы
(model/, venv/), пользовательские данные (userdata/, backups/) и служебный мусор.
Архив, собранный этим скриптом, гарантированно проходит updater.validate_archive.

Использование:
    python scripts/make_release.py 1.6
    python scripts/make_release.py 1.6 --out /путь/к/выходу.zip

Перед сборкой версия записывается в version.json (в корне проекта).
"""

import argparse
import json
import os
import sys
import zipfile

APP_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Папки верхнего уровня, которые НЕ входят в релиз.
EXCLUDE_TOP = {
    "model", "venv", ".venv", "userdata", "backups",
    ".git", "__pycache__", ".pytest_cache", ".idea", ".vscode",
    "dist", "tests",
}
# Файлы, которые не нужны в релизе.
EXCLUDE_FILE_SUFFIX = (".pyc", ".pyo", ".log", ".zip", ".rar")
EXCLUDE_FILE_NAMES = {".DS_Store"}


def _included(rel_path: str) -> bool:
    parts = rel_path.replace("\\", "/").split("/")
    if parts[0] in EXCLUDE_TOP:
        return False
    name = parts[-1]
    if name in EXCLUDE_FILE_NAMES or name.endswith(EXCLUDE_FILE_SUFFIX):
        return False
    return True


def write_version(version: str):
    with open(os.path.join(APP_ROOT, "version.json"), "w", encoding="utf-8") as f:
        json.dump({"version": version}, f, ensure_ascii=False, indent=2)
        f.write("\n")


def build(version: str, out_path: str):
    write_version(version)

    files = []
    for root, dirs, names in os.walk(APP_ROOT):
        rel_root = os.path.relpath(root, APP_ROOT)
        if rel_root == ".":
            dirs[:] = [d for d in dirs if d not in EXCLUDE_TOP]
        for n in names:
            rel = os.path.relpath(os.path.join(root, n), APP_ROOT)
            if _included(rel):
                files.append(rel)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for rel in files:
            zf.write(os.path.join(APP_ROOT, rel), rel)

    size_mb = os.path.getsize(out_path) / (1024 * 1024)
    print(f"Готово: {out_path}")
    print(f"Версия: {version} · файлов: {len(files)} · размер: {size_mb:.1f} МБ")


def main():
    parser = argparse.ArgumentParser(description="Сборка релизного архива")
    parser.add_argument("version", help="Новая версия, например 1.6")
    parser.add_argument(
        "--out",
        default=None,
        help="Путь к выходному zip (по умолчанию dist/spell_corrector_v<version>.zip)",
    )
    args = parser.parse_args()

    out = args.out or os.path.join(
        APP_ROOT, "dist", f"spell_corrector_v{args.version}.zip"
    )
    build(args.version, out)


if __name__ == "__main__":
    sys.exit(main())
