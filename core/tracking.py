"""
Управление снимками документов для режима сравнения.

Снимки хранятся в userdata/snapshots/ как .docx-файлы (вне кода — переживают
обновление приложения). Индекс — index.json в той же директории.

В индексе хранится относительное имя файла (filename); абсолютный путь
резолвится через SNAPSHOTS_DIR при чтении. Это делает снимки независимыми от
расположения папки приложения. Поддерживается старая схема с ключом "path".
"""

import json
import os
import shutil
from datetime import datetime

from core import paths

SNAPSHOTS_DIR = paths.SNAPSHOTS_DIR
INDEX_FILE = paths.INDEX_FILE


class TrackingManager:
    def __init__(self):
        os.makedirs(SNAPSHOTS_DIR, exist_ok=True)
        self._index = self._load_index()

    # ─── Индекс ─────────────────────────────────────────────────────────

    def _load_index(self) -> dict:
        if os.path.isfile(INDEX_FILE):
            try:
                with open(INDEX_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"snapshots": []}

    def _save_index(self):
        try:
            with open(INDEX_FILE, "w", encoding="utf-8") as f:
                json.dump(self._index, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ─── Вспомогательные ─────────────────────────────────────────────────

    @staticmethod
    def _base_name(doc_name: str) -> str:
        """Имя документа без расширения."""
        name, _ = os.path.splitext(doc_name)
        return name

    @staticmethod
    def _resolve_path(entry: dict) -> str:
        """Абсолютный путь к файлу снимка по записи индекса.

        Новая схема — ключ "filename"; старая — "path" (резолвим по basename
        в SNAPSHOTS_DIR, если исходный путь больше не существует).
        """
        filename = entry.get("filename")
        if filename:
            return os.path.join(SNAPSHOTS_DIR, filename)
        old = entry.get("path", "")
        if old and os.path.isfile(old):
            return old
        return os.path.join(SNAPSHOTS_DIR, os.path.basename(old))

    def _public(self, entry: dict) -> dict:
        """Запись с заполненным абсолютным path для UI-диалогов."""
        out = dict(entry)
        out["path"] = self._resolve_path(entry)
        return out

    def _make_snapshot_path(self, doc_name: str, create_new: bool) -> str:
        """Вернуть путь для нового снимка.

        create_new=False → путь основного снимка (может перезаписать).
        create_new=True  → добавляем числовой суффикс _1, _2, …
        """
        base = self._base_name(doc_name)
        primary = os.path.join(SNAPSHOTS_DIR, f"{base}.docx")

        if not create_new:
            return primary

        i = 1
        while True:
            candidate = os.path.join(SNAPSHOTS_DIR, f"{base}_{i}.docx")
            if not os.path.isfile(candidate):
                return candidate
            i += 1

    # ─── Публичный API ───────────────────────────────────────────────────

    def has_snapshot(self, doc_name: str) -> bool:
        """Есть ли хотя бы один снимок для данного имени документа."""
        return any(
            e["original_doc_name"] == doc_name
            for e in self._index["snapshots"]
        )

    def list_snapshots(self, doc_name: str) -> list[dict]:
        """Все снимки для конкретного имени документа.

        Returns: list of {original_doc_name, path, display_name, created_at}
        """
        return [
            self._public(e) for e in self._index["snapshots"]
            if e["original_doc_name"] == doc_name
        ]

    def list_all_snapshots(self) -> list[dict]:
        """Все сохранённые снимки (для диалога управления)."""
        return [self._public(e) for e in self._index["snapshots"]]

    def save_snapshot(self, doc_com, doc_name: str, create_new: bool = False) -> str:
        """Сохранить снимок документа Word.

        Если документ сохранён без изменений — копируем файл напрямую.
        Если есть несохранённые изменения — используем COM SaveAs2 с возвратом
        к оригинальному пути.

        Args:
            doc_com:     COM-объект Word.Document.
            doc_name:    Имя документа (doc_com.Name).
            create_new:  True → создать новую копию с суффиксом; False → заменить основную.

        Returns: путь к сохранённому снимку.
        Raises: RuntimeError при невозможности сохранить.
        """
        dest_path = self._make_snapshot_path(doc_name, create_new=create_new)
        has_path = bool(doc_com.Path)
        is_saved = bool(doc_com.Saved)

        if not has_path:
            raise RuntimeError(
                "Документ ещё не сохранён. Сохраните документ перед созданием снимка."
            )

        if is_saved:
            shutil.copy2(doc_com.FullName, dest_path)
        else:
            original_path = doc_com.FullName
            original_fmt = doc_com.SaveFormat
            try:
                doc_com.SaveAs2(FileName=dest_path, FileFormat=16)
                doc_com.SaveAs2(FileName=original_path, FileFormat=original_fmt)
            except Exception as exc:
                raise RuntimeError(f"Не удалось сохранить снимок: {exc}") from exc

        filename = os.path.basename(dest_path)
        display_name = os.path.splitext(filename)[0]
        entry = {
            "original_doc_name": doc_name,
            "filename": filename,
            "display_name": display_name,
            "created_at": datetime.now().isoformat(timespec="seconds"),
        }

        if not create_new:
            # Заменить старую основную запись (по тому же файлу)
            self._index["snapshots"] = [
                e for e in self._index["snapshots"]
                if not (
                    e["original_doc_name"] == doc_name
                    and self._resolve_path(e) == dest_path
                )
            ]

        self._index["snapshots"].append(entry)
        self._save_index()
        return dest_path

    def delete_snapshot(self, path: str):
        """Удалить снимок по пути к файлу."""
        try:
            if os.path.isfile(path):
                os.remove(path)
        except OSError:
            pass
        self._index["snapshots"] = [
            e for e in self._index["snapshots"]
            if self._resolve_path(e) != path
        ]
        self._save_index()
