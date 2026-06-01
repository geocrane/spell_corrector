"""
Диалог управления сохранёнными снимками документов.

Отображает все снимки, сгруппированные по имени исходного документа,
с возможностью удаления каждого снимка.
Доступен без открытых документов — через кнопку «Снимки» в режиме отслеживания.
"""

import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

from core.tracking import TrackingManager
from ui.constants import RIBBON_BG, RIBBON_NORMAL_FG, RIBBON_HOVER_BG, RIBBON_HOVER_BORDER

_TITLE_FONT = ("Segoe UI", 10, "bold")
_GROUP_FONT = ("Segoe UI", 9, "bold")
_NORMAL_FONT = ("Segoe UI", 9)
_SMALL_FONT = ("Segoe UI", 8)
_DATE_FMT = "%d.%m.%Y %H:%M"


def _fmt_date(iso: str) -> str:
    try:
        return datetime.fromisoformat(iso).strftime(_DATE_FMT)
    except Exception:
        return iso


class SnapshotManagerDialog(tk.Toplevel):
    """Диалог просмотра и удаления всех сохранённых снимков."""

    def __init__(self, parent, tracking_manager: TrackingManager):
        super().__init__(parent)
        self.title("Управление снимками")
        self.resizable(False, True)
        self.grab_set()
        self.configure(bg=RIBBON_BG)

        self._tracking = tracking_manager
        self._build_ui()
        self._refresh()
        self._center_on_parent(parent)

    # ─── UI ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        outer = tk.Frame(self, bg=RIBBON_BG, padx=14, pady=12)
        outer.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            outer, text="Сохранённые снимки",
            font=_TITLE_FONT, bg=RIBBON_BG, fg=RIBBON_NORMAL_FG,
        ).pack(anchor="w", pady=(0, 8))

        # Прокручиваемая область
        container = tk.Frame(outer, bg=RIBBON_BG, bd=1, relief="solid",
                             highlightbackground=RIBBON_HOVER_BORDER)
        container.pack(fill=tk.BOTH, expand=True)

        self._canvas = tk.Canvas(container, bg=RIBBON_BG, highlightthickness=0,
                                 width=360, height=300)
        sb = ttk.Scrollbar(container, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._content = tk.Frame(self._canvas, bg=RIBBON_BG)
        self._content_window = self._canvas.create_window(
            (0, 0), window=self._content, anchor="nw",
        )
        self._content.bind("<Configure>", self._on_content_configure)
        self._canvas.bind("<Configure>", self._on_canvas_configure)

        # Кнопка закрытия
        btn_row = tk.Frame(outer, bg=RIBBON_BG)
        btn_row.pack(fill=tk.X, pady=(10, 0))

        tk.Button(
            btn_row, text="Закрыть",
            command=self.destroy,
            font=_NORMAL_FONT, relief="flat", padx=14, pady=4,
            bg=RIBBON_HOVER_BG, fg=RIBBON_NORMAL_FG,
            activebackground="#d0d0d0", cursor="hand2",
        ).pack(side=tk.RIGHT)

        self.minsize(400, 200)

    def _refresh(self):
        for w in self._content.winfo_children():
            w.destroy()

        all_snaps = self._tracking.list_all_snapshots()

        if not all_snaps:
            tk.Label(
                self._content,
                text="Снимков нет. Нажмите «Отслеживать» в режиме\nсравнения, чтобы сохранить снимок документа.",
                font=_SMALL_FONT, fg="#999", bg=RIBBON_BG, justify="center",
            ).pack(pady=30)
            return

        # Группировка по имени исходного документа
        groups: dict[str, list[dict]] = {}
        for snap in all_snaps:
            groups.setdefault(snap["original_doc_name"], []).append(snap)

        for doc_name, snaps in groups.items():
            self._add_group(doc_name, snaps)

    def _add_group(self, doc_name: str, snaps: list[dict]):
        # Заголовок группы
        group_frame = tk.Frame(self._content, bg=RIBBON_BG)
        group_frame.pack(fill=tk.X, padx=6, pady=(8, 2))

        tk.Label(
            group_frame, text=f"📄 {doc_name}",
            font=_GROUP_FONT, bg=RIBBON_BG, fg=RIBBON_NORMAL_FG,
            anchor="w",
        ).pack(fill=tk.X)

        ttk.Separator(group_frame, orient="horizontal").pack(fill=tk.X, pady=2)

        # Строки снимков
        for snap in snaps:
            self._add_snap_row(snap)

    def _add_snap_row(self, snap: dict):
        row = tk.Frame(self._content, bg=RIBBON_BG)
        row.pack(fill=tk.X, padx=18, pady=1)

        tk.Label(
            row,
            text=f"{snap['display_name']}",
            font=_NORMAL_FONT, bg=RIBBON_BG, fg=RIBBON_NORMAL_FG,
            anchor="w",
        ).pack(side=tk.LEFT)

        tk.Label(
            row,
            text=f"  {_fmt_date(snap['created_at'])}",
            font=_SMALL_FONT, bg=RIBBON_BG, fg="#888",
            anchor="w",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Button(
            row, text="×",
            font=("Segoe UI", 9, "bold"),
            fg="#dc3545", bg=RIBBON_BG,
            activeforeground="white", activebackground="#dc3545",
            relief="flat", bd=0, cursor="hand2", padx=6,
            command=lambda p=snap["path"]: self._delete_snapshot(p),
        ).pack(side=tk.RIGHT)

    def _delete_snapshot(self, path: str):
        if not messagebox.askyesno(
            "Удалить снимок",
            "Удалить этот снимок? Действие нельзя отменить.",
            parent=self,
        ):
            return
        self._tracking.delete_snapshot(path)
        self._refresh()

    # ─── Скролл ─────────────────────────────────────────────────────────

    def _on_content_configure(self, _event=None):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self._canvas.itemconfig(self._content_window, width=event.width)

    # ─── Позиционирование ────────────────────────────────────────────────

    def _center_on_parent(self, parent):
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        dw = self.winfo_reqwidth()
        dh = self.winfo_reqheight()
        x = px + (pw - dw) // 2
        y = py + (ph - dh) // 2
        self.geometry(f"+{max(0, x)}+{max(0, y)}")
