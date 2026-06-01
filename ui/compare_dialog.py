"""
Модальное окно выбора документа и снимка для сравнения.

Левая панель:  список открытых Word-документов → выбор «текущего» (в нём идёт навигация).
Правая панель: ВСЕ сохранённые снимки, сгруппированные по имени исходного документа.
               Доступны всегда, независимо от того, открыт ли исходный файл.

Callback on_compare(doc_dict, snapshot_path) вызывается при нажатии «Сравнить».
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


class CompareDialog(tk.Toplevel):
    """Диалог выбора документа и снимка для сравнения."""

    def __init__(self, parent, open_docs: list[dict],
                 tracking_manager: TrackingManager,
                 on_compare):
        """
        open_docs:        список открытых Word-документов (doc dict с полем "name").
        tracking_manager: экземпляр TrackingManager.
        on_compare:       callable(doc_dict, snapshot_path) — вызывается при старте сравнения.
        """
        super().__init__(parent)
        self.title("Сравнение документов")
        self.resizable(False, False)
        self.grab_set()
        self.configure(bg=RIBBON_BG)

        self._open_docs = open_docs
        self._tracking = tracking_manager
        self._on_compare = on_compare
        self._selected_doc = None
        self._selected_snapshot_path = None
        # Sentinel: не совпадает ни с одним path И не равен tristatevalue("")
        # → все радиокнопки пустые, пока пользователь не выберет.
        self._snap_var = tk.StringVar(value="\x00none")

        self._build_ui()
        self._refresh_snapshots()  # показать все снимки сразу
        self._center_on_parent(parent)

    # ─── UI ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        outer = tk.Frame(self, bg=RIBBON_BG, padx=14, pady=12)
        outer.pack(fill=tk.BOTH, expand=True)

        # ── Левая панель: список открытых документов ──────────────────
        left = tk.Frame(outer, bg=RIBBON_BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        tk.Label(
            left, text="Текущий документ:",
            font=_TITLE_FONT, bg=RIBBON_BG, fg=RIBBON_NORMAL_FG,
        ).pack(anchor="w", pady=(0, 4))

        tk.Label(
            left,
            text="(навигация по изменениям\nпроисходит в нём)",
            font=_SMALL_FONT, bg=RIBBON_BG, fg="#888",
            justify="left",
        ).pack(anchor="w", pady=(0, 6))

        doc_frame = tk.Frame(left, bg=RIBBON_BG, bd=1, relief="solid",
                             highlightbackground=RIBBON_HOVER_BORDER)
        doc_frame.pack(fill=tk.BOTH, expand=True)

        self._doc_listbox = tk.Listbox(
            doc_frame, selectmode=tk.SINGLE,
            font=_NORMAL_FONT, activestyle="none",
            selectbackground="#cfe4f9", selectforeground=RIBBON_NORMAL_FG,
            bg=RIBBON_BG, fg=RIBBON_NORMAL_FG,
            relief="flat", bd=0, width=26, height=10,
            exportselection=False,
        )
        doc_sb = ttk.Scrollbar(doc_frame, orient="vertical",
                               command=self._doc_listbox.yview)
        self._doc_listbox.configure(yscrollcommand=doc_sb.set)
        doc_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._doc_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        if self._open_docs:
            for doc in self._open_docs:
                self._doc_listbox.insert(tk.END, f"  {doc['name']}")
        else:
            self._doc_listbox.insert(tk.END, "  (нет открытых документов)")
            self._doc_listbox.config(state="disabled")

        self._doc_listbox.bind("<<ListboxSelect>>", self._on_doc_select)

        # ── Правая панель: ВСЕ снимки ─────────────────────────────────
        right = tk.Frame(outer, bg=RIBBON_BG)
        right.grid(row=0, column=1, sticky="nsew")

        tk.Label(
            right, text="Снимок для сравнения:",
            font=_TITLE_FONT, bg=RIBBON_BG, fg=RIBBON_NORMAL_FG,
        ).pack(anchor="w", pady=(0, 4))

        tk.Label(
            right,
            text="(доступны все сохранённые снимки)",
            font=_SMALL_FONT, bg=RIBBON_BG, fg="#888",
            justify="left",
        ).pack(anchor="w", pady=(0, 6))

        snap_outer = tk.Frame(right, bg=RIBBON_BG, bd=1, relief="solid",
                              highlightbackground=RIBBON_HOVER_BORDER)
        snap_outer.pack(fill=tk.BOTH, expand=True)

        self._snap_canvas = tk.Canvas(snap_outer, bg=RIBBON_BG,
                                      highlightthickness=0, width=240)
        snap_sb = ttk.Scrollbar(snap_outer, orient="vertical",
                                command=self._snap_canvas.yview)
        self._snap_canvas.configure(yscrollcommand=snap_sb.set)
        snap_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._snap_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._snap_frame = tk.Frame(self._snap_canvas, bg=RIBBON_BG)
        self._snap_window = self._snap_canvas.create_window(
            (0, 0), window=self._snap_frame, anchor="nw",
        )
        self._snap_frame.bind("<Configure>", self._on_snap_frame_configure)
        self._snap_canvas.bind("<Configure>", self._on_snap_canvas_configure)

        outer.columnconfigure(0, weight=1)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(0, weight=1)

        # ── Кнопки ────────────────────────────────────────────────────
        btn_row = tk.Frame(outer, bg=RIBBON_BG)
        btn_row.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))

        self._compare_btn = tk.Button(
            btn_row, text="Сравнить",
            command=self._on_confirm,
            state="disabled",
            font=_NORMAL_FONT,
            bg="#007bff", fg="white",
            activebackground="#0056b3", activeforeground="white",
            relief="flat", padx=12, pady=4, cursor="hand2",
        )
        self._compare_btn.pack(side=tk.RIGHT, padx=(6, 0))

        tk.Button(
            btn_row, text="Отмена",
            command=self.destroy,
            font=_NORMAL_FONT, relief="flat", padx=12, pady=4,
            bg=RIBBON_HOVER_BG, fg=RIBBON_NORMAL_FG,
            activebackground="#d0d0d0", cursor="hand2",
        ).pack(side=tk.RIGHT)

        self.minsize(500, 300)

    # ─── Документы ──────────────────────────────────────────────────────

    def _on_doc_select(self, _event=None):
        sel = self._doc_listbox.curselection()
        if not sel:
            return
        self._selected_doc = self._open_docs[sel[0]]
        self._update_compare_btn()

    # ─── Снимки (ВСЕ, сгруппированные по документу) ─────────────────────

    def _refresh_snapshots(self):
        for w in self._snap_frame.winfo_children():
            w.destroy()
        self._snap_var.set("\x00none")
        self._selected_snapshot_path = None
        self._update_compare_btn()

        all_snaps = self._tracking.list_all_snapshots()
        if not all_snaps:
            tk.Label(
                self._snap_frame,
                text="Нет сохранённых снимков.\nНажмите «Отслеживать» в режиме\nсравнения, чтобы создать снимок.",
                font=_SMALL_FONT, fg="#999", bg=RIBBON_BG, justify="center",
            ).pack(pady=20)
            return

        # Группировать по original_doc_name
        groups: dict[str, list[dict]] = {}
        for snap in all_snaps:
            groups.setdefault(snap["original_doc_name"], []).append(snap)

        for doc_name, snaps in groups.items():
            self._add_group(doc_name, snaps)

    def _add_group(self, doc_name: str, snaps: list[dict]):
        group_frame = tk.Frame(self._snap_frame, bg=RIBBON_BG)
        group_frame.pack(fill=tk.X, padx=4, pady=(6, 2))

        tk.Label(
            group_frame, text=f"📄 {doc_name}",
            font=_GROUP_FONT, bg=RIBBON_BG, fg=RIBBON_NORMAL_FG,
            anchor="w",
        ).pack(fill=tk.X)

        ttk.Separator(group_frame, orient="horizontal").pack(fill=tk.X, pady=(1, 3))

        for snap in snaps:
            self._add_snapshot_row(snap)

    def _add_snapshot_row(self, snap: dict):
        row = tk.Frame(self._snap_frame, bg=RIBBON_BG, pady=1)
        row.pack(fill=tk.X, padx=12)

        # Все RadioButton делят ОДНУ self._snap_var → взаимоисключающий выбор
        rb = tk.Radiobutton(
            row,
            text=f"{snap['display_name']}  {_fmt_date(snap['created_at'])}",
            variable=self._snap_var,
            value=snap["path"],
            justify="left", anchor="w",
            font=_SMALL_FONT, bg=RIBBON_BG,
            activebackground=RIBBON_HOVER_BG,
            command=lambda p=snap["path"]: self._on_snap_select(p),
        )
        rb.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Button(
            row, text="×",
            font=("Segoe UI", 9, "bold"),
            fg="#dc3545", bg=RIBBON_BG,
            activeforeground="white", activebackground="#dc3545",
            relief="flat", bd=0, cursor="hand2", padx=4,
            command=lambda p=snap["path"]: self._delete_snapshot(p),
        ).pack(side=tk.RIGHT)

    def _on_snap_select(self, path: str):
        self._selected_snapshot_path = path
        self._update_compare_btn()

    def _delete_snapshot(self, path: str):
        if not messagebox.askyesno(
            "Удалить снимок",
            "Удалить этот снимок? Действие нельзя отменить.",
            parent=self,
        ):
            return
        self._tracking.delete_snapshot(path)
        if self._selected_snapshot_path == path:
            self._selected_snapshot_path = None
        self._refresh_snapshots()

    # ─── Скролл списка снимков ──────────────────────────────────────────

    def _on_snap_frame_configure(self, _event=None):
        self._snap_canvas.configure(scrollregion=self._snap_canvas.bbox("all"))

    def _on_snap_canvas_configure(self, event):
        self._snap_canvas.itemconfig(self._snap_window, width=event.width)

    # ─── Кнопка «Сравнить» ───────────────────────────────────────────────

    def _update_compare_btn(self):
        ready = (
            self._selected_doc is not None
            and self._selected_snapshot_path is not None
        )
        self._compare_btn.configure(state="normal" if ready else "disabled")

    def _on_confirm(self):
        if self._selected_doc is None or self._selected_snapshot_path is None:
            return
        doc = self._selected_doc
        path = self._selected_snapshot_path
        self.destroy()
        self._on_compare(doc, path)

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
