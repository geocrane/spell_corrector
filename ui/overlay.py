"""
Универсальный overlay-индикатор долгих операций.

Накладывается поверх scrollable-области (canvas), показывает сообщение,
анимированный спиннер и опциональную кнопку «Остановить». Используется для
поиска документов, извлечения предложений, унификации форматирования,
перерисовки интерфейса и загрузки модели.
"""

import logging
import tkinter as tk

from ui.constants import (
    SPINNER_FRAMES,
    SPINNER_DELAY,
    RIBBON_TOGGLE_MIXED_BG,
    RIBBON_TOGGLE_ON_BORDER,
    RIBBON_NORMAL_FG,
    RIBBON_FONT_FAMILY,
)

logger = logging.getLogger("ui.overlay")


class BusyOverlay:
    """Полупрозрачная заглушка с сообщением, спиннером и кнопкой отмены."""

    def __init__(self, parent: tk.Misc):
        self._parent = parent
        self._frame: tk.Frame | None = None
        self._inner: tk.Frame | None = None
        self._spinner_label: tk.Label | None = None
        self._message_label: tk.Label | None = None
        self._cancel_button: tk.Button | None = None
        self._spinner_index = 0
        self._spinner_job = None
        self._on_cancel = None

    def _build(self) -> None:
        # Используем чисто белый фон для всего фрейма, чтобы на 100% скрыть
        # то, что происходит под ним (отрисовка плиток).
        self._frame = tk.Frame(self._parent, bg="#ffffff", bd=0)
        self._inner = tk.Frame(self._frame, bg="#ffffff")
        self._inner.place(relx=0.5, rely=0.5, anchor="center")

        self._spinner_label = tk.Label(
            self._inner,
            text=SPINNER_FRAMES[0],
            font=(RIBBON_FONT_FAMILY, 28),
            bg="#ffffff",
            fg=RIBBON_TOGGLE_ON_BORDER,
        )
        self._spinner_label.pack(pady=(0, 10))

        self._message_label = tk.Label(
            self._inner,
            text="",
            font=(RIBBON_FONT_FAMILY, 11),
            bg="#ffffff",
            fg=RIBBON_NORMAL_FG,
            wraplength=380,
            justify="center",
        )
        self._message_label.pack()

        self._cancel_button = tk.Button(
            self._inner,
            text="Остановить",
            font=(RIBBON_FONT_FAMILY, 10),
            bg="#ffffff",
            fg=RIBBON_NORMAL_FG,
            activebackground="#f3f3f3",
            activeforeground=RIBBON_NORMAL_FG,
            relief="solid",
            bd=1,
            padx=12,
            pady=4,
            cursor="hand2",
            command=self._handle_cancel,
        )

    def show(
        self,
        message: str,
        *,
        cancelable: bool = True,
        on_cancel=None,
    ) -> None:
        """Показать overlay поверх parent.

        Args:
            message: Сообщение в центре.
            cancelable: Если True И задан on_cancel — снизу кнопка «Остановить».
            on_cancel: Колбэк по клику «Остановить». Вызывается один раз.
        """
        if self._frame is None:
            self._build()

        self._on_cancel = on_cancel
        assert self._message_label is not None
        self._message_label.config(text=message)

        assert self._cancel_button is not None
        if cancelable and on_cancel is not None:
            self._cancel_button.config(state="normal")
            self._cancel_button.pack(pady=(12, 0))
        else:
            self._cancel_button.pack_forget()

        assert self._frame is not None
        self._frame.place(x=0, y=0, relwidth=1, relheight=1)
        self._frame.lift()

        self._spinner_index = 0
        self._cancel_spinner_job()
        self._spinner_job = self._parent.after(SPINNER_DELAY, self._animate)

    def update_message(self, message: str) -> None:
        """Обновить текст сообщения (например, прогресс)."""
        if self._message_label is not None and self._message_label.winfo_exists():
            self._message_label.config(text=message)

    def is_visible(self) -> bool:
        return self._frame is not None and self._frame.winfo_ismapped()

    def hide(self) -> None:
        """Скрыть overlay."""
        self._cancel_spinner_job()
        self._on_cancel = None
        if self._frame is not None:
            self._frame.place_forget()

    def _animate(self) -> None:
        if (
            self._frame is not None
            and self._frame.winfo_ismapped()
            and self._spinner_label is not None
        ):
            self._spinner_index = (self._spinner_index + 1) % len(SPINNER_FRAMES)
            self._spinner_label.config(text=SPINNER_FRAMES[self._spinner_index])
            self._spinner_job = self._parent.after(SPINNER_DELAY, self._animate)
        else:
            self._spinner_job = None

    def _handle_cancel(self) -> None:
        if self._cancel_button is not None:
            self._cancel_button.config(state="disabled")
        self.update_message("Останавливаем…")
        cb = self._on_cancel
        self._on_cancel = None
        if cb is not None:
            try:
                cb()
            except Exception:
                logger.exception("on_cancel callback failed")

    def _cancel_spinner_job(self) -> None:
        if self._spinner_job is not None:
            try:
                self._parent.after_cancel(self._spinner_job)
            except Exception:
                pass
            self._spinner_job = None
