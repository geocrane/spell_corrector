"""
Точка входа приложения Corrector.

Минимальный файл: логирование, обработчики исключений, создание Engine + MainWindow.
"""

import logging
import os
import sys
import threading as _threading

# ─── Логирование ────────────────────────────────────────────────────────

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_LOG_FILE = os.path.join(_THIS_DIR, "spell_debug.log")
logger = logging.getLogger("main_app")
logger.setLevel(logging.DEBUG)
_fh = logging.FileHandler(_LOG_FILE, encoding="utf-8", mode="a")
_fh.setFormatter(
    logging.Formatter(
        "%(asctime)s [%(levelname)s] [main] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
)
logger.addHandler(_fh)


# ─── Глобальные обработчики исключений ─────────────────────────────────

def _global_except_hook(exc_type, exc_value, exc_tb):
    logger.critical("Unhandled exception", exc_info=(exc_type, exc_value, exc_tb))
    sys.__excepthook__(exc_type, exc_value, exc_tb)


def _thread_except_hook(args):
    logger.critical(
        "Unhandled thread exception [%s]",
        args.thread.name if args.thread else "?",
        exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
    )


sys.excepthook = _global_except_hook
_threading.excepthook = _thread_except_hook


# ─── Точка входа ────────────────────────────────────────────────────────

from core.engine import Engine
from ui.main_window import MainWindow


def main():
    """Создать Engine и MainWindow, запустить главный цикл приложения."""
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
