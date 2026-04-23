"""
Управление конфигурацией приложения.

Загрузка из config.json с fallback на значения по умолчанию.
Сохранение изменений в config.json.
"""

import json
import logging
import os

logger = logging.getLogger("core.config")

_THIS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_CONFIG_PATH = os.path.join(_THIS_DIR, "config.json")

_DEFAULT_CONFIG = {
    "default_adapter": "lora_adapter_v2.1",
    "hide_clean_sentences": True,
    "strict_protection": False,
    "auditor_format": False,
    "skip_tables": False,
    "word_blocklist": [],
}


def load_config():
    """Загрузить конфигурацию из config.json.

    При отсутствии файла или некорректном JSON возвращает копию _DEFAULT_CONFIG.

    Returns:
        dict: Конфигурация.
    """
    config = _DEFAULT_CONFIG.copy()
    try:
        with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        config.update(data)
    except (FileNotFoundError, json.JSONDecodeError):
        pass
    return config


def save_config(config):
    """Сохранить конфигурацию в config.json.

    Args:
        config: Словарь конфигурации.
    """
    try:
        with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except OSError as e:
        logger.error("Failed to save config: %s", e)
