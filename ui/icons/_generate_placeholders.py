"""
Однократный скрипт для генерации PNG-заглушек иконок.

Создаёт цветной квадрат 64×64 с короткой подписью для каждого ключа
из ICONS (см. ui/icons.py). Запускать вручную один раз:

    python ui/icons/_generate_placeholders.py

После — заменяйте файлы по тем же путям своими финальными иконками.
"""

import hashlib
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

from PIL import Image, ImageDraw, ImageFont  # noqa: E402

from ui.icons import ICONS, ICON_DIR, ICON_FALLBACK_TEXT  # noqa: E402


# Палитра приятных цветов
PALETTE = [
    "#2B579A",  # синий Word
    "#217346",  # зелёный Excel
    "#0078D4",  # светло-синий Outlook
    "#D24726",  # оранжевый
    "#7B83EB",  # фиолетовый
    "#5B6770",  # серо-синий
    "#28A745",  # зелёный
    "#DC3545",  # красный
    "#FD7E14",  # оранжево-жёлтый
    "#6F42C1",  # фиолетовый
    "#17A2B8",  # циан
    "#E83E8C",  # розовый
    "#20C997",  # бирюзовый
]


def color_for_key(key: str) -> str:
    """Детерминированный цвет для ключа."""
    h = hashlib.md5(key.encode("utf-8")).digest()
    return PALETTE[h[0] % len(PALETTE)]


def _font(size: int) -> ImageFont.ImageFont:
    """Подобрать TTF-шрифт для подписи. Не падает, если ни один не найден."""
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial Bold.ttf",
        "/Library/Fonts/Arial.ttf",
        "C:\\Windows\\Fonts\\arialbd.ttf",
        "C:\\Windows\\Fonts\\arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for path in candidates:
        if os.path.isfile(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    return ImageFont.load_default()


def generate_one(key: str, filename: str):
    out_path = os.path.join(ICON_DIR, filename)
    size = 64
    bg = color_for_key(key)
    img = Image.new("RGBA", (size, size), bg)
    draw = ImageDraw.Draw(img)

    label = ICON_FALLBACK_TEXT.get(key, key[:2].upper())
    # Если fallback состоит из одного экзотического символа — берём 1-2 буквы ключа
    if len(label) > 3:
        label = label[:2]
    font = _font(28 if len(label) <= 2 else 22)

    bbox = draw.textbbox((0, 0), label, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    pos = ((size - tw) // 2 - bbox[0], (size - th) // 2 - bbox[1])
    draw.text(pos, label, fill="#ffffff", font=font)

    img.save(out_path, "PNG")
    print(f"  → {out_path}")


def main():
    os.makedirs(ICON_DIR, exist_ok=True)
    print(f"Генерация {len(ICONS)} иконок в {ICON_DIR}")
    for key, filename in ICONS.items():
        generate_one(key, filename)
    print("Готово.")


if __name__ == "__main__":
    main()
