"""Пути к ресурсам программы (шаблоны, заглушка постера, кэш).

При запуске из сборки PyInstaller ресурсы лежат в каталоге ``sys._MEIPASS``,
при запуске из исходников в корне проекта.
"""

from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image

TEMPLATE_A4 = "template.docx"
TEMPLATE_A5 = "template_a5.docx"
NO_POSTER = "no_poster.jpg"
CACHE_NAME = "cache"


def base_path() -> Path:
    """Каталог, в котором лежат ресурсы."""
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    return Path(__file__).resolve().parent.parent


def resource_path(name: str) -> Path:
    return base_path() / name


def template_path(a5: bool = False) -> Path:
    return resource_path(TEMPLATE_A5 if a5 else TEMPLATE_A4)


def no_poster() -> Image.Image:
    return Image.open(resource_path(NO_POSTER))


def cache_path() -> Path:
    return resource_path(CACHE_NAME)
