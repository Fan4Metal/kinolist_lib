"""Модель данных: карточка фильма и вспомогательные функции."""

from __future__ import annotations

import io
from dataclasses import dataclass, field

from PIL import Image

# Порядок определяет приоритет при выборе основного жанра фильма.
GENRES_HIERARCHY = [
    "мультфильм",
    "мюзикл",
    "ужасы",
    "фантастика",
    "фэнтези",
    "военный",
    "история",
    "приключения",
    "боевик",
    "триллер",
    "детектив",
    "комедия",
    "мелодрама",
    "драма",
]


def main_genre(genres: list[str], hierarchy: list[str] = GENRES_HIERARCHY) -> str:
    """Возвращает основной жанр фильма по иерархии, либо первый жанр из списка."""
    if not genres:
        return ""
    genres_set = set(genres)
    for genre in hierarchy:
        if genre in genres_set:
            return genre
    return genres[0]


def image_to_png(image: Image.Image) -> io.BytesIO:
    """Возвращает изображение как файлоподобный объект в формате PNG."""
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    buffer.seek(0)
    return buffer


@dataclass
class Film:
    """Карточка фильма.

    Поле ``rating`` хранится строкой: ``"7.5"`` рейтинг Кинопоиска, ``"i6.7"`` рейтинг IMDb,
    пустая строка означает отсутствие рейтинга.
    """

    title: str
    year: int | None = None
    rating: str = ""
    countries: list[str] = field(default_factory=list)
    description: str = ""
    directors: list[str] = field(default_factory=list)
    actors: list[str] = field(default_factory=list)
    poster: Image.Image | None = None
    kp_id: int | None = None
    genres: list[str] = field(default_factory=list)
    main_genre: str = ""
    poster_url: str = ""
    poster_preview_url: str = ""

    @property
    def rating_text(self) -> str:
        """Рейтинг в виде текста для списков."""
        if not self.rating or self.rating == "None":
            return "нет рейтинга"
        if self.rating.startswith("i"):
            return f"IMDb {self.rating[1:]}"
        return f"Кинопоиск {self.rating}"

    @property
    def directors_text(self) -> str:
        """Строка вида ``Режиссер: X`` или ``Режиссеры: X, Y``; пустая, если режиссёров нет."""
        if not self.directors:
            return ""
        label = "Режиссеры" if len(self.directors) > 1 else "Режиссер"
        return f"{label}: {', '.join(self.directors)}"
