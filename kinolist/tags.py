"""Чтение и запись тегов в файлы mp4.

Формат тегов совместим с предыдущими версиями программы: все данные карточки фильма
хранятся в стандартных атомах и в свободных атомах ``----:com.apple.iTunes:*``.
"""

from __future__ import annotations

import io
import logging
import os
import re

from mutagen import MutagenError
from mutagen.mp4 import MP4, AtomDataType, MP4Cover, MP4FreeForm
from PIL import Image

from .models import Film, image_to_png
from .resources import no_poster

log = logging.getLogger(__name__)

TITLE = "\xa9nam"
YEAR = "\xa9day"
GENRE = "\xa9gen"
DESCRIPTION = "desc"
LONG_DESCRIPTION = "ldes"
COVER = "covr"
DIRECTORS = "----:com.apple.iTunes:DIRECTOR"
ACTORS = "----:com.apple.iTunes:Actors"
RATING = "----:com.apple.iTunes:kpra"
COUNTRIES = "----:com.apple.iTunes:countr"
KP_ID = "----:com.apple.iTunes:kpid"
GENRES = "----:com.apple.iTunes:genre"

LIST_SEPARATOR = ";"
ACTORS_SEPARATOR = "\r\n"


def _freeform(text: str) -> MP4FreeForm:
    return MP4FreeForm(text.encode(), AtomDataType.UTF8)


def _open(file_path: str) -> MP4 | None:
    try:
        return MP4(file_path)
    except (MutagenError, OSError) as error:
        log.error(f"{os.path.basename(file_path)}: не удалось открыть файл ({error})")
        return None


def write_tags(film: Film, file_path: str) -> bool:
    """Записывает карточку фильма в теги mp4. Существующие теги удаляются."""
    video = _open(file_path)
    if video is None:
        return False
    try:
        video.delete()
    except Exception as error:
        log.error(f"{os.path.basename(file_path)}: теги не записаны ({error})")
        return False

    video[TITLE] = film.title
    # Пустая строка в desc не сохраняется, поэтому вместо неё записывается пробел.
    video[DESCRIPTION] = film.description or " "
    video[LONG_DESCRIPTION] = film.description or " "
    if film.year:
        video[YEAR] = str(film.year)
    poster = film.poster or no_poster()
    video[COVER] = [MP4Cover(image_to_png(poster).getvalue(), imageformat=MP4Cover.FORMAT_PNG)]
    video[DIRECTORS] = _freeform(LIST_SEPARATOR.join(film.directors))
    # Актёры хранятся парами "пустая строка, имя", разделёнными переводом строки.
    actors_lines: list[str] = []
    for actor in film.actors:
        actors_lines.extend(("", actor))
    video[ACTORS] = _freeform(ACTORS_SEPARATOR.join(actors_lines))
    video[RATING] = _freeform(film.rating or "")
    video[COUNTRIES] = _freeform(LIST_SEPARATOR.join(film.countries))
    video[KP_ID] = _freeform(str(film.kp_id) if film.kp_id is not None else "")
    video[GENRES] = _freeform(LIST_SEPARATOR.join(film.genres))
    video[GENRE] = film.main_genre or ""

    try:
        video.save()
    except Exception as error:
        log.error(f"{os.path.basename(file_path)}: теги не записаны ({error})")
        return False
    return True


def _text(video: MP4, key: str) -> str:
    """Текстовое значение тега (стандартного или свободного), пустая строка при отсутствии."""
    values = video.get(key)
    if not values:
        return ""
    value = values[0]
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return str(value)


def _list(video: MP4, key: str, separator: str = LIST_SEPARATOR) -> list[str]:
    return [item for item in _text(video, key).split(separator) if item]


def parse_year(value: str) -> int | None:
    """Извлекает год из строки вида ``2019`` или ``2019-05-01``."""
    match = re.match(r"\s*(\d{4})", value)
    return int(match.group(1)) if match else None


def read_tags(file_path: str) -> Film | None:
    """Читает карточку фильма из тегов mp4. Возвращает ``None``, если тегов нет или файл не открылся."""
    video = _open(file_path)
    if video is None or not video.tags:
        return None
    title = _text(video, TITLE)
    if not title:
        return None

    poster = None
    covers = video.get(COVER)
    if covers:
        try:
            poster = Image.open(io.BytesIO(bytes(covers[0])))
        except OSError:
            poster = None

    kp_id = _text(video, KP_ID)
    return Film(
        title=title,
        year=parse_year(_text(video, YEAR)),
        rating=_text(video, RATING),
        countries=_list(video, COUNTRIES),
        description=_text(video, DESCRIPTION).strip(),
        directors=_list(video, DIRECTORS),
        actors=_text(video, ACTORS).split(ACTORS_SEPARATOR)[1::2],
        poster=poster or no_poster(),
        kp_id=int(kp_id) if kp_id.isdigit() else None,
        genres=_list(video, GENRES),
        main_genre=_text(video, GENRE),
    )


def clear_tags(file_path: str) -> bool:
    """Удаляет все теги из файла mp4."""
    video = _open(file_path)
    if video is None:
        return False
    try:
        video.delete()
        video.save()
    except Exception as error:
        log.error(f"{os.path.basename(file_path)}: теги не удалены ({error})")
        return False
    return True
