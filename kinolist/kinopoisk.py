"""Клиент неофициального API Кинопоиска (kinopoiskapiunofficial.tech).

Используются три метода API: поиск по ключевому слову, карточка фильма и список персонала.
"""

from __future__ import annotations

import io
import logging
import re
import textwrap
import time
from dataclasses import dataclass

import requests
from PIL import Image

from .models import Film, main_genre
from .resources import no_poster

log = logging.getLogger(__name__)

API_URL = "https://kinopoiskapiunofficial.tech"
KP_TAG_RE = re.compile(r"KP~(\d+)")
PAREN_RE = re.compile(r"\(([^()]*)\)")
TRAILING_YEAR_RE = re.compile(r"(?:\((\d{4})\)|\b(\d{4}))\s*$")
YEAR_TOLERANCE = 1
POSTER_SIZE = (360, 540)
POSTER_RATIO = 1.5
DESCRIPTION_LIMIT = 665
ACTORS_LIMIT = 10
REQUEST_TIMEOUT = 30


class KinopoiskError(Exception):
    """Ошибка обращения к API."""


@dataclass(frozen=True)
class SearchResult:
    kp_id: int
    title: str
    year: str | int | None


def kp_id_from_title(title: str) -> int | None:
    """Находит тег ``KP~xxx`` в названии и возвращает xxx (kinopoisk id)."""
    match = KP_TAG_RE.search(title)
    return int(match.group(1)) if match else None


def _strip_year(text: str) -> str:
    return TRAILING_YEAR_RE.sub("", text).strip(" -–—.,;:")


def parse_title(title: str) -> tuple[list[str], int | None]:
    """Разбирает строку вида ``Название (Original Title) 2006``.

    Возвращает варианты запроса в порядке убывания точности (исходная строка, текст без скобок и года,
    содержимое каждой скобки) и год, если он указан в конце строки.
    """
    title = " ".join(title.split())
    year_match = TRAILING_YEAR_RE.search(title)
    year = int(year_match.group(1) or year_match.group(2)) if year_match else None

    candidates = [title, _strip_year(" ".join(PAREN_RE.sub(" ", title).split()))]
    for group in PAREN_RE.findall(title):
        group = _strip_year(" ".join(group.split()))
        if group and not group.isdigit():
            candidates.append(group)

    variants: list[str] = []
    for candidate in candidates:
        if candidate and candidate not in variants:
            variants.append(candidate)
    return variants, year


def _film_year(film: dict) -> int | None:
    match = re.match(r"\d{4}", str(film.get("year") or ""))
    return int(match.group()) if match else None


def _pick_by_year(films: list[dict], year: int, tolerance: int = YEAR_TOLERANCE) -> dict:
    """Первый результат с точным годом, иначе ближайший в пределах допуска, иначе просто первый.

    Год в исходной строке нередко взят из другой базы и отличается на единицу, поэтому он
    служит предпочтением при выборе, а не условием поиска.
    """
    ranked = []
    for index, film in enumerate(films):
        film_year = _film_year(film)
        if film_year is not None and abs(film_year - year) <= tolerance:
            ranked.append((abs(film_year - year), index, film))
    return min(ranked)[2] if ranked else films[0]


def fit_poster(image: Image.Image) -> Image.Image:
    """Обрезает постер до соотношения сторон 2:3, уменьшает до 360x540 и переводит в RGB."""
    width, height = image.size
    if width > height / POSTER_RATIO:
        new_width = height / POSTER_RATIO
        left = (width - new_width) / 2
        image = image.crop((left, 0, left + new_width, height))
    elif height > POSTER_RATIO * width:
        new_height = width * POSTER_RATIO
        top = (height - new_height) / 2
        image = image.crop((0, top, width, top + new_height))
    image.thumbnail(POSTER_SIZE)
    return image.convert("RGB")


def shorten_description(text: str, limit: int = DESCRIPTION_LIMIT) -> str:
    text = text.replace("\n\n", " ")
    return textwrap.shorten(text, limit, fix_sentence_endings=True, break_long_words=False, placeholder="...")


class Kinopoisk:
    """Обёртка над API. Между сетевыми запросами выдерживается пауза ``delay`` секунд."""

    def __init__(self, token: str, session: requests.Session | None = None, delay: float = 0.2):
        self.token = token
        self.session = session or requests.Session()
        self.delay = delay

    def _get(self, path: str, **params):
        response = self.session.get(
            API_URL + path,
            headers={"X-API-KEY": self.token, "Content-Type": "application/json"},
            params=params or None,
            timeout=REQUEST_TIMEOUT,
        )
        if response.status_code != 200:
            raise KinopoiskError(f"HTTP {response.status_code} при запросе {path}")
        if not getattr(response, "from_cache", False):
            time.sleep(self.delay)
        return response.json()

    def search(self, query: str, year: int | None = None) -> SearchResult | None:
        """Ищет фильм по названию или по тегу ``KP~id``. Возвращает ``None``, если ничего не найдено.

        При заданном ``year`` предпочитается первый результат с таким годом выпуска.
        """
        kp_id = kp_id_from_title(query)
        if kp_id is not None:
            data = self._get(f"/api/v2.2/films/{kp_id}")
            return SearchResult(kp_id, _film_title(data), data.get("year"))

        data = self._get("/api/v2.1/films/search-by-keyword", keyword=query, page=1)
        films = data.get("films") or []
        if not data.get("searchFilmsCountResult") or not films:
            return None
        film = _pick_by_year(films, year) if year is not None else films[0]
        return SearchResult(int(film["filmId"]), film.get("nameRu") or film.get("nameEn") or "", film.get("year"))

    def search_any(self, title: str) -> tuple[SearchResult | None, str]:
        """Перебирает варианты запроса из ``parse_title``. Возвращает результат и сработавший запрос."""
        variants, year = parse_title(title)
        for query in variants:
            result = self.search(query, year)
            if result is not None:
                return result, query
        return None, title

    def film(self, kp_id: int, shorten: bool = False) -> Film:
        """Загружает полную карточку фильма вместе с постером."""
        directors, actors = self._staff(kp_id)
        data = self._get(f"/api/v2.2/films/{kp_id}")

        description = data.get("description") or ""
        if shorten:
            description = shorten_description(description)
        rating = data.get("ratingKinopoisk")
        genres = [item["genre"] for item in data.get("genres", [])]
        poster_url = data.get("posterUrl") or ""

        return Film(
            title=_film_title(data),
            year=data.get("year"),
            rating=str(rating) if rating else "",
            countries=[item["country"] for item in data.get("countries", [])],
            description=description,
            directors=directors,
            actors=actors,
            poster=self._poster(poster_url),
            kp_id=int(kp_id),
            genres=genres,
            main_genre=main_genre(genres),
            poster_url=poster_url,
            poster_preview_url=data.get("posterUrlPreview") or "",
        )

    def _staff(self, kp_id: int) -> tuple[list[str], list[str]]:
        items = self._get("/api/v1/staff", filmId=kp_id)
        directors = [_person_name(item) for item in items if item.get("professionText") == "Режиссеры"]
        actors = [_person_name(item) for item in items if item.get("professionText") == "Актеры"]
        return directors, actors[:ACTORS_LIMIT]

    def _poster(self, url: str) -> Image.Image:
        if url:
            try:
                response = self.session.get(url, timeout=REQUEST_TIMEOUT)
                if response.status_code == 200:
                    return fit_poster(Image.open(io.BytesIO(response.content)))
            except (requests.RequestException, OSError) as error:
                log.warning(f"не удалось загрузить постер {url} ({error})")
        return no_poster()


def _film_title(data: dict) -> str:
    return data.get("nameRu") or data.get("nameOriginal") or data.get("nameEn") or ""


def _person_name(item: dict) -> str:
    return item.get("nameRu") or item.get("nameEn") or ""
