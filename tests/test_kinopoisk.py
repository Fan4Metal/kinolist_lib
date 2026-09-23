import pytest
from PIL import Image

from kinolist.kinopoisk import (
    Kinopoisk,
    KinopoiskError,
    SearchResult,
    fit_poster,
    kp_id_from_title,
    parse_title,
    shorten_description,
)

FILM_JSON = {
    "kinopoiskId": 507,
    "nameRu": "Терминатор",
    "nameOriginal": "The Terminator",
    "posterUrl": "https://example.com/poster.jpg",
    "posterUrlPreview": "https://example.com/preview.jpg",
    "ratingKinopoisk": 8.0,
    "year": 1984,
    "description": "Первый абзац.\n\nВторой абзац. " + "Очень длинное описание. " * 60,
    "countries": [{"country": "США"}, {"country": "Великобритания"}],
    "genres": [{"genre": "драма"}, {"genre": "фантастика"}, {"genre": "боевик"}],
}

STAFF_JSON = [
    {"nameRu": "Джеймс Кэмерон", "nameEn": "James Cameron", "professionText": "Режиссеры"},
    {"nameRu": "", "nameEn": "Gale Anne Hurd", "professionText": "Продюсеры"},
] + [{"nameRu": f"Актер {i}", "nameEn": f"Actor {i}", "professionText": "Актеры"} for i in range(15)]


class FakeKinopoisk(Kinopoisk):
    """Клиент с подменёнными ответами API и без загрузки постера."""

    def __init__(self, responses):
        super().__init__("token", delay=0)
        self.responses = responses
        self.calls = []

    def _get(self, path, **params):
        self.calls.append((path, params))
        response = self.responses[path]
        if isinstance(response, Exception):
            raise response
        return response

    def _poster(self, url):
        return Image.new("RGB", (360, 540))


def test_kp_id_from_title():
    assert kp_id_from_title("Терминатор KP~507") == 507
    assert kp_id_from_title("Терминатор") is None


def test_search_by_keyword():
    kp = FakeKinopoisk(
        {
            "/api/v2.1/films/search-by-keyword": {
                "searchFilmsCountResult": 2,
                "films": [{"filmId": 507, "nameRu": "Терминатор", "nameEn": "The Terminator", "year": "1984"}],
            }
        }
    )
    result = kp.search("Terminator")
    assert (result.kp_id, result.title, result.year) == (507, "Терминатор", "1984")
    assert kp.calls[0][1] == {"keyword": "Terminator", "page": 1}


@pytest.mark.parametrize(
    ("title", "variants", "year"),
    [
        ("Терминатор", ["Терминатор"], None),
        (
            "Терминатор (The Terminator) 1984",
            ["Терминатор (The Terminator) 1984", "Терминатор", "The Terminator"],
            1984,
        ),
        ("24: Искупление (2008)", ["24: Искупление (2008)", "24: Искупление"], 2008),
        (
            "Грязный Гарри 2. Высшая сила (Сила магнума) (Dirty Harry 2. Magnum Force) 1973",
            [
                "Грязный Гарри 2. Высшая сила (Сила магнума) (Dirty Harry 2. Magnum Force) 1973",
                "Грязный Гарри 2. Высшая сила",
                "Сила магнума",
                "Dirty Harry 2. Magnum Force",
            ],
            1973,
        ),
        ("Элитный отряд 2 (Tropa de Elite 2 - O Inimigo ) 2011", None, 2011),
    ],
)
def test_parse_title(title, variants, year):
    parsed_variants, parsed_year = parse_title(title)
    assert parsed_year == year
    if variants is not None:
        assert parsed_variants == variants
    else:
        assert parsed_variants[1:] == ["Элитный отряд 2", "Tropa de Elite 2 - O Inimigo"]


def test_search_prefers_matching_year():
    films = [
        {"filmId": 444, "nameRu": "Терминатор 2", "year": "1991"},
        {"filmId": 507, "nameRu": "Терминатор", "year": "1984"},
    ]
    kp = FakeKinopoisk({"/api/v2.1/films/search-by-keyword": {"searchFilmsCountResult": 2, "films": films}})
    assert kp.search("Terminator").kp_id == 444
    assert kp.search("Terminator", year=1984).kp_id == 507
    assert kp.search("Terminator", year=1985).kp_id == 507  # год из другой базы отличается на единицу
    assert kp.search("Terminator", year=1990).kp_id == 444  # точное совпадение важнее порядка
    assert kp.search("Terminator", year=2000).kp_id == 444  # вне допуска: первый результат


def test_search_any_falls_back_to_simpler_queries():
    class StubKinopoisk(Kinopoisk):
        def __init__(self):
            super().__init__("token", delay=0)
            self.queries = []

        def search(self, query, year=None):
            self.queries.append((query, year))
            if query == "The Terminator":
                return SearchResult(507, "Терминатор", 1984)
            return None

    kp = StubKinopoisk()
    result, query = kp.search_any("Терминатор (The Terminator) 1984")
    assert (result.kp_id, query) == (507, "The Terminator")
    assert kp.queries == [
        ("Терминатор (The Terminator) 1984", 1984),
        ("Терминатор", 1984),
        ("The Terminator", 1984),
    ]
    assert kp.search_any("Ничего") == (None, "Ничего")


def test_search_not_found():
    kp = FakeKinopoisk({"/api/v2.1/films/search-by-keyword": {"searchFilmsCountResult": 0, "films": []}})
    assert kp.search("nothing") is None


def test_search_by_kp_tag_uses_film_endpoint():
    kp = FakeKinopoisk({"/api/v2.2/films/507": FILM_JSON})
    result = kp.search("Whatever KP~507")
    assert (result.kp_id, result.title, result.year) == (507, "Терминатор", 1984)


def test_search_raises_on_http_error():
    kp = FakeKinopoisk({"/api/v2.1/films/search-by-keyword": KinopoiskError("HTTP 402")})
    with pytest.raises(KinopoiskError):
        kp.search("x")


def test_film_mapping():
    kp = FakeKinopoisk({"/api/v2.2/films/507": FILM_JSON, "/api/v1/staff": STAFF_JSON})
    film = kp.film(507)
    assert film.title == "Терминатор"
    assert film.year == 1984
    assert film.rating == "8.0"
    assert film.countries == ["США", "Великобритания"]
    assert film.directors == ["Джеймс Кэмерон"]
    assert len(film.actors) == 10
    assert film.actors[0] == "Актер 0"
    assert film.genres == ["драма", "фантастика", "боевик"]
    assert film.main_genre == "фантастика"
    assert film.kp_id == 507
    assert film.description.startswith("Первый абзац.\n\nВторой")
    assert film.poster.size == (360, 540)


def test_film_shorten_and_missing_fields():
    data = dict(FILM_JSON, ratingKinopoisk=None, description=None, nameRu=None)
    kp = FakeKinopoisk({"/api/v2.2/films/507": data, "/api/v1/staff": []})
    film = kp.film(507, shorten=True)
    assert film.title == "The Terminator"
    assert film.rating == ""
    assert film.description == ""
    assert film.directors == [] and film.actors == []


def test_shorten_description():
    text = shorten_description(FILM_JSON["description"])
    assert len(text) <= 665
    assert text.endswith("...")
    assert "\n\n" not in text


@pytest.mark.parametrize("size", [(1000, 1000), (200, 1000), (100, 150), (400, 600)])
def test_fit_poster_ratio_and_size(size):
    poster = fit_poster(Image.new("RGBA", size))
    width, height = poster.size
    assert poster.mode == "RGB"
    assert width <= 360 and height <= 540
    assert abs(height / width - 1.5) < 0.02
