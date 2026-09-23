from kinolist.models import Film, main_genre


def test_main_genre_follows_hierarchy():
    assert main_genre(["драма", "комедия", "ужасы"]) == "ужасы"
    assert main_genre(["драма", "мелодрама"]) == "мелодрама"


def test_main_genre_fallbacks():
    assert main_genre(["вестерн", "нуар"]) == "вестерн"
    assert main_genre([]) == ""


def test_rating_text():
    assert Film("x", rating="").rating_text == "нет рейтинга"
    assert Film("x", rating="None").rating_text == "нет рейтинга"
    assert Film("x", rating="7.5").rating_text == "Кинопоиск 7.5"
    assert Film("x", rating="i6.7").rating_text == "IMDb 6.7"


def test_directors_text():
    assert Film("x").directors_text == ""
    assert Film("x", directors=["A"]).directors_text == "Режиссер: A"
    assert Film("x", directors=["A", "B"]).directors_text == "Режиссеры: A, B"
