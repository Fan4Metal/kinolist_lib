from mutagen.mp4 import MP4

from kinolist import tags
from kinolist.models import Film
from kinolist.tags import clear_tags, parse_year, read_tags, write_tags


def test_parse_year():
    assert parse_year("1984") == 1984
    assert parse_year("2019-05-01") == 2019
    assert parse_year("") is None
    assert parse_year("n/a") is None


def test_write_and_read_round_trip(mp4_file, film):
    assert write_tags(film, mp4_file)
    loaded = read_tags(mp4_file)
    assert loaded is not None
    assert loaded.title == film.title
    assert loaded.year == film.year
    assert loaded.rating == film.rating
    assert loaded.countries == film.countries
    assert loaded.description == film.description
    assert loaded.directors == film.directors
    assert loaded.actors == film.actors
    assert loaded.kp_id == film.kp_id
    assert loaded.genres == film.genres
    assert loaded.main_genre == film.main_genre
    assert loaded.poster.size == film.poster.size


def test_legacy_tag_layout(mp4_file, film):
    """Раскладка тегов совпадает с прежними версиями программы."""
    write_tags(film, mp4_file)
    video = MP4(mp4_file)
    assert video[tags.TITLE] == ["Терминатор"]
    assert video[tags.YEAR] == ["1984"]
    assert video[tags.GENRE] == ["фантастика"]
    assert video[tags.DIRECTORS][0].decode() == "Джеймс Кэмерон"
    assert video[tags.COUNTRIES][0].decode() == "США;Великобритания"
    assert video[tags.GENRES][0].decode() == "фантастика;боевик;триллер"
    assert video[tags.KP_ID][0].decode() == "507"
    assert video[tags.RATING][0].decode() == "8.0"
    actors = video[tags.ACTORS][0].decode().split("\r\n")
    assert actors[:4] == ["", "Арнольд Шварценеггер", "", "Линда Хэмилтон"]


def test_read_tags_tolerates_missing_and_odd_values(mp4_file):
    video = MP4(mp4_file)
    video[tags.TITLE] = "Фильм"
    video[tags.YEAR] = "2019-05-01"
    video[tags.COUNTRIES] = tags._freeform("США")
    video.save()
    loaded = read_tags(mp4_file)
    assert loaded.title == "Фильм"
    assert loaded.year == 2019
    assert loaded.countries == ["США"]
    assert loaded.rating == "" and loaded.kp_id is None
    assert loaded.actors == [] and loaded.directors == []
    assert loaded.poster is not None


def test_read_tags_skips_files_tagged_by_other_tools(mp4_file):
    """Файл с одним названием от релиз-группы не должен превращаться в пустую карточку."""
    video = MP4(mp4_file)
    video[tags.TITLE] = "-= HDee =-"
    video[tags.YEAR] = "2020"
    video.save()
    assert read_tags(mp4_file) is None


def test_empty_description_written_as_space(mp4_file):
    write_tags(Film("Без описания"), mp4_file)
    assert MP4(mp4_file)[tags.DESCRIPTION] == [" "]
    assert read_tags(mp4_file).description == ""


def test_clear_tags(mp4_file, film):
    write_tags(film, mp4_file)
    assert clear_tags(mp4_file)
    assert read_tags(mp4_file) is None


def test_missing_file():
    assert read_tags("no_such_file.mp4") is None
    assert not write_tags(Film("x"), "no_such_file.mp4")
    assert not clear_tags("no_such_file.mp4")
