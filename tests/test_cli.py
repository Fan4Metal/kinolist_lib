import logging

import pytest

from kinolist import cli
from kinolist.kinopoisk import SearchResult


class FakeKinopoisk:
    """Клиент, находящий фильм по любому из его названий без обращения к сети."""

    def __init__(self, films: dict, aliases: dict):
        self.films = films
        self.aliases = aliases

    def search(self, query):
        kp_id = self.aliases.get(query.lower())
        if kp_id is None:
            return None
        film = self.films[kp_id]
        return SearchResult(kp_id, film.title, film.year)

    def film(self, kp_id, shorten=False):
        return self.films[kp_id]


@pytest.fixture
def fake_api(monkeypatch, film):
    api = FakeKinopoisk({507: film}, {"терминатор": 507, "terminator": 507})
    monkeypatch.setattr(cli, "load_token", lambda: "token")
    monkeypatch.setattr(cli, "Kinopoisk", lambda token: api)
    monkeypatch.setattr(cli.requests_cache, "install_cache", lambda *a, **k: None)
    monkeypatch.setattr(cli.requests_cache, "uninstall_cache", lambda: None)
    return api


def test_version(capsys):
    with pytest.raises(SystemExit) as exit_info:
        cli.main(["--version"])
    assert exit_info.value.code == 0
    assert "Kinolist Lib" in capsys.readouterr().out


def test_output_must_be_docx(fake_api, tmp_path):
    assert cli.main(["-m", "Терминатор", "-o", str(tmp_path / "list.pdf")]) == 1


def test_movie_list_and_txt(fake_api, tmp_path):
    output = tmp_path / "out" / "list.docx"
    assert cli.main(["-m", "Терминатор", "Неизвестный", "-o", str(output), "--txtlist", "--nocache"]) == 0
    assert output.exists()
    assert (tmp_path / "out" / "list.txt").read_text(encoding="utf-8") == "Терминатор\n"


def test_movie_test_mode_creates_nothing(fake_api, tmp_path, caplog):
    caplog.set_level(logging.INFO, logger="kinolist")
    output = tmp_path / "list.docx"
    assert cli.main(["-m", "Терминатор", "-o", str(output), "--test"]) == 0
    assert not output.exists()
    assert "Найдено фильмов: 1, не найдено: 0" in caplog.text


def test_file_list(fake_api, tmp_path):
    titles = tmp_path / "movies.txt"
    titles.write_text("Терминатор\n\n", encoding="utf-8")
    output = tmp_path / "list.docx"
    assert cli.main(["-f", str(titles), "-o", str(output), "--newformat"]) == 0
    assert output.exists()


def test_tag_and_loc(fake_api, mp4_file, tmp_path):
    assert cli.main(["-t", mp4_file]) == 0
    output = tmp_path / "loc.docx"
    assert cli.main(["--loc", str(tmp_path), "-o", str(output), "--a5", "--genres", "--sort", "name"]) == 0
    assert output.exists()
    assert cli.main(["--cleartags", mp4_file]) == 0
    assert cli.main(["--loc", str(tmp_path), "-o", str(tmp_path / "empty.docx")]) == 0
    assert not (tmp_path / "empty.docx").exists()


def test_tag_with_explicit_kp_id(fake_api, mp4_file):
    from kinolist.tags import read_tags

    assert cli.main(["-t", mp4_file, "-kp", "507"]) == 0
    assert read_tags(mp4_file).kp_id == 507
