import logging

import pytest

from kinolist import cli
from kinolist.kinopoisk import Kinopoisk, SearchResult


class FakeKinopoisk:
    """Клиент, находящий фильм по любому из его названий без обращения к сети."""

    def __init__(self, films: dict, aliases: dict):
        self.films = films
        self.aliases = aliases

    def search(self, query, year=None):
        kp_id = self.aliases.get(query.lower())
        if kp_id is None:
            return None
        film = self.films[kp_id]
        return SearchResult(kp_id, film.title, film.year)

    search_any = Kinopoisk.search_any

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


def test_movie_test_mode_creates_nothing(fake_api, tmp_path, capsys):
    output = tmp_path / "list.docx"
    assert cli.main(["-m", "Терминатор", "Нет такого", "-o", str(output), "--test"]) == 0
    assert not output.exists()
    out = capsys.readouterr().out
    assert "√ Терминатор (1984)  KP 507" in out
    assert "× Нет такого" in out
    assert "Найдено: 1, не найдено: 1" in out
    assert "\nНе найдены (1)\n   1. Нет такого\n" in out
    assert "\x1b[" not in out  # без терминала цвета отключены


def test_cover_names_output_file(fake_api, mp4_file, tmp_path, monkeypatch):
    movies = tmp_path / "Рекомендации #231"
    movies.mkdir()
    (movies / "Terminator.mp4").write_bytes(open(mp4_file, "rb").read())
    assert cli.main(["-t", str(movies / "Terminator.mp4")]) == 0
    monkeypatch.chdir(tmp_path)
    assert cli.main(["--loc", str(movies), "--cover", "--cover-name", "--txtlist"]) == 0
    assert (tmp_path / "Рекомендации #231.docx").exists()
    assert (tmp_path / "Рекомендации #231.txt").exists()
    assert cli.main(["--loc", str(movies), "--cover", 'Топ: "лучшее"?', "--cover-name", "-o", "out/x.docx"]) == 0
    assert (tmp_path / "out" / "Топ лучшее.docx").exists()
    assert cli.main(["--loc", str(movies), "--cover"]) == 0
    assert (tmp_path / "list.docx").exists()


def test_search_falls_back_to_title_variants(fake_api, tmp_path, capsys):
    assert cli.main(["-m", "Терминатор (The Terminator) 1984", "--test"]) == 0
    out = capsys.readouterr().out
    assert "√ Терминатор (1984)  KP 507, по запросу «Терминатор»" in out


def test_library_log_messages_go_to_console(fake_api, capsys):
    logging.getLogger("kinolist.tags").error("file.mp4: не удалось открыть файл")
    assert cli.main(["--cleartags", "no_such_dir_or_file"]) == 0
    out = capsys.readouterr().out
    assert "× file.mp4: не удалось открыть файл" in out
    assert "Ошибка: неверно указан путь." in out


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
    assert cli.main(["--cleartags", mp4_file, "--no-confirm"]) == 0
    assert cli.main(["--loc", str(tmp_path), "-o", str(tmp_path / "empty.docx")]) == 0
    assert not (tmp_path / "empty.docx").exists()


def test_loc_with_cover(fake_api, mp4_file, tmp_path):
    from docx import Document

    def cover_texts(path):
        return {text.text for text in Document(str(path)).element.body[0].xpath(".//w:t")}

    assert cli.main(["-t", mp4_file]) == 0
    by_dir = tmp_path / "by_dir.docx"
    assert cli.main(["--loc", str(tmp_path), "-o", str(by_dir), "--cover"]) == 0
    assert cover_texts(by_dir) == {tmp_path.name}
    custom = tmp_path / "custom.docx"
    assert cli.main(["--loc", str(tmp_path), "-o", str(custom), "--cover", "Свой текст", "--newformat"]) == 0
    assert cover_texts(custom) == {"Свой текст"}


def test_cleartags_with_confirmation(fake_api, mp4_file, monkeypatch):
    from kinolist.tags import read_tags

    assert cli.main(["-t", mp4_file, "-kp", "507"]) == 0
    monkeypatch.setattr("builtins.input", lambda prompt: "n")
    assert cli.main(["--cleartags", mp4_file]) == 0
    assert read_tags(mp4_file) is not None
    monkeypatch.setattr("builtins.input", lambda prompt: "y")
    assert cli.main(["--cleartags", mp4_file]) == 0
    assert read_tags(mp4_file) is None

    assert cli.main(["-t", mp4_file, "-kp", "507"]) == 0
    monkeypatch.setattr("builtins.input", lambda prompt: pytest.fail("подтверждение не должно запрашиваться"))
    assert cli.main(["--cleartags", mp4_file, "--no-confirm"]) == 0
    assert read_tags(mp4_file) is None


def test_tag_with_explicit_kp_id(fake_api, mp4_file):
    from kinolist.tags import read_tags

    assert cli.main(["-t", mp4_file, "-kp", "507"]) == 0
    assert read_tags(mp4_file).kp_id == 507
