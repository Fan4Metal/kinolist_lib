import os
import time

from kinolist.files import (
    file_title,
    find_mp4_files,
    is_mp4,
    read_lines,
    rename_destination,
    safe_filename,
    sort_files,
    torrent_title,
    write_lines,
)


def test_is_mp4_and_title():
    assert is_mp4("a/b/Film.MP4")
    assert not is_mp4("a/b/Film.mkv")
    assert file_title("a/b/Film (1984).mp4") == "Film (1984)"


def test_find_mp4_files(tmp_path):
    (tmp_path / "a.mp4").write_bytes(b"")
    (tmp_path / "b.mkv").write_bytes(b"")
    files = find_mp4_files(str(tmp_path))
    assert [os.path.basename(f) for f in files] == ["a.mp4"]


def test_sort_files(tmp_path):
    first = tmp_path / "b.mp4"
    second = tmp_path / "a.mp4"
    first.write_bytes(b"")
    time.sleep(0.05)
    second.write_bytes(b"")
    files = [str(first), str(second)]

    by_name, message = sort_files(files, "name")
    assert by_name == [str(second), str(first)] and message == "по имени"
    assert sort_files(files, "name_r")[0] == [str(first), str(second)]
    assert sort_files(files, "datem")[0] == [str(first), str(second)]
    assert sort_files(files, "datem_r")[0] == [str(second), str(first)]
    assert sort_files(files, None)[0] == by_name
    assert sort_files(files, "unknown")[0] == by_name


def test_read_write_lines(tmp_path):
    path = tmp_path / "list.txt"
    write_lines(str(path), ["Один", "Два"])
    path.write_text(path.read_text(encoding="utf-8") + "\n   \nТри  \n", encoding="utf-8")
    assert read_lines(str(path)) == ["Один", "Два", "Три"]


def test_safe_filename():
    assert safe_filename('Что: "Где"? <Когда>|*/\\') == "Что Где Когда"


def test_torrent_title():
    assert torrent_title("x/The.Terminator.1984.1080p.BluRay.x264.mp4") == "The Terminator"


def test_rename_destination():
    assert rename_destination(os.path.join("dir", "old.mp4"), "Терминатор: 2", 1991) == os.path.join(
        "dir", "Терминатор 2 (1991).mp4"
    )
    assert rename_destination("old.mp4", "Фильм", None) == "Фильм.mp4"
