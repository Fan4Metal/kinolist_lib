"""Работа с файлами: поиск mp4, ярлыки, сортировка, текстовые списки, разбор имён торрентов."""

from __future__ import annotations

import glob
import logging
import os
from collections.abc import Callable
from dataclasses import dataclass

import PTN

log = logging.getLogger(__name__)

VIDEO_EXT = ".mp4"
LNK_EXT = ".lnk"
FORBIDDEN_CHARS = '\\/:*?"<>|'

# Ключ параметра --sort: (функция ключа, обратный порядок, описание).
SORT_OPTIONS: dict[str, tuple[Callable[[str], object], bool, str]] = {
    "date": (os.path.getctime, False, "по дате создания"),
    "date_r": (os.path.getctime, True, "по дате создания в обратном порядке"),
    "datem": (os.path.getmtime, False, "по дате изменения"),
    "datem_r": (os.path.getmtime, True, "по дате изменения в обратном порядке"),
    "name": (os.path.basename, False, "по имени"),
    "name_r": (os.path.basename, True, "по имени в обратном порядке"),
}


def lnk_target(lnk_path: str) -> str:
    """Путь, на который указывает ярлык Windows (.lnk)."""
    import win32com.client

    shell = win32com.client.Dispatch("WScript.Shell")
    return shell.CreateShortCut(lnk_path).Targetpath


def find_mp4_files(directory: str, follow_lnk: bool = False) -> list[str]:
    """Список mp4-файлов в каталоге. При ``follow_lnk`` учитываются и ярлыки на mp4-файлы."""
    files = glob.glob(os.path.join(directory, f"*{VIDEO_EXT}"))
    if follow_lnk:
        for lnk in glob.glob(os.path.join(directory, f"*{LNK_EXT}")):
            try:
                target = lnk_target(lnk)
            except Exception as error:
                log.warning(f"Не удалось прочитать ярлык {os.path.basename(lnk)}: {error}")
                continue
            if is_mp4(target) and os.path.isfile(target):
                files.append(target)
    return files


def sort_files(files: list[str], option: str | None) -> tuple[list[str], str]:
    """Сортирует список файлов по варианту из ``SORT_OPTIONS``. Возвращает список и описание сортировки."""
    key, reverse, message = SORT_OPTIONS.get(option or "", SORT_OPTIONS["name"])
    return sorted(files, key=key, reverse=reverse), message


def is_mp4(path: str) -> bool:
    return os.path.splitext(path)[1].lower() == VIDEO_EXT


def file_title(path: str) -> str:
    """Имя файла без каталога и расширения."""
    return os.path.splitext(os.path.basename(path))[0]


def read_lines(path: str) -> list[str]:
    """Непустые строки текстового файла."""
    with open(path, encoding="utf-8") as file:
        return [line.strip() for line in file if line.strip()]


def write_lines(path: str, lines: list[str]) -> None:
    with open(path, "w", encoding="utf-8") as file:
        for line in lines:
            file.write(line + "\n")


def safe_filename(name: str) -> str:
    """Удаляет символы, запрещённые в именах файлов Windows."""
    return name.translate(str.maketrans("", "", FORBIDDEN_CHARS))


def torrent_title(path: str) -> str:
    """Название фильма, извлечённое из имени торрент-файла, либо пустая строка."""
    return PTN.parse(file_title(path)).get("title", "") or ""


@dataclass(frozen=True)
class Rename:
    source: str
    destination: str


def rename_destination(path: str, title: str, year: str | int | None) -> str:
    """Новый путь файла вида ``Название (год).ext`` в том же каталоге."""
    ext = os.path.splitext(path)[1]
    name = f"{safe_filename(title)} ({year}){ext}" if year else f"{safe_filename(title)}{ext}"
    return os.path.join(os.path.dirname(path), name)


def apply_renames(renames: list[Rename]) -> None:
    for item in renames:
        try:
            os.rename(item.source, item.destination)
            log.info(f"Переименование файла: {item.source} -> {item.destination}")
        except OSError as error:
            log.error(f"Ошибка переименования файла: {item.source} -> {item.destination}")
            log.error(error)
