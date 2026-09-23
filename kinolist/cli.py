"""Интерфейс командной строки ``kl``."""

from __future__ import annotations

import logging
import os
from pathlib import Path

import requests
import requests_cache
from tqdm import tqdm

from . import __version__
from .argparse_ru import argparse
from .docx_out import txt_path_for, write_simple_list, write_table_list, write_txt_list
from .files import (
    Rename,
    apply_renames,
    file_title,
    find_mp4_files,
    is_mp4,
    read_lines,
    rename_destination,
    sort_files,
    torrent_title,
)
from .kinopoisk import Kinopoisk, KinopoiskError, SearchResult
from .models import Film
from .resources import cache_path, template_path
from .tags import clear_tags, read_tags, write_tags

log = logging.getLogger("kinolist")

DEFAULT_OUTPUT = "list.docx"
CACHE_EXPIRE_SECONDS = 3600

EPILOG = R"""
Примеры:
kl -m "Terminator" "Terminator 2" KP~319  --создает список list.docx из 3 фильмов: Terminator,
                                                Terminator 2 и Terminator 3 (*)
kl -f movies.txt -o movies.docx           --создает список movies.docx из всех фильмов в файле movies.txt
kl -t ./Terminator.mp4                    --записывает теги в файл Terminator.mp4 в текущем каталоге
kl -t c:\movies\Terminator.mp4            --записывает теги в файл Terminator.mp4 в каталоге c:\movies
kl -t c:\movies\Chuzhie.mp4 -kp 406       --записывает в файл Chuzhie.mp4 теги фильма Чужие (Kinopoisk_id 406)
kl -t                                     --записывает теги во все mp4 файлы в текущем каталоге
kl -t c:\movies                           --записывает теги во все mp4 файлы в каталоге c:\movies
kl --cleartags                            --удаляет все теги во всех mp4 файлах в текущем каталоге
kl -r *.mp4                               --переименовывает mp4 файлы в текущем каталоге (торрент -> название.mp4)
kl -l                                     --создает список list.docx из всех mp4 файлов в текущем каталоге.
kl --loc                                  --создает список list.docx из всех mp4 файлов в текущем каталоге, используя
                                                только теги файлов (все теги должны быть предварительно записаны в
                                                файл). Рейтинг в теге kpra, начинающийся с "i" (например: i6.7),
                                                интерпретируется как рейтинг IMDb.
kl --loc --newformat                      --создает список из тегов файлов в новом формате
kl --loc --a5                             --создает список из тегов файлов в формате A5 (для планшетов)
kl --loc --cover                          --создает список из тегов файлов с обложкой, заголовок обложки
                                                берется из имени каталога
kl --loc --cover "Рекомендации #231"      --создает список из тегов файлов с обложкой с указанным заголовком


* Можно указать Kinopoisk_id напрямую, используя тег KP~XXX в названии фильма (где XXX - Kinopoisk_id)
"""

SORT_HELP = (
    "Сортировка списка по тегам. Варианты: date - по дате создания, date_r - по дате создания в обратном порядке, "
    "datem - по дате изменения, datem_r - по дате изменения в обратном порядке, name - по имени, "
    "name_r - по имени в обратном порядке"
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="kl",
        description=f"Библиотека для создания списков фильмов в формате docx. Версия {__version__}.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=EPILOG,
    )
    cwd = os.getcwd()
    add = parser.add_argument
    add(
        "-ver",
        "--version",
        action="version",
        version=f"Kinolist Lib {__version__}",
        help="выводит версию программы и завершает работу",
    )
    add("-f", "--file", help="создает список фильмов в формате docx из текстового файла в формате txt")
    add("--txtlist", action="store_true", help="дополнительно сохраняет текстовый список с названиями фильмов")
    add("-m", "--movie", nargs="+", help="создает список фильмов в формате docx из указанных фильмов")
    add(
        "--test",
        action="store_true",
        help="тестовый поиск фильмов без создания списка, работает с параметрами --file, --movie и --tag",
    )
    add("-o", "--output", help="имя выходного файла (list.docx по умолчанию)")
    add(
        "-s",
        "--shorten",
        action="store_true",
        help="сокращает описания фильмов, чтобы поместились два фильма на странице",
    )
    add(
        "-t", "--tag", nargs="?", const=cwd, help="записывает теги в файл mp4 (или во все mp4 файлы в текущем каталоге)"
    )
    add("-kp", "--kinopoisk_id", type=int, help="указывает значение kinopoisk_id для записи в тег")
    add(
        "--cleartags",
        nargs="?",
        const=cwd,
        help="удаляет все теги в файле mp4 (или во всех mp4 файлах в текущем каталоге)",
    )
    add("-r", "--rename", nargs="?", const=cwd, help="переименовывает mp4 файлы в текущем каталоге")
    add(
        "-l",
        "--list",
        nargs="?",
        const=cwd,
        help="создает список фильмов в формате docx из mp4 файлов в текущем каталоге",
    )
    add(
        "--loc",
        nargs="?",
        const=cwd,
        help="создает список фильмов в формате docx из тегов mp4 файлов в текущем каталоге",
    )
    add("-nf", "--newformat", action="store_true", help="модификатор для создания списка фильмов в новом формате")
    add("-g", "--genres", action="store_true", help="модификатор добавляет жанры в список фильмов")
    add("--a5", action="store_true", help="список в формате A5 (для списков с постерами)")
    add(
        "--cover",
        nargs="?",
        const="",
        metavar="ТЕКСТ",
        help="добавляет обложку первой страницей; без текста заголовок берется из имени каталога или файла",
    )
    add("--sort", help=SORT_HELP)
    add("--nocache", action="store_true", help="не использовать кэш")
    add("--clearcache", action="store_true", help="очистить кэш")
    return parser


def load_token() -> str | None:
    """Токен API из модуля ``config`` (файл config.py рядом с программой)."""
    try:
        from config import KINOPOISK_API_TOKEN
    except ImportError:
        log.error("Не найден файл config.py с переменной KINOPOISK_API_TOKEN.")
        return None
    return KINOPOISK_API_TOKEN


def resolve_output(output: str | None) -> str | None:
    """Проверяет имя выходного файла и создаёт его каталог."""
    if not output:
        return DEFAULT_OUTPUT
    if os.path.splitext(output)[1].lower() != ".docx":
        log.error("Выходной файл должен иметь расширение docx.")
        return None
    output_dir = os.path.dirname(output)
    if output_dir:
        Path(output_dir).mkdir(parents=True, exist_ok=True)
    return output


def resolve_titles(kp: Kinopoisk, titles: list[str]) -> tuple[list[SearchResult], list[str]]:
    """Ищет фильмы по названиям. Возвращает найденные фильмы и список ненайденных названий."""
    found: list[SearchResult] = []
    not_found: list[str] = []
    for title in titles:
        try:
            result = kp.search(title)
        except (KinopoiskError, requests.RequestException) as error:
            log.warning(f"Ошибка поиска «{title}»: {error}")
            result = None
        if result is None:
            log.info(f"{title} не найден")
            not_found.append(title)
        else:
            log.info(f"Найден фильм: {result.title} ({result.year}), kinopoisk id: {result.kp_id}")
            found.append(result)
    return found, not_found


def load_films(kp: Kinopoisk, kp_ids: list[int], shorten: bool = False) -> list[Film]:
    films: list[Film] = []
    for kp_id in tqdm(kp_ids, desc="Загрузка информации...   "):
        try:
            films.append(kp.film(kp_id, shorten))
        except (KinopoiskError, requests.RequestException, KeyError, ValueError) as error:
            log.warning(f"Не удалось загрузить фильм {kp_id}: {error}")
    return films


def cover_title(args: argparse.Namespace, default: str) -> str | None:
    """Заголовок обложки: текст параметра ``--cover``, либо ``default``, если параметр указан без текста."""
    if args.cover is None:
        return None
    return args.cover or default


def dir_name(path: str) -> str:
    return os.path.basename(os.path.abspath(path))


def save_lists(films: list[Film], output: str, args: argparse.Namespace, cover_default: str = "") -> None:
    cover = cover_title(args, cover_default)
    if args.newformat:
        write_simple_list(films, output, genres=args.genres, cover=cover)
    else:
        write_table_list(films, output, template_path(a5=args.a5), genres=args.genres, cover=cover)
    if args.txtlist:
        write_txt_list(films, txt_path_for(output))


def make_list_from_titles(
    kp: Kinopoisk, titles: list[str], output: str, args: argparse.Namespace, cover_default: str = ""
) -> None:
    found, not_found = resolve_titles(kp, titles)
    for title in not_found:
        log.warning(f"Фильм не найден: {title}")
    if args.test:
        log.info(f"Найдено фильмов: {len(found)}, не найдено: {len(not_found)}")
        return
    if not found:
        log.warning("Фильмы не найдены.")
        return
    films = load_films(kp, [item.kp_id for item in found], args.shorten)
    if not films:
        log.error("Ошибка, список не создан!")
        return
    save_lists(films, output, args, cover_default)


def list_mp4_dir(path: str, follow_lnk: bool = False) -> list[str]:
    log.info(f"Поиск файлов mp4 в каталоге: {os.path.abspath(path)}")
    files = find_mp4_files(path, follow_lnk)
    if not files:
        log.warning(f'В каталоге "{path}" файлы mp4 не найдены.')
        return []
    for file in files:
        log.info(f"Найден файл: {os.path.basename(file)}")
    log.info(f"Всего файлов: {len(files)}")
    return files


def tag_file(kp: Kinopoisk, path: str, kp_id: int | None = None) -> bool:
    """Записывает теги в один файл. Фильм ищется по имени файла, если id не задан."""
    name = os.path.basename(path)
    if kp_id is None:
        found, _ = resolve_titles(kp, [file_title(path)])
        if not found:
            log.warning(f"Фильм не найден: {name}")
            return False
        kp_id = found[0].kp_id
    try:
        film = kp.film(kp_id)
    except (KinopoiskError, requests.RequestException, KeyError, ValueError) as error:
        log.warning(f"Не удалось загрузить фильм {kp_id}: {error}")
        return False
    if not write_tags(film, path):
        log.warning(f"Тег не записан в файл: {name}")
        return False
    log.info(f"Записан тег в файл: {name}")
    return True


def cmd_file(kp: Kinopoisk, path: str, output: str, args: argparse.Namespace) -> None:
    if not os.path.isfile(path):
        log.error(f"Файл {path} не найден.")
        return
    titles = read_lines(path)
    if not titles:
        log.warning("Фильмы не найдены.")
        return
    log.info(f"Запрос из {path} ({len(titles)}): " + ", ".join(titles))
    make_list_from_titles(kp, titles, output, args, cover_default=file_title(path))


def cmd_tag(kp: Kinopoisk, path: str, args: argparse.Namespace) -> None:
    if os.path.isfile(path):
        if not is_mp4(path):
            log.error("Можно записывать теги только в файлы mp4.")
            return
        tag_file(kp, path, args.kinopoisk_id)
    elif os.path.isdir(path):
        files = list_mp4_dir(path)
        if args.test:
            _, not_found = resolve_titles(kp, [file_title(file) for file in files])
            if not_found:
                print("Следующие фильмы не найдены:")
                print("\n".join(not_found))
            return
        for file in files:
            tag_file(kp, file)
    else:
        log.error("Неверно указан путь.")


def cmd_cleartags(path: str) -> None:
    if os.path.isfile(path):
        if not is_mp4(path):
            log.error("Можно удалять теги только в файлах mp4.")
            return
        files = [path]
    elif os.path.isdir(path):
        files = list_mp4_dir(path)
    else:
        log.error("Неверно указан путь.")
        return
    for file in files:
        if clear_tags(file):
            log.info(f"Теги удалены в файле: {os.path.basename(file)}")
        else:
            log.warning(f"Теги не удалены в файле: {os.path.basename(file)}")


def cmd_list(kp: Kinopoisk, path: str, output: str, args: argparse.Namespace) -> None:
    if not os.path.isdir(path):
        log.error("Ошибка! В качестве параметра должен быть путь до каталога с файлами mp4.")
        return
    files = list_mp4_dir(path)
    if files:
        make_list_from_titles(kp, [file_title(file) for file in files], output, args, cover_default=dir_name(path))


def cmd_rename(kp: Kinopoisk, pattern: str) -> None:
    """Переименовывает файлы из торрент-имён в ``Название (год).ext`` после подтверждения."""
    import glob

    files = glob.glob(pattern)
    if not files:
        log.warning("Файлы не найдены.")
        return
    renames: list[Rename] = []
    for file in files:
        name = os.path.basename(file)
        log.info(f"Поиск названия фильма в имени файла: {name}")
        title = torrent_title(file)
        found = resolve_titles(kp, [title])[0] if title else []
        if not found:
            log.info(f"Не найдено название фильма в имени файла: {name}")
            continue
        renames.append(Rename(file, rename_destination(file, found[0].title, found[0].year)))
    if not renames:
        log.warning("Нечего переименовывать.")
        return

    print("\nБудут переименованы файлы:")
    for number, item in enumerate(renames, start=1):
        print(f"{number:2d}:", item.source, "->", item.destination)
    print()
    if input("Продолжить? [y/n] ").lower() == "y":
        apply_renames(renames)
        log.info("Файлы переименованы.")
    else:
        log.info("Отмена переименования файлов.")


def cmd_loc(path: str, output: str, args: argparse.Namespace) -> None:
    """Список по тегам mp4-файлов каталога без обращения к API."""
    if not os.path.isdir(path):
        log.error("Ошибка! В качестве параметра должен быть путь до каталога с файлами mp4.")
        return
    log.info(f"Поиск файлов mp4 в каталоге: {os.path.abspath(path)}")
    files = find_mp4_files(path, follow_lnk=True)
    if not files:
        log.warning(f'В каталоге "{path}" файлы mp4 не найдены.')
        return
    files, message = sort_files(files, args.sort)
    log.info(f"Сортировка файлов: {message}")
    for file in files:
        log.info(f"Найден файл: {os.path.basename(file)}")
    log.info(f"Всего: {len(files)}")

    films: list[Film] = []
    for file in tqdm(files, desc="Загрузка тегов...        "):
        film = read_tags(file)
        if film is None:
            log.warning(f"Не удалось прочитать теги в файле: '{os.path.basename(file)}'! Файл пропущен.")
        else:
            films.append(film)
    if not films:
        log.error("Ошибка, список не создан!")
        return
    save_lists(films, output, args, cover_default=dir_name(path))


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="[%(asctime)s]%(levelname)s:%(name)s:%(message)s",
        datefmt="%d.%m.%Y %H:%M:%S",
    )
    args = build_parser().parse_args(argv)

    requests_cache.install_cache(str(cache_path()), expire_after=CACHE_EXPIRE_SECONDS)
    if args.clearcache:
        requests_cache.clear()
        log.info("Кэш очищен.")
        return 0
    if args.nocache:
        requests_cache.uninstall_cache()

    output = resolve_output(args.output)
    if output is None:
        return 1

    needs_api = any((args.file, args.movie, args.tag, args.list, args.rename))
    kp: Kinopoisk | None = None
    if needs_api:
        token = load_token()
        if token is None:
            return 1
        kp = Kinopoisk(token)

    if args.file:
        cmd_file(kp, args.file, output, args)
    elif args.movie:
        make_list_from_titles(kp, args.movie, output, args, cover_default=file_title(output))
    elif args.tag:
        cmd_tag(kp, args.tag, args)
    elif args.cleartags:
        cmd_cleartags(args.cleartags)
    elif args.list:
        cmd_list(kp, args.list, output, args)
    elif args.rename:
        cmd_rename(kp, args.rename)
    elif args.loc:
        cmd_loc(args.loc, output, args)
    else:
        print(f"Kinolist Lib {__version__}\nДля помощи используйте параметр --help")
    return 0
