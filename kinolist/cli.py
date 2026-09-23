"""Интерфейс командной строки ``kl``."""

from __future__ import annotations

import glob
import os
import sys
from pathlib import Path

import requests
import requests_cache

from . import __version__
from .argparse_ru import argparse
from .console import Console, install_logging, progress
from .docx_out import txt_path_for, write_simple_list, write_table_list, write_txt_list
from .files import (
    Rename,
    apply_renames,
    file_title,
    find_mp4_files,
    is_mp4,
    read_lines,
    rename_destination,
    safe_filename,
    sort_files,
    torrent_title,
)
from .kinopoisk import Kinopoisk, KinopoiskError, SearchResult
from .models import Film
from .resources import cache_path, template_path
from .tags import clear_tags, read_tags, write_tags

DEFAULT_OUTPUT = "list.docx"
CACHE_EXPIRE_SECONDS = 3600

console = Console()

EPILOG = R"""
Примеры

Списки по названиям (нужен доступ к API):
  kl -m "Terminator" "Terminator 2" KP~319   список list.docx из трех фильмов
  kl -f movies.txt -o movies.docx            список movies.docx из названий в файле movies.txt
  kl -f movies.txt --test                    только поиск фильмов, без создания списка
  kl -l c:\movies                            список по именам mp4-файлов в каталоге

Списки по тегам (без доступа к API, теги должны быть записаны заранее):
  kl --loc                                   список list.docx из тегов файлов текущего каталога
  kl --loc c:\movies --sort datem_r          для каталога c:\movies, новые файлы первыми
  kl --loc --cover                           с обложкой, заголовок из имени каталога
  kl --loc --cover "Рекомендации #231"       с обложкой с указанным заголовком
  kl --loc --cover --cover-name              файл списка называется как заголовок обложки

Оформление списка (для любого способа создания):
  -a5                                        шаблон A5 (для планшетов)
  -nf                                        простой нумерованный список без постеров
  -g                                         жанр фильма в карточке
  -s                                         сокращенные описания, два фильма на странице
  -tl                                        дополнительно текстовый файлс названиями

Теги mp4 (нужен доступ к API):
  kl -t                                      записать теги во все mp4-файлы текущего каталога
  kl -t c:\movies\Terminator.mp4             теги в один файл, фильм ищется по имени файла
  kl -t c:\movies\Chuzhie.mp4 -kp 406        записать теги фильма с указанным Kinopoisk id
  kl -t c:\movies --test                     только поиск фильмов по именам файлов
  kl --cleartags Alien.mp4                   удалить теги в файле после подтверждения
  kl --cleartags c:\movies --no-confirm      удалить теги во всех файлах без вопроса
  kl -r *.mp4                                переименовать файлы: торрент-имя -> Название (год).mp4

Прочее:
  --nocache, --clearcache                    не использовать кэш запросов, очистить кэш
  --pause                                    ждать Enter перед выходом (для контекстного меню)

Поиск фильма. Тег KP~XXX в названии задает Kinopoisk id напрямую. Для строки вида
"Название (Original Title) 2006" перебираются варианты: строка целиком, название без скобок,
содержимое каждой скобки; год помогает выбрать нужный результат. Рейтинг в теге kpra,
начинающийся с "i" (например i6.7), интерпретируется как рейтинг IMDb.
"""

SORT_HELP = (
    "порядок файлов для --loc: name (по имени, по умолчанию), date (по дате создания), "
    "datem (по дате изменения); суффикс _r дает обратный порядок, например datem_r"
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
    add("-f", "--file", metavar="FILE.txt", help="список фильмов по названиям из текстового файла, по одному в строке")
    add("-tl", "--txtlist", action="store_true", help="дополнительно сохраняет текстовый файл с названиями фильмов")
    add("-m", "--movie", nargs="+", help="список фильмов по указанным названиям")
    add(
        "--test",
        action="store_true",
        help="только поиск фильмов, без создания списка и записи тегов (с --file, --movie, --list и --tag)",
    )
    add("-o", "--output", metavar="FILE.docx", help="имя выходного файла (list.docx по умолчанию)")
    add("-s", "--shorten", action="store_true", help="сокращает описания, чтобы два фильма помещались на странице")
    add(
        "-t",
        "--tag",
        nargs="?",
        const=cwd,
        metavar="PATH",
        help="записывает теги в файл mp4 или во все mp4-файлы каталога (по умолчанию текущего)",
    )
    add("-kp", "--kinopoisk_id", type=int, metavar="ID", help="Kinopoisk id фильма для записи в тег (с --tag)")
    add(
        "--cleartags",
        nargs="?",
        const=cwd,
        metavar="PATH",
        help="удаляет все теги в файле mp4 или во всех mp4-файлах каталога (по умолчанию текущего)",
    )
    add("--no-confirm", action="store_true", help="удаляет теги без запроса подтверждения (с --cleartags)")
    add(
        "-r",
        "--rename",
        nargs="?",
        const=cwd,
        metavar="MASK",
        help="переименовывает файлы по маске из торрент-имен в «Название (год)», с подтверждением",
    )
    add(
        "-l",
        "--list",
        nargs="?",
        const=cwd,
        metavar="DIR",
        help="список фильмов по именам mp4-файлов каталога (по умолчанию текущего)",
    )
    add(
        "--loc",
        nargs="?",
        const=cwd,
        metavar="DIR",
        help="список фильмов по тегам mp4-файлов каталога (по умолчанию текущего), без обращения к API",
    )
    add("-nf", "--newformat", action="store_true", help="простой нумерованный список без постеров")
    add("-g", "--genres", action="store_true", help="добавляет жанр фильма в список")
    add("-a5", "--a5", action="store_true", help="список в формате A5 (для списков с постерами)")
    add(
        "--cover",
        nargs="?",
        const="",
        metavar="ТЕКСТ",
        help="добавляет обложку первой страницей; без текста заголовок берется из имени каталога или файла",
    )
    add(
        "--cover-name",
        action="store_true",
        help="называет выходной файл как заголовок обложки (без запрещенных символов), работает вместе с --cover",
    )
    add("--sort", metavar="ORDER", help=SORT_HELP)
    add("--nocache", action="store_true", help="не использовать кэш запросов к API")
    add("--clearcache", action="store_true", help="очищает кэш запросов к API и завершает работу")
    add("--pause", action="store_true", help="ждет нажатия Enter перед выходом (для запуска из контекстного меню)")
    return parser


def load_token() -> str | None:
    """Токен API из модуля ``config`` (файл config.py рядом с программой)."""
    try:
        from config import KINOPOISK_API_TOKEN
    except ImportError:
        console.error("не найден файл config.py с переменной KINOPOISK_API_TOKEN.")
        return None
    return KINOPOISK_API_TOKEN


def resolve_output(output: str | None) -> str | None:
    """Проверяет имя выходного файла и создаёт его каталог."""
    if not output:
        return DEFAULT_OUTPUT
    if os.path.splitext(output)[1].lower() != ".docx":
        console.error("выходной файл должен иметь расширение docx.")
        return None
    output_dir = os.path.dirname(output)
    if output_dir:
        Path(output_dir).mkdir(parents=True, exist_ok=True)
    return output


def found_text(result: SearchResult) -> str:
    return f"{result.title} ({result.year})" if result.year else result.title


def resolve_titles(kp: Kinopoisk, titles: list[str]) -> tuple[list[SearchResult], list[str]]:
    """Ищет фильмы по названиям. Возвращает найденные фильмы и список ненайденных названий."""
    found: list[SearchResult] = []
    not_found: list[str] = []
    for title in titles:
        try:
            result, query = kp.search_any(title)
        except (KinopoiskError, requests.RequestException) as error:
            console.warn(f"{title}: ошибка поиска ({error})")
            result, query = None, title
        if result is None:
            console.fail(title)
            not_found.append(title)
        else:
            note = f"KP {result.kp_id}"
            if query != title:
                note += f", по запросу «{query}»"
            console.ok(found_text(result), note)
            found.append(result)
    return found, not_found


def search_summary(found: list[SearchResult], not_found: list[str]) -> None:
    console.info(f"Найдено: {len(found)}, не найдено: {len(not_found)}")
    if not_found:
        # Отдельный список ненайденных названий: так их удобнее искать вручную.
        console.section(f"Не найдены ({len(not_found)})")
        for number, title in enumerate(not_found, start=1):
            console.item(number, title)


def load_films(kp: Kinopoisk, kp_ids: list[int], shorten: bool = False) -> list[Film]:
    films: list[Film] = []
    for kp_id in progress(kp_ids, "Загрузка информации"):
        try:
            films.append(kp.film(kp_id, shorten))
        except (KinopoiskError, requests.RequestException, KeyError, ValueError) as error:
            console.warn(f"не удалось загрузить фильм {kp_id}: {error}")
    return films


def cover_title(args: argparse.Namespace, default: str) -> str | None:
    """Заголовок обложки: текст параметра ``--cover``, либо ``default``, если параметр указан без текста."""
    if args.cover is None:
        return None
    return args.cover or default


def dir_name(path: str) -> str:
    return os.path.basename(os.path.abspath(path))


def output_for_cover(output: str, cover: str) -> str:
    """Имя файла по заголовку обложки (без запрещённых символов) в каталоге ``output``."""
    name = safe_filename(cover).strip(" .")
    return os.path.join(os.path.dirname(output), name + ".docx") if name else output


def save_lists(films: list[Film], output: str, args: argparse.Namespace, cover_default: str = "") -> None:
    cover = cover_title(args, cover_default)
    console.info(f"Фильмов в списке: {len(films)}")
    if cover:
        console.info(f"Обложка: {cover}")
        if args.cover_name:
            output = output_for_cover(output, cover)
    if args.newformat:
        saved = write_simple_list(films, output, genres=args.genres, cover=cover)
    else:
        saved = write_table_list(films, output, template_path(a5=args.a5), genres=args.genres, cover=cover)
    if not saved:
        console.error(f'нет доступа на запись к файлу "{output}". Список не сохранен.')
        return
    if args.txtlist:
        write_txt_list(films, txt_path_for(output))
        console.info(f"Текстовый список: {txt_path_for(output)}")
    console.result(f"Список создан: {output}")


def make_list_from_titles(
    kp: Kinopoisk, titles: list[str], output: str, args: argparse.Namespace, cover_default: str = ""
) -> None:
    console.section(f"Поиск фильмов ({len(titles)})")
    found, not_found = resolve_titles(kp, titles)
    search_summary(found, not_found)
    if args.test or not found:
        return
    console.section("Создание списка")
    films = load_films(kp, [item.kp_id for item in found], args.shorten)
    if not films:
        console.error("список не создан.")
        return
    save_lists(films, output, args, cover_default)


def list_mp4_dir(path: str, follow_lnk: bool = False, sort: str | None = None) -> list[str]:
    console.section(f"Каталог: {os.path.abspath(path)}")
    files = find_mp4_files(path, follow_lnk)
    if not files:
        console.warn("файлы mp4 не найдены.")
        return []
    files, message = sort_files(files, sort)
    if sort:
        console.note(f"Сортировка: {message}")
    for number, file in enumerate(files, start=1):
        console.item(number, os.path.basename(file))
    console.info(f"Всего файлов: {len(files)}")
    return files


def tag_file(kp: Kinopoisk, path: str, kp_id: int | None = None) -> bool:
    """Записывает теги в один файл. Фильм ищется по имени файла, если id не задан."""
    name = os.path.basename(path)
    if kp_id is None:
        found, _ = resolve_titles(kp, [file_title(path)])
        if not found:
            return False
        kp_id = found[0].kp_id
    try:
        film = kp.film(kp_id)
    except (KinopoiskError, requests.RequestException, KeyError, ValueError) as error:
        console.warn(f"не удалось загрузить фильм {kp_id}: {error}")
        return False
    if not write_tags(film, path):
        return False
    console.ok(f"{name} {console.arrow()} {film.title} ({film.year})", "теги записаны")
    return True


def cmd_file(kp: Kinopoisk, path: str, output: str, args: argparse.Namespace) -> None:
    if not os.path.isfile(path):
        console.error(f"файл {path} не найден.")
        return
    titles = read_lines(path)
    if not titles:
        console.warn("в файле нет названий фильмов.")
        return
    make_list_from_titles(kp, titles, output, args, cover_default=file_title(path))


def cmd_tag(kp: Kinopoisk, path: str, args: argparse.Namespace) -> None:
    if os.path.isfile(path):
        if not is_mp4(path):
            console.error("можно записывать теги только в файлы mp4.")
            return
        console.section(f"Запись тегов: {os.path.basename(path)}")
        tag_file(kp, path, args.kinopoisk_id)
    elif os.path.isdir(path):
        files = list_mp4_dir(path)
        if not files:
            return
        if args.test:
            console.section(f"Поиск фильмов ({len(files)})")
            found, not_found = resolve_titles(kp, [file_title(file) for file in files])
            search_summary(found, not_found)
            return
        console.section("Запись тегов")
        written = sum(tag_file(kp, file) for file in files)
        console.result(f"Теги записаны: {written} из {len(files)}")
    else:
        console.error("неверно указан путь.")


def cmd_cleartags(path: str, confirm: bool = True) -> None:
    if os.path.isfile(path):
        if not is_mp4(path):
            console.error("можно удалять теги только в файлах mp4.")
            return
        files = [path]
    elif os.path.isdir(path):
        files = list_mp4_dir(path)
        if not files:
            return
    else:
        console.error("неверно указан путь.")
        return
    console.section("Удаление тегов")
    if confirm:
        question = (
            f"Удалить теги в файле {os.path.basename(path)}? [y/n] "
            if len(files) == 1
            else f"Удалить теги во всех файлах ({len(files)})? [y/n] "
        )
        if console.prompt(question).lower() != "y":
            console.info("Удаление отменено.")
            return
    cleared = 0
    for file in files:
        if clear_tags(file):
            console.ok(os.path.basename(file))
            cleared += 1
    console.result(f"Теги удалены: {cleared} из {len(files)}")


def cmd_list(kp: Kinopoisk, path: str, output: str, args: argparse.Namespace) -> None:
    if not os.path.isdir(path):
        console.error("в качестве параметра должен быть путь до каталога с файлами mp4.")
        return
    files = list_mp4_dir(path)
    if files:
        make_list_from_titles(kp, [file_title(file) for file in files], output, args, cover_default=dir_name(path))


def cmd_rename(kp: Kinopoisk, pattern: str) -> None:
    """Переименовывает файлы из торрент-имён в ``Название (год).ext`` после подтверждения."""
    files = glob.glob(pattern)
    if not files:
        console.warn("файлы не найдены.")
        return
    console.section(f"Поиск названий ({len(files)})")
    renames: list[Rename] = []
    for file in files:
        name = os.path.basename(file)
        title = torrent_title(file)
        found = resolve_titles(kp, [title])[0] if title else []
        if not found:
            console.fail(f"{name}: название не определено")
            continue
        renames.append(Rename(file, rename_destination(file, found[0].title, found[0].year)))
    if not renames:
        console.warn("нечего переименовывать.")
        return

    console.section("Будут переименованы файлы")
    for number, item in enumerate(renames, start=1):
        console.item(number, f"{os.path.basename(item.source)} {console.arrow()} {os.path.basename(item.destination)}")
    console.write()
    if console.prompt("Продолжить? [y/n] ").lower() == "y":
        apply_renames(renames)
        console.result(f"Переименовано файлов: {len(renames)}")
    else:
        console.info("Переименование отменено.")


def cmd_loc(path: str, output: str, args: argparse.Namespace) -> None:
    """Список по тегам mp4-файлов каталога без обращения к API."""
    if not os.path.isdir(path):
        console.error("в качестве параметра должен быть путь до каталога с файлами mp4.")
        return
    files = list_mp4_dir(path, follow_lnk=True, sort=args.sort)
    if not files:
        return

    console.section("Чтение тегов")
    films: list[Film] = []
    for file in files:
        film = read_tags(file)
        if film is None:
            console.fail(f"{os.path.basename(file)}: нет тегов Kinolist, файл пропущен")
        else:
            console.ok(f"{film.title} ({film.year})" if film.year else film.title)
            films.append(film)
    if not films:
        console.error("список не создан.")
        return
    console.section("Создание списка")
    save_lists(films, output, args, cover_default=dir_name(path))


def run(args: argparse.Namespace) -> int:
    requests_cache.install_cache(str(cache_path()), expire_after=CACHE_EXPIRE_SECONDS)
    if args.clearcache:
        requests_cache.clear()
        console.result("Кэш очищен.")
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
        cmd_cleartags(args.cleartags, confirm=not args.no_confirm)
    elif args.list:
        cmd_list(kp, args.list, output, args)
    elif args.rename:
        cmd_rename(kp, args.rename)
    elif args.loc:
        cmd_loc(args.loc, output, args)
    else:
        console.info("Для помощи используйте параметр --help")
    return 0


def main(argv: list[str] | None = None) -> int:
    # При перенаправлении вывода кодировка может не содержать части символов: заменяем их, а не падаем.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")
    install_logging(console)
    args = build_parser().parse_args(argv)
    console.header(f"Kinolist Lib {__version__}")
    try:
        return run(args)
    except KeyboardInterrupt:
        console.write()
        console.warn("прервано.")
        return 130
    finally:
        if args.pause:
            console.write()
            console.prompt("Нажмите Enter для выхода...")
