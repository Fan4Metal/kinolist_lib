# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Что это

Kinolist Lib — консольная утилита `kl` для Windows: строит списки фильмов в docx и пишет карточки фильмов в теги mp4 по данным неофициального API Кинопоиска. Код, docstrings, комментарии, справка и сообщения программы — на русском. README двуязычный: `README.md` (русский, основной) и `README.en.md`; изменения в одном должны отражаться в другом.

## Команды

Окружение управляется через `uv` (Python 3.14 в `.python-version`, минимум 3.13).

```
uv sync                                   # установка зависимостей и dev-группы
uv run kl --help                          # запуск CLI из исходников
uv run pytest                             # все тесты
uv run pytest tests/test_tags.py          # один файл
uv run pytest tests/test_tags.py::test_parse_year   # один тест
uv run ruff check .                       # линтер (E, F, I, UP, B, W; строка 120)
uv run ruff format .
uv run python tools/make_release.py [--no-installer]   # PyInstaller -> dist/kl, затем Inno Setup -> dist/*.exe
```

Тесты тегов используют фикстуру `mp4_file`, которая создаёт файл через `ffmpeg`; без `ffmpeg` в `PATH` они пропускаются, а не падают. Сетевые запросы в тестах не выполняются: `Kinopoisk` подменяется через переопределение `_get`/`_poster` (tests/test_kinopoisk.py) или monkeypatch `cli.Kinopoisk` и `cli.load_token` (tests/test_cli.py).

Токен API читается из `config.py` в корне (переменная `KINOPOISK_API_TOKEN`). Файл в `.gitignore`, как и `cache.sqlite`.

## Архитектура

Точка входа `kinolist_lib.py` (нужна PyInstaller) вызывает `kinolist.cli:main`; тот же вызов зарегистрирован как скрипт `kl` в pyproject.

Поток данных: `cli.run` → `Kinopoisk.search_any` (варианты названия из `parse_title`, тег `KP~id`) → `Kinopoisk.film` → `Film` → `docx_out.*` или `tags.write_tags`. Ветка `--loc` идёт в обход API: `tags.read_tags` → `Film` → `docx_out`.

- **cli.py** — argparse и диспетчер `run()`: за один запуск выполняется одна команда (цепочка `elif` по приоритету `-f`, `-m`, `-t`, `--cleartags`, `-l`, `-r`, `--loc`). Токен и клиент создаются только если команда попала в `needs_api`; `--loc` и `--cleartags` работают без сети. `requests_cache` устанавливается глобально здесь же (файл `cache.sqlite` рядом с ресурсами, TTL 1 час).
- **argparse_ru.py** — подменяет `gettext` до импорта `argparse`, чтобы служебные сообщения были русскими. Импортировать argparse нужно только так: `from .argparse_ru import argparse`.
- **kinopoisk.py** — тонкая обёртка над тремя методами API (`search-by-keyword`, `films/{id}`, `staff`). Год из названия — предпочтение при выборе результата (допуск ±1), а не фильтр. Пауза между запросами делается только для ответов не из кэша. Постер приводится к 2:3 и 360×540.
- **models.py** — `Film`. Рейтинг хранится строкой: `"7.5"` Кинопоиск, `"i6.7"` IMDb, `""` нет рейтинга. `GENRES_HIERARCHY` задаёт приоритет основного жанра.
- **tags.py** — раскладка атомов mp4 (стандартные плюс `----:com.apple.iTunes:*`). Формат должен оставаться совместимым с файлами, протегированными старыми версиями (тест `test_legacy_tag_layout`). «Своим» считается файл, у которого есть хотя бы один атом из `OWN_TAGS`; актёры хранятся парами «пустая строка, имя» через `\r\n`.
- **docx_out.py** — три формата: таблицы по шаблону (первая таблица шаблона клонируется на каждый фильм), простой нумерованный список, txt. Обложка — абзац из `templates/template_cover.docx`, копируемый в начало документа и масштабируемый под размер страницы (A5).
- **resources.py** — единственное место, знающее о `sys._MEIPASS`: из сборки ресурсы берутся оттуда, из исходников — из корня проекта. Каталоги `templates/` и `images/` попадают в сборку через `DATA_DIRS` в tools/make_release.py; новый ресурс нужно класть в один из них.
- **console.py** — весь вывод. `cli` пишет через объект `Console`; остальные модули пишут в `logging.getLogger(__name__)`, и `install_logging` направляет логгер `kinolist` в ту же консоль. Цвета и маркеры деградируют до ASCII при перенаправлении вывода.
- **files.py** — поиск mp4, ярлыки `.lnk` через pywin32 (только для `--loc`), разбор торрент-имён через `parse-torrent-title`, варианты `--sort`.

## Ограничения, которые нельзя нарушать без запроса

- Флаги CLI, на которые завязан установщик (`kinolist_lib.iss`, контекстное меню Проводника): `--genres --loc "%V" --pause`, `--cover --cover-name`, `-t "%1"`, `--cleartags "%1"`, `-f "%1"`. Их имена и семантику менять нельзя.
- Совместимость раскладки тегов mp4 и структуры шаблонов docx с предыдущими версиями.
- Версия задаётся только в `kinolist/__init__.py`; `make_release.py` передаёт её в ресурс версии exe и в Inno Setup (`/DMyAppVersion`). Значение в `.iss` — запасное для ручной сборки.
- Проект рассчитан на Windows: `pywin32` и `colorama` подключены условно по `sys_platform`.
