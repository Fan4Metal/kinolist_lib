"""Модуль argparse с русскими служебными сообщениями.

Подмена ``gettext.gettext`` должна произойти до первого импорта ``argparse``,
поэтому импортировать argparse следует только отсюда: ``from .argparse_ru import argparse``.
"""

import gettext

TRANSLATIONS = {
    "usage": "Применение",
    "show this help message and exit": "выводит это сообщение и завершает работу",
    "error:": "Ошибка:",
    "the following arguments are required:": "Следующие аргументы обязательны:",
    "options": "Параметры",
    "show program's version number and exit": "Показывает версию и завершает работу",
    "unrecognized arguments": "нераспознанные параметры",
    "examples:": "Примеры:",
}


def localize(text: str) -> str:
    for source, target in TRANSLATIONS.items():
        text = text.replace(source, target)
    return text


gettext.gettext = localize

import argparse  # noqa: E402

__all__ = ["argparse"]
