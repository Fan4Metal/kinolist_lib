"""Форматированный вывод в консоль.

Цвета включаются только для терминала (и отключаются переменной ``NO_COLOR``). В классической
консоли Windows (cmd.exe) поддержку управляющих последовательностей включает colorama.
Маркеры подобраны из символов, которые есть в стандартных шрифтах консоли (Consolas, Lucida Console).
"""

from __future__ import annotations

import logging
import os
import sys
from collections.abc import Iterable
from typing import TextIO

from tqdm import tqdm

OK = "√"
FAIL = "×"
WARN = "!"
ARROW = "→"
# Запасные маркеры для потоков, кодировка которых не содержит символов выше (например, cp1251 при
# перенаправлении вывода в файл из cmd.exe).
ASCII_SYMBOLS = {OK: "+", FAIL: "-", ARROW: "->"}

RESET = "\x1b[0m"
BOLD = "\x1b[1m"
GRAY = "\x1b[90m"
RED = "\x1b[31m"
GREEN = "\x1b[32m"
YELLOW = "\x1b[33m"
CYAN = "\x1b[36m"

INDENT = "  "
PROGRESS_FORMAT = INDENT + "{desc}  {bar}  {n}/{total} ({percentage:3.0f}%)  {elapsed} < {remaining}"
PROGRESS_WIDTH = 72
# Дорожка из светлых блоков, заливка сплошная: оба символа есть в шрифтах консоли и в cp866.
PROGRESS_CHARS = "░█"


def progress(items: Iterable, desc: str) -> tqdm:
    """Индикатор выполнения в едином стиле; после завершения строка убирается."""
    return tqdm(
        items,
        desc=desc,
        bar_format=PROGRESS_FORMAT,
        ncols=PROGRESS_WIDTH,
        ascii=PROGRESS_CHARS,
        colour="white" if supports_color(sys.stderr) else None,
        leave=False,
    )


def supports_color(stream: TextIO) -> bool:
    if os.environ.get("NO_COLOR"):
        return False
    return hasattr(stream, "isatty") and stream.isatty()


class Console:
    """Без явного ``stream`` вывод идёт в текущий ``sys.stdout``, который определяется при каждой записи."""

    def __init__(self, stream: TextIO | None = None, color: bool | None = None):
        self._stream = stream
        self._color = color
        self._terminal_ready = False

    @property
    def stream(self) -> TextIO:
        return self._stream or sys.stdout

    def symbol(self, symbol: str) -> str:
        """Маркер либо его ASCII-замена, если кодировка потока его не поддерживает."""
        encoding = getattr(self.stream, "encoding", None) or "utf-8"
        try:
            symbol.encode(encoding)
        except (UnicodeEncodeError, LookupError):
            return ASCII_SYMBOLS.get(symbol, symbol)
        return symbol

    @property
    def color(self) -> bool:
        if self._color is None:
            self._color = supports_color(self.stream)
        if self._color and not self._terminal_ready:
            self._terminal_ready = True
            if sys.platform == "win32":
                # Включает обработку управляющих последовательностей в классической консоли Windows.
                import colorama

                colorama.just_fix_windows_console()
        return self._color

    def style(self, text: str, *codes: str) -> str:
        if not self.color or not codes:
            return text
        return "".join(codes) + text + RESET

    def write(self, text: str = "") -> None:
        # tqdm.write не ломает активный индикатор выполнения; без индикатора это обычный print.
        tqdm.write(text, file=self.stream)

    def header(self, text: str) -> None:
        self.write(self.style(text, BOLD))

    def section(self, text: str) -> None:
        """Заголовок раздела, отделённый пустой строкой."""
        self.write()
        self.write(self.style(text, BOLD, CYAN))

    def info(self, text: str) -> None:
        self.write(INDENT + text)

    def note(self, text: str) -> None:
        """Второстепенная строка серым цветом."""
        self.write(INDENT + self.style(text, GRAY))

    def item(self, index: int, text: str) -> None:
        self.write(f"{index:>4}. {text}")

    def ok(self, text: str, note: str = "") -> None:
        line = f"{INDENT}{self.style(self.symbol(OK), GREEN)} {text}"
        if note:
            line += f"  {self.style(note, GRAY)}"
        self.write(line)

    def fail(self, text: str) -> None:
        self.write(f"{INDENT}{self.style(self.symbol(FAIL), RED)} {text}")

    def arrow(self) -> str:
        return self.symbol(ARROW)

    def warn(self, text: str) -> None:
        self.write(f"{INDENT}{self.style(WARN + ' ' + text, YELLOW)}")

    def error(self, text: str) -> None:
        self.write(self.style(f"Ошибка: {text}", RED, BOLD))

    def result(self, text: str) -> None:
        """Итог работы команды."""
        self.write()
        self.write(self.style(f"{self.symbol(OK)} {text}", GREEN, BOLD))

    def prompt(self, text: str) -> str:
        """Запрос ввода; при закрытом стандартном вводе возвращает пустую строку."""
        try:
            return input(self.style(text, BOLD))
        except EOFError:
            return ""


class ConsoleHandler(logging.Handler):
    """Выводит записи журнала библиотечных модулей в том же стиле, что и остальной вывод."""

    def __init__(self, console: Console):
        super().__init__()
        self.console = console

    def emit(self, record: logging.LogRecord) -> None:
        message = record.getMessage()
        if record.levelno >= logging.ERROR:
            self.console.fail(message)
        elif record.levelno >= logging.WARNING:
            self.console.warn(message)
        else:
            self.console.info(message)


def install_logging(console: Console, level: int = logging.INFO, name: str = "kinolist") -> None:
    """Направляет журнал пакета в консоль. Прежние обработчики этого типа заменяются.

    Обработчик вешается на логгер пакета, а не на корневой, чтобы сообщения сторонних библиотек
    (например, requests_cache) не попадали в вывод.
    """
    logger = logging.getLogger(name)
    for handler in list(logger.handlers):
        if isinstance(handler, ConsoleHandler):
            logger.removeHandler(handler)
    logger.addHandler(ConsoleHandler(console))
    logger.setLevel(level)
