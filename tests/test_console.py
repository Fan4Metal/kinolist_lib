import io
import logging

from kinolist.console import Console, install_logging


def make_console(color: bool):
    stream = io.StringIO()
    return Console(stream, color=color), stream


def test_plain_output():
    console, stream = make_console(color=False)
    console.header("Kinolist Lib")
    console.section("Раздел")
    console.item(1, "файл.mp4")
    console.ok("Фильм (1984)", "KP 507")
    console.fail("Не найден")
    console.warn("предупреждение")
    console.error("сбой")
    console.result("Готово")
    assert stream.getvalue() == (
        "Kinolist Lib\n\nРаздел\n   1. файл.mp4\n  √ Фильм (1984)  KP 507\n  × Не найден\n  ! предупреждение\n"
        "Ошибка: сбой\n\n√ Готово\n"
    )


def test_colored_output_wraps_in_escape_codes():
    console, stream = make_console(color=True)
    console.ok("Фильм")
    console.error("сбой")
    out = stream.getvalue()
    assert "\x1b[32m√\x1b[0m Фильм" in out
    assert "\x1b[31m\x1b[1mОшибка: сбой\x1b[0m" in out


def test_ascii_fallback_for_limited_encoding():
    raw = io.BytesIO()
    stream = io.TextIOWrapper(raw, encoding="cp1251", newline="\n")
    console = Console(stream, color=False)
    console.ok("Фильм")
    console.fail("Нет")
    console.result(f"a {console.arrow()} b")
    stream.flush()
    assert raw.getvalue().decode("cp1251") == "  + Фильм\n  - Нет\n\n+ a -> b\n"


def test_logging_handler_maps_levels():
    console, stream = make_console(color=False)
    install_logging(console)
    log = logging.getLogger("kinolist.test")
    log.info("сведения")
    log.warning("осторожно")
    log.error("плохо")
    logging.getLogger("requests_cache").info("Clearing all items from the cache")
    assert stream.getvalue() == "  сведения\n  ! осторожно\n  × плохо\n"
    install_logging(console)
    assert sum(type(h).__name__ == "ConsoleHandler" for h in logging.getLogger("kinolist").handlers) == 1
