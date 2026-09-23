"""Формирование списков фильмов: docx по шаблону, docx в простом формате, txt."""

from __future__ import annotations

import logging
import os
import re
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.table import Table

from .console import progress
from .files import write_lines
from .models import Film, image_to_png
from .resources import cover_template_path, no_poster

log = logging.getLogger(__name__)

FONT = "Arial"
TITLE_SIZE = Pt(11)
TEXT_SIZE = Pt(10)
ACTORS_LABEL_COLOR = RGBColor(255, 102, 0)
ACTORS_COLOR = RGBColor(0, 0, 255)
POSTER_WIDTH = Cm(7)
SIMPLE_FONT = "Times New Roman"
SIMPLE_FONT_SIZE = Pt(14)
SIMPLE_ACTORS_LIMIT = 3
# Размер страницы, под который свёрстан шаблон обложки; для других страниц объекты масштабируются.
COVER_PAGE_WIDTH = Cm(21.0)
COVER_PAGE_HEIGHT = Cm(29.7)
VML_SIZE_RE = re.compile(r"(margin-left|margin-top|width|height):(-?[\d.]+)pt")


def _scale_vml_style(style: str, sx: float, sy: float) -> str:
    """Масштабирует размеры в пунктах внутри атрибута ``style`` резервного представления VML."""

    def scale(match: re.Match) -> str:
        name, value = match.group(1), float(match.group(2))
        factor = sx if name in ("margin-left", "width") else sy
        return f"{name}:{value * factor:g}pt"

    return VML_SIZE_RE.sub(scale, style)


def _scale_cover(paragraph, sx: float, sy: float) -> None:
    """Масштабирует плавающие объекты обложки: размеры, смещения и кегль текста."""
    for element in paragraph.xpath(".//wp:extent | .//a:ext"):
        element.set("cx", str(round(int(element.get("cx")) * sx)))
        element.set("cy", str(round(int(element.get("cy")) * sy)))
    for element in paragraph.xpath(".//a:off"):
        element.set("x", str(round(int(element.get("x")) * sx)))
        element.set("y", str(round(int(element.get("y")) * sy)))
    for element in paragraph.xpath(".//wp:positionH/wp:posOffset"):
        element.text = str(round(int(element.text) * sx))
    for element in paragraph.xpath(".//wp:positionV/wp:posOffset"):
        element.text = str(round(int(element.text) * sy))
    font_scale = min(sx, sy)
    for element in paragraph.xpath(".//w:sz | .//w:szCs"):
        element.set(qn("w:val"), str(round(int(element.get(qn("w:val"))) * font_scale)))
    for element in paragraph.iter():
        style = element.get("style")
        if style:
            element.set("style", _scale_vml_style(style, sx, sy))


def add_cover(document, title: str) -> None:
    """Вставляет обложку с заголовком первой страницей документа.

    Обложка берётся из шаблона ``templates/template_cover.docx``: абзац с плавающим текстовым полем,
    подложкой на всю страницу и разрывом страницы в конце. Внешних связей у абзаца нет,
    поэтому его можно переносить между документами простым копированием.
    """
    cover = Document(str(cover_template_path()))
    paragraph = deepcopy(cover.paragraphs[0]._p)
    for text in paragraph.xpath(".//w:t"):
        text.text = title
    section = document.sections[0]
    sx = section.page_width / COVER_PAGE_WIDTH
    sy = section.page_height / COVER_PAGE_HEIGHT
    if abs(sx - 1) > 0.01 or abs(sy - 1) > 0.01:
        _scale_cover(paragraph, sx, sy)
    document.element.body.insert(0, paragraph)


def _clone_first_table(document, count: int) -> None:
    """Клонирует первую таблицу документа ``count`` раз, разделяя копии пустыми абзацами."""
    template = document.tables[0]._tbl
    paragraph = document.paragraphs[0]
    for _ in range(count):
        paragraph._p.addnext(deepcopy(template))
        paragraph = document.add_paragraph()


def _add_run(paragraph, text: str, size=TEXT_SIZE, bold=False, color: RGBColor | None = None, underline=False):
    run = paragraph.add_run(text)
    run.font.name = FONT
    run.font.size = size
    run.font.bold = bold
    run.font.underline = underline
    if color is not None:
        run.font.color.rgb = color
    return run


def _fill_table(table: Table, film: Film, genres: bool) -> None:
    """Заполняет одну таблицу шаблона данными фильма."""
    _add_run(table.cell(0, 1).paragraphs[0], f"{film.title} - {film.rating_text}", size=TITLE_SIZE, bold=True)

    info = table.cell(1, 1)
    _add_run(info.add_paragraph(), str(film.year) if film.year else "")
    _add_run(info.add_paragraph(), ", ".join(film.countries))
    _add_run(info.add_paragraph(), film.directors_text)
    if genres and film.main_genre:
        _add_run(info.add_paragraph(), f"Жанр: {film.main_genre}")

    info.add_paragraph()
    paragraph = info.add_paragraph()
    _add_run(paragraph, "В главных ролях: ", color=ACTORS_LABEL_COLOR)
    _add_run(paragraph, ", ".join(film.actors), color=ACTORS_COLOR, underline=True)

    info.add_paragraph()
    info.add_paragraph()
    _add_run(info.add_paragraph(), film.description)
    info.add_paragraph()

    poster = film.poster or no_poster()
    table.cell(0, 0).paragraphs[1].add_run().add_picture(image_to_png(poster), width=POSTER_WIDTH)


def _save(document, path: str) -> bool:
    """Сохраняет документ; ``False``, если файл открыт в другой программе или недоступен для записи."""
    try:
        document.save(path)
    except PermissionError:
        return False
    return True


def write_table_list(
    films: list[Film], path: str, template: Path, genres: bool = False, cover: str | None = None
) -> bool:
    """Список с постерами по шаблону docx (одна таблица на фильм); ``cover`` добавляет обложку."""
    document = Document(str(template))
    if len(films) > 1:
        _clone_first_table(document, len(films) - 1)
    for table, film in zip(progress(document.tables, "Запись в таблицы"), films, strict=True):
        _fill_table(table, film, genres)
    if cover:
        add_cover(document, cover)
    return _save(document, path)


def write_simple_list(films: list[Film], path: str, genres: bool = False, cover: str | None = None) -> bool:
    """Нумерованный текстовый список без постеров («новый формат»); ``cover`` добавляет обложку."""
    document = Document()
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.75)
    section.bottom_margin = Cm(2)

    font = document.styles["Normal"].font
    font.name = SIMPLE_FONT
    font.size = SIMPLE_FONT_SIZE

    for number, film in enumerate(progress(films, "Запись в файл"), start=1):
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(12)
        paragraph.add_run(f"{number}. ")
        paragraph.add_run(f"{film.title} ({film.year}) ").bold = True
        if genres and film.main_genre:
            paragraph.add_run(f"Жанр: {film.main_genre}\n")
        paragraph.add_run(f"{film.directors_text}\n")
        paragraph.add_run(f"Актеры: {', '.join(film.actors[:SIMPLE_ACTORS_LIMIT])}")
    if cover:
        add_cover(document, cover)
    return _save(document, path)


def write_txt_list(films: list[Film], path: str) -> None:
    """Текстовый файл с названиями фильмов, по одному в строке."""
    write_lines(path, [film.title for film in films])


def txt_path_for(docx_path: str) -> str:
    return os.path.splitext(docx_path)[0] + ".txt"
