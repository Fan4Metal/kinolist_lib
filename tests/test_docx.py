from dataclasses import replace

from docx import Document
from docx.oxml.ns import qn

from kinolist.docx_out import txt_path_for, write_simple_list, write_table_list, write_txt_list
from kinolist.resources import template_path


def cover_paragraph(path):
    """Первый элемент тела документа; тест падает, если это не абзац обложки."""
    first = Document(str(path)).element.body[0]
    assert first.tag == qn("w:p")
    return first


def cover_texts(paragraph) -> set[str]:
    return {text.text for text in paragraph.xpath(".//w:t")}


def test_write_table_list_clones_tables(tmp_path, film):
    output = tmp_path / "list.docx"
    films = [film, replace(film, title="Терминатор 2", rating="", directors=[])]
    assert write_table_list(films, str(output), template_path(a5=False), genres=True)

    document = Document(str(output))
    assert len(document.tables) == 2
    first, second = document.tables
    assert first.cell(0, 1).text == "Терминатор - Кинопоиск 8.0"
    assert "Режиссер: Джеймс Кэмерон" in first.cell(1, 1).text
    assert "Жанр: фантастика" in first.cell(1, 1).text
    assert "В главных ролях: Арнольд Шварценеггер" in first.cell(1, 1).text
    assert second.cell(0, 1).text == "Терминатор 2 - нет рейтинга"
    assert len(document.inline_shapes) == 2


def test_write_table_list_a5(tmp_path, film):
    output = tmp_path / "a5.docx"
    assert write_table_list([film], str(output), template_path(a5=True))
    assert round(Document(str(output)).sections[0].page_height.cm, 1) == 14.8


def test_write_simple_list(tmp_path, film):
    output = tmp_path / "simple.docx"
    assert write_simple_list([film, film], str(output), genres=True)
    paragraphs = [p.text for p in Document(str(output)).paragraphs]
    assert paragraphs[0].startswith("1. Терминатор (1984) Жанр: фантастика\nРежиссер: Джеймс Кэмерон\nАктеры: ")
    assert paragraphs[0].count(",") == 2
    assert paragraphs[1].startswith("2. ")


def test_table_list_with_cover(tmp_path, film):
    output = tmp_path / "cover.docx"
    assert write_table_list([film, film], str(output), template_path(), cover="Рекомендации #231")
    cover = cover_paragraph(output)
    assert cover_texts(cover) == {"Рекомендации #231"}
    assert cover.xpath(".//w:br[@w:type='page']")
    assert len(Document(str(output)).tables) == 2


def test_cover_is_scaled_for_a5(tmp_path, film):
    output = tmp_path / "cover_a5.docx"
    assert write_table_list([film], str(output), template_path(a5=True), cover="A5")
    cover = cover_paragraph(output)
    sizes = {int(element.get(qn("w:val"))) for element in cover.xpath(".//w:sz")}
    assert max(sizes) == 72  # 72 пт в шаблоне A4 -> 36 пт (72 полупункта) на странице высотой 14,8 см
    page_height = Document(str(output)).sections[0].page_height
    assert all(int(element.get("cy")) < page_height * 1.1 for element in cover.xpath(".//wp:extent"))


def test_simple_list_with_cover(tmp_path, film):
    output = tmp_path / "simple_cover.docx"
    assert write_simple_list([film], str(output), cover="Обложка")
    assert cover_texts(cover_paragraph(output)) == {"Обложка"}


def test_write_txt_list(tmp_path, film):
    output = str(tmp_path / "list.docx")
    write_txt_list([film, replace(film, title="Другой")], txt_path_for(output))
    assert (tmp_path / "list.txt").read_text(encoding="utf-8") == "Терминатор\nДругой\n"
