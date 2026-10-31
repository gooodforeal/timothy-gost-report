"""Builds a GOST formatted .docx report from a plain folder.

Purpose:
    The agent writes report content as Markdown files in a plain folder; this script
    is the only place that knows how the content must look on paper: A4 page,
    margins 30/15/20/20 mm, Times New Roman 14 pt, 1.5 spacing, 1.25 cm indent,
    justified text, bold headings aligned left from the paragraph indent,
    "Рисунок N — ..." / "Таблица N — ..." / "Листинг N — ..." captions, Consolas
    listings, formulas numbered at the right edge, generated table of contents and
    page numbers centered in the footer. There is no title page: it is made
    separately. Re-running the script after editing the plain folder regenerates
    the whole document.

Usage:
    python3 build_report.py path/to/plain -o path/to/out/report.docx

    Without -o the report goes to out/report.docx next to the plain folder.

plain folder layout:
    style.json    optional overrides of the GOST style defaults
    *.md          report parts, assembled in file-name order (01_..., 02_...)
    images/       pictures referenced from the Markdown files

Requires:
    python-docx (pip install python-docx)
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
from typing import Sequence

from docx import Document
from docx.document import Document as DocxDocument
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Mm, Pt, RGBColor
from docx.styles.style import ParagraphStyle
from docx.table import Table as DocxTable
from docx.table import _Row
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.text.run import Run

from plain_parser import (
    BIBLIOGRAPHY_TITLES,
    TABLE_OF_CONTENTS_TITLE,
    Block,
    BulletList,
    Figure,
    Formula,
    Heading,
    Listing,
    PageBreak,
    Paragraph,
    PlainError,
    Table,
    number_blocks,
    parse_inline,
    parse_plain_file,
    resolve_references,
)
from report_style import ReportStyle, StyleError, load_report_style

logger = logging.getLogger(__name__)

STYLE_FILE_NAME = "style.json"
OUTPUT_DIR_NAME = "out"  # finished .docx files live here, next to plain/
DEFAULT_OUTPUT_NAME = "report.docx"
PAGE_WIDTH_MM = 210
PAGE_HEIGHT_MM = 297
CAPTION_DASH = "—"  # GOST: "Рисунок 1 — Название" uses an em dash with spaces
LIST_BULLET = "–"  # GOST: list items start with a dash
CYRILLIC_LIST_LETTERS = "абвгдежиклмнпрстуфхцшщэюя"
BLOCK_GAP_PT = 12  # one empty line between figures/tables/listings and text
TOC_INSTRUCTION = 'TOC \\o "1-3" \\h \\z \\u'
TOC_PLACEHOLDER = "Обновите поле: правый клик → «Обновить поле» (или F9)"
TABLE_FULL_WIDTH_PCT = "5000"  # OOXML percentages are in fiftieths of a percent
TAB_SIZE_SPACES = 4


class ReportRenderer:
    """Renders parsed plain blocks into a python-docx Document with GOST formatting.

    Keeps the small amount of state a document flow needs: whether the next
    paragraph must start a new page and whether we are inside the bibliography
    (its numbered items use "1." instead of "1)").
    """

    def __init__(self, style: ReportStyle) -> None:
        """Creates an empty A4 document with GOST page setup and styles.

        Args:
            style: Typography and page settings.
        """
        self._style = style
        self._document: DocxDocument = Document()
        self._is_page_break_pending = False
        self._is_in_bibliography = False
        self._text_width = Mm(PAGE_WIDTH_MM - style.margin_left_mm - style.margin_right_mm)
        self._setup_page()
        self._setup_styles()

    def render(self, blocks: Sequence[Block]) -> DocxDocument:
        """Builds the whole document: table of contents and body.

        Args:
            blocks: Numbered blocks with resolved references.

        Returns:
            The finished document, ready to save.
        """
        if self._style.has_table_of_contents:
            self._add_table_of_contents()
        for block in blocks:
            self._add_block(block)
        return self._document

    # ---------------------------------------------------------------- setup

    def _setup_page(self) -> None:
        """Sets A4, GOST margins and a centered page number on every page.

        Numbering may start from 2 when a separate title page goes in front.
        """
        section = self._document.sections[0]
        section.page_width = Mm(PAGE_WIDTH_MM)
        section.page_height = Mm(PAGE_HEIGHT_MM)
        section.left_margin = Mm(self._style.margin_left_mm)
        section.right_margin = Mm(self._style.margin_right_mm)
        section.top_margin = Mm(self._style.margin_top_mm)
        section.bottom_margin = Mm(self._style.margin_bottom_mm)
        if self._style.first_page_number != 1:
            page_numbering = OxmlElement("w:pgNumType")
            page_numbering.set(qn("w:start"), str(self._style.first_page_number))
            section._sectPr.append(page_numbering)

        footer_paragraph = section.footer.paragraphs[0]
        self._format_block_paragraph(footer_paragraph, WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
        _add_field(footer_paragraph, "PAGE", "1")

    def _setup_styles(self) -> None:
        """Configures Normal, Heading 1-3, Listing and TOC styles."""
        style = self._style
        styles = self._document.styles

        normal = styles["Normal"]
        _set_style_font(normal, style.font_name, style.font_size_pt)
        normal_format = normal.paragraph_format
        normal_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        normal_format.first_line_indent = Cm(style.first_line_indent_cm)
        normal_format.line_spacing = style.line_spacing
        normal_format.space_before = Pt(0)
        normal_format.space_after = Pt(0)
        normal_format.widow_control = True
        _set_style_language(normal, "ru-RU")

        # Headings: bold, body-size font, left aligned from the paragraph indent, 0 pt spacing, no trailing dot.
        for level in (1, 2, 3):
            heading = styles[f"Heading {level}"]
            _set_style_font(heading, style.font_name, style.font_size_pt, is_bold=True)
            heading_format = heading.paragraph_format
            heading_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            heading_format.first_line_indent = Cm(style.first_line_indent_cm)
            heading_format.left_indent = Cm(0)
            heading_format.line_spacing = style.line_spacing
            # No extra spacing around headings: 0 pt before and after.
            heading_format.space_before = Pt(0)
            heading_format.space_after = Pt(0)
            heading_format.keep_with_next = True

        listing = styles.add_style("Listing", WD_STYLE_TYPE.PARAGRAPH)
        listing.base_style = normal
        _set_style_font(listing, style.listing_font_name, style.listing_font_size_pt)
        listing.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        listing.paragraph_format.first_line_indent = Cm(0)
        listing.paragraph_format.line_spacing = style.listing_line_spacing

        # Word picks these built-in names for TOC entries; without them entries inherit
        # the body first-line indent and look broken.
        for level in (1, 2, 3):
            toc_style = styles.add_style(f"toc {level}", WD_STYLE_TYPE.PARAGRAPH)
            toc_style.base_style = normal
            toc_style.paragraph_format.first_line_indent = Cm(0)
            toc_style.paragraph_format.left_indent = Cm(0.5 * (level - 1))
            toc_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # ------------------------------------------------------------------ TOC

    def _add_table_of_contents(self) -> None:
        """Adds the "Содержание" title and a TOC field that Word fills on open.

        The title looks like a heading but uses body style, so it does not list itself.
        """
        title = self._start_paragraph()
        title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        title.add_run(TABLE_OF_CONTENTS_TITLE.capitalize()).bold = True

        toc_paragraph = self._add_text_paragraph("", WD_ALIGN_PARAGRAPH.LEFT)
        _add_field(toc_paragraph, TOC_INSTRUCTION, TOC_PLACEHOLDER)

        # Ask Word to refresh fields on open so the TOC gets real page numbers.
        update_fields = OxmlElement("w:updateFields")
        update_fields.set(qn("w:val"), "true")
        self._document.settings.element.append(update_fields)
        self._is_page_break_pending = True


    # ---------------------------------------------------------------- blocks

    def _add_block(self, block: Block) -> None:
        """Dispatches one block to its renderer."""
        if isinstance(block, Heading):
            self._add_heading(block)
        elif isinstance(block, Paragraph):
            self._add_inline_text(self._start_paragraph(), block.text)
        elif isinstance(block, BulletList):
            self._add_list(block)
        elif isinstance(block, Figure):
            self._add_figure(block)
        elif isinstance(block, Table):
            self._add_table(block)
        elif isinstance(block, Listing):
            self._add_listing(block)
        elif isinstance(block, Formula):
            self._add_formula(block)
        elif isinstance(block, PageBreak):
            self._is_page_break_pending = True

    def _add_heading(self, heading: Heading) -> None:
        """Renders a heading: bold, left aligned from the paragraph indent, no caps.

        Sections: "1.2 Title". Structural elements (Введение, Заключение...): title
        without a number. Appendices: "Приложение А. Title". Every level-1 heading
        starts a new page.
        """
        paragraph = self._document.add_paragraph(style=f"Heading {heading.level}")
        is_top_level = heading.level == 1
        # Every level-1 heading always starts a new page; this is not configurable.
        paragraph.paragraph_format.page_break_before = is_top_level or self._is_page_break_pending
        self._is_page_break_pending = False
        if is_top_level:
            self._is_in_bibliography = heading.title.upper() in BIBLIOGRAPHY_TITLES

        if heading.kind == "structural":
            # Normalizes "ВВЕДЕНИЕ" / "введение" to "Введение"; these titles have no acronyms.
            text = heading.title.capitalize()
        elif heading.kind == "appendix":
            text = f"Приложение {heading.number}. {heading.title}"
        else:
            text = f"{heading.number} {heading.title}" if heading.number else heading.title
        paragraph.add_run(text)

    def _add_list(self, bullet_list: BulletList) -> None:
        """Renders list items with GOST markers: "–" for bullets, "1)" / "а)" for ordered.

        Items of the bibliography are numbered "1." as GOST requires.
        """
        counter_by_level: dict[int, int] = {}
        for item in bullet_list.items:
            # Deeper levels restart numbering when we come back up.
            counter_by_level = {level: count for level, count in counter_by_level.items() if level <= item.level}
            if item.is_ordered:
                counter_by_level[item.level] = counter_by_level.get(item.level, 0) + 1
                marker = _ordered_marker(counter_by_level[item.level], item.level, self._is_in_bibliography)
            else:
                marker = LIST_BULLET

            paragraph = self._start_paragraph()
            paragraph.paragraph_format.left_indent = Cm(self._style.first_line_indent_cm * item.level)
            paragraph.add_run(f"{marker} ")
            self._add_inline_text(paragraph, item.text)

    def _add_figure(self, figure: Figure) -> None:
        """Adds a centered picture with "Рисунок N — caption" below it."""
        picture_paragraph = self._start_paragraph()
        self._format_block_paragraph(picture_paragraph, WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
        picture_paragraph.paragraph_format.space_before = Pt(BLOCK_GAP_PT)
        picture_paragraph.paragraph_format.keep_with_next = True

        # Explicit width wins; otherwise keep native size but never exceed the text width.
        run = picture_paragraph.add_run()
        if figure.width_cm is not None:
            run.add_picture(str(figure.image_path), width=min(Cm(figure.width_cm), self._text_width))
        else:
            picture = run.add_picture(str(figure.image_path))
            if picture.width > self._text_width:
                scale = self._text_width / picture.width
                picture.width = Emu(int(picture.width * scale))
                picture.height = Emu(int(picture.height * scale))

        caption = self._document.add_paragraph()
        self._format_block_paragraph(caption, WD_ALIGN_PARAGRAPH.CENTER)
        caption.paragraph_format.space_after = Pt(BLOCK_GAP_PT)
        self._add_inline_text(caption, f"Рисунок {figure.number} {CAPTION_DASH} {figure.caption}")

    def _add_table(self, table_block: Table) -> None:
        """Adds "Таблица N — caption" above a full-width grid table (GOST)."""
        caption = self._start_paragraph()
        self._format_block_paragraph(caption, WD_ALIGN_PARAGRAPH.LEFT)
        caption.paragraph_format.space_before = Pt(BLOCK_GAP_PT)
        caption.paragraph_format.keep_with_next = True
        self._add_inline_text(caption, f"Таблица {table_block.number} {CAPTION_DASH} {table_block.caption}")

        column_count = max(len(row) for row in table_block.rows)
        table = self._document.add_table(rows=len(table_block.rows), cols=column_count)
        table.style = self._document.styles["Table Grid"]
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_full_width(table)

        for row_index, (row, cells_text) in enumerate(zip(table.rows, table_block.rows)):
            is_header = row_index == 0
            for cell, text in zip(row.cells, cells_text):
                paragraph = cell.paragraphs[0]
                alignment = WD_ALIGN_PARAGRAPH.CENTER if is_header else WD_ALIGN_PARAGRAPH.LEFT
                self._format_block_paragraph(paragraph, alignment, line_spacing=1.0)
                self._add_inline_text(paragraph, text, font_size_pt=self._style.table_font_size_pt)
        # The header row repeats on every page the table spans (GOST).
        _mark_header_row(table.rows[0])

        # Empty line after the table so the next paragraph does not stick to it.
        self._add_text_paragraph("", WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.0)

    def _add_listing(self, listing: Listing) -> None:
        """Adds "Листинг N — caption" and the code, one paragraph per line, monospace."""
        if listing.caption:
            caption = self._start_paragraph()
            self._format_block_paragraph(caption, WD_ALIGN_PARAGRAPH.LEFT)
            caption.paragraph_format.space_before = Pt(BLOCK_GAP_PT)
            caption.paragraph_format.keep_with_next = True
            self._add_inline_text(caption, f"Листинг {listing.number} {CAPTION_DASH} {listing.caption}")

        code_lines = listing.code.expandtabs(TAB_SIZE_SPACES).split("\n")
        for line_index, line in enumerate(code_lines):
            paragraph = self._start_paragraph(style_name="Listing")
            # Short listings stay on one page; long ones are allowed to break.
            paragraph.paragraph_format.keep_with_next = len(code_lines) <= 15
            if line_index == len(code_lines) - 1:
                paragraph.paragraph_format.space_after = Pt(BLOCK_GAP_PT)
            if line:
                paragraph.add_run(line)

    def _add_formula(self, formula: Formula) -> None:
        """Centers the formula with its number "(N)" at the right margin (GOST)."""
        paragraph = self._start_paragraph()
        self._format_block_paragraph(paragraph, WD_ALIGN_PARAGRAPH.LEFT)
        paragraph.paragraph_format.space_before = Pt(BLOCK_GAP_PT / 2)
        paragraph.paragraph_format.space_after = Pt(BLOCK_GAP_PT / 2)
        tab_stops = paragraph.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Emu(self._text_width // 2), WD_TAB_ALIGNMENT.CENTER)
        tab_stops.add_tab_stop(self._text_width, WD_TAB_ALIGNMENT.RIGHT)

        paragraph.add_run("\t")
        self._add_inline_text(paragraph, formula.expression)
        paragraph.add_run(f"\t({formula.number})")

    # --------------------------------------------------------------- helpers

    def _start_paragraph(self, style_name: str | None = None) -> DocxParagraph:
        """Adds a body paragraph, applying a pending page break to it."""
        paragraph = self._document.add_paragraph(style=style_name)
        if self._is_page_break_pending:
            paragraph.paragraph_format.page_break_before = True
            self._is_page_break_pending = False
        return paragraph

    def _add_text_paragraph(
        self, text: str, alignment: WD_ALIGN_PARAGRAPH, *, line_spacing: float | None = None
    ) -> DocxParagraph:
        """Adds a plain paragraph without indent (gaps, TOC field)."""
        paragraph = self._start_paragraph()
        self._format_block_paragraph(paragraph, alignment, line_spacing=line_spacing)
        if text:
            paragraph.add_run(text)
        return paragraph

    def _format_block_paragraph(
        self, paragraph: DocxParagraph, alignment: WD_ALIGN_PARAGRAPH, *, line_spacing: float | None = None
    ) -> None:
        """Removes the body indent and sets alignment for captions, cells and similar."""
        paragraph_format = paragraph.paragraph_format
        paragraph_format.alignment = alignment
        paragraph_format.first_line_indent = Cm(0)
        if line_spacing is not None:
            paragraph_format.line_spacing = line_spacing

    def _add_inline_text(self, paragraph: DocxParagraph, text: str, *, font_size_pt: float | None = None) -> None:
        """Appends text with **bold**, *italic* and `code` runs to a paragraph."""
        for inline_run in parse_inline(text):
            run = paragraph.add_run(inline_run.text)
            run.bold = inline_run.is_bold or None
            run.italic = inline_run.is_italic or None
            if inline_run.is_code:
                _set_run_font_name(run, self._style.listing_font_name)
            if font_size_pt is not None:
                run.font.size = Pt(font_size_pt)


def _ordered_marker(number: int, level: int, is_bibliography: bool) -> str:
    """Returns the marker of an ordered list item: "1.", "1)" or "а)"."""
    if is_bibliography and level == 0:
        return f"{number}."
    if level == 0:
        return f"{number})"
    return f"{CYRILLIC_LIST_LETTERS[(number - 1) % len(CYRILLIC_LIST_LETTERS)]})"


def _set_style_font(style: ParagraphStyle, font_name: str, size_pt: float, *, is_bold: bool = False) -> None:
    """Sets font, size, weight and black color on a style, dropping theme fonts.

    Built-in heading styles reference theme fonts (Calibri Light) through
    ``asciiTheme`` attributes that override ``w:ascii``; they must be removed or
    Word keeps showing the theme font.
    """
    style.font.size = Pt(size_pt)
    style.font.bold = is_bold
    style.font.italic = False
    style.font.color.rgb = RGBColor(0, 0, 0)
    fonts = style.element.get_or_add_rPr().get_or_add_rFonts()
    for theme_attribute in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        fonts.attrib.pop(qn(theme_attribute), None)
    for script_attribute in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(script_attribute), font_name)


def _set_run_font_name(run: Run, font_name: str) -> None:
    """Sets a run font for all scripts, Cyrillic included."""
    run.font.name = font_name
    fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    fonts.set(qn("w:eastAsia"), font_name)
    fonts.set(qn("w:cs"), font_name)


def _set_style_language(style: ParagraphStyle, language: str) -> None:
    """Marks style text as Russian so Word spell-checks and hyphenates it correctly."""
    run_properties = style.element.get_or_add_rPr()
    language_element = OxmlElement("w:lang")
    for attribute in ("w:val", "w:eastAsia"):
        language_element.set(qn(attribute), language)
    run_properties.append(language_element)


def _add_field(paragraph: DocxParagraph, instruction: str, placeholder: str) -> None:
    """Appends a Word field (PAGE, TOC...) with placeholder text shown until update."""
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    instruction_text = OxmlElement("w:instrText")
    instruction_text.set(qn("xml:space"), "preserve")
    instruction_text.text = f" {instruction} "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    paragraph.add_run()._element.append(begin)
    paragraph.add_run()._element.append(instruction_text)
    paragraph.add_run()._element.append(separate)
    paragraph.add_run(placeholder)
    paragraph.add_run()._element.append(end)


def _set_table_full_width(table: DocxTable) -> None:
    """Stretches a table to 100% of the text width."""
    table_properties = table._tbl.tblPr
    width = table_properties.find(qn("w:tblW"))
    if width is None:
        width = OxmlElement("w:tblW")
        table_properties.append(width)
    width.set(qn("w:type"), "pct")
    width.set(qn("w:w"), TABLE_FULL_WIDTH_PCT)


def _mark_header_row(row: _Row) -> None:
    """Makes a table row repeat at the top of every page."""
    header_marker = OxmlElement("w:tblHeader")
    header_marker.set(qn("w:val"), "true")
    row._tr.get_or_add_trPr().append(header_marker)


def build_report(plain_dir: Path, output_path: Path) -> None:
    """Parses the plain folder and writes the .docx report.

    Args:
        plain_dir: Folder with *.md files, images/ and an optional style.json.
        output_path: Where to save the .docx file.

    Raises:
        StyleError: If style.json is invalid.
        PlainError: If the Markdown files are invalid.
        OSError: If the output cannot be written (e.g. the file is open in Word).
    """
    style = load_report_style(plain_dir / STYLE_FILE_NAME)
    markdown_paths = sorted(plain_dir.glob("*.md"))
    if not markdown_paths:
        raise PlainError(f"No .md files in {plain_dir}")

    blocks: list[Block] = []
    for markdown_path in markdown_paths:
        blocks.extend(parse_plain_file(markdown_path, plain_dir))
    number_by_label = number_blocks(blocks, is_numbered_by_section=style.caption_numbering == "by_section")
    resolve_references(blocks, number_by_label)

    document = ReportRenderer(style).render(blocks)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_path))
    logger.info("Built %s from %d files (%d blocks)", output_path, len(markdown_paths), len(blocks))


def main(argv: Sequence[str] | None = None) -> int:
    """Command-line entry point; returns the process exit code."""
    parser = argparse.ArgumentParser(description="Build a GOST .docx report from a plain folder.")
    parser.add_argument("plain_dir", type=Path, help="plain folder with *.md files")
    parser.add_argument("-o", "--output", type=Path, help=f"output .docx (default: <plain_dir>/../{OUTPUT_DIR_NAME}/{DEFAULT_OUTPUT_NAME})")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    plain_dir = args.plain_dir.resolve()
    output_path = (args.output or plain_dir.parent / OUTPUT_DIR_NAME / DEFAULT_OUTPUT_NAME).resolve()
    try:
        build_report(plain_dir, output_path)
    except (StyleError, PlainError) as error:
        logger.error("%s", error)
        return 1
    except OSError as error:
        logger.error("Cannot write %s (is it open in Word?): %s", output_path, error)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
