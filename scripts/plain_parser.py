"""Parses plain Markdown files into typed blocks and numbers them by GOST.

Purpose:
    The report builder works on a small, strict Markdown dialect (described in
    SKILL.md and references/plain-format.md). This module turns plain files into a
    flat list of blocks, numbers sections, appendices, figures, tables, listings and
    formulas, and replaces cross-references like ``[@fig:scheme]`` with numbers.
    Unsupported or broken syntax fails with "file:line: reason" instead of being
    silently rendered wrong.

Components:
    Heading, Paragraph, BulletList, Figure, Table, Listing, Formula, PageBreak: blocks.
    parse_plain_file: one Markdown file -> list of blocks.
    number_blocks: assigns GOST numbers and returns the label -> number index.
    resolve_references: substitutes ``[@label]`` in all text.
    parse_inline: splits text into bold / italic / code runs.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterator, Mapping, Union

STRUCTURAL_TITLES = frozenset(
    {
        "СПИСОК ИСПОЛНИТЕЛЕЙ",
        "РЕФЕРАТ",
        "АННОТАЦИЯ",
        "ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ",
        "ОПРЕДЕЛЕНИЯ",
        "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ",
        "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ",
        "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ",
        "ВВЕДЕНИЕ",
        "ЗАКЛЮЧЕНИЕ",
        "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
        "СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ",
        "СПИСОК ЛИТЕРАТУРЫ",
        "БИБЛИОГРАФИЧЕСКИЙ СПИСОК",
    }
)
BIBLIOGRAPHY_TITLES = frozenset(title for title in STRUCTURAL_TITLES if "СПИСОК" in title and "ИСПОЛНИТЕЛЕЙ" not in title)
TABLE_OF_CONTENTS_TITLE = "СОДЕРЖАНИЕ"

# GOST: appendix letters skip Ё, З, Й, О, Ч, Ь, Ы, Ъ.
APPENDIX_LETTERS = "АБВГДЕЖИКЛМНПРСТУФХЦШЩЭЮЯ"
LIST_INDENT_SPACES = 2
PAGE_BREAK_MARKERS = frozenset({"\\newpage", "\\pagebreak"})

_HEADING_RE = re.compile(r"^(?P<hashes>#{1,3})\s+(?P<title>.+?)\s*$")
_ATTRIBUTES_RE = re.compile(r"\s*\{(?P<body>[^{}]*)\}\s*$")
_FIGURE_RE = re.compile(r"^!\[(?P<caption>[^\]]*)\]\((?P<path>[^)]+)\)\s*(?P<attrs>\{[^{}]*\})?\s*$")
_TABLE_CAPTION_RE = re.compile(r"^(?:Таблица|Table):\s*(?P<caption>.+)$")
_LISTING_CAPTION_RE = re.compile(r"^(?:Листинг|Listing):\s*(?P<caption>.+)$")
_FORMULA_RE = re.compile(r"^\$\$(?P<expression>.+?)\$\$\s*(?P<attrs>\{[^{}]*\})?\s*$")
_LIST_ITEM_RE = re.compile(r"^(?P<indent>[ \t]*)(?P<marker>[-*+]|\d+[.)])\s+(?P<text>.+)$")
_TABLE_SEPARATOR_RE = re.compile(r"^\|?(\s*:?-{3,}:?\s*\|)+\s*:?-*:?\s*$")
# "Приложение А. Название", "Приложение: Название", "Приложение — Название".
_APPENDIX_RE = re.compile(r"^приложение(?:\s+[а-яa-z](?=[\s.:—–-]|$))?[\s.:—–-]*(?P<title>.*)$", re.IGNORECASE)
_REFERENCE_RE = re.compile(r"\[@(?P<label>[\w.:-]+)\]")
# Bold, inline code, italic; italic needs non-space next to both stars so "a * b" stays text.
_INLINE_RE = re.compile(r"(\*\*.+?\*\*|`[^`]+`|\*[^*\s](?:[^*]*?[^*\s])?\*)")
_TABLE_CELL_SPLIT_RE = re.compile(r"(?<!\\)\|")


class PlainError(Exception):
    """Raised when plain files break the supported Markdown dialect."""


@dataclass
class Heading:
    """Section heading; ``kind`` decides how GOST wants it laid out."""

    level: int
    title: str
    kind: str  # "section" | "structural" | "appendix"
    is_numbered: bool
    label: str | None
    number: str = ""


@dataclass
class Paragraph:
    """Body text paragraph with inline Markdown markup."""

    text: str


@dataclass
class ListItem:
    """One list item; ``level`` 0 is the outermost list."""

    level: int
    text: str
    is_ordered: bool


@dataclass
class BulletList:
    """Consecutive list items (ordered and unordered may be mixed by level)."""

    items: list[ListItem]


@dataclass
class Figure:
    """Picture with a mandatory caption, rendered as "Рисунок N — caption"."""

    image_path: Path
    caption: str
    label: str | None
    width_cm: float | None
    number: str = ""


@dataclass
class Table:
    """Pipe table; the first row is the header and repeats on page breaks."""

    caption: str
    label: str | None
    rows: list[list[str]]
    number: str = ""


@dataclass
class Listing:
    """Code block; numbered only when it has a "Листинг:" caption."""

    caption: str
    label: str | None
    code: str
    number: str = ""


@dataclass
class Formula:
    """Display formula, centered with its number "(N)" at the right edge."""

    expression: str
    label: str | None
    number: str = ""


@dataclass
class PageBreak:
    """Explicit page break requested with \\newpage."""


Block = Union[Heading, Paragraph, BulletList, Figure, Table, Listing, Formula, PageBreak]


@dataclass(frozen=True)
class InlineRun:
    """Piece of text with uniform formatting."""

    text: str
    is_bold: bool = False
    is_italic: bool = False
    is_code: bool = False


@dataclass(frozen=True)
class _Attributes:
    """Values from a trailing ``{#label - width=12}`` attribute block."""

    label: str | None = None
    is_unnumbered: bool = False
    width_cm: float | None = None


@dataclass
class _PendingCaption:
    """A "Таблица:" / "Листинг:" line waiting for its table or code block."""

    kind: str  # "table" | "listing"
    caption: str
    label: str | None
    line_number: int


def parse_plain_file(markdown_path: Path, plain_dir: Path) -> list[Block]:
    """Parses one plain Markdown file.

    Args:
        markdown_path: The Markdown file to parse.
        plain_dir: plain folder; image paths are resolved against it.

    Returns:
        Blocks in document order.

    Raises:
        PlainError: With "file:line: reason" on any syntax problem or missing image.
    """
    return _PlainFileParser(markdown_path, plain_dir).parse()


def number_blocks(blocks: list[Block], *, is_numbered_by_section: bool) -> dict[str, str]:
    """Assigns GOST numbers to headings and captioned blocks in place.

    Sections get "1", "1.1", "1.1.1"; appendices get letters and their subsections
    "А.1". Figures, tables, listings and formulas are numbered per kind, either
    through the whole report or within each section ("2.3"); inside appendices
    they are always "А.1" (GOST).

    Args:
        blocks: All blocks of the report in order.
        is_numbered_by_section: Number captions within sections instead of continuously.

    Returns:
        Mapping from every label to its number, for resolve_references.

    Raises:
        PlainError: On duplicate labels or too many appendices.
    """
    number_by_label: dict[str, str] = {}
    section_count = 0
    appendix_count = 0
    top_prefix = ""  # number/letter of the current level-1 heading, "" if unnumbered
    subsection_counts = [0, 0]  # counters for heading levels 2 and 3
    caption_scope = ""
    caption_counts: Counter[str] = Counter()

    for block in blocks:
        if isinstance(block, Heading):
            if block.level == 1:
                subsection_counts = [0, 0]
                if block.kind == "appendix":
                    if appendix_count >= len(APPENDIX_LETTERS):
                        raise PlainError("Too many appendices for GOST letter numbering")
                    top_prefix = APPENDIX_LETTERS[appendix_count]
                    appendix_count += 1
                elif block.is_numbered:
                    section_count += 1
                    top_prefix = str(section_count)
                else:
                    top_prefix = ""
                block.number = top_prefix

                # Caption counters restart when their scope (section / appendix) changes.
                is_scoped = block.kind == "appendix" or is_numbered_by_section
                new_scope = top_prefix if is_scoped else ""
                if new_scope != caption_scope:
                    caption_scope = new_scope
                    caption_counts.clear()
            elif block.is_numbered and top_prefix:
                # Level 2 resets level 3; a level 3 heading needs a numbered parent.
                depth = block.level - 2
                subsection_counts[depth] += 1
                if depth == 0:
                    subsection_counts[1] = 0
                parent_numbers = subsection_counts[: depth + 1]
                if all(parent_numbers):
                    block.number = ".".join([top_prefix, *map(str, parent_numbers)])
            _register_label(number_by_label, block.label, block.number)
            continue

        if isinstance(block, (Figure, Table, Formula)) or (isinstance(block, Listing) and block.caption):
            kind = type(block).__name__
            caption_counts[kind] += 1
            count = caption_counts[kind]
            block.number = f"{caption_scope}.{count}" if caption_scope else str(count)
            # GOST: formula references are written in brackets, "по формуле (1)".
            reference_text = f"({block.number})" if isinstance(block, Formula) else block.number
            _register_label(number_by_label, block.label, reference_text)

    return number_by_label


def resolve_references(blocks: list[Block], number_by_label: Mapping[str, str]) -> None:
    """Replaces every ``[@label]`` in text with the label's number, in place.

    Args:
        blocks: Numbered blocks.
        number_by_label: Index returned by number_blocks.

    Raises:
        PlainError: If a reference points to an unknown label.
    """

    def substitute(match: re.Match[str]) -> str:
        label = match.group("label")
        if label not in number_by_label:
            raise PlainError(f"Unknown reference [@{label}]; known labels: {', '.join(sorted(number_by_label))}")
        return number_by_label[label]

    def resolve(text: str) -> str:
        return _REFERENCE_RE.sub(substitute, text)

    for block in blocks:
        if isinstance(block, Paragraph):
            block.text = resolve(block.text)
        elif isinstance(block, BulletList):
            for item in block.items:
                item.text = resolve(item.text)
        elif isinstance(block, Table):
            block.caption = resolve(block.caption)
            block.rows = [[resolve(cell) for cell in row] for row in block.rows]
        elif isinstance(block, (Figure, Listing)):
            block.caption = resolve(block.caption)


def parse_inline(text: str) -> list[InlineRun]:
    """Splits text into runs by ``**bold**``, ``*italic*`` and ```code``` markup."""
    runs = []
    # re.split with a capturing group puts the matched markup at odd indexes.
    for index, part in enumerate(_INLINE_RE.split(text)):
        if not part:
            continue
        if index % 2 == 0:
            runs.append(InlineRun(part))
        elif part.startswith("**"):
            runs.append(InlineRun(part[2:-2], is_bold=True))
        elif part.startswith("`"):
            runs.append(InlineRun(part[1:-1], is_code=True))
        else:
            runs.append(InlineRun(part[1:-1], is_italic=True))
    return runs


def _register_label(number_by_label: dict[str, str], label: str | None, number: str) -> None:
    """Adds a label to the index, rejecting duplicates and labels on unnumbered blocks."""
    if label is None:
        return
    if label in number_by_label:
        raise PlainError(f"Duplicate label #{label}")
    if not number:
        raise PlainError(f"Label #{label} is attached to an unnumbered element")
    number_by_label[label] = number


def _split_attributes(text: str) -> tuple[str, _Attributes]:
    """Cuts a trailing ``{...}`` attribute block off ``text`` and parses it.

    Supported tokens: ``#label``, ``-`` (unnumbered heading), ``width=12`` (cm).

    Raises:
        ValueError: On an unknown token or a bad width.
    """
    match = _ATTRIBUTES_RE.search(text)
    if match is None:
        return text, _Attributes()

    label = None
    is_unnumbered = False
    width_cm = None
    for token in match.group("body").split():
        if token in ("-", ".unnumbered"):
            is_unnumbered = True
        elif token.startswith("#"):
            label = token[1:]
        elif token.startswith("width="):
            width_cm = float(token[len("width=") :].removesuffix("cm"))
        else:
            raise ValueError(f"Unknown attribute '{token}' (allowed: #label, -, width=N)")
    return text[: match.start()].rstrip(), _Attributes(label, is_unnumbered, width_cm)


def _split_table_row(line: str) -> list[str]:
    """Splits ``| a | b |`` into cells, honouring escaped ``\\|``."""
    inner = line.strip().removeprefix("|").removesuffix("|")
    return [cell.strip().replace("\\|", "|") for cell in _TABLE_CELL_SPLIT_RE.split(inner)]


class _PlainFileParser:
    """Line-by-line parser for one plain file; keeps position for error messages."""

    def __init__(self, markdown_path: Path, plain_dir: Path) -> None:
        """Reads the file.

        Args:
            markdown_path: The Markdown file to parse.
            plain_dir: plain folder for resolving image paths.
        """
        self._path = markdown_path
        self._plain_dir = plain_dir
        self._lines = markdown_path.read_text(encoding="utf-8").splitlines()
        self._position = 0
        self._blocks: list[Block] = []
        self._pending_caption: _PendingCaption | None = None

    def parse(self) -> list[Block]:
        """Parses the whole file.

        Returns:
            Blocks in document order.

        Raises:
            PlainError: With file and line on any syntax problem.
        """
        while self._position < len(self._lines):
            stripped = self._lines[self._position].strip()
            if not stripped:
                self._position += 1
                continue
            try:
                self._parse_block(stripped)
            except ValueError as error:
                raise self._error(str(error)) from error

        if self._pending_caption is not None:
            raise self._error("Caption is not followed by a table or code block", self._pending_caption.line_number)
        return self._blocks

    def _error(self, message: str, line_number: int | None = None) -> PlainError:
        """Builds a PlainError pointing at the current (or given) line."""
        return PlainError(f"{self._path.name}:{line_number or self._position + 1}: {message}")

    def _parse_block(self, stripped: str) -> None:
        """Dispatches on the first line of a block and consumes the whole block."""
        # A "Таблица:" / "Листинг:" line must be immediately followed by its block.
        if self._pending_caption is not None:
            expected_start = "|" if self._pending_caption.kind == "table" else "```"
            if not stripped.startswith(expected_start):
                raise self._error(
                    f"'{self._pending_caption.caption}' must be followed by a {self._pending_caption.kind}",
                    self._pending_caption.line_number,
                )

        if stripped in PAGE_BREAK_MARKERS:
            self._blocks.append(PageBreak())
            self._position += 1
        elif stripped.startswith("```"):
            self._parse_listing()
        elif stripped.startswith("|"):
            self._parse_table()
        elif stripped.startswith("####"):
            raise self._error("Only heading levels 1-3 are allowed (GOST: section, subsection, item)")
        elif heading_match := _HEADING_RE.match(stripped):
            self._parse_heading(len(heading_match.group("hashes")), heading_match.group("title"))
        elif caption_match := _TABLE_CAPTION_RE.match(stripped):
            self._remember_caption("table", caption_match.group("caption"))
        elif caption_match := _LISTING_CAPTION_RE.match(stripped):
            self._remember_caption("listing", caption_match.group("caption"))
        elif figure_match := _FIGURE_RE.match(stripped):
            self._parse_figure(figure_match)
        elif formula_match := _FORMULA_RE.match(stripped):
            _, attributes = _split_attributes(formula_match.group("attrs") or "")
            self._blocks.append(Formula(formula_match.group("expression").strip(), attributes.label))
            self._position += 1
        elif _LIST_ITEM_RE.match(self._lines[self._position]):
            self._parse_list()
        else:
            self._parse_paragraph()

    def _is_block_start(self, stripped: str) -> bool:
        """Checks whether a line opens a new block, ending the current paragraph."""
        return (
            stripped in PAGE_BREAK_MARKERS
            or stripped.startswith(("```", "|", "#", "![", "$$"))
            or bool(_TABLE_CAPTION_RE.match(stripped) or _LISTING_CAPTION_RE.match(stripped))
            or bool(_LIST_ITEM_RE.match(stripped))
        )

    def _parse_heading(self, level: int, raw_title: str) -> None:
        """Parses ``# Title {#label -}`` and classifies it as section/structural/appendix."""
        title, attributes = _split_attributes(raw_title)
        kind = "section"
        if level == 1:
            upper_title = title.upper()
            if upper_title == TABLE_OF_CONTENTS_TITLE:
                raise self._error("Do not write СОДЕРЖАНИЕ: the builder generates it")
            appendix_match = _APPENDIX_RE.match(title)
            if upper_title in STRUCTURAL_TITLES:
                kind = "structural"
            elif appendix_match:
                kind = "appendix"
                title = appendix_match.group("title").strip()
                if not title:
                    raise self._error("Appendix needs a title: '# Приложение: Название'")

        is_numbered = kind != "structural" and not attributes.is_unnumbered
        self._blocks.append(Heading(level, title, kind, is_numbered, attributes.label))
        self._position += 1

    def _remember_caption(self, kind: str, raw_caption: str) -> None:
        """Stores a table/listing caption until its block arrives."""
        caption, attributes = _split_attributes(raw_caption)
        self._pending_caption = _PendingCaption(kind, caption, attributes.label, self._position + 1)
        self._position += 1

    def _take_caption(self, kind: str) -> _PendingCaption | None:
        """Returns and clears the pending caption if it belongs to ``kind``."""
        caption = self._pending_caption
        self._pending_caption = None
        return caption if caption is not None and caption.kind == kind else None

    def _parse_figure(self, figure_match: re.Match[str]) -> None:
        """Parses ``![Caption](images/x.png){#fig:x width=12}`` and checks the file exists."""
        caption = figure_match.group("caption").strip()
        if not caption:
            raise self._error("Figure needs a caption: ![Caption](images/file.png)")
        image_path = (self._plain_dir / figure_match.group("path").strip()).resolve()
        if not image_path.is_file():
            raise self._error(f"Image not found: {image_path}")
        _, attributes = _split_attributes(figure_match.group("attrs") or "")
        self._blocks.append(Figure(image_path, caption, attributes.label, attributes.width_cm))
        self._position += 1

    def _parse_table(self) -> None:
        """Consumes consecutive ``|`` lines; the caption line above is mandatory."""
        caption = self._take_caption("table")
        if caption is None:
            raise self._error("Table needs a caption line above it: 'Таблица: Название {#tbl:label}'")

        rows = []
        for line in self._consume_while(lambda stripped: stripped.startswith("|")):
            if not _TABLE_SEPARATOR_RE.match(line):
                rows.append(_split_table_row(line))
        self._blocks.append(Table(caption.caption, caption.label, rows))

    def _parse_listing(self) -> None:
        """Consumes a fenced code block; it is numbered only with a "Листинг:" caption."""
        caption = self._take_caption("listing")
        opening_line_number = self._position + 1
        self._position += 1

        code_lines = []
        while self._position < len(self._lines):
            line = self._lines[self._position]
            self._position += 1
            if line.strip().startswith("```"):
                break
            code_lines.append(line)
        else:
            raise self._error("Code block is not closed with ```", opening_line_number)

        self._blocks.append(
            Listing(
                caption=caption.caption if caption else "",
                label=caption.label if caption else None,
                code="\n".join(code_lines),
            )
        )

    def _parse_list(self) -> None:
        """Consumes list items; lines without a marker continue the previous item."""
        items: list[ListItem] = []
        while self._position < len(self._lines):
            line = self._lines[self._position]
            stripped = line.strip()
            item_match = _LIST_ITEM_RE.match(line)
            if item_match:
                indent_width = len(item_match.group("indent").expandtabs(4))
                is_ordered = item_match.group("marker")[0].isdigit()
                items.append(ListItem(indent_width // LIST_INDENT_SPACES, item_match.group("text").strip(), is_ordered))
            elif stripped and items and not self._is_block_start(stripped):
                items[-1].text = f"{items[-1].text} {stripped}"
            else:
                break
            self._position += 1
        self._blocks.append(BulletList(items))

    def _parse_paragraph(self) -> None:
        """Joins lines up to a blank line or the start of another block."""
        lines = [self._lines[self._position].strip()]
        self._position += 1
        lines.extend(self._consume_while(lambda stripped: bool(stripped) and not self._is_block_start(stripped)))
        self._blocks.append(Paragraph(" ".join(lines)))

    def _consume_while(self, predicate: Callable[[str], bool]) -> Iterator[str]:
        """Yields stripped lines while ``predicate`` holds, advancing the position."""
        while self._position < len(self._lines):
            stripped = self._lines[self._position].strip()
            if not predicate(stripped):
                return
            self._position += 1
            yield stripped

