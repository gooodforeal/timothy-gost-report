"""Flags phrases and habits that make plain text read as machine-written.

Purpose:
    Before building the report the agent runs this check over the plain folder and
    rewrites every flagged sentence. It catches the typical tells of generated
    Russian text (filler openers, "важно отметить", "играет ключевую роль"),
    formatting that GOST forbids in body text (bold, straight quotes, hyphen used
    as a dash) and overused words. It is a linter, not a judge: a clean run does
    not prove the text is good, see references/writing-style.md.

Usage:
    python3 check_text.py path/to/plain

Exit code:
    0 when nothing was found, 1 when there are findings or the folder is invalid.
"""

from __future__ import annotations

import re
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Sequence

# (pattern, hint). Patterns are matched case-insensitively against prose lines.
CLICHE_RULES: tuple[tuple[str, str], ...] = (
    (r"в современном мире|в наше время|в наши дни|в эпоху|в условиях стремительн", "пустой зачин — начните с сути"),
    (r"игра\w* (ключев|важн|важнейш|значительн|решающ)\w* роль", "штамп — скажите, что именно делает объект"),
    (r"(важно|стоит|следует|необходимо|нельзя не|хотелось бы) (отметить|подчеркнуть|понимать|заметить)", "вводная вода — удалите"),
    (r"неотъемлем", "штамп «неотъемлемая часть»"),
    (r"широк\w* (спектр|круг|ряд)|огромн\w* (количеств|числ)", "размыто — дайте число или перечень"),
    (r"подводя итог|в заключение,? (можно|следует|стоит)|резюмируя", "штамп концовки"),
    (r"давайте|рассмотрим подробнее|погрузимся|окунемся", "разговорный оборот ИИ"),
    (r"уникальн|инновационн|революционн|передов\w+|мощн\w+ инструмент", "рекламная лексика"),
    (r"комплексн\w* подход|ключев\w* (аспект|момент|фактор)|данн\w+ аспект", "канцелярский штамп"),
    (r"синерги|бесшовн|трансформаци\w+ (бизнес|процесс)", "модное слово без смысла"),
    (r"это не просто|не просто \w+, а ", "конструкция «не просто X, а Y»"),
    (r"по сути|в целом|так или иначе|в той или иной степени", "слово-паразит"),
    (r"эффективн\w+ (решени|инструмент|способ)|надежн\w+ решени", "оценка без доказательства"),
    (r"позволя\w+ (значительно|существенно|эффективно)", "усилитель без цифр"),
)
OVERUSED_WORD_LIMITS: dict[str, int] = {"является": 4, "таким образом": 2, "данный": 4, "также": 6}
MIN_SENTENCES_FOR_RHYTHM_CHECK = 6
MAX_SIMILAR_LENGTH_SHARE = 0.7  # share of sentences within ±25% of the median length

_COMPILED_RULES = [(re.compile(pattern, re.IGNORECASE), hint) for pattern, hint in CLICHE_RULES]
_BOLD_RE = re.compile(r"\*\*[^*]+\*\*")
_STRAIGHT_QUOTES_RE = re.compile(r'"[^"]+"')
_HYPHEN_AS_DASH_RE = re.compile(r"\s-\s")
_EM_DASH_RE = re.compile(r"\s—\s")
_SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?])\s+")
_SKIPPED_LINE_PREFIXES = ("#", "|", "![", "$$", "Таблица:", "Листинг:")


@dataclass(frozen=True)
class Finding:
    """One problem in a plain file."""

    file_name: str
    line_number: int
    fragment: str
    hint: str

    def __str__(self) -> str:
        """Formats the finding as "file:line: «fragment» — hint"."""
        location = f"{self.file_name}:{self.line_number}" if self.line_number else self.file_name
        return f"{location}: «{self.fragment}» — {self.hint}"


def check_file(markdown_path: Path) -> list[Finding]:
    """Checks one plain file.

    Args:
        markdown_path: The Markdown file to check.

    Returns:
        Findings in line order, followed by whole-file findings.
    """
    findings: list[Finding] = []
    prose_lines = list(_iterate_prose_lines(markdown_path))
    name = markdown_path.name

    for line_number, line in prose_lines:
        for pattern, hint in _COMPILED_RULES:
            findings.extend(Finding(name, line_number, match.group(0), hint) for match in pattern.finditer(line))
        if _BOLD_RE.search(line):
            findings.append(Finding(name, line_number, "**…**", "ГОСТ: полужирный только в заголовках"))
        if _STRAIGHT_QUOTES_RE.search(line):
            findings.append(Finding(name, line_number, '"…"', "используйте кавычки-ёлочки «…»"))
        if _HYPHEN_AS_DASH_RE.search(line):
            findings.append(Finding(name, line_number, " - ", "дефис вместо тире — пишите « – »"))
        if len(_EM_DASH_RE.findall(line)) >= 2:
            findings.append(Finding(name, line_number, "— … —", "много тире в одном абзаце — типичный признак ИИ"))

    # Whole-file checks: overused words and monotonous sentence length.
    full_text = " ".join(line for _, line in prose_lines).lower()
    for word, limit in OVERUSED_WORD_LIMITS.items():
        count = len(re.findall(rf"\b{word}\b", full_text))
        if count > limit:
            findings.append(Finding(name, 0, word, f"встречается {count} раз (не больше {limit})"))
    if _has_monotonous_rhythm(full_text):
        findings.append(Finding(name, 0, "ритм", "предложения почти одной длины — чередуйте короткие и длинные"))
    return findings


def _iterate_prose_lines(markdown_path: Path) -> Iterator[tuple[int, str]]:
    """Yields (line number, line) for body text, skipping code, tables, captions, headings."""
    is_in_code_block = False
    for line_number, line in enumerate(markdown_path.read_text(encoding="utf-8").splitlines(), start=1):
        stripped = line.strip()
        if stripped.startswith("```"):
            is_in_code_block = not is_in_code_block
            continue
        if is_in_code_block or not stripped or stripped.startswith(_SKIPPED_LINE_PREFIXES):
            continue
        yield line_number, stripped


def _has_monotonous_rhythm(text: str) -> bool:
    """Checks whether most sentences have nearly the same length (a generated-text tell)."""
    lengths = sorted(len(sentence.split()) for sentence in _SENTENCE_SPLIT_RE.split(text) if sentence.strip())
    if len(lengths) < MIN_SENTENCES_FOR_RHYTHM_CHECK:
        return False
    median = lengths[len(lengths) // 2]
    similar_count = sum(1 for length in lengths if abs(length - median) <= median * 0.25)
    return similar_count / len(lengths) > MAX_SIMILAR_LENGTH_SHARE


def main(argv: Sequence[str] | None = None) -> int:
    """Command-line entry point; prints findings and returns the exit code."""
    arguments = list(sys.argv[1:] if argv is None else argv)
    if len(arguments) != 1:
        print("usage: check_text.py PLAIN_DIR", file=sys.stderr)
        return 1
    plain_dir = Path(arguments[0])
    markdown_paths = sorted(plain_dir.glob("*.md"))
    if not markdown_paths:
        print(f"No .md files in {plain_dir}", file=sys.stderr)
        return 1

    findings = [finding for path in markdown_paths for finding in check_file(path)]
    for finding in findings:
        print(finding)
    hint_counts = Counter(finding.hint for finding in findings)
    print(f"\nНайдено замечаний: {len(findings)}" + (f" (чаще всего: {hint_counts.most_common(1)[0][0]})" if findings else ""))
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
