#!/usr/bin/env python3
# 주요통계집 마크다운의 HTML/Markdown 표를 Excel 파일로 저장
# "1-1" 같은 번호 줄과 그 직전 제목 줄을 표 제목으로 사용

from __future__ import annotations

import html
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import TypeAlias

from table_utils import (
    MergeRange,
    TABLE_RE,
    parse_table_to_grid,
    sanitize_filename,
    unique_filename,
    write_excel,
)

# 주요통계집 본문 제목 번호 탐색
TITLE_CODE_RE = re.compile(r"^\s*(\d+-\d+)\.?\s*$")
MARKDOWN_TABLE_ROW_RE = re.compile(r"^\s*\|.*\|\s*$")
MARKDOWN_SEPARATOR_CELL_RE = re.compile(r"^:?-{3,}:?$")
HTML_BR_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)
MARKDOWN_ESCAPE_RE = re.compile(r"\\([\\`*_{}\[\]()#+\-.!|~])")
IMAGE_LINE_RE = re.compile(r"^!\[.*?\]\(.*?\)$")
ParsedGrid: TypeAlias = tuple[list[list[str]], list[MergeRange], list[bool]]


@dataclass(frozen=True)
class KeyStatHeading:
    code: str
    title: str
    start: int

    @property
    def file_title(self) -> str:
        return f"{self.code} {self.title}".strip()


@dataclass(frozen=True)
class KeyStatTableMatch:
    heading: KeyStatHeading | None
    source: str
    table_type: str
    start: int
    line_no: int
    index: int


# 파일 위치와 줄 번호 계산용 라인 정보 생성
def _iter_lines_with_offsets(text: str) -> list[tuple[int, int, str]]:
    lines: list[tuple[int, int, str]] = []
    offset = 0

    for line_no, raw_line in enumerate(text.splitlines(keepends=True), start=1):
        lines.append((line_no, offset, raw_line.rstrip("\r\n")))
        offset += len(raw_line)

    return lines


# 제목 번호 바로 위의 실제 제목 줄 탐색
def _previous_title_line(lines: list[tuple[int, int, str]], code_line_index: int) -> str:
    for _, _, line in reversed(lines[:code_line_index]):
        title = line.strip()
        if not title or IMAGE_LINE_RE.fullmatch(title):
            continue
        return title

    return ""


# "1-1" 번호 줄을 기준으로 주요통계집 섹션 제목 목록 생성
def find_key_stat_headings(text: str) -> list[KeyStatHeading]:
    lines = _iter_lines_with_offsets(text)
    headings: list[KeyStatHeading] = []

    for line_index, (_, offset, line) in enumerate(lines):
        code_match = TITLE_CODE_RE.fullmatch(line)
        if code_match is None:
            continue

        title = _previous_title_line(lines, line_index)
        headings.append(
            KeyStatHeading(
                code=code_match.group(1),
                title=title,
                start=offset,
            )
        )

    return headings


# Markdown 표 행 분리. 이스케이프된 파이프(\|)는 셀 구분자로 보지 않는다.
def _split_markdown_row(line: str) -> list[str]:
    stripped = line.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]

    cells: list[str] = []
    current: list[str] = []
    escaped = False

    for char in stripped:
        if escaped:
            current.append(char)
            escaped = False
            continue

        if char == "\\":
            escaped = True
            continue

        if char == "|":
            cells.append("".join(current).strip())
            current = []
            continue

        current.append(char)

    if escaped:
        current.append("\\")

    cells.append("".join(current).strip())
    return cells


# Markdown 표 구분선 행 여부
def _is_markdown_separator(line: str) -> bool:
    if MARKDOWN_TABLE_ROW_RE.fullmatch(line) is None:
        return False

    cells = _split_markdown_row(line)
    return bool(cells) and all(
        MARKDOWN_SEPARATOR_CELL_RE.fullmatch(cell.replace(" ", "")) is not None
        for cell in cells
    )


# Markdown 표 시작 행 여부
def _is_markdown_table_start(lines: list[tuple[int, int, str]], index: int) -> bool:
    if index + 1 >= len(lines):
        return False

    line = lines[index][2]
    next_line = lines[index + 1][2]
    return (
        MARKDOWN_TABLE_ROW_RE.fullmatch(line) is not None
        and _is_markdown_separator(next_line)
    )


# HTML table 범위 안에 있는 Markdown 후보는 제외
def _is_inside_ranges(position: int, ranges: list[tuple[int, int]]) -> bool:
    return any(start <= position < end for start, end in ranges)


# Markdown 파이프 표 블록 탐색
def _find_markdown_table_blocks(
    text: str,
    html_ranges: list[tuple[int, int]],
) -> list[KeyStatTableMatch]:
    lines = _iter_lines_with_offsets(text)
    matches: list[KeyStatTableMatch] = []
    line_index = 0

    while line_index < len(lines):
        line_no, start, _ = lines[line_index]
        if _is_inside_ranges(start, html_ranges) or not _is_markdown_table_start(
            lines,
            line_index,
        ):
            line_index += 1
            continue

        end_index = line_index + 2
        while end_index < len(lines):
            _, _, line = lines[end_index]
            if MARKDOWN_TABLE_ROW_RE.fullmatch(line) is None:
                break
            end_index += 1

        end = (
            lines[end_index][1]
            if end_index < len(lines)
            else len(text)
        )
        matches.append(
            KeyStatTableMatch(
                heading=None,
                source=text[start:end].strip(),
                table_type="markdown",
                start=start,
                line_no=line_no,
                index=0,
            )
        )
        line_index = end_index

    return matches


# HTML table 블록 탐색
def _find_html_table_blocks(text: str) -> list[KeyStatTableMatch]:
    matches: list[KeyStatTableMatch] = []

    for table_match in TABLE_RE.finditer(text):
        line_no = text.count("\n", 0, table_match.start()) + 1
        matches.append(
            KeyStatTableMatch(
                heading=None,
                source=table_match.group(0),
                table_type="html",
                start=table_match.start(),
                line_no=line_no,
                index=0,
            )
        )

    return matches


# 표와 가장 가까운 앞쪽 주요통계집 제목 매칭
def find_tables_with_headings(text: str) -> list[KeyStatTableMatch]:
    headings = find_key_stat_headings(text)
    html_tables = _find_html_table_blocks(text)
    html_ranges = [
        (table.start, table.start + len(table.source))
        for table in html_tables
    ]
    markdown_tables = _find_markdown_table_blocks(text, html_ranges)
    tables = sorted([*html_tables, *markdown_tables], key=lambda table: table.start)
    matches: list[KeyStatTableMatch] = []

    for index, table in enumerate(tables, start=1):
        heading = next(
            (
                candidate
                for candidate in reversed(headings)
                if candidate.start < table.start
            ),
            None,
        )
        matches.append(
            KeyStatTableMatch(
                heading=heading,
                source=table.source,
                table_type=table.table_type,
                start=table.start,
                line_no=table.line_no,
                index=index,
            )
        )

    return matches


# Markdown 표 셀 텍스트 정리
def _clean_markdown_cell(value: str) -> str:
    value = MARKDOWN_ESCAPE_RE.sub(r"\1", value)
    value = HTML_BR_RE.sub("\n", value)
    value = html.unescape(value)
    lines = [
        re.sub(r"[ \t]+", " ", line).strip()
        for line in value.split("\n")
    ]
    return "\n".join(lines).strip()


# Markdown 파이프 표를 Excel용 그리드로 변환
def parse_markdown_table_to_grid(table_markdown: str) -> ParsedGrid:
    raw_rows = [
        line
        for line in table_markdown.splitlines()
        if MARKDOWN_TABLE_ROW_RE.fullmatch(line)
    ]
    if len(raw_rows) < 2 or not _is_markdown_separator(raw_rows[1]):
        return [], [], []

    rows = [
        [_clean_markdown_cell(cell) for cell in _split_markdown_row(row)]
        for index, row in enumerate(raw_rows)
        if index != 1
    ]
    max_cols = max((len(row) for row in rows), default=0)
    grid = [row + [""] * (max_cols - len(row)) for row in rows]
    is_header = [index == 0 for index in range(len(grid))]

    return grid, [], is_header


# 출력 파일 기본 이름 생성
def _build_base_name(table: KeyStatTableMatch) -> str:
    if table.heading is None:
        return f"table_{table.index:03d}"

    return sanitize_filename(table.heading.file_title) or f"table_{table.index:03d}"


# 표 종류별 파싱
def _parse_table(table: KeyStatTableMatch) -> ParsedGrid:
    if table.table_type == "html":
        return parse_table_to_grid(table.source)
    if table.table_type == "markdown":
        return parse_markdown_table_to_grid(table.source)
    raise ValueError(f"지원하지 않는 표 형식입니다: {table.table_type}")


# 주요통계집 마크다운의 표를 Excel로 변환
def convert_md_to_excel(md_path: Path, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    text = md_path.read_text(encoding="utf-8")

    used_names: dict[str, int] = {}
    saved_paths: list[Path] = []

    for table in find_tables_with_headings(text):
        file_stem = unique_filename(_build_base_name(table), used_names)
        output_path = output_dir / f"{file_stem}.xlsx"

        try:
            grid, merges, is_header = _parse_table(table)
            if not grid:
                continue

            write_excel(grid, merges, is_header, output_path)
            saved_paths.append(output_path)
        except Exception as error:
            print(
                f"  [error] line {table.line_no} "
                f"({file_stem}, {table.table_type}): {error}"
            )

    return saved_paths


# CLI 실행 진입점
def main() -> None:
    if len(sys.argv) < 3:
        print("Usage: python extract_tables_keystat.py <input.md> <output_dir>")
        sys.exit(1)

    saved = convert_md_to_excel(Path(sys.argv[1]), Path(sys.argv[2]))
    print(f"저장: {len(saved)}개")
    print(f"출력 폴더: {sys.argv[2]}")


if __name__ == "__main__":
    main()
