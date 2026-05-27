#!/usr/bin/env python3
# 마크다운 문서의 HTML/Markdown 표 탐색, Markdown 표 파싱, 표별 저장 에러 처리 모듈

from __future__ import annotations

import html
import re
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Callable, Iterable, Protocol, TypeAlias, TypeVar

from table_utils import (
    MergeRange,
    TABLE_RE,
    parse_table_to_grid,
    unique_filename,
    write_excel,
)

ParsedGrid: TypeAlias = tuple[list[list[str]], list[MergeRange], list[bool]]

MARKDOWN_TABLE_ROW_RE = re.compile(r"^\s*\|.*\|\s*$")
MARKDOWN_SEPARATOR_CELL_RE = re.compile(r"^:?-{3,}:?$")
HTML_BR_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)
MARKDOWN_ESCAPE_RE = re.compile(r"\\([\\`*_{}\[\]()#+\-.!|~])")

NO_TABLES_FOUND_MESSAGE = "표가 발견되지 않았습니다. 표가 포함된 파일인지 확인해 주세요."

T = TypeVar("T")
R = TypeVar("R")


class NoTablesFoundError(ValueError):
    pass


class TableExportItem(Protocol):
    source: str
    table_type: str
    line_no: int
    index: int


TableExportT = TypeVar("TableExportT", bound=TableExportItem)


@dataclass(frozen=True)
class TableBlock:
    source: str
    table_type: str
    start: int
    line_no: int
    index: int


# 파일 위치와 줄 번호 계산용 라인 정보 생성
def iter_lines_with_offsets(text: str) -> list[tuple[int, int, str]]:
    lines: list[tuple[int, int, str]] = []
    offset = 0

    for line_no, raw_line in enumerate(text.splitlines(keepends=True), start=1):
        lines.append((line_no, offset, raw_line.rstrip("\r\n")))
        offset += len(raw_line)

    return lines


# 표와 가장 가까운 앞쪽 제목 매칭
def find_nearest_heading(
    headings: list[tuple[int, T]],
    table_start: int,
) -> T | None:
    return next(
        (
            heading
            for heading_start, heading in reversed(headings)
            if heading_start < table_start
        ),
        None,
    )


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


# HTML table 블록 탐색
def find_html_table_blocks(text: str) -> list[TableBlock]:
    matches: list[TableBlock] = []

    for table_match in TABLE_RE.finditer(text):
        line_no = text.count("\n", 0, table_match.start()) + 1
        matches.append(
            TableBlock(
                source=table_match.group(0),
                table_type="html",
                start=table_match.start(),
                line_no=line_no,
                index=0,
            )
        )

    return matches


# Markdown 파이프 표 블록 탐색
def find_markdown_table_blocks(
    text: str,
    html_ranges: list[tuple[int, int]] | None = None,
) -> list[TableBlock]:
    lines = iter_lines_with_offsets(text)
    ranges = html_ranges or []
    matches: list[TableBlock] = []
    line_index = 0

    while line_index < len(lines):
        line_no, start, _ = lines[line_index]
        if _is_inside_ranges(start, ranges) or not _is_markdown_table_start(
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

        end = lines[end_index][1] if end_index < len(lines) else len(text)
        matches.append(
            TableBlock(
                source=text[start:end].strip(),
                table_type="markdown",
                start=start,
                line_no=line_no,
                index=0,
            )
        )
        line_index = end_index

    return matches


# 문서 안의 HTML/Markdown 표 블록을 문서 순서대로 탐색
def find_table_blocks(text: str) -> list[TableBlock]:
    html_tables = find_html_table_blocks(text)
    html_ranges = [
        (table.start, table.start + len(table.source))
        for table in html_tables
    ]
    markdown_tables = find_markdown_table_blocks(text, html_ranges)
    tables = sorted([*html_tables, *markdown_tables], key=lambda table: table.start)

    return [
        replace(table, index=index)
        for index, table in enumerate(tables, start=1)
    ]


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


# 표 종류별 파싱
def parse_table_block(table: TableExportItem) -> ParsedGrid:
    if table.table_type == "html":
        return parse_table_to_grid(table.source)
    if table.table_type == "markdown":
        return parse_markdown_table_to_grid(table.source)
    raise ValueError(f"지원하지 않는 표 형식입니다: {table.table_type}")


# 표별 변환 에러 출력
def print_table_error(
    table: TableExportItem,
    table_name: str,
    error: Exception,
) -> None:
    print(
        f"  [error] line {table.line_no} "
        f"({table_name}, {table.table_type}): {error}"
    )


# 표 하나를 변환/저장할 때 공통 에러 처리 적용
def export_table_safely(
    table: TableExportItem,
    table_name: str,
    action: Callable[[], R],
) -> R | None:
    try:
        return action()
    except Exception as error:
        print_table_error(table, table_name, error)
        return None


# 표마다 별도 Excel 파일로 저장
def write_tables_to_excel_files(
    tables: Iterable[TableExportT],
    output_dir: Path,
    build_base_name: Callable[[TableExportT], str],
) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    used_names: dict[str, int] = {}
    saved_paths: list[Path] = []

    for table in tables:
        base_name = build_base_name(table) or f"table_{table.index:03d}"
        file_stem = unique_filename(base_name, used_names)
        output_path = output_dir / f"{file_stem}.xlsx"

        def _save_table() -> Path | None:
            grid, merges, is_header = parse_table_block(table)
            if not grid:
                return None

            write_excel(grid, merges, is_header, output_path)
            return output_path

        saved_path = export_table_safely(table, file_stem, _save_table)
        if saved_path is not None:
            saved_paths.append(saved_path)

    return saved_paths
