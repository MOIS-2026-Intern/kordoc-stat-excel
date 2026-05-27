#!/usr/bin/env python3
# 통계연보 마크다운의 HTML/Markdown 표를 Excel 파일로 저장
# 공용 표 탐색/파싱 함수는 table_extractor과 table_utils에서 재사용

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook

from table_extractor import (
    export_table_safely,
    find_nearest_heading,
    find_table_blocks,
    parse_table_block,
)
from table_utils import (
    sanitize_filename,
    unique_filename,
    unique_sheet_title,
    write_table_to_worksheet,
)

# 통계연보 코드형 헤딩 탐색
HEADING_RE = re.compile(r"^(\d+(?:-\d+)+)\t(.+)$", re.MULTILINE)

MAX_SHEETS_PER_WORKBOOK = 100


@dataclass(frozen=True)
class TableMatch:
    heading: str | None
    source: str
    table_type: str
    start: int
    line_no: int
    index: int


# 통계연보 코드형 헤딩 목록 생성
def find_statistics_yearbook_headings(text: str) -> list[tuple[int, str]]:
    return [(match.start(), match.group(0)) for match in HEADING_RE.finditer(text)]


# 표와 가장 가까운 앞쪽 헤딩 매칭
def find_tables_with_headings(text: str) -> list[TableMatch]:
    headings = find_statistics_yearbook_headings(text)
    matches: list[TableMatch] = []

    for table in find_table_blocks(text):
        matches.append(
            TableMatch(
                heading=find_nearest_heading(headings, table.start),
                source=table.source,
                table_type=table.table_type,
                start=table.start,
                line_no=table.line_no,
                index=table.index,
            )
        )

    return matches


# 100개 시트 단위의 출력 파일 경로 생성
def _workbook_output_path(md_path: Path, output_dir: Path, index: int) -> Path:
    base_name = sanitize_filename(md_path.stem) or "tables"
    return output_dir / f"{base_name}_{index}.xlsx"


# 통계연보 마크다운의 표를 Excel로 변환
def convert_md_to_excel(md_path: Path, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    text = md_path.read_text(encoding="utf-8")

    used_table_names: dict[str, int] = {}
    used_sheet_titles: set[str] = set()
    saved_paths: list[Path] = []
    workbook_index = 1
    sheets_in_workbook = 0
    workbook = Workbook()
    workbook.remove(workbook.active)

    for table in find_tables_with_headings(text):
        if table.heading is None:
            continue

        base_name = sanitize_filename(table.heading)
        table_name = unique_filename(base_name, used_table_names)

        def _append_table() -> bool:
            nonlocal workbook
            nonlocal workbook_index
            nonlocal sheets_in_workbook
            nonlocal used_sheet_titles

            grid, merges, is_header = parse_table_block(table)
            if not grid:
                return False

            if sheets_in_workbook == MAX_SHEETS_PER_WORKBOOK:
                output_path = _workbook_output_path(md_path, output_dir, workbook_index)
                workbook.save(output_path)
                saved_paths.append(output_path)

                workbook_index += 1
                sheets_in_workbook = 0
                used_sheet_titles = set()
                workbook = Workbook()
                workbook.remove(workbook.active)

            sheet_name = unique_sheet_title(table_name, used_sheet_titles)
            write_table_to_worksheet(workbook, sheet_name, grid, merges, is_header)
            sheets_in_workbook += 1
            return True

        export_table_safely(table, table_name, _append_table)

    if sheets_in_workbook:
        output_path = _workbook_output_path(md_path, output_dir, workbook_index)
        workbook.save(output_path)
        saved_paths.append(output_path)

    return saved_paths


# CLI 실행 진입점
def main() -> None:
    if len(sys.argv) < 3:
        print("Usage: python extract_tables.py <input.md> <output_dir>")
        sys.exit(1)

    output_dir = Path(sys.argv[2])
    saved = convert_md_to_excel(Path(sys.argv[1]), output_dir)
    print(f"저장: {len(saved)}개")
    print(f"출력 폴더: {output_dir}")


if __name__ == "__main__":
    main()
