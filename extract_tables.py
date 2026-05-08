#!/usr/bin/env python3
# 통계연보 마크다운의 HTML 표를 Excel 파일로 저장
# 공용 표 파싱/Excel 저장 함수는 table_utils에서 재사용

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook

from table_utils import (
    TABLE_RE,
    parse_table_to_grid,
    sanitize_filename,
    unique_filename,
    unique_sheet_title,
    write_table_to_worksheet,
)

INPUT_FILE = Path("/Users/song/dev/mois/kordoc/통계연보파싱.md")
OUTPUT_DIR = Path("/Users/song/dev/mois/kordoc/statoutput")

# 통계연보 코드형 헤딩 탐색
HEADING_RE = re.compile(r"^(\d+(?:-\d+)+)\t(.+)$", re.MULTILINE)

MAX_SHEETS_PER_WORKBOOK = 100


@dataclass(frozen=True)
class TableMatch:
    heading: str | None
    html: str
    line_no: int


# 표와 가장 가까운 앞쪽 헤딩 매칭
def find_tables_with_headings(text: str) -> list[TableMatch]:
    headings = [(match.start(), match.group(0)) for match in HEADING_RE.finditer(text)]
    matches: list[TableMatch] = []

    for table_match in TABLE_RE.finditer(text):
        heading = next(
            (
                heading_text
                for heading_start, heading_text in reversed(headings)
                if heading_start < table_match.start()
            ),
            None,
        )
        line_no = text.count("\n", 0, table_match.start()) + 1
        matches.append(
            TableMatch(
                heading=heading,
                html=table_match.group(0),
                line_no=line_no,
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

    used_names: dict[str, int] = {}
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
        table_name = unique_filename(base_name, used_names)

        try:
            grid, merges, is_header = parse_table_to_grid(table.html)
            if not grid:
                continue

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
        except Exception as error:
            print(f"  [error] line {table.line_no} ({table_name}): {error}")

    if sheets_in_workbook:
        output_path = _workbook_output_path(md_path, output_dir, workbook_index)
        workbook.save(output_path)
        saved_paths.append(output_path)

    return saved_paths


# CLI 실행 진입점
def main() -> None:
    saved = convert_md_to_excel(INPUT_FILE, OUTPUT_DIR)
    print(f"저장: {len(saved)}개")
    print(f"출력 폴더: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
