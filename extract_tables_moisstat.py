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


@dataclass(frozen=True)
class TableMatch:
    heading: str | None
    source: str
    table_type: str
    start: int
    line_no: int
    index: int


@dataclass
class WorkbookGroup:
    output_path: Path
    workbook: Workbook
    used_sheet_titles: set[str]
    sheet_count: int = 0


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


# 통계연보 번호 앞 그룹 이름 생성: 1-1-1, 1-2-1 -> output_1.xlsx
def _build_group_base_name(md_path: Path, heading: str) -> str:
    code_match = HEADING_RE.fullmatch(heading)
    if code_match is None:
        return sanitize_filename(md_path.stem) or "tables"

    group_name = code_match.group(1).split("-", maxsplit=1)[0]
    file_prefix = sanitize_filename(md_path.stem) or "tables"
    return sanitize_filename(f"{file_prefix}_{group_name}") or file_prefix


def _create_workbook() -> Workbook:
    workbook = Workbook()
    workbook.remove(workbook.active)
    return workbook


# 통계연보 마크다운의 표를 Excel로 변환
def convert_md_to_excel(md_path: Path, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    text = md_path.read_text(encoding="utf-8")

    used_table_names: dict[str, int] = {}
    used_file_names: dict[str, int] = {}
    workbook_groups: dict[str, WorkbookGroup] = {}

    for table in find_tables_with_headings(text):
        if table.heading is None:
            continue

        base_name = sanitize_filename(table.heading)
        table_name = unique_filename(base_name, used_table_names)
        group_base_name = _build_group_base_name(md_path, table.heading)

        def _append_table() -> bool:
            grid, merges, is_header = parse_table_block(table)
            if not grid:
                return False

            group = workbook_groups.get(group_base_name)
            if group is None:
                file_stem = unique_filename(group_base_name, used_file_names)
                group = WorkbookGroup(
                    output_path=output_dir / f"{file_stem}.xlsx",
                    workbook=_create_workbook(),
                    used_sheet_titles=set(),
                )
                workbook_groups[group_base_name] = group

            sheet_name = unique_sheet_title(table_name, group.used_sheet_titles)
            write_table_to_worksheet(group.workbook, sheet_name, grid, merges, is_header)
            group.sheet_count += 1
            return True

        export_table_safely(table, table_name, _append_table)

    saved_paths: list[Path] = []
    for group in workbook_groups.values():
        if group.sheet_count == 0:
            continue
        group.workbook.save(group.output_path)
        saved_paths.append(group.output_path)

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
