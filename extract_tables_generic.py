#!/usr/bin/env python3
# 범용 마크다운의 HTML/Markdown 표를 Excel 파일로 저장
# 일반 마크다운 헤딩을 표 제목으로 사용

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path

from table_extractor import (
    NO_TABLES_FOUND_MESSAGE,
    NoTablesFoundError,
    find_nearest_heading,
    find_table_blocks,
    write_tables_to_excel_files,
)
from table_utils import sanitize_filename

# ATX 스타일 마크다운 헤딩 탐색
MD_HEADING_RE = re.compile(r"^#{1,6}\s+(.+?)\s*#*\s*$", re.MULTILINE)


@dataclass(frozen=True)
class GenericTableMatch:
    heading: str | None
    source: str
    table_type: str
    start: int
    line_no: int
    index: int

    @property
    def html(self) -> str:
        return self.source


# 일반 마크다운 헤딩 목록 생성
def find_generic_headings(text: str) -> list[tuple[int, str]]:
    return [
        (match.start(), match.group(1).strip())
        for match in MD_HEADING_RE.finditer(text)
    ]


# 표와 가장 가까운 앞쪽 마크다운 헤딩 매칭
def find_tables_with_headings(text: str) -> list[GenericTableMatch]:
    headings = find_generic_headings(text)
    matches: list[GenericTableMatch] = []

    for table in find_table_blocks(text):
        matches.append(
            GenericTableMatch(
                heading=find_nearest_heading(headings, table.start),
                source=table.source,
                table_type=table.table_type,
                start=table.start,
                line_no=table.line_no,
                index=table.index,
            )
        )

    return matches


# 출력 파일 기본 이름 생성
def _build_base_name(md_path: Path, table: GenericTableMatch) -> str:
    prefix = sanitize_filename(md_path.stem)
    table_name = sanitize_filename(table.heading) if table.heading else f"table_{table.index:03d}"
    return f"{prefix}_{table_name}" if prefix else table_name


# 범용 마크다운의 표를 Excel로 변환
def convert_md_to_excel(md_path: Path, output_dir: Path) -> list[Path]:
    text = md_path.read_text(encoding="utf-8")
    tables = find_tables_with_headings(text)

    if not tables:
        raise NoTablesFoundError(NO_TABLES_FOUND_MESSAGE)

    saved_paths = write_tables_to_excel_files(
        tables,
        output_dir,
        lambda table: _build_base_name(md_path, table),
    )

    if not saved_paths:
        raise NoTablesFoundError(NO_TABLES_FOUND_MESSAGE)

    return saved_paths


# CLI 실행 진입점
def main() -> None:
    if len(sys.argv) < 3:
        print("Usage: python extract_tables_generic.py <input.md> <output_dir>")
        sys.exit(1)

    saved = convert_md_to_excel(Path(sys.argv[1]), Path(sys.argv[2]))
    print(f"저장: {len(saved)}개")
    print(f"출력 폴더: {sys.argv[2]}")


if __name__ == "__main__":
    main()
