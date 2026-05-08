#!/usr/bin/env python3
# 범용 마크다운의 HTML 표를 Excel 파일로 저장
# 일반 마크다운 헤딩을 표 제목으로 사용

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path

from table_utils import (
    TABLE_RE,
    parse_table_to_grid,
    sanitize_filename,
    unique_filename,
    write_excel,
)

# ATX 스타일 마크다운 헤딩 탐색
MD_HEADING_RE = re.compile(r"^#{1,6}\s+(.+?)\s*#*\s*$", re.MULTILINE)


@dataclass(frozen=True)
class GenericTableMatch:
    heading: str | None
    html: str
    line_no: int
    index: int


# 표와 가장 가까운 앞쪽 마크다운 헤딩 매칭
def find_tables_with_headings(text: str) -> list[GenericTableMatch]:
    headings = [
        (match.start(), match.group(1).strip())
        for match in MD_HEADING_RE.finditer(text)
    ]
    matches: list[GenericTableMatch] = []

    for index, table_match in enumerate(TABLE_RE.finditer(text), start=1):
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
            GenericTableMatch(
                heading=heading,
                html=table_match.group(0),
                line_no=line_no,
                index=index,
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
    output_dir.mkdir(parents=True, exist_ok=True)
    text = md_path.read_text(encoding="utf-8")

    used_names: dict[str, int] = {}
    saved_paths: list[Path] = []

    for table in find_tables_with_headings(text):
        file_stem = unique_filename(_build_base_name(md_path, table), used_names)
        output_path = output_dir / f"{file_stem}.xlsx"

        try:
            grid, merges, is_header = parse_table_to_grid(table.html)
            if not grid:
                continue

            write_excel(grid, merges, is_header, output_path)
            saved_paths.append(output_path)
        except Exception as error:
            print(f"  [error] line {table.line_no} ({file_stem}): {error}")

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
