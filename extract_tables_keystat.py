#!/usr/bin/env python3
# 주요통계집 마크다운의 HTML/Markdown 표를 Excel 파일로 저장
# "1-1", "참고-1" 같은 번호 줄과 그 직전 제목 줄을 표 제목으로 사용

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path

from table_extractor import (
    find_nearest_heading,
    find_table_blocks,
    iter_lines_with_offsets,
    write_tables_to_excel_files,
)
from table_utils import sanitize_filename

# 주요통계집 본문 제목 번호 탐색
TITLE_CODE_RE = re.compile(r"^\s*((?:\d+-\d+)|(?:참고-\d+))\.?\s*$")
IMAGE_LINE_RE = re.compile(r"^!\[.*?\]\(.*?\)$")


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


# 제목 번호 바로 위의 실제 제목 줄 탐색
def _previous_title_line(lines: list[tuple[int, int, str]], code_line_index: int) -> str:
    for _, _, line in reversed(lines[:code_line_index]):
        title = line.strip()
        if not title or IMAGE_LINE_RE.fullmatch(title):
            continue
        return title

    return ""


# "1-1", "참고-1" 번호 줄을 기준으로 주요통계집 섹션 제목 목록 생성
def find_key_stat_headings(text: str) -> list[KeyStatHeading]:
    lines = iter_lines_with_offsets(text)
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


# 표와 가장 가까운 앞쪽 주요통계집 제목 매칭
def find_tables_with_headings(text: str) -> list[KeyStatTableMatch]:
    headings = [
        (heading.start, heading)
        for heading in find_key_stat_headings(text)
    ]
    matches: list[KeyStatTableMatch] = []

    for table in find_table_blocks(text):
        matches.append(
            KeyStatTableMatch(
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
def _build_base_name(table: KeyStatTableMatch) -> str:
    if table.heading is None:
        return f"table_{table.index:03d}"

    return sanitize_filename(table.heading.file_title) or f"table_{table.index:03d}"


# 주요통계집 마크다운의 표를 Excel로 변환
def convert_md_to_excel(md_path: Path, output_dir: Path) -> list[Path]:
    text = md_path.read_text(encoding="utf-8")
    return write_tables_to_excel_files(
        find_tables_with_headings(text),
        output_dir,
        _build_base_name,
    )


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
