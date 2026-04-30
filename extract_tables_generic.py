#!/usr/bin/env python3
"""
범용 마크다운 → 엑셀 표 추출기
extract_tables.py와 동일한 convert_md_to_excel 시그니처를 제공한다.
pipeline.py의 import만 바꿔주면 통계연보 형식이 아닌 다른 문서에도 동작한다.

차이점:
- 통계연보 전용 코드형 헤딩(`1-1-1-2\\t제목`) 대신 일반 마크다운 헤딩(`#`, `##`...)을 사용
- 헤딩이 없는 표도 `table_001` 형태의 일련번호로 저장
- 같은 출력 폴더에 여러 입력의 결과가 섞여도 구분되도록 입력 파일명을 prefix로 붙임
"""

import re
import sys
from pathlib import Path

from extract_tables import (
    TABLE_RE,
    parse_table_to_grid,
    sanitize_filename,
    write_excel,
)

# `# 제목`, `## 제목` 등 ATX 스타일 마크다운 헤딩
MD_HEADING_RE = re.compile(r"^#{1,6}\s+(.+?)\s*#*\s*$", re.MULTILINE)


# 각 <table>을 가장 가까운 앞 헤딩(없으면 None)과 짝지어 반환
def find_tables_with_headings(text: str):
    headings = [(m.start(), m.group(1).strip()) for m in MD_HEADING_RE.finditer(text)]
    out = []
    for tm in TABLE_RE.finditer(text):
        heading = next(
            (h_text for h_start, h_text in reversed(headings) if h_start < tm.start()),
            None,
        )
        line_no = text.count("\n", 0, tm.start()) + 1
        out.append((heading, tm.group(0), line_no))
    return out


# md 안의 모든 <table>을 xlsx로 저장하고 경로 목록을 돌려줌
def convert_md_to_excel(md_path: Path, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    text = md_path.read_text(encoding="utf-8")
    prefix = sanitize_filename(md_path.stem)

    used_names: dict[str, int] = {}
    saved_paths: list[Path] = []

    for idx, (heading, table_html, line_no) in enumerate(
        find_tables_with_headings(text), start=1
    ):
        base = sanitize_filename(heading) if heading else f"table_{idx:03d}"
        name = f"{prefix}_{base}" if prefix else base

        used_names[name] = used_names.get(name, 0) + 1
        if used_names[name] > 1:
            name = f"{name}_{used_names[name]}"
        out_path = output_dir / f"{name}.xlsx"

        try:
            grid, merges, is_header = parse_table_to_grid(table_html)
            if not grid:
                continue
            write_excel(grid, merges, is_header, out_path)
            saved_paths.append(out_path)
        except Exception as e:
            print(f"  [error] line {line_no} ({name}): {e}")

    return saved_paths


def main():
    if len(sys.argv) < 3:
        print("Usage: python extract_tables_generic.py <input.md> <output_dir>")
        sys.exit(1)
    saved = convert_md_to_excel(Path(sys.argv[1]), Path(sys.argv[2]))
    print(f"저장: {len(saved)}개")
    print(f"출력 폴더: {sys.argv[2]}")


if __name__ == "__main__":
    main()
