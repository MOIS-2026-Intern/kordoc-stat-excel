#!/usr/bin/env python3
"""
마크다운 안의 <table>들을 추출해 엑셀(.xlsx)로 저장
파일명은 표 바로 위에 있는 코드형 헤딩(예: '1-1-1-2\t정부조직 변천 ...')에서 가져옴
"""

import re
from pathlib import Path

from bs4 import BeautifulSoup, NavigableString
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

INPUT_FILE = Path("/Users/song/dev/mois/kordoc/통계연보파싱.md")
OUTPUT_DIR = Path("/Users/song/dev/mois/kordoc/statoutput")

# '1-1-1-2\t정부조직 변천 ...' 형태의 코드형 헤딩
HEADING_RE = re.compile(r"^(\d+(?:-\d+)+)\t(.+)$", re.MULTILINE)
TABLE_RE = re.compile(r"<table\b.*?</table>", re.IGNORECASE | re.DOTALL)

# 파일 시스템에서 쓸 수 없는 문자/공백을 정리하고 길이 제한
def sanitize_filename(text: str) -> str:
    text = text.replace("\t", " ")
    text = re.sub(r'[\/\\:*?"<>|]', "_", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:180].rstrip()


# 셀 내부 텍스트를 뽑되, <br>은 줄바꿈으로 변환
def get_cell_text(cell) -> str:
    parts = []
    for elem in cell.descendants:
        if isinstance(elem, NavigableString):
            parts.append(str(elem))
        elif elem.name == "br":
            parts.append("\n")
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in "".join(parts).split("\n")]
    return "\n".join(lines).strip()


# rowspan/colspan 속성을 정수로 변환 (잘못된 값이면 1).
def _intspan(cell, attr: str) -> int:
    try:
        return max(1, int(cell.get(attr, 1)))
    except (TypeError, ValueError):
        return 1


"""
    rowspan/colspan을 반영한 2D 그리드로 변환
    반환: (grid, merges, is_header)
      - grid:      List[List[str]]
      - merges:    [(r1, c1, r2, c2)]  (0-based, inclusive)
      - is_header: 행별 헤더 여부 (해당 행이 <th>를 포함하면 True)
    """
def parse_table_to_grid(table_html: str):
    soup = BeautifulSoup(table_html, "html.parser")
    table = soup.find("table")
    if not table:
        return [], [], []

    rows = table.find_all("tr")
    placed = []           # (row, col, text) — 실제 셀이 놓인 위치
    occupied = set()      # 다른 셀의 span으로 점유된 좌표 (해당 칸엔 새 셀을 둘 수 없음)
    merges = []
    is_header = [False] * len(rows)
    max_rows = len(rows)
    max_cols = 0

    for r, tr in enumerate(rows):
        c = 0
        for cell in tr.find_all(["th", "td"]):
            if cell.name == "th":
                is_header[r] = True
            # 이미 윗 행에서 rowspan으로 점유된 칸은 건너뜀
            while (r, c) in occupied:
                c += 1

            rowspan = _intspan(cell, "rowspan")
            colspan = _intspan(cell, "colspan")
            placed.append((r, c, get_cell_text(cell)))

            for rr in range(r, r + rowspan):
                for cc in range(c, c + colspan):
                    if (rr, cc) != (r, c):
                        occupied.add((rr, cc))

            if rowspan > 1 or colspan > 1:
                merges.append((r, c, r + rowspan - 1, c + colspan - 1))

            max_rows = max(max_rows, r + rowspan)
            max_cols = max(max_cols, c + colspan)
            c += colspan

    # rowspan이 마지막 행을 넘어가는 비정상 케이스를 위해 is_header 길이를 맞춤.
    is_header.extend([False] * (max_rows - len(is_header)))

    grid = [[""] * max_cols for _ in range(max_rows)]
    for r, c, text in placed:
        grid[r][c] = text
    return grid, merges, is_header


# 그리드/병합 정보를 받아 xlsx 파일로 저장
def write_excel(grid, merges, is_header, output_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    bold = Font(bold=True)
    align = Alignment(wrap_text=True, vertical="center", horizontal="center")

    # 셀 값/스타일 채우기 (openpyxl은 1-based)
    for r, row in enumerate(grid, start=1):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.alignment = align
            if is_header[r - 1]:
                cell.font = bold

    # 병합 적용
    for (sr, sc, er, ec) in merges:
        if sr == er and sc == ec:
            continue
        ws.merge_cells(
            start_row=sr + 1, start_column=sc + 1,
            end_row=er + 1, end_column=ec + 1,
        )

    # 컬럼 너비: 가장 긴 줄 기준 + 여유, 단 [10, 50] 범위로 클램프
    if grid:
        for c in range(len(grid[0])):
            longest = max(
                (len(ln) for row in grid if c < len(row) for ln in row[c].split("\n")),
                default=8,
            )
            ws.column_dimensions[get_column_letter(c + 1)].width = min(max(longest + 2, 10), 50)

    wb.save(output_path)


"""
    문서에서 <table>...</table> 블록을 모두 찾고,
    각 표의 시작 위치보다 앞에 있는 코드형 헤딩 중 가장 가까운 것을 짝지어 반환
    반환: [(heading_or_None, table_html, start_line)]
    """
def find_tables_with_headings(text: str):
    headings = [(m.start(), m.group(0)) for m in HEADING_RE.finditer(text)]
    out = []
    for tm in TABLE_RE.finditer(text):
        # 표 시작점 앞의 헤딩 중 가장 가까운 것을 찾음
        heading = next(
            (h_text for h_start, h_text in reversed(headings) if h_start < tm.start()),
            None,
        )
        line_no = text.count("\n", 0, tm.start()) + 1
        out.append((heading, tm.group(0), line_no))
    return out


# md 파일에서 표를 모두 추출해 output_dir에 xlsx로 저장하고 경로 목록을 돌려줌
def convert_md_to_excel(md_path: Path, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    text = md_path.read_text(encoding="utf-8")

    used_names: dict[str, int] = {}
    saved_paths: list[Path] = []

    for heading, table_html, line_no in find_tables_with_headings(text):
        if heading is None:
            continue

        base = sanitize_filename(heading)
        # 같은 헤딩이 여러 번 나오면 _2, _3, ... 으로 충돌 회피
        used_names[base] = used_names.get(base, 0) + 1
        name = base if used_names[base] == 1 else f"{base}_{used_names[base]}"
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
    saved = convert_md_to_excel(INPUT_FILE, OUTPUT_DIR)
    print(f"저장: {len(saved)}개")
    print(f"출력 폴더: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
