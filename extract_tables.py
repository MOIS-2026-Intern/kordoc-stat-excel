#!/usr/bin/env python3
"""
통계연보파싱.md 파일에서 <table> 태그 안의 내용을 추출해
table 바로 위에 있는 코드 형식 헤딩(예: '1-1-1-2\t정부조직 변천 ...')을
파일명으로 하여 statoutput 폴더에 엑셀(.xlsx)로 저장한다.
"""

import re
from pathlib import Path

from bs4 import BeautifulSoup, NavigableString
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

INPUT_FILE = Path("/Users/song/dev/mois/kordoc/통계연보파싱.md")
OUTPUT_DIR = Path("/Users/song/dev/mois/kordoc/statoutput")

# '1-1-1-2\t정부조직 변천 ...' 형태
HEADING_RE = re.compile(r"^(\d+(?:-\d+)+)\t(.+)$")


def sanitize_filename(text: str) -> str:
    text = text.replace("\t", " ")
    text = re.sub(r'[\/\\:*?"<>|]', "_", text)
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > 180:
        text = text[:180].rstrip()
    return text


def get_cell_text(cell) -> str:
    parts = []
    for elem in cell.descendants:
        if isinstance(elem, NavigableString):
            parts.append(str(elem))
        elif elem.name == "br":
            parts.append("\n")
    text = "".join(parts)
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in text.split("\n")]
    # 빈 줄 정리하지 않고 그대로 두되, 양 끝만 strip
    return "\n".join(lines).strip()


def parse_table_to_grid(table_html: str):
    """
    rowspan/colspan을 반영하여 2D 그리드와 병합 정보를 만든다.
    반환: (grid: List[List[str]], merges: List[(r1,c1,r2,c2)])  # 0-based inclusive
    """
    soup = BeautifulSoup(table_html, "html.parser")
    table = soup.find("table")
    if not table:
        return [], []

    rows = table.find_all("tr")
    grid_dicts = []  # list of {col: text}
    occupied = set()  # set of (row, col) — 다른 셀의 span으로 점유됨
    merges = []
    is_header = []  # bool per row, row is header if it contains <th>

    for r, tr in enumerate(rows):
        while r >= len(grid_dicts):
            grid_dicts.append({})
            is_header.append(False)
        c = 0
        row_has_th = False
        for cell in tr.find_all(["th", "td"]):
            if cell.name == "th":
                row_has_th = True
            while (r, c) in occupied:
                c += 1
            text = get_cell_text(cell)
            try:
                rowspan = int(cell.get("rowspan", 1))
            except (TypeError, ValueError):
                rowspan = 1
            try:
                colspan = int(cell.get("colspan", 1))
            except (TypeError, ValueError):
                colspan = 1
            rowspan = max(1, rowspan)
            colspan = max(1, colspan)

            grid_dicts[r][c] = text
            for rr in range(r, r + rowspan):
                while rr >= len(grid_dicts):
                    grid_dicts.append({})
                    is_header.append(False)
                for cc in range(c, c + colspan):
                    if (rr, cc) != (r, c):
                        occupied.add((rr, cc))
            if rowspan > 1 or colspan > 1:
                merges.append((r, c, r + rowspan - 1, c + colspan - 1))
            c += colspan
        if row_has_th:
            is_header[r] = True

    max_cols = 0
    for row in grid_dicts:
        if row:
            max_cols = max(max_cols, max(row.keys()) + 1)
    for (rr, cc) in occupied:
        max_cols = max(max_cols, cc + 1)

    grid = []
    for row in grid_dicts:
        line = [""] * max_cols
        for c, t in row.items():
            line[c] = t
        grid.append(line)
    return grid, merges, is_header


def write_excel(grid, merges, is_header, output_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    bold = Font(bold=True)
    align = Alignment(wrap_text=True, vertical="center", horizontal="center")

    for r, row in enumerate(grid, start=1):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.alignment = align
            if r - 1 < len(is_header) and is_header[r - 1]:
                cell.font = bold

    for (sr, sc, er, ec) in merges:
        if sr == er and sc == ec:
            continue
        ws.merge_cells(
            start_row=sr + 1, start_column=sc + 1,
            end_row=er + 1, end_column=ec + 1,
        )

    if grid:
        max_cols = len(grid[0])
        for c in range(max_cols):
            max_len = 8
            for row in grid:
                if c < len(row) and row[c]:
                    for ln in row[c].split("\n"):
                        max_len = max(max_len, len(ln))
            ws.column_dimensions[get_column_letter(c + 1)].width = min(
                max(max_len + 2, 10), 50
            )

    wb.save(output_path)


def find_tables_with_headings(text: str):
    """파일을 줄 단위로 읽으며 <table>...</table> 블록을 찾고
    각 블록 위쪽의 가장 가까운 코드형 헤딩을 매칭한다."""
    lines = text.split("\n")
    out = []  # (heading_or_None, table_html, start_line)
    n = len(lines)
    i = 0
    while i < n:
        line = lines[i]
        if re.search(r"<table\b", line, re.IGNORECASE):
            start = i
            buf = [line]
            while "</table>" not in buf[-1].lower():
                i += 1
                if i >= n:
                    break
                buf.append(lines[i])
            table_html = "\n".join(buf)

            heading = None
            for j in range(start - 1, -1, -1):
                m = HEADING_RE.match(lines[j])
                if m:
                    heading = lines[j]
                    break
            out.append((heading, table_html, start + 1))
        i += 1
    return out


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    text = INPUT_FILE.read_text(encoding="utf-8")

    tables = find_tables_with_headings(text)
    print(f"발견된 <table> 블록 수: {len(tables)}")

    used_names = {}
    saved = 0
    skipped = 0

    for heading, table_html, line_no in tables:
        if heading is None:
            skipped += 1
            print(f"  [skip] line {line_no}: 위쪽에 코드형 헤딩이 없음")
            continue

        base = sanitize_filename(heading)
        if base in used_names:
            used_names[base] += 1
            name = f"{base}_{used_names[base]}"
        else:
            used_names[base] = 1
            name = base

        out_path = OUTPUT_DIR / f"{name}.xlsx"
        try:
            grid, merges, is_header = parse_table_to_grid(table_html)
            if not grid:
                skipped += 1
                print(f"  [skip] line {line_no}: 빈 테이블 ({name})")
                continue
            write_excel(grid, merges, is_header, out_path)
            saved += 1
        except Exception as e:
            skipped += 1
            print(f"  [error] line {line_no} ({name}): {e}")

    print(f"\n저장: {saved}개  /  건너뜀: {skipped}개")
    print(f"출력 폴더: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
