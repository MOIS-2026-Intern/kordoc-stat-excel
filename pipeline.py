"""
업로드된 hwp/pdf 등을 kordoc로 마크다운으로 변환한 뒤,
그 안의 <table>들을 엑셀(.xlsx)로 추출해 zip 바이트로 묶어주는 백엔드 파이프라인.
"""

import io
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Callable, Optional

from extract_tables import convert_md_to_excel
from extract_tables_generic import convert_md_to_excel as convert_md_to_excel_generic

# (percent: 0~100, message: str) 를 받는 진행률 콜백 타입
ProgressCallback = Callable[[int, str], None]


# kordoc CLI를 실행해 입력 파일을 마크다운으로 변환
def run_kordoc(input_path: Path, output_path: Path) -> None:
    subprocess.run(
        ["npx", "-y", "kordoc", str(input_path), "-o", str(output_path)],
        check=True,
    )


# 엑셀 파일 목록을 메모리상의 zip 바이트로 묶어 반환
def build_zip_bytes(xlsx_paths: list[Path]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for p in xlsx_paths:
            zf.write(p, arcname=p.name)
    return buf.getvalue()


# 업로드 파일 바이트를 받아 (md 변환 → xlsx 추출 → zip 묶기)까지 한 번에 처리
# on_progress가 주어지면 단계별로 (percent, message)를 호출
def convert_upload_to_zip(
    filename: str,
    data: bytes,
    on_progress: Optional[ProgressCallback] = None,
) -> bytes:
    def report(percent: int, message: str) -> None:
        if on_progress is not None:
            on_progress(percent, message)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        input_path = tmp / filename
        output_md = tmp / "output.md"
        xlsx_dir = tmp / "xlsx"

        report(5, "파일 저장 중…")
        input_path.write_bytes(data)

        report(15, "마크다운 변환 중… (시간이 좀 걸려요)")
        run_kordoc(input_path, output_md)

        report(75, "표를 엑셀로 추출 중…")
        xlsx_paths = convert_md_to_excel_generic(output_md, xlsx_dir)

        report(95, "zip 압축 중…")
        result = build_zip_bytes(xlsx_paths)

        report(100, "완료")
        return result
