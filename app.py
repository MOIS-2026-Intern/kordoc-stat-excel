"""
업로드한 hwp/pdf 등을 kordoc로 마크다운으로 변환한 뒤,
그 안의 <table>들을 엑셀(.xlsx)로 추출해 zip으로 묶어 다운로드시키는 Streamlit 앱.
"""

import io
import subprocess
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from extract_tables import convert_md_to_excel

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


# 파일 업로드
uploaded_files = st.file_uploader("파일 업로드", accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        # 파일마다 임시 디렉토리에서 (원본 저장 → md 변환 → xlsx 추출) 파이프라인 돌림
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            input_path = tmp / file.name
            output_md = tmp / "output.md"
            xlsx_dir = tmp / "xlsx"

            input_path.write_bytes(file.read())
            run_kordoc(input_path, output_md)
            xlsx_paths = convert_md_to_excel(output_md, xlsx_dir)

            zip_bytes = build_zip_bytes(xlsx_paths)

        base_name = Path(file.name).stem
        st.download_button(
            f"{base_name} 엑셀 다운로드",
            zip_bytes,
            file_name=f"{base_name}.zip",
            mime="application/zip",
            key=file.name,
        )
