# 업로드 파일을 Excel 표 묶음 zip으로 변환
# 원본 저장, 마크다운 변환, Excel 추출, zip 생성 순서

from __future__ import annotations

import io
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Callable

from extract_tables_moisstat import convert_md_to_excel as convert_statistics_yearbook_md_to_excel
from extract_tables_keystat import convert_md_to_excel as convert_key_year_statistics_md_to_excel
from extract_tables_generic import convert_md_to_excel as convert_generic_md_to_excel

# 진행률 콜백 타입
ProgressCallback = Callable[[int, str], None]
TableExtractor = Callable[[Path, Path], list[Path]]

STATISTICS_YEARBOOK_EXTRACTOR = "행정안전통계연보 추출기"
KEY_YEAR_STATISTICS_EXTRACTOR = "주요통계집 추출기"
GENERIC_EXTRACTOR = "범용 추출기"
DEFAULT_EXTRACTOR = GENERIC_EXTRACTOR

EXTRACTORS: dict[str, TableExtractor] = {
    STATISTICS_YEARBOOK_EXTRACTOR: convert_statistics_yearbook_md_to_excel,
    KEY_YEAR_STATISTICS_EXTRACTOR: convert_key_year_statistics_md_to_excel,
    GENERIC_EXTRACTOR: convert_generic_md_to_excel,
}


# kordoc CLI 실행
def run_kordoc(input_path: Path, output_path: Path) -> None:
    subprocess.run(
        ["npx", "-y", "kordoc", str(input_path), "-o", str(output_path)],
        check=True,
    )


# Excel 파일 zip bytes 생성
def build_zip_bytes(xlsx_paths: list[Path]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as zip_file:
        for path in xlsx_paths:
            zip_file.write(path, arcname=path.name)
    return buffer.getvalue()


# 진행률 콜백 호출
def _report(
    on_progress: ProgressCallback | None,
    percent: int,
    message: str,
) -> None:
    if on_progress is not None:
        on_progress(percent, message)


# 업로드 파일 bytes를 zip bytes로 변환
def convert_upload_to_zip(
    filename: str,
    data: bytes,
    extractor_name: str = DEFAULT_EXTRACTOR,
    on_progress: ProgressCallback | None = None,
) -> bytes:
    extractor = EXTRACTORS.get(extractor_name)
    if extractor is None:
        raise ValueError(f"지원하지 않는 추출기입니다: {extractor_name}")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        input_path = tmp_path / filename
        markdown_path = tmp_path / "output.md"
        xlsx_dir = tmp_path / "xlsx"

        _report(on_progress, 5, "파일 저장 중…")
        input_path.write_bytes(data)

        _report(on_progress, 15, "마크다운 변환 중… (시간이 좀 걸려요)")
        run_kordoc(input_path, markdown_path)

        _report(on_progress, 75, f"{extractor_name}로 표를 엑셀로 추출 중…")
        xlsx_paths = extractor(markdown_path, xlsx_dir)

        _report(on_progress, 95, "zip 압축 중…")
        result = build_zip_bytes(xlsx_paths)

        _report(on_progress, 100, "완료")
        return result
