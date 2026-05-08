# Streamlit UI 진입점
# 파일 업로드와 zip 다운로드 화면 구성

from __future__ import annotations

from pathlib import Path

import streamlit as st

from pipeline import convert_upload_to_zip

APP_TITLE = "HWP → Excel 변환기"
APP_CAPTION = "아직은 베타 버전입니다. 변환이 안 되는 파일이 있을 수 있어요."
DOWNLOAD_MIME = "application/zip"
LOGO_PATH = Path(__file__).parent / "src" / "logo2.png"


# 상단 제목 영역 렌더링
def render_header() -> None:
    st.title(APP_TITLE)
    st.caption(APP_CAPTION)
    st.divider()


# 하단 안내 영역 렌더링
def render_footer() -> None:
    st.divider()
    st.markdown(
        "<div style='text-align:center; color:#888; font-size:0.85rem; padding:0.5rem 0;'>"
        "© 행정안전부 (MOIS) · 표 추출기"
        "</div>",
        unsafe_allow_html=True,
    )


# 변환 결과 세션 캐시 조회
def _get_conversion_cache() -> dict[str, bytes]:
    return st.session_state.setdefault("converted", {})


# 업로드 파일 변환
def _convert_uploaded_file(file) -> bytes:
    progress_bar = st.progress(0, text=f"{file.name} 준비 중…")

    def update_progress(percent: int, message: str) -> None:
        progress_bar.progress(percent, text=f"{file.name} — {message}")

    try:
        return convert_upload_to_zip(
            filename=file.name,
            data=file.read(),
            on_progress=update_progress,
        )
    finally:
        progress_bar.empty()


# 다운로드 버튼 렌더링
def render_download_button(file, zip_bytes: bytes) -> None:
    base_name = Path(file.name).stem
    st.download_button(
        label=f"{base_name} 엑셀 다운로드",
        data=zip_bytes,
        file_name=f"{base_name}.zip",
        mime=DOWNLOAD_MIME,
        key=file.file_id,
    )


# 파일 업로드 영역 렌더링
def render_uploader() -> None:
    uploaded_files = st.file_uploader(
        "파일 업로드 (hwp, pdf 등)",
        accept_multiple_files=True,
    )
    if not uploaded_files:
        return

    cache = _get_conversion_cache()
    for file in uploaded_files:
        if file.file_id not in cache:
            cache[file.file_id] = _convert_uploaded_file(file)

        render_download_button(file, cache[file.file_id])


# 앱 실행
def main() -> None:
    st.set_page_config(
        page_title="한글 표 추출기",
        page_icon=str(LOGO_PATH),
        layout="centered",
    )
    render_header()
    render_uploader()
    render_footer()


main()
