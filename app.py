"""
통계연보 표 추출기 Streamlit UI.
"""

from pathlib import Path

import streamlit as st

from pipeline import convert_upload_to_zip

LOGO_PATH = Path(__file__).parent / "src" / "logo2.png"


def render_header() -> None:
    st.title("HWP → Excel 변환기")
    st.caption("아직은 베타 버전입니다. 변환이 안 되는 파일이 있을 수 있어요.")
    st.divider()


def render_footer() -> None:
    st.divider()
    st.markdown(
        "<div style='text-align:center; color:#888; font-size:0.85rem; padding:0.5rem 0;'>"
        "© 행정안전부 (MOIS) · 표 추출기"
        "</div>",
        unsafe_allow_html=True,
    )


def render_uploader() -> None:
    uploaded_files = st.file_uploader(
        "파일 업로드 (hwp, pdf 등)",
        accept_multiple_files=True,
    )
    if not uploaded_files:
        return

    for file in uploaded_files:
        progress_bar = st.progress(0, text=f"{file.name} 준비 중…")

        def on_progress(percent: int, message: str, name: str = file.name) -> None:
            progress_bar.progress(percent, text=f"{name} — {message}")

        zip_bytes = convert_upload_to_zip(
            file.name, file.read(), on_progress=on_progress
        )
        progress_bar.empty()

        base_name = Path(file.name).stem
        st.download_button(
            f"{base_name} 엑셀 다운로드",
            zip_bytes,
            file_name=f"{base_name}.zip",
            mime="application/zip",
            key=file.name,
        )


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
