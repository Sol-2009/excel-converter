import streamlit as st
import tempfile
import subprocess
import os
from pathlib import Path

st.set_page_config(page_title="Excel to XLSX 변환기")
st.title("Excel to .xlsx 변환기")
st.caption("LibreOffice를 사용하여 서식을 유지하면서 .xlsx로 변환합니다.")

SUPPORTED_EXTENSIONS = ["xls", "xlsb", "xlsm", "xlt", "xltx", "xltm", "csv"]

st.info(f"지원 형식: {', '.join('.' + ext for ext in SUPPORTED_EXTENSIONS)}")

uploaded_files = st.file_uploader(
    "변환할 Excel 파일을 업로드하세요",
    type=SUPPORTED_EXTENSIONS,
    accept_multiple_files=True,
)


def convert_to_xlsx(input_path, output_dir):
    """LibreOffice headless를 사용하여 파일을 .xlsx로 변환"""
    env = os.environ.copy()
    env["HOME"] = tempfile.mkdtemp()

    result = subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--norestore",
            "--convert-to",
            "xlsx",
            "--outdir",
            output_dir,
            input_path,
        ],
        capture_output=True,
        text=True,
        timeout=120,
        env=env,
    )
    if result.returncode != 0:
        raise RuntimeError(f"변환 실패: {result.stderr}")

    output_path = os.path.join(output_dir, Path(input_path).stem + ".xlsx")
    if not os.path.exists(output_path):
        raise RuntimeError("변환된 파일을 찾을 수 없습니다.")
    return output_path


if "results" not in st.session_state:
    st.session_state.results = []

if uploaded_files and st.button("변환 시작", type="primary"):
    st.session_state.results = []
    progress = st.progress(0)

    for i, uploaded_file in enumerate(uploaded_files):
        file_stem = Path(uploaded_file.name).stem
        file_suffix = Path(uploaded_file.name).suffix

        with st.spinner(f"변환 중: {uploaded_file.name}"):
            tmp_dir = None
            try:
                tmp_dir = tempfile.mkdtemp()
                tmp_in_path = os.path.join(tmp_dir, uploaded_file.name)
                with open(tmp_in_path, "wb") as f:
                    f.write(uploaded_file.read())

                output_path = convert_to_xlsx(tmp_in_path, tmp_dir)

                with open(output_path, "rb") as f:
                    converted_data = f.read()

                st.session_state.results.append(
                    {
                        "name": uploaded_file.name,
                        "stem": file_stem,
                        "data": converted_data,
                        "success": True,
                    }
                )
            except Exception as e:
                st.session_state.results.append(
                    {
                        "name": uploaded_file.name,
                        "stem": file_stem,
                        "error": str(e),
                        "success": False,
                    }
                )
            finally:
                if tmp_dir and os.path.exists(tmp_dir):
                    import shutil

                    shutil.rmtree(tmp_dir, ignore_errors=True)

        progress.progress((i + 1) / len(uploaded_files))

    progress.empty()
    st.rerun()

if st.session_state.results:
    st.divider()
    success_count = sum(1 for r in st.session_state.results if r["success"])
    fail_count = len(st.session_state.results) - success_count
    st.subheader(f"결과: {success_count}건 성공, {fail_count}건 실패")

    for result in st.session_state.results:
        if result["success"]:
            col1, col2 = st.columns([3, 1])
            with col1:
                st.success(f"{result['name']} -> {result['stem']}.xlsx")
            with col2:
                st.download_button(
                    label="다운로드",
                    data=result["data"],
                    file_name=f"{result['stem']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{result['name']}",
                )
        else:
            st.error(f"{result['name']} 변환 실패: {result['error']}")
