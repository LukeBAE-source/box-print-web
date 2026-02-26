# app.py
import io
import time
import zipfile
from pathlib import Path

import streamlit as st
from renderer import run_render

PROJECT_DIR = Path(__file__).resolve().parent
TMP_DIR = PROJECT_DIR / "_tmp_uploads"
TMP_DIR.mkdir(exist_ok=True)

st.set_page_config(page_title="포장박스 인쇄 자동화", layout="wide")
st.title("포장박스 인쇄 자동화")

st.write("box_data.xlsx 파일을 업로드하고 실행을 누르면 결과를 다운로드할 수 있습니다.")

uploaded = st.file_uploader("box_data.xlsx 업로드", type=["xlsx"])

run_btn = st.button("실행", type="primary", disabled=(uploaded is None))

if run_btn:
    ts = time.strftime("%Y%m%d_%H%M%S")

    # 1) 업로드 파일 저장
    excel_path = TMP_DIR / f"box_data_{ts}.xlsx"
    with open(excel_path, "wb") as f:
        f.write(uploaded.getbuffer())

    # 2) 렌더 실행 (renderer.py의 run_render() 인자에 맞춤)
    with st.spinner("렌더링 중..."):
        out_paths = run_render(
            excel_path=str(excel_path),
            limit=0,  # 0이면 전체 처리, 테스트 시 5 같은 숫자로 바꿔도 됨
            template_root=str(PROJECT_DIR / "templates"),
            icon_dir=str(PROJECT_DIR / "icons"),
            font_path=str(PROJECT_DIR / "fonts" / "NotoSansKR-Medium.ttf"),
            coords_json_path=str(PROJECT_DIR / "coords" / "coords.json"),
            out_dir=str(PROJECT_DIR / "output_pdf"),
        )

    # 3) 결과 zip 생성
    if not out_paths:
        st.error("결과 파일이 생성되지 않았습니다. output_pdf 폴더를 확인하세요.")
    else:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in out_paths:
                p = Path(p)
                if p.exists() and p.is_file():
                    zf.write(p, arcname=p.name)

        st.success(f"완료: {len(out_paths)}개 파일 생성")
        st.download_button(
            "결과 다운로드(zip)",
            data=zip_buffer.getvalue(),
            file_name=f"output_{ts}.zip",
            mime="application/zip",
        )