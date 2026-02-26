# app.py
import io
import time
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from renderer import run_render

PROJECT_DIR = Path(__file__).resolve().parent
TMP_DIR = PROJECT_DIR / "_tmp_uploads"
TMP_DIR.mkdir(exist_ok=True)

FONT_PATH = PROJECT_DIR / "fonts" / "NotoSansKR-Medium.ttf"
COORDS_JSON = PROJECT_DIR / "coords" / "coords.json"
TEMPLATES_DIR = PROJECT_DIR / "templates"
ICONS_DIR = PROJECT_DIR / "icons"
OUT_DIR = PROJECT_DIR / "output_pdf"

REQUIRED_COLS = [
    "brand",
    "box_type",
    "box_group",
    "item_code",
    "product_name_ko",
    "product_name_en",
    "origin_country",
]

st.set_page_config(page_title="포장박스 인쇄 시안 자동화", layout="wide")
st.title("포장박스 인쇄 시안 자동화")

# ✅ 탭 순서: 개별 품목 입력 → 엑셀 업로드
tab_manual, tab_upload = st.tabs(["개별 품목 입력", "엑셀 업로드"])


def render_and_zip(excel_path: Path, ts: str):
    """공통: renderer 실행 후 zip 생성/다운로드 버튼 제공"""
    with st.spinner("렌더링 중..."):
        out_paths = run_render(
            excel_path=str(excel_path),
            limit=0,
            template_root=str(TEMPLATES_DIR),
            icon_dir=str(ICONS_DIR),
            font_path=str(FONT_PATH),
            coords_json_path=str(COORDS_JSON),
            out_dir=str(OUT_DIR),
        )

    if not out_paths:
        st.error("결과 파일이 생성되지 않았습니다. output_pdf 폴더를 확인하세요.")
        return

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


with tab_manual:
    st.write("개별 품목 정보를 입력하면 1건짜리 엑셀을 자동 생성해서 렌더링합니다.")

    with st.form("manual_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            brand = st.text_input("brand (예: iloom, desker, sloubed)", value="")
            box_type = st.text_input("box_type (예: BASIC, PANEL 등)", value="")
            box_group = st.text_input("box_group (예: M, S, MS 등)", value="")
        with c2:
            item_code = st.text_input("item_code (품목코드)", value="")
            origin_country = st.text_input("origin_country (예: KR, CN, VN)", value="")
        with c3:
            product_name_ko = st.text_input("product_name_ko", value="")
            product_name_en = st.text_input("product_name_en", value="")

        limit_preview = st.checkbox("미리보기(입력값 확인)", value=True)
        submit = st.form_submit_button("실행(개별 입력)", type="primary")

    if limit_preview:
        preview_df = pd.DataFrame(
            [{
                "brand": brand,
                "box_type": box_type,
                "box_group": box_group,
                "item_code": item_code,
                "product_name_ko": product_name_ko,
                "product_name_en": product_name_en,
                "origin_country": origin_country,
            }],
            columns=REQUIRED_COLS,
        )
        st.dataframe(preview_df, use_container_width=True)

    if submit:
        required_values = {
            "brand": brand,
            "box_type": box_type,
            "box_group": box_group,
            "item_code": item_code,
            "product_name_ko": product_name_ko,
            "product_name_en": product_name_en,
            "origin_country": origin_country,
        }

        missing = [k for k, v in required_values.items() if not str(v).strip()]
        if missing:
            st.error(f"필수 입력이 비어있습니다: {missing}")
        else:
            ts = time.strftime("%Y%m%d_%H%M%S")
            excel_path = TMP_DIR / f"manual_{item_code}_{ts}.xlsx"

            df = pd.DataFrame([required_values], columns=REQUIRED_COLS)
            df.to_excel(excel_path, index=False)

            render_and_zip(excel_path, ts)


with tab_upload:
    st.write("box_data.xlsx 파일을 업로드하고 실행을 누르면 결과를 다운로드할 수 있습니다.")
    uploaded = st.file_uploader("box_data.xlsx 업로드", type=["xlsx"])
    run_btn = st.button("실행(업로드)", type="primary", disabled=(uploaded is None))

    if run_btn:
        ts = time.strftime("%Y%m%d_%H%M%S")
        excel_path = TMP_DIR / f"box_data_{ts}.xlsx"
        with open(excel_path, "wb") as f:
            f.write(uploaded.getbuffer())

        render_and_zip(excel_path, ts)