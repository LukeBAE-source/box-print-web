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
    st.write("개별 품목 정보를 입력하면 1건 렌더링(PDF 1개) 후 다운로드합니다.")

    # templates 폴더를 스캔해서 brand / box_type / box_group 옵션 자동 생성
    brand_options = sorted([p.name for p in TEMPLATES_DIR.iterdir() if p.is_dir()])

    brand_to_pairs = {}
    for b in brand_options:
        pairs = set()
        for pdf in (TEMPLATES_DIR / b).glob("*.pdf"):
            stem = pdf.stem  # 예: BASIC_M
            if "_" in stem:
                bt, bg = stem.split("_", 1)
                pairs.add((bt, bg))
        brand_to_pairs[b] = sorted(pairs)

    # ✅ 좌/우 레이아웃 (왼쪽 입력, 오른쪽 사용법)
    left, right = st.columns([1, 1], gap="large")

    with left:
        # ✅ 각 항목을 "각 행"으로(한 줄씩) 표시
        brand = st.selectbox(
            "brand (예: iloom, desker, sloubed) - 선택",
            options=[""] + brand_options,
            index=0,
        )

        item_code = st.text_input("item_code (품목코드) - 입력", value="")

        origin_country = st.selectbox(
            "origin_country (원산지, 예: KR, CN, VN) - 선택",
            options=["", "KR", "CN", "VN"],
            index=0,
        )

        product_name_ko = st.text_input("product_name_ko (단품명) - 입력", value="")
        product_name_en = st.text_input("product_name_en (단품명_영문) - 입력", value="")

        pairs = brand_to_pairs.get(brand, []) if brand else []
        box_type_options = sorted({bt for bt, _ in pairs})

        box_type = st.selectbox(
            "box_type (예: BASIC, PANEL 등) - 선택",
            options=[""] + box_type_options,
            index=0,
            disabled=(not brand),
        )

        if brand and box_type:
            box_group_options = sorted({bg for bt, bg in pairs if bt == box_type})
        else:
            box_group_options = []

        box_group = st.selectbox(
            "box_group (예: M, S, MS 등) - 선택",
            options=[""] + box_group_options,
            index=0,
            disabled=(not brand or not box_type),
        )

        limit_preview = st.checkbox("미리보기(입력값 확인)", value=True)

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

        run_manual = st.button("실행(개별 입력)", type="primary")

    with right:
        st.subheader("사용법")
        # ✅ 여기 내용을 원하는대로 수정해서 사용법을 채우면 됩니다.
        st.markdown(
            """
            1. **brand**를 선택합니다. (templates 폴더에 있는 브랜드만 표시됩니다.)  
            2. **item_code**(품목코드)를 입력합니다.  
            3. **원산지 / 단품명(국문/영문)**을 입력합니다.  
            4. **box_type → box_group**을 순서대로 선택합니다.  
               - brand 선택 후 box_type이 활성화됩니다.  
               - box_type 선택 후 box_group이 활성화됩니다.  
            5. **실행(개별 입력)**을 누르면 PDF 1개가 생성되고 다운로드 버튼이 나타납니다.
            
            **주의**
            - 선택한 조합에 해당하는 템플릿 파일이 `templates/<brand>/<box_type>_<box_group>.pdf` 로 존재해야 합니다.
            - 좌표가 `coords/coords.json`에 없으면 오류가 날 수 있습니다.
            """
        )

    # ✅ 실행 로직은 레이아웃 밖에서 처리(버튼 클릭 시)
    if run_manual:
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
            else:
                pdf_path = Path(out_paths[0])
                if not pdf_path.exists():
                    st.error("PDF 파일을 찾을 수 없습니다.")
                else:
                    st.success("완료: PDF 1개 생성")
                    st.download_button(
                        "PDF 다운로드",
                        data=pdf_path.read_bytes(),
                        file_name=pdf_path.name,
                        mime="application/pdf",
                    )

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