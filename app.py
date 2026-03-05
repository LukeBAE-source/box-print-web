# app.py
import io
import time
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from renderer import run_render, safe_filename

PROJECT_DIR = Path(__file__).resolve().parent
TMP_DIR = PROJECT_DIR / "_tmp_uploads"
TMP_DIR.mkdir(exist_ok=True)

# assets
ASSETS_DIR = PROJECT_DIR / "assets"
TEMPLATE_TABLE_IMG = ASSETS_DIR / "template_table.png"
MANUALS_DIR = ASSETS_DIR / "manuals"

# ✅ 엑셀 양식 다운로드 파일
FORMS_DIR = ASSETS_DIR / "forms"
BOX_DATA_TEMPLATE = FORMS_DIR / "box_data.xlsx"

# 브랜드 한글명
BRAND_NAME_KO = {
    "iloom": "일룸",
    "desker": "데스커",
    "sloubed": "슬로우베드",
}

TEMPLATES_DIR = PROJECT_DIR / "templates"
COORDS_JSON = PROJECT_DIR / "coords" / "coords.json"
ICONS_DIR = PROJECT_DIR / "icons"
OUT_DIR = PROJECT_DIR / "output_pdf"
OUT_DIR.mkdir(exist_ok=True)

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

tab_manual, tab_upload = st.tabs(["개별 품목 입력", "엑셀 업로드"])


def _scan_brand_templates():

    brand_options = sorted([p.name for p in TEMPLATES_DIR.iterdir() if p.is_dir()])

    brand_to_pairs = {}

    for b in brand_options:
        pairs = set()
        for pdf in (TEMPLATES_DIR / b).glob("*.pdf"):
            stem = pdf.stem
            if "_" in stem:
                bt, bg = stem.split("_", 1)
                pairs.add((bt, bg))
        brand_to_pairs[b] = sorted(pairs)

    return brand_options, brand_to_pairs


def _render_single_pdf(row: dict, ts: str) -> Path:

    brand = str(row["brand"]).strip()
    box_type = str(row["box_type"]).strip()
    box_group = str(row["box_group"]).strip()

    item_code = str(row["item_code"]).strip()
    name_ko = str(row["product_name_ko"]).strip()
    name_en = str(row["product_name_en"]).strip()
    origin_country = str(row["origin_country"]).strip()

    template_key = f"{box_type}_{box_group}".lower()

    filename = safe_filename(f"{brand}_{template_key}_{item_code}.pdf")

    output_path = OUT_DIR / filename

    run_render(
        brand=brand,
        template_key=template_key,
        item_code=item_code,
        name_ko=name_ko,
        name_en=name_en,
        origin_country=origin_country,
        output_path=str(output_path),
    )

    return output_path


def _render_excel_to_zip(excel_path: Path, ts: str):

    df = pd.read_excel(excel_path, dtype=str).fillna("")

    ok_paths = []
    fail_rows = []

    with st.spinner("렌더링 중..."):

        for i, r in df.iterrows():

            row = {c: r.get(c, "") for c in REQUIRED_COLS}

            try:

                p = _render_single_pdf(row, ts)

                ok_paths.append(p)

            except Exception as e:

                fail_rows.append((i + 2, str(e)))

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:

        for p in ok_paths:
            zf.write(p, arcname=p.name)

    zip_buffer.seek(0)

    st.success(f"완료: {len(ok_paths)}개 PDF 생성")

    st.download_button(
        "결과 다운로드(zip)",
        data=zip_buffer.getvalue(),
        file_name=f"output_{ts}.zip",
        mime="application/zip",
    )


brand_options, brand_to_pairs = _scan_brand_templates()

# -----------------------------
# 개별 입력 탭
# -----------------------------

with tab_manual:

    st.write("개별 품목 정보를 입력하면 PDF 1개를 생성하고 다운로드합니다.")

    left, right = st.columns([1, 1], gap="large")

    with left:

        brand = st.selectbox("brand", options=[""] + brand_options)

        item_code = st.text_input("item_code")
        product_name_ko = st.text_input("product_name_ko")
        product_name_en = st.text_input("product_name_en")

        origin_country = st.text_input("origin_country")

        pairs = brand_to_pairs.get(brand, []) if brand else []

        box_type_options = sorted({bt for bt, _ in pairs})

        box_type = st.selectbox("box_type", options=[""] + box_type_options)

        if brand and box_type:
            box_group_options = sorted({bg for bt, bg in pairs if bt == box_type})
        else:
            box_group_options = []

        box_group = st.selectbox("box_group", options=[""] + box_group_options)

        run_manual = st.button("실행(개별 입력)")

with right:
    usage_col, manual_col = st.columns([1.2, 1], gap="large")

    with usage_col:
        st.subheader("사용법")
        st.markdown(
            """
1. **brand** 선택  
2. **item_code / 단품명(국문/영문) / 원산지** 입력  
3. **box_type → box_group 선택 (<span style="color:red;font-weight:700;">템플릿 기준표 참조</span>)**  
4. **실행(개별 입력)** 클릭 → PDF 다운로드  
""",
            unsafe_allow_html=True,
        )

    with manual_col:
        st.subheader("브랜드 매뉴얼")
        st.caption("다운로드해서 포장 규격/박스 타입 확인 후 사용하세요.")

        for b in brand_options:
            manual_path = MANUALS_DIR / f"manual_{b}.pdf"
            brand_ko = BRAND_NAME_KO.get(b, b)

            crow_text, row_btn = st.columns([8,1], gap="small")
            with col1:
                st.markdown(f"**{brand_ko} 포장박스 매뉴얼**")
            with col2:
                if manual_path.exists():
                    with open(manual_path, "rb") as f:
                        st.download_button(
                            "다운로드",
                            data=f,
                            file_name=manual_path.name,
                            mime="application/pdf",
                            key=f"manual_{b}",
                        )
                else:
                    st.caption("없음")

    # ✅ 템플릿 기준표를 right 영역 안으로 이동
    st.markdown("---")
    st.subheader("템플릿 기준표")
    if TEMPLATE_TABLE_IMG.exists():
        st.image(str(TEMPLATE_TABLE_IMG), use_container_width=True)
    else:
        st.warning("assets/template_table.png 파일이 없습니다.")

# -----------------------------
# 엑셀 업로드 탭
# -----------------------------

with tab_upload:

    st.write(
        "양식을 다운 받아 작성 후 box_data.xlsx 파일을 업로드하고 실행을 누르면 결과를 ZIP으로 다운로드할 수 있습니다."
    )

    left_u, right_u = st.columns([1, 1], gap="large")

    with left_u:

        # ✅ 양식 다운로드
        col_a, col_b = st.columns([3,1])

        with col_a:
            st.markdown("**양식 다운로드**")

        with col_b:
            if BOX_DATA_TEMPLATE.exists():
                with open(BOX_DATA_TEMPLATE, "rb") as f:
                    st.download_button(
                        "다운로드",
                        data=f,
                        file_name="box_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        uploaded = st.file_uploader("box_data.xlsx 업로드", type=["xlsx"])

        run_btn = st.button("실행(업로드)", type="primary", disabled=(uploaded is None))

        if run_btn and uploaded is not None:

            ts = time.strftime("%Y%m%d_%H%M%S")

            excel_path = TMP_DIR / f"box_data_{ts}.xlsx"

            with open(excel_path, "wb") as f:
                f.write(uploaded.getbuffer())

            _render_excel_to_zip(excel_path, ts)

    with right_u:

        st.subheader("사용법")

        st.markdown(
            """
### 업로드 방법

1. **box_data.xlsx** 업로드  
2. **실행(업로드)** 클릭  
3. 완료 후 ZIP 다운로드  

### 엑셀 필수 컬럼

- brand  
- box_type  
- box_group  
- item_code  
- product_name_ko  
- product_name_en  
- origin_country  
"""
        )