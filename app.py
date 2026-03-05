# app.py
import io
import time
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from renderer import run_render, safe_filename

# -----------------------------
# Paths
# -----------------------------
PROJECT_DIR = Path(__file__).resolve().parent
TMP_DIR = PROJECT_DIR / "_tmp_uploads"
TMP_DIR.mkdir(exist_ok=True)

ASSETS_DIR = PROJECT_DIR / "assets"
TEMPLATE_TABLE_IMG = ASSETS_DIR / "template_table.png"
MANUALS_DIR = ASSETS_DIR / "manuals"

FORMS_DIR = ASSETS_DIR / "forms"
BOX_DATA_TEMPLATE = FORMS_DIR / "box_data.xlsx"  # ✅ 양식 다운로드 파일명

TEMPLATES_DIR = PROJECT_DIR / "templates"
COORDS_JSON = PROJECT_DIR / "coords" / "coords.json"
ICONS_DIR = PROJECT_DIR / "icons"
OUT_DIR = PROJECT_DIR / "output_pdf"
OUT_DIR.mkdir(exist_ok=True)

# -----------------------------
# Const
# -----------------------------
BRAND_NAME_KO = {
    "iloom": "일룸",
    "desker": "데스커",
    "sloubed": "슬로우베드",
}

REQUIRED_COLS = [
    "brand",
    "box_type",
    "box_group",
    "item_code",
    "product_name_ko",
    "product_name_en",
    "origin_country",
]

# -----------------------------
# Page
# -----------------------------
st.set_page_config(page_title="포장박스 인쇄 시안 자동화", layout="wide")
st.title("포장박스 인쇄 시안 자동화")

# ✅ UI 정리용 CSS (버튼/텍스트/여백)
st.markdown(
    """
<style>
/* 다운로드 버튼: 작게 + 줄바꿈 방지 */
div.stDownloadButton > button{
  padding: 0.25rem 0.65rem !important;
  font-size: 0.85rem !important;
  line-height: 1.1 !important;
  white-space: nowrap !important;
}

/* 오른쪽 패널 타이틀 간격 */
h3 { margin-bottom: 0.2rem !important; }
</style>
""",
    unsafe_allow_html=True,
)

tab_manual, tab_upload = st.tabs(["개별 품목 입력", "엑셀 업로드"])


# -----------------------------
# Helpers
# -----------------------------
def _ensure_prerequisites():
    problems = []
    if not TEMPLATES_DIR.exists():
        problems.append(f"templates 폴더가 없습니다: {TEMPLATES_DIR}")
    if not COORDS_JSON.exists():
        problems.append(f"coords.json 파일이 없습니다: {COORDS_JSON}")
    if problems:
        st.error("프로젝트 필수 파일/폴더가 누락되었습니다:\n- " + "\n- ".join(problems))
        st.stop()


def _scan_brand_templates():
    if not TEMPLATES_DIR.exists():
        return [], {}

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


def _render_single_pdf(row: dict) -> Path:
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
    missing_cols = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing_cols:
        st.error(f"엑셀 필수 컬럼 누락: {missing_cols}\n현재 컬럼: {list(df.columns)}")
        return

    ok_paths = []
    fail_rows = []

    with st.spinner("렌더링 중..."):
        for i, r in df.iterrows():
            row = {c: r.get(c, "") for c in REQUIRED_COLS}

            # 최소 필수값 체크
            if (
                not str(row["brand"]).strip()
                or not str(row["box_type"]).strip()
                or not str(row["box_group"]).strip()
                or not str(row["item_code"]).strip()
            ):
                fail_rows.append((i + 2, "필수값(brand/box_type/box_group/item_code) 누락"))
                continue

            try:
                p = _render_single_pdf(row)
                if p.exists() and p.stat().st_size > 0:
                    ok_paths.append(p)
                else:
                    fail_rows.append((i + 2, "PDF 생성 실패(파일 없음 또는 0바이트)"))
            except Exception as e:
                fail_rows.append((i + 2, str(e)))

    if not ok_paths:
        st.error("생성된 PDF가 없습니다. 템플릿/좌표/입력값을 확인하세요.")
        if fail_rows:
            st.info("\n".join([f"- row {n}: {msg}" for n, msg in fail_rows[:30]]))
        return

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in ok_paths:
            zf.write(p, arcname=p.name)
    zip_buffer.seek(0)

    st.success(f"완료: {len(ok_paths)}개 PDF 생성")
    if fail_rows:
        st.warning("실패 내역(최대 30건):\n" + "\n".join([f"- row {n}: {msg}" for n, msg in fail_rows[:30]]))

    st.download_button(
        "결과 다운로드(zip)",
        data=zip_buffer.getvalue(),
        file_name=f"output_{ts}.zip",
        mime="application/zip",
    )


# -----------------------------
# Init
# -----------------------------
_ensure_prerequisites()
brand_options, brand_to_pairs = _scan_brand_templates()


# -----------------------------
# Tab 1: Manual input
# -----------------------------
with tab_manual:
    st.write("개별 품목 정보를 입력하면 PDF 1개를 생성하고 다운로드합니다.")
    left, right = st.columns([1, 1], gap="large")

    # ---- Left: inputs
    with left:
        brand = st.selectbox(
            "brand (예: iloom, desker, sloubed) - 선택",
            options=[""] + brand_options,
            index=0,
        )

        item_code = st.text_input("item_code (품목코드) - 입력", value="")
        product_name_ko = st.text_input("product_name_ko (단품명) - 입력", value="")
        product_name_en = st.text_input("product_name_en (단품명_영문) - 입력", value="")

        icon_keys = sorted([p.stem.replace("icon_", "", 1) for p in ICONS_DIR.glob("icon_*.png") if p.is_file()])
        origin_country = st.selectbox(
            "origin_country (원산지) - 선택",
            options=[""] + [k.upper() for k in icon_keys],
            index=0,
        )

        pairs = brand_to_pairs.get(brand, []) if brand else []
        box_type_options = sorted({bt for bt, _ in pairs})
        box_type = st.selectbox(
            "box_type (예: BASIC, PANEL 등) - 선택",
            options=[""] + box_type_options,
            index=0,
            disabled=(not brand),
        )

        box_group_options = sorted({bg for bt, bg in pairs if bt == box_type}) if (brand and box_type) else []
        box_group = st.selectbox(
            "box_group (예: M, S, MS 등) - 선택",
            options=[""] + box_group_options,
            index=0,
            disabled=(not brand or not box_type),
        )

        preview = st.checkbox("미리보기(입력값 확인)", value=True)
        if preview:
            preview_df = pd.DataFrame(
                [
                    {
                        "brand": brand,
                        "box_type": box_type,
                        "box_group": box_group,
                        "item_code": item_code,
                        "product_name_ko": product_name_ko,
                        "product_name_en": product_name_en,
                        "origin_country": origin_country,
                    }
                ],
                columns=REQUIRED_COLS,
            )
            st.dataframe(preview_df, use_container_width=True)

        run_manual = st.button("실행(개별 입력)", type="primary")

    # ---- Right: usage + manuals + template table (all inside right)
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
            st.caption("포장 규격/박스 타입 확인 후 사용하세요.")

            if not MANUALS_DIR.exists():
                st.warning("assets/manuals 폴더가 없습니다.")
            else:
                for b in brand_options:
                    manual_path = MANUALS_DIR / f"manual_{b}.pdf"
                    brand_ko = BRAND_NAME_KO.get(b, b)

                    # ✅ 한 줄: "~매뉴얼" [다운로드]
                    row_text, row_btn = st.columns([7, 3], gap="small")
                    with row_text:
                        st.markdown(f"{brand_ko} 포장박스 매뉴얼")
                    with row_btn:
                        if manual_path.exists():
                            with open(manual_path, "rb") as f:
                                st.download_button(
                                    "다운로드",
                                    data=f,
                                    file_name=manual_path.name,
                                    mime="application/pdf",
                                    key=f"manual_{b}",
                                    use_container_width=True,
                                )
                        else:
                            st.caption("없음")

        st.markdown("---")
        st.subheader("템플릿 기준표")
        if TEMPLATE_TABLE_IMG.exists():
            st.image(str(TEMPLATE_TABLE_IMG), use_container_width=True)
        else:
            st.warning("assets/template_table.png 파일이 없습니다.")

    # ---- Run manual
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
            try:
                with st.spinner("렌더링 중..."):
                    pdf_path = _render_single_pdf(required_values)

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
            except Exception as e:
                st.error(f"렌더링 실패: {e}")


# -----------------------------
# Tab 2: Excel upload
# -----------------------------
with tab_upload:
    st.write("양식을 다운 받아 작성 후 box_data.xlsx 파일을 업로드하고 실행을 누르면 결과를 ZIP으로 다운로드할 수 있습니다.")

    left_u, right_u = st.columns([1, 1], gap="large")

    with left_u:
        # ✅ 한 줄: 양식 다운로드 [다운로드]
        tcol, bcol = st.columns([3, 1], gap="small")
        with tcol:
            st.markdown("**양식 다운로드**")
        with bcol:
            if BOX_DATA_TEMPLATE.exists():
                with open(BOX_DATA_TEMPLATE, "rb") as f:
                    st.download_button(
                        "다운로드",
                        data=f,
                        file_name="box_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_box_data_template",
                        use_container_width=True,
                    )
            else:
                st.caption("없음")

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

### 주의사항
- 템플릿: `templates/<brand>/<box_type>_<box_group>.pdf`
- 좌표: `coords/coords.json`
- 원산지 아이콘: `icons/icon_<origin>.png`
"""
        )