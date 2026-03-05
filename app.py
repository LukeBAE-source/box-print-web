# app.py
import io
import time
import zipfile
import traceback
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
BOX_DATA_TEMPLATE = FORMS_DIR / "box_data.xlsx"

TEMPLATES_DIR = PROJECT_DIR / "templates"
COORDS_JSON = PROJECT_DIR / "coords" / "coords.json"
ICONS_DIR = PROJECT_DIR / "icons"

OUT_DIR = PROJECT_DIR / "output_pdf"
OUT_DIR.mkdir(exist_ok=True)

# --------------------------
# Config
# --------------------------
BRAND_NAME_KO = {
    "iloom": "일룸",
    "desker": "데스커",
    "sloubed": "슬로우",
}

REQUIRED_COLS = [
    "brand",
    "item_code",
    "origin_country",
    "product_name_ko",
    "product_name_en",
    "box_type",
    "box_group",
]

ORIGIN_OPTIONS = ["KOREA", "CHINA", "VIETNAM"]


# --------------------------
# Helpers
# --------------------------
def _ensure_prerequisites():
    missing = []
    if not TEMPLATES_DIR.exists():
        missing.append(f"templates 폴더 없음: {TEMPLATES_DIR}")
    if not COORDS_JSON.exists():
        missing.append(f"coords.json 없음: {COORDS_JSON}")
    if not ICONS_DIR.exists():
        # 아이콘은 없어도 텍스트 대체 가능하지만, 프로젝트 요구에 따라 막을지 선택
        missing.append(f"icons 폴더 없음: {ICONS_DIR}")

    if missing:
        st.error("필수 리소스가 없습니다:\n\n- " + "\n- ".join(missing))
        st.stop()


def _scan_template_options(brand: str):
    brand = (brand or "").strip().lower()
    brand_dir = TEMPLATES_DIR / brand
    if not brand_dir.exists():
        return [], {}

    pdfs = list(brand_dir.glob("*.pdf"))
    options = []
    mapping = {}
    for p in pdfs:
        key = p.stem.strip()
        if not key:
            continue
        options.append(key)
        mapping[key.lower()] = key

    options_sorted = sorted(set(options), key=lambda x: x.lower())
    return options_sorted, mapping


def _split_template_key(template_key: str):
    s = (template_key or "").strip()
    if "_" not in s:
        return s, ""
    a, b = s.split("_", 1)
    return a, b


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
        raise ValueError(f"엑셀에 필수 컬럼이 없습니다: {missing_cols}")

    out_mem = io.BytesIO()
    failures = []

    with zipfile.ZipFile(out_mem, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for idx, r in df.iterrows():
            row = {c: str(r.get(c, "")).strip() for c in REQUIRED_COLS}
            missing = [k for k, v in row.items() if not str(v).strip()]
            if missing:
                failures.append((idx + 2, "필수 입력 누락", f"{missing}"))
                continue

            try:
                pdf_path = _render_single_pdf(row)
                zf.writestr(pdf_path.name, pdf_path.read_bytes())
            except Exception as e:
                failures.append((idx + 2, "렌더링 실패", str(e)))

    out_mem.seek(0)
    zip_name = f"output_{ts}.zip"
    return out_mem, zip_name, failures


# --------------------------
# UI
# --------------------------
st.set_page_config(page_title="포장박스 인쇄 자동화", layout="wide")
_ensure_prerequisites()

st.title("포장박스 인쇄 자동화")

tab_single, tab_upload = st.tabs(["개별 품목 입력", "엑셀 업로드"])

# -----------------------------
# Tab 1: Single item
# -----------------------------
with tab_single:
    left, right = st.columns([1, 1], gap="large")
    brand_options = list(BRAND_NAME_KO.keys())

    with left:
        with st.form("single_item_form"):
            st.subheader("입력")

            brand = st.selectbox("brand", brand_options, format_func=lambda x: BRAND_NAME_KO.get(x, x))
            template_options, _ = _scan_template_options(brand)

            if not template_options:
                st.warning(f"templates/{brand} 폴더에 PDF 템플릿이 없습니다.")
                box_type_options = [""]
                box_group_options = [""]
            else:
                bt = []
                bg = []
                for k in template_options:
                    a, b = _split_template_key(k)
                    if a:
                        bt.append(a.upper())
                    if b:
                        bg.append(b.upper())
                box_type_options = sorted(set(bt)) or [""]
                box_group_options = sorted(set(bg)) or [""]

            box_type = st.selectbox("box_type", box_type_options)
            box_group = st.selectbox("box_group", box_group_options)

            item_code = st.text_input("item_code")
            product_name_ko = st.text_input("product_name_ko")
            product_name_en = st.text_input("product_name_en")
            origin_country = st.selectbox("origin_country", ORIGIN_OPTIONS)

            run_manual = st.form_submit_button("실행(개별 입력)")

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
                    st.code(traceback.format_exc())

    with right:
        st.subheader("사용법")
        st.markdown(
            """
1) brand 선택  
2) box_type/box_group 선택  
3) item_code, 단품명(국문/영문), 원산지 입력  
4) 실행(개별 입력) → PDF 다운로드  
"""
        )
        if TEMPLATE_TABLE_IMG.exists():
            st.image(str(TEMPLATE_TABLE_IMG), caption="템플릿 기준표", use_container_width=True)

        st.subheader("브랜드 매뉴얼")
        if MANUALS_DIR.exists():
            for b in brand_options:
                manual_path = MANUALS_DIR / f"manual_{b}.pdf"
                brand_ko = BRAND_NAME_KO.get(b, b)
                row_text, row_btn = st.columns([6, 2], gap="small")
                with row_text:
                    st.write(f"{brand_ko} 매뉴얼")
                with row_btn:
                    if manual_path.exists():
                        st.download_button(
                            "다운로드",
                            data=manual_path.read_bytes(),
                            file_name=manual_path.name,
                            mime="application/pdf",
                            key=f"manual_{b}",
                        )

# -----------------------------
# Tab 2: Excel upload
# -----------------------------
with tab_upload:
    st.write("양식을 다운 받아 작성 후 box_data.xlsx 파일을 업로드하고 실행을 누르면 결과를 ZIP으로 다운로드할 수 있습니다.")

    tcol, bcol = st.columns([3, 1], gap="small")
    with tcol:
        st.markdown("**양식 다운로드**")
    with bcol:
        if BOX_DATA_TEMPLATE.exists():
            st.download_button(
                "다운로드",
                data=BOX_DATA_TEMPLATE.read_bytes(),
                file_name=BOX_DATA_TEMPLATE.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("assets/forms/box_data.xlsx 파일이 없습니다.")

    st.divider()
    uploaded = st.file_uploader("box_data.xlsx 업로드", type=["xlsx"])
    run_upload = st.button("실행(엑셀 업로드)", type="primary", disabled=(uploaded is None))

    if run_upload and uploaded is not None:
        ts = time.strftime("%Y%m%d_%H%M%S")
        tmp_excel = TMP_DIR / f"box_data_{ts}.xlsx"
        tmp_excel.write_bytes(uploaded.read())

        try:
            with st.spinner("일괄 렌더링 중..."):
                zip_mem, zip_name, failures = _render_excel_to_zip(tmp_excel, ts)

            st.success("완료: ZIP 생성")
            st.download_button(
                "ZIP 다운로드",
                data=zip_mem.getvalue(),
                file_name=zip_name,
                mime="application/zip",
            )

            if failures:
                st.warning("일부 행이 실패했습니다.")
                fail_df = pd.DataFrame(failures, columns=["excel_row", "type", "detail"])
                st.dataframe(fail_df, use_container_width=True)

        except Exception as e:
            st.error(f"실행 실패: {e}")
            st.code(traceback.format_exc())