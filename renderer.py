import os
import re
import json
import argparse
import pandas as pd

from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from pypdf import PdfReader, PdfWriter

# =========================
# 기본 경로(단독 실행 시 사용)
# =========================
XLSX_PATH = "box_data.xlsx"
TEMPLATE_ROOT = "templates"
ICON_DIR = "icons"
FONT_PATH = os.path.join("fonts", "NotoSansKR-Medium.ttf")
COORDS_JSON_PATH = os.path.join("coords", "coords.json")
OUT_DIR = "output_pdf"

# =========================
# 유틸
# =========================
def norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s or "")).strip().lower()

def register_fonts(font_path: str = FONT_PATH):
    if not os.path.exists(font_path):
        raise FileNotFoundError(f"폰트 파일 없음: {font_path}")
    pdfmetrics.registerFont(TTFont("NotoSansKR-Medium", font_path))

def draw_text_rotated_180(c, x, y, text):
    c.saveState()
    c.translate(x, y)
    c.rotate(180)
    c.drawString(0, 0, text)
    c.restoreState()

def draw_image_rotated_180(c, img_path, x, y, w, h):
    c.saveState()
    c.translate(x + w, y + h)
    c.rotate(180)
    c.drawImage(ImageReader(img_path), 0, 0, width=w, height=h, mask="auto")
    c.restoreState()

def load_coords(coords_json_path: str = COORDS_JSON_PATH):
    if not os.path.exists(coords_json_path):
        raise FileNotFoundError(f"좌표 JSON 없음: {coords_json_path}")

    with open(coords_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    defaults = data.get("defaults", {}) or {}
    templates = data.get("templates", {}) or {}

    template_coords = {}
    for brand, mp in templates.items():
        if not isinstance(mp, dict):
            continue
        for template_key, cfg in mp.items():
            template_coords[(brand, template_key)] = cfg

    return defaults, template_coords

def get_icon_path(country: str, icon_dir: str = ICON_DIR) -> str:
    c = re.sub(r"\s+", "", str(country or "")).strip().upper()
    return os.path.join(icon_dir, f"icon_{c}.png")

# templates/<brand>/<box_type>_<box_group>.pdf (공백/대소문자 무시)
def find_template_pdf(brand: str, box_type: str, box_group: str, template_root: str = TEMPLATE_ROOT) -> str:
    brand_dir = os.path.join(template_root, brand)
    if not os.path.isdir(brand_dir):
        raise FileNotFoundError(f"템플릿 브랜드 폴더 없음: {brand_dir}")

    target = norm(f"{box_type}_{box_group}")

    for fn in os.listdir(brand_dir):
        if not fn.lower().endswith(".pdf"):
            continue
        key = norm(os.path.splitext(fn)[0])
        if key == target:
            return os.path.join(brand_dir, fn)

    raise FileNotFoundError(f"템플릿 없음: {brand}/{box_type}_{box_group}.pdf")

def get_cfg_for_template(defaults: dict, template_coords: dict, brand: str, template_pdf_path: str) -> dict:
    stem = os.path.splitext(os.path.basename(template_pdf_path))[0]
    template_key = norm(stem)

    cfg = template_coords.get((brand, template_key))
    if cfg:
        return cfg

    fallback = defaults.get(brand)
    if fallback:
        return fallback

    raise KeyError(f"좌표 없음: brand={brand}, template_key={template_key} (coords.json에 등록 필요)")

def is_zero_xy(xy) -> bool:
    try:
        return float(xy[0]) == 0 and float(xy[1]) == 0
    except Exception:
        return False

def is_zero_rect(rect) -> bool:
    try:
        return float(rect[0]) == 0 and float(rect[1]) == 0 and float(rect[2]) == 0 and float(rect[3]) == 0
    except Exception:
        return False

# =========================
# 오버레이 생성
# =========================
def make_overlay_pdf(overlay_path: str, page_w: float, page_h: float, cfg: dict, row: dict, icon_dir: str = ICON_DIR):
    c = canvas.Canvas(overlay_path, pagesize=(page_w, page_h))

    cover_rects = cfg.get("cover_rects", []) or []
    pos = cfg.get("pos", {}) or {}
    icon_pos = cfg.get("icon_pos", {}) or {}

    # 회전/숨김 설정(JSON에서 제어)
    rotate_cfg = cfg.get("rotate_180", {}) or {}              # 텍스트용: pos 키 기준
    icon_rotate_cfg = cfg.get("icon_rotate_180", {}) or {}    # 아이콘용: icon_pos 키 기준
    hide_cfg = cfg.get("hide", {}) or {}                      # 텍스트 숨김: pos 키 기준
    icon_hide_cfg = cfg.get("icon_hide", {}) or {}            # 아이콘 숨김: icon_pos 키 기준

    font_main_size = int(cfg.get("font_main_size", 26))
    font_sub_size = int(cfg.get("font_sub_size", 12))

    # 1) 가리기(흰 박스)
    c.setFillColorRGB(1, 1, 1)
    c.setStrokeColorRGB(1, 1, 1)
    for rect in cover_rects:
        if is_zero_rect(rect):
            continue
        x, y, w, h = rect
        c.rect(x, y, w, h, fill=1, stroke=0)

    # 2) 텍스트
    sku = str(row["item_code"])
    name_ko = str(row["product_name_ko"])
    name_en = str(row["product_name_en"])

    c.setFillColorRGB(0, 0, 0)

    # SKU(메인)
    c.setFont("NotoSansKR-Medium", font_main_size)
    for key, val in [("L_item_code", sku), ("L1_item_code", sku), ("L2_item_code", sku), ("L3_item_code", sku), ("R_item_code", sku)]:
        if key not in pos:
            continue
        if hide_cfg.get(key, False):
            continue
        if is_zero_xy(pos[key]):
            continue
        x, y = pos[key]
        if rotate_cfg.get(key, False):
            draw_text_rotated_180(c, x, y, val)
        else:
            c.drawString(x, y, val)

    # 이름(서브)
    c.setFont("NotoSansKR-Medium", font_sub_size)
    for key, val in [
        ("L_name_ko", name_ko),
        ("L_name_en", name_en),
        ("R_name_ko", name_ko),
        ("R_name_en", name_en),
        ("L1_name_ko", name_ko),
        ("L1_name_en", name_en),
        ("L2_name_ko", name_ko),
        ("L2_name_en", name_en),
        ("L3_name_ko", name_ko),
        ("L3_name_en", name_en),
    ]:
        if key not in pos:
            continue
        if hide_cfg.get(key, False):
            continue
        if is_zero_xy(pos[key]):
            continue
        x, y = pos[key]
        if rotate_cfg.get(key, False):
            draw_text_rotated_180(c, x, y, val)
        else:
            c.drawString(x, y, val)

    # 3) 원산지 아이콘
    origin = str(row.get("origin_country", "")).strip()
    ip = get_icon_path(origin, icon_dir=icon_dir)

    if os.path.exists(ip):
        for k in ("L_origin", "R_origin", "origin"):
            if k not in icon_pos:
                continue
            if icon_hide_cfg.get(k, False):
                continue
            if is_zero_rect(icon_pos[k]):
                continue
            x, y, w, h = icon_pos[k]
            if icon_rotate_cfg.get(k, False):
                draw_image_rotated_180(c, ip, x, y, w, h)
            else:
                c.drawImage(ImageReader(ip), x, y, width=w, height=h, mask="auto")
    else:
        # 아이콘 없으면 텍스트 대체(원하면 삭제 가능)
        msg = f"MADE IN {origin}"
        c.setFont("NotoSansKR-Medium", 10)
        for k in ("L_origin", "R_origin", "origin"):
            if k not in icon_pos:
                continue
            if icon_hide_cfg.get(k, False):
                continue
            if is_zero_rect(icon_pos[k]):
                continue
            x, y, w, h = icon_pos[k]
            if icon_rotate_cfg.get(k, False):
                draw_text_rotated_180(c, x, y + h + 2, msg)
            else:
                c.drawString(x, y + h + 2, msg)

    c.save()

# =========================
# 한 행 렌더링
# =========================
def render_row(
    defaults: dict,
    template_coords: dict,
    row: dict,
    *,
    template_root: str = TEMPLATE_ROOT,
    icon_dir: str = ICON_DIR,
    out_dir: str = OUT_DIR,
) -> str:
    brand = str(row["brand"]).strip()
    box_type = str(row["box_type"]).strip()
    box_group = str(row["box_group"]).strip()

    template_pdf = find_template_pdf(brand, box_type, box_group, template_root=template_root)
    cfg = get_cfg_for_template(defaults, template_coords, brand, template_pdf)

    reader = PdfReader(template_pdf)
    base_page = reader.pages[0]
    page_w = float(base_page.mediabox.width)
    page_h = float(base_page.mediabox.height)

    sku = str(row["item_code"])
    overlay_path = os.path.join(out_dir, f"__overlay_{sku}.pdf")
    make_overlay_pdf(overlay_path, page_w, page_h, cfg, row, icon_dir=icon_dir)

    overlay_reader = PdfReader(overlay_path)
    base_page.merge_page(overlay_reader.pages[0])

    out_path = os.path.join(out_dir, f"{brand}_{box_type}_{box_group}_{sku}.pdf")
    writer = PdfWriter()
    writer.add_page(base_page)
    with open(out_path, "wb") as f:
        writer.write(f)

    try:
        os.remove(overlay_path)
    except:
        pass

    return out_path

# =========================
# ✅ Streamlit에서 호출할 함수(추가)
# =========================
def run_render(
    excel_path: str,
    *,
    limit: int = 0,
    template_root: str = TEMPLATE_ROOT,
    icon_dir: str = ICON_DIR,
    font_path: str = FONT_PATH,
    coords_json_path: str = COORDS_JSON_PATH,
    out_dir: str = OUT_DIR,
) -> list[str]:
    """
    Streamlit/외부에서 호출하는 진입점 함수.

    excel_path: 업로드된 box_data.xlsx 경로
    limit: 0이면 전체, N이면 N행만 처리
    나머지 경로들은 기본값 유지 가능(프로젝트 구조 그대로 사용)
    return: 생성된 PDF 파일 경로 리스트
    """
    os.makedirs(out_dir, exist_ok=True)

    register_fonts(font_path)
    defaults, template_coords = load_coords(coords_json_path)

    df = pd.read_excel(excel_path)

    required = ["brand","box_type","box_group","item_code","product_name_ko","product_name_en","origin_country"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"엑셀에 필수 컬럼이 없음: {missing}")

    out_files = []
    count = 0
    for _, r in df.iterrows():
        out_files.append(
            render_row(
                defaults,
                template_coords,
                r.to_dict(),
                template_root=template_root,
                icon_dir=icon_dir,
                out_dir=out_dir,
            )
        )
        count += 1
        if limit and count >= limit:
            break

    return out_files

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=0, help="처리 행 수 제한(0이면 전체)")
    args = parser.parse_args()

    # 기존 단독 실행은 그대로 유지
    out_files = run_render(
        XLSX_PATH,
        limit=args.limit,
        template_root=TEMPLATE_ROOT,
        icon_dir=ICON_DIR,
        font_path=FONT_PATH,
        coords_json_path=COORDS_JSON_PATH,
        out_dir=OUT_DIR,
    )

    print("RENDER DONE:")
    for p in out_files[:20]:
        print(" -", p)

if __name__ == "__main__":
    main()