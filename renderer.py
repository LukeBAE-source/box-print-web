import os
import re
import json
from typing import Dict, Any

import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import Color
from pypdf import PdfReader, PdfWriter

TEMPLATE_ROOT = "templates"
COORDS_JSON_PATH = os.path.join("coords", "coords.json")
ICON_DIR = "icons"

FONT_MEDIUM_PATH = os.path.join("fonts", "NotoSansKR-Medium.ttf")
FONT_BOLD_PATH = os.path.join("fonts", "NotoSansKR-Bold.ttf")

FONT_MEDIUM_NAME = "NotoSansKR-Medium"
FONT_BOLD_NAME = "NotoSansKR-Bold"


# --------------------------------------------------
# Utils
# --------------------------------------------------

def normalize(s: str) -> str:
    return str(s or "").strip().lower()


def register_fonts():
    if os.path.exists(FONT_MEDIUM_PATH):
        pdfmetrics.registerFont(TTFont(FONT_MEDIUM_NAME, FONT_MEDIUM_PATH))
    if os.path.exists(FONT_BOLD_PATH):
        pdfmetrics.registerFont(TTFont(FONT_BOLD_NAME, FONT_BOLD_PATH))


def load_coords():
    with open(COORDS_JSON_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)

    templates = data.get("templates", {}) or {}
    result = {}

    for brand, mp in templates.items():
        b = normalize(brand)
        if not isinstance(mp, dict):
            continue
        for key, cfg in mp.items():
            result[(b, normalize(key))] = cfg

    return result


def get_icon_path(country: str) -> str:
    if not country:
        return ""

    # icons 폴더가 없으면 icon 대신 텍스트로 대체 (예외 방지)
    if not os.path.isdir(ICON_DIR):
        return ""

    key = re.sub(r"\s+", "", country).lower()

    for fn in os.listdir(ICON_DIR):
        if not fn.lower().endswith(".png"):
            continue
        if fn.lower().startswith("icon_"):
            name = fn[5:-4].lower()
            if name == key:
                return os.path.join(ICON_DIR, fn)

    return ""


def resolve_font(cfg: Dict[str, Any], key: str, is_main: bool):
    main_size = int(cfg.get("font_main_size", 26))
    sub_size = int(cfg.get("font_sub_size", 12))

    profiles = cfg.get("font_profiles", {}) or {}
    key_map = cfg.get("font_key_profile", {}) or {}
    bold_keys = set(cfg.get("font_bold_keys", []) or [])

    size = main_size if is_main else sub_size
    bold = key in bold_keys

    profile_name = key_map.get(key)
    if profile_name and profile_name in profiles:
        profile = profiles[profile_name]
        size = int(profile.get("size", size))
        bold = bool(profile.get("bold", False)) or bold

    if bold and FONT_BOLD_NAME in pdfmetrics.getRegisteredFontNames():
        return FONT_BOLD_NAME, size

    return FONT_MEDIUM_NAME, size


def draw_text_rotated_180(c, x, y, text):
    c.saveState()
    c.translate(x, y)
    c.rotate(180)
    c.drawString(0, 0, text)
    c.restoreState()


def draw_image_rotated_180(c, x, y, w, h, image_path):
    """
    (x,y,w,h) 영역에 이미지를 180도 회전해서 배치
    ReportLab은 drawImage에 rotate 옵션이 없으므로 좌표계를 회전시켜 그린다.
    """
    c.saveState()
    c.translate(x + w, y + h)  # 박스의 우상단으로 이동 후 180도 회전
    c.rotate(180)
    c.drawImage(image_path, 0, 0, w, h, preserveAspectRatio=True)
    c.restoreState()


def safe_filename(s: str) -> str:
    s = str(s or "").strip()
    s = re.sub(r'[\\/*?:"<>|]+', "_", s)  # Windows 금지문자 치환
    s = re.sub(r"\s+", " ", s).strip()
    return s or "output"


# --------------------------------------------------
# Main render function
# --------------------------------------------------

def run_render(
    brand: str,
    template_key: str,
    item_code: str,
    name_ko: str,
    name_en: str,
    origin_country: str,
    output_path: str
):
    register_fonts()
    coords = load_coords()

    b = normalize(brand)
    t = normalize(template_key)

    if (b, t) not in coords:
        raise ValueError(f"coords 없음: {b}/{t}")

    cfg = coords[(b, t)]

    template_path = os.path.join(TEMPLATE_ROOT, b, f"{t}.pdf")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"템플릿 없음: {os.path.abspath(template_path)}")

    # output 폴더 보장
    out_dir = os.path.dirname(output_path) or "."
    os.makedirs(out_dir, exist_ok=True)

    abs_out = os.path.abspath(output_path)
    print("CWD:", os.getcwd())
    print("OUTPUT:", abs_out)

    overlay_path = output_path + ".overlay.pdf"

    try:
        reader = PdfReader(template_path)
        page = reader.pages[0]

        w = float(page.mediabox.width)
        h = float(page.mediabox.height)

        c = canvas.Canvas(overlay_path, pagesize=(w, h))

        # --------------------------------------------------
        # COVER (가리기) - 텍스트/아이콘보다 먼저 그려야 함
        # --------------------------------------------------
        cover_rects = cfg.get("cover_rects", []) or []
        if cover_rects:
            c.saveState()
            c.setFillColor(Color(1, 1, 1))    # 흰색
            c.setStrokeColor(Color(1, 1, 1))  # 테두리도 흰색(표시 안 되게)
            for r in cover_rects:
                if not isinstance(r, (list, tuple)) or len(r) < 4:
                    continue
                x, y, rw, rh = map(float, r[:4])
                if rw <= 0 or rh <= 0:
                    continue
                c.rect(x, y, rw, rh, stroke=0, fill=1)
            c.restoreState()

        pos = cfg.get("pos", {}) or {}
        icon_pos = cfg.get("icon_pos", {}) or {}
        rotate_cfg = cfg.get("rotate_180", {}) or {}
        icon_rotate_cfg = cfg.get("icon_rotate_180", {}) or {}

        # --------------------------------------------------
        # TEXT
        # --------------------------------------------------
        for key, xy in pos.items():
            if not isinstance(xy, (list, tuple)) or len(xy) < 2:
                continue

            x, y = float(xy[0]), float(xy[1])
            if x == 0 and y == 0:
                continue

            if "item_code" in key:
                text = item_code
                is_main = True
            elif key.endswith("_name_ko"):
                text = name_ko
                is_main = False
            elif key.endswith("_name_en"):
                text = name_en
                is_main = False
            else:
                continue

            font_name, size = resolve_font(cfg, key, is_main)
            c.setFont(font_name, size)

            if rotate_cfg.get(key, False):
                draw_text_rotated_180(c, x, y, text)
            else:
                c.drawString(x, y, text)

        # --------------------------------------------------
        # ICON
        # --------------------------------------------------
        icon_path = get_icon_path(origin_country)

        for key, r in icon_pos.items():
            if not isinstance(r, (list, tuple)) or len(r) < 4:
                continue

            x, y, rw, rh = map(float, r)
            do_rot = bool(icon_rotate_cfg.get(key, False))

            if icon_path and os.path.exists(icon_path):
                if do_rot:
                    draw_image_rotated_180(c, x, y, rw, rh, icon_path)
                else:
                    c.drawImage(icon_path, x, y, rw, rh, preserveAspectRatio=True)
            else:
                c.setFont(FONT_MEDIUM_NAME, 8)
                if do_rot:
                    draw_text_rotated_180(c, x, y, origin_country or "")
                else:
                    c.drawString(x, y, origin_country or "")

        c.showPage()
        c.save()

        overlay_reader = PdfReader(overlay_path)
        overlay_page = overlay_reader.pages[0]

        writer = PdfWriter()
        page.merge_page(overlay_page)
        writer.add_page(page)

        with open(output_path, "wb") as f:
            writer.write(f)

        if not os.path.exists(output_path):
            raise RuntimeError(f"output 파일이 생성되지 않음: {abs_out}")
        if os.path.getsize(output_path) == 0:
            raise RuntimeError(f"output 파일 0바이트: {abs_out}")

        print("SAVED:", abs_out, "size:", os.path.getsize(output_path))

    finally:
        if os.path.exists(overlay_path):
            try:
                os.remove(overlay_path)
            except Exception as e:
                print("overlay 삭제 실패:", overlay_path, e)


# --------------------------------------------------
# Excel Batch Mode
# --------------------------------------------------

def run_batch_from_excel(excel_path="box_data.xlsx", output_dir="output_pdf"):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"엑셀 파일 없음: {os.path.abspath(excel_path)}")

    os.makedirs(output_dir, exist_ok=True)

    df = pd.read_excel(excel_path, dtype=str).fillna("")

    required_cols = [
        "brand", "box_type", "box_group", "item_code",
        "product_name_ko", "product_name_en", "origin_country"
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"엑셀 필수 컬럼 누락: {missing} / 현재 컬럼: {list(df.columns)}")

    success = 0
    fail = 0

    for idx, row in df.iterrows():
        brand = row.get("brand", "").strip()
        box_type = row.get("box_type", "").strip()
        box_group = row.get("box_group", "").strip()
        item_code = row.get("item_code", "").strip()
        name_ko = row.get("product_name_ko", "").strip()
        name_en = row.get("product_name_en", "").strip()
        origin_country = row.get("origin_country", "").strip()

        if not brand or not box_type or not box_group or not item_code:
            continue

        template_key = f"{box_type}_{box_group}".lower()

        filename = safe_filename(f"{brand}_{template_key}_{item_code}.pdf")
        output_path = os.path.join(output_dir, filename)

        try:
            run_render(
                brand=brand,
                template_key=template_key,
                item_code=item_code,
                name_ko=name_ko,
                name_en=name_en,
                origin_country=origin_country,
                output_path=output_path,
            )
            success += 1
            print(f"[OK] row {idx+2} → {output_path}")
        except Exception as e:
            fail += 1
            print(f"[FAIL] row {idx+2} → {e}")

    print(f"\n완료: 성공 {success}건 / 실패 {fail}건")


# --------------------------------------------------
# Main Entry
# --------------------------------------------------

if __name__ == "__main__":
    # renderer.py 실행 시: box_data.xlsx 기반으로 output_pdf/에 일괄 생성
    run_batch_from_excel()
