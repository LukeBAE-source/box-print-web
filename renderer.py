# renderer.py
import os
import re
import json
import glob
from typing import Dict, Any, Tuple
from pathlib import Path

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import Color
from pypdf import PdfReader, PdfWriter

# --------------------------------------------------
# Paths (CWD 영향 제거: 이 파일 위치 기준)
# --------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

TEMPLATE_ROOT = str(BASE_DIR / "templates")
COORDS_JSON_PATH = str(BASE_DIR / "coords" / "coords.json")
ICON_DIR = str(BASE_DIR / "icons")

FONT_MEDIUM_PATH = str(BASE_DIR / "fonts" / "NotoSansKR-Medium.ttf")
FONT_BOLD_PATH = str(BASE_DIR / "fonts" / "NotoSansKR-Bold.ttf")

FONT_MEDIUM_NAME = "NotoSansKR-Medium"
FONT_BOLD_NAME = "NotoSansKR-Bold"


# --------------------------------------------------
# Utils
# --------------------------------------------------
def normalize(s: str) -> str:
    return str(s or "").strip().lower()


def safe_filename(name: str) -> str:
    s = re.sub(r"[^\w\-.]+", "_", str(name)).strip("_")
    return s or "output.pdf"


def register_fonts():
    if os.path.exists(FONT_MEDIUM_PATH):
        pdfmetrics.registerFont(TTFont(FONT_MEDIUM_NAME, FONT_MEDIUM_PATH))
    if os.path.exists(FONT_BOLD_PATH):
        pdfmetrics.registerFont(TTFont(FONT_BOLD_NAME, FONT_BOLD_PATH))


def load_coords() -> Dict[Tuple[str, str], Dict[str, Any]]:
    if not os.path.exists(COORDS_JSON_PATH):
        raise FileNotFoundError(f"coords.json 없음: {COORDS_JSON_PATH}")

    with open(COORDS_JSON_PATH, "r", encoding="utf-8") as f:
        raw = json.load(f)

    templates = raw.get("templates", {}) or {}
    out: Dict[Tuple[str, str], Dict[str, Any]] = {}

    for brand, tmap in templates.items():
        for template_key, cfg in (tmap or {}).items():
            out[(normalize(brand), normalize(template_key))] = cfg or {}

    return out


def find_template_path(template_root: str, brand_norm: str, template_key_norm: str) -> str:
    brand_dir = os.path.join(template_root, brand_norm)
    if not os.path.isdir(brand_dir):
        raise FileNotFoundError(f"템플릿 브랜드 폴더 없음: {brand_dir}")

    candidates = glob.glob(os.path.join(brand_dir, "*.pdf"))
    for p in candidates:
        base = os.path.splitext(os.path.basename(p))[0]
        if normalize(base) == template_key_norm:
            return p

    sample = [os.path.basename(x) for x in sorted(candidates)[:30]]
    raise FileNotFoundError(
        f"템플릿 없음: {os.path.join(brand_dir, template_key_norm + '.pdf')}\n"
        f"폴더 내 PDF 예시(최대 30개): {sample}"
    )


def get_icon_path(country: str) -> str:
    if not country:
        return ""

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

    prof_name = key_map.get(key)
    prof = profiles.get(prof_name, {}) if prof_name else {}

    size = int(prof.get("size", main_size if is_main else sub_size))
    bold = bool(prof.get("bold", key in bold_keys))

    font_name = FONT_BOLD_NAME if bold else FONT_MEDIUM_NAME
    return font_name, size


def draw_text_rotated_180(c: canvas.Canvas, x: float, y: float, text: str):
    c.saveState()
    c.translate(x, y)
    c.rotate(180)
    c.drawString(0, 0, text)
    c.restoreState()


def draw_image_rotated_180(c: canvas.Canvas, x: float, y: float, w: float, h: float, path: str):
    c.saveState()
    c.translate(x + w, y + h)
    c.rotate(180)
    c.drawImage(path, 0, 0, w, h, preserveAspectRatio=True, mask="auto")
    c.restoreState()


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
    template_path = find_template_path(TEMPLATE_ROOT, b, t)

    out_dir = os.path.dirname(output_path) or "."
    os.makedirs(out_dir, exist_ok=True)

    overlay_path = output_path + ".overlay.pdf"

    try:
        reader = PdfReader(template_path)
        base_page = reader.pages[0]

        w = float(base_page.mediabox.width)
        h = float(base_page.mediabox.height)

        c = canvas.Canvas(overlay_path, pagesize=(w, h))

        # cover rects (가리기)
        cover_rects = cfg.get("cover_rects", []) or []
        if cover_rects:
            c.saveState()
            c.setFillColor(Color(1, 1, 1))
            c.setStrokeColor(Color(1, 1, 1))
            for r in cover_rects:
                if not isinstance(r, (list, tuple)) or len(r) < 4:
                    continue
                x, y, rw, rh = map(float, r[:4])
                c.rect(x, y, rw, rh, fill=1, stroke=0)
            c.restoreState()

        # text
        pos = cfg.get("pos", {}) or {}
        text_rotate_cfg = cfg.get("text_rotate_180", {}) or {}

        text_map = {
            "L_item_code": item_code,
            "L1_item_code": item_code,
            "L2_item_code": item_code,
            "R_item_code": item_code,
            "R1_item_code": item_code,
            "R2_item_code": item_code,

            "L_name_ko": name_ko,
            "L1_name_ko": name_ko,
            "L2_name_ko": name_ko,
            "R_name_ko": name_ko,
            "R1_name_ko": name_ko,
            "R2_name_ko": name_ko,

            "L_name_en": name_en,
            "L1_name_en": name_en,
            "L2_name_en": name_en,
            "R_name_en": name_en,
            "R1_name_en": name_en,
            "R2_name_en": name_en,
        }

        for key, p in pos.items():
            if not isinstance(p, (list, tuple)) or len(p) < 2:
                continue
            x, y = map(float, p[:2])

            text = text_map.get(key, "")
            if not text:
                continue

            is_main = ("name_ko" in key) or ("item_code" in key)
            font_name, font_size = resolve_font(cfg, key, is_main=is_main)
            c.setFont(font_name, font_size)

            do_rot = bool(text_rotate_cfg.get(key, False))
            if do_rot:
                draw_text_rotated_180(c, x, y, text)
            else:
                c.drawString(x, y, text)

        # icon
        icon_path = get_icon_path(origin_country)
        icon_pos = cfg.get("icon_pos", {}) or {}
        icon_rotate_cfg = cfg.get("icon_rotate_180", {}) or {}

        for key, r in icon_pos.items():
            if not isinstance(r, (list, tuple)) or len(r) < 4:
                continue
            x, y, rw, rh = map(float, r)
            do_rot = bool(icon_rotate_cfg.get(key, False))

            if icon_path and os.path.exists(icon_path):
                if do_rot:
                    draw_image_rotated_180(c, x, y, rw, rh, icon_path)
                else:
                    c.drawImage(icon_path, x, y, rw, rh, preserveAspectRatio=True, mask="auto")
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
        base_page.merge_page(overlay_page)
        writer.add_page(base_page)

        with open(output_path, "wb") as f:
            writer.write(f)

    finally:
        try:
            if os.path.exists(overlay_path):
                os.remove(overlay_path)
        except Exception:
            pass