import os
import re
import json
import csv
import argparse

from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter

TEMPLATE_ROOT = "templates"
COORDS_JSON_PATH = os.path.join("coords", "coords.json")

GUIDE_ROOT = "guide_pdf"
os.makedirs(GUIDE_ROOT, exist_ok=True)


def norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s or "")).strip().lower()


def load_coords():
    if not os.path.exists(COORDS_JSON_PATH):
        raise FileNotFoundError(f"좌표 JSON 없음: {COORDS_JSON_PATH}")

    with open(COORDS_JSON_PATH, "r", encoding="utf-8") as f:
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


def make_guide_overlay_pdf(overlay_path: str, page_w: float, page_h: float, cfg: dict | None, note: str | None = None):
    c = canvas.Canvas(overlay_path, pagesize=(page_w, page_h))

    # grid
    step = 10
    c.setStrokeColorRGB(0.85, 0.85, 0.85)
    c.setLineWidth(0.5)
    for x in range(0, int(page_w) + 1, step):
        c.line(x, 0, x, page_h)
    for y in range(0, int(page_h) + 1, step):
        c.line(0, y, page_w, y)

    # labels
    c.setFillColorRGB(0, 0, 1)
    c.setFont("Helvetica", 8)
    for x in range(0, int(page_w) + 1, step * 2):
        c.drawString(x + 2, 2, f"x={x}")
    for y in range(0, int(page_h) + 1, step * 2):
        c.drawString(2, y + 2, f"y={y}")

    # note
    if note:
        c.setFillColorRGB(1, 0, 0)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(20, page_h - 30, note)

    if not cfg:
        c.save()
        return

    cover_rects = cfg.get("cover_rects", []) or []
    pos = cfg.get("pos", {}) or {}
    icon_pos = cfg.get("icon_pos", {}) or {}

    # cover_rects (red)
    c.setStrokeColorRGB(1, 0, 0)
    c.setLineWidth(1)
    for i, rect in enumerate(cover_rects, start=1):
        x, y, w, h = rect
        c.rect(x, y, w, h, fill=0, stroke=1)
        c.setFillColorRGB(1, 0, 0)
        c.setFont("Helvetica", 8)
        c.drawString(x, y + h + 2, f"cover#{i}")

    # pos (green)
    c.setFillColorRGB(0, 0.6, 0)
    c.setStrokeColorRGB(0, 0.6, 0)
    c.setFont("Helvetica", 8)
    for k, xy in pos.items():
        x, y = xy
        c.circle(x, y, 3, stroke=1, fill=1)
        c.drawString(x + 5, y + 3, k)

    # icon_pos (blue)
    c.setStrokeColorRGB(0, 0, 1)
    c.setFillColorRGB(0, 0, 1)
    c.setFont("Helvetica", 8)
    for k, rect in icon_pos.items():
        x, y, w, h = rect
        c.rect(x, y, w, h, fill=0, stroke=1)
        c.drawString(x, y + h + 2, k)

    c.save()


def scan_templates():
    items = []
    if not os.path.isdir(TEMPLATE_ROOT):
        raise FileNotFoundError("templates 폴더가 없습니다.")

    for brand in os.listdir(TEMPLATE_ROOT):
        brand_dir = os.path.join(TEMPLATE_ROOT, brand)
        if not os.path.isdir(brand_dir):
            continue

        for fn in os.listdir(brand_dir):
            if not fn.lower().endswith(".pdf"):
                continue
            path = os.path.join(brand_dir, fn)
            stem = os.path.splitext(fn)[0]
            template_key = norm(stem)
            items.append((brand, fn, path, template_key))

    return items


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--all", action="store_true", help="templates의 모든 pdf 템플릿 가이드 생성")
    args = parser.parse_args()

    if not args.all:
        print("사용법: python guide_maker.py --all")
        return

    defaults, template_coords = load_coords()
    templates = scan_templates()

    missing_rows = []
    total = 0

    for brand, fn, path, template_key in templates:
        brand_out_dir = os.path.join(GUIDE_ROOT, brand)
        os.makedirs(brand_out_dir, exist_ok=True)

        # cfg 선택: template-specific > brand default > None
        cfg = template_coords.get((brand, template_key))
        note = None
        if cfg is None:
            cfg = defaults.get(brand)
            note = "NO TEMPLATE COORDS (using brand defaults)" if cfg else "NO COORDS (grid only)"
            missing_rows.append([brand, template_key, fn])

        reader = PdfReader(path)
        base_page = reader.pages[0]
        page_w = float(base_page.mediabox.width)
        page_h = float(base_page.mediabox.height)

        overlay_tmp = os.path.join(brand_out_dir, "__tmp_overlay.pdf")
        make_guide_overlay_pdf(overlay_tmp, page_w, page_h, cfg, note)

        overlay_reader = PdfReader(overlay_tmp)
        base_page.merge_page(overlay_reader.pages[0])

        out_path = os.path.join(brand_out_dir, f"GUIDE_{fn}")
        writer = PdfWriter()
        writer.add_page(base_page)
        with open(out_path, "wb") as f:
            writer.write(f)

        try:
            os.remove(overlay_tmp)
        except:
            pass

        total += 1

    # missing CSV
    if missing_rows:
        csv_path = os.path.join(GUIDE_ROOT, "missing_templates.csv")
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["brand", "template_key", "template_filename"])
            w.writerows(missing_rows)
        print(f"missing_templates.csv 생성: {csv_path}")

    print(f"GUIDE 생성 완료: {total} files")
    print(f"가이드 폴더: {GUIDE_ROOT}/<brand>/GUIDE_*.pdf")


if __name__ == "__main__":
    main()
