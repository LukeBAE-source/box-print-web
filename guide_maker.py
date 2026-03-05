import os
import json
import argparse
from typing import Dict, Tuple, Any

from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color, red, blue, black
from pypdf import PdfReader, PdfWriter

TEMPLATE_ROOT = "templates"
COORDS_JSON_PATH = os.path.join("coords", "coords.json")
GUIDE_ROOT = "guide_pdf"

GRID_SPACING = 10
GRID_MAJOR_EVERY = 5
GRID_LABEL_FONT = 7


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def normalize(s: str) -> str:
    return str(s or "").strip().lower()


def load_coords() -> Dict[Tuple[str, str], Dict[str, Any]]:
    path = COORDS_JSON_PATH
    raw = open(path, "r", encoding="utf-8").read()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        # context 출력
        lines = raw.splitlines()
        start = max(1, e.lineno - 5)
        end = min(len(lines), e.lineno + 5)

        print("\n[coords.json JSONDecodeError]")
        print(f"- message : {e.msg}")
        print(f"- line    : {e.lineno}")
        print(f"- column  : {e.colno}")
        print("\n[Context]")
        for ln in range(start, end + 1):
            prefix = ">>" if ln == e.lineno else "  "
            print(f"{prefix} {ln:4d}: {lines[ln-1]}")
            if ln == e.lineno:
                print(" " * (e.colno + 7) + "^")
        raise

    templates = data.get("templates", {}) or {}
    result = {}
    for brand, mp in templates.items():
        b = normalize(brand)
        if not isinstance(mp, dict):
            continue
        for key, cfg in mp.items():
            result[(b, normalize(key))] = cfg
    return result

def find_templates() -> Dict[Tuple[str, str], str]:
    result = {}

    if not os.path.isdir(TEMPLATE_ROOT):
        return result

    for brand in os.listdir(TEMPLATE_ROOT):
        brand_dir = os.path.join(TEMPLATE_ROOT, brand)
        if not os.path.isdir(brand_dir):
            continue

        b = normalize(brand)

        for fn in os.listdir(brand_dir):
            if not fn.lower().endswith(".pdf"):
                continue

            stem = os.path.splitext(fn)[0]
            result[(b, normalize(stem))] = os.path.join(brand_dir, fn)

    return result


def draw_grid_with_axes(c: canvas.Canvas, w: float, h: float):
    minor_color = Color(0.85, 0.85, 0.85)
    major_color = Color(0.65, 0.65, 0.65)

    # Minor grid
    c.setStrokeColor(minor_color)
    c.setLineWidth(0.25)

    x = 0
    while x <= w:
        c.line(x, 0, x, h)
        x += GRID_SPACING

    y = 0
    while y <= h:
        c.line(0, y, w, y)
        y += GRID_SPACING

    # Major grid + axis numbers
    c.setStrokeColor(major_color)
    c.setLineWidth(0.5)
    c.setFont("Helvetica", GRID_LABEL_FONT)

    step_major = GRID_SPACING * GRID_MAJOR_EVERY

    x = 0
    while x <= w:
        c.line(x, 0, x, h)
        c.drawString(x + 2, 2, f"{int(x)}")
        x += step_major

    y = 0
    while y <= h:
        c.line(0, y, w, y)
        c.drawString(2, y + 2, f"{int(y)}")
        y += step_major

    c.setFillColor(black)
    c.drawString(3, 3, "0,0")


def make_guide(template_path: str, out_path: str, cfg: Dict[str, Any]):
    reader = PdfReader(template_path)
    page = reader.pages[0]

    w = float(page.mediabox.width)
    h = float(page.mediabox.height)

    overlay_path = out_path + ".overlay.pdf"
    c = canvas.Canvas(overlay_path, pagesize=(w, h))

    # 1️⃣ Grid
    draw_grid_with_axes(c, w, h)

    # 2️⃣ cover_rects (빨간색 + 번호)
    c.setStrokeColor(red)
    c.setFillColor(red)
    c.setLineWidth(0.7)
    c.setFont("Helvetica", 8)

    cover_rects = cfg.get("cover_rects", []) or []

    for idx, r in enumerate(cover_rects, start=1):

        if not isinstance(r, (list, tuple)) or len(r) < 4:
            continue

        x, y, rw, rh = map(float, r)

        # 박스
        c.rect(x, y, rw, rh, stroke=1, fill=0)

        # 번호(박스 밖 상단 우선)
        label = str(idx)
        pad = 3
        label_h = 9

        lx = x
        ly = y + rh + pad

        if ly + label_h > h:
            ly = y - label_h - pad

        if ly < 0:
            lx = x + rw + pad
            ly = y + rh - label_h

        c.drawString(lx, ly, label)

    # 3️⃣ pos (파란색)
    c.setStrokeColor(blue)
    c.setFillColor(blue)
    c.setLineWidth(0.5)
    c.setFont("Helvetica", 7)

    pos = cfg.get("pos", {}) or {}

    for key, xy in pos.items():
        if not isinstance(xy, (list, tuple)) or len(xy) < 2:
            continue

        x, y = float(xy[0]), float(xy[1])

        if x == 0 and y == 0:
            continue

        c.circle(x, y, 1.3, stroke=1, fill=1)
        c.drawString(x + 3, y + 2, str(key))

    c.showPage()
    c.save()

    # merge overlay
    overlay_reader = PdfReader(overlay_path)
    overlay_page = overlay_reader.pages[0]

    writer = PdfWriter()
    page.merge_page(overlay_page)
    writer.add_page(page)

    with open(out_path, "wb") as f:
        writer.write(f)

    os.remove(overlay_path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--all", action="store_true")
    parser.add_argument("--brand", type=str, default="")
    parser.add_argument("--template", type=str, default="")
    args = parser.parse_args()

    ensure_dir(GUIDE_ROOT)

    coords = load_coords()
    templates = find_templates()

    targets = []

    if args.all:
        for (b, t), path in templates.items():
            targets.append((b, t, path))
    else:
        b = normalize(args.brand)
        t = normalize(args.template)

        if not b or not t:
            print("사용법: --all 또는 --brand <b> --template <t>")
            return

        if (b, t) not in templates:
            print("템플릿 없음")
            return

        targets.append((b, t, templates[(b, t)]))

    for b, t, path in targets:
        cfg = coords.get((b, t))

        if not cfg:
            print(f"[SKIP] coords 없음: {b}/{t}")
            continue

        brand_dir = os.path.join(GUIDE_ROOT, b)
        ensure_dir(brand_dir)

        out_path = os.path.join(brand_dir, f"GUIDE_{b}_{t}.pdf")
        make_guide(path, out_path, cfg)

        print(f"[OK] {out_path}")


if __name__ == "__main__":
    main()