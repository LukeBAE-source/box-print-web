import os
import re
import argparse
import pandas as pd

from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from pypdf import PdfReader, PdfWriter

# =========================
# 기본 경로
# =========================
XLSX_PATH = "box_data.xlsx"
TEMPLATE_ROOT = "templates"
ICON_DIR = "icons"
FONT_PATH = os.path.join("fonts", "NotoSansKR-Medium.ttf")

OUT_DIR = "output_pdf"
GUIDE_DIR = "guide_pdf"
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(GUIDE_DIR, exist_ok=True)

# =========================
# 폰트 등록
# =========================
def register_fonts():
    if not os.path.exists(FONT_PATH):
        raise FileNotFoundError(f"폰트 파일 없음: {FONT_PATH}")
    pdfmetrics.registerFont(TTFont("NotoSansKR-Medium", FONT_PATH))

# =========================
# 문자열 정규화(공백 제거 + 소문자)
# =========================
def norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s or "")).strip().lower()

# =========================
# 템플릿 찾기
# templates/<brand>/<box_type>_<box_group>.pdf
# 공백/대소문자 무시 매칭
# =========================
def find_template_pdf(brand: str, box_type: str, box_group: str) -> str:
    brand_dir = os.path.join(TEMPLATE_ROOT, brand)
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

# =========================
# 아이콘 경로: icons/icon_<country>.png
# =========================
def get_icon_path(country: str) -> str:
    c = re.sub(r"\s+", "", str(country or "")).strip().upper()
    return os.path.join(ICON_DIR, f"icon_{c}.png")

# =========================
# 브랜드별 좌표 설정 (예시)
# - 좌표 단위: PDF 포인트
# - 기준: 좌하단 (0,0)
# =========================
COORDS = {
    "iloom": {
        "cover_rects": [
            (180, 295, 160, 55),
            (560, 295, 160, 55),
            (463, 139, 60, 6),
            (717, 139, 60, 6),
        ],
        "pos": {
            "L_item_code": (180, 330),
            "L_name_ko": (180, 310),
            "L_name_en": (180, 295),
            "R_item_code": (560, 330),
            "R_name_ko": (560, 310),
            "R_name_en": (560, 295),
        },
        "icon_pos": {
            "L_origin": (463, 139, 60, 6),
            "R_origin": (717, 139, 60, 6),
        },
        "font_main_size": 26,
        "font_sub_size": 12,
    },

    "desker": {
        "cover_rects": [
            (55, 350, 520, 150),
            (675, 350, 520, 150),
            (480, 515, 170, 40),   # 데스커 예시: 상단 중앙 MADE IN 영역(이미지 기준 참고)
        ],
        "pos": {
            "L_item_code": (80, 55),
            "L_name_ko": (80, 80),
            "L_name_en": (80, 105),
            "R_item_code": (700, 55),
            "R_name_ko": (700, 80),
            "R_name_en": (700, 105),
        },
        "icon_pos": {
            "L_origin": (480, 520, 100, 28),
            "R_origin": (480, 520, 100, 28),  # 데스커는 한 곳만 찍으면 된다면 두 값을 같게 둬도 됨
        },
        "font_main_size": 22,
        "font_sub_size": 12,
    },

    "sloubed": {
        "cover_rects": [
            (60, 360, 520, 140),
            (680, 360, 520, 140),
            (520, 20, 200, 50),
            (1100, 20, 200, 50),
        ],
        "pos": {
            "L_item_code": (90, 450),
            "L_name_ko": (90, 420),
            "L_name_en": (90, 395),
            "R_item_code": (710, 450),
            "R_name_ko": (710, 420),
            "R_name_en": (710, 395),
        },
        "icon_pos": {
            "L_origin": (520, 25, 100, 32),
            "R_origin": (1100, 25, 100, 32),
        },
        "font_main_size": 26,
        "font_sub_size": 12,
    },
}

# =========================
# 오버레이 PDF 생성
# =========================
def make_overlay_pdf(overlay_path: str, page_w: float, page_h: float, cfg: dict, row: dict):
    c = canvas.Canvas(overlay_path, pagesize=(page_w, page_h))

    # 1) 가리기(흰 박스)
    c.setFillColorRGB(1, 1, 1)
    c.setStrokeColorRGB(1, 1, 1)
    for x, y, w, h in cfg["cover_rects"]:
        c.rect(x, y, w, h, fill=1, stroke=0)

    # 2) 텍스트
    sku = str(row["item_code"])
    name_ko = str(row["product_name_ko"])
    name_en = str(row["product_name_en"])

    c.setFillColorRGB(0, 0, 0)

    c.setFont("NotoSansKR-Medium", cfg["font_main_size"])
    c.drawString(*cfg["pos"]["L_item_code"], sku)
    c.drawString(*cfg["pos"]["R_item_code"], sku)

    c.setFont("NotoSansKR-Medium", cfg["font_sub_size"])
    c.drawString(*cfg["pos"]["L_name_ko"], name_ko)
    c.drawString(*cfg["pos"]["L_name_en"], name_en)
    c.drawString(*cfg["pos"]["R_name_ko"], name_ko)
    c.drawString(*cfg["pos"]["R_name_en"], name_en)

    # 3) 원산지 아이콘
    origin = str(row.get("origin_country", "")).strip()
    ip = get_icon_path(origin)

    if os.path.exists(ip):
        # 좌/우 둘 다 찍는 구조
        for k in ("L_origin", "R_origin"):
            x, y, w, h = cfg["icon_pos"][k]
            c.drawImage(ImageReader(ip), x, y, width=w, height=h, mask="auto")
    else:
        # 아이콘 없으면 텍스트 대체(원하면 삭제)
        c.setFont("NotoSansKR-Medium", 10)
        c.drawString(cfg["icon_pos"]["L_origin"][0], cfg["icon_pos"]["L_origin"][1] + 10, f"MADE IN {origin}")
        c.drawString(cfg["icon_pos"]["R_origin"][0], cfg["icon_pos"]["R_origin"][1] + 10, f"MADE IN {origin}")

    c.save()

# =========================
# 템플릿 위에 좌표 가이드(격자/라벨/사각형) 오버레이 생성
# =========================
def make_guide_overlay_pdf(overlay_path: str, page_w: float, page_h: float, cfg: dict):
    c = canvas.Canvas(overlay_path, pagesize=(page_w, page_h))

    # 격자 간격(포인트)
    step = 50

    # 옅은 격자선
    c.setStrokeColorRGB(0.8, 0.8, 0.8)
    c.setLineWidth(0.5)
    for x in range(0, int(page_w) + 1, step):
        c.line(x, 0, x, page_h)
    for y in range(0, int(page_h) + 1, step):
        c.line(0, y, page_w, y)

    # 좌표 라벨
    c.setFillColorRGB(0, 0, 1)
    c.setFont("Helvetica", 8)
    for x in range(0, int(page_w) + 1, step * 2):
        c.drawString(x + 2, 2, f"x={x}")
    for y in range(0, int(page_h) + 1, step * 2):
        c.drawString(2, y + 2, f"y={y}")

    # cover_rects 표시(빨간 테두리)
    c.setStrokeColorRGB(1, 0, 0)
    c.setLineWidth(1)
    for i, (x, y, w, h) in enumerate(cfg["cover_rects"], start=1):
        c.rect(x, y, w, h, fill=0, stroke=1)
        c.drawString(x, y + h + 2, f"cover#{i}")

    # pos 표시(초록 점)
    c.setFillColorRGB(0, 0.6, 0)
    for k, (x, y) in cfg["pos"].items():
        c.circle(x, y, 3, stroke=1, fill=1)
        c.drawString(x + 5, y + 3, k)

    # icon_pos 표시(파란 테두리)
    c.setStrokeColorRGB(0, 0, 1)
    for k, (x, y, w, h) in cfg["icon_pos"].items():
        c.rect(x, y, w, h, fill=0, stroke=1)
        c.drawString(x, y + h + 2, k)

    c.save()

# =========================
# 1행 처리(템플릿+오버레이 merge)
# =========================
def render_row(row: dict, create_guide: bool = False):
    brand = str(row["brand"]).strip()
    box_type = str(row["box_type"]).strip()
    box_group = str(row["box_group"]).strip()

    if brand not in COORDS:
        raise KeyError(f"COORDS에 없는 brand: {brand} (허용: {list(COORDS.keys())})")

    template_pdf = find_template_pdf(brand, box_type, box_group)
    reader = PdfReader(template_pdf)
    base_page = reader.pages[0]

    page_w = float(base_page.mediabox.width)
    page_h = float(base_page.mediabox.height)

    cfg = COORDS[brand]

    # 가이드 PDF 생성 모드
    if create_guide:
        overlay_path = os.path.join(GUIDE_DIR, f"__guide_overlay_{brand}_{norm(box_type)}_{norm(box_group)}.pdf")
        make_guide_overlay_pdf(overlay_path, page_w, page_h, cfg)

        overlay_reader = PdfReader(overlay_path)
        base_page.merge_page(overlay_reader.pages[0])

        out_guide = os.path.join(GUIDE_DIR, f"GUIDE_{brand}_{box_type}_{box_group}.pdf")
        w = PdfWriter()
        w.add_page(base_page)
        with open(out_guide, "wb") as f:
            w.write(f)

        try:
            os.remove(overlay_path)
        except:
            pass

        return out_guide

    # 실제 출력 모드
    sku = str(row["item_code"])
    overlay_path = os.path.join(OUT_DIR, f"__overlay_{sku}.pdf")
    make_overlay_pdf(overlay_path, page_w, page_h, cfg, row)

    overlay_reader = PdfReader(overlay_path)
    base_page.merge_page(overlay_reader.pages[0])

    out_path = os.path.join(OUT_DIR, f"{brand}_{box_type}_{box_group}_{sku}.pdf")
    w = PdfWriter()
    w.add_page(base_page)
    with open(out_path, "wb") as f:
        w.write(f)

    try:
        os.remove(overlay_path)
    except:
        pass

    return out_path

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--guide", action="store_true", help="좌표 가이드 PDF를 생성(템플릿별 1장)")
    parser.add_argument("--limit", type=int, default=0, help="처리 행 수 제한(0이면 전체)")
    args = parser.parse_args()

    register_fonts()

    df = pd.read_excel(XLSX_PATH)

    required = ["brand","box_type","box_group","item_code","product_name_ko","product_name_en","origin_country"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"엑셀에 필수 컬럼이 없음: {missing}")

    # guide 모드면: 템플릿 조합(brand/box_type/box_group)별로 1장씩만 생성
    if args.guide:
        made = set()
        out_files = []
        for _, r in df.iterrows():
            row = r.to_dict()
            key = (str(row["brand"]).strip(), str(row["box_type"]).strip(), str(row["box_group"]).strip())
            if key in made:
                continue
            made.add(key)
            out_files.append(render_row(row, create_guide=True))

        print("GUIDE DONE:")
        for p in out_files:
            print(" -", p)
        return

    # 실제 출력
    out_files = []
    it = df.iterrows()
    count = 0
    for _, r in it:
        out_files.append(render_row(r.to_dict(), create_guide=False))
        count += 1
        if args.limit and count >= args.limit:
            break

    print("RENDER DONE:")
    for p in out_files[:20]:
        print(" -", p)

if __name__ == "__main__":
    main()
