import os
import re
import json
import argparse
from datetime import datetime

TEMPLATE_ROOT = "templates"
COORDS_DIR = "coords"
COORDS_JSON_PATH = os.path.join(COORDS_DIR, "coords.json")

BRANDS = ["iloom", "desker", "sloubed"]

def norm(s: str) -> str:
    """파일명(확장자 제외)에서 공백 제거 + 소문자"""
    return re.sub(r"\s+", "", str(s or "")).strip().lower()

def list_templates():
    """
    templates/<brand>/*.pdf 스캔해서
    return: { brand: { template_key: original_filename } }
    """
    result = {b: {} for b in BRANDS}

    if not os.path.isdir(TEMPLATE_ROOT):
        raise FileNotFoundError(f"'{TEMPLATE_ROOT}' 폴더가 없습니다.")

    for brand in BRANDS:
        brand_dir = os.path.join(TEMPLATE_ROOT, brand)
        if not os.path.isdir(brand_dir):
            # 브랜드 폴더가 없으면 비워두고 진행
            continue

        for fn in os.listdir(brand_dir):
            if not fn.lower().endswith(".pdf"):
                continue
            stem = os.path.splitext(fn)[0]
            key = norm(stem)
            result[brand][key] = fn

    return result

def empty_cfg(font_main_size: int, font_sub_size: int):
    """
    템플릿별 cfg 빈 뼈대.
    좌표는 사용자가 추후 수정.
    """
    return {
        "cover_rects": [],
        "pos": {
            "L_item_code": [0, 0],
            "L_name_ko": [0, 0],
            "L_name_en": [0, 0],
            "R_item_code": [0, 0],
            "R_name_ko": [0, 0],
            "R_name_en": [0, 0]
        },
        "icon_pos": {
            "L_origin": [0, 0, 0, 0],
            "R_origin": [0, 0, 0, 0]
        },
        "font_main_size": font_main_size,
        "font_sub_size": font_sub_size
    }

def load_existing():
    if not os.path.exists(COORDS_JSON_PATH):
        return None
    with open(COORDS_JSON_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--overwrite", action="store_true",
                        help="기존 coords.json이 있어도 템플릿 섹션을 완전히 재생성(덮어쓰기)")
    parser.add_argument("--fill-zeros", action="store_true",
                        help="pos/icon_pos를 0으로 강제(기존 값이 있어도 0으로 덮어씀)")
    args = parser.parse_args()

    os.makedirs(COORDS_DIR, exist_ok=True)

    templates = list_templates()
    existing = load_existing()

    # 기본 defaults (원하면 여기 값을 바꿔서 기본 폰트 크기 등 유지)
    defaults = {
        "iloom":   {"cover_rects": [], "pos": {}, "icon_pos": {}, "font_main_size": 26, "font_sub_size": 12},
        "desker":  {"cover_rects": [], "pos": {}, "icon_pos": {}, "font_main_size": 22, "font_sub_size": 12},
        "sloubed": {"cover_rects": [], "pos": {}, "icon_pos": {}, "font_main_size": 26, "font_sub_size": 12},
    }

    # 기존 defaults 유지
    if existing and isinstance(existing.get("defaults"), dict):
        for b in BRANDS:
            if b in existing["defaults"]:
                defaults[b] = existing["defaults"][b]

    data = {
        "meta": {
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "template_root": TEMPLATE_ROOT,
            "note": "templates/<brand>/*.pdf 를 스캔해서 templates.<brand>.<template_key> 뼈대를 생성한 파일입니다."
        },
        "defaults": defaults,
        "templates": {b: {} for b in BRANDS}
    }

    # 기존 templates 유지/머지
    if existing and isinstance(existing.get("templates"), dict) and not args.overwrite:
        for b in BRANDS:
            if b in existing["templates"] and isinstance(existing["templates"][b], dict):
                data["templates"][b].update(existing["templates"][b])

    # 템플릿 파일 목록 반영(키 생성)
    for brand in BRANDS:
        for template_key, original_fn in templates.get(brand, {}).items():
            if args.overwrite or (template_key not in data["templates"][brand]):
                # 신규 생성
                dft = defaults.get(brand, {})
                data["templates"][brand][template_key] = empty_cfg(
                    int(dft.get("font_main_size", 26)),
                    int(dft.get("font_sub_size", 12))
                )

            # 선택 옵션: 좌표 0으로 강제
            if args.fill_zeros:
                dft = defaults.get(brand, {})
                data["templates"][brand][template_key] = empty_cfg(
                    int(dft.get("font_main_size", 26)),
                    int(dft.get("font_sub_size", 12))
                )

    # 템플릿 파일명 참조용 매핑(검수 편하게)
    # template_key -> 원본 파일명
    data["template_files"] = templates

    with open(COORDS_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # 요약 출력
    print(f"Saved: {COORDS_JSON_PATH}")
    for b in BRANDS:
        print(f"- {b}: {len(templates.get(b, {}))} templates")

if __name__ == "__main__":
    main()
