import os
import re
import requests
import pathlib
import pandas as pd
from selenium.webdriver.common.by import By
from item import ItemRow
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
from dataclasses import asdict

def ensure_dir(path: str | os.PathLike) -> str:
    p = pathlib.Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return str(p.resolve())


def only_digits(text: str) -> int:
    """문자열에서 숫자만 뽑아 정수로 변환. 없으면 0."""
    nums = re.findall(r"\d+", text)
    if not nums:
        return 0
    return int("".join(nums))

def download_image(img_url: str, save_dir: str, filename_hint: str) -> str:
    """이미지 다운로드 (requests). 실패 시 빈 문자열 반환."""
    try:
        r = requests.get(img_url, timeout=15)
        r.raise_for_status()
        ext = ".jpg"
        # 간단한 확장자 추정
        if "png" in r.headers.get("content-type", ""):
            ext = ".png"
        safe_name = re.sub(r"[^a-zA-Z0-9_\-]+", "_", filename_hint)[:80]
        save_path = os.path.join(save_dir, f"{safe_name}{ext}")
        with open(save_path, "wb") as f:
            f.write(r.content)
        return save_path
    except Exception:
        return ""
    
def try_text(parent, sel):
    try: return parent.find_element(By.CSS_SELECTOR, sel).text.strip()
    except: return ""

def try_attr(parent, sel, attr):
    try: return parent.find_element(By.CSS_SELECTOR, sel).get_attribute(attr) or ""
    except: return ""

# --- 엑셀 유틸: 이미지 삽입 + 자동 너비 ---
def insert_images_and_autofit(xlsx_path: str, rows: list[ItemRow],
                               img_col_letter: str = "A",
                               thumb_size: tuple[int, int] = (120, 120)):
    """
    xlsx_path: pandas가 저장한 엑셀 경로
    rows: self.results (ItemRow 리스트)
    img_col_letter: 이미지를 넣을 컬럼 (기본 A)
    thumb_size: 삽입용 썸네일 크기
    """
    wb = load_workbook(xlsx_path)
    ws = wb.active

    # 1) 이미지용 컬럼을 맨 앞에 삽입하고 헤더 추가
    ws.insert_cols(1)
    ws[f"{img_col_letter}1"] = "Image"

    # 2) 각 행에 이미지 추가 (행 높이도 조정)
    for r_idx, item in enumerate(rows, start=2):  # 헤더 다음부터
        p = getattr(item, "image_path", "") or ""
        if not p or not os.path.exists(p):
            continue
        try:
            # 썸네일 생성(원본 보존)
            with PILImage.open(p) as im:
                im.thumbnail(thumb_size)
                thumb_path = p  # 원본 위에 덮지 않고 임시 파일을 쓰고 싶다면 별도 경로 지정
                # 필요 시 별도 파일로 저장하고 싶으면 아래 두 줄 사용
                # base, ext = os.path.splitext(p)
                # thumb_path = f"{base}_thumb{ext}"
                im.save(thumb_path)

            xlimg = XLImage(thumb_path)
            ws.add_image(xlimg, f"{img_col_letter}{r_idx}")
            # 이미지가 보이도록 행 높이 살짝 늘리기
            ws.row_dimensions[r_idx].height = max(ws.row_dimensions[r_idx].height or 0, 90)
        except Exception:
            # 이미지 삽입 실패는 무시하고 계속
            pass

    # 3) 이미지 컬럼 고정 너비
    ws.column_dimensions[img_col_letter].width = 18  # 대략 120px 정도 보기 좋게

    # 4) 나머지 컬럼 자동 너비(문자열 길이 기반)
    #    많이 쓰는 패턴: 각 컬럼의 최대 텍스트 길이 + 여백으로 width 설정
    for col in ws.columns:
        col_letter = col[0].column_letter
        if col_letter == img_col_letter:
            continue
        max_len = 0
        for cell in col:
            v = "" if cell.value is None else str(cell.value)
            if len(v) > max_len:
                max_len = len(v)
        # 폰트/픽셀 차이를 감안해 약간의 보정치 추가
        ws.column_dimensions[col_letter].width = max(10, min(60, max_len + 2))
    wb.save(xlsx_path)


def results_to_dataframe(rows: list[ItemRow]) -> pd.DataFrame:
    """엑셀에 넣을 DataFrame 생성(이미지 컬럼은 썸네일로 따로 넣으므로 경로는 유지)"""
    df = pd.DataFrame([asdict(r) for r in rows])
    # 보기 좋은 컬럼 순서로 재배치 (원하면 수정)
    preferred = [
        "shop_name", "name", "price_jpy", "price_krw", "review_count",
        "total_count", "product_url", "image_url", "image_path"
    ]
    df = df[[c for c in preferred if c in df.columns]]
    return df
