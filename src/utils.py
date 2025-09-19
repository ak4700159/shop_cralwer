import os
import re
import requests
import pathlib
import urllib.request
import pandas as pd

from selenium.webdriver.common.by import By
from item import ItemRow
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
from dataclasses import asdict
from datetime import timedelta

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
    
def try_text(parent, sel):
    try: return parent.find_element(By.CSS_SELECTOR, sel).text.strip()
    except: return ""

def try_attr(parent, sel, attr):
    try: return parent.find_element(By.CSS_SELECTOR, sel).get_attribute(attr) or ""
    except: return ""

def fetch_image_bytes(url: str) -> bytes:
    # urllib.request 대신 selenium로도 가능하지만 간단히 requests/urllib로 받아옵니다.
    # utils에 같은 함수가 없다면 아래 구현을 사용하세요.
    import urllib.request
    with urllib.request.urlopen(url) as resp:
        return resp.read()

def guess_ext_from_url(url: str, default: str = "jpg") -> str:
    m = re.search(r"\.(png|jpe?g|gif|webp|bmp)(?:\?|$)", url, re.IGNORECASE)
    if m:
        ext = m.group(1).lower()
        return "jpg" if ext in ("jpeg", "jpg") else ext
    return default

# 엑셀 단위 변환(대략값, Calibri 11 가정)
def excel_col_width_to_pixels(width: float) -> int:
    """엑셀 열 너비 → 픽셀(대략식)"""
    return int(width * 7 + 5)

def pixels_to_row_height_points(px: float) -> float:
    """픽셀 → 엑셀 행 높이(포인트, 96DPI 가정)"""
    return px * 72 / 96

def autosize_text_columns(ws, skip_letters: set[str]):
    """텍스트 열 폭 간단 자동화 (이미지 열은 제외)"""
    for col in ws.columns:
        letter = col[0].column_letter
        if letter in skip_letters:
            continue
        max_len = 0
        for cell in col:
            v = "" if cell.value is None else str(cell.value)
            if len(v) > max_len:
                max_len = len(v)
        ws.column_dimensions[letter].width = min(60, max(10, max_len + 2))

def timer(func):
    import time
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        sec = end_time - start_time
        time_result = timedelta(seconds=sec)
        time_result = str(timedelta(seconds=sec)).split(".")
        print(f"{func.__name__} 함수 실행 시간: {sec:.2f} 초 ({time_result[0]})")
        return result
    return wrapper