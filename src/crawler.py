# crawler_qoo10.py
from __future__ import annotations

import os
import io
from datetime import datetime, timedelta, timezone
from typing import List, Dict, Any

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from item import ItemRow
from image import Image
from utils import *
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage


BASE_URL = "https://m.qoo10.jp/shop/"
JPY_TO_KRW = 9.40
VALID_PERIODS = {"D": "日", "W": "週", "M": "月"}

class Crawler:
    def __init__(self, shop_name: str, save_path: str = "./results", period: str = "W"):
        self.shop_name:     str = shop_name
        self.period:        str = period.upper()
        if self.period not in VALID_PERIODS:
            raise ValueError(f"period must be one of {list(VALID_PERIODS.keys())}")
        self.save_root:     str = ensure_dir(save_path)  # 폴더는 존재 보장만 하고, 하위 디렉토리 생성은 없음
        self.results:       List[ItemRow] = []
        self._snap:         List[Dict[str, Any]] = []
        # {"idx" : 검색된 이미지 인덱스(일종의 순서), "bytes" : 실제 이미지 데이터, "ext" : 파일 형식}
        self.images:       List[Image] = []  

    def setup_driver(self):
        """ chrom driver 설정 함수 """
        options = Options()
        # enable-logging: 크롬 실행 시 콘솔에 뜨는 디버깅 로그(DevTools 프로토콜 등)를 끔.
        # enable-automation: 크롬 우측 상단에 뜨는 "Chrome is being controlled by automated software" 경고 메시지를 숨김.
        options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
        # 로그 출력 정도 수정, 0=INFO..3=ERROR만 
        options.add_argument("--log-level=2")  
        # pageLoadStrategy 추가 -> 페이지 로딩 전략 변경
            # normal: 기본값. 모든 리소스 로딩 완료까지 대기.
            # eager: HTML만 로드되면 제어권 반환.
            # none: 로딩 완료를 기다리지 않고 즉시 반환.
        options.set_capability("pageLoadStrategy", "eager")
        # 화면 표시 없이 백그라운드에서 브라우저 실행.
        options.add_argument("--headless=new")
        # 보안 샌드박스 모드 off
        options.add_argument("--no-sandbox")
        # headless 모드에서 gpu 가속이 필요없음(화면 랜더링을 하지 않기 때문에)
        options.add_argument("--disable-gpu")
        # 밑에 3개는 이미지 로딩 관련 설정
            # --disable-images: 크롬 플래그 수준에서 이미지 비활성화.
            # prefs 설정: 사용자 프로필에서 "이미지 로드 안 함"으로 지정.
            # blink-settings: 렌더링 엔진에 직접 "이미지 끔" 설정.
        options.add_argument('--disable-images')
        options.add_experimental_option("prefs", {'profile.managed_default_content_settings.images': 2})
        options.add_argument('--blink-settings=imagesEnabled=false')
        # 모바일 설정
        options.add_experimental_option("mobileEmulation", {"deviceName": "Galaxy S8"})
        self.driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=options
        )

        self.wait = WebDriverWait(self.driver, 10)
        KST = timezone(timedelta(hours=9))
        self.search_datetime = datetime.now(KST).strftime("%Y-%m-%d_%H%M%S")
        print(f"[INIT] WebDriver ready at {self.search_datetime}")

    def select_period(self):
        self.wait.until(EC.presence_of_element_located((By.ID, "ul_ranking_period")))
        old_list = self.wait.until(EC.presence_of_element_located((By.ID, "ul_minishop_ranking")))
        btn_sel = f'#ul_ranking_period button[value="{self.period}"]'
        btn = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, btn_sel)))
        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        try:
            btn.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", btn)
        try:
            WebDriverWait(self.driver, 10).until(EC.staleness_of(old_list))
        except Exception:
            self.wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, f'#ul_ranking_period li.selected button[value="{self.period}"]')
            ))
        self.wait.until(EC.presence_of_element_located((By.ID, "ul_minishop_ranking")))
        print(f"[PERIOD] switched to {self.period} ({VALID_PERIODS[self.period]})")

    def collect_items(self):
        self.driver.get(f"{BASE_URL}/{self.shop_name}")
        self.select_period()
        self.wait.until(EC.presence_of_element_located((By.ID, "ul_minishop_ranking")))
        lis_sel = "ul#ul_minishop_ranking > li"
        lis = self.driver.find_elements(By.CSS_SELECTOR, lis_sel)
        count = min(len(lis), 10)

        for i in range(count):
            li = self.driver.find_elements(By.CSS_SELECTOR, lis_sel)[i]
            name = try_text(li, "p.text_item")
            price_jpy = only_digits(try_text(li, "strong.price_original"))
            price_krw = round(price_jpy * JPY_TO_KRW, 2)
            href = try_attr(li, "div.top_wrap a", "href")
            total_count = try_text(li, "span.option_text")
            self._snap.append({
                "idx": i,
                "name": name,
                "price_jpy": price_jpy,
                "price_krw": price_krw,
                "product_url": href,
                "total_count": total_count
            })

        for row in self._snap:
            self.driver.get(row["product_url"])
            try:
                review_txt = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "p.reviewstar_text"))
                ).text
            except Exception:
                review_txt = "0"
            review_cnt = only_digits(review_txt)

            img_el = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button.imgLink img"))
            )
            image_url = img_el.get_attribute("src")
            img_bytes = fetch_image_bytes(image_url)
            ext = guess_ext_from_url(image_url)

            # 디스크에 이미지 저장하지 않음(메모리 전용)
            self.results.append(ItemRow(
                name=row["name"],
                price_jpy=row["price_jpy"],
                price_krw=row["price_krw"],
                review_count=review_cnt,
                image_url=image_url,
                image_path="",  # 저장하지 않으므로 빈 문자열
                product_url=row["product_url"],
                shop_name=self.shop_name,
                total_count=row["total_count"]
            ))
            self.images.append(
                Image(
                    idx=row["idx"],
                    img_bytes= img_bytes,
                    ext=ext
                ))

    def save_outputs(self) -> str:
        if not self.results:
            print("[INFO] 저장할 결과가 없습니다.")
            return ""

        # 결과 파일은 save_path 바로 아래에 저장(하위 디렉토리 생성 X)
        xlsx_path = os.path.join(self.save_root, f"qoo10_top_{self.shop_name}_{self.search_datetime}.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "ranking"
        headers_xlsx = ["Rank", "Name", "Price(JPY)", "Price(KRW)", "Reviews",
                        "Product URL", "Shop", "Total Count", "Image"]
        ws.append(headers_xlsx)

        # 이미지가 들어갈 열(I) 폭 지정(적절한 썸네일 폭)
        img_col_letter = "I"
        ws.column_dimensions[img_col_letter].width = 25  # 필요 시 조정 가능
        target_col_px = excel_col_width_to_pixels(ws.column_dimensions[img_col_letter].width)

        for i, r in enumerate(self.results, start=1):
            ws.append([i, r.name, r.price_jpy, r.price_krw, r.review_count,
                       r.product_url, r.shop_name, r.total_count, ""])

            img_info = self.images[i-1]
            img = XLImage(io.BytesIO(img_info["bytes"]))

            # 열 폭 기준으로 비율 유지 축소(너비가 열폭보다 크면 축소, 작으면 원본 유지)
            orig_w, orig_h = float(img.width), float(img.height)
            scale = min(1.0, target_col_px / orig_w) if orig_w > 0 else 1.0
            img.width = orig_w * scale
            img.height = orig_h * scale

            # 최종 이미지 높이에 맞춰 행 높이를 자동 설정
            row_idx = ws.max_row
            ws.row_dimensions[row_idx].height = pixels_to_row_height_points(img.height)

            # 이미지 삽입
            anchor = f"{img_col_letter}{row_idx}"
            ws.addimage(img, anchor)

        # 텍스트 열 자동 너비(이미지 열 I는 제외)
        autosize_text_columns(ws, skip_letters={img_col_letter})

        wb.save(xlsx_path)
        print(f"[SAVE] XLSX(이미지 포함) 저장 완료: {xlsx_path}")
        return xlsx_path

    @timer
    def run(self):
        try:
            self.setup_driver()
            self.collect_items()
        finally:
            try:
                self.driver.quit()
            except Exception:
                pass
        # 테스트 시 주석을 해제하고 제대로 저장되는지 확인
        # self.save_outputs()


if __name__ == "__main__":
    # 테스트 코드
    crawler = Crawler(shop_name="anua", save_path="./results", period="D")
    crawler.run()