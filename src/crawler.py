# crawler_qoo10.py
from __future__ import annotations

import os
import pandas as pd
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta, timezone
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from item import ItemRow
from utils import *


BASE_URL = "https://m.qoo10.jp/shop/"
JPY_TO_KRW = 9.40  # 요구사항: 엔 * 9.40

# ---------- 크롤러 ----------
class Crawler:
    def __init__(self, shop_name: str, save_img_path: str = "./imgs", save_path: str = "./results"):
        self.shop_name = shop_name
        self.save_img_path = ensure_dir(save_img_path)
        self.save_path = ensure_dir(save_path)
        self.results: list[ItemRow] = []

    def setup_driver(self):
        options = Options()
        options.add_argument("--headless=new")  # UI 숨김옵션
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-gpu")
        # 모바일 플랫폼으로 변경
        mobile_emulation = {
            "deviceName": "Galaxy S8"   # 원하는 기기 이름 (Chrome DevTools 지원 기기)
        }
        options.add_experimental_option("mobileEmulation", mobile_emulation)
        self.driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=options
        )
        self.wait = WebDriverWait(self.driver, 10)

        # 현재 시각(KST)
        KST = timezone(timedelta(hours=9))
        self.search_datetime = datetime.now(KST).strftime("%Y-%m-%d_%H%M%S")
        print(f"[INIT] WebDriver ready at {self.search_datetime}")

    def collect_items(self):
        self.driver.get(f"{BASE_URL}/{self.shop_name}")
        self.wait.until(EC.presence_of_element_located((By.ID, "ul_minishop_ranking")))
        snapshot = []
        lis_sel = "ul#ul_minishop_ranking > li"
        lis = self.driver.find_elements(By.CSS_SELECTOR, lis_sel)
        count = min(len(lis), 10)

        for i in range(count):
            # 매 반복마다 '다시 찾기'로 stale 예방
            li = self.driver.find_elements(By.CSS_SELECTOR, lis_sel)[i]
            name = try_text(li, "p.text_item")
            price_jpy = only_digits(try_text(li, "strong.price_original"))
            price_krw = round(price_jpy * JPY_TO_KRW, 2)
            href = try_attr(li, "div.top_wrap a", "href")
            total_count = try_text(li, "span.option_text")
            # print("total count : ", total_count)
            snapshot.append({"name": name, 
                             "price_jpy": price_jpy, 
                             "price_krw": price_krw,
                            "product_url": href, 
                            "total_count" : total_count})

        # 상세 페이지에서 리뷰만 가져오기 (목록 페이지는 그대로 유지)
        for idx, row in enumerate(snapshot):
            self.driver.get(row["product_url"])
            review_txt = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "p.reviewstar_text"))
            ).text
            review_cnt = only_digits(review_txt)
            image_url = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button.imgLink img"))
            ).get_attribute('src')
            img_path = download_image(image_url, self.save_img_path, f"{idx}_{self.shop_name}_{self.search_datetime}")
            self.results.append(ItemRow(
                name=row["name"], 
                price_jpy=row["price_jpy"], 
                price_krw=row["price_krw"],
                review_count=review_cnt, 
                image_url=image_url,
                image_path=img_path, 
                product_url=row["product_url"],
                shop_name=self.shop_name, 
                total_count=row["total_count"]))

    def extract_revicw_count(self, item_url:str) -> int:
        """ 리뷰 카운트 함수  """
        self.driver.get(item_url)
        review_text = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'p.reviewstar_text'))
            ).text.strip()
        review_count = only_digits(review_text)
        return review_count

    def save_to_excel(self) -> str:
        if not self.results:
            print("[INFO] 저장할 결과가 없습니다.")
            return ""

        df = results_to_dataframe(self.results)
        out_path = os.path.join(self.save_path, f"qoo10_top_{self.shop_name}_{self.search_datetime}.xlsx")
        # 1) pandas로 우선 저장
        df.to_excel(out_path, index=False, engine="openpyxl")
        # 2) 저장된 파일에 이미지 삽입 + 자동 너비
        insert_images_and_autofit(out_path, self.results)
        print(f"[SAVE] Excel 저장 완료: {out_path}")
        return out_path

    def run(self):
        try:
            self.setup_driver()
            self.collect_items()
        finally:
            try:
                self.driver.quit()
            except Exception:
                pass
        self.save_to_excel()

if __name__ == "__main__":
    # 사용 예: 상위 10개 수집
    crawler = Crawler(shop_name="anua", save_img_path="./imgs", save_path="./results")
    crawler.run()
