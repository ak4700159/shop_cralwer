from cralwer_manager import CrawlerManager
import queue
import FreeSimpleGUI as sg
import pandas as pd 
import traceback
import os 
import io
from datetime import datetime
from openpyxl.drawing.image import Image as XLImage
from utils import excel_col_width_to_pixels, pixels_to_row_height_points
from openpyxl import Workbook


def run_all_sequential(window: sg.Window, shops: list[str], outdir: str, period: str, log_q: queue.Queue):
    try:
        # {"shop": s, "path": latest} 형식으로 상점 크롤링 결과를 저장
        results = []
        # 크롤링마다 
        data_results = []
        images = []
        manager = CrawlerManager.get(save_path=outdir, period=period)

        # ✅ 통합 워크북/시트 생성(헤더 1회 작성)
        wb_all = Workbook()
        ws_all = wb_all.active
        ws_all.title = "ranking"
        headers_xlsx = ["Rank", "Name", "Price(JPY)", "Price(KRW)", "Reviews",
                        "Product URL", "Shop", "Total Count", "Image"]
        ws_all.append(headers_xlsx)
        # 이미지 열 기본 폭(개별 시트와 동일 기준)
        ws_all.column_dimensions["I"].width = 25

        for s in shops:
            try:
                log_q.put(f"[START] {s} 수집 시작 (period={period})")
                crawler = manager.run_shop(s)  # ✅ 크롤러 인스턴스 받기

                # 최신 개별 엑셀 경로 찾기(기존 그대로)
                latest = None
                for f in os.listdir(outdir):
                    if f.startswith(f"qoo10_top_{s}_") and f.endswith(".xlsx"):
                        p = os.path.join(outdir, f)
                        if latest is None or os.path.getmtime(p) > os.path.getmtime(latest):
                            latest = p

                # ✅ 통합 워크시트에 '기존 저장 방식 그대로' 행/이미지 추가
                _ = append_to_worksheet(ws_all, data_results, images)

                log_q.put(f"[DONE] {s} 완료")
                step_payload = {"shop": s, "path": latest}
                results.append(step_payload)
                data_results.extend(crawler.results)
                images.extend(crawler._images)
                print(len(images))
                window.write_event_value("-STEP_DONE-", step_payload)

            except Exception as e:
                log_q.put("[ERROR] " + repr(e))
                log_q.put(traceback.format_exc())

        # ✅ 모든 상점 처리 후 통합 파일 저장
        try:
            ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            combined_path = os.path.join(outdir, f"qoo10_top5_{ts}.xlsx")
            wb_all.save(combined_path)
            log_q.put(f"[SAVE] 통합 결과 저장: {combined_path}")
        except Exception as e:
            log_q.put(f"[WARN] 통합 파일 저장 실패: {e}")

        window.write_event_value("-ALL_DONE-", True)
    except Exception as e:
        log_q.put("[ERROR] " + repr(e))
        log_q.put(traceback.format_exc())
        window.write_event_value("-ALL_DONE-", True)

# Crawler 클래스 내부에 메서드 추가
def append_to_worksheet(ws, data_results, images, start_row: int = None) -> int:
    """
    현재 self.results / self._images 내용을
    외부 워크시트(ws)에 '기존 저장 형식' 그대로 추가한다.
    - 헤더는 쓰지 않음(외부에서 1회만 작성)
    - 이미지 스케일/행 높이 계산은 save_outputs와 동일
    반환값: 추가한 행 수
    """
    if not data_results:
        return 0

    # 이미지 열/열폭/픽셀폭 계산 (save_outputs와 동일)
    img_col_letter = "I"
    # 외부 ws의 이미지 열 폭이 설정되어 있지 않다면 기본값을 준다
    if ws.column_dimensions[img_col_letter].width in (None, 0):
        ws.column_dimensions[img_col_letter].width = 25
    target_col_px = excel_col_width_to_pixels(ws.column_dimensions[img_col_letter].width)

    # 시작 행
    row_idx = ws.max_row + 1 if start_row is None else start_row

    # 결과 행 추가
    for i, r in enumerate(data_results, start=1):
        ws.cell(row=row_idx, column=1, value=i)                 # Rank (개별 시트에서는 1..N)
        ws.cell(row=row_idx, column=2, value=r.name)            # Name
        ws.cell(row=row_idx, column=3, value=r.price_jpy)       # Price(JPY)
        ws.cell(row=row_idx, column=4, value=r.price_krw)       # Price(KRW)
        ws.cell(row=row_idx, column=5, value=r.review_count)    # Reviews
        ws.cell(row=row_idx, column=6, value=r.product_url)     # Product URL
        ws.cell(row=row_idx, column=7, value=r.shop_name)       # Shop
        ws.cell(row=row_idx, column=8, value=r.total_count)     # Total Count
        # column 9(Image)은 비워둔 뒤 실제 이미지를 add_image

        # 이미지 생성 및 스케일
        img_info = images
        img = XLImage(io.BytesIO(img_info["bytes"]))

        orig_w, orig_h = float(img.width), float(img.height)
        scale = min(1.0, target_col_px / orig_w) if orig_w > 0 else 1.0
        img.width = orig_w * scale
        img.height = orig_h * scale

        # 행 높이 = 이미지 높이에 맞춰 자동 설정
        ws.row_dimensions[row_idx].height = pixels_to_row_height_points(img.height)

        # 이미지 삽입 (I열)
        anchor = f"{img_col_letter}{row_idx}"
        ws.add_image(img, anchor)
        row_idx += 1
    return len(data_results)

def normalize_shop(line: str) -> str:
    line = line.strip()
    if not line:
        return ""
    if "qoo10.jp" in line:
        return line.rstrip("/").split("/")[-1]
    return line

def rows_from_one_file(path: str, shop: str) -> list[list]:
    """
    방금 끝난 엑셀 파일 하나를 읽어 미리보기 테이블용 행을 반환.
    크롤러가 저장한 사람친화 헤더(대문자/공백 포함)와
    내부 스키마(소문자 스네이크케이스) 둘 다 지원한다.
    """
    try:
        if not path:
            return []
        df = pd.read_excel(path)

        # 목표 테이블 헤더 순서
        targets = ["Shop", "Name", "JPY", "KRW", "Reviews", "URL"]

        # 각 타겟 컬럼이 될 수 있는 '후보 이름'들(우선순위 순)
        candidates = {
            "Shop":    ["shop_name", "Shop"],
            "Name":    ["name", "Name"],
            "JPY":     ["price_jpy", "Price(JPY)", "JPY"],
            "KRW":     ["price_krw", "Price(KRW)", "KRW"],
            "Reviews": ["review_count", "Reviews"],
            "URL":     ["product_url", "Product URL", "URL"],
        }

        # 실제 df에 존재하는 컬럼 이름을 타겟별로 매핑
        actual = {}
        cols_set = set(df.columns)
        for tgt in targets:
            found = None
            for cand in candidates[tgt]:
                if cand in cols_set:
                    found = cand
                    break
            actual[tgt] = found

        # Shop이 없으면 파라미터로 받은 shop 값 사용
        if actual["Shop"] is None:
            df["__shop"] = shop
            actual["Shop"] = "__shop"

        # 존재하지 않는 나머지는 빈 문자열로 채움
        for tgt in targets:
            if actual[tgt] is None:
                df[f"__empty_{tgt}"] = ""
                actual[tgt] = f"__empty_{tgt}"

        # 타겟 순서로 DataFrame 구성 후 리스트로 변환
        out = df[[actual[t] for t in targets]].values.tolist()
        return out

    except Exception as e:
        return []