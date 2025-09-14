# app.py  (FreeSimpleGUI + 단일 워커로 순차 처리)
import os
import threading
import queue
import traceback
import webbrowser
import FreeSimpleGUI as sg   # pip install FreeSimpleGUI
from datetime import datetime
import pandas as pd

from crawler import Crawler
from utils import ensure_dir

# -----------------------------
# 단일 워커: 여러 상점을 순차 처리
# -----------------------------
def run_all_sequential(window: sg.Window, shops: list[str], imgdir: str, outdir: str, log_q: queue.Queue):
    """
    하나의 스레드에서 shops를 순차 처리.
    각 상점이 끝날 때마다 -STEP_DONE- 이벤트를 GUI로 보냄.
    모두 끝나면 -ALL_DONE- 이벤트 전송.
    """
    try:
        results = []  # [{"shop":..., "path":...}]
        for s in shops:
            try:
                log_q.put(f"[START] {s} 수집 시작")
                crawler = Crawler(shop_name=s, save_img_path=imgdir, save_path=outdir)
                crawler.run()  # 내부에서 저장까지 수행

                # 가장 최근 엑셀 찾기
                latest = None
                for f in os.listdir(outdir):
                    if f.startswith(f"qoo10_top_{s}_") and f.endswith(".xlsx"):
                        p = os.path.join(outdir, f)
                        if latest is None or os.path.getmtime(p) > os.path.getmtime(latest):
                            latest = p

                log_q.put(f"[DONE] {s} 완료")
                step_payload = {"shop": s, "path": latest}
                results.append(step_payload)

                # GUI에 "한 단계 완료" 통지
                # (메인 이벤트 루프에서 window.read() 시 수신)
                window.write_event_value("-STEP_DONE-", step_payload)  # 권장 패턴 :contentReference[oaicite:2]{index=2}

            except Exception as e:
                log_q.put("[ERROR] " + repr(e))
                log_q.put(traceback.format_exc())

        # 전체 완료 통지
        window.write_event_value("-ALL_DONE-", True)
    except Exception as e:
        log_q.put("[ERROR] " + repr(e))
        log_q.put(traceback.format_exc())
        window.write_event_value("-ALL_DONE-", True)


# -----------------------------
# UI 레이아웃
# -----------------------------
sg.theme("SystemDefaultForReal")
if hasattr(sg, "set_options"):
    sg.set_options(font=("Segoe UI", 10), dpi_awareness=True)

header = [
    [sg.Text("Qoo10 베스트셀러 수집기", font=("Segoe UI", 14, "bold"))],
    [sg.Text("상점 입력 (한 줄에 하나: 'anua' 처럼 shop 이름 또는 전체 m.qoo10 URL)"),
     sg.Push(),
     sg.Button("예시 붙여넣기", key="-EXAMPLE-")],
    [sg.Multiline(size=(80,6), key="-INPUT-", expand_x=True, expand_y=False)]
]

paths = [
    [sg.Text("이미지 저장 폴더"), sg.Input("./imgs", key="-IMGDIR-", size=(40,1)), sg.FolderBrowse(target="-IMGDIR-")],
    [sg.Text("엑셀 저장 폴더"), sg.Input("./results", key="-OUTDIR-", size=(40,1)), sg.FolderBrowse(target="-OUTDIR-")],
]

controls = [
    [sg.Button("수집 시작", key="-START-", button_color=("white","#0078D7")),
     sg.Button("중지", key="-STOP-", disabled=True),
     sg.Push(),
     sg.Button("엑셀 열기(최근)", key="-OPENXLS-", disabled=True)],
    [sg.ProgressBar(max_value=100, orientation="h", size=(40,20), key="-PROG-")],
]

body = [
    [sg.Frame("실행 로그", [[sg.Multiline(size=(80,12), key="-LOG-", autoscroll=True, disabled=True, expand_x=True)]], expand_x=True)],
    [sg.Frame("결과 미리보기 (최근 실행 전체 합본)", [[
        sg.Table(
            values=[],
            headings=["Shop","Name","JPY","KRW","Reviews","URL"],
            key="-TABLE-",
            auto_size_columns=False,
            col_widths=[12,28,8,8,8,40],
            expand_x=True,
            expand_y=True,
            justification="left",
            enable_click_events=True,        # 셀 좌표 이벤트
            enable_events=True,              # 선택 이벤트
            right_click_menu=[
                "", 
                ["선택 셀 복사", "선택 행 복사", "URL 복사"]
            ],
        )
    ]], expand_x=True)]
]
layout = header + paths + controls + body
# 아이콘 파일 지정 (프로젝트 폴더 안에 icon.ico 두었다고 가정)
ICON_PATH = os.path.join(os.path.dirname(__file__), "icon.ico")
window = sg.Window(
    "Qoo10 Crawler",
    layout,
    resizable=True,
    finalize=True,
    icon=ICON_PATH   # ← 여기 추가
)

# -----------------------------
# 상태
# -----------------------------
log_q: queue.Queue[str] = queue.Queue()
latest_results: list[dict] = []  # [{"shop":..., "path":...}]
running = False
total_shops = 0
processed = 0

def log(text):
    window["-LOG-"].update(disabled=False)
    window["-LOG-"].print(text)
    window["-LOG-"].update(disabled=True)

def normalize_shop(line: str) -> str:
    line = line.strip()
    if not line:
        return ""
    if "qoo10.jp" in line:
        return line.rstrip("/").split("/")[-1]
    return line

def table_values_from_files(rows: list[dict]) -> list[list]:
    frames = []
    for r in rows:
        try:
            if not r.get("path"):
                continue
            df = pd.read_excel(r["path"])
            if "shop_name" not in df.columns:
                df["shop_name"] = r["shop"]
            frames.append(df)
        except Exception as e:
            log(f"[WARN] 미리보기 로드 실패({r.get('shop','?')}): {e}")

    if not frames:
        return []
    df_all = pd.concat(frames, ignore_index=True)  # 여러 파일 합치기(권장) :contentReference[oaicite:3]{index=3}
    cols = ["shop_name","name","price_jpy","price_krw","review_count","product_url"]
    cols = [c for c in cols if c in df_all.columns]
    return df_all[cols].values.tolist()

# -----------------------------
# 이벤트 루프
# -----------------------------
while True:
    event, values = window.read(timeout=100)
    if event in (sg.WINDOW_CLOSE_ATTEMPTED_EVENT, sg.WIN_CLOSED, "Exit"):
        break

    if event == "-EXAMPLE-":
        window["-INPUT-"].update("anua\nromand\nzenb\n")

    if event == "-START-" and not running:
        shops = [normalize_shop(s) for s in values["-INPUT-"].splitlines()]
        shops = [s for s in shops if s]
        if not shops:
            sg.popup_error("상점 이름(또는 URL)을 한 줄에 하나씩 입력하세요.")
            continue

        imgdir = values["-IMGDIR-"] or "./imgs"
        outdir = values["-OUTDIR-"] or "./results"
        ensure_dir(imgdir); ensure_dir(outdir)

        # 상태 초기화
        latest_results.clear()
        total_shops = len(shops)
        processed = 0
        running = True
        window["-STOP-"].update(disabled=False)
        window["-OPENXLS-"].update(disabled=True)
        window["-PROG-"].update(0)
        log(f"[INFO] 총 {total_shops}개 작업(순차) 시작")

        # ★ 단일 워커 스레드 시작 (여러 스레드 아님: 순차 처리)
        t = threading.Thread(
            target=run_all_sequential,
            args=(window, shops, imgdir, outdir, log_q),
            daemon=True
        )
        t.start()

    if event == "-STOP-" and running:
        # 간단 안내(현재 구조상 즉시 중지는 어려움)
        log("[INFO] 중지 요청됨 (현재 작업이 끝나면 멈춥니다)")

    # 로그 플러시
    try:
        while True:
            log(log_q.get_nowait())
    except queue.Empty:
        pass

    # 한 상점 완료 시
    if event == "-STEP_DONE-":
        payload = values["-STEP_DONE-"]  # {"shop":..., "path":...}
        latest_results.append(payload)
        processed += 1
        # 진행도 업데이트 (update 호출은 표준 방식) :contentReference[oaicite:4]{index=4}
        pct = int((processed / max(1, total_shops)) * 100)
        window["-PROG-"].update(pct)

        # 미리보기 즉시 반영
        table_vals = table_values_from_files(latest_results)
        window["-TABLE-"].update(values=table_vals)  # Table.update로 갱신 :contentReference[oaicite:5]{index=5}

        # 최근 엑셀 버튼 활성화
        window["-OPENXLS-"].update(disabled=(len(latest_results) == 0))

    # 전체 완료
    if event == "-ALL_DONE-":
        running = False
        window["-STOP-"].update(disabled=True)
        # 100% 보정
        window["-PROG-"].update(100)

    if event == "-OPENXLS-":
        if latest_results:
            p = latest_results[-1]["path"]  # 최근 파일
            try:
                webbrowser.open(p)
            except Exception:
                sg.popup_ok(f"엑셀 위치: {p}")

    # 테이블 셀 클릭 좌표 저장
    if isinstance(event, tuple) and len(event) >= 3 and event[0] == "-TABLE-" and event[1] == "+CLICKED+":
        _, _, (r, c) = event
        if r is not None and c is not None and r >= 0 and c >= 0:
            last_clicked_cell = (r, c)

    # 우클릭 메뉴 - 선택 셀 복사
    if event == "선택 셀 복사":
        try:
            if last_clicked_cell is None:
                sg.popup_ok("먼저 복사할 셀을 클릭하세요.")
            else:
                r, c = last_clicked_cell
                table_vals = window["-TABLE-"].get()
                text = str(table_vals[r][c]) if 0 <= r < len(table_vals) and 0 <= c < len(table_vals[r]) else ""
                sg.clipboard_set(text)
                log(f"[COPY] 셀({r},{c}) -> 클립보드")
        except Exception as e:
            log(f"[ERROR] 셀 복사 실패: {e}")

    # 우클릭 메뉴 - 선택 행 복사 (탭으로 구분)
    if event == "선택 행 복사":
        try:
            selected_rows = values.get("-TABLE-", [])
            table_vals = window["-TABLE-"].get()
            if not selected_rows:
                # 선택이 없으면 마지막 클릭 행으로 대체
                if last_clicked_cell is None:
                    sg.popup_ok("먼저 복사할 행을 선택(또는 셀을 클릭)하세요.")
                    continue
                selected_rows = [last_clicked_cell[0]]
            # 여러 행도 한 번에
            lines = []
            for r in selected_rows:
                if 0 <= r < len(table_vals):
                    lines.append("\t".join(str(x) for x in table_vals[r]))
            sg.clipboard_set("\n".join(lines))
            log(f"[COPY] {len(selected_rows)}개 행 -> 클립보드")
        except Exception as e:
            log(f"[ERROR] 행 복사 실패: {e}")

    # 우클릭 메뉴 - URL 복사 (URL 컬럼만)
    if event == "URL 복사":
        try:
            table_vals = window["-TABLE-"].get()
            url_col_idx = 5  # ["Shop","Name","JPY","KRW","Reviews","URL"] 기준
            targets = []
            selected_rows = values.get("-TABLE-", [])
            if selected_rows:
                for r in selected_rows:
                    if 0 <= r < len(table_vals):
                        targets.append(str(table_vals[r][url_col_idx]))
            else:
                # 선택이 없으면 마지막 클릭한 셀의 행을 사용
                if last_clicked_cell is None:
                    sg.popup_ok("URL을 복사할 행을 선택(또는 URL 셀을 클릭)하세요.")
                    continue
                r, _ = last_clicked_cell
                if 0 <= r < len(table_vals):
                    targets.append(str(table_vals[r][url_col_idx]))

            if not targets:
                sg.popup_ok("복사할 URL이 없습니다.")
            else:
                sg.clipboard_set("\n".join(targets))
                log(f"[COPY] URL {len(targets)}개 -> 클립보드")
        except Exception as e:
            log(f"[ERROR] URL 복사 실패: {e}")

window.close()
