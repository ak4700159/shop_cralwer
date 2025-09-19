import queue
import os
import FreeSimpleGUI as sg 
import threading
from app_process import *
from utils import ensure_dir
import webbrowser

# Layout 기본 설정
sg.theme("SystemDefaultForReal")
if hasattr(sg, "set_options"):
    sg.set_options(font=("Segoe UI", 10), dpi_awareness=True)

ICON_PATH = os.path.join(os.path.dirname(__file__), "icon.ico")

""" APP BUILDER : application frame build """
class AppBuilder:
    def __init__(self):
        self.log_q: queue.Queue[str] = queue.Queue()
        self.latest_results: list[str] = []  # [path path .... ], 크롤링 결과 저장 파일 경로
        self.running = False
        self.total_shops = 0
        self.processed = 0
        self.current_period = "W"
        self.last_clicked_cell = None  # 셀 복사용 좌표
        # ✅ 미리보기 누적 버퍼
        self.preview_rows: list[list] = []  # [["Shop","Name","JPY","KRW","Reviews","URL"] 형태의 데이터 누적]

    def log(self, text):
        self.window["-LOG-"].update(disabled=False)
        self.window["-LOG-"].print(text)
        self.window["-LOG-"].update(disabled=True)

    def make_header(self) -> list[list]:
        return [
            [sg.Text("Qoo10 베스트셀러 수집기", font=("Segoe UI", 14, "bold"))],
            [sg.Text("상점 입력 (한 줄에 하나: 'anua' 처럼 shop 이름 또는 전체 m.qoo10 URL)"),
            sg.Push(),
            sg.Button("예시 붙여넣기", key="-EXAMPLE-")],
            [sg.Multiline(size=(80,6), key="-INPUT-", expand_x=True, expand_y=False)]
        ]

    def period_buttons_row(self, selected: str):
        def style(key, label, is_sel):
            return sg.Button(
                label, key=key,
                button_color=("white", "#0078D7") if is_sel else ("#333333", "#E5E5E5"),
                border_width=1, size=(8,1), pad=(2,2)
            )
        return [
            sg.Text("기간 선택", size=(8,1)),
            style("-PERIOD_D-", "일(D)", selected == "D"),
            style("-PERIOD_W-", "주(W)", selected == "W"),
            style("-PERIOD_M-", "월(M)", selected == "M"),
            sg.Text("", key="-PERIOD_LABEL-", size=(16,1), pad=(8,0))
        ]

    def make_path_frame(self):
        return [
            # 이미지 폴더는 최신 Crawler에선 쓰지 않지만, 기존 UI 호환을 위해 남겨둠(무시됨)
            [sg.Text("엑셀 저장 폴더"), sg.Input("./results", key="-OUTDIR-", size=(40,1)), sg.FolderBrowse(target="-OUTDIR-")],
        ]

    def update_period_buttons(self, sel: str):
        # 버튼 색 갱신, sel = 현재 선택된 기간
        def set_btn(key, is_sel):
            self.window[key].update(button_color=("white", "#0078D7") if is_sel else ("#333333", "#E5E5E5"))
        set_btn("-PERIOD_D-", sel == "D")
        set_btn("-PERIOD_W-", sel == "W")
        set_btn("-PERIOD_M-", sel == "M")
        self.window["-PERIOD_LABEL-"].update(f"선택: {sel}")

    def make_period_frame(self):
        default_period = "W"
        return [
            [sg.Frame("기간", [self.period_buttons_row(default_period)], expand_x=True)]
        ]

    def make_control_frame(self):
        return [
            [sg.Button("수집 시작", key="-START-", button_color=("white","#0078D7")),
            sg.Button("중지", key="-STOP-", disabled=True),
            sg.Push(),
            sg.Button("엑셀 열기(최근)", key="-OPENXLS-", disabled=True)],
            [sg.ProgressBar(max_value=100, orientation="h", size=(40,20), key="-PROG-")],
        ]

    def make_result_frame(self):
        return [
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
                    enable_click_events=True,
                    enable_events=True,
                    right_click_menu=["", ["선택 셀 복사", "선택 행 복사", "URL 복사"]],
                )
            ]], expand_x=True)]
        ]

    def make_layout(self):
        header = self.make_header()
        period_frame = self.make_period_frame()
        path_frame = self.make_path_frame()
        control_frame = self.make_control_frame()
        result_frame = self.make_result_frame()
        self.layout = header + period_frame + path_frame + control_frame + result_frame

    def make_window(self):
        self.window = sg.Window(
            "Qoo10 Crawler",
            self.layout,
            resizable=True,
            finalize=True,
            icon=ICON_PATH
        )

    def exec_event_loop(self):
        """
        발생할 수 있는 전체 이벤트 목록
            1. -EXAMPLE- : 상점 예시 추가
            2. -PERIOD_D-, -PERIOD_W-, -PERIOD_M- : 기간 변경
            3. -START- : 수집 시작 버튼 
            4. -STOP- : 중단 버튼
            5. -STEP_DONE- : 상점 하나 크롤링 완료
            6. -ALL_DONE- : 모든 크롤링 완료
            7. -OPENXLS- : 엑셀 파일 열기
            8. 선택 셀 복사, 선택 행 복사, URL 복사
        """
        while True:
            # 100ms 마다 프레임에서 발생한 이벤트를 읽는다 
            event, values = self.window.read(timeout=100)
            if event in (sg.WINDOW_CLOSE_ATTEMPTED_EVENT, sg.WIN_CLOSED, "Exit"):
                break

            # 로그 플러시
            try:
                while True:
                    self.log(self.log_q.get_nowait())
            except queue.Empty:
                pass

            if event == "-EXAMPLE-":
                self.window["-INPUT-"].update("anua\nromand\nzenb\n")

            # 기간 버튼
            if event in ("-PERIOD_D-", "-PERIOD_W-", "-PERIOD_M-"):
                self.current_period = {"-PERIOD_D-":"D", "-PERIOD_W-":"W", "-PERIOD_M-":"M"}[event]
                self.update_period_buttons(self.current_period)
                self.log(f"[INFO] period = {self.current_period}")

            if event == "-START-" and not self.running:
                self.preview_rows.clear()  # ✅ 누적 미리보기 리셋
                # 입력된 상점이름 정규화
                shops = [normalize_shop(s) for s in values["-INPUT-"].splitlines()]
                shops = [s for s in shops if s]
                # 상점을 아무것도 추가하지 않고 수집 시작할 경우
                if not shops:
                    sg.popup_error("상점 이름(또는 URL)을 한 줄에 하나씩 입력하세요.")
                    continue

                outdir = values["-OUTDIR-"] or "./results"
                ensure_dir(outdir)

                # 상태 초기화
                total_shops = len(shops)
                processed = 0
                self.running = True
                self.window["-TABLE-"].update(values=self.preview_rows)
                self.window["-STOP-"].update(disabled=False)
                self.window["-OPENXLS-"].update(disabled=True)
                self.window["-PROG-"].update(0)

                self.log(f"[INFO] 총 {total_shops}개 작업(순차) 시작 / period={self.current_period}")

                # 단일 워커 스레드 시작
                t = threading.Thread(
                    target=run_all_sequential,
                    args=(self.window, shops, outdir, self.current_period, self.log_q),
                    daemon=True
                )
                t.start()

            if event == "-STOP-" and self.running:
                self.log("[INFO] 중지 요청됨 (현재 작업이 끝나면 멈춥니다)")

            # 한 상점 완료 시
            if event == "-STEP_DONE-":
                payload = values["-STEP_DONE-"]  # save path 추출
                self.latest_results.append(payload)
                # 진행 퍼센트 업데이트(로그창)
                processed += 1
                pct = int((processed / max(1, total_shops)) * 100)
                self.window["-PROG-"].update(pct)

                # 미리보기: 이번에 끝난 파일만 읽어서 누적
                # 현재 중간 과정마다 실행된 결과는 확인할 수 없다. 
                # 나중에 처리해보셈
                new_rows = rows_from_one_file(payload)
                if new_rows:
                    self.preview_rows.extend(new_rows)
                    self.window["-TABLE-"].update(values=self.preview_rows)

                # 최근 엑셀 버튼 활성화
                self.window["-OPENXLS-"].update(disabled=(len(self.latest_results) == 0))

            # 전체 완료
            if event == "-ALL_DONE-":
                self.running = False
                self.window["-STOP-"].update(disabled=True)
                self.window["-PROG-"].update(100)

            if event == "-OPENXLS-":
                if self.latest_results:
                    p = self.latest_results[-1]  # 최근 파일
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
                        data = self.preview_rows
                        if last_clicked_cell is None:
                            sg.popup_ok("먼저 복사할 셀을 클릭하세요.")
                        else:
                            r, c = last_clicked_cell
                            if 0 <= r < len(data) and 0 <= c < len(data[r]):
                                sg.clipboard_set(str(data[r][c]))
                                self.log(f"[COPY] 셀({r},{c}) -> 클립보드")
                            else:
                                sg.popup_ok("잘못된 셀 위치입니다.")
                except Exception as e:
                    self.log(f"[ERROR] 셀 복사 실패: {e}")

            # 우클릭 메뉴 - 선택 행 복사
            if event == "선택 행 복사":
                try:
                    selected_rows = values.get("-TABLE-", [])
                    data = self.preview_rows
                    selected_rows = values.get("-TABLE-", [])
                    if not selected_rows:
                        if last_clicked_cell is None:
                            sg.popup_ok("먼저 복사할 행을 선택(또는 셀을 클릭)하세요.")
                            continue
                        selected_rows = [last_clicked_cell[0]]

                    lines = []
                    for r in selected_rows:
                        if 0 <= r < len(data):
                            lines.append("\t".join(str(x) for x in data[r]))

                    if lines:
                        sg.clipboard_set("\n".join(lines))
                        self.log(f"[COPY] {len(selected_rows)}개 행 -> 클립보드")
                    else:
                        sg.popup_ok("복사할 행이 없습니다.")
                except Exception as e:
                    self.log(f"[ERROR] 행 복사 실패: {e}")

            # 우클릭 메뉴 - URL 복사
            if event == "URL 복사":
                try:
                    data = self.preview_rows
                    url_col_idx = 5  # ["Shop","Name","JPY","KRW","Reviews","URL"]
                    targets = []

                    selected_rows = values.get("-TABLE-", [])
                    if selected_rows:
                        for r in selected_rows:
                            if 0 <= r < len(data) and url_col_idx < len(data[r]):
                                targets.append(str(data[r][url_col_idx]))
                    else:
                        if last_clicked_cell is None:
                            sg.popup_ok("URL을 복사할 행을 선택(또는 URL 셀을 클릭)하세요.")
                            continue
                        r, _ = last_clicked_cell
                        if 0 <= r < len(data) and url_col_idx < len(data[r]):
                            targets.append(str(data[r][url_col_idx]))

                    if targets:
                        sg.clipboard_set("\n".join(targets))
                        self.log(f"[COPY] URL {len(targets)}개 -> 클립보드")
                    else:
                        sg.popup_ok("복사할 URL이 없습니다.")
                except Exception as e:
                    self.log(f"[ERROR] URL 복사 실패: {e}")


    def exit_app(self):
        self.window.close()

    def make_app(self):
        self.make_layout()
        self.make_window()

    def exec_app(self):
        self.exec_event_loop()
        self.exit_app()