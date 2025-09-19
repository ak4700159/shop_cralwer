# 개발환경
- python==3.11.13
- 가상환경(conda) + pip 사용
    
        conda create -n shop python==3.11.13
        conda activate shop
        pip install -r requirements

- 


# exe 만들기 (PyInstaller)

설치

    pip install -U pyinstaller


프로젝트 루트에서 빌드

- 콘솔 숨기고 단일 파일로 빌드 (권장)

       pyinstaller --noconsole --onefile --name Qoo10Crawler app.py

빌드 결과는 dist/Qoo10Crawler.exe에 생성된다 (기본 동작/폴더 구조는 PyInstaller 문서 참고)