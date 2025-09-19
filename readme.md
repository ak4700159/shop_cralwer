# Qoo10 간단 크롤러
- 각 상점별 Top5 아이템을 자동으로 추출해 이미지 파일로 저장


# 개발환경
- python==3.11.13
- 가상환경(conda) + pip 사용
    
        conda create -n shop python==3.11.13
        conda activate shop
        pip install -r requirements



# exe 만들기 (PyInstaller)

- 설치 명령어

      pip install -U pyinstaller



- 콘솔 숨기고 단일 파일로 빌드 명령어 (권장) 

       pyinstaller --noconsole --onefile --name Qoo10Crawler app.py

빌드 결과는 dist/Qoo10Crawler.exe에 생성된다 (기본 동작/폴더 구조는 PyInstaller 문서 참고)