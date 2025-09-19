from crawler import Crawler
from utils import ensure_dir
import threading

class CrawlerManager:
    _instance = None
    _lock = threading.Lock()

    def __init__(self, save_path: str, period: str):
        self.save_path = save_path
        self.period = period
        self._crawler = None  # 단일 Crawler 인스턴스

    @classmethod
    def get(cls, save_path: str, period: str) -> "CrawlerManager":
        with cls._lock:
            if cls._instance is None:
                cls._instance = CrawlerManager(save_path, period)
            else:
                # 최신 설정으로 갱신
                cls._instance.save_path = save_path
                cls._instance.period = period
            return cls._instance

    def run_shop(self, shop_name: str) -> Crawler:
        """
        단일 Crawler 인스턴스를 재사용하되,
        shop_name/period/save_path를 갱신해서 실행합니다.
        """
        if self._crawler is None:
            self._crawler = Crawler(shop_name=shop_name, save_path=self.save_path, period=self.period)
        else:
            # 필드 갱신(크롤러 구현에 맞춰 필드명 조정)
            self._crawler.shop_name = shop_name
            self._crawler.period = self.period
            self._crawler.results = []
            self._crawler.images = []
            self._crawler._snap = []
            # 저장 경로가 바뀔 수 있으니 보장
            ensure_dir(self.save_path)
            # 일부 구현에서는 save_root 같은 속성을 쓰기도 하므로 방어적 갱신
            if hasattr(self._crawler, "save_root"):
                self._crawler.save_root = self.save_path

        self._crawler.run()
        return self._crawler