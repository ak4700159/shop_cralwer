from dataclasses import dataclass

@dataclass
class ItemRow:
    shop_name: str
    name: str
    price_jpy: int
    price_krw: float
    review_count: int
    image_url: str
    image_path: str
    product_url: str
    total_count: str