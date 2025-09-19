from dataclasses import dataclass

@dataclass
class Image:
    idx: int
    img_bytes: bytes
    ext: str
    