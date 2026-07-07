from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any


@dataclass
class Review:
    platform: str
    review_id: str
    review_date: datetime
    rating: int
    brand: str = ""
    seller: str = ""
    product_name: str = ""
    product_id: str = ""
    sku: str = ""
    text: str = ""
    pros: str = ""
    cons: str = ""
    photo_urls: list[str] = field(default_factory=list)
    video_urls: list[str] = field(default_factory=list)
    source_url: str = ""
    source_id: str = ""
    pains: list[str] = field(default_factory=list)
    sentiment: str = ""
    raw: dict[str, Any] = field(default_factory=dict)

    @property
    def has_photo(self) -> bool:
        return bool(self.photo_urls)

    @property
    def has_video(self) -> bool:
        return bool(self.video_urls)

    @property
    def full_text(self) -> str:
        return " ".join(x.strip() for x in (self.text, self.pros, self.cons) if x and x.strip())
