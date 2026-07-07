from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Iterable

from .models import Review
from .utils import json_dumps, parse_datetime


SCHEMA = """
CREATE TABLE IF NOT EXISTS reviews (
    platform TEXT NOT NULL,
    review_id TEXT NOT NULL,
    review_date TEXT NOT NULL,
    rating INTEGER NOT NULL,
    brand TEXT NOT NULL DEFAULT '',
    seller TEXT NOT NULL DEFAULT '',
    product_name TEXT NOT NULL DEFAULT '',
    product_id TEXT NOT NULL DEFAULT '',
    sku TEXT NOT NULL DEFAULT '',
    text TEXT NOT NULL DEFAULT '',
    pros TEXT NOT NULL DEFAULT '',
    cons TEXT NOT NULL DEFAULT '',
    photo_urls TEXT NOT NULL DEFAULT '[]',
    video_urls TEXT NOT NULL DEFAULT '[]',
    source_url TEXT NOT NULL DEFAULT '',
    source_id TEXT NOT NULL DEFAULT '',
    pains TEXT NOT NULL DEFAULT '[]',
    sentiment TEXT NOT NULL DEFAULT '',
    raw TEXT NOT NULL DEFAULT '{}',
    collected_at TEXT NOT NULL,
    PRIMARY KEY (platform, review_id)
);
CREATE INDEX IF NOT EXISTS idx_reviews_date ON reviews(review_date DESC);
CREATE INDEX IF NOT EXISTS idx_reviews_brand_seller ON reviews(brand, seller);
CREATE INDEX IF NOT EXISTS idx_reviews_rating ON reviews(rating);
"""


class ReviewStore:
    def __init__(self, path: str):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.path)
        self.conn.executescript(SCHEMA)

    def close(self) -> None:
        self.conn.close()

    def upsert_many(self, reviews: Iterable[Review]) -> int:
        now = datetime.utcnow().isoformat(timespec="seconds") + "Z"
        rows = []
        for r in reviews:
            rows.append((
                r.platform, r.review_id, r.review_date.isoformat(), int(r.rating),
                r.brand, r.seller, r.product_name, r.product_id, r.sku,
                r.text, r.pros, r.cons, json_dumps(r.photo_urls),
                json_dumps(r.video_urls), r.source_url, r.source_id,
                json_dumps(r.pains), r.sentiment, json_dumps(r.raw), now,
            ))
        if not rows:
            return 0
        self.conn.executemany(
            """
            INSERT INTO reviews (
                platform, review_id, review_date, rating, brand, seller,
                product_name, product_id, sku, text, pros, cons,
                photo_urls, video_urls, source_url, source_id, pains,
                sentiment, raw, collected_at
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(platform, review_id) DO UPDATE SET
                review_date=excluded.review_date,
                rating=excluded.rating,
                brand=excluded.brand,
                seller=excluded.seller,
                product_name=excluded.product_name,
                product_id=excluded.product_id,
                sku=excluded.sku,
                text=excluded.text,
                pros=excluded.pros,
                cons=excluded.cons,
                photo_urls=excluded.photo_urls,
                video_urls=excluded.video_urls,
                source_url=excluded.source_url,
                source_id=excluded.source_id,
                pains=excluded.pains,
                sentiment=excluded.sentiment,
                raw=excluded.raw,
                collected_at=excluded.collected_at
            """,
            rows,
        )
        self.conn.commit()
        return len(rows)

    def all(self, limit: int | None = None) -> list[Review]:
        sql = "SELECT * FROM reviews ORDER BY review_date DESC"
        params: tuple = ()
        if limit:
            sql += " LIMIT ?"
            params = (limit,)
        cursor = self.conn.execute(sql, params)
        names = [x[0] for x in cursor.description]
        return [self._row_to_review(dict(zip(names, row))) for row in cursor.fetchall()]

    def between(self, date_from: datetime, date_to: datetime, limit: int | None = None) -> list[Review]:
        sql = (
            "SELECT * FROM reviews WHERE datetime(review_date) >= datetime(?) "
            "AND datetime(review_date) <= datetime(?) ORDER BY datetime(review_date) DESC"
        )
        params: list = [date_from.isoformat(), date_to.isoformat()]
        if limit:
            sql += " LIMIT ?"
            params.append(limit)
        cursor = self.conn.execute(sql, tuple(params))
        names = [x[0] for x in cursor.description]
        return [self._row_to_review(dict(zip(names, row))) for row in cursor.fetchall()]

    @staticmethod
    def _row_to_review(row: dict) -> Review:
        return Review(
            platform=row["platform"],
            review_id=row["review_id"],
            review_date=parse_datetime(row["review_date"]),
            rating=int(row["rating"]),
            brand=row["brand"],
            seller=row["seller"],
            product_name=row["product_name"],
            product_id=row["product_id"],
            sku=row["sku"],
            text=row["text"],
            pros=row["pros"],
            cons=row["cons"],
            photo_urls=json.loads(row["photo_urls"] or "[]"),
            video_urls=json.loads(row["video_urls"] or "[]"),
            source_url=row["source_url"],
            source_id=row["source_id"],
            pains=json.loads(row["pains"] or "[]"),
            sentiment=row["sentiment"],
            raw=json.loads(row["raw"] or "{}"),
        )
