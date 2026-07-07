from __future__ import annotations

import hashlib
import re
import time
from datetime import datetime, timezone
from typing import Any
from urllib.parse import urljoin

from .base import Collector
from ..models import Review

MONTHS_RU = {
    "января": 1, "февраля": 2, "марта": 3, "апреля": 4,
    "мая": 5, "июня": 6, "июля": 7, "августа": 8,
    "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12,
}


class PublicMarketplaceBrowserCollector(Collector):
    """Сбор публичных отзывов обычным Chromium без обхода защит."""

    def __init__(self, config: dict[str, Any]):
        self.config = config
        self.platform = str(config["platform"]).strip().lower()
        if self.platform not in {"ozon", "wildberries"}:
            raise ValueError("platform должен быть ozon или wildberries")
        self.seller_url = str(config["seller_url"]).strip()
        self.seller_name = str(config.get("seller_name") or self.seller_url)
        self.brands = [str(x).upper() for x in config.get("brands", []) if x]
        self.source_id = str(config.get("id", f"{self.platform}-public"))
        self.headless = bool(config.get("headless", True))
        self.max_products = int(config.get("max_products_per_seller", 100))
        self.max_reviews = int(config.get("max_reviews_per_product", 500))
        self.scroll_rounds = int(config.get("scroll_rounds", 25))
        self.review_scroll_rounds = int(config.get("review_scroll_rounds", 20))
        self.timeout = int(config.get("page_timeout_ms", 60000))
        self.delay = int(config.get("delay_ms", 1500)) / 1000

    def collect(self, date_from: datetime, date_to: datetime) -> list[Review]:
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as exc:
            raise RuntimeError("Установите Playwright: pip install playwright") from exc

        reviews: dict[str, Review] = {}
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=self.headless)
            context = browser.new_context(
                locale="ru-RU",
                timezone_id=str(self.config.get("timezone", "Europe/Moscow")),
                viewport={"width": 1440, "height": 1000},
            )
            page = context.new_page()
            page.set_default_timeout(self.timeout)
            page.goto(self.seller_url, wait_until="domcontentloaded")
            self._assert_not_blocked(page)
            products = self._collect_products(page)
            print(f"[public] найдено товаров брендов: {len(products)}")

            for index, product in enumerate(products, start=1):
                print(f"[public] {index}/{len(products)} {product['name']}")
                product_page = context.new_page()
                try:
                    product_page.set_default_timeout(self.timeout)
                    product_page.goto(product["url"], wait_until="domcontentloaded")
                    self._assert_not_blocked(product_page)
                    self._open_reviews(product_page)
                    for review in self._collect_review_cards(product_page, product):
                        if date_from <= review.review_date <= date_to:
                            reviews[review.review_id] = review
                finally:
                    product_page.close()
                time.sleep(self.delay)
            browser.close()
        return sorted(reviews.values(), key=lambda x: x.review_date, reverse=True)

    def _collect_products(self, page: Any) -> list[dict[str, str]]:
        seen: dict[str, dict[str, str]] = {}
        for _ in range(self.scroll_rounds):
            for item in page.locator("a[href]").evaluate_all(
                "els => els.map(a => ({href:a.href, text:(a.innerText || a.closest('article,div')?.innerText || '').trim()}))"
            ):
                href = str(item.get("href") or "")
                text = str(item.get("text") or "")
                if not self._is_product_url(href):
                    continue
                brand = self._detect_brand(text)
                if not brand:
                    continue
                url = self._canonical_product_url(href)
                product_id = self._product_id(url)
                seen[url] = {
                    "url": url,
                    "product_id": product_id,
                    "name": self._clean_product_name(text, brand, product_id),
                    "brand": brand,
                }
                if len(seen) >= self.max_products:
                    return list(seen.values())
            page.mouse.wheel(0, 1600)
            time.sleep(self.delay)
        return list(seen.values())[: self.max_products]

    def _open_reviews(self, page: Any) -> None:
        candidates = [
            "text=Отзывы", "a:has-text('Отзывы')", "button:has-text('Отзывы')",
            "[href*='review']", "[href*='feedback']",
        ]
        for selector in candidates:
            try:
                locator = page.locator(selector).first
                if locator.count() and locator.is_visible():
                    locator.click()
                    page.wait_for_timeout(1200)
                    return
            except Exception:
                continue

    def _collect_review_cards(self, page: Any, product: dict[str, str]) -> list[Review]:
        cards: dict[str, Review] = {}
        selectors = [
            "[data-widget*='review'] article",
            "[data-widget*='review'] > div",
            "[class*='review'] article",
            "[class*='feedback']",
            "article",
        ]
        for _ in range(self.review_scroll_rounds):
            for selector in selectors:
                try:
                    locator = page.locator(selector)
                    count = min(locator.count(), self.max_reviews)
                    for i in range(count):
                        review = self._card_to_review(locator.nth(i), product)
                        if review:
                            cards[review.review_id] = review
                except Exception:
                    continue
            if len(cards) >= self.max_reviews:
                break
            page.mouse.wheel(0, 1800)
            time.sleep(self.delay)
        return list(cards.values())[: self.max_reviews]

    def _card_to_review(self, card: Any, product: dict[str, str]) -> Review | None:
        try:
            text = (card.inner_text() or "").strip()
        except Exception:
            return None
        if len(text) < 10:
            return None
        rating = self._extract_rating(card, text)
        review_date = self._market_date(text)
        if not rating or not review_date:
            return None
        photos, videos = self._media_urls(card)
        digest = hashlib.sha1(
            f"{self.platform}|{product['product_id']}|{review_date.isoformat()}|{rating}|{text}".encode("utf-8")
        ).hexdigest()
        return Review(
            platform=self.platform,
            review_id=f"public-{self.platform}-{digest}",
            review_date=review_date,
            rating=rating,
            brand=product["brand"],
            seller=self.seller_name,
            product_name=product["name"],
            product_id=product["product_id"],
            text=text,
            photo_urls=photos,
            video_urls=videos,
            source_url=product["url"],
            source_id=self.source_id,
        )

    def _extract_rating(self, card: Any, text: str) -> int:
        try:
            labels = card.locator("[aria-label]").evaluate_all("els => els.map(x => x.getAttribute('aria-label'))")
            for label in labels:
                match = re.search(r"([1-5])(?:\s*из\s*5|\s*зв)", str(label), re.I)
                if match:
                    return int(match.group(1))
            stars = card.locator("svg, [class*='star']")
            filled = 0
            for i in range(min(stars.count(), 10)):
                cls = str(stars.nth(i).get_attribute("class") or "").lower()
                if "active" in cls or "filled" in cls:
                    filled += 1
            if 1 <= filled <= 5:
                return filled
        except Exception:
            pass
        match = re.search(r"(?:оценка|рейтинг)?\s*([1-5])\s*(?:из\s*5|звезд|★)", text, re.I)
        return int(match.group(1)) if match else 0

    def _market_date(self, text: str) -> datetime | None:
        match = re.search(r"\b(\d{1,2})\s+(" + "|".join(MONTHS_RU) + r")\s+(\d{4})\b", text.lower())
        if match:
            return datetime(int(match.group(3)), MONTHS_RU[match.group(2)], int(match.group(1)), tzinfo=timezone.utc)
        match = re.search(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{4})\b", text)
        if match:
            return datetime(int(match.group(3)), int(match.group(2)), int(match.group(1)), tzinfo=timezone.utc)
        return None

    def _media_urls(self, card: Any) -> tuple[list[str], list[str]]:
        photos: list[str] = []
        videos: list[str] = []
        try:
            photos = [str(x) for x in card.locator("img").evaluate_all("els => els.map(x => x.currentSrc || x.src).filter(Boolean)")]
            videos = [str(x) for x in card.locator("video, video source").evaluate_all("els => els.map(x => x.currentSrc || x.src).filter(Boolean)")]
        except Exception:
            pass
        return list(dict.fromkeys(photos)), list(dict.fromkeys(videos))

    def _detect_brand(self, text: str) -> str:
        upper = text.upper()
        for brand in self.brands:
            if re.search(rf"(?<![A-ZА-Я0-9]){re.escape(brand)}(?![A-ZА-Я0-9])", upper):
                return brand
        return ""

    def _is_product_url(self, url: str) -> bool:
        return (self.platform == "ozon" and "/product/" in url) or (self.platform == "wildberries" and "/catalog/" in url)

    def _canonical_product_url(self, url: str) -> str:
        absolute = urljoin(self.seller_url, url)
        return absolute.split("?")[0].split("#")[0]

    def _product_id(self, url: str) -> str:
        patterns = [r"-(\d+)/?$", r"/catalog/(\d+)/", r"/(\d+)/detail"]
        for pattern in patterns:
            match = re.search(pattern, url)
            if match:
                return match.group(1)
        return hashlib.sha1(url.encode("utf-8")).hexdigest()[:16]

    @staticmethod
    def _clean_product_name(text: str, brand: str, product_id: str) -> str:
        lines = [x.strip() for x in text.splitlines() if x.strip()]
        for line in lines:
            if brand.lower() in line.lower() and 4 <= len(line) <= 200:
                return line
        return f"{brand} {product_id}"

    @staticmethod
    def _assert_not_blocked(page: Any) -> None:
        body = (page.locator("body").inner_text() or "").lower()
        markers = ["captcha", "проверка, что вы не робот", "доступ ограничен", "access denied"]
        if any(marker in body for marker in markers):
            raise RuntimeError("Маркетплейс ограничил публичный браузерный доступ; обход защиты не выполняется")
