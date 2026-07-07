from __future__ import annotations

import re
from collections import Counter, defaultdict
from typing import Iterable

from .models import Review


PAIN_PATTERNS: dict[str, tuple[str, ...]] = {
    "Упаковка / протечки / повреждения": (
        r"\bтеч", r"\bпротек", r"\bразлит", r"\bмокр", r"\bупаков",
        r"\bканистр.*помят", r"\bкрышк.*сломан", r"\bповрежд", r"\bгермет",
    ),
    "Доставка": (
        r"\bдостав", r"\bкурьер", r"\bзадерж", r"\bопозда", r"\bдолго ехал",
        r"\bпункт выдач", r"\bпвз\b",
    ),
    "Качество / подлинность": (
        r"\bподдел", r"\bоригинал", r"\bкачест", r"\bбрак", r"\bконтрафакт",
        r"\bсомнен", r"\bсертифик", r"\bпломб",
    ),
    "Эффект / эксплуатационные свойства": (
        r"\bне помог", r"\bэффект.*нет", r"\bрасход", r"\bшум", r"\bстук",
        r"\bдвигател", r"\bперегрев", r"\bзапуск", r"\bдым", r"\bнагар",
    ),
    "Совместимость / характеристики": (
        r"\bне подош", r"\bсовмест", r"\bдопуск", r"\bвязкост", r"\bспецификац",
        r"\bне соответствует", r"\bартикул", r"\bмодель",
    ),
    "Цена": (r"\bдорог", r"\bцен", r"\bпереплат", r"\bдешевле", r"\bскидк"),
    "Запах / цвет / консистенция": (
        r"\bзапах", r"\bвон", r"\bцвет", r"\bгуст", r"\bжидк", r"\bосад", r"\bконсистенц",
    ),
    "Сервис / коммуникация": (
        r"\bпродавец", r"\bне ответ", r"\bподдержк", r"\bобщен", r"\bхам", r"\bвозврат", r"\bобмен",
    ),
    "Инструкция / информация": (
        r"\bинструкц", r"\bописан", r"\bнет информац", r"\bнепонятно", r"\bмаркировк", r"\bэтикет",
    ),
}


def classify_review(review: Review) -> Review:
    text = review.full_text.lower()
    review.pains = [
        label for label, patterns in PAIN_PATTERNS.items()
        if any(re.search(pattern, text, flags=re.IGNORECASE) for pattern in patterns)
    ]
    review.sentiment = "negative" if review.rating <= 2 else "neutral" if review.rating == 3 else "positive"
    return review


def classify_many(reviews: Iterable[Review]) -> list[Review]:
    return [classify_review(r) for r in reviews]


def summary(reviews: list[Review]) -> dict:
    total = len(reviews)
    ratings = Counter(r.rating for r in reviews)
    pains = Counter(p for r in reviews for p in r.pains)
    brands = Counter(r.brand or "Не указан" for r in reviews)
    sellers = Counter(r.seller or "Не указан" for r in reviews)
    platforms = Counter(r.platform for r in reviews)
    negative = sum(1 for r in reviews if r.rating <= 2)
    media = sum(1 for r in reviews if r.has_photo or r.has_video)
    avg = round(sum(r.rating for r in reviews) / total, 2) if total else 0.0
    return {
        "total": total,
        "average_rating": avg,
        "negative_count": negative,
        "negative_share": round(negative / total * 100, 1) if total else 0.0,
        "media_count": media,
        "media_share": round(media / total * 100, 1) if total else 0.0,
        "ratings": ratings,
        "pains": pains,
        "brands": brands,
        "sellers": sellers,
        "platforms": platforms,
    }


def weekly_trend(reviews: list[Review]) -> list[dict]:
    buckets: dict[tuple[str, str, str], list[Review]] = defaultdict(list)
    for r in reviews:
        iso = r.review_date.isocalendar()
        week = f"{iso.year}-W{iso.week:02d}"
        buckets[(week, r.brand or "Не указан", r.seller or "Не указан")].append(r)
    rows = []
    for (week, brand, seller), items in sorted(buckets.items()):
        rows.append({
            "week": week,
            "brand": brand,
            "seller": seller,
            "reviews": len(items),
            "average_rating": round(sum(x.rating for x in items) / len(items), 2),
            "negative_share": round(sum(1 for x in items if x.rating <= 2) / len(items) * 100, 1),
            "with_media": sum(1 for x in items if x.has_photo or x.has_video),
        })
    return rows
