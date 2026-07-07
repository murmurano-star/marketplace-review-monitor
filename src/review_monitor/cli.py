from __future__ import annotations

import argparse
import sys
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

from dateutil import parser as date_parser

from .analysis import classify_many
from .collectors import ImportFileCollector, PublicMarketplaceBrowserCollector
from .config import load_config
from .report import write_dashboard, write_excel, write_markdown
from .storage import ReviewStore

COLLECTORS = {"public_marketplace_browser": PublicMarketplaceBrowserCollector}


def _parse_boundary(value: str, *, end_of_day: bool, timezone_name: str) -> datetime:
    parsed = date_parser.parse(value)
    tz = ZoneInfo(timezone_name)
    if len(value.strip()) <= 10 and isinstance(parsed.date(), date):
        parsed = datetime.combine(parsed.date(), time.max if end_of_day else time.min, tzinfo=tz)
    elif parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=tz)
    return parsed.astimezone(timezone.utc)


def resolve_date_range(config: dict, date_from_arg: str | None = None, date_to_arg: str | None = None) -> tuple[datetime, datetime]:
    range_config = config.get("date_range") or {}
    timezone_name = str(config.get("timezone", "Europe/Moscow"))
    date_from_value = (date_from_arg or range_config.get("date_from") or "").strip()
    date_to_value = (date_to_arg or range_config.get("date_to") or "").strip()
    date_to = _parse_boundary(date_to_value, end_of_day=True, timezone_name=timezone_name) if date_to_value else datetime.now(timezone.utc)
    date_from = _parse_boundary(date_from_value, end_of_day=False, timezone_name=timezone_name) if date_from_value else date_to - timedelta(days=int(range_config.get("default_lookback_days", 7)))
    if date_from > date_to:
        raise ValueError("date_from не может быть позже date_to")
    return date_from, date_to


def run(config_path: str, date_from_arg: str | None = None, date_to_arg: str | None = None) -> int:
    config = load_config(config_path)
    date_from, date_to = resolve_date_range(config, date_from_arg, date_to_arg)
    print(f"[range] {date_from.isoformat()} — {date_to.isoformat()}")
    collected = []

    for source in config.get("sources", []):
        if not source or not source.get("enabled", True):
            continue
        collector_name = source.get("collector")
        collector_class = COLLECTORS.get(collector_name)
        if not collector_class:
            raise ValueError(f"Неизвестный collector: {collector_name}")
        source = dict(source)
        source.setdefault("timezone", config.get("timezone", "Europe/Moscow"))
        print(f"[collect] {source.get('id', collector_name)}")
        reviews = collector_class(source).collect(date_from, date_to)
        print(f"[collect] получено: {len(reviews)}")
        collected.extend(reviews)

    for item in config.get("import_files", []) or []:
        reviews = ImportFileCollector(item).collect(date_from, date_to)
        collected.extend(reviews)

    classified = classify_many(collected)
    store = ReviewStore(config["database"])
    try:
        written = store.upsert_many(classified)
        range_reviews = store.between(date_from, date_to)
        database_total = len(store.all())
    finally:
        store.close()

    output = Path(config["output_dir"])
    output.mkdir(parents=True, exist_ok=True)
    write_excel(range_reviews, output / "latest.xlsx", date_from=date_from, date_to=date_to)
    write_markdown(range_reviews, output / "latest.md", date_from=date_from, date_to=date_to)
    write_dashboard(range_reviews, output / "dashboard.html", max_rows=int(config.get("max_dashboard_rows", 5000)))
    print(f"[done] обработано: {written}; в диапазоне: {len(range_reviews)}; в базе: {database_total}")
    return 0


def main() -> None:
    parser = argparse.ArgumentParser(description="Мониторинг отзывов Wildberries и Ozon")
    sub = parser.add_subparsers(dest="command", required=True)
    run_parser = sub.add_parser("run")
    run_parser.add_argument("--config", default="config/targets.yml")
    run_parser.add_argument("--date-from", default=None)
    run_parser.add_argument("--date-to", default=None)
    args = parser.parse_args()
    try:
        if args.command == "run":
            raise SystemExit(run(args.config, args.date_from, args.date_to))
    except Exception as exc:
        print(f"[error] {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
