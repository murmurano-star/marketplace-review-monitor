from __future__ import annotations

import json
from datetime import datetime, timezone
from typing import Any

from dateutil import parser as date_parser


def parse_datetime(value: Any) -> datetime:
    if isinstance(value, datetime):
        dt = value
    elif isinstance(value, (int, float)):
        dt = datetime.fromtimestamp(value, tz=timezone.utc)
    elif value:
        dt = date_parser.parse(str(value))
    else:
        dt = datetime.now(timezone.utc)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


def json_dumps(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def list_from_value(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(x) for x in value if x]
    if isinstance(value, tuple):
        return [str(x) for x in value if x]
    if isinstance(value, str):
        return [x.strip() for x in value.split("|") if x.strip()]
    return [str(value)]
