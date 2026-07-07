from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml


def load_config(path: str | Path) -> dict[str, Any]:
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"Не найден файл конфигурации: {config_path}")

    data = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    data.setdefault("timezone", "Europe/Moscow")
    data.setdefault("database", "data/reviews.sqlite3")
    data.setdefault("output_dir", "reports")
    data.setdefault("max_dashboard_rows", 5000)
    data.setdefault("sources", [])
    data.setdefault("import_files", [])
    return data
