from __future__ import annotations

import os
from pathlib import Path


def app_data_dir() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    path = base / "DailyFlowAnalyzer"
    path.mkdir(parents=True, exist_ok=True)
    return path


def imports_dir() -> Path:
    path = app_data_dir() / "imports"
    path.mkdir(parents=True, exist_ok=True)
    return path


def cache_dir() -> Path:
    path = app_data_dir() / "cache"
    path.mkdir(parents=True, exist_ok=True)
    return path


def database_path() -> Path:
    return app_data_dir() / "daily_flow.sqlite3"


def token_path() -> Path:
    return app_data_dir() / "tushare.token"

