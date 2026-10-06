from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path


@dataclass(frozen=True)
class FlowRow:
    trade_date: date
    code: str
    name: str
    pct_change: float | None
    turnover: float | None
    main_net: float | None
    main_ratio: float | None
    super_net: float | None
    big_net: float | None
    mid_net: float | None
    small_net: float | None


@dataclass(frozen=True)
class ParsedWorkbook:
    path: Path
    trade_date: date
    sha256: str
    rows: tuple[FlowRow, ...]
    warnings: tuple[str, ...]


@dataclass(frozen=True)
class ImportResult:
    trade_date: date
    row_count: int
    status: str
    message: str

