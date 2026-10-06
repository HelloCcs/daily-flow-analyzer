from __future__ import annotations

import shutil
import sqlite3
from datetime import date, datetime
from pathlib import Path

from .models import ImportResult, ParsedWorkbook
from .paths import database_path, imports_dir


SCHEMA = """
CREATE TABLE IF NOT EXISTS imports (
    trade_date TEXT PRIMARY KEY,
    sha256 TEXT NOT NULL,
    original_name TEXT NOT NULL,
    saved_path TEXT NOT NULL,
    row_count INTEGER NOT NULL,
    imported_at TEXT NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS idx_import_hash ON imports(sha256);
CREATE TABLE IF NOT EXISTS flows (
    trade_date TEXT NOT NULL,
    code TEXT NOT NULL,
    name TEXT NOT NULL,
    pct_change REAL,
    turnover REAL,
    main_net REAL,
    main_ratio REAL,
    super_net REAL,
    big_net REAL,
    mid_net REAL,
    small_net REAL,
    PRIMARY KEY (trade_date, code)
);
CREATE INDEX IF NOT EXISTS idx_flows_code_date ON flows(code, trade_date);
"""


class HistoryStore:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or database_path()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as connection:
            connection.executescript(SCHEMA)

    def connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path, timeout=30)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA foreign_keys=ON")
        return connection

    def imported_dates(self) -> list[date]:
        with self.connect() as connection:
            values = connection.execute("SELECT trade_date FROM imports ORDER BY trade_date").fetchall()
        return [date.fromisoformat(row[0]) for row in values]

    def import_records(self) -> list[dict]:
        with self.connect() as connection:
            rows = connection.execute("SELECT * FROM imports ORDER BY trade_date DESC").fetchall()
        return [dict(row) for row in rows]

    def import_workbook(self, parsed: ParsedWorkbook, replace: bool = False) -> ImportResult:
        day = parsed.trade_date.isoformat()
        with self.connect() as connection:
            same_hash = connection.execute("SELECT trade_date FROM imports WHERE sha256=?", (parsed.sha256,)).fetchone()
            if same_hash:
                return ImportResult(parsed.trade_date, len(parsed.rows), "skipped", f"该文件已作为 {same_hash[0]} 导入，无需重复处理。")
            existing = connection.execute("SELECT sha256 FROM imports WHERE trade_date=?", (day,)).fetchone()
            if existing and not replace:
                return ImportResult(parsed.trade_date, len(parsed.rows), "conflict", "该交易日已有不同文件，需要确认后替换。")

        target_dir = imports_dir() / parsed.trade_date.strftime("%Y%m%d")
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / f"{parsed.sha256[:12]}_{parsed.path.name}"
        if not target.exists():
            shutil.copy2(parsed.path, target)

        values = [
            (
                row.trade_date.isoformat(), row.code, row.name, row.pct_change, row.turnover,
                row.main_net, row.main_ratio, row.super_net, row.big_net, row.mid_net, row.small_net,
            )
            for row in parsed.rows
        ]
        with self.connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            if existing:
                connection.execute("DELETE FROM flows WHERE trade_date=?", (day,))
                connection.execute("DELETE FROM imports WHERE trade_date=?", (day,))
            connection.executemany(
                """INSERT INTO flows(
                    trade_date,code,name,pct_change,turnover,main_net,main_ratio,
                    super_net,big_net,mid_net,small_net
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                values,
            )
            connection.execute(
                "INSERT INTO imports VALUES (?,?,?,?,?,?)",
                (day, parsed.sha256, parsed.path.name, str(target), len(values), datetime.now().isoformat(timespec="seconds")),
            )
        status = "replaced" if existing else "imported"
        verb = "替换" if existing else "导入"
        return ImportResult(parsed.trade_date, len(values), status, f"已{verb} {day}，共 {len(values):,} 只股票。")

    def delete_date(self, trade_date: date) -> None:
        day = trade_date.isoformat()
        with self.connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            connection.execute("DELETE FROM flows WHERE trade_date=?", (day,))
            connection.execute("DELETE FROM imports WHERE trade_date=?", (day,))

    def load_window(self, report_date: date, max_days: int = 11) -> tuple[list[date], list[dict]]:
        with self.connect() as connection:
            rows = connection.execute(
                "SELECT trade_date FROM imports WHERE trade_date<=? ORDER BY trade_date DESC LIMIT ?",
                (report_date.isoformat(), max_days),
            ).fetchall()
            dates = [date.fromisoformat(row[0]) for row in rows]
            if not dates:
                return [], []
            placeholders = ",".join("?" for _ in dates)
            flow_rows = connection.execute(
                f"SELECT * FROM flows WHERE trade_date IN ({placeholders}) ORDER BY code, trade_date DESC",
                tuple(day.isoformat() for day in dates),
            ).fetchall()
        return dates, [dict(row) for row in flow_rows]

