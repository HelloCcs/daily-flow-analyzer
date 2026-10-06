from __future__ import annotations

import json
import time
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Callable

import requests

from .paths import cache_dir


Progress = Callable[[str], None]


class ProviderError(RuntimeError):
    pass


@dataclass(frozen=True)
class MarketRecord:
    code: str
    close: float | None
    total_mv_wan: float | None


@dataclass(frozen=True)
class LimitRecord:
    code: str
    name: str
    limit_pct: float | None
    consecutive_boards: int | None
    board_text: str
    seal_amount_yuan: float | None
    first_seal_time: str
    last_seal_time: str
    break_count: int | None


def validate_tushare_token(token: str, trade_date: date, progress: Progress = lambda _: None) -> tuple[int, int]:
    """Verify both APIs required by a complete report without relying on cache."""
    token = token.strip()
    if not token:
        raise ProviderError("Token 为空。")
    counts = []
    for api_name, fields in (
        ("daily", "ts_code,trade_date,close,vol"),
        ("daily_basic", "ts_code,trade_date,close,total_mv"),
    ):
        progress(f"正在验证 Tushare {api_name} 接口...")
        payload = {
            "api_name": api_name,
            "token": token,
            "params": {"trade_date": trade_date.strftime("%Y%m%d")},
            "fields": fields,
        }
        body = _request_json("POST", "https://api.tushare.pro", progress, json=payload, timeout=(10, 60))
        if body.get("code") != 0:
            raise ProviderError(f"Tushare {api_name} 验证失败：{body.get('msg') or body.get('code')}")
        count = len(((body.get("data") or {}).get("items") or []))
        if count < 4000:
            raise ProviderError(f"Tushare {api_name} 仅返回 {count} 只股票，请确认所选日期是交易日。")
        counts.append(count)
    return counts[0], counts[1]


def fetch_volume_history(
    token: str,
    trade_dates: list[date],
    progress: Progress = lambda _: None,
) -> dict[str, dict[str, float]]:
    if not token.strip():
        return {}
    history: dict[str, dict[str, float]] = {}
    for index, trade_date in enumerate(trade_dates, start=1):
        cache_path = cache_dir() / f"volume_{trade_date:%Y%m%d}.json"
        cached = _read_cache(cache_path)
        if isinstance(cached, list) and len(cached) >= 4000:
            records = cached
        else:
            progress(f"正在获取成交量 {index}/{len(trade_dates)}：{trade_date}...")
            payload = {
                "api_name": "daily",
                "token": token.strip(),
                "params": {"trade_date": trade_date.strftime("%Y%m%d")},
                "fields": "ts_code,trade_date,vol",
            }
            body = _request_json("POST", "https://api.tushare.pro", progress, json=payload, timeout=(10, 60))
            if body.get("code") != 0:
                raise ProviderError(f"Tushare 成交量返回错误：{body.get('msg') or body.get('code')}")
            data = body.get("data") or {}
            fields, items = data.get("fields") or [], data.get("items") or []
            lookup = {name: position for position, name in enumerate(fields)}
            if not {"ts_code", "vol"}.issubset(lookup):
                raise ProviderError("Tushare 成交量响应缺少必要字段。")
            records = [
                {
                    "code": str(item[lookup["ts_code"]]).upper(),
                    "volume_hands": _float(item[lookup["vol"]]),
                }
                for item in items
            ]
            if len(records) < 4000:
                raise ProviderError(f"{trade_date} 成交量仅返回 {len(records)} 只股票，数据不完整。")
            _write_cache(cache_path, records)
        day = trade_date.isoformat()
        for record in records:
            volume = _float(record.get("volume_hands"))
            if volume is not None:
                history.setdefault(record["code"], {})[day] = volume
    progress(f"成交量历史获取完成（{len(trade_dates)} 个交易日，{len(history):,} 只股票）。")
    return history


def _read_cache(path: Path) -> dict | list | None:
    try:
        return json.loads(path.read_text("utf-8"))
    except (OSError, ValueError):
        return None


def _write_cache(path: Path, value: dict | list) -> None:
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(value, ensure_ascii=False), "utf-8")
    temp.replace(path)


def _float(value) -> float | None:
    try:
        return float(value) if value not in (None, "", "-") else None
    except (TypeError, ValueError):
        return None


def _request_json(method: str, url: str, progress: Progress, **kwargs) -> dict:
    last_error: Exception | None = None
    session = requests.Session()
    session.trust_env = False
    for attempt in range(1, 4):
        try:
            response = session.request(method, url, **kwargs)
            response.raise_for_status()
            return response.json()
        except Exception as exc:
            last_error = exc
            if attempt < 3:
                wait_seconds = attempt * 2
                progress(f"联网请求第 {attempt} 次失败，{wait_seconds} 秒后重试：{type(exc).__name__}")
                time.sleep(wait_seconds)
    raise ProviderError(f"连续 3 次联网请求失败：{type(last_error).__name__}: {last_error}") from last_error


def fetch_market_data(token: str, trade_date: date, progress: Progress = lambda _: None) -> dict[str, MarketRecord]:
    cache_path = cache_dir() / f"market_{trade_date:%Y%m%d}.json"
    cached = _read_cache(cache_path)
    if isinstance(cached, list) and len(cached) >= 4000:
        progress(f"已读取 {trade_date} 的市值缓存（{len(cached):,} 只）。")
        return {
            row["code"]: MarketRecord(row["code"], row.get("close"), row.get("total_mv_wan"))
            for row in cached
        }
    if not token.strip():
        raise ProviderError("未设置 Tushare Token，无法取得报告日的精确总市值。")
    progress(f"正在通过 Tushare 获取 {trade_date} 收盘价和总市值...")
    payload = {
        "api_name": "daily_basic",
        "token": token.strip(),
        "params": {"trade_date": trade_date.strftime("%Y%m%d")},
        "fields": "ts_code,trade_date,close,total_mv",
    }
    body = _request_json("POST", "https://api.tushare.pro", progress, json=payload, timeout=(10, 60))
    if body.get("code") != 0:
        raise ProviderError(f"Tushare 返回错误：{body.get('msg') or body.get('code')}")
    data = body.get("data") or {}
    fields, items = data.get("fields") or [], data.get("items") or []
    lookup = {name: index for index, name in enumerate(fields)}
    required = {"ts_code", "close", "total_mv"}
    if not required.issubset(lookup):
        raise ProviderError("Tushare 市值响应缺少必要字段。")
    records = []
    for item in items:
        records.append(
            {
                "code": str(item[lookup["ts_code"]]).upper(),
                "close": _float(item[lookup["close"]]),
                "total_mv_wan": _float(item[lookup["total_mv"]]),
            }
        )
    if len(records) < 4000:
        raise ProviderError(f"Tushare 仅返回 {len(records)} 只股票，疑似非交易日或权限/数据不完整。")
    _write_cache(cache_path, records)
    progress(f"市值数据获取完成（{len(records):,} 只）。")
    return {row["code"]: MarketRecord(**row) for row in records}


def _format_time(value) -> str:
    text = str(value or "").strip()
    if not text or text == "0":
        return ""
    text = text.zfill(6)[-6:]
    return f"{text[:2]}:{text[2:4]}:{text[4:]}"


def fetch_limit_pool(trade_date: date, progress: Progress = lambda _: None) -> dict[str, LimitRecord]:
    cache_path = cache_dir() / f"limit_{trade_date:%Y%m%d}.json"
    cached = _read_cache(cache_path)
    if isinstance(cached, list):
        progress(f"已读取 {trade_date} 的涨停缓存（{len(cached)} 只）。")
        return {row["code"]: LimitRecord(**row) for row in cached}
    progress(f"正在获取 {trade_date} 涨停池和封单数据...")
    params = {
        "ut": "7eea3edcaed734bea9cbfc24409ed989",
        "dpt": "wz.ztzt",
        "Pageindex": 0,
        "pagesize": 10000,
        "sort": "fbt:asc",
        "date": trade_date.strftime("%Y%m%d"),
        "_": int(time.time() * 1000),
    }
    body = _request_json("GET", "https://push2ex.eastmoney.com/getTopicZTPool", progress, params=params, timeout=(10, 45))
    pool = ((body.get("data") or {}).get("pool") or [])
    records = []
    for item in pool:
        code = str(item.get("c") or "").zfill(6)
        if code.startswith(("8", "9")):
            suffix = ".BJ"
        elif code.startswith(("5", "6")):
            suffix = ".SH"
        else:
            suffix = ".SZ"
        stats = item.get("zttj") or {}
        days, count = stats.get("days"), stats.get("ct")
        board_text = f"{days}天{count}板" if days and count else ""
        records.append(
            {
                "code": code + suffix,
                "name": str(item.get("n") or ""),
                "limit_pct": _float(item.get("zdp")),
                "consecutive_boards": int(item["lbc"]) if item.get("lbc") is not None else None,
                "board_text": board_text,
                "seal_amount_yuan": _float(item.get("fund")),
                "first_seal_time": _format_time(item.get("fbt")),
                "last_seal_time": _format_time(item.get("lbt")),
                "break_count": int(item["zbc"]) if item.get("zbc") is not None else None,
            }
        )
    if not records:
        raise ProviderError("涨停池返回为空，不能据此判定当日涨停情况。")
    _write_cache(cache_path, records)
    progress(f"涨停数据获取完成（{len(records)} 只）。")
    return {row["code"]: LimitRecord(**row) for row in records}
