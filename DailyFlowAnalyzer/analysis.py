from __future__ import annotations

from collections import defaultdict
from datetime import date

from .providers import LimitRecord, MarketRecord


CATEGORIES = (
    ("main_net", "主力"),
    ("super_net", "特大单"),
    ("big_net", "大单"),
    ("mid_net", "中单"),
)
WINDOWS = (10, 7, 5, 3)
PROGRESSION_PAIRS = ((10, 7), (7, 5), (5, 3))


def growth(current: float | None, previous: float | None) -> tuple[float | None, str]:
    if current is None or previous is None:
        return None, "数据不足"
    if previous == 0:
        if current > 0:
            return None, "由零转正"
        if current < 0:
            return None, "由零转负"
        return 0.0, "持平"
    rate = (current - previous) / abs(previous)
    if previous < 0 < current:
        return rate, "由负转正"
    if previous > 0 > current:
        return rate, "由正转负"
    return rate, "增加" if current > previous else "减少" if current < previous else "持平"


def _sum(values: list[float | None]) -> float | None:
    present = [value for value in values if value is not None]
    return sum(present) if present else None


def build_analysis_rows(
    dates: list[date],
    flow_rows: list[dict],
    market: dict[str, MarketRecord],
    limits: dict[str, LimitRecord],
    volume_history: dict[str, dict[str, float]] | None = None,
) -> list[dict]:
    if not dates:
        return []
    latest = dates[0].isoformat()
    history: dict[str, dict[str, dict]] = defaultdict(dict)
    for row in flow_rows:
        history[row["code"]][row["trade_date"]] = row

    volume_history = volume_history or {}
    output = []
    latest_codes = [code for code, items in history.items() if latest in items]
    latest_codes.sort(key=lambda code: ((history[code][latest].get("main_net") is None), -(history[code][latest].get("main_net") or 0)))
    for sequence, code in enumerate(latest_codes, start=1):
        current = history[code][latest]
        market_row = market.get(code)
        limit_row = limits.get(code)
        total_mv_wan = market_row.total_mv_wan if market_row else None
        row = {
            "序号": sequence,
            "代码": code,
            "名称": current["name"],
            "交易日期": dates[0],
            "涨跌幅(%)": current.get("pct_change"),
            "收盘价(元)": market_row.close if market_row else None,
            "总市值(亿元)": total_mv_wan / 10000 if total_mv_wan is not None else None,
            "成交额(万元)": current.get("turnover"),
            "主力资金流入比率(%)": current.get("main_ratio"),
            "主力净额占总市值(%)": (current.get("main_net") / total_mv_wan * 100)
            if current.get("main_net") is not None and total_mv_wan
            else None,
            "是否涨停": "是" if limit_row else "否",
            "涨停幅度(%)": limit_row.limit_pct if limit_row else None,
            "几天几板": limit_row.board_text if limit_row else "",
            "连续涨停板数": limit_row.consecutive_boards if limit_row else None,
            "封单资金(万元)": limit_row.seal_amount_yuan / 10000 if limit_row and limit_row.seal_amount_yuan is not None else None,
            "首次封板时间": limit_row.first_seal_time if limit_row else "",
            "最后封板时间": limit_row.last_seal_time if limit_row else "",
            "炸板次数": limit_row.break_count if limit_row else None,
        }
        today_turnover = current.get("turnover")
        row["今日成交额占总市值(%)"] = (
            today_turnover / total_mv_wan * 100 if today_turnover is not None and total_mv_wan else None
        )
        volumes = volume_history.get(code, {})
        for index, day in enumerate(dates[:10]):
            volume = volumes.get(day.isoformat())
            previous_volume = volumes.get(dates[index + 1].isoformat()) if index + 1 < len(dates) else None
            row[f"成交量日期{index + 1}"] = day
            row[f"成交量{index + 1}(万手)"] = volume / 10000 if volume is not None else None
            row[f"倍量{index + 1}(倍)"] = (
                volume / previous_volume
                if volume is not None and previous_volume not in (None, 0)
                else None
            )
        previous_main_ratio = history[code].get(dates[1].isoformat(), {}).get("main_ratio") if len(dates) >= 2 else None
        row["今日主力资金流入比率(%)"] = current.get("main_ratio")
        row["昨日主力资金流入比率(%)"] = previous_main_ratio
        row["日比率分差(百分点)"] = (
            current.get("main_ratio") - previous_main_ratio
            if current.get("main_ratio") is not None and previous_main_ratio is not None
            else None
        )
        for field, label in CATEGORIES:
            today = current.get(field)
            previous_row = history[code].get(dates[1].isoformat(), {}) if len(dates) >= 2 else {}
            previous = previous_row.get(field)
            previous_turnover = previous_row.get("turnover")
            rate, status = growth(today, previous)
            row[f"今日{label}净额(万元)"] = today
            row[f"昨日{label}净额(万元)"] = previous
            row[f"{label}净额增减(万元)"] = today - previous if today is not None and previous is not None else None
            row[f"{label}净额增长率(%)"] = rate * 100 if rate is not None else None
            row[f"{label}变化状态"] = status
            row[f"今日{label}资金流入比率(%)"] = (
                today / today_turnover * 100 if today is not None and today_turnover else None
            )
            row[f"昨日{label}资金流入比率(%)"] = (
                previous / previous_turnover * 100 if previous is not None and previous_turnover else None
            )
            row[f"{label}资金流入比率分差(百分点)"] = (
                row[f"今日{label}资金流入比率(%)"] - row[f"昨日{label}资金流入比率(%)"]
                if row[f"今日{label}资金流入比率(%)"] is not None
                and row[f"昨日{label}资金流入比率(%)"] is not None
                else None
            )
            row[f"今日{label}净额占总市值(%)"] = (
                today / total_mv_wan * 100 if today is not None and total_mv_wan else None
            )
        for window in WINDOWS:
            current_dates = dates[:window]
            previous_dates = dates[1 : window + 1]
            current_turnover = _sum([history[code].get(day.isoformat(), {}).get("turnover") for day in current_dates]) if len(current_dates) == window else None
            previous_turnover = _sum([history[code].get(day.isoformat(), {}).get("turnover") for day in previous_dates]) if len(previous_dates) == window else None
            for field, label in CATEGORIES:
                current_sum = _sum([history[code].get(day.isoformat(), {}).get(field) for day in current_dates]) if len(current_dates) == window else None
                previous_sum = _sum([history[code].get(day.isoformat(), {}).get(field) for day in previous_dates]) if len(previous_dates) == window else None
                period_rate, _ = growth(current_sum, previous_sum)
                row[f"{window}日{label}净额合计(万元)"] = current_sum
                row[f"{window}日{label}较上期差值(万元)"] = (
                    current_sum - previous_sum if current_sum is not None and previous_sum is not None else None
                )
                row[f"{window}日含今日{label}净额(万元)"] = current_sum
                row[f"{window}日不含今日{label}净额(万元)"] = previous_sum
                row[f"{window}日{label}增减量(万元)"] = (
                    current_sum - previous_sum if current_sum is not None and previous_sum is not None else None
                )
                row[f"{window}日{label}增减率(%)"] = period_rate * 100 if period_rate is not None else None
                current_flow_ratio = current_sum / current_turnover * 100 if current_sum is not None and current_turnover else None
                previous_flow_ratio = previous_sum / previous_turnover * 100 if previous_sum is not None and previous_turnover else None
                row[f"{window}日含今日{label}资金流入比率(%)"] = current_flow_ratio
                row[f"{window}日不含今日{label}资金流入比率(%)"] = previous_flow_ratio
                row[f"{window}日{label}资金流入比率分差(百分点)"] = (
                    current_flow_ratio - previous_flow_ratio
                    if current_flow_ratio is not None and previous_flow_ratio is not None
                    else None
                )
                row[f"{window}日含今日{label}净额占总市值(%)"] = (
                    current_sum / total_mv_wan * 100 if current_sum is not None and total_mv_wan else None
                )
            current_main = row.get(f"{window}日主力净额合计(万元)")
            previous_main = _sum([history[code].get(day.isoformat(), {}).get("main_net") for day in previous_dates]) if len(previous_dates) == window else None
            current_ratio = current_main / current_turnover * 100 if current_main is not None and current_turnover else None
            previous_ratio = previous_main / previous_turnover * 100 if previous_main is not None and previous_turnover else None
            row[f"{window}日含今日成交额(万元)"] = current_turnover
            row[f"{window}日不含今日成交额(万元)"] = previous_turnover
            row[f"{window}日含今日主力资金流入比率(%)"] = current_ratio
            row[f"{window}日不含今日主力资金流入比率(%)"] = previous_ratio
            row[f"{window}日比率分差(百分点)"] = (
                current_ratio - previous_ratio if current_ratio is not None and previous_ratio is not None else None
            )
            row[f"{window}日含今日成交额占总市值(%)"] = (
                current_turnover / total_mv_wan * 100 if current_turnover is not None and total_mv_wan else None
            )
        observation_windows = (10, 7, 5, 3)
        for shorter, longer in zip(observation_windows, observation_windows[1:]):
            shorter_ratio = row.get(f"{shorter}日含今日主力资金流入比率(%)")
            longer_ratio = row.get(f"{longer}日含今日主力资金流入比率(%)")
            row[f"{shorter}至{longer}日比率变化(百分点)"] = (
                longer_ratio - shorter_ratio
                if shorter_ratio is not None and longer_ratio is not None
                else None
            )
        observation_ratios = [row.get(f"{window}日含今日主力资金流入比率(%)") for window in observation_windows]
        if all(value is not None for value in observation_ratios):
            row["逐步增减值(百分点)"] = observation_ratios[-1] - observation_ratios[0]
            if all(left < right for left, right in zip(observation_ratios, observation_ratios[1:])):
                row["是否逐步增长"] = "是"
                row["周期资金变化观察"] = "越接近今日逐步增长"
            elif all(left > right for left, right in zip(observation_ratios, observation_ratios[1:])):
                row["是否逐步增长"] = "否"
                row["周期资金变化观察"] = "越接近今日逐步减少"
            elif observation_ratios[-1] < 0 < observation_ratios[0]:
                row["是否逐步增长"] = "否"
                row["周期资金变化观察"] = "由流出转流入"
            elif observation_ratios[-1] > 0 > observation_ratios[0]:
                row["是否逐步增长"] = "否"
                row["周期资金变化观察"] = "由流入转流出"
            else:
                row["是否逐步增长"] = "否"
                row["周期资金变化观察"] = "区间震荡变化"
        else:
            row["逐步增减值(百分点)"] = None
            row["是否逐步增长"] = "数据不足"
            row["周期资金变化观察"] = "数据不足"
        for scope in ("含今日", "不含今日"):
            for shorter, longer in PROGRESSION_PAIRS:
                for _, label in CATEGORIES:
                    shorter_value = row.get(f"{shorter}日{scope}{label}净额(万元)")
                    longer_value = row.get(f"{longer}日{scope}{label}净额(万元)")
                    increase = (
                        longer_value - shorter_value
                        if longer_value is not None and shorter_value is not None
                        else None
                    )
                    increase_rate, _ = growth(longer_value, shorter_value)
                    row[f"{scope}{shorter}至{longer}日{label}增加量(万元)"] = increase
                    row[f"{scope}{shorter}至{longer}日{label}增减率(%)"] = (
                        increase_rate * 100 if increase_rate is not None else None
                    )
                    shorter_ratio = row.get(f"{shorter}日{scope}{label}资金流入比率(%)")
                    longer_ratio = row.get(f"{longer}日{scope}{label}资金流入比率(%)")
                    row[f"{scope}{shorter}至{longer}日{label}比率变化(百分点)"] = (
                        longer_ratio - shorter_ratio
                        if longer_ratio is not None and shorter_ratio is not None
                        else None
                    )
        output.append(row)
    return output
