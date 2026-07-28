from __future__ import annotations

import os

import pandas as pd


def today() -> pd.Timestamp:
    override = os.environ.get("WXT_TODAY", "").strip()
    if override:
        parsed = pd.to_datetime(override, errors="coerce")
        if not pd.isna(parsed):
            return parsed.normalize()
    return pd.Timestamp.today().normalize()


def current_previous_windows(period: str, base_date: pd.Timestamp | None = None) -> tuple[tuple[pd.Timestamp, pd.Timestamp], tuple[pd.Timestamp, pd.Timestamp]]:
    base = (base_date or today()).normalize()
    if period == "week":
        current_end = base
        current_start = current_end - pd.Timedelta(days=6)
        previous_end = current_start - pd.Timedelta(days=1)
        previous_start = previous_end - pd.Timedelta(days=6)
    elif period == "half-month":
        current_end = base
        current_start = current_end - pd.Timedelta(days=14)
        previous_end = current_start - pd.Timedelta(days=1)
        previous_start = previous_end - pd.Timedelta(days=14)
    else:
        current_end = base
        current_start = current_end - pd.Timedelta(days=29)
        previous_end = current_start - pd.Timedelta(days=1)
        previous_start = previous_end - pd.Timedelta(days=29)
    return (current_start, current_end), (previous_start, previous_end)


def range_label(start: pd.Timestamp, end: pd.Timestamp) -> str:
    return f"{start:%Y-%m-%d}~{end:%Y-%m-%d}"


def label_for_date(dates: pd.Series, period: str, base_date: pd.Timestamp | None = None) -> pd.Series:
    current, previous = current_previous_windows(period, base_date)
    current_label = range_label(*current)
    previous_label = range_label(*previous)
    labels = pd.Series(pd.NA, index=dates.index, dtype="object")
    labels.loc[dates.between(current[0], current[1], inclusive="both")] = current_label
    labels.loc[dates.between(previous[0], previous[1], inclusive="both")] = previous_label
    return labels
