from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd


@dataclass(frozen=True)
class LoadedReport:
    path: Path
    sheet_name: str
    raw: pd.DataFrame


def read_excel_reports(paths: Iterable[Path]) -> list[LoadedReport]:
    reports: list[LoadedReport] = []
    for path in paths:
        workbook = pd.ExcelFile(path)
        for sheet_name in workbook.sheet_names:
            frame = pd.read_excel(path, sheet_name=sheet_name, dtype=object)
            frame = frame.dropna(axis=0, how="all")
            if not frame.empty:
                frame.columns = [str(column).strip() for column in frame.columns]
                reports.append(LoadedReport(path=path, sheet_name=sheet_name, raw=frame))
    return reports
