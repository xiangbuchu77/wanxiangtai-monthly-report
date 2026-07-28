from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook


def open_workbook(path: str | Path):
    return load_workbook(Path(path))


def save_workbook(workbook, path: str | Path) -> Path:
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    return output
