from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


REPORT_SUFFIXES = {".xlsx", ".xls", ".csv"}


@dataclass(frozen=True)
class InputPacket:
    text: str
    files: list[Path]
    settings: dict[str, Any]


@dataclass(frozen=True)
class RequestDecision:
    session_key: str
    action: str
    period: str
    target: str | None
    report_files: list[Path]
    document_files: list[Path]
    should_collect: bool
    should_generate: bool


def split_uploaded_files(files: list[Path]) -> tuple[list[Path], list[Path]]:
    report_files = [path.resolve() for path in files if path.suffix.lower() in REPORT_SUFFIXES and path.exists()]
    document_files = [path.resolve() for path in files if path.suffix.lower() not in REPORT_SUFFIXES and path.exists()]
    return report_files, document_files


def user_has_confirmed_upload(action: str) -> bool:
    return action in {"confirm", "supplement", "new"}
