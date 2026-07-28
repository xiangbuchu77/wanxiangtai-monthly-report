from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from .utils import ensure_dir


SUPPORTED_INPUT_SUFFIXES = {".csv", ".xls", ".xlsx", ".doc", ".docx", ".png", ".jpg", ".jpeg", ".zip"}


def clean_dir(path: str | Path) -> Path:
    target = Path(path)
    if target.exists():
        shutil.rmtree(target)
    target.mkdir(parents=True, exist_ok=True)
    return target


def extract_zip(path: str | Path, output_dir: str | Path) -> list[Path]:
    source = Path(path)
    target = ensure_dir(output_dir)
    extracted: list[Path] = []
    with zipfile.ZipFile(source) as archive:
        for member in archive.infolist():
            if member.is_dir():
                continue
            archive.extract(member, target)
            extracted.append(target / member.filename)
    return extracted
