from __future__ import annotations

import re
import shutil
import unicodedata
import zipfile
from pathlib import Path


WINDOWS_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{index}" for index in range(1, 10)),
    *(f"LPT{index}" for index in range(1, 10)),
}
WINDOWS_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
MAX_EXTRACTED_FILENAME_LENGTH = 120


class ZipArchiveError(ValueError):
    """A ZIP cannot be used as a report input."""


def sanitize_windows_filename(filename: str | None) -> str:
    """Return a flat filename that is valid on Windows, macOS and Linux."""
    if not filename:
        return ""
    leaf = str(filename).replace("\x00", "").replace("\\", "/").rsplit("/", 1)[-1]
    leaf = unicodedata.normalize("NFKC", leaf)
    leaf = WINDOWS_INVALID_FILENAME_CHARS.sub("_", leaf).strip().rstrip(". ")
    if not leaf or leaf in {".", ".."}:
        return ""
    if leaf.split(".", 1)[0].upper() in WINDOWS_RESERVED_NAMES:
        leaf = f"_{leaf}"
    if len(leaf) > MAX_EXTRACTED_FILENAME_LENGTH:
        suffix = Path(leaf).suffix
        stem_limit = max(1, MAX_EXTRACTED_FILENAME_LENGTH - len(suffix))
        leaf = f"{leaf[:stem_limit].rstrip('. ')}{suffix}"
    return leaf


def decode_zip_member_name(info: zipfile.ZipInfo) -> str:
    """Recover common GBK/GB18030 filenames written by older Windows tools."""
    name = info.filename
    if not name or name.isascii() or info.flag_bits & 0x800:
        return name
    try:
        raw_name = name.encode("cp437")
    except UnicodeEncodeError:
        return name
    for encoding in ("gb18030", "utf-8"):
        try:
            candidate = raw_name.decode(encoding)
        except UnicodeDecodeError:
            continue
        if _contains_cjk(candidate) and not _contains_cjk(name):
            return candidate
    return name


def extract_zip_archive(zip_path: Path, target_dir: Path, allowed_suffixes: set[str]) -> list[Path]:
    source = Path(zip_path).expanduser()
    if not source.exists():
        raise ZipArchiveError(f"压缩包不存在：{source.name}")
    if not source.is_file() or not zipfile.is_zipfile(source):
        raise ZipArchiveError(f"压缩包损坏或不是有效的 ZIP 文件：{source.name}")

    target = Path(target_dir)
    target.mkdir(parents=True, exist_ok=True)
    allowed = {suffix.lower() for suffix in allowed_suffixes} - {".zip"}
    extracted: list[Path] = []
    try:
        with zipfile.ZipFile(source, mode="r") as archive:
            for info in archive.infolist():
                member_name = decode_zip_member_name(info).replace("\\", "/")
                parts = [part for part in member_name.split("/") if part not in {"", "."}]
                if info.is_dir() or not parts:
                    continue
                if ".." in parts or "__MACOSX" in parts or parts[-1].startswith("."):
                    continue
                filename = sanitize_windows_filename(parts[-1])
                if not filename or Path(filename).suffix.lower() not in allowed:
                    continue
                if info.flag_bits & 0x1:
                    raise ZipArchiveError(f"压缩包包含加密文件，暂不支持自动解包：{filename}")
                destination = _unique_path(target / filename)
                partial = destination.with_name(f".{destination.name}.part")
                try:
                    with archive.open(info, mode="r") as src, partial.open("wb") as out:
                        shutil.copyfileobj(src, out)
                    partial.replace(destination)
                finally:
                    partial.unlink(missing_ok=True)
                extracted.append(destination.resolve())
    except ZipArchiveError:
        shutil.rmtree(target, ignore_errors=True)
        raise
    except (zipfile.BadZipFile, zipfile.LargeZipFile, RuntimeError, OSError, EOFError) as exc:
        shutil.rmtree(target, ignore_errors=True)
        raise ZipArchiveError(f"压缩包损坏、被加密或读取失败：{source.name}") from exc

    if not extracted:
        shutil.rmtree(target, ignore_errors=True)
        supported = "、".join(sorted(allowed))
        raise ZipArchiveError(f"压缩包中没有可处理的报表或图片（支持：{supported}）：{source.name}")
    return extracted


def _contains_cjk(value: str) -> bool:
    return any("\u3400" <= char <= "\u9fff" for char in value)


def _unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    for index in range(2, 1000):
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise ZipArchiveError(f"压缩包内重名文件过多，无法保存：{path.name}")
