from __future__ import annotations

import base64
import io
import json
import mimetypes
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from .ai_provider import AIConfig, call_llm_vision_json


IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
MAX_IMAGES_PER_VISION_REQUEST = 8
CANONICAL_COLUMNS = [
    "统计日期",
    "店铺名称",
    "访客数",
    "支付金额",
    "支付买家数",
    "支付转化率",
    "全站推广花费",
    "广告引导成交金额",
    "点击量",
    "展现量",
]
ALIASES = {
    "统计日期": ("统计日期", "日期", "时间范围", "统计周期", "周期"),
    "店铺名称": ("店铺名称", "店铺", "店铺名"),
    "访客数": ("访客数", "店铺总访客数", "访客"),
    "支付金额": ("支付金额", "店铺总成交金额", "成交金额", "成交额"),
    "支付买家数": ("支付买家数", "支付人数", "支付客户数", "成交买家数"),
    "支付转化率": ("支付转化率", "转化率"),
    "全站推广花费": ("全站推广花费", "广告总花费", "推广花费", "广告花费", "消耗"),
    "广告引导成交金额": ("广告引导成交金额", "广告带来的成交金额", "广告成交金额", "引导成交金额"),
    "点击量": ("点击量", "总点击量", "点击"),
    "展现量": ("展现量", "曝光量", "展示量"),
}


@dataclass(frozen=True)
class ScreenshotRecognitionResult:
    report_files: list[Path]
    audits: list[dict[str, Any]]


def extract_screenshot_report_files(
    image_files: list[Path], output_dir: Path, ai_config: AIConfig
) -> ScreenshotRecognitionResult:
    """Merge complementary metric screenshots into one structured workbook input."""
    output_dir.mkdir(parents=True, exist_ok=True)
    audits: list[dict[str, Any]] = []
    prepared: list[tuple[Path, str]] = []
    for source in image_files:
        try:
            prepared.append((source, image_data_url(source)))
            audits.append({"path": str(source), "status": "submitted", "kind": "screenshot_metrics"})
        except Exception as exc:
            audits.append({"path": str(source), "status": "failed", "kind": "screenshot_metrics", "reason": str(exc)})

    records: list[dict[str, Any]] = []
    batch_errors: list[str] = []
    for start in range(0, len(prepared), MAX_IMAGES_PER_VISION_REQUEST):
        chunk = prepared[start : start + MAX_IMAGES_PER_VISION_REQUEST]
        try:
            payload = call_llm_vision_json(_prompt(len(chunk)), [url for _, url in chunk], ai_config)
            records.extend(normalize_metric_records(payload))
        except Exception as exc:  # Other chunks can still provide usable data.
            batch_errors.append(str(exc))

    merged_records, conflicts = merge_complementary_records(records)
    usable_records = [record for record in merged_records if record_is_usable(record)]
    if not usable_records:
        reason = "截图中没有识别到同时包含日期和经营指标的可用数据"
        if batch_errors:
            reason += "；" + "；".join(batch_errors[:2])
        audits.append({"status": "failed", "kind": "screenshot_batch", "reason": reason})
        return ScreenshotRecognitionResult(report_files=[], audits=audits)

    target = output_dir / "_screenshot_metrics.xlsx"
    pd.DataFrame(usable_records, columns=CANONICAL_COLUMNS).to_excel(target, index=False)
    for audit in audits:
        if audit.get("status") == "submitted":
            audit["status"] = "ok"
            audit["output"] = str(target)
    audits.append(
        {
            "status": "ok",
            "kind": "screenshot_batch",
            "output": str(target),
            "image_count": len(prepared),
            "record_count": len(usable_records),
            "conflicts": conflicts,
            "warnings": batch_errors,
        }
    )
    return ScreenshotRecognitionResult(report_files=[target.resolve()], audits=audits)


def image_data_url(path: Path) -> str:
    try:
        from PIL import Image

        with Image.open(path) as image:
            image = image.convert("RGB")
            image.thumbnail((2200, 2200))
            buffer = io.BytesIO()
            image.save(buffer, format="JPEG", quality=88, optimize=True)
        payload = buffer.getvalue()
        mime = "image/jpeg"
    except (ImportError, OSError):
        payload = path.read_bytes()
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    encoded = base64.b64encode(payload).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def normalize_metric_records(payload: Any) -> list[dict[str, Any]]:
    candidates: Any = payload
    if isinstance(payload, dict):
        if isinstance(payload.get("records"), list):
            candidates = payload["records"]
        elif isinstance(payload.get("data"), list):
            candidates = payload["data"]
        elif isinstance(payload.get("data"), dict):
            candidates = [payload["data"]]
        else:
            candidates = [payload]
    if not isinstance(candidates, list):
        raise ValueError("截图识别结果不是指标记录列表")
    records = [normalize_metric_record(item) for item in candidates if isinstance(item, dict)]
    if not records:
        raise ValueError("截图识别结果没有可用指标记录")
    return records


def merge_complementary_records(records: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[str]]:
    prepared = [{column: record.get(column, "") for column in CANONICAL_COLUMNS} for record in records]
    known_dates = {str(record["统计日期"]).strip() for record in prepared if str(record["统计日期"]).strip()}
    known_stores = {str(record["店铺名称"]).strip() for record in prepared if str(record["店铺名称"]).strip()}
    for record in prepared:
        if not str(record["统计日期"]).strip() and len(known_dates) == 1:
            record["统计日期"] = next(iter(known_dates))
        if not str(record["店铺名称"]).strip() and len(known_stores) == 1:
            record["店铺名称"] = next(iter(known_stores))

    grouped: dict[tuple[str, str], dict[str, Any]] = {}
    conflicts: list[str] = []
    for record in prepared:
        key = (str(record["统计日期"]).strip(), str(record["店铺名称"]).strip())
        target = grouped.setdefault(key, {column: "" for column in CANONICAL_COLUMNS})
        for column in CANONICAL_COLUMNS:
            value = record.get(column, "")
            if value in (None, ""):
                continue
            existing = target.get(column, "")
            if existing in (None, ""):
                target[column] = value
            elif existing != value:
                conflicts.append(f"{key[0] or '未知日期'}/{key[1] or '未知店铺'}：{column}存在多个值，保留首个识别值")
    return list(grouped.values()), list(dict.fromkeys(conflicts))


def record_is_usable(record: dict[str, Any]) -> bool:
    if not str(record.get("统计日期") or "").strip():
        return False
    return any(record.get(column) not in (None, "") for column in CANONICAL_COLUMNS if column not in {"统计日期", "店铺名称"})


def normalize_metric_record(payload: Any) -> dict[str, Any]:
    if isinstance(payload, dict) and isinstance(payload.get("data"), dict):
        payload = payload["data"]
    if not isinstance(payload, dict):
        raise ValueError("截图识别结果不是指标对象")
    result: dict[str, Any] = {column: "" for column in CANONICAL_COLUMNS}
    for canonical, aliases in ALIASES.items():
        value = next((payload.get(alias) for alias in aliases if payload.get(alias) not in (None, "")), "")
        result[canonical] = normalize_value(canonical, value)
    return result


def normalize_value(column: str, value: Any) -> Any:
    if value in (None, ""):
        return ""
    if column in {"统计日期", "店铺名称"}:
        return str(value).strip()
    text = str(value).replace(",", "").replace("¥", "").replace("元", "").strip()
    if column == "支付转化率":
        if text.endswith("%"):
            return _number(text[:-1]) / 100
        number = _number(text)
        return number / 100 if number > 1 else number
    return _number(text)


def _number(value: str) -> float:
    match = re.search(r"[-+]?\d+(?:\.\d+)?", value)
    if not match:
        return 0.0
    return float(match.group())


def _prompt(image_count: int = 1) -> str:
    return f"""读取这一批共 {image_count} 张电商/万相台/生意参谋数据截图，提取清晰可见的店铺汇总指标。
这些图片可能是同一店铺、同一日期的不同页面。请合并互补字段，但不要把重复截图中的数值相加；不同日期或不同店铺必须拆成不同记录。
只返回 JSON 对象，格式为 {{"records": [{{...}}]}}；不能猜测，无法确认的字段返回空字符串。
每条记录字段仅允许：统计日期、店铺名称、访客数、支付金额、支付买家数、支付转化率、全站推广花费、广告引导成交金额、点击量、展现量。
统计日期必须原样保留截图中可见的日期或日期范围。支付转化率可返回百分号字符串。不要写任何解释文字。"""
