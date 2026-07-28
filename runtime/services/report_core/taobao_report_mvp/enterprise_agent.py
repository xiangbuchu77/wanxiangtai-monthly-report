from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path
from typing import Any

from .report_workflow_agent import (
    DEFAULT_TEMPLATE_CANDIDATES,
    build_output_paths,
    first_existing,
    run_workflow,
)


SUPPORTED_SUFFIXES = {".xlsx", ".xls", ".csv"}
CORE_METRIC_LABELS = [
    "店铺总访客数",
    "店铺总成交金额（元）",
    "店铺支付转化率",
    "广告总花费（元）",
    "广告带来的成交金额（元）",
    "广告投入产出比（ROI）",
    "平均点击成本（元）",
    "总点击量",
]
PERIOD_ALIASES = {
    "month": "month",
    "monthly": "month",
    "月报": "month",
    "月": "month",
    "week": "week",
    "weekly": "week",
    "周报": "week",
    "周": "week",
    "half-month": "half-month",
    "half_month": "half-month",
    "半月报": "half-month",
    "半月": "half-month",
}


def main() -> None:
    parser = argparse.ArgumentParser(
        description="企业版万相台报表 Agent：从上传原始报表到最终 Excel 月报/周报/半月报。"
    )
    parser.add_argument("uploads", nargs="*", help="上传的原始报表文件或文件夹，支持 .xlsx/.xls/.csv")
    parser.add_argument("--input-dir", type=Path, action="append", help="上传报表目录；可传多次")
    parser.add_argument("--template", type=Path, help="万相台月报模板路径；不传则自动寻找默认模板")
    parser.add_argument("--output-dir", type=Path, help="输出目录；不传则自动创建企业交付目录")
    parser.add_argument(
        "--report-type",
        "--period",
        dest="period",
        default="month",
        help="报表周期：月报/month、周报/week、半月报/half-month；默认月报",
    )
    parser.add_argument("--target", "--month", dest="target", help="指定目标周期；不传则按数据自动拆分生成")
    parser.add_argument("--keep-json", action="store_true", help="保留 AI JSON 和任务包；默认仍保留便于审计")
    args = parser.parse_args()

    period = normalize_period(args.period)
    uploads = collect_uploads(args.uploads, args.input_dir or [])
    output_dir = args.output_dir or default_output_dir(period)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not uploads:
        core_metric_text = "、".join(CORE_METRIC_LABELS)
        result = {
            "status": "blocked",
            "message": f"没有找到可处理的数据。请补充以下指标：{core_metric_text}。文件名不限，系统会按字段识别。",
            "core_dataset_name": "店铺收益核心指标",
            "expected_core_data": [
                *CORE_METRIC_LABELS,
                "商品经营数据：商品名称、商品访客、支付金额、支付转化率（可选，用于主要商品模块）",
            ],
        }
        write_enterprise_manifest(output_dir, uploads, period, args.target, result)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return

    template = args.template or first_existing(DEFAULT_TEMPLATE_CANDIDATES)
    outputs = build_output_paths(output_dir)
    result = run_workflow(uploads, template, outputs, args.target, period)
    enterprise_result = build_enterprise_result(result, uploads, output_dir, period, args.target)
    write_enterprise_manifest(output_dir, uploads, period, args.target, enterprise_result)
    write_handoff_note(output_dir, enterprise_result)
    print(json.dumps(enterprise_result, ensure_ascii=False, indent=2))


def normalize_period(value: str) -> str:
    key = str(value).strip().lower()
    if key not in PERIOD_ALIASES:
        raise SystemExit(f"不支持的报表周期: {value}。请使用 月报/month、周报/week、半月报/half-month。")
    return PERIOD_ALIASES[key]


def collect_uploads(upload_args: list[str], input_dirs: list[Path]) -> list[Path]:
    candidates: list[Path] = []
    for item in upload_args:
        path = Path(item).expanduser()
        if path.is_dir():
            candidates.extend(scan_dir(path))
        else:
            candidates.append(path)
    for directory in input_dirs:
        candidates.extend(scan_dir(directory.expanduser()))

    unique: list[Path] = []
    seen: set[Path] = set()
    for path in candidates:
        if path.suffix.lower() not in SUPPORTED_SUFFIXES:
            continue
        resolved = path.resolve()
        if resolved.name.startswith("~$"):
            continue
        if resolved not in seen:
            unique.append(resolved)
            seen.add(resolved)
    return unique


def scan_dir(directory: Path) -> list[Path]:
    if not directory.exists():
        return []
    return sorted(path for path in directory.rglob("*") if path.is_file() and path.suffix.lower() in SUPPORTED_SUFFIXES)


def default_output_dir(period: str) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return Path("outputs/enterprise_runs") / f"{period}_{stamp}"


def build_enterprise_result(
    result: dict[str, Any],
    uploads: list[Path],
    output_dir: Path,
    period: str,
    target: str | None,
) -> dict[str, Any]:
    final_report = result.get("final_report")
    final_reports = final_report if isinstance(final_report, list) else ([final_report] if final_report else [])
    blocked_periods = result.get("blocked_periods", [])
    resolved_target = target or result_period_label(result)
    return {
        "status": result.get("status"),
        "message": user_message(result),
        "report_type": period_label(period),
        "target_period": resolved_target,
        "upload_count": len(uploads),
        "output_dir": str(output_dir),
        "final_reports": final_reports,
        "validation_report": result.get("validation_report"),
        "standard_total": result.get("standard_total"),
        "blocked_periods": blocked_periods,
        "blocking_missing": result.get("blocking_missing", []),
        "recommended_missing": result.get("recommended_missing", []),
        "source_files": result.get("source_files", [str(path) for path in uploads]),
        "csv_classifications": result.get("csv_classifications", []),
        "core_dataset_name": "店铺收益核心指标",
        "raw_result": result,
    }


def result_period_label(result: dict[str, Any]) -> str:
    periods = result.get("periods") or []
    if isinstance(periods, list) and len(periods) == 1:
        return str(periods[0])
    if isinstance(periods, list) and periods:
        return "、".join(str(item) for item in periods)
    return "最近可生成周期"


def user_message(result: dict[str, Any]) -> str:
    if result.get("status") == "ok":
        blocked = result.get("blocked_periods") or []
        if blocked:
            return "已生成可用周期的最终报表；部分周期缺少店铺收益核心指标，已列入缺失清单。"
        return "已完成从上传报表到最终报表生成。"
    return result.get("message") or "生成受阻，请查看缺失清单。"


def period_label(period: str) -> str:
    return {"month": "月报", "week": "周报", "half-month": "半月报"}.get(period, period)


def write_enterprise_manifest(
    output_dir: Path,
    uploads: list[Path],
    period: str,
    target: str | None,
    result: dict[str, Any],
) -> None:
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "period": period,
        "target": target,
        "uploads": [str(path) for path in uploads],
        "result": result,
    }
    (output_dir / "enterprise_manifest.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_handoff_note(output_dir: Path, result: dict[str, Any]) -> None:
    lines = [
        "# 万相台企业报表 Agent 交付说明",
        "",
        f"- 状态：{result.get('status')}",
        f"- 报表类型：{result.get('report_type')}",
        f"- 目标周期：{result.get('target_period')}",
        f"- 上传文件数：{result.get('upload_count')}",
        f"- 输出目录：{result.get('output_dir')}",
        "",
        "## 最终报表",
    ]
    final_reports = result.get("final_reports") or []
    if final_reports:
        lines.extend(f"- {item}" for item in final_reports)
    else:
        lines.append("- 未生成；请查看缺失清单。")

    lines.extend(["", "## 审计与校验"])
    for key in ["validation_report", "standard_total"]:
        if result.get(key):
            lines.append(f"- {result[key]}")

    blocked = result.get("blocked_periods") or []
    if blocked:
        lines.extend(["", "## 未生成周期"])
        for item in blocked:
            lines.append(f"- {item.get('period')}：缺少 {'、'.join(item.get('missing', []))}")

    recommended = result.get("recommended_missing") or []
    if recommended:
        lines.extend(["", "## 建议补充"])
        lines.extend(f"- {item}" for item in recommended)

    (output_dir / "交付说明.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
