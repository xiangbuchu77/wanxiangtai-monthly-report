from __future__ import annotations

import argparse
import json
import os
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from .ai_provider import AIConfig, config_from_env, maybe_enhance_analysis
from .classifier import ReportType, classify
from .exporter import write_workbook
from .io import read_excel_reports
from .marketing_scene_importer import convert_marketing_scene, parse_period_start, read_csv
from .metrics import build_report
from .monthly_report_agent import AgentInputs, build_context, display_period_for_title, fill_template, load_ad_metrics, validate_required_data


APP_ROOT = Path(os.environ.get("WXT_APP_ROOT", "."))
DEFAULT_TEMPLATE_CANDIDATES = [
    APP_ROOT / "templates/万相台月报模板.xlsx",
    Path("templates/万相台月报模板.xlsx"),
    Path("outputs/ecommerce-monthly/万相台AI分析月报.xlsx"),
    Path("/Users/gordon/Desktop/万相台月报模板.xlsx"),
]
PERIOD_TEMPLATE_NAMES = {
    "month": "万相台月报模板.xlsx",
    "half-month": "万相台半月报模板.xlsx",
    "week": "万相台周报模板.xlsx",
}
DEFAULT_OUTPUT_DIR = Path("outputs/ecommerce-monthly/agent_run")
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
STORE_REQUIRED_METRICS = CORE_METRIC_LABELS[:3]
AD_ENHANCEMENT_METRICS = CORE_METRIC_LABELS[3:]
STORE_AD_COST_COLUMNS = {"全站推广花费", "关键词推广花费", "精准人群推广花费", "智能场景花费", "淘宝客佣金"}
RECOMMENDED_REPORT_TYPES = {
    ReportType.PRODUCT_ROI_DAILY: "商品投放效果数据：商品访客、推广消耗、点击量、引导成交金额、商品ROI",
}
OPTIONAL_REPORT_GROUPS = {
    "店铺流量来源数据：一级来源、访客数、支付买家数、支付金额": {
        ReportType.STORE_TRAFFIC_MONTHLY_NEW,
        ReportType.STORE_TRAFFIC_MONTHLY_OLD,
    },
    "商品流量来源数据：商品名称、来源、访客数、支付金额、支付转化": {
        ReportType.PRODUCT_TRAFFIC_MONTHLY_NEW,
        ReportType.PRODUCT_TRAFFIC_MONTHLY_OLD,
    },
}


@dataclass(frozen=True)
class CsvClassification:
    path: Path
    kind: str
    start_date: str
    end_date: str
    rows: int
    columns: tuple[str, ...]


@dataclass(frozen=True)
class WorkflowOutputs:
    standard_total: Path
    ad_semifinished: Path
    ai_json: Path
    ai_packet: Path
    final_report: Path
    validation_report: Path


def main() -> None:
    parser = argparse.ArgumentParser(description="淘宝/天猫运营月报自动化agent")
    parser.add_argument("inputs", nargs="*", help="原始Excel/CSV文件路径")
    parser.add_argument("--input-dir", type=Path, help="读取目录下的 .xlsx/.csv 文件")
    parser.add_argument("--template", type=Path, help="月报模板路径")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--month", help="目标周期，例如 2026-05、2026-05上半月、2026-05-04~2026-05-10；不传则自动拆分生成所有周期")
    parser.add_argument("--period", choices=["month", "week", "half-month"], default="month", help="报表周期：month 月报、week 周报、half-month 半月报")
    parser.add_argument("--ai-provider", choices=["", "deepseek"], default=os.environ.get("WXT_AI_PROVIDER", ""), help="可选AI文案增强：deepseek")
    parser.add_argument("--deepseek-api-key", default=os.environ.get("DEEPSEEK_API_KEY", ""), help="DeepSeek API Key；不填则只用本地规则")
    parser.add_argument("--deepseek-model", default=os.environ.get("DEEPSEEK_MODEL", "deepseek-v4-flash"), help="DeepSeek模型名")
    parser.add_argument("--deepseek-base-url", default=os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com"), help="DeepSeek兼容接口地址")
    parser.add_argument("--ai-mode", choices=["off", "fast", "standard", "full"], default=os.environ.get("WXT_AI_MODE", "fast"), help="AI增强模式：off 本地模板、fast 极速、standard 标准、full 精修")
    args = parser.parse_args()

    paths = collect_paths(args.inputs, args.input_dir)
    outputs = build_output_paths(args.output_dir)
    template = args.template or first_existing(DEFAULT_TEMPLATE_CANDIDATES)
    ai_config = AIConfig(args.ai_provider, args.deepseek_api_key, args.deepseek_base_url, args.deepseek_model, args.ai_mode)
    result = run_workflow(paths, template, outputs, args.month, args.period, ai_config=ai_config)
    print(json.dumps(result, ensure_ascii=False, indent=2))


def run_workflow(
    paths: list[Path],
    template: Path,
    outputs: WorkflowOutputs,
    month: str | None = None,
    period: str = "month",
    ai_config: AIConfig | None = None,
) -> dict[str, Any]:
    ai_config = ai_config or config_from_env()
    template = template_for_period(template, period)
    outputs.standard_total.parent.mkdir(parents=True, exist_ok=True)
    excel_paths = [path for path in paths if path.suffix.lower() in {".xlsx", ".xls"}]
    csv_paths = [path for path in paths if path.suffix.lower() == ".csv"]

    loaded = read_excel_reports([path for path in excel_paths if path.suffix.lower() == ".xlsx"])
    classified = [classify(report.path, report.sheet_name, report.raw) for report in loaded]
    csv_classified = [classify_csv(path) for path in csv_paths]

    preflight = preflight_validate(classified, csv_classified)
    write_validation_report(preflight, classified, csv_classified, outputs.validation_report)
    if preflight["blocking_missing"]:
        return {
            "status": "blocked",
            "message": "缺少可生成店铺收益模块的核心指标："
            + "、".join(preflight["blocking_missing"])
            + "。请补充包含这些指标的数据后再生成。",
            "blocking_missing": preflight["blocking_missing"],
            "recommended_missing": preflight["recommended_missing"],
            "validation_report": str(outputs.validation_report),
            "source_files": [str(path) for path in paths],
            "csv_classifications": [csv_classification_dict(item) for item in csv_classified],
        }

    total = build_report(classified, period=period)
    write_workbook(total, outputs.standard_total)

    plan_inputs = [item.path for item in csv_classified if item.kind == "plan_report"]
    scene_inputs = [item.path for item in csv_classified if item.kind == "marketing_scene_report"]
    ad_inputs = plan_inputs or scene_inputs
    product_input = next((item.path for item in csv_classified if item.kind == "product_ad_report"), None)
    ad_result = convert_marketing_scene(ad_inputs, product_input, period=period)
    with pd.ExcelWriter(outputs.ad_semifinished, engine="openpyxl") as writer:
        ad_result["monthly_summary"].to_excel(writer, sheet_name="monthly_summary", index=False)
        ad_result["plan_detail"].to_excel(writer, sheet_name="plan_detail", index=False)
        ad_result["product_ad"].to_excel(writer, sheet_name="product_ad", index=False)

    ad_metrics = load_ad_metrics(outputs.ad_semifinished)
    sheets = pd.read_excel(outputs.standard_total, sheet_name=None)
    target_months = [resolve_requested_period(outputs.standard_total, month)] if month else infer_latest_period(outputs.standard_total)
    final_reports: list[str] = []
    ai_json_paths: list[str] = []
    packet_paths: list[str] = []
    blocked: list[dict[str, Any]] = []
    for target_month in target_months:
        validation = validate_required_data(sheets, ad_metrics, target_month)
        context = build_context(sheets, ad_metrics, target_month, validation)
        context["period_type"] = period
        if validation["blocking_missing"]:
            blocked.append({"period": target_month, "missing": validation["blocking_missing"], "warnings": validation["warnings"]})
            continue

        ai_json = build_analysis_json(context, ai_config)
        ai_json_path = period_path(outputs.ai_json, target_month, len(target_months) > 1)
        ai_json_path.write_text(json.dumps(ai_json, ensure_ascii=False, indent=2), encoding="utf-8")
        ai_json_paths.append(str(ai_json_path))

        final_report = final_report_path(outputs.final_report.parent, context, target_month, period)
        fill_template(template, final_report, context, ai_json)
        final_reports.append(str(final_report))
        packet = {
            "role": "资深 Python 数据工程师、BI分析师和电商数据产品经理",
            "period": target_month,
            "period_type": period,
            "data": context,
            "analysis_json": ai_json,
        }
        packet_path = period_path(outputs.ai_packet, target_month, len(target_months) > 1)
        packet_path.write_text(json.dumps(packet, ensure_ascii=False, indent=2), encoding="utf-8")
        packet_paths.append(str(packet_path))

    if outputs.ad_semifinished.exists():
        outputs.ad_semifinished.unlink()
    if not final_reports:
        blocked_missing = sorted({missing for item in blocked for missing in item.get("missing", [])})
        return {
            "status": "blocked",
            "message": "已识别数据，但所有周期都缺少店铺收益核心指标："
            + "、".join(blocked_missing or STORE_REQUIRED_METRICS)
            + "，暂时无法生成最终xlsx。",
            "blocking_missing": blocked_missing,
            "blocked_periods": blocked,
            "standard_total": str(outputs.standard_total),
            "ad_semifinished_deleted": str(outputs.ad_semifinished),
            "ai_json": ai_json_paths,
            "validation_report": str(outputs.validation_report),
            "source_files": [str(path) for path in paths],
            "csv_classifications": [csv_classification_dict(item) for item in csv_classified],
        }
    return {
        "status": "ok",
        "period": period,
        "periods": target_months,
        "defaulted_to_latest_period": month is None,
        "standard_total": str(outputs.standard_total),
        "ad_semifinished_deleted": str(outputs.ad_semifinished),
        "ai_json": ai_json_paths[0] if len(ai_json_paths) == 1 else ai_json_paths,
        "ai_packet": packet_paths[0] if len(packet_paths) == 1 else packet_paths,
        "final_report": final_reports[0] if len(final_reports) == 1 else final_reports,
        "blocked_periods": blocked,
        "validation_report": str(outputs.validation_report),
        "recommended_missing": preflight["recommended_missing"],
        "source_files": [str(path) for path in paths],
        "csv_classifications": [csv_classification_dict(item) for item in csv_classified],
    }


def collect_paths(inputs: list[str], input_dir: Path | None) -> list[Path]:
    paths: list[Path] = []
    if input_dir:
        paths.extend(sorted(path for path in input_dir.iterdir() if path.suffix.lower() in {".xlsx", ".xls", ".csv", ".zip"}))
    paths.extend(Path(item) for item in inputs)
    paths = expand_zip_inputs(paths, (input_dir or Path("outputs/ecommerce-monthly/agent_run")) / "_unzipped")
    unique: list[Path] = []
    seen: set[Path] = set()
    for path in paths:
        resolved = path.expanduser()
        if resolved not in seen:
            unique.append(resolved)
            seen.add(resolved)
    return unique


def expand_zip_inputs(paths: list[Path], extract_root: Path) -> list[Path]:
    expanded: list[Path] = []
    for path in paths:
        source = path.expanduser()
        if source.suffix.lower() == ".zip" and source.exists() and zipfile.is_zipfile(source):
            target_dir = extract_root / source.stem
            target_dir.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(source) as archive:
                for info in archive.infolist():
                    if info.is_dir():
                        continue
                    inner = Path(info.filename)
                    if inner.name.startswith(".") or "__MACOSX" in inner.parts:
                        continue
                    if inner.suffix.lower() not in {".xlsx", ".xls", ".csv"}:
                        continue
                    filename = sanitize_filename(inner.name)
                    destination = unique_path(target_dir / filename)
                    with archive.open(info) as src, destination.open("wb") as out:
                        shutil.copyfileobj(src, out)
                    expanded.append(destination.resolve())
        else:
            expanded.append(source)
    return expanded


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    for index in range(1, 1000):
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"无法保存解压文件: {path.name}")


def build_output_paths(output_dir: Path) -> WorkflowOutputs:
    return WorkflowOutputs(
        standard_total=output_dir / "01_标准经营总表.xlsx",
        ad_semifinished=output_dir / "02_月报数据半成品_广告汇总.xlsx",
        ai_json=output_dir / "03_AI分析文案.json",
        ai_packet=output_dir / "04_AI任务包.json",
        final_report=output_dir / "05_万相台月报_最终版.xlsx",
        validation_report=output_dir / "00_数据缺失与识别报告.xlsx",
    )


def first_existing(paths: list[Path]) -> Path:
    for path in paths:
        if path.exists():
            return path
    raise FileNotFoundError("未找到月报模板，请用 --template 指定模板路径。")


def template_for_period(template: Path, period: str) -> Path:
    expected_name = PERIOD_TEMPLATE_NAMES.get(period, PERIOD_TEMPLATE_NAMES["month"])
    candidates = [
        APP_ROOT / "templates" / expected_name,
        template.parent / expected_name,
        Path("templates") / expected_name,
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return template


def classify_csv(path: Path) -> CsvClassification:
    frame = read_csv(path)
    columns = tuple(frame.columns)
    dates = parse_period_start(frame["日期"]) if "日期" in frame.columns else pd.Series(dtype="datetime64[ns]")
    kind = "unknown_csv"
    column_set = set(columns)
    if {"计划名字", "花费", "投入产出比", "点击量", "展现量"}.issubset(column_set):
        kind = "plan_report"
    elif {"主体ID", "主体名称", "花费", "投入产出比"}.issubset(column_set):
        kind = "product_ad_report"
    elif {"词名字/词包名字", "花费", "投入产出比"}.issubset(column_set):
        kind = "keyword_report"
    elif {"场景名字", "花费", "投入产出比", "点击量", "展现量"}.issubset(column_set):
        kind = "marketing_scene_report"
    return CsvClassification(
        path=path,
        kind=kind,
        start_date="" if dates.empty else str(dates.min().date()),
        end_date="" if dates.empty else str(dates.max().date()),
        rows=len(frame),
        columns=columns,
    )


def csv_classification_dict(item: CsvClassification) -> dict[str, Any]:
    return {
        "path": str(item.path),
        "kind": item.kind,
        "start_date": item.start_date,
        "end_date": item.end_date,
        "rows": item.rows,
        "columns": list(item.columns),
    }


def preflight_validate(classified: list[Any], csv_classified: list[CsvClassification]) -> dict[str, Any]:
    present_types = {item.report_type for item in classified}
    available_metrics = infer_available_core_metrics(classified, csv_classified)
    blocking_missing: list[str] = []
    recommended_missing: list[str] = []
    for label in STORE_REQUIRED_METRICS:
        if label not in available_metrics:
            blocking_missing.append(label)
    for label in AD_ENHANCEMENT_METRICS:
        if label not in available_metrics:
            recommended_missing.append(f"缺少指标：{label}（缺少时广告分析生成简版）")
    for report_type, label in RECOMMENDED_REPORT_TYPES.items():
        if report_type not in present_types:
            recommended_missing.append(label)
    for label, options in OPTIONAL_REPORT_GROUPS.items():
        if present_types.isdisjoint(options):
            recommended_missing.append(label)

    csv_kinds = {item.kind for item in csv_classified}
    if "plan_report" not in csv_kinds:
        recommended_missing.append("推广计划明细数据：计划名称、计划类型、花费、展现量、点击量、ROI（缺少时操作规划会简化）")
    if "product_ad_report" not in csv_kinds:
        recommended_missing.append("商品广告数据：商品ID、商品名称、广告花费、广告ROI（缺少时主要商品广告建议会简化）")
    if "keyword_report" not in csv_kinds:
        recommended_missing.append("关键词投放数据：关键词/词包、花费、点击量、成交或ROI（缺少时关键词层建议会简化）")

    return {
        "blocking_missing": sorted(set(blocking_missing)),
        "recommended_missing": sorted(set(recommended_missing)),
    }


def infer_available_core_metrics(classified: list[Any], csv_classified: list[CsvClassification]) -> set[str]:
    available: set[str] = set()
    for item in classified:
        columns = set(map(str, item.frame.columns))
        if item.report_type == ReportType.STORE_CORE_MONTHLY:
            if "访客数" in columns:
                available.add("店铺总访客数")
            if "支付金额" in columns:
                available.add("店铺总成交金额（元）")
            if "支付转化率" in columns or {"支付买家数", "访客数"}.issubset(columns):
                available.add("店铺支付转化率")
            if columns.intersection(STORE_AD_COST_COLUMNS):
                available.add("广告总花费（元）")
        if item.report_type == ReportType.PRODUCT_ROI_DAILY:
            if "点击量" in columns:
                available.add("总点击量")
            if {"推广消耗金额", "点击量"}.issubset(columns):
                available.add("平均点击成本（元）")
            if "推广消耗金额" in columns:
                available.add("广告总花费（元）")
            if "推广引导总成交金额" in columns:
                available.add("广告带来的成交金额（元）")
            if "推广ROI" in columns or {"推广引导总成交金额", "推广消耗金额"}.issubset(columns):
                available.add("广告投入产出比（ROI）")

    for item in csv_classified:
        columns = set(map(str, item.columns))
        if item.kind in {"marketing_scene_report", "plan_report"}:
            if "花费" in columns:
                available.add("广告总花费（元）")
            if "点击量" in columns:
                available.add("总点击量")
            if {"花费", "点击量"}.issubset(columns):
                available.add("平均点击成本（元）")
            if {"花费", "投入产出比"}.issubset(columns):
                available.add("广告带来的成交金额（元）")
                available.add("广告投入产出比（ROI）")
        if item.kind == "product_ad_report":
            if "花费" in columns:
                available.add("广告总花费（元）")
            if "投入产出比" in columns:
                available.add("广告投入产出比（ROI）")
    return available


def write_validation_report(
    preflight: dict[str, Any],
    classified: list[Any],
    csv_classified: list[CsvClassification],
    output: Path,
) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    excel_rows = [
        {
            "文件名": item.path.name,
            "工作表": item.sheet_name,
            "识别类型": item.report_type.value,
            "行数": len(item.frame),
            "列数": len(item.frame.columns),
            "匹配字段": "、".join(item.matched_columns),
        }
        for item in classified
    ]
    csv_rows = [
        {
            "文件名": item.path.name,
            "识别类型": item.kind,
            "起始日期": item.start_date,
            "结束日期": item.end_date,
            "行数": item.rows,
            "列数": len(item.columns),
        }
        for item in csv_classified
    ]
    missing_rows = [missing_row("阻断缺失", item) for item in preflight["blocking_missing"]]
    missing_rows.extend(missing_row("建议补充", item) for item in preflight["recommended_missing"])
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(excel_rows).to_excel(writer, sheet_name="Excel识别", index=False)
        pd.DataFrame(csv_rows).to_excel(writer, sheet_name="CSV识别", index=False)
        pd.DataFrame(missing_rows).to_excel(writer, sheet_name="缺失清单", index=False)


def missing_row(level: str, item: str) -> dict[str, str]:
    return {
        "类型": level,
        "缺少数据内容": item,
        "影响模块": missing_impact(item),
    }


def missing_impact(item: str) -> str:
    if "商品" in item:
        return "主要商品数据；商品操作建议"
    if "计划" in item:
        return "推广计划明细；下月重点操作规划"
    if "关键词" in item:
        return "关键词诊断；投放优化建议"
    if "流量来源" in item:
        return "流量结构分析；来源承接建议"
    if "店铺收益" in item or "访客数" in item or "支付金额" in item or "支付转化" in item:
        return "一、店铺收益；核心指标总览；无法生成完整报表"
    if "广告" in item or "ROI" in item or "点击量" in item or "展现量" in item:
        return "推广进展分析；预算分配；运营总结"
    return "分析完整度"


def infer_month(total_path: Path) -> str:
    core = pd.read_excel(total_path, sheet_name="经营总表")
    return str(core["月份"].dropna().max())


def infer_periods(total_path: Path) -> list[str]:
    core = pd.read_excel(total_path, sheet_name="经营总表")
    if core.empty or "月份" not in core.columns:
        return []
    return [str(item) for item in sorted(core["月份"].dropna().astype(str).unique())]


def infer_latest_period(total_path: Path) -> list[str]:
    periods = infer_periods(total_path)
    if not periods:
        return []
    return [periods[-1]]


def resolve_requested_period(total_path: Path, requested: str | None) -> str:
    if not requested:
        periods = infer_latest_period(total_path)
        return periods[0] if periods else ""
    text = str(requested).strip()
    periods = infer_periods(total_path)
    if text in periods:
        return text
    if len(text) == 7 and text[4] == "-":
        matches = [period for period in periods if period_matches_month(period, text)]
        if matches:
            return matches[-1]
    return text


def period_matches_month(period_value: str, month: str) -> bool:
    text = str(period_value)
    if text.startswith(month):
        return True
    if "~" in text:
        start_text, end_text = text.split("~", 1)
        start = pd.to_datetime(start_text, errors="coerce")
        end = pd.to_datetime(end_text, errors="coerce")
        return (
            (not pd.isna(start) and start.strftime("%Y-%m") == month)
            or (not pd.isna(end) and end.strftime("%Y-%m") == month)
        )
    return False


def period_path(path: Path, period_value: str, split: bool) -> Path:
    if not split:
        return path
    safe_period = period_value.replace("/", "-").replace("~", "_")
    return path.with_name(f"{path.stem}_{safe_period}{path.suffix}")


def final_report_path(output_dir: Path, context: dict[str, Any], period_value: str, period: str) -> Path:
    store_name = sanitize_filename(str(context.get("core", {}).get("店铺名称") or "店铺"))
    label, _is_partial = display_period_for_title(period_value, period, context.get("core", {}))
    suffix = {"month": "月报", "week": "周报", "half-month": "半月报"}.get(period, "报表")
    return output_dir / f"{store_name}{label}{suffix}.xlsx"


def display_period(period_value: str, period: str) -> str:
    text = str(period_value)
    if period == "month" and len(text) == 7 and text[4] == "-":
        year, month = text.split("-", 1)
        return f"{year}年{month}月"
    if period == "half-month" and len(text) >= 10 and text[4] == "-":
        year = text[:4]
        month = text[5:7]
        half = text[7:]
        return f"{year}年{month}月{half}"
    if period == "week":
        return text.replace("~", "至")
    return text


def sanitize_filename(value: str) -> str:
    cleaned = "".join("_" if char in r'\/:*?"<>|' else char for char in value).strip()
    return cleaned or "店铺"


def build_analysis_json(context: dict[str, Any], ai_config: AIConfig | None = None) -> dict[str, Any]:
    core = context["core"]
    previous = context["previous_core"]
    monthly_ad = context["monthly_ad"]
    previous_ad = context["previous_monthly_ad"]
    top_products = context["top_products"]
    product_ad = {normalize_id(row.get("商品ID")): row for row in context["product_ad"]}
    plan_rows = context["plan_ad"]
    best_plans = sorted(plan_rows, key=lambda row: float(row.get("ROI") or 0), reverse=True)[:2]
    high_spend_plans = sorted(plan_rows, key=lambda row: float(row.get("花费") or 0), reverse=True)[:3]
    visible_plan_rows = plan_rows[:3]
    visible_best_plans = sorted(visible_plan_rows, key=lambda row: float(row.get("ROI") or 0), reverse=True)[:2]
    visible_high_spend_plans = sorted(visible_plan_rows, key=lambda row: float(row.get("花费") or 0), reverse=True)[:3]
    high_spend_product = max(
        top_products,
        key=lambda row: float(row.get("广告花费") or product_ad.get(normalize_id(row.get("商品ID")), {}).get("广告花费") or 0),
        default=top_products[0] if top_products else {},
    )
    best_product = max(
        top_products,
        key=lambda row: float(row.get("广告ROI") or product_ad.get(normalize_id(row.get("商品ID")), {}).get("广告ROI") or 0),
        default=top_products[0] if top_products else {},
    )

    product_actions = []
    for product in top_products[:4]:
        product_id = normalize_id(product.get("商品ID"))
        ad = product_ad.get(product_id, {})
        product_actions.append(
            {
                "商品名称": product.get("商品名称", ""),
                "操作建议": product_action_text(product, ad, monthly_ad),
            }
        )

    draft = {
        "product_actions": product_actions,
        "goal_paths": goal_paths(core, previous, monthly_ad, previous_ad, best_plans),
        "budget_usage": budget_usage(
            visible_plan_rows,
            visible_best_plans,
            visible_high_spend_plans,
            best_product,
            high_spend_product,
            monthly_ad,
        ),
        "weekly_plan": weekly_plan(best_plans, high_spend_plans, best_product, high_spend_product, monthly_ad),
        "summary": summary_text(
            core,
            previous,
            monthly_ad,
            previous_ad,
            best_plans,
            high_spend_plans,
            best_product,
            high_spend_product,
        ),
    }
    return maybe_enhance_analysis(draft, context, ai_config)


def product_action_text(product: dict[str, Any], ad: dict[str, Any], monthly_ad: dict[str, Any]) -> str:
    spend = num(ad.get("广告花费") or product.get("广告花费"))
    roi = num(ad.get("广告ROI") or product.get("广告ROI"))
    total_roi = num(monthly_ad.get("总ROI"))
    if spend and spend < 200:
        return f"广告花费{money(spend)}样本偏小，先观察3-7天；ROI稳定高于账户均值后再放大预算。"
    if spend and roi >= total_roi:
        return f"广告花费{money(spend)}、ROI {one(roi)}高于账户{one(total_roi)}，保留预算；若量级不足，出价最多上调10%。"
    if spend:
        return f"广告花费{money(spend)}、ROI {one(roi)}低于账户{one(total_roi)}，先回收低效预算10%，保留观察位。"
    return "缺少商品广告花费和ROI，暂不主动加预算；补充商品广告报表后再判断。"


def goal_paths(
    core: dict[str, Any],
    previous: dict[str, Any],
    monthly_ad: dict[str, Any],
    previous_ad: dict[str, Any],
    best_plans: list[dict[str, Any]],
) -> dict[str, str]:
    plan_names = "、".join(str(plan.get("推广计划名称")) for plan in best_plans if plan.get("推广计划名称"))
    return {
        "月销售额": f"以本月支付{money(core.get('支付金额'))}为基准设+12%目标，增量优先来自{plan_names or '高ROI计划'}和Top商品放量。",
        "广告ROI": f"本月广告ROI {one(monthly_ad.get('总ROI'))}，下月目标按高于当前5%设定；预算优先保留{plan_names or '高ROI计划'}。",
        "访客数": f"本月访客{whole(core.get('访客数'))}，下月目标按+8%设定，新增访客主要来自全站推广和关键词推广。",
        "转化率": f"本月支付转化率{percent(core.get('支付转化率_计算'))}，重点处理高访客低转化商品页面和客服承接。"
    }


def budget_usage(
    plan_rows: list[dict[str, Any]],
    best_plans: list[dict[str, Any]],
    high_spend_plans: list[dict[str, Any]],
    best_product: dict[str, Any],
    high_spend_product: dict[str, Any],
    monthly_ad: dict[str, Any],
) -> list[dict[str, str]]:
    high_spend_plan = high_spend_plans[0] if high_spend_plans else {}
    best_product_name = str(best_product.get("商品名称", "高ROI商品"))[:24]
    high_spend_name = str(high_spend_product.get("商品名称", "高花费商品"))[:24]
    account_roi = num(monthly_ad.get("总ROI"))
    channel_ratios = budget_ratios_from_plans(plan_rows, account_roi)
    items: list[dict[str, str]] = []
    for channel, ratio in channel_ratios.items():
        names = channel_plan_names(plan_rows, channel)
        if channel == "关键词推广":
            usage = f"{names or '关键词计划'}优先保留成交稳定、ROI高于账户均值{one(account_roi)}的预算；30次点击无成交词先降价5%-10%。"
        elif channel == "全站推广":
            usage = f"优先承接{names or best_product_name}；若ROI连续3天高于账户均值，预算可增加20%，反之回收10%。"
        elif channel == "人群推广":
            usage = f"{names or high_spend_name}围绕高意向人群拆分测试，ROI高于账户{one(account_roi)}时保留预算。"
        else:
            usage = f"{names or channel}按成交效率分配预算，ROI低于账户均值时先回收低效消耗。"
        items.append({"推广渠道": channel, "占比": ratio, "用途说明": usage})
    return items


def budget_ratios_from_plans(plan_rows: list[dict[str, Any]], account_roi: float) -> dict[str, float]:
    channels = active_budget_channels(plan_rows)
    if not channels:
        return {}
    stats = {channel: {"spend": 0.0, "weighted_roi": 0.0} for channel in channels}
    for plan in plan_rows:
        channel = budget_channel(plan)
        if channel not in stats:
            continue
        spend = num(plan.get("花费"))
        roi = num(plan.get("ROI"))
        stats[channel]["spend"] += spend
        stats[channel]["weighted_roi"] += spend * roi
    total_spend = sum(item["spend"] for item in stats.values())
    if total_spend <= 0:
        share = 1 / len(channels)
        return {channel: share for channel in channels}
    scores: dict[str, float] = {}
    for channel, item in stats.items():
        spend = item["spend"]
        spend_share = spend / total_spend if total_spend else 0
        roi = item["weighted_roi"] / spend if spend else 0
        if spend <= 0:
            factor = 0.25
        elif account_roi > 0 and roi >= account_roi * 1.15:
            factor = 1.25
        elif account_roi > 0 and roi < account_roi * 0.85:
            factor = 0.75
        else:
            factor = 1.0
        scores[channel] = max(spend_share * factor, 0.03 if spend > 0 else 0.0)
    total_score = sum(scores.values())
    if total_score <= 0:
        share = 1 / len(channels)
        return {channel: share for channel in channels}
    ratios = {channel: scores[channel] / total_score for channel in channels}
    return rebalance_ratios(ratios)


def rebalance_ratios(ratios: dict[str, float]) -> dict[str, float]:
    floors = {"关键词推广": 0.20, "全站推广": 0.15, "人群推广": 0.05}
    caps = {"关键词推广": 0.70, "全站推广": 0.65, "人群推广": 0.30}
    adjusted = {
        channel: min(max(value, floors.get(channel, 0.03)), caps.get(channel, 0.80))
        for channel, value in ratios.items()
    }
    total = sum(adjusted.values())
    if total <= 0:
        return {}
    return {channel: adjusted[channel] / total for channel in adjusted}


def active_budget_channels(plan_rows: list[dict[str, Any]]) -> list[str]:
    channels: list[str] = []
    for plan in plan_rows:
        if num(plan.get("花费")) <= 0:
            continue
        channel = budget_channel(plan)
        if channel not in channels:
            channels.append(channel)
    return channels


def budget_channel(plan: dict[str, Any]) -> str:
    text = f"{plan.get('推广计划名称', '')} {plan.get('计划类型', '')}"
    if "人群" in text:
        return "人群推广"
    if "全站" in text:
        return "全站推广"
    return "关键词推广"


def channel_plan_names(plan_rows: list[dict[str, Any]], channel: str) -> str:
    rows = [row for row in plan_rows if budget_channel(row) == channel]
    rows = sorted(rows, key=lambda row: num(row.get("ROI")), reverse=True)[:2]
    return "、".join(str(row.get("推广计划名称")) for row in rows if row.get("推广计划名称"))


def weekly_plan(
    best_plans: list[dict[str, Any]],
    high_spend_plans: list[dict[str, Any]],
    best_product: dict[str, Any],
    high_spend_product: dict[str, Any],
    monthly_ad: dict[str, Any],
) -> list[dict[str, str]]:
    best_plan = str(best_plans[0].get("推广计划名称", "最高ROI计划")) if best_plans else "最高ROI计划"
    high_spend_plan = str(high_spend_plans[0].get("推广计划名称", "最高花费计划")) if high_spend_plans else "最高花费计划"
    high_spend_roi = num(high_spend_plans[0].get("ROI")) if high_spend_plans else 0.0
    best_product_name = str(best_product.get("商品名称", "最高ROI商品"))[:24]
    high_spend_name = str(high_spend_product.get("商品名称", "高花费商品"))[:24]
    account_roi_value = num(monthly_ad.get("总ROI"))
    account_roi = one(account_roi_value)
    total_spend = num(monthly_ad.get("总推广花费"))
    total_clicks = num(monthly_ad.get("点击量"))
    cpc = total_spend / total_clicks if total_clicks else 0.0
    product_roi = num(high_spend_product.get("广告ROI") or high_spend_product.get("ROI"))

    actions: list[dict[str, str]] = []
    if high_spend_roi and account_roi_value and high_spend_roi < account_roi_value:
        actions.append({
            "操作事项": "低ROI计划止损",
            "具体内容": f"优先检查{high_spend_plan}，ROI {one(high_spend_roi)}低于账户{account_roi}；先回收10%预算，拆分关键词、人群和时段看低效来源。",
            "预期效果": "减少低产出消耗，把预算留给更确定的成交来源。",
        })
    else:
        actions.append({
            "操作事项": "高ROI计划扩量",
            "具体内容": f"把{best_plan}作为优先加权对象，单次加预算不超过20%，同时观察点击、收藏加购和成交是否同步增长。",
            "预期效果": "在控制风险的前提下扩大高产出计划流量。",
        })

    if product_roi and account_roi_value and product_roi < account_roi_value:
        actions.append({
            "操作事项": "高花费商品止损",
            "具体内容": f"复盘{high_spend_name}的广告花费和ROI，若连续3天低于账户均值，降低出价或暂停低成交词。",
            "预期效果": "避免预算继续集中在低回报商品上。",
        })
    else:
        actions.append({
            "操作事项": "重点商品放量",
            "具体内容": f"围绕{best_product_name}增加优质计划承接，主图、详情首屏和优惠表达同步补强。",
            "预期效果": "把高ROI商品的点击优势转化为更多成交。",
        })

    if cpc >= 1.5:
        actions.append({
            "操作事项": "点击成本压降",
            "具体内容": f"当前CPC约{one(cpc)}元，优先下调高点击低成交词出价，保留成交词和高ROI人群。",
            "预期效果": "降低无效点击消耗，提高同等预算下的成交机会。",
        })
    else:
        actions.append({
            "操作事项": "优质流量加权",
            "具体内容": f"当前CPC约{one(cpc)}元，可围绕{best_plan}扩大稳定流量，但需每日盯ROI和成交金额。",
            "预期效果": "利用较低点击成本获取更多有效访客。",
        })

    actions.append({
        "操作事项": "预算结构微调",
        "具体内容": "根据3.1计划表现，把低ROI预算小步迁移到成交稳定计划；每次调整后至少观察2到3天。",
        "预期效果": "让预算分配跟随真实成交效率变化。",
    })
    actions.append({
        "操作事项": "详情页成交承接",
        "具体内容": f"优先优化{best_product_name}详情页首屏、规格卖点、评价背书和客服话术，减少点击后流失。",
        "预期效果": "提升高意向流量的支付转化。",
    })

    return [
        {"时间节点": f"第{index}周", **item}
        for index, item in enumerate(actions[:4], start=1)
    ]


def summary_text(
    core: dict[str, Any],
    previous: dict[str, Any],
    monthly_ad: dict[str, Any],
    previous_ad: dict[str, Any],
    best_plans: list[dict[str, Any]],
    high_spend_plans: list[dict[str, Any]],
    best_product: dict[str, Any],
    high_spend_product: dict[str, Any],
) -> dict[str, list[str]]:
    best_plan = best_plans[0] if best_plans else {}
    high_spend_plan = high_spend_plans[0] if high_spend_plans else {}
    visitors = num(core.get("访客数"))
    prev_visitors = num(previous.get("访客数"))
    sales = num(core.get("支付金额"))
    prev_sales = num(previous.get("支付金额"))
    conversion = num(core.get("支付转化率_计算"))
    prev_conversion = num(previous.get("支付转化率_计算"))
    ad_spend = num(monthly_ad.get("总推广花费"))
    prev_ad_spend = num(previous_ad.get("总推广花费"))
    ad_sales = num(monthly_ad.get("广告引导成交金额"))
    prev_ad_sales = num(previous_ad.get("广告引导成交金额"))
    account_roi = num(monthly_ad.get("总ROI"))
    prev_roi = num(previous_ad.get("总ROI"))
    clicks = num(monthly_ad.get("点击量"))
    prev_clicks = num(previous_ad.get("点击量"))
    cpc = safe_div(ad_spend, clicks)
    prev_cpc = safe_div(prev_ad_spend, prev_clicks)
    best_plan_roi = num(best_plan.get("ROI"))
    high_spend_roi = num(high_spend_plan.get("ROI"))
    high_spend_risk = (
        f"预算预警：{high_spend_plan.get('推广计划名称', '最高花费计划')}花费{money(high_spend_plan.get('花费'))}、ROI {one(high_spend_roi)}高于账户{one(account_roi)}，可保留但需看转化是否连续稳定。"
        if high_spend_roi >= account_roi
        else f"预算风险：{high_spend_plan.get('推广计划名称', '最高花费计划')}花费{money(high_spend_plan.get('花费'))}、ROI {one(high_spend_roi)}低于账户{one(account_roi)}，先降10%预算。"
    )
    return {
        "本月亮点": [
            f"访客{whole(visitors)}环比{pct(visitors, prev_visitors)}，说明流量入口被放大；但成交额仅{money(sales)}、环比{pct(sales, prev_sales)}，下月不能只继续放量，要把新增流量导向高转化商品。",
            f"广告引导成交{money(ad_sales)}环比{pct(ad_sales, prev_ad_sales)}，高于广告花费环比{pct(ad_spend, prev_ad_spend)}，说明投放拉动成交有效；保留高ROI计划并做预算迁移。",
            f"{best_plan.get('推广计划名称', '最高ROI计划')} ROI {one(best_plan_roi)}，高于账户均值{one(account_roi)}，属于可扩量计划；下月可先测试加20%预算。",
            f"CPC从{one(prev_cpc)}元降到{one(cpc)}元，点击量环比{pct(clicks, prev_clicks)}，说明流量采购成本下降；可继续保留低CPC且有成交的计划。",
            f"账户ROI从{one(prev_roi)}到{one(account_roi)}，环比{pct(account_roi, prev_roi)}，投放效率有改善；预算应优先给ROI高于账户均值的计划。"
        ],
        "存在问题": [
            f"异常预警：访客环比{pct(visitors, prev_visitors)}，但支付转化率从{percent(prev_conversion)}降到{percent(conversion)}，新增流量质量或商品承接存在问题，先查定向、人群和详情页。",
            high_spend_risk,
            f"成交承接风险：支付金额环比{pct(sales, prev_sales)}低于访客环比{pct(visitors, prev_visitors)}，说明流量没有充分转成成交；下月优先提升转化率。",
            f"{str(high_spend_product.get('商品名称', '高花费商品'))[:24]}属于重点消耗商品，需逐日看ROI、CPC和成交，连续3天低于账户均值就回收预算。",
            "数据风险：若缺少搜索词、人群、地域拆分表，只能判断账户层异常，无法定位具体低效词和低效人群。"
        ],
        "下月重点关注": [
            f"预算优化：{best_plan.get('推广计划名称', '高ROI计划')} ROI {one(best_plan_roi)}高于账户{one(account_roi)}，先加20%预算测试；低于账户均值的高花费计划同步回收10%。",
            f"增长动作：访客增长{pct(visitors, prev_visitors)}但转化率下降{abs_pct_change(conversion, prev_conversion)}，下月重点提升转化率，而不是继续无差别扩大流量。",
            f"商品动作：优先优化{str(best_product.get('商品名称', '高ROI商品'))[:24]}详情页首屏、评价背书、优惠表达和客服响应，承接新增点击。",
            "风控动作：建立高点击低成交清单，30次点击无成交先降价5%-10%，100次点击无成交加入否定或暂停。",
            "复盘动作：每周输出计划、商品、关键词、人群四层表，把ROI高于账户均值的预算加20%，低ROI高花费对象降10%。"
        ],
    }


def num(value: Any) -> float:
    if value is None or pd.isna(value):
        return 0.0
    return float(value)


def pct(current: float, previous: float) -> str:
    if not previous:
        return "无上月基数"
    return f"{(current / previous - 1):+.1%}"


def abs_pct_change(current: float, previous: float) -> str:
    if not previous:
        return "无上月基数"
    return f"{abs(current / previous - 1):.1%}"


def safe_div(numerator: float, denominator: float) -> float:
    return numerator / denominator if denominator else 0.0


def one(value: Any) -> str:
    return f"{num(value):.1f}"


def whole(value: Any) -> str:
    return f"{num(value):,.0f}"


def money(value: Any) -> str:
    return f"{num(value):,.0f}元"


def percent(value: Any) -> str:
    return f"{num(value):.1%}"


def normalize_id(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().replace(",", "")
    if text.endswith(".0"):
        text = text[:-2]
    return text


if __name__ == "__main__":
    main()
