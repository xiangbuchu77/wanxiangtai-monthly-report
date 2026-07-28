from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from .classifier import ClassifiedReport, ReportType
from .cleaning import clean_frame
from .periods import label_for_date


AD_COST_COLUMNS = ["全站推广花费", "关键词推广花费", "精准人群推广花费", "智能场景花费", "淘宝客佣金"]
STORE_SUM_COLUMNS = [
    "访客数",
    "浏览量",
    "商品访客数",
    "商品浏览量",
    "加购人数",
    "加购件数",
    "支付金额",
    "支付买家数",
    "支付子订单数",
    "支付件数",
    "成功退款金额",
    "老买家支付金额",
    *AD_COST_COLUMNS,
]
PRODUCT_SUM_COLUMNS = [
    "商品访客数",
    "商品浏览量",
    "加购人数",
    "加购件数",
    "商品收藏人数",
    "支付金额",
    "支付件数",
    "支付买家数",
    "成功退款金额",
]
PRODUCT_ROI_SUM_COLUMNS = [
    "支付金额",
    "支付件数",
    "支付人数",
    "成功退款金额",
    "商品访客数",
    "商品浏览量",
    "推广消耗金额",
    "展现量",
    "点击量",
    "推广总加购数",
    "推广引导总成交金额",
    "总成交笔数",
    "关键词推广花费",
    "关键词推广成交金额",
    "人群推广花费",
    "人群推广成交金额",
    "场景推广花费",
    "场景推广成交金额",
    "全站推广花费",
    "全站推广成交金额",
]


@dataclass(frozen=True)
class BuildResult:
    sheets: dict[str, pd.DataFrame]
    audit_log: pd.DataFrame
    quality_log: pd.DataFrame


def safe_divide(numerator: pd.Series | float, denominator: pd.Series | float) -> pd.Series | float:
    if isinstance(denominator, pd.Series):
        return numerator / denominator.where(denominator != 0)
    return numerator / denominator if denominator else 0.0


def build_report(classified_reports: list[ClassifiedReport], period: str = "month") -> BuildResult:
    frames_by_type: dict[ReportType, list[pd.DataFrame]] = {}
    audit_rows: list[dict[str, object]] = []
    quality_rows: list[dict[str, object]] = []

    for report in classified_reports:
        cleaned = add_period_bucket(clean_frame(report.frame), period)
        frames_by_type.setdefault(report.report_type, []).append(cleaned)
        period_min = cleaned["period_start"].min() if "period_start" in cleaned.columns else ""
        period_max = cleaned["period_end"].max() if "period_end" in cleaned.columns else ""
        audit_rows.append(
            {
                "文件名": report.path.name,
                "工作表": report.sheet_name,
                "识别类型": report.report_type.value,
                "行数": len(cleaned),
                "列数": len(cleaned.columns),
                "起始日期": period_min,
                "结束日期": period_max,
                "匹配字段": "、".join(report.matched_columns),
            }
        )

    sheets = {
        "经营总表": build_store_core(frames_by_type.get(ReportType.STORE_CORE_MONTHLY, [])),
        "店铺流量来源": build_store_traffic(frames_by_type),
        "商品经营": build_product_effect(frames_by_type.get(ReportType.PRODUCT_EFFECT_MONTHLY, [])),
        "商品投产比日报": build_product_roi_daily(frames_by_type.get(ReportType.PRODUCT_ROI_DAILY, [])),
        "商品流量来源": build_product_traffic(frames_by_type),
    }

    individual_types = {
        ReportType.STORE_CORE_MONTHLY: "店铺收益核心指标",
        ReportType.PRODUCT_ROI_DAILY: "商品投放效果数据",
        ReportType.PRODUCT_EFFECT_MONTHLY: "商品经营效果数据",
    }
    for report_type, label in individual_types.items():
        count = sum(1 for row in audit_rows if row["识别类型"] == report_type.value)
        quality_rows.append({"检查项": f"数据维度可用:{label}", "状态": "通过" if count else "缺失", "数量": count})
    report_groups = {
        "店铺流量来源数据（新旧口径任一）": [
            ReportType.STORE_TRAFFIC_MONTHLY_NEW,
            ReportType.STORE_TRAFFIC_MONTHLY_OLD,
        ],
        "商品流量来源数据（新旧口径任一）": [
            ReportType.PRODUCT_TRAFFIC_MONTHLY_NEW,
            ReportType.PRODUCT_TRAFFIC_MONTHLY_OLD,
        ],
    }
    for label, report_types in report_groups.items():
        count = sum(1 for row in audit_rows if row["识别类型"] in {report_type.value for report_type in report_types})
        quality_rows.append({"检查项": f"数据维度可用:{label}", "状态": "通过" if count else "缺失", "数量": count})

    for sheet_name, frame in sheets.items():
        quality_rows.append({"检查项": f"输出sheet非空:{sheet_name}", "状态": "通过" if not frame.empty else "空表", "数量": len(frame)})

    unknown_count = sum(1 for row in audit_rows if row["识别类型"] == ReportType.UNKNOWN.value)
    quality_rows.append({"检查项": "未知报表数量", "状态": "通过" if unknown_count == 0 else "存在未知报表", "数量": unknown_count})

    return BuildResult(
        sheets=sheets,
        audit_log=pd.DataFrame(audit_rows),
        quality_log=pd.DataFrame(quality_rows),
    )


def concat(frames: list[pd.DataFrame]) -> pd.DataFrame:
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def existing(columns: list[str], frame: pd.DataFrame) -> list[str]:
    return [column for column in columns if column in frame.columns]


def add_missing_numeric(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    result = frame.copy()
    for column in columns:
        if column not in result.columns:
            result[column] = 0.0
    return result


def add_period_bucket(frame: pd.DataFrame, period: str) -> pd.DataFrame:
    result = frame.copy()
    if "period_start" not in result.columns:
        return result
    starts = pd.to_datetime(result["period_start"], errors="coerce")
    if period in {"month", "half-month", "week"}:
        return add_recent_window_bucket(result, starts, period)
    else:
        result["period_bucket"] = starts.dt.strftime("%Y-%m")
    return result


def add_recent_window_bucket(frame: pd.DataFrame, starts: pd.Series, period: str) -> pd.DataFrame:
    result = frame.copy()
    ends = pd.to_datetime(result.get("period_end", starts), errors="coerce")
    date_ref = ends.where(ends.notna(), starts)
    if date_ref.dropna().empty:
        result["period_bucket"] = starts.dt.strftime("%Y-%m")
        return result
    result["period_bucket"] = label_for_date(date_ref, period, date_ref.max())
    return result[result["period_bucket"].notna()].copy()


def build_store_core(frames: list[pd.DataFrame]) -> pd.DataFrame:
    source = concat(frames)
    if source.empty:
        return pd.DataFrame()

    source = add_missing_numeric(source, STORE_SUM_COLUMNS)
    group_keys = ["店铺名称", "period_bucket"]
    grouped = (
        source.groupby(group_keys, as_index=False, dropna=False)[STORE_SUM_COLUMNS]
        .sum()
        .sort_values(["店铺名称", "period_bucket"])
    )
    if "period_start" in source.columns:
        date_bounds = (
            source.groupby(group_keys, as_index=False, dropna=False)
            .agg(起始日期=("period_start", "min"), 结束日期=("period_end", "max"))
        )
        grouped = grouped.merge(date_bounds, on=group_keys, how="left")
    grouped["推广花费合计"] = grouped[AD_COST_COLUMNS].sum(axis=1)
    grouped["支付转化率_计算"] = safe_divide(grouped["支付买家数"], grouped["访客数"])
    grouped["客单价_计算"] = safe_divide(grouped["支付金额"], grouped["支付买家数"])
    grouped["UV价值_计算"] = safe_divide(grouped["支付金额"], grouped["访客数"])
    grouped["推广费比_计算"] = safe_divide(grouped["推广花费合计"], grouped["支付金额"])
    grouped["推广ROI_计算"] = safe_divide(grouped["支付金额"], grouped["推广花费合计"])

    grouped["支付金额_环比"] = grouped.groupby("店铺名称")["支付金额"].pct_change()
    grouped["访客数_环比"] = grouped.groupby("店铺名称")["访客数"].pct_change()
    grouped["支付买家数_环比"] = grouped.groupby("店铺名称")["支付买家数"].pct_change()
    return grouped.rename(columns={"period_bucket": "月份"})


def build_store_traffic(frames_by_type: dict[ReportType, list[pd.DataFrame]]) -> pd.DataFrame:
    parts: list[pd.DataFrame] = []
    new_frame = concat(frames_by_type.get(ReportType.STORE_TRAFFIC_MONTHLY_NEW, []))
    if not new_frame.empty:
        new_frame = add_missing_numeric(new_frame, ["访客数", "新访客数", "支付买家数", "支付金额", "加购人数"])
        top = new_frame[new_frame["来源层级"].astype(str).isin(["1", "1.0"])].copy()
        if top.empty:
            top = new_frame.copy()
        top["来源口径"] = "新版"
        top["一级来源"] = top["来源名称"]
        top["二级来源"] = ""
        parts.append(top)

    old_frame = concat(frames_by_type.get(ReportType.STORE_TRAFFIC_MONTHLY_OLD, []))
    if not old_frame.empty:
        old_frame = add_missing_numeric(old_frame, ["访客数", "新访客数", "支付买家数", "支付金额", "加购人数"])
        top = old_frame[
            old_frame["来源层级"].astype(str).isin(["1", "1.0"])
            & old_frame["归属原则"].astype(str).eq("每一次访问来源")
        ].copy()
        if top.empty:
            top = old_frame.copy()
        top["来源口径"] = "旧版-每一次访问来源"
        top["一级来源"] = top["一级流量来源"]
        top["二级来源"] = top["二级流量来源"]
        parts.append(top)

    source = concat(parts)
    if source.empty:
        return pd.DataFrame()
    grouped = (
        source.groupby(["来源口径", "店铺名称", "period_bucket", "一级来源", "二级来源"], as_index=False, dropna=False)[
            ["访客数", "新访客数", "支付买家数", "支付金额", "加购人数"]
        ]
        .sum()
        .sort_values(["period_bucket", "来源口径", "支付金额"], ascending=[True, True, False])
    )
    grouped["支付转化率_计算"] = safe_divide(grouped["支付买家数"], grouped["访客数"])
    grouped["UV价值_计算"] = safe_divide(grouped["支付金额"], grouped["访客数"])
    return grouped.rename(columns={"period_bucket": "月份"})


def build_product_effect(frames: list[pd.DataFrame]) -> pd.DataFrame:
    source = concat(frames)
    if source.empty:
        return pd.DataFrame()
    source = add_missing_numeric(source, PRODUCT_SUM_COLUMNS)
    dims = ["店铺名称", "period_bucket", "商品ID", "商品名称", "商品状态", "一级类目名称", "二级类目名称", "叶子类目名称"]
    dims = existing(dims, source)
    grouped = source.groupby(dims, as_index=False, dropna=False)[PRODUCT_SUM_COLUMNS].sum()
    grouped["支付转化率_计算"] = safe_divide(grouped["支付买家数"], grouped["商品访客数"])
    grouped["客单价_计算"] = safe_divide(grouped["支付金额"], grouped["支付买家数"])
    grouped["UV价值_计算"] = safe_divide(grouped["支付金额"], grouped["商品访客数"])
    return grouped.sort_values(["period_bucket", "支付金额"], ascending=[True, False]).rename(columns={"period_bucket": "月份"})


def build_product_roi_daily(frames: list[pd.DataFrame]) -> pd.DataFrame:
    source = concat(frames)
    if source.empty:
        return pd.DataFrame()
    source = add_missing_numeric(source, PRODUCT_ROI_SUM_COLUMNS)
    dims = ["店铺名称", "period_bucket", "商品ID", "商品名称", "品牌名称", "一级类目名称", "二级类目名称", "叶子类目名称", "商品状态"]
    dims = existing(dims, source)
    grouped = source.groupby(dims, as_index=False, dropna=False)[PRODUCT_ROI_SUM_COLUMNS].sum()
    grouped["推广ROI_计算"] = safe_divide(grouped["推广引导总成交金额"], grouped["推广消耗金额"])
    grouped["推广费比_计算"] = safe_divide(grouped["推广消耗金额"], grouped["支付金额"])
    grouped["点击率_计算"] = safe_divide(grouped["点击量"], grouped["展现量"])
    grouped["平均点击花费_计算"] = safe_divide(grouped["推广消耗金额"], grouped["点击量"])
    grouped["支付转化率_计算"] = safe_divide(grouped["支付人数"], grouped["商品访客数"])
    return grouped.sort_values(["period_bucket", "推广消耗金额"], ascending=[True, False]).rename(columns={"period_bucket": "月份"})


def build_product_traffic(frames_by_type: dict[ReportType, list[pd.DataFrame]]) -> pd.DataFrame:
    parts: list[pd.DataFrame] = []
    new_frame = concat(frames_by_type.get(ReportType.PRODUCT_TRAFFIC_MONTHLY_NEW, []))
    if not new_frame.empty:
        new_frame = add_missing_numeric(new_frame, ["访客数", "浏览量", "支付买家数", "支付件数", "支付金额", "加购人数"])
        top = new_frame[new_frame["来源层级"].astype(str).isin(["1", "1.0"])].copy()
        if top.empty:
            top = new_frame.copy()
        top["来源口径"] = "新版"
        top["一级来源"] = top["来源名称"]
        top["二级来源"] = ""
        parts.append(top)

    old_frame = concat(frames_by_type.get(ReportType.PRODUCT_TRAFFIC_MONTHLY_OLD, []))
    if not old_frame.empty:
        old_frame = add_missing_numeric(old_frame, ["访客数", "支付买家数", "支付件数", "支付金额", "加购人数"])
        top = old_frame[
            old_frame["来源层级"].astype(str).isin(["1", "1.0"])
            & old_frame["归属原则"].astype(str).eq("每一次访问来源")
        ].copy()
        if top.empty:
            top = old_frame.copy()
        top["来源口径"] = "旧版-每一次访问来源"
        top["一级来源"] = top["一级流量来源"]
        top["二级来源"] = top["二级流量来源"]
        parts.append(top)

    source = concat(parts)
    if source.empty:
        return pd.DataFrame()
    dims = ["来源口径", "店铺名称", "period_bucket", "商品ID", "商品名称", "一级来源", "二级来源"]
    dims = existing(dims, source)
    metrics = existing(["访客数", "浏览量", "支付买家数", "支付件数", "支付金额", "加购人数"], source)
    grouped = source.groupby(dims, as_index=False, dropna=False)[metrics].sum()
    grouped["支付转化率_计算"] = safe_divide(grouped["支付买家数"], grouped["访客数"])
    grouped["UV价值_计算"] = safe_divide(grouped["支付金额"], grouped["访客数"])
    return grouped.sort_values(["period_bucket", "支付金额"], ascending=[True, False]).rename(columns={"period_bucket": "月份"})
