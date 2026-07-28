from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path

import pandas as pd


class ReportType(str, Enum):
    STORE_CORE_MONTHLY = "store_core_monthly"
    STORE_TRAFFIC_MONTHLY_NEW = "store_traffic_monthly_new"
    STORE_TRAFFIC_MONTHLY_OLD = "store_traffic_monthly_old"
    PRODUCT_ROI_DAILY = "product_roi_daily"
    PRODUCT_TRAFFIC_MONTHLY_NEW = "product_traffic_monthly_new"
    PRODUCT_TRAFFIC_MONTHLY_OLD = "product_traffic_monthly_old"
    PRODUCT_EFFECT_MONTHLY = "product_effect_monthly"
    UNKNOWN = "unknown"


@dataclass(frozen=True)
class ClassifiedReport:
    path: Path
    sheet_name: str
    report_type: ReportType
    frame: pd.DataFrame
    matched_columns: tuple[str, ...]


RULES: tuple[tuple[ReportType, frozenset[str]], ...] = (
    (
        ReportType.PRODUCT_ROI_DAILY,
        frozenset({"统计日期", "店铺名称", "商品ID", "推广消耗金额", "推广ROI", "点击量"}),
    ),
    (
        ReportType.PRODUCT_EFFECT_MONTHLY,
        frozenset({"统计日期", "店铺名称", "商品ID", "商品名称", "商品访客数", "商品详情页跳出率", "支付金额"}),
    ),
    (
        ReportType.PRODUCT_TRAFFIC_MONTHLY_OLD,
        frozenset({"统计日期", "店铺名称", "商品ID", "商品名称", "一级流量来源", "二级流量来源", "归属原则"}),
    ),
    (
        ReportType.PRODUCT_TRAFFIC_MONTHLY_NEW,
        frozenset({"统计日期", "店铺名称", "商品ID", "流量时期", "上级来源名称", "来源名称"}),
    ),
    (
        ReportType.STORE_TRAFFIC_MONTHLY_OLD,
        frozenset({"统计日期", "店铺名称", "一级流量来源", "二级流量来源", "三级流量来源", "归属原则"}),
    ),
    (
        ReportType.STORE_TRAFFIC_MONTHLY_NEW,
        frozenset({"统计日期", "店铺名称", "流量时期", "来源类型", "上级来源名称", "来源名称"}),
    ),
    (
        ReportType.STORE_CORE_MONTHLY,
        frozenset({"统计日期", "店铺名称", "访客数", "支付金额", "支付买家数", "支付转化率", "全站推广花费"}),
    ),
)


def classify(path: Path, sheet_name: str, frame: pd.DataFrame) -> ClassifiedReport:
    columns = set(map(str, frame.columns))
    for report_type, required in RULES:
        if required.issubset(columns):
            return ClassifiedReport(
                path=path,
                sheet_name=sheet_name,
                report_type=report_type,
                frame=frame.copy(),
                matched_columns=tuple(sorted(required)),
            )
    return ClassifiedReport(
        path=path,
        sheet_name=sheet_name,
        report_type=ReportType.UNKNOWN,
        frame=frame.copy(),
        matched_columns=tuple(),
    )
