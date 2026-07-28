from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd


NUMERIC_COLUMNS = {
    "访客数",
    "直播间访客数",
    "短视频访客数",
    "图文访客数",
    "店铺页访客数",
    "3D应用访客数",
    "无线端访客数",
    "PC端访客数",
    "浏览量",
    "商品访客数",
    "商品浏览量",
    "平均停留时长",
    "平均停留时长（秒）",
    "商品详情页跳出率",
    "跳失率",
    "新访客数",
    "支付金额",
    "支付买家数",
    "支付人数",
    "支付件数",
    "支付笔数",
    "支付子订单数",
    "支付父订单数",
    "支付转化率",
    "客单价",
    "UV价值",
    "成功退款金额",
    "下单金额",
    "下单买家数",
    "下单件数",
    "下单转化率",
    "加购人数",
    "加购件数",
    "商品收藏人数",
    "商品收藏买家数",
    "收藏人数",
    "关注店铺人数",
    "老访客数",
    "新访客数",
    "老买家支付金额",
    "全站推广花费",
    "关键词推广花费",
    "精准人群推广花费",
    "智能场景花费",
    "淘宝客佣金",
    "推广消耗金额",
    "推广费比",
    "推广ROI",
    "展现量",
    "点击量",
    "点击率",
    "平均点击花费",
    "无界推广CPM",
    "推广总加购数",
    "推广引导总成交金额",
    "总成交笔数",
    "无界推广CVR",
    "无界行动转化率",
    "关键词推广ROI",
    "关键词推广展现量",
    "关键词推广点击量",
    "关键词推广CTR",
    "关键词推广CPC",
    "关键词推广CPM",
    "关键词推广成交金额",
    "人群推广花费",
    "人群推广ROI",
    "人群推广展现量",
    "人群推广点击量",
    "人群推广CTR",
    "人群推广CPC",
    "人群推广成交金额",
    "场景推广花费",
    "场景推广ROI",
    "场景推广成交金额",
    "全站推广ROI",
    "全站推广展现量",
    "全站推广点击量",
    "全站推广CTR",
    "全站推广CPC",
    "全站推广成交金额",
    "广告流量",
    "广告流量占比",
    "平台流量",
    "平台流量占比",
}

TEXT_COLUMNS = {
    "统计日期",
    "店铺名称",
    "商品ID",
    "商品名称",
    "品牌名称",
    "一级类目名称",
    "二级类目名称",
    "叶子类目名称",
    "商品状态",
    "流量时期",
    "来源类型",
    "上级来源名称",
    "来源名称",
    "一级流量来源",
    "二级流量来源",
    "三级流量来源",
    "来源层级",
    "归属原则",
}


@dataclass(frozen=True)
class DateParts:
    period_start: pd.Timestamp
    period_end: pd.Timestamp
    period_month: str


def parse_number(value: object) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text in {"", "-", "--", "NaN", "nan", "None"}:
        return 0.0
    is_percent = text.endswith("%")
    text = text.removesuffix("%").replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return 0.0
    number = float(match.group(0))
    return number / 100 if is_percent else number


def normalize_date(value: object) -> DateParts:
    text = str(value).strip()
    range_match = re.match(
        r"^\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})\s*(?:~|至|到|—|–)\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})\s*$",
        text,
    )
    if range_match:
        start_text, end_text = range_match.groups()
        start = pd.to_datetime(start_text)
        end = pd.to_datetime(end_text)
    else:
        start = pd.to_datetime(text)
        end = start
    return DateParts(period_start=start, period_end=end, period_month=start.strftime("%Y-%m"))


def clean_frame(frame: pd.DataFrame) -> pd.DataFrame:
    cleaned = frame.copy()
    for column in cleaned.columns:
        if column in TEXT_COLUMNS:
            cleaned[column] = cleaned[column].map(lambda value: "" if pd.isna(value) else str(value).strip())
        elif column in NUMERIC_COLUMNS:
            cleaned[column] = cleaned[column].map(parse_number)

    if "统计日期" in cleaned.columns:
        dates = cleaned["统计日期"].map(normalize_date)
        cleaned["period_start"] = [item.period_start for item in dates]
        cleaned["period_end"] = [item.period_end for item in dates]
        cleaned["period_month"] = [item.period_month for item in dates]

    if "商品ID" in cleaned.columns:
        cleaned["商品ID"] = cleaned["商品ID"].map(normalize_product_id)

    return cleaned


def normalize_product_id(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().replace(",", "")
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text
