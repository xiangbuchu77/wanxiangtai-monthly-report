from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd

from .periods import label_for_date


DEFAULT_INPUT = Path("/Users/gordon/Downloads/计划报表_20260616_100007.csv")
DEFAULT_PRODUCT_INPUT = Path("/Users/gordon/Downloads/商品报表_20260616_102609.csv")
DEFAULT_OUTPUT = Path("outputs/ecommerce-monthly/月报数据半成品_广告汇总.xlsx")


def main() -> None:
    parser = argparse.ArgumentParser(description="将万相台/营销场景日报汇总为月报agent需要的广告数据表")
    parser.add_argument("--input", type=Path, nargs="+", default=[DEFAULT_INPUT])
    parser.add_argument("--product-input", type=Path, default=DEFAULT_PRODUCT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    result = convert_marketing_scene(args.input, args.product_input)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        result["monthly_summary"].to_excel(writer, sheet_name="monthly_summary", index=False)
        result["plan_detail"].to_excel(writer, sheet_name="plan_detail", index=False)
        result["product_ad"].to_excel(writer, sheet_name="product_ad", index=False)
    print(f"output={args.output}")
    print(result["monthly_summary"].to_string(index=False))
    print(result["plan_detail"].to_string(index=False))


def convert_marketing_scene(paths: list[Path], product_path: Path | None = None, period: str = "month") -> dict[str, pd.DataFrame]:
    frames = [read_csv(path) for path in paths]
    product_frame = read_csv(product_path) if product_path and product_path.exists() else pd.DataFrame()
    if not frames:
        return {
            "monthly_summary": pd.DataFrame(columns=["月份", "广告引导成交金额", "总推广花费", "点击量", "总ROI"]),
            "plan_detail": pd.DataFrame(columns=["月份", "推广计划名称", "计划类型", "花费", "展现量", "点击量", "ROI"]),
            "product_ad": convert_product_ad(product_frame, period),
        }
    frame = pd.concat(frames, ignore_index=True)
    return convert_frames(frame, product_frame, period)


def read_csv(path: Path) -> pd.DataFrame:
    frame = pd.read_csv(path, encoding="gb18030")
    frame.columns = [str(column).strip() for column in frame.columns]
    frame["来源文件"] = path.name
    return frame


def convert_frames(frame: pd.DataFrame, product_frame: pd.DataFrame, period: str = "month") -> dict[str, pd.DataFrame]:
    frame["日期"] = parse_period_start(frame["日期"])
    frame["月份"] = period_label(frame["日期"], period)

    for column in ["展现量", "点击量", "花费", "总成交金额", "投入产出比"]:
        if column in frame.columns:
            frame[column] = pd.to_numeric(frame[column], errors="coerce").fillna(0)

    frame["ROI成交金额_按投入产出比"] = frame["花费"] * frame["投入产出比"]

    monthly = (
        frame.groupby("月份", as_index=False)[["ROI成交金额_按投入产出比", "花费", "点击量"]]
        .sum()
        .rename(columns={"ROI成交金额_按投入产出比": "广告引导成交金额", "花费": "总推广花费"})
    )
    monthly["总ROI"] = monthly["广告引导成交金额"] / monthly["总推广花费"].where(monthly["总推广花费"] != 0)
    monthly = monthly[["月份", "广告引导成交金额", "总推广花费", "点击量", "总ROI"]]

    plan_name_column = "计划名字" if "计划名字" in frame.columns else "场景名字"
    group_columns = ["月份", "场景名字"]
    if plan_name_column != "场景名字":
        group_columns.append(plan_name_column)
    plan = (
        frame.groupby(group_columns, as_index=False)[["花费", "展现量", "点击量", "ROI成交金额_按投入产出比"]]
        .sum()
    )
    plan["推广计划名称"] = plan[plan_name_column]
    plan["计划类型"] = plan["场景名字"].map(plan_type)
    plan["ROI"] = plan["ROI成交金额_按投入产出比"] / plan["花费"].where(plan["花费"] != 0)
    plan = plan[["月份", "推广计划名称", "计划类型", "花费", "展现量", "点击量", "ROI"]]

    product_ad = convert_product_ad(product_frame, period)
    return {"monthly_summary": monthly, "plan_detail": plan, "product_ad": product_ad}


def convert_product_ad(frame: pd.DataFrame, period: str = "month") -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame(columns=["月份", "商品ID", "商品名称", "广告花费", "广告ROI"])
    frame = frame.copy()
    frame.columns = [str(column).strip() for column in frame.columns]
    frame["日期"] = parse_period_start(frame["日期"])
    frame["月份"] = period_label(frame["日期"], period)
    for column in ["花费", "投入产出比"]:
        frame[column] = pd.to_numeric(frame[column], errors="coerce").fillna(0)
    frame["ROI成交金额_按投入产出比"] = frame["花费"] * frame["投入产出比"]
    product = (
        frame.groupby(["月份", "主体ID", "主体名称"], as_index=False)[["花费", "ROI成交金额_按投入产出比"]]
        .sum()
        .rename(columns={"主体ID": "商品ID", "主体名称": "商品名称", "花费": "广告花费"})
    )
    product["广告ROI"] = product["ROI成交金额_按投入产出比"] / product["广告花费"].where(product["广告花费"] != 0)
    product["商品ID"] = product["商品ID"].astype(str)
    product = product[product["广告花费"] > 0]
    return product[["月份", "商品ID", "商品名称", "广告花费", "广告ROI"]]


def period_label(dates: pd.Series, period: str) -> pd.Series:
    if period in {"month", "half-month", "week"}:
        return label_for_date(dates, period)
    return dates.dt.strftime("%Y-%m")


def rolling_30_day_label(dates: pd.Series) -> pd.Series:
    return label_for_date(dates, "month")


def parse_period_start(values: pd.Series) -> pd.Series:
    return values.map(parse_period_start_value)


def parse_period_start_value(value: object) -> pd.Timestamp:
    text = str(value).strip()
    match = re.match(
        r"^\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})\s*(?:~|至|到|—|–)\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})\s*$",
        text,
    )
    if match:
        text = match.group(1)
    return pd.to_datetime(text)


def plan_type(name: object) -> str:
    text = str(name)
    if "关键词" in text:
        return "标准计划"
    if "全站" in text:
        return "全站推广"
    if "人群" in text:
        return "人群推广"
    if "场景" in text:
        return "场景推广"
    return "其他"


if __name__ == "__main__":
    main()
