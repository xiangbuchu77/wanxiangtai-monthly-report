from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


DEFAULT_TOTAL = Path("outputs/ecommerce-monthly/淘宝天猫运营标准经营总表_MVP.xlsx")
DEFAULT_TEMPLATE = Path("/Users/gordon/Desktop/万相台月报模板.xlsx")
DEFAULT_OUTPUT = Path("outputs/ecommerce-monthly/月报模板数据覆盖检查.xlsx")


@dataclass(frozen=True)
class Requirement:
    section: str
    template_range: str
    field: str
    source_sheet: str
    source_columns: tuple[str, ...]
    availability: str
    evidence: str
    action: str


def main() -> None:
    parser = argparse.ArgumentParser(description="检查标准经营总表对万相台月报模板的数据覆盖情况")
    parser.add_argument("--total", type=Path, default=DEFAULT_TOTAL)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    requirements = build_coverage(args.total, args.template)
    frame = pd.DataFrame([item.__dict__ for item in requirements])
    summary = (
        frame.groupby(["section", "availability"], as_index=False)
        .size()
        .rename(columns={"size": "count"})
        .sort_values(["section", "availability"])
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="覆盖明细", index=False)
        summary.to_excel(writer, sheet_name="覆盖汇总", index=False)
    print(f"output={args.output}")
    print(frame[["section", "field", "availability", "evidence"]].to_string(index=False))


def build_coverage(total_path: Path, template_path: Path) -> list[Requirement]:
    sheets = pd.read_excel(total_path, sheet_name=None)
    template = pd.read_excel(template_path, sheet_name="Sheet1", header=None)
    latest_month = latest_complete_month(sheets)

    checks: list[Requirement] = []
    checks.extend(check_core_metrics(sheets, latest_month))
    checks.extend(check_product_metrics(sheets, latest_month))
    checks.extend(check_promotion_metrics(sheets, latest_month))
    checks.extend(check_planning_and_summary(template))
    return checks


def latest_complete_month(sheets: dict[str, pd.DataFrame]) -> str:
    core = sheets.get("经营总表", pd.DataFrame())
    if core.empty or "月份" not in core.columns:
        return ""
    return str(core["月份"].dropna().max())


def has_columns(frame: pd.DataFrame, columns: tuple[str, ...]) -> bool:
    return all(column in frame.columns for column in columns)


def nonzero_sum(frame: pd.DataFrame, column: str, month: str | None = None) -> bool:
    if frame.empty or column not in frame.columns:
        return False
    scoped = frame
    if month and "月份" in scoped.columns:
        scoped = scoped[scoped["月份"].astype(str).eq(month)]
    return pd.to_numeric(scoped[column], errors="coerce").fillna(0).sum() > 0


def row_count(frame: pd.DataFrame, month: str | None = None) -> int:
    if frame.empty:
        return 0
    if month and "月份" in frame.columns:
        return int(frame[frame["月份"].astype(str).eq(month)].shape[0])
    return int(frame.shape[0])


def req(
    section: str,
    template_range: str,
    field: str,
    source_sheet: str,
    source_columns: tuple[str, ...],
    ok: bool,
    evidence: str,
    action: str,
) -> Requirement:
    return Requirement(
        section=section,
        template_range=template_range,
        field=field,
        source_sheet=source_sheet,
        source_columns=source_columns,
        availability="可用" if ok else "缺失",
        evidence=evidence,
        action=action,
    )


def check_core_metrics(sheets: dict[str, pd.DataFrame], month: str) -> list[Requirement]:
    core = sheets.get("经营总表", pd.DataFrame())
    traffic = sheets.get("店铺流量来源", pd.DataFrame())
    rows = row_count(core, month)
    paid_traffic = traffic[
        traffic.get("一级来源", pd.Series(dtype=object)).astype(str).isin(["广告流量", "付费推广"])
        & traffic.get("月份", pd.Series(dtype=object)).astype(str).eq(month)
    ] if not traffic.empty else pd.DataFrame()
    paid_amount_ok = nonzero_sum(paid_traffic, "支付金额") if not paid_traffic.empty else False
    ad_cost_ok = nonzero_sum(core, "推广花费合计", month)
    click_ok = nonzero_sum(sheets.get("商品投产比日报", pd.DataFrame()), "点击量", month)
    impressions_ok = nonzero_sum(sheets.get("商品投产比日报", pd.DataFrame()), "展现量", month)

    return [
        req("一、核心数据总览", "B3:D3", "店铺总访客数", "经营总表", ("访客数", "访客数_环比"), rows > 0 and nonzero_sum(core, "访客数", month), f"{month} 行数={rows}", "可直接填充"),
        req("一、核心数据总览", "B4:D4", "店铺总成交金额", "经营总表", ("支付金额", "支付金额_环比"), rows > 0 and nonzero_sum(core, "支付金额", month), f"{month} 行数={rows}", "可直接填充"),
        req("一、核心数据总览", "B5:D5", "店铺支付转化率", "经营总表", ("支付转化率_计算",), rows > 0 and nonzero_sum(core, "支付转化率_计算", month), f"{month} 行数={rows}", "可直接填充"),
        req("一、核心数据总览", "B6:D6", "广告总花费", "经营总表", ("推广花费合计", "全站推广花费", "关键词推广花费"), ad_cost_ok, f"{month} 推广花费合计非零={ad_cost_ok}", "可直接填充"),
        req("一、核心数据总览", "B7:D7", "广告带来的成交金额", "店铺流量来源", ("一级来源", "支付金额"), paid_amount_ok, f"{month} 广告/付费来源支付金额非零={paid_amount_ok}", "可用流量来源归因填充"),
        req("一、核心数据总览", "B8:D8", "广告ROI", "经营总表+店铺流量来源", ("推广花费合计", "支付金额"), ad_cost_ok and paid_amount_ok, f"花费={ad_cost_ok}; 广告成交={paid_amount_ok}", "用广告成交金额/广告总花费计算"),
        req("一、核心数据总览", "B9:D9", "平均点击成本", "商品投产比日报", ("推广消耗金额", "点击量"), ad_cost_ok and click_ok, f"{month} 点击量非零={click_ok}", "缺点击量时不可严谨计算"),
        req("一、核心数据总览", "B10:D10", "总点击量", "商品投产比日报", ("点击量",), click_ok, f"{month} 点击量非零={click_ok}", "需补充有点击量的万相台报表"),
        req("三、推广进展分析", "D21:F26", "展现量/点击率", "商品投产比日报", ("展现量", "点击量"), impressions_ok and click_ok, f"{month} 展现={impressions_ok}; 点击={click_ok}", "需补充计划层级展现点击数据"),
    ]


def check_product_metrics(sheets: dict[str, pd.DataFrame], month: str) -> list[Requirement]:
    product = sheets.get("商品经营", pd.DataFrame())
    roi = sheets.get("商品投产比日报", pd.DataFrame())
    product_rows = row_count(product, month)
    ad_spend_ok = nonzero_sum(roi, "推广消耗金额", month)
    ad_roi_ok = nonzero_sum(roi, "推广ROI_计算", month)
    return [
        req("二、主要商品数据", "A14:E18", "商品名称/访客/支付/件数/转化率", "商品经营", ("商品名称", "商品访客数", "支付金额", "支付件数", "支付转化率_计算"), product_rows > 0, f"{month} 商品行数={product_rows}", "可取支付金额Top商品填充"),
        req("二、主要商品数据", "F14:G18", "商品广告花费/广告ROI", "商品投产比日报", ("推广消耗金额", "推广ROI_计算"), ad_spend_ok and ad_roi_ok, f"{month} 商品推广消耗非零={ad_spend_ok}; ROI非零={ad_roi_ok}", "需补充商品推广消耗与成交"),
        req("二、主要商品数据", "H14:H18", "操作建议", "AI分析", ("商品经营", "商品流量来源"), product_rows > 0, f"{month} 商品行数={product_rows}", "可由AI基于数据生成；非确定性字段"),
    ]


def check_promotion_metrics(sheets: dict[str, pd.DataFrame], month: str) -> list[Requirement]:
    core = sheets.get("经营总表", pd.DataFrame())
    roi = sheets.get("商品投产比日报", pd.DataFrame())
    plan_cost_cols = ("关键词推广花费", "全站推广花费", "人群推广花费", "场景推广花费")
    plan_cost_ok = any(nonzero_sum(core, col, month) or nonzero_sum(roi, col, month) for col in plan_cost_cols)
    plan_click_ok = any(nonzero_sum(roi, col, month) for col in ("关键词推广点击量", "全站推广点击量", "人群推广点击量", "场景推广点击量", "点击量"))
    plan_revenue_ok = any(nonzero_sum(roi, col, month) for col in ("关键词推广成交金额", "全站推广成交金额", "人群推广成交金额", "场景推广成交金额", "推广引导总成交金额"))
    return [
        req("三、推广进展分析", "C22:C26", "计划花费", "经营总表/商品投产比日报", plan_cost_cols, plan_cost_ok, f"{month} 计划花费非零={plan_cost_ok}", "可部分填充花费，但计划名称需映射"),
        req("三、推广进展分析", "D22:G26", "计划展现/点击/CPC", "商品投产比日报", ("展现量", "点击量", "平均点击花费_计算"), plan_click_ok, f"{month} 计划点击非零={plan_click_ok}", "缺计划层级点击展现"),
        req("三、推广进展分析", "H22:H26", "计划ROI", "商品投产比日报", ("推广消耗金额", "推广引导总成交金额"), plan_cost_ok and plan_revenue_ok, f"{month} 计划成交非零={plan_revenue_ok}", "缺计划层级推广成交"),
    ]


def check_planning_and_summary(template: pd.DataFrame) -> list[Requirement]:
    return [
        req("四、下月推广规划", "B39:D42", "下月核心目标", "AI分析+经营总表", ("支付金额", "访客数", "支付转化率_计算"), True, "可用历史指标派生，需业务目标规则或AI", "可由agent填充"),
        req("四、下月推广规划", "B46:D51", "下月预算分配", "AI分析+推广表现", ("推广花费合计", "推广ROI_计算"), True, "可用历史花费作为约束，分配方案需AI/规则", "可由agent填充"),
        req("四、下月推广规划", "C55:D58", "下月重点操作规划", "AI分析", ("经营总表", "商品经营", "流量来源"), True, "模板固定周次，内容需AI生成", "可由agent填充"),
        req("五、运营总结与建议", "A61:D69", "本月亮点/存在问题/下月重点关注", "AI分析", ("经营总表", "商品经营", "流量来源"), True, "截图所示文本区需AI总结", "可由agent填充"),
        req("三、推广进展分析", "A30:D34", "3.2 本月主要操作记录", "手工记录", tuple(), True, "用户指定可以不要", "跳过不填"),
    ]


if __name__ == "__main__":
    main()
