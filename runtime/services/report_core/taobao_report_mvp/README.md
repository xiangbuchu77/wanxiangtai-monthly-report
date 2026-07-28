# 淘宝/天猫运营月报自动化系统 MVP

本模块用于从生意参谋、万相台无界版导出的 Excel 报表中，确定性生成标准经营总表。

## 第一阶段边界

- 不接入 OpenAI。
- 不调用任何 AI 分析。
- 不生成运营建议。
- 不生成 PPT。
- 仅用 Python 规则识别、清洗、汇总和计算指标。

## 已支持报表

- 店铺经营核心月报
- 店铺流量来源构成月报（新版）
- 店铺流量来源构成月报旧版
- 商品经营投产比核心日报
- 商品流量来源构成月报（新版）
- 商品流量来源构成月报旧版
- 商品整体效果月报

## 运行

```bash
/Users/gordon/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m taobao_report_mvp.cli
```

指定目录：

```bash
/Users/gordon/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m taobao_report_mvp.cli --input-dir /path/to/exported_excels --output outputs/ecommerce-monthly/report.xlsx
```

## 输出 Sheet

- `经营总表`
- `店铺流量来源`
- `商品经营`
- `商品投产比日报`
- `商品流量来源`
- `报表识别日志`
- `数据质量`

## AI 月报填充 Agent

第二阶段可以使用 `monthly_report_agent`。它的边界是：

- 数值字段只能来自标准经营总表或结构化广告导出表。
- AI 只生成文字字段，例如商品操作建议、达成路径、预算用途、周计划、总结区。
- 如果点击量、花费、ROI 等关键广告数据缺失，脚本会阻断，不会编造。
- `3.2 本月主要操作记录` 按当前需求跳过。

运行覆盖和任务包生成：

```bash
/Users/gordon/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m taobao_report_mvp.template_coverage
/Users/gordon/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m taobao_report_mvp.monthly_report_agent --ad-metrics outputs/ecommerce-monthly/广告数据导入模板.xlsx
```

广告数据导入表需要包含：

- `monthly_summary`：`月份`、`广告引导成交金额`、`总推广花费`、`点击量`、`总ROI`
- `plan_detail`：`月份`、`推广计划名称`、`计划类型`、`花费`、`展现量`、`点击量`、`ROI`
- `product_ad`：`月份`、`商品ID`、`商品名称`、`广告花费`、`广告ROI`
