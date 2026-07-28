# 万相台企业报表自动化 Agent

这个项目用于把淘宝/天猫运营人员从生意参谋、万相台无界版导出的原始 Excel/CSV 报表，自动生成可交付的万相台月报、周报或半月报。

企业用户可以只上传一个文件夹，Agent 会自动完成报表识别、数据校验、确定性计算、分析文案生成、Excel 填充、美化排版和交付清单输出。

## 核心流程

1. 用户上传原始 Excel/CSV 表格或一个报表文件夹。
2. Agent 自动识别报表类型、清洗字段、按周期汇总数据。
3. 默认生成月报；用户指定周报或半月报时按要求生成。
4. 用户未指定月份/周期时，默认只生成数据里最近一个可生成周期；用户明确要求多周期时再按周期拆分生成。
5. 如果缺少核心数据，输出缺失清单并说明需要补哪些报表。
6. 如果数据齐全，生成最终 Excel 报表，并删除广告半成品临时文件。
7. AI 只介入文字分析字段：
   - 主要商品数据的操作建议
   - 四、下月推广规划
   - 五、运营总结与建议
8. 脚本把数值和 AI 文案写入模板，并统一字体、边框、列宽、换行和合并单元格。

## 关键原则

- 数值字段必须由 Python 确定性计算。
- AI 不得编造数字。
- AI 不负责计算指标，只负责基于半成品数据生成具体分析文案。
- 缺少店铺收益核心指标时才阻断，并且必须提示具体缺少的数据内容，例如“店铺总访客数”“店铺总成交金额（元）”“店铺支付转化率”；不要用“月度推广核心指标总览数据”这类文件/模块名代替指标名。
- 非核心增强文件缺失时生成简版，并在交付说明里提示建议补充。

## 企业入口

```bash
python -m taobao_report_mvp.enterprise_agent /path/to/uploaded_reports --template /path/to/万相台月报模板.xlsx
```

安装为命令行脚本后：

```bash
wanxiangtai-agent /path/to/uploaded_reports --template /path/to/万相台月报模板.xlsx
```

## 钉钉企业入口

面向企业员工使用时，推荐部署为统一 Web 服务并嵌入钉钉 H5 应用。这样员工不再需要区分 macOS 或 Windows，也不需要安装本地启动包。

```bash
PYTHONPATH=src python3 start_dingtalk_service.py
```

默认监听：

```text
http://服务器地址:8788
```

钉钉接入方式见：

```text
docs/DINGTALK_INTEGRATION.md
```

## QClaw 本地入口

当前桌面侧只保留 QClaw/钉钉服务入口，不再维护旧的本地网页启动包。QClaw 调用本地 Agent 时会自动启动服务；也可以手动双击：

```text
启动QClaw万相台Agent.command
```

默认监听：

```text
http://127.0.0.1:8799
```

生成成功后，QClaw skill 会调用系统打开最终 `.xlsx`，用户无需复制本地路径。

生成周报：

```bash
wanxiangtai-agent /path/to/uploaded_reports --report-type 周报
```

生成半月报：

```bash
wanxiangtai-agent /path/to/uploaded_reports --report-type 半月报
```

指定某个周期：

```bash
wanxiangtai-agent /path/to/uploaded_reports --target 2026-05
wanxiangtai-agent /path/to/uploaded_reports --report-type 半月报 --target 2026-05上半月
wanxiangtai-agent /path/to/uploaded_reports --report-type 周报 --target 2026-05-04~2026-05-10
```

## 输出内容

企业入口会生成：

- `00_数据缺失与识别报告.xlsx`
- `01_标准经营总表.xlsx`
- `03_AI分析文案.json`
- `04_AI任务包.json`
- `05_万相台月报_最终版.xlsx`
- `enterprise_manifest.json`
- `交付说明.md`

说明：

- `02_月报数据半成品_广告汇总.xlsx` 是临时文件，最终版生成后会自动删除。
- 如果核心数据里包含多个周期，最终报表会带周期后缀，例如 `05_万相台月报_最终版_2026-05.xlsx`。
- 如果部分周期缺核心数据，只生成可用周期，缺失周期进入交付说明和识别报告。

## 核心输入

最简生成优先依赖“店铺收益”相关数据，而不是固定文件名。文件可以任意命名，Agent 会按字段和数据结构识别。

最小阻塞指标通常是：

- 店铺总访客数
- 店铺总成交金额（元）
- 店铺支付转化率

广告相关指标如“广告总花费（元）”“广告带来的成交金额（元）”“广告投入产出比（ROI）”“平均点击成本（元）”“总点击量”能从上传数据中识别或计算时就自动补入；缺少时只作为建议补充，不应固定要求用户补某个指定文件。

只有部分核心数据时，Agent 会尽量生成简版报表，并略去主要商品、计划明细等无数据模块。

## 建议输入

- 商品整体效果月报
- 计划报表
- 商品报表
- 店铺流量来源构成月报新版/旧版
- 商品流量来源构成月报新版/旧版
- 商品经营投产比核心日报
- 关键词报表

建议输入不阻断生成，但会提升商品、计划、关键词、人群和流量来源分析颗粒度。

## 内部工作流入口

如需调试内部流程：

```bash
python -m taobao_report_mvp.report_workflow_agent <原始表格1> <原始表格2> --output-dir outputs/run
```
