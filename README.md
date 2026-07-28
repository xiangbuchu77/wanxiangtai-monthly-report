# 万相台企业报表自动化 Agent

把淘宝/天猫运营人员从生意参谋、万相台无界版导出的原始 Excel、CSV 或截图，自动转换为可交付的万相台月报、半月报或周报。

这个项目面向企业内部运营场景：用户只需上传一个报表文件夹，Agent 会完成文件识别、数据校验、确定性指标计算、分析文案生成、Excel 模板填充与交付清单输出。

![万相台企业报表自动化 Agent 工作流](https://raw.githubusercontent.com/xiangbuchu77/wanxiangtai-monthly-report/main/dashboard/public/og.png)

## 为什么做这个项目

电商运营月报通常要从多份结构不同的报表中整理数据，再人工计算指标、填写模板并撰写结论。这个 Agent 把重复、易错的步骤收敛为一条可检查的交付链路：**程序计算数字，模型只辅助解释数字。**

## 能做什么

- 接收 Excel、CSV、ZIP、截图和补充说明材料，不要求固定文件名。
- 自动识别文件类型、清洗字段，并按最新有效日期生成月报、半月报或周报。
- 对店铺收益、广告花费、成交金额、ROI、CPC、点击量等指标进行确定性计算。
- 只在需要判断的文字区域使用 AI，例如商品操作建议、推广规划和运营总结。
- 严格写入既有 Excel 模板，保留合并单元格、边框、列宽和数值格式。
- 输出最终 Excel、缺失数据说明、标准经营总表和交付清单；临时半成品会自动清理。

## 工作流

1. **接收文件**：识别上传的结构化报表、压缩包和截图，并归类到同一任务。
2. **识别请求**：根据用户指令判断需要生成月报、半月报还是周报。
3. **计算与校验**：统一周期、校验核心指标，明确指出缺少的具体数据项。
4. **填充模板**：把计算结果写入对应报表模板，同时保持模板布局不变。
5. **生成分析**：模型只补充基于数据的业务解释与操作建议，不参与数值计算。
6. **交付与清理**：输出用户指定的最终 Excel，并清理中间文件。

## 规则与模型的边界

| 由 Python 确定 | 由 AI 辅助 |
| --- | --- |
| 指标汇总与环比计算 | 商品操作建议 |
| ROI、CPC、点击量和成交金额 | 下周期推广规划 |
| 周期判断与模板填充 | 运营总结与业务解读 |
| 数据缺失检查与交付清单 | 仅在有证据时生成文字结论 |

AI 不得编造数字，也不负责计算指标；模型未返回有效结论时会保留空白。

## 交付内容

企业入口在数据足够时会生成：

- `00_数据缺失与识别报告.xlsx`
- `01_标准经营总表.xlsx`
- `03_AI分析文案.json`
- `04_AI任务包.json`
- `05_万相台月报_最终版.xlsx`
- `enterprise_manifest.json`
- `交付说明.md`

`02_月报数据半成品_广告汇总.xlsx` 只作为临时文件使用，最终交付前会被删除。多个周期会拆分为带日期后缀的独立报表。

## 快速开始

安装项目依赖后，传入包含原始报表的目录和 Excel 模板：

```bash
python -m taobao_report_mvp.enterprise_agent /path/to/uploaded_reports \
  --template /path/to/万相台月报模板.xlsx
```

安装为命令行脚本后：

```bash
wanxiangtai-agent /path/to/uploaded_reports \
  --template /path/to/万相台月报模板.xlsx
```

指定输出类型或周期：

```bash
wanxiangtai-agent /path/to/uploaded_reports --report-type 周报
wanxiangtai-agent /path/to/uploaded_reports --report-type 半月报 --target 2026-05上半月
wanxiangtai-agent /path/to/uploaded_reports --target 2026-05
```

## 部署入口

### 企业 Web / 钉钉

推荐部署为统一 Web 服务并嵌入钉钉 H5 应用，让员工在浏览器或钉钉内直接提交报表任务：

```bash
PYTHONPATH=src python3 start_dingtalk_service.py
```

默认监听 `http://服务器地址:8788`。钉钉接入说明见 `docs/DINGTALK_INTEGRATION.md`。

### QClaw 本地入口

桌面侧保留 QClaw 与钉钉服务入口。QClaw 调用本地 Agent 时会自动启动服务，也可以手动运行：

```text
启动QClaw万相台Agent.command
```

默认监听 `http://127.0.0.1:8799`。生成成功后，QClaw 会调用系统打开最终 `.xlsx`。

## 数据要求

项目通过字段和数据结构识别报表，不依赖固定文件名。最小阻塞指标通常是：

- 店铺总访客数
- 店铺总成交金额（元）
- 店铺支付转化率

广告花费、广告成交金额、ROI、CPC 和点击量会从上传数据中识别或计算；缺失时作为建议补充，而不阻断简版报表的生成。

建议额外上传商品整体效果、计划、商品、流量来源、商品经营投产比和关键词等报表，以提高商品、计划、关键词、人群与流量来源分析的颗粒度。

## 仓库结构

```text
src/                 核心 Agent、报表识别与计算逻辑
runtime/             企业服务与 Excel 渲染能力
dashboard/           工作流可视化页面
skills/              QClaw Skill、提示词与报表模板
tests/               核心流程测试
```

## 内部调试

```bash
python -m taobao_report_mvp.report_workflow_agent <原始表格1> <原始表格2> \
  --output-dir outputs/run
```
