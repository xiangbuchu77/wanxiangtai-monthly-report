"use client";

import { useMemo, useState } from "react";

type ReportType = "month" | "half" | "week";
type NavKey = "workflow" | "inputs" | "periods" | "delivery";

const reportTypes: Array<{
  key: ReportType;
  label: string;
  window: string;
  comparison: string;
  plan: string;
  filename: string;
}> = [
  {
    key: "month",
    label: "月报",
    window: "最近 30 天",
    comparison: "再往前 30 天",
    plan: "下月推广规划",
    filename: "店铺名YYYY年MM月月报.xlsx",
  },
  {
    key: "half",
    label: "半月报",
    window: "最近 15 天",
    comparison: "再往前 15 天",
    plan: "下半月推广规划",
    filename: "店铺名YYYY年MM月DD日至MM月DD日半月报.xlsx",
  },
  {
    key: "week",
    label: "周报",
    window: "最近 7 天",
    comparison: "再往前 7 天",
    plan: "下周推广规划",
    filename: "店铺名YYYY年MM月DD日至MM月DD日周报.xlsx",
  },
];

const workflowSteps = [
  {
    index: "01",
    title: "接收文件",
    summary: "Excel、CSV、ZIP 或截图进入同一任务。",
    details: ["ZIP 自动解包", "截图识别为补充数据", "同店铺文件归入本次任务"],
  },
  {
    index: "02",
    title: "识别请求",
    summary: "先判断用户是否已明确说出报表类型。",
    details: ["明确说月报、半月报或周报即直接执行", "未说明时才询问是否还有文件", "新指令会使旧任务失效"],
  },
  {
    index: "03",
    title: "计算与校验",
    summary: "以文件中最新有效日期为截止日，统一拆分周期。",
    details: ["核心指标按当前窗口汇总", "仅店铺收益表使用对比窗口", "缺少什么指标就提示什么指标"],
  },
  {
    index: "04",
    title: "填充模板",
    summary: "严格写入对应报表模板，不改动原有布局。",
    details: ["月报、半月报、周报独立模板", "保留合并单元格与边框", "数值格式、百分比与字体统一"],
  },
  {
    index: "05",
    title: "生成分析",
    summary: "QClaw 内置模型只补充需要判断的文字内容。",
    details: ["环比说明避免固定句式", "预算与操作规划基于计划表现", "无有效结论时留空，不编造数据"],
  },
  {
    index: "06",
    title: "交付与清理",
    summary: "仅返回用户请求的 Excel，临时文件自动清理。",
    details: ["钉钉直接发送 Excel", "不额外生成推广月报等文件", "旧临时文件超过 1 天自动清理"],
  },
];

const inputRules = [
  ["结构化报表", "xlsx / csv", "优先读取店铺收益、推广计划、商品与流量数据"],
  ["压缩包", "zip", "自动解包后按内部文件类型读取，避免手工逐个上传"],
  ["经营截图", "png / jpg", "识别指标和表格，用于补充生意参谋等非结构化信息"],
  ["说明材料", "doc / docx", "用于补充业务背景或诊断规则，不覆盖原始数值"],
];

function Dot({ type }: { type: "rule" | "ai" | "output" }) {
  return <span className={`legend-dot ${type}`} aria-hidden="true" />;
}

export default function Home() {
  const [reportType, setReportType] = useState<ReportType>("month");
  const [nav, setNav] = useState<NavKey>("workflow");

  const navigateTo = (key: NavKey) => {
    const sectionId: Record<NavKey, string> = {
      workflow: "workflow",
      inputs: "inputs",
      periods: "periods",
      delivery: "delivery",
    };

    setNav(key);
    document.getElementById(sectionId[key])?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const selectedReport = useMemo(
    () => reportTypes.find((item) => item.key === reportType) ?? reportTypes[0],
    [reportType],
  );

  return (
    <main className="app-shell">
      <aside className="sidebar" aria-label="工作流导航">
        <div className="brand">
          <div className="brand-mark">万</div>
          <div>
            <strong>万相台报表</strong>
            <span>Skill 工作流看板</span>
          </div>
        </div>

        <nav className="nav-list" aria-label="主导航">
          {[
            ["workflow", "工作流总览", "▦"],
            ["inputs", "数据投喂", "◫"],
            ["periods", "周期引擎", "◷"],
            ["delivery", "交付规则", "↗"],
          ].map(([key, label, glyph]) => (
            <button
              className={`nav-item ${nav === key ? "active" : ""}`}
              key={key}
              onClick={() => navigateTo(key as NavKey)}
              type="button"
            >
              <span aria-hidden="true">{glyph}</span>
              {label}
            </button>
          ))}
        </nav>

        <div className="sidebar-note">
          <span className="pulse" aria-hidden="true" />
          <div>
            <strong>QClaw 本地运行</strong>
            <p>无需单独 API Key</p>
          </div>
        </div>

        <div className="account">
          <div className="avatar">WF</div>
          <div>
            <strong>月报 Skill</strong>
            <span>流程说明版</span>
          </div>
        </div>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">万相台 / 自动化报表工作流</p>
            <h1>从数据投喂到 Excel 交付</h1>
            <p className="intro">这不是数据大屏，而是对月报 Skill 如何接收、判断、计算、分析与交付的可视化说明。</p>
          </div>
          <div className="report-switch" aria-label="演示报表类型">
            {reportTypes.map((item) => (
              <button
                className={reportType === item.key ? "selected" : ""}
                key={item.key}
                onClick={() => setReportType(item.key)}
                type="button"
              >
                {item.label}
              </button>
            ))}
          </div>
        </header>

        <section className="hero-strip" aria-label="当前流程演示" id="periods">
          <div className="hero-step"><span>用户指令</span><strong>“生成{selectedReport.label}”</strong></div>
          <span className="flow-arrow" aria-hidden="true">→</span>
          <div className="hero-step"><span>当前窗口</span><strong>{selectedReport.window}</strong></div>
          <span className="flow-arrow" aria-hidden="true">→</span>
          <div className="hero-step"><span>环比窗口</span><strong>{selectedReport.comparison}</strong></div>
          <span className="flow-arrow" aria-hidden="true">→</span>
          <div className="hero-step"><span>最终输出</span><strong>{selectedReport.label} Excel</strong></div>
          <p>截止日始终取已导入文件里的最新有效日期</p>
        </section>

        <section className="workflow-panel" aria-labelledby="workflow-title" id="workflow">
          <div className="section-heading">
            <div>
              <p className="panel-kicker">主流程</p>
              <h2 id="workflow-title">六步生成链路</h2>
            </div>
            <div className="legend" aria-label="流程图图例"><span><Dot type="rule" />规则处理</span><span><Dot type="ai" />模型增强</span><span><Dot type="output" />交付结果</span></div>
          </div>

          <div className="workflow-track">
            {workflowSteps.map((step, index) => (
              <article className={`workflow-card ${index === 4 ? "ai-step" : ""} ${index === 5 ? "output-step" : ""}`} key={step.index}>
                <div className="step-topline"><span className="step-number">{step.index}</span><span className="step-kind">{index === 4 ? "模型增强" : index === 5 ? "交付" : "规则"}</span></div>
                <h3>{step.title}</h3>
                <p>{step.summary}</p>
                <ul>
                  {step.details.map((detail) => <li key={detail}>{detail}</li>)}
                </ul>
              </article>
            ))}
          </div>
        </section>

        <section className="rule-grid">
          <article className="panel input-panel" id="inputs">
            <div className="panel-heading">
              <div>
                <p className="panel-kicker">输入层</p>
                <h2>用户可以怎么投喂数据</h2>
              </div>
              <span className="panel-note">不要求固定文件名</span>
            </div>
            <div className="input-table" role="table" aria-label="支持的输入类型">
              <div className="input-row input-head" role="row"><span>输入</span><span>格式</span><span>处理方式</span></div>
              {inputRules.map(([name, format, detail]) => (
                <div className="input-row" role="row" key={name}>
                  <strong>{name}</strong><span className="format-pill">{format}</span><span>{detail}</span>
                </div>
              ))}
            </div>
          </article>

          <article className="panel control-panel">
            <div className="panel-heading">
              <div>
                <p className="panel-kicker">判断层</p>
                <h2>减少等待与重复生成</h2>
              </div>
            </div>
            <div className="decision-list">
              <div><span className="decision-index">A</span><p><strong>用户明确说生成什么</strong>：立即执行对应类型，只生成被点名的月报、半月报或周报。</p></div>
              <div><span className="decision-index">B</span><p><strong>用户只上传文件未说明类型</strong>：才询问是否还有文件，避免在文件没发完时抢先出报表。</p></div>
              <div><span className="decision-index">C</span><p><strong>用户重新下达生成指令</strong>：此前任务标记作废，避免重复返回多份同名 Excel。</p></div>
            </div>
          </article>
        </section>

        <section className="boundary-section" id="delivery">
          <div className="section-heading">
            <div><p className="panel-kicker">处理与输出层</p><h2>规则计算与模型分析的边界</h2></div>
            <span className="panel-note">数值优先来源于文件，文字才需要增强</span>
          </div>
          <div className="boundary-grid">
            <article className="boundary-card deterministic">
              <div className="boundary-title"><Dot type="rule" /><h3>由程序确定</h3></div>
              <ul>
                <li>按 {selectedReport.window} 汇总店铺收益与推广数据</li>
                <li>店铺总访客数、成交金额、转化率、花费、成交、ROI、CPC、点击量统一计算</li>
                <li>主要商品固定保留 Top4；没有商品数据则省略该板块</li>
                <li>严格按对应模板填充，保留单元格、合并与格式</li>
              </ul>
            </article>
            <article className="boundary-card augmented">
              <div className="boundary-title"><Dot type="ai" /><h3>由 QClaw 内置模型增强</h3></div>
              <ul>
                <li>店铺收益中的环比说明与业务解读</li>
                <li>{selectedReport.plan}中的预算、操作事项、具体内容与预期效果</li>
                <li>运营总结与建议，避免机械重复固定句式</li>
                <li>模型未返回有效结论时保留空白，不伪造文本</li>
              </ul>
            </article>
            <article className="boundary-card delivery-card">
              <div className="boundary-title"><Dot type="output" /><h3>最终交付约束</h3></div>
              <ul>
                <li>仅输出用户要求的 Excel，不产生额外“推广月报”或半成品</li>
                <li>本次演示文件名：{selectedReport.filename}</li>
                <li>生成文件作为钉钉附件回传给发起员工</li>
                <li>临时输出超过 1 天自动清理</li>
              </ul>
            </article>
          </div>
        </section>

        <footer className="footer-note">万相台月报 Skill · 输入层 / 判断层 / 处理层 / 输出层 · QClaw 工作流说明</footer>
      </section>
    </main>
  );
}
