from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from typing import Any
from uuid import uuid4


@dataclass(frozen=True)
class AIConfig:
    provider: str = ""
    api_key: str = ""
    base_url: str = "https://api.deepseek.com"
    model: str = "deepseek-v4-flash"
    mode: str = "fast"


def config_from_env() -> AIConfig:
    provider = os.environ.get("WXT_AI_PROVIDER", "qclaw").strip() or "qclaw"
    api_key = os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if api_key and not provider:
        provider = "deepseek"
    if provider.lower() in {"qclaw", "openclaw", "claw"}:
        return AIConfig(
            provider=provider,
            api_key="",
            base_url=os.environ.get("QCLAW_LLM_BASE_URL", "http://127.0.0.1:19100/proxy/llm").strip()
            or "http://127.0.0.1:19100/proxy/llm",
            model=os.environ.get("QCLAW_LLM_MODEL", "modelroute").strip() or "modelroute",
            mode=os.environ.get("WXT_AI_MODE", "claw").strip() or "claw",
        )
    return AIConfig(
        provider=provider,
        api_key=api_key,
        base_url=os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com").strip() or "https://api.deepseek.com",
        model=os.environ.get("DEEPSEEK_MODEL", "deepseek-v4-flash").strip() or "deepseek-v4-flash",
        mode=os.environ.get("WXT_AI_MODE", "fast").strip() or "fast",
    )


def maybe_enhance_analysis(draft: dict[str, Any], context: dict[str, Any], config: AIConfig | None = None) -> dict[str, Any]:
    config = config or config_from_env()
    if config.mode.lower() in {"off", "local", "template", "none"}:
        result = dict(draft)
        result["ai_provider_status"] = "local_template"
        return apply_expression_variation(result)
    provider = config.provider.lower()
    if provider == "deepseek" and not config.api_key:
        result = dict(draft)
        result["ai_provider_status"] = "local_template"
        return apply_expression_variation(result)
    if provider not in {"deepseek", "qclaw", "openclaw", "claw"}:
        result = dict(draft)
        result["ai_provider_status"] = "local_template"
        return apply_expression_variation(result)
    try:
        enhanced = call_llm_analysis(draft, compact_context(context), config)
    except Exception as exc:  # noqa: BLE001 - report generation must not fail if AI is unavailable
        result = dict(draft)
        result["ai_provider_status"] = f"{provider}_failed: {exc}"
        return result
    return merge_analysis(draft, enhanced)


def call_llm_analysis(draft: dict[str, Any], context: dict[str, Any], config: AIConfig) -> dict[str, Any]:
    url = config.base_url.rstrip("/") + "/chat/completions"
    draft_for_ai = compact_draft(draft, config.mode)
    style_id = uuid4().hex[:8]
    style_options = [
        "表达更偏运营复盘，先判断趋势，再给动作。",
        "表达更偏管理汇报，先说结论，再说明原因和执行重点。",
        "表达更偏投放优化，突出预算、ROI、点击质量和承接动作。",
        "表达更偏店铺增长，突出流量、转化、商品结构和下月机会。",
    ]
    style_instruction = style_options[int(style_id[:2], 16) % len(style_options)]
    payload = {
        "model": config.model,
        "temperature": 0.65,
        "top_p": 0.9,
        "response_format": {"type": "json_object"},
        "messages": [
            {
                "role": "system",
                "content": (
                    "你是电商万相台月报运营顾问。只能基于用户给出的数据和草稿改写分析文案，"
                    "不得新增、编造或重新计算数字。输出必须是JSON，字段结构与草稿一致。"
                    "每次生成都要保持数字和结论稳健，但避免照搬固定句式，允许在措辞、句序、侧重点上自然变化。"
                ),
            },
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "本次生成批次": {
                            "style_id": style_id,
                            "generated_at": datetime.now().isoformat(timespec="seconds"),
                            "表达风格": style_instruction,
                        },
                        "要求": [
                            "所有数字保留草稿中的表达，不要发明新指标。",
                            "文字采用现象-原因-动作，避免泛泛而谈。",
                            "不要机械复用上一版或草稿中的原句；同义表达可以变化，但业务含义不能变。",
                            "如果两个数据版本接近，也要在文字表达上体现不同的复盘角度。",
                            "本月亮点、存在问题、下月重点关注各至少5条。",
                            "budget_usage由你决定生成几行，字段为推广渠道、预算、占比、用途说明；如果无法判断可返回空数组。",
                            "weekly_plan由你决定操作事项、具体内容、预期效果；时间节点按报表周期最多输出对应周数，如果无法判断可返回空数组。",
                            "预算、操作规划和商品建议要具体可执行。",
                            "如果是极速模式，优先增强summary、goal_paths、budget_usage和weekly_plan；product_actions可以保持简洁。",
                        ],
                        "核心数据": context,
                        "增强模式": config.mode,
                        "草稿JSON": draft_for_ai,
                    },
                    ensure_ascii=False,
                ),
            },
        ],
    }
    headers = {"Content-Type": "application/json"}
    if config.api_key:
        headers["Authorization"] = f"Bearer {config.api_key}"
    request = urllib.request.Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    try:
        with urllib.request.urlopen(request) as response:
            data = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"{config.provider or 'llm'} HTTP {exc.code}: {body[:300]}") from exc
    content = data["choices"][0]["message"]["content"]
    result = parse_json_object(content)
    result["_ai_generation"] = {"provider": config.provider or "llm", "style_id": style_id, "style": style_instruction}
    return result


def call_llm_vision_json(prompt: str, image_url: str | list[str], config: AIConfig) -> dict[str, Any]:
    """Use the configured QClaw/OpenClaw model for a compact screenshot-to-JSON task."""
    if config.mode.lower() in {"off", "local", "template", "none"}:
        raise RuntimeError("当前 AI 增强模式已关闭，无法识别截图数据")
    url = config.base_url.rstrip("/") + "/chat/completions"
    image_urls = [image_url] if isinstance(image_url, str) else list(image_url)
    payload = {
        "model": config.model,
        "temperature": 0,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": "你是严谨的数据识别助手，只输出可解析的 JSON。"},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    *({"type": "image_url", "image_url": {"url": url}} for url in image_urls),
                ],
            },
        ],
    }
    headers = {"Content-Type": "application/json"}
    if config.api_key:
        headers["Authorization"] = f"Bearer {config.api_key}"
    request = urllib.request.Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    try:
        with urllib.request.urlopen(request) as response:
            data = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"截图识别请求失败：HTTP {exc.code} {body[:200]}") from exc
    content = data["choices"][0]["message"]["content"]
    return parse_json_object(content)


def call_deepseek(draft: dict[str, Any], context: dict[str, Any], config: AIConfig) -> dict[str, Any]:
    return call_llm_analysis(draft, context, config)


def chat_deepseek(messages: list[dict[str, str]], config: AIConfig) -> str:
    if not config.api_key:
        raise RuntimeError("请先填写 DeepSeek API Key。")
    url = config.base_url.rstrip("/") + "/chat/completions"
    payload = {
        "model": config.model,
        "temperature": 0.4,
        "messages": [
            {
                "role": "system",
                "content": (
                    "你是万相台诊断师，面向电商运营人员回答问题。"
                    "回答要先讲清概念，再给可执行动作；涉及预算、ROI、关键词、商品、转化时，"
                    "必须提醒用户结合其报表数据判断，不要编造具体店铺数字。"
                ),
            },
            *normalize_chat_messages(messages),
        ],
    }
    request = urllib.request.Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {config.api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            data = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"DeepSeek HTTP {exc.code}: {body[:300]}") from exc
    return str(data["choices"][0]["message"]["content"]).strip()


def normalize_chat_messages(messages: list[dict[str, str]]) -> list[dict[str, str]]:
    clean: list[dict[str, str]] = []
    for item in messages[-12:]:
        role = str(item.get("role", "")).strip()
        content = str(item.get("content", "")).strip()
        if role not in {"user", "assistant"} or not content:
            continue
        clean.append({"role": role, "content": content[:4000]})
    if not clean:
        raise RuntimeError("请输入要咨询的问题。")
    return clean


def parse_json_object(text: str) -> dict[str, Any]:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    return json.loads(text)


def merge_analysis(draft: dict[str, Any], enhanced: dict[str, Any]) -> dict[str, Any]:
    result = dict(draft)
    for key in ["goal_paths", "weekly_plan", "summary"]:
        if valid_value(enhanced.get(key)):
            result[key] = enhanced[key]
    if valid_value(enhanced.get("product_actions")):
        result["product_actions"] = merge_named_items(
            draft.get("product_actions", []),
            enhanced.get("product_actions", []),
            "商品名称",
        )
    if valid_value(enhanced.get("budget_usage")):
        result["budget_usage"] = merge_named_items(
            draft.get("budget_usage", []),
            enhanced.get("budget_usage", []),
            "推广渠道",
        )
    generation = enhanced.get("_ai_generation") or enhanced.get("_deepseek_generation")
    provider = generation.get("provider") if isinstance(generation, dict) else "ai"
    result["ai_provider_status"] = f"{provider}_ok"
    if isinstance(generation, dict):
        result["ai_provider_generation"] = generation
    result = apply_expression_variation(result)
    return result


def merge_named_items(base: Any, updates: Any, key: str) -> list[Any]:
    if not isinstance(base, list) or not isinstance(updates, list):
        return updates if isinstance(updates, list) else base
    update_map = {item.get(key): item for item in updates if isinstance(item, dict)}
    merged = []
    used = set()
    for item in base:
        if isinstance(item, dict) and item.get(key) in update_map:
            merged.append({**item, **update_map[item.get(key)]})
            used.add(item.get(key))
        else:
            merged.append(item)
    for item in updates:
        if isinstance(item, dict) and item.get(key) not in used:
            merged.append(item)
    return merged


def compact_draft(draft: dict[str, Any], mode: str) -> dict[str, Any]:
    if mode in {"full", "精修"}:
        return draft
    if mode in {"standard", "标准"}:
        return {
            "product_actions": draft.get("product_actions", [])[:5],
            "goal_paths": draft.get("goal_paths", {}),
            "budget_usage": draft.get("budget_usage", []),
            "weekly_plan": draft.get("weekly_plan", []),
            "summary": draft.get("summary", {}),
        }
    return {
        "product_actions": draft.get("product_actions", [])[:3],
        "goal_paths": draft.get("goal_paths", {}),
        "budget_usage": draft.get("budget_usage", []),
        "weekly_plan": draft.get("weekly_plan", []),
        "summary": draft.get("summary", {}),
    }


def apply_expression_variation(result: dict[str, Any]) -> dict[str, Any]:
    generation = result.get("ai_provider_generation") or {}
    style_id = str(generation.get("style_id") or uuid4().hex[:8])
    variant = int(style_id[:2], 16) % 4
    varied = dict(result)
    if isinstance(varied.get("product_actions"), list):
        varied["product_actions"] = [
            {
                **item,
                "操作建议": vary_text(str(item.get("操作建议", "")), variant),
            }
            if isinstance(item, dict)
            else item
            for item in varied["product_actions"]
        ]
    if isinstance(varied.get("goal_paths"), dict):
        varied["goal_paths"] = {key: vary_text(str(value), variant) for key, value in varied["goal_paths"].items()}
    if isinstance(varied.get("budget_usage"), list):
        varied["budget_usage"] = [
            {
                **item,
                "用途说明": vary_text(str(item.get("用途说明", "")), variant),
            }
            if isinstance(item, dict)
            else item
            for item in varied["budget_usage"]
        ]
    if isinstance(varied.get("weekly_plan"), list):
        varied["weekly_plan"] = [
            {
                **item,
                "具体内容": vary_text(str(item.get("具体内容", "")), variant),
                "预期效果": vary_text(str(item.get("预期效果", "")), variant),
            }
            if isinstance(item, dict)
            else item
            for item in varied["weekly_plan"]
        ]
    summary = varied.get("summary")
    if isinstance(summary, dict):
        varied["summary"] = {
            key: [vary_text(str(line), variant) for line in value] if isinstance(value, list) else value
            for key, value in summary.items()
        }
    generation = dict(generation)
    generation.setdefault("style_id", style_id)
    generation["local_variant"] = variant
    varied["ai_provider_generation"] = generation
    return varied


def vary_text(text: str, variant: int) -> str:
    if not text:
        return text
    replacements = [
        (),
        (("本月", "本期"), ("下月", "下阶段"), ("重点关注", "优先跟进"), ("建议", "建议")),
        (("本月", "当前周期"), ("下月", "下一周期"), ("需", "需要"), ("关注", "跟进")),
        (("本月", "本轮"), ("下月", "后续"), ("可以", "可"), ("同时", "同步"), ("控制", "压控")),
    ][variant]
    varied = text
    for old, new in replacements:
        varied = varied.replace(old, new)
    return varied


def valid_value(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, (list, dict, str)) and not value:
        return False
    return True


def compact_context(context: dict[str, Any]) -> dict[str, Any]:
    return {
        "period_type": context.get("period_type"),
        "core": context.get("core", {}),
        "previous_core": context.get("previous_core", {}),
        "monthly_ad": context.get("monthly_ad", {}),
        "previous_monthly_ad": context.get("previous_monthly_ad", {}),
        "top_products": context.get("top_products", [])[:6],
        "plan_ad": context.get("plan_ad", [])[:6],
        "product_ad": context.get("product_ad", [])[:6],
        "validation": context.get("validation", {}),
    }
