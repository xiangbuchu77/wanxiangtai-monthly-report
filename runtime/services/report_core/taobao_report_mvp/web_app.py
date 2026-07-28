from __future__ import annotations

import html
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
import traceback
from datetime import date, datetime
from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, quote, unquote, urlparse

import argparse
import threading
import webbrowser

import pandas as pd

from .ai_provider import AIConfig, chat_deepseek
from .archive_utils import extract_zip_archive, sanitize_windows_filename
from .enterprise_agent import build_enterprise_result, normalize_period, write_enterprise_manifest, write_handoff_note
from .report_workflow_agent import DEFAULT_TEMPLATE_CANDIDATES, build_output_paths, first_existing, run_workflow
from .screenshot_input import IMAGE_SUFFIXES, extract_screenshot_report_files


APP_TITLE = os.environ.get("WXT_APP_TITLE", "万相台报表 Agent")
SUPPORTED_SUFFIXES = {".xlsx", ".xls", ".csv", ".zip"}
QCLAW_UPLOAD_SUFFIXES = {".xlsx", ".xls", ".csv", ".doc", ".docx", ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".zip"}
QCLAW_REPORT_SUFFIXES = {".xlsx", ".xls", ".csv"}
QCLAW_EXTRACTABLE_SUFFIXES = QCLAW_REPORT_SUFFIXES | {".doc", ".docx", ".png", ".jpg", ".jpeg", ".webp", ".bmp"}
APP_ROOT = Path(os.environ.get("WXT_APP_ROOT", Path(__file__).resolve().parents[2]))
DEFAULT_HOST = os.environ.get("WXT_HOST", "127.0.0.1")
DEFAULT_PORT = int(os.environ.get("WXT_PORT", "8788"))
OUTPUT_ROOT = Path(os.environ.get("WXT_OUTPUT_ROOT", str(APP_ROOT / "outputs" / "web_runs")))
FINAL_REPORT_DIR = Path(os.environ.get("WXT_FINAL_REPORT_DIR", str(APP_ROOT / "outputs" / "final_reports")))
SOURCE_CACHE_DIR = Path(os.environ.get("WXT_SOURCE_CACHE_DIR", str(APP_ROOT / "cache" / "source")))
DAILY_SOURCE_CACHE_DIR = Path(os.environ.get("WXT_DAILY_SOURCE_CACHE_DIR", str(SOURCE_CACHE_DIR / "daily")))
TEMP_OUTPUT_MAX_AGE_DAYS = int(os.environ.get("WXT_TEMP_OUTPUT_MAX_AGE_DAYS", "7"))
CORE_METRIC_LABELS = [
    "店铺总访客数",
    "店铺总成交金额（元）",
    "店铺支付转化率",
    "广告总花费（元）",
    "广告带来的成交金额（元）",
    "广告投入产出比（ROI）",
    "平均点击成本（元）",
    "总点击量",
]
REPORT_SESSIONS: dict[str, "ReportCollectionSession"] = {}
LAST_REPORTS: dict[str, "LastReportSession"] = {}
REPORT_SESSIONS_LOCK = threading.Lock()
ACTIVE_GENERATION_TOKENS: dict[str, str] = {}
ACTIVE_GENERATION_LOCK = threading.Lock()


def start_generation_request(session_key: str) -> str:
    token = f"{time.time_ns()}"
    with ACTIVE_GENERATION_LOCK:
        ACTIVE_GENERATION_TOKENS[session_key] = token
    return token


def generation_is_current(session_key: str, token: str) -> bool:
    with ACTIVE_GENERATION_LOCK:
        return ACTIVE_GENERATION_TOKENS.get(session_key) == token


def obsolete_generation_response() -> dict[str, Any]:
    return {
        "type": "cancelled",
        "text": "已收到新的生成指令，上一轮生成已停止返回，请以最新生成结果为准。",
        "files": [],
        "audit": {"cancelled_by_new_generation_request": True},
    }


def discard_obsolete_generation(result: dict[str, Any] | None = None, output_dir: Path | None = None) -> dict[str, Any]:
    if result:
        for item in result.get("final_reports") or []:
            try:
                Path(str(item)).unlink(missing_ok=True)
            except Exception:
                pass
    if output_dir:
        shutil.rmtree(output_dir, ignore_errors=True)
    return obsolete_generation_response()


@dataclass
class FormPart:
    name: str
    value: str = ""
    filename: str | None = None
    data: bytes = b""


@dataclass
class ReportCollectionSession:
    session_key: str
    period: str = "month"
    target: str | None = None
    store_name: str | None = None
    files: list[Path] | None = None
    created_at: str = ""
    updated_at: str = ""

    def __post_init__(self) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        if self.files is None:
            self.files = []
        if not self.created_at:
            self.created_at = now
        if not self.updated_at:
            self.updated_at = now


@dataclass
class LastReportSession:
    session_key: str
    period: str = "month"
    target: str | None = None
    source_files: list[Path] | None = None
    final_reports: list[Path] | None = None
    recommended_missing: list[str] | None = None
    updated_at: str = ""

    def __post_init__(self) -> None:
        if self.source_files is None:
            self.source_files = []
        if self.final_reports is None:
            self.final_reports = []
        if self.recommended_missing is None:
            self.recommended_missing = []
        if not self.updated_at:
            self.updated_at = datetime.now().isoformat(timespec="seconds")


@dataclass
class DailyArchiveResult:
    stores: dict[str, list[Path]]
    unknown_files: list[Path]


def main() -> None:
    parser = argparse.ArgumentParser(description="启动万相台企业报表 Agent 网页界面")
    parser.add_argument("--host", default=DEFAULT_HOST, help="监听地址，默认 127.0.0.1")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT, help="监听端口，默认 8788")
    parser.add_argument("--open", action="store_true", help="启动后自动打开浏览器")
    args = parser.parse_args()

    cleanup_old_output_runs()
    cleanup_daily_source_cache()
    server, port = create_server(args.host, args.port)
    url = f"http://{args.host}:{port}"
    print(f"{APP_TITLE} 已启动: {url}")
    print("按 Ctrl+C 停止服务。")
    if args.open:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    server.serve_forever()


def create_server(host: str, preferred_port: int) -> tuple[ThreadingHTTPServer, int]:
    cleanup_old_output_runs()
    errors: list[str] = []
    for port in range(preferred_port, preferred_port + 20):
        try:
            return ThreadingHTTPServer((host, port), ReportAppHandler), port
        except OSError as exc:
            if exc.errno not in {1, 13, 48, 49, 98, 99, 10013, 10048}:
                raise
            errors.append(f"{port}: {exc}")
    message = "无法启动本地网页服务，已尝试端口 " + f"{preferred_port}-{preferred_port + 19}。"
    if errors:
        message += "\n" + "\n".join(errors[-5:])
    raise RuntimeError(message)


class ReportAppHandler(BaseHTTPRequestHandler):
    server_version = "WanxiangtaiReportAgent/1.0"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_html(render_home())
        elif parsed.path == "/chat":
            self.send_html(render_chat())
        elif parsed.path == "/healthz":
            self.send_json({"status": "ok", "app": APP_TITLE})
        elif parsed.path == "/download":
            self.handle_download(parsed.query)
        else:
            self.send_error(HTTPStatus.NOT_FOUND, "页面不存在")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/chat":
            self.handle_chat_api()
            return
        if parsed.path == "/qclaw/message":
            self.handle_qclaw_message()
            return
        if parsed.path != "/generate":
            self.send_error(HTTPStatus.NOT_FOUND, "页面不存在")
            return
        try:
            result = self.handle_generate()
            self.send_html(render_result(result))
        except Exception as exc:  # noqa: BLE001 - web UI should show friendly failure
            error = {
                "status": "error",
                "message": str(exc),
                "trace": traceback.format_exc(),
            }
            self.send_html(render_error(error), status=HTTPStatus.INTERNAL_SERVER_ERROR)

    def log_message(self, format: str, *args: Any) -> None:
        print(f"[web] {self.address_string()} - {format % args}")

    def handle_generate(self) -> dict[str, Any]:
        form = self.parse_form()
        period = normalize_period(field_value(form, "period", "month"))
        target = field_value(form, "target", "").strip() or None
        run_name = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = OUTPUT_ROOT / f"{period}_{run_name}"
        upload_dir = output_dir / "uploads"
        upload_dir.mkdir(parents=True, exist_ok=True)

        upload_sources = save_uploads(form, "reports", upload_dir)
        uploads = expand_archives(upload_sources, output_dir / "unzipped", SUPPORTED_SUFFIXES)
        template_files = save_uploads(form, "template", upload_dir / "template")
        template = template_files[0] if template_files else first_existing(DEFAULT_TEMPLATE_CANDIDATES)
        deepseek_api_key = field_value(form, "deepseek_api_key", "").strip() or os.environ.get("DEEPSEEK_API_KEY", "")
        ai_provider = field_value(form, "ai_provider", "").strip()
        if deepseek_api_key and not ai_provider:
            ai_provider = "deepseek"
        ai_config = AIConfig(
            provider=ai_provider,
            api_key=deepseek_api_key,
            base_url=field_value(form, "deepseek_base_url", "").strip() or os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
            model=field_value(form, "deepseek_model", "").strip() or os.environ.get("DEEPSEEK_MODEL", "deepseek-v4-flash"),
            mode=field_value(form, "ai_mode", "").strip() or os.environ.get("WXT_AI_MODE", "fast"),
        )

        if not uploads:
            core_metric_text = "、".join(CORE_METRIC_LABELS)
            result = {
                "status": "blocked",
                "message": f"没有找到可处理的数据。请补充以下指标：{core_metric_text}。文件名不限，系统会按字段识别。",
                "core_dataset_name": "店铺收益核心指标",
                "expected_core_data": [
                    *CORE_METRIC_LABELS,
                    "商品经营数据：商品名称、商品访客、支付金额、支付转化率（可选，用于主要商品模块）",
                ],
                "output_dir": str(output_dir),
                "final_reports": [],
            }
            write_enterprise_manifest(output_dir, uploads, period, target, result)
            write_handoff_note(output_dir, result)
            return result

        outputs = build_output_paths(output_dir)
        raw_result = run_workflow(uploads, template, outputs, target, period, ai_config=ai_config)
        result = build_enterprise_result(raw_result, uploads, output_dir, period, target)
        write_enterprise_manifest(output_dir, uploads, period, target, result)
        write_handoff_note(output_dir, result)
        return result

    def handle_chat_api(self) -> None:
        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            config = AIConfig(
                provider="deepseek",
                api_key=str(payload.get("api_key", "")).strip(),
                base_url=str(payload.get("base_url", "")).strip() or "https://api.deepseek.com",
                model=str(payload.get("model", "")).strip() or "deepseek-v4-flash",
            )
            answer = chat_deepseek(payload.get("messages") or [], config)
            self.send_json({"status": "ok", "answer": answer})
        except Exception as exc:  # noqa: BLE001 - return friendly chatbot error
            self.send_json({"status": "error", "message": str(exc)}, status=HTTPStatus.BAD_REQUEST)

    def handle_qclaw_message(self) -> None:
        try:
            content_type = self.headers.get("Content-Type", "")
            if content_type.startswith("multipart/form-data"):
                result = self.handle_qclaw_multipart()
            else:
                result = self.handle_qclaw_json()
            self.send_json(result)
        except Exception as exc:  # noqa: BLE001 - qclaw needs machine readable errors
            self.send_json(
                {
                    "type": "error",
                    "text": f"处理失败：{exc}",
                    "files": [],
                    "audit": {"error": traceback.format_exc()},
                },
                status=HTTPStatus.BAD_REQUEST,
            )

    def handle_qclaw_json(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0"))
        payload = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
        text = str(payload.get("text", "")).strip()
        files = [Path(str(item.get("path", ""))).expanduser() for item in payload.get("files", []) if item.get("path")]
        settings = dict(payload.get("settings", {}))
        return self.run_qclaw_request(text, files, settings)

    def handle_qclaw_multipart(self) -> dict[str, Any]:
        form = self.parse_form()
        text = field_value(form, "text", "").strip()
        settings = {
            "deepseek_api_key": field_value(form, "deepseek_api_key", "").strip(),
            "deepseek_model": field_value(form, "deepseek_model", "deepseek-v4-flash").strip(),
            "deepseek_base_url": field_value(form, "deepseek_base_url", "https://api.deepseek.com").strip(),
            "ai_mode": field_value(form, "ai_mode", "fast").strip(),
            "period": field_value(form, "period", "").strip(),
            "target": field_value(form, "target", "").strip(),
            "session_key": field_value(form, "session_key", "").strip(),
            "conversation_id": field_value(form, "conversation_id", "").strip(),
            "user_id": field_value(form, "user_id", "").strip(),
            "store_name": field_value(form, "store_name", "").strip(),
            "action": field_value(form, "action", "").strip(),
            "report_mode": field_value(form, "report_mode", "").strip(),
        }
        run_name = datetime.now().strftime("%Y%m%d_%H%M%S")
        upload_dir = OUTPUT_ROOT / f"qclaw_{run_name}" / "uploads"
        files = save_uploads(form, "files", upload_dir, allowed_suffixes=QCLAW_UPLOAD_SUFFIXES)
        files.extend(save_uploads(form, "reports", upload_dir, allowed_suffixes=QCLAW_UPLOAD_SUFFIXES))
        return self.run_qclaw_request(text, files, settings)

    def run_qclaw_request(self, text: str, files: list[Path], settings: dict[str, Any]) -> dict[str, Any]:
        cleanup_daily_source_cache()
        period = requested_period_from_text_or_settings(text, settings)
        target = str(settings.get("target") or "").strip() or None
        selected_store = selected_store_name(settings)
        files = expand_archives(
            files,
            OUTPUT_ROOT / "qclaw_unzipped" / datetime.now().strftime("%Y%m%d_%H%M%S_%f"),
            QCLAW_EXTRACTABLE_SUFFIXES,
        )
        report_files = [path.resolve() for path in files if path.suffix.lower() in QCLAW_REPORT_SUFFIXES and path.exists()]
        image_files = [path.resolve() for path in files if path.suffix.lower() in IMAGE_SUFFIXES and path.exists()]
        screenshot_dir = OUTPUT_ROOT / "qclaw_screenshot_inputs" / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        screenshot_audit: list[dict[str, Any]] = []
        if image_files:
            ai_config = qclaw_ai_config(settings)
            recognition = extract_screenshot_report_files(image_files, screenshot_dir, ai_config)
            report_files = unique_paths([*report_files, *recognition.report_files])
            screenshot_audit = recognition.audits
        document_files = [
            path.resolve()
            for path in files
            if path.suffix.lower() not in QCLAW_REPORT_SUFFIXES and path.suffix.lower() not in IMAGE_SUFFIXES and path.exists()
        ]
        archive_result = archive_daily_report_files(report_files, selected_store)
        document_audit = parse_document_files(document_files, OUTPUT_ROOT / "qclaw_file2md")
        route = classify_qclaw_intent(text, report_files, document_files)
        if route["intent"] == "awaiting_report_request":
            return {
                "type": "needs_input",
                "text": "已收到可用数据文件。请直接说“生成月报”“生成半月报”或“生成周报”，我会立即生成对应的最终 Excel。",
                "files": [],
                "audit": {
                    "intent": "awaiting_report_request",
                    "route": route,
                    "report_files": [str(path) for path in report_files],
                    "screenshot_recognition": screenshot_audit,
                    "daily_store_cache": archive_result_summary(archive_result),
                },
            }
        wants_report = route["intent"] == "report_generation"
        session_key = report_session_key(settings, self.client_address[0] if self.client_address else "") if wants_report else ""
        generation_token = start_generation_request(session_key) if wants_report else ""
        if wants_report and not report_files:
            last_report = get_last_report(session_key)
            if last_report and last_report.source_files and wants_cached_generation(text):
                cached_period = period or last_report.period or "month"
                result = self.generate_report_from_collected_files(
                    list(last_report.source_files),
                    document_files,
                    cached_period,
                    target or last_report.target,
                    settings,
                )
                if not generation_is_current(session_key, generation_token):
                    return discard_obsolete_generation(result)
                result.setdefault("audit", {})["reused_last_source_cache"] = True
                return result
            return {
                "type": "needs_input",
                "text": "可以生成周期报表。请上传万相台或生意参谋的 csv/xlsx 原始报表，文件名不限，新旧版任选其一；最终会返回 xlsx。若想复用上一次数据，也可以说“用刚才的数据生成半月报/月报”。",
                "files": [],
                "audit": {
                    "intent": "report_generation",
                    "route": route,
                    "missing": ["csv_or_xlsx_report_files"],
                    "document_files_received": [str(path) for path in document_files],
                    "document_parse": document_audit,
                    "screenshot_recognition": screenshot_audit,
                },
            }
        if wants_report and report_files:
            periods = requested_periods_from_text_or_settings(text, settings)
            results: list[dict[str, Any]] = []
            cache_summary = archive_result_summary(archive_result)
            for item_period in periods:
                if not generation_is_current(session_key, generation_token):
                    return obsolete_generation_response()
                item_report_files, store_choice = resolve_daily_store_files_for_generation(report_files, selected_store, settings, item_period, target)
                if store_choice:
                    return store_choice
                run_name = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_dir = OUTPUT_ROOT / f"qclaw_{item_period}_{run_name}"
                output_dir.mkdir(parents=True, exist_ok=True)
                template = first_existing(DEFAULT_TEMPLATE_CANDIDATES)
                ai_config = qclaw_ai_config(settings)
                raw_result = run_workflow(item_report_files, template, build_output_paths(output_dir), target, item_period, ai_config=ai_config)
                if not generation_is_current(session_key, generation_token):
                    return discard_obsolete_generation(output_dir=output_dir)
                result = build_enterprise_result(raw_result, item_report_files, output_dir, item_period, target)
                if not generation_is_current(session_key, generation_token):
                    return discard_obsolete_generation(result, output_dir)
                results.append(result)
            final_result = results[0] if len(results) == 1 else combine_period_results(results, report_files)
            if not generation_is_current(session_key, generation_token):
                return discard_obsolete_generation(final_result)
            response = qclaw_report_response(final_result, document_files, document_audit)
            final_result["final_reports"] = [item.get("path") for item in response.get("files", []) if item.get("kind") == "final_report"]
            store_last_report(session_key, final_result, report_files, periods[-1], target)
            for item in results:
                cleanup_transient_run_artifacts(item)
            shutil.rmtree(screenshot_dir, ignore_errors=True)
            response.setdefault("audit", {}).update(
                {
                    "daily_store_cache": cache_summary,
                    "screenshot_recognition": screenshot_audit,
                }
            )
            return response
        if document_files:
            return {
                "type": "needs_input",
                "text": "我已收到文档/图片类文件并尝试读取。若需要生成报表，请再上传万相台或生意参谋的 csv/xlsx 原始报表；文件名不需要固定。",
                "files": [],
                "audit": {
                    "intent": "file_read",
                    "route": route,
                    "document_files": [str(path) for path in document_files],
                    "document_parse": document_audit,
                    "screenshot_recognition": screenshot_audit,
                    "report_files": [str(path) for path in report_files],
                    "next_step": "upload_report_csv_or_xlsx_for_final_xlsx",
                },
            }
        return self.run_qclaw_chat(text, settings)

    def handle_report_collection(
        self,
        text: str,
        report_files: list[Path],
        document_files: list[Path],
        period: str,
        target: str | None,
        settings: dict[str, Any],
    ) -> dict[str, Any]:
        session_key = report_session_key(settings, self.client_address[0] if self.client_address else "")
        action = normalize_report_action(text, settings)
        selected_store = selected_store_name(settings)
        if action == "cancel":
            with REPORT_SESSIONS_LOCK:
                REPORT_SESSIONS.pop(session_key, None)
            return report_collection_response(
                "已取消本次月报文件收集。需要重新生成时，请再次进入月报板块。",
                session_key,
                [],
                status="cancelled",
            )

        if action == "supplement":
            with REPORT_SESSIONS_LOCK:
                active_session = REPORT_SESSIONS.get(session_key)
                pending_files = list(active_session.files or []) if active_session else []
            last_report = get_last_report(session_key)
            if not last_report or not last_report.source_files:
                return report_collection_response(
                    "没有找到可补充的上一份月报记录。请重新上传本次要生成的全部文件。",
                    session_key,
                    report_files,
                    status="waiting_files",
                )
            combined_files = unique_paths([*(last_report.source_files or []), *pending_files, *report_files])
            result_period = period or last_report.period or "month"
            result = self.generate_report_from_collected_files(combined_files, document_files, result_period, last_report.target or target, settings)
            with REPORT_SESSIONS_LOCK:
                REPORT_SESSIONS.pop(session_key, None)
            result.setdefault("audit", {})["supplemented_previous_report"] = True
            return result

        if action == "new":
            with REPORT_SESSIONS_LOCK:
                session = ReportCollectionSession(session_key=session_key, period=period or "month", target=target, store_name=selected_store, files=list(report_files))
                REPORT_SESSIONS[session_key] = session
                collected_files = list(session.files or [])
            if report_files and wants_report_generation(text):
                result = self.generate_report_from_collected_files(collected_files, document_files, session.period, target, settings)
                with REPORT_SESSIONS_LOCK:
                    REPORT_SESSIONS.pop(session_key, None)
                return result
            return report_collection_response(
                f"已按新月报开始收集，当前共 {len(collected_files)} 个可用表格。{describe_collected_files(collected_files)}还有文件可以继续上传；全部发完后点“确认生成”。",
                session_key,
                collected_files,
                status="collecting",
            )

        with REPORT_SESSIONS_LOCK:
            session = REPORT_SESSIONS.get(session_key)
            if session is None:
                session = ReportCollectionSession(session_key=session_key, period=period or "month", target=target)
                REPORT_SESSIONS[session_key] = session
            if period:
                session.period = period
            session.target = target or session.target
            session.store_name = selected_store or session.store_name
            if report_files:
                existing = {str(path) for path in session.files or []}
                for path in report_files:
                    if str(path) not in existing:
                        session.files.append(path)
                        existing.add(str(path))
            session.updated_at = datetime.now().isoformat(timespec="seconds")
            collected_files = list(session.files or [])
            selected_store = session.store_name or selected_store

        if selected_store:
            daily_files = daily_cached_files_for_store(selected_store)
            if daily_files:
                collected_files = unique_paths([*daily_files, *collected_files])
                with REPORT_SESSIONS_LOCK:
                    session = REPORT_SESSIONS.get(session_key)
                    if session:
                        session.store_name = selected_store
                        session.files = collected_files
                        session.updated_at = datetime.now().isoformat(timespec="seconds")

        if action == "confirm":
            if not selected_store:
                store_names = daily_store_names()
                if len(store_names) > 1:
                    return store_choice_response(
                        "今天已缓存多个店铺的数据。请选择要生成哪个店铺的报表，我会调用该店铺今天上传过的所有文件。",
                        session_key,
                        store_names,
                        session.period,
                        target,
                    )
                if len(store_names) == 1:
                    selected_store = store_names[0]
                    collected_files = unique_paths([*daily_cached_files_for_store(selected_store), *collected_files])
            if not collected_files:
                return report_collection_response(
                    "还没有收到可用于生成报表的 csv/xlsx 原始表。请继续上传文件，发完后点“确认生成”。",
                    session_key,
                    collected_files,
                    status="waiting_files",
                )
            result = self.generate_report_from_collected_files(collected_files, document_files, session.period, session.target, settings)
            with REPORT_SESSIONS_LOCK:
                REPORT_SESSIONS.pop(session_key, None)
            return result

        if action == "enter":
            return report_collection_response(
                "已进入报表生成板块。请把本次需要生成周期报表的 csv/xlsx 原始表都发到这里；文件名不限，日报/月报、范围日期、新旧版报表都可以识别。发完后点“确认生成”，我再开始生成，避免文件没发完就误生成。",
                session_key,
                collected_files,
                status="collecting",
            )

        if report_files:
            last_report = get_last_report(session_key)
            if last_report:
                return supplement_choice_response(session_key, collected_files, last_report, len(report_files))
            if not selected_store:
                store_names = daily_store_names()
                if len(store_names) > 1:
                    return store_choice_response(
                        f"已收到 {len(report_files)} 个新文件，并已按店铺归档。今天有多个店铺缓存，请选择这批文件属于哪个店铺或要继续操作哪个店铺。",
                        session_key,
                        store_names,
                        session.period,
                        target,
                    )
            return report_collection_response(
                f"已收到 {len(report_files)} 个新文件，当前共 {len(collected_files)} 个可用表格。{selected_store_prefix(selected_store)}{describe_collected_files(collected_files)}还有文件可以继续上传；全部发完后点“确认生成”。",
                session_key,
                collected_files,
                status="collecting",
            )

        return report_collection_response(
            f"当前月报板块已暂存 {len(collected_files)} 个可用表格。{selected_store_prefix(selected_store)}{describe_collected_files(collected_files)}请继续上传，或点“确认生成”。",
            session_key,
            collected_files,
            status="collecting",
        )

    def generate_report_from_collected_files(
        self,
        report_files: list[Path],
        document_files: list[Path],
        period: str,
        target: str | None,
        settings: dict[str, Any],
    ) -> dict[str, Any]:
        document_audit = parse_document_files(document_files, OUTPUT_ROOT / "qclaw_file2md")
        run_name = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = OUTPUT_ROOT / f"qclaw_{period}_{run_name}"
        output_dir.mkdir(parents=True, exist_ok=True)
        template = first_existing(DEFAULT_TEMPLATE_CANDIDATES)
        ai_config = AIConfig(
            provider="qclaw",
            api_key="",
            base_url=os.environ.get("QCLAW_LLM_BASE_URL", "http://127.0.0.1:19100/proxy/llm"),
            model=os.environ.get("QCLAW_LLM_MODEL", "modelroute"),
            mode=str(settings.get("ai_mode") or os.environ.get("WXT_AI_MODE", "fast")),
        )
        raw_result = run_workflow(report_files, template, build_output_paths(output_dir), target, period, ai_config=ai_config)
        session_key = report_session_key(settings, self.client_address[0] if self.client_address else "")
        current_token = ACTIVE_GENERATION_TOKENS.get(session_key, "")
        if current_token and not generation_is_current(session_key, current_token):
            return discard_obsolete_generation(output_dir=output_dir)
        result = build_enterprise_result(raw_result, report_files, output_dir, period, target)
        response = qclaw_report_response(result, document_files, document_audit)
        result["final_reports"] = [item.get("path") for item in response.get("files", []) if item.get("kind") == "final_report"]
        if current_token and not generation_is_current(session_key, current_token):
            return discard_obsolete_generation(result, output_dir)
        store_last_report(session_key, result, report_files, period, target)
        cleanup_transient_run_artifacts(result)
        response.setdefault("audit", {})["collection_mode"] = True
        return response

    def run_qclaw_chat(self, text: str, settings: dict[str, Any]) -> dict[str, Any]:
        api_key = str(settings.get("deepseek_api_key") or os.environ.get("DEEPSEEK_API_KEY", "")).strip()
        if not api_key:
            return {
                "type": "needs_input",
                "text": "请先提供 DeepSeek API Key，或让管理员在服务端配置 DEEPSEEK_API_KEY。",
                "files": [],
                "audit": {"intent": "chat", "missing": ["deepseek_api_key"]},
            }
        answer = chat_deepseek(
            [{"role": "user", "content": text or "请介绍一下万相台诊断师可以做什么。"}],
            AIConfig(
                provider="deepseek",
                api_key=api_key,
                base_url=str(settings.get("deepseek_base_url") or os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com")),
                model=str(settings.get("deepseek_model") or os.environ.get("DEEPSEEK_MODEL", "deepseek-v4-flash")),
            ),
        )
        return {
            "type": "message",
            "text": answer,
            "files": [],
            "audit": {"intent": "chat", "used_tools": ["deepseek"]},
        }

    def handle_download(self, query: str) -> None:
        params = parse_qs(query)
        raw_path = params.get("path", [""])[0]
        path = Path(unquote(raw_path)).resolve()
        root = OUTPUT_ROOT.resolve()
        try:
            path.relative_to(root)
        except ValueError:
            self.send_error(HTTPStatus.FORBIDDEN, "不能下载输出目录外文件")
            return
        if not path.exists() or not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND, "文件不存在")
            return
        content = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type_for(path))
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(path.name)}")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def parse_form(self) -> dict[str, list[FormPart]]:
        content_type = self.headers.get("Content-Type", "")
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length)
        if not content_type.startswith("multipart/form-data"):
            return {}
        raw = (
            f"Content-Type: {content_type}\r\n"
            "MIME-Version: 1.0\r\n"
            "\r\n"
        ).encode("utf-8") + body
        message = BytesParser(policy=policy.default).parsebytes(raw)
        form: dict[str, list[FormPart]] = {}
        for part in message.iter_parts():
            name = part.get_param("name", header="content-disposition")
            if not name:
                continue
            filename = part.get_filename()
            payload = part.get_payload(decode=True) or b""
            if filename:
                item = FormPart(name=name, filename=filename, data=payload)
            else:
                charset = part.get_content_charset() or "utf-8"
                item = FormPart(name=name, value=payload.decode(charset, errors="replace"))
            form.setdefault(name, []).append(item)
        return form

    def send_html(self, body: str, status: HTTPStatus = HTTPStatus.OK) -> None:
        data = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, payload: dict[str, Any], status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def field_value(form: dict[str, list[FormPart]], name: str, default: str = "") -> str:
    if name not in form:
        return default
    value = form[name][0]
    if value.value is None:
        return default
    return str(value.value)


def save_uploads(
    form: dict[str, list[FormPart]],
    name: str,
    target_dir: Path,
    allowed_suffixes: set[str] | None = None,
) -> list[Path]:
    if name not in form:
        return []
    fields = form[name]
    target_dir.mkdir(parents=True, exist_ok=True)
    saved: list[Path] = []
    for field in fields:
        filename = safe_filename(field.filename)
        if not filename:
            continue
        suffix = Path(filename).suffix.lower()
        if suffix not in (allowed_suffixes or SUPPORTED_SUFFIXES):
            continue
        destination = unique_path(target_dir / filename)
        with destination.open("wb") as out:
            out.write(field.data)
        saved.append(destination.resolve())
    return saved


def expand_archives(files: list[Path], target_dir: Path, allowed_suffixes: set[str]) -> list[Path]:
    expanded: list[Path] = []
    seen_sources: set[Path] = set()
    for path in files:
        source = Path(path).expanduser()
        if not source.exists():
            continue
        resolved = source.resolve()
        if resolved in seen_sources:
            continue
        seen_sources.add(resolved)
        if source.suffix.lower() == ".zip" and source.exists():
            expanded.extend(expand_zip_archive(source, target_dir / source.stem, allowed_suffixes))
        else:
            expanded.append(resolved)
    return unique_paths(expanded)


def expand_zip_archive(zip_path: Path, target_dir: Path, allowed_suffixes: set[str]) -> list[Path]:
    return extract_zip_archive(zip_path, target_dir, allowed_suffixes)


def safe_filename(filename: str | None) -> str:
    return sanitize_windows_filename(filename)


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for index in range(1, 1000):
        candidate = path.with_name(f"{stem}_{index}{suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"无法保存上传文件: {path.name}")


def content_type_for(path: Path) -> str:
    if path.suffix.lower() == ".xlsx":
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if path.suffix.lower() == ".json":
        return "application/json; charset=utf-8"
    if path.suffix.lower() == ".md":
        return "text/markdown; charset=utf-8"
    return "application/octet-stream"


def report_session_key(settings: dict[str, Any], fallback: str = "") -> str:
    for key in ("session_key", "conversation_id", "conversationId", "chat_id", "chatId", "user_id", "userId", "sender_id", "senderId"):
        value = str(settings.get(key) or "").strip()
        if value:
            return value
    return fallback or "default"


def should_prompt_for_supplement(text: str, report_files: list[Path], settings: dict[str, Any]) -> bool:
    if not report_files:
        return False
    action = normalize_report_action(text, settings)
    if action in {"supplement", "new"}:
        return True
    session_key = report_session_key(settings)
    with REPORT_SESSIONS_LOCK:
        has_active_session = session_key in REPORT_SESSIONS
    last_report = get_last_report(session_key)
    return bool(last_report and not has_active_session and action == "collect")


def wants_cached_generation(text: str) -> bool:
    stripped = text.strip()
    return any(keyword in stripped for keyword in ("刚才", "上次", "之前", "上一份", "同一批", "这些数据", "原来的数据"))


def selected_store_name(settings: dict[str, Any]) -> str:
    for key in ("store_name", "storeName", "shop_name", "shopName", "店铺名称", "店铺"):
        value = str(settings.get(key) or "").strip()
        if value:
            return value
    return ""


def today_cache_root() -> Path:
    return DAILY_SOURCE_CACHE_DIR / date.today().isoformat()


def cleanup_daily_source_cache() -> None:
    root = DAILY_SOURCE_CACHE_DIR
    if not root.exists():
        return
    today_name = date.today().isoformat()
    for child in root.iterdir():
        try:
            if child.is_dir() and child.name != today_name:
                shutil.rmtree(child)
            elif child.is_file() and child.name != today_name:
                child.unlink()
        except Exception:
            continue


def archive_daily_report_files(report_files: list[Path], preferred_store: str = "") -> DailyArchiveResult:
    files = unique_files_by_content(report_files)
    if not files:
        return DailyArchiveResult(stores={}, unknown_files=[])
    inferred = {path: (preferred_store or infer_store_name_from_report(path)) for path in files}
    known_stores = sorted({store for store in inferred.values() if store and store != "未识别店铺"})
    if not preferred_store and len(known_stores) == 1:
        inferred = {path: (store if store and store != "未识别店铺" else known_stores[0]) for path, store in inferred.items()}

    stores: dict[str, list[Path]] = {}
    unknown: list[Path] = []
    for source, store in inferred.items():
        store = store or "未识别店铺"
        if store == "未识别店铺":
            unknown.append(source)
        cache_dir = today_cache_root() / safe_cache_name(store)
        cache_dir.mkdir(parents=True, exist_ok=True)
        existing = daily_cached_files_for_store(store)
        target = matching_content_path(source, existing)
        if target is None:
            target = unique_path(cache_dir / source.name)
            shutil.copy2(source, target)
        stores.setdefault(store, []).append(target.resolve())
    write_daily_cache_index(stores)
    return DailyArchiveResult(stores=stores, unknown_files=unknown)


def write_daily_cache_index(stores: dict[str, list[Path]]) -> None:
    for store, paths in stores.items():
        cache_dir = today_cache_root() / safe_cache_name(store)
        metadata_path = cache_dir / "cache.json"
        existing: dict[str, Any] = {}
        if metadata_path.exists():
            try:
                existing = json.loads(metadata_path.read_text(encoding="utf-8"))
            except Exception:
                existing = {}
        existing_paths = [Path(path) for path in existing.get("files", []) if Path(path).exists()]
        all_paths = unique_files_by_content([*existing_paths, *paths])
        payload = {
            "store_name": store,
            "cache_date": date.today().isoformat(),
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "files": [str(path) for path in all_paths],
        }
        metadata_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def daily_store_names() -> list[str]:
    cleanup_daily_source_cache()
    root = today_cache_root()
    if not root.exists():
        return []
    stores: list[str] = []
    for child in sorted(root.iterdir()):
        if not child.is_dir():
            continue
        metadata = child / "cache.json"
        if metadata.exists():
            try:
                payload = json.loads(metadata.read_text(encoding="utf-8"))
                name = str(payload.get("store_name") or "").strip()
            except Exception:
                name = ""
        else:
            name = ""
        stores.append(name or child.name)
    return [store for store in stores if store]


def daily_cached_files_for_store(store_name: str) -> list[Path]:
    if not store_name:
        return []
    cleanup_daily_source_cache()
    cache_dir = today_cache_root() / safe_cache_name(store_name)
    if not cache_dir.exists():
        return []
    metadata = cache_dir / "cache.json"
    if metadata.exists():
        try:
            payload = json.loads(metadata.read_text(encoding="utf-8"))
            return unique_files_by_content([Path(path) for path in payload.get("files", [])])
        except Exception:
            pass
    return unique_files_by_content([path for path in cache_dir.iterdir() if path.suffix.lower() in QCLAW_REPORT_SUFFIXES])


def infer_store_name_from_report(path: Path) -> str:
    suffix = path.suffix.lower()
    try:
        if suffix in {".xlsx", ".xls"}:
            return infer_store_name_from_excel(path)
        if suffix == ".csv":
            return infer_store_name_from_csv(path)
    except Exception:
        return infer_store_name_from_filename(path)
    return infer_store_name_from_filename(path)


def infer_store_name_from_excel(path: Path) -> str:
    excel = pd.ExcelFile(path)
    for sheet_name in excel.sheet_names[:5]:
        frame = excel.parse(sheet_name, nrows=80)
        store = infer_store_name_from_frame(frame)
        if store:
            return store
    return infer_store_name_from_filename(path)


def infer_store_name_from_csv(path: Path) -> str:
    for encoding in ("gb18030", "utf-8-sig", "utf-8"):
        try:
            frame = pd.read_csv(path, encoding=encoding, nrows=80)
            store = infer_store_name_from_frame(frame)
            if store:
                return store
            break
        except Exception:
            continue
    return infer_store_name_from_filename(path)


def infer_store_name_from_frame(frame: pd.DataFrame) -> str:
    if frame.empty:
        return ""
    columns = [str(column).strip() for column in frame.columns]
    frame = frame.copy()
    frame.columns = columns
    for column in ("店铺名称", "店铺", "店铺名"):
        if column in frame.columns:
            values = frame[column].dropna().astype(str).str.strip()
            values = values[values.ne("")]
            if not values.empty:
                return values.iloc[0]
    # Some exports place the store name in the first rows instead of a dedicated column.
    for value in frame.head(10).to_numpy().flatten().tolist():
        text = str(value).strip()
        if not text or text.lower() == "nan":
            continue
        if "店铺" in text and len(text) <= 80:
            for sep in ("：", ":", " "):
                if sep in text:
                    candidate = text.split(sep)[-1].strip()
                    if candidate:
                        return candidate
            return text
    return ""


def infer_store_name_from_filename(path: Path) -> str:
    stem = path.stem
    for token in ("店铺经营核心月报", "店铺流量来源构成月报", "商品流量来源构成月报", "商品整体效果月报", "商品经营投产比核心日报", "营销场景报表", "关键词报表", "计划报表", "商品报表"):
        if token in stem:
            prefix = stem.split(token, 1)[0].strip("_- ")
            if prefix:
                return prefix
    return "未识别店铺"


def resolve_daily_store_files_for_generation(
    report_files: list[Path],
    selected_store: str,
    settings: dict[str, Any],
    period: str,
    target: str | None,
) -> tuple[list[Path], dict[str, Any] | None]:
    session_key = report_session_key(settings)
    if selected_store:
        cached = daily_cached_files_for_store(selected_store)
        return (unique_files_by_content([*report_files, *cached]) or report_files, None)
    stores = daily_store_names()
    if len(stores) == 1:
        cached = daily_cached_files_for_store(stores[0])
        return (unique_files_by_content([*report_files, *cached]) or report_files, None)
    if len(stores) > 1:
        return report_files, store_choice_response(
            "今天已缓存多个店铺的数据。请选择要生成哪个店铺的报表，我会调用该店铺今天上传过的所有文件。",
            session_key,
            stores,
            period,
            target,
        )
    return report_files, None


def store_choice_response(text: str, session_key: str, store_names: list[str], period: str, target: str | None) -> dict[str, Any]:
    unique_stores = []
    for store in store_names:
        if store not in unique_stores:
            unique_stores.append(store)
    return {
        "type": "report_collection",
        "text": text,
        "files": [],
        "audit": {
            "intent": "report_collection",
            "status": "awaiting_store_choice",
            "session_key": session_key,
            "daily_cache_date": date.today().isoformat(),
            "stores": unique_stores,
            "used_tools": ["daily_store_cache"],
        },
        "actions": [
            {
                "id": f"select_store_{idx}",
                "label": store,
                "type": "submit",
                "settings": {
                    "action": "select_store",
                    "report_mode": "1",
                    "session_key": session_key,
                    "store_name": store,
                    "period": period,
                    "target": target or "",
                },
            }
            for idx, store in enumerate(unique_stores, start=1)
        ],
    }


def selected_store_prefix(store_name: str) -> str:
    return f"当前店铺：{store_name}。" if store_name else ""


def archive_result_summary(result: DailyArchiveResult) -> dict[str, Any]:
    return {
        "cache_date": date.today().isoformat(),
        "stores": {store: [str(path) for path in paths] for store, paths in result.stores.items()},
        "unknown_files": [str(path) for path in result.unknown_files],
    }


def get_last_report(session_key: str) -> LastReportSession | None:
    with REPORT_SESSIONS_LOCK:
        last_report = LAST_REPORTS.get(session_key)
    if last_report:
        return last_report
    return load_last_report_metadata(session_key)


def store_last_report(session_key: str, result: dict[str, Any], report_files: list[Path], period: str, target: str | None) -> None:
    if result.get("status") != "ok":
        return
    final_reports = [Path(path) for path in result.get("final_reports", [])]
    recommended_missing = list(result.get("recommended_missing") or [])
    cached_source_files = persist_source_files(session_key, report_files)
    last_report = LastReportSession(
        session_key=session_key,
        period=period,
        target=target,
        source_files=cached_source_files,
        final_reports=final_reports,
        recommended_missing=recommended_missing,
    )
    save_last_report_metadata(last_report)
    with REPORT_SESSIONS_LOCK:
        LAST_REPORTS[session_key] = last_report


def safe_cache_name(value: str) -> str:
    name = "".join("_" if char in r'\/:*?"<>| ' else char for char in value).strip("._")
    return name[:80] or "default"


def persist_source_files(session_key: str, report_files: list[Path]) -> list[Path]:
    existing_files = unique_paths(report_files)
    if not existing_files:
        return []
    cache_dir = SOURCE_CACHE_DIR / safe_cache_name(session_key)
    payloads: list[tuple[str, bytes]] = []
    for source in existing_files:
        payloads.append((source.name, source.read_bytes()))
    if cache_dir.exists():
        shutil.rmtree(cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)
    cached: list[Path] = []
    for name, data in payloads:
        target = unique_path(cache_dir / name)
        target.write_bytes(data)
        cached.append(target.resolve())
    return cached


def save_last_report_metadata(last_report: LastReportSession) -> None:
    cache_dir = SOURCE_CACHE_DIR / safe_cache_name(last_report.session_key)
    cache_dir.mkdir(parents=True, exist_ok=True)
    payload = {
        "session_key": last_report.session_key,
        "period": last_report.period,
        "target": last_report.target,
        "source_files": [str(path) for path in last_report.source_files or []],
        "final_reports": [str(path) for path in last_report.final_reports or []],
        "recommended_missing": list(last_report.recommended_missing or []),
        "updated_at": last_report.updated_at,
    }
    (cache_dir / "last_report.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_last_report_metadata(session_key: str) -> LastReportSession | None:
    metadata_path = SOURCE_CACHE_DIR / safe_cache_name(session_key) / "last_report.json"
    if not metadata_path.exists():
        return None
    try:
        payload = json.loads(metadata_path.read_text(encoding="utf-8"))
        source_files = [Path(path) for path in payload.get("source_files", []) if Path(path).exists()]
        if not source_files:
            return None
        return LastReportSession(
            session_key=str(payload.get("session_key") or session_key),
            period=str(payload.get("period") or "month"),
            target=payload.get("target"),
            source_files=source_files,
            final_reports=[Path(path) for path in payload.get("final_reports", [])],
            recommended_missing=list(payload.get("recommended_missing") or []),
            updated_at=str(payload.get("updated_at") or ""),
        )
    except Exception:
        return None


def unique_paths(paths: list[Path]) -> list[Path]:
    unique: list[Path] = []
    seen: set[str] = set()
    for path in paths:
        resolved = Path(path).resolve()
        key = str(resolved)
        if key not in seen and resolved.exists():
            unique.append(resolved)
            seen.add(key)
    return unique


def unique_files_by_content(paths: list[Path]) -> list[Path]:
    unique: list[Path] = []
    seen_fingerprints: set[str] = set()
    for path in unique_paths(paths):
        fingerprint = file_content_fingerprint(path)
        if fingerprint in seen_fingerprints:
            continue
        unique.append(path)
        seen_fingerprints.add(fingerprint)
    return unique


def matching_content_path(source: Path, candidates: list[Path]) -> Path | None:
    source_fingerprint = file_content_fingerprint(source)
    for candidate in unique_paths(candidates):
        if file_content_fingerprint(candidate) == source_fingerprint:
            return candidate
    return None


def file_content_fingerprint(path: Path) -> str:
    resolved = Path(path).resolve()
    try:
        stat = resolved.stat()
        digest = hashlib.sha256()
        with resolved.open("rb") as source:
            for chunk in iter(lambda: source.read(1024 * 1024), b""):
                digest.update(chunk)
        return f"{stat.st_size}:{digest.hexdigest()}"
    except OSError:
        return f"path:{resolved}"


def cleanup_old_output_runs(max_age_days: int | None = None) -> None:
    max_age_days = TEMP_OUTPUT_MAX_AGE_DAYS if max_age_days is None else max_age_days
    if max_age_days <= 0:
        return
    root = OUTPUT_ROOT
    if not root.exists() or not root.is_dir():
        return
    cutoff = time.time() - max_age_days * 24 * 60 * 60
    for child in root.iterdir():
        try:
            if not child.is_dir():
                continue
            if child.resolve() in {FINAL_REPORT_DIR.resolve(), SOURCE_CACHE_DIR.resolve()}:
                continue
            if child.stat().st_mtime < cutoff:
                shutil.rmtree(child)
        except Exception:
            continue


def cleanup_transient_run_artifacts(result: dict[str, Any]) -> None:
    if truthy(os.environ.get("WXT_KEEP_RUN_ARTIFACTS")):
        return
    output_dir_text = str(result.get("output_dir") or "").strip()
    if not output_dir_text:
        return
    try:
        output_dir = Path(output_dir_text).resolve()
        output_root = OUTPUT_ROOT.resolve()
        if output_dir == output_root or output_root not in output_dir.parents:
            return
        shutil.rmtree(output_dir, ignore_errors=True)
        result["transient_artifacts_deleted"] = True
    except Exception:
        result["transient_artifacts_deleted"] = False


def truthy(value: Any) -> bool:
    text = str(value or "").strip().lower()
    return text in {"1", "true", "yes", "y", "on", "月报", "report", "report_mode"}


def normalize_report_action(text: str, settings: dict[str, Any]) -> str:
    action = str(settings.get("action") or "").strip().lower()
    if action in {"enter_report", "enter_report_board", "start_report", "start_collection", "report_enter"}:
        return "enter"
    if action in {"confirm_report", "generate_report", "finish_upload", "done", "confirm"}:
        return "confirm"
    if action in {"cancel_report", "cancel_collection", "cancel"}:
        return "cancel"
    if action in {"supplement_report", "append_report", "补充上一份", "补充"}:
        return "supplement"
    if action in {"new_report", "start_new_report", "生成新月报", "新月报"}:
        return "new"
    if action in {"select_store", "choose_store", "store_select"}:
        return "select_store"
    stripped = text.strip()
    if any(keyword in stripped for keyword in ("进入月报板块", "月报板块", "开始月报", "我要生成月报", "收集月报文件")):
        return "enter"
    if any(keyword in stripped for keyword in ("发完了", "发送完了", "上传完了", "确认生成", "开始生成", "现在生成", "生成吧", "生成月报")):
        return "confirm"
    if any(keyword in stripped for keyword in ("取消月报", "取消生成", "不生成了", "清空月报", "重新开始")):
        return "cancel"
    if any(keyword in stripped for keyword in ("补充上一份", "补充之前", "补充到之前", "补充上次", "追加到上次")):
        return "supplement"
    if any(keyword in stripped for keyword in ("生成新月报", "新月报", "重新生成一份", "另生成")):
        return "new"
    return "collect"


def should_use_report_collection(text: str, settings: dict[str, Any]) -> bool:
    session_key = report_session_key(settings)
    with REPORT_SESSIONS_LOCK:
        has_session = session_key in REPORT_SESSIONS
    return (
        truthy(settings.get("report_mode"))
        or bool(str(settings.get("action") or "").strip())
        or has_session
        or normalize_report_action(text, settings) in {"enter", "confirm", "cancel", "supplement", "new", "select_store"}
    )


def report_collection_response(
    text: str,
    session_key: str,
    files: list[Path],
    status: str = "collecting",
) -> dict[str, Any]:
    action_settings = report_action_settings(session_key)
    response = {
        "type": "report_collection",
        "text": text,
        "files": [],
        "audit": {
            "intent": "report_collection",
            "status": status,
            "session_key": session_key,
            "collected_count": len(files),
            "collected_files": [str(path) for path in files],
            "used_tools": ["report_collection_router"],
        },
    }
    if status != "cancelled":
        response["actions"] = [
            {"id": "continue_upload", "label": "继续上传", "type": "hint"},
            {"id": "confirm_report", "label": "确认生成", "type": "submit", "settings": {**action_settings, "action": "confirm_report"}},
            {"id": "cancel_report", "label": "取消", "type": "submit", "settings": {**action_settings, "action": "cancel_report"}},
        ]
    return response


def report_action_settings(session_key: str) -> dict[str, str]:
    settings = {"report_mode": "1", "session_key": session_key}
    with REPORT_SESSIONS_LOCK:
        session = REPORT_SESSIONS.get(session_key)
    if session:
        if session.period:
            settings["period"] = session.period
        if session.target:
            settings["target"] = session.target
        if session.store_name:
            settings["store_name"] = session.store_name
    return settings


def supplement_choice_response(session_key: str, files: list[Path], last_report: LastReportSession, new_count: int) -> dict[str, Any]:
    missing_preview = "、".join((last_report.recommended_missing or [])[:3])
    if len(last_report.recommended_missing or []) > 3:
        missing_preview += "等"
    text = (
        f"已收到 {new_count} 个新文件。上一份月报是简版，仍有建议补充项：{missing_preview or '无'}。\n"
        "如果这是同一个店铺、同一个周期的补充数据，建议选“补充上一份月报”；如果是新店铺或新周期，请选“生成新月报”。"
    )
    return {
        "type": "report_collection",
        "text": text,
        "files": [],
        "audit": {
            "intent": "report_collection",
            "status": "awaiting_supplement_choice",
            "session_key": session_key,
            "collected_count": len(files),
            "collected_files": [str(path) for path in files],
            "last_report_files": [str(path) for path in last_report.final_reports or []],
            "recommended_missing": last_report.recommended_missing or [],
            "used_tools": ["report_collection_router"],
        },
        "actions": [
            {"id": "supplement_report", "label": "补充上一份月报", "type": "submit", "settings": {"action": "supplement_report", "report_mode": "1", "session_key": session_key, "period": last_report.period}},
            {"id": "new_report", "label": "生成新月报", "type": "submit", "settings": {"action": "new_report", "report_mode": "1", "session_key": session_key}},
            {"id": "cancel_report", "label": "取消", "type": "submit", "settings": {"action": "cancel_report", "report_mode": "1", "session_key": session_key}},
        ],
    }


def describe_collected_files(files: list[Path]) -> str:
    if not files:
        return ""
    labels: list[str] = []
    for path in files:
        name = path.name
        if "店铺经营" in name or "核心" in name:
            labels.append("店铺收益核心数据")
        elif "店铺流量" in name:
            labels.append("店铺流量来源数据")
        elif "商品流量" in name:
            labels.append("商品流量来源数据")
        elif "商品整体" in name or "商品经营" in name:
            labels.append("商品经营数据")
        elif "商品报表" in name:
            labels.append("商品广告数据")
        elif "营销场景" in name or "计划报表" in name:
            labels.append("推广计划/广告场景数据")
        elif "关键词" in name:
            labels.append("关键词投放数据")
    if not labels:
        return ""
    unique = []
    for label in labels:
        if label not in unique:
            unique.append(label)
    return "我初步识别到：" + "、".join(unique[:6]) + "。"


def wants_report_generation(text: str) -> bool:
    lowered = text.lower().replace(" ", "")
    explicit_phrases = [
        "生成报表",
        "制作报表",
        "导出报表",
        "输出报表",
        "生成excel",
        "导出excel",
        "输出excel",
    ]
    if any(phrase in lowered for phrase in explicit_phrases):
        return True
    actions = ("生成", "制作", "做一份", "做个", "导出", "输出")
    periods = ("月报", "半月报", "周报")
    return any(action in lowered for action in actions) and any(period in lowered for period in periods)


def classify_qclaw_intent(text: str, report_files: list[Path], document_files: list[Path]) -> dict[str, Any]:
    stripped = text.strip()
    lowered = stripped.lower()
    if wants_report_generation(stripped):
        return {
            "intent": "report_generation",
            "reason": "explicit_report_keyword",
            "uses_monthly_program": True,
        }
    file_analysis_keywords = [
        "分析",
        "诊断",
        "复盘",
        "看下",
        "看看",
        "解读",
        "优化建议",
        "经营建议",
        "投放建议",
        "总结",
        "生成",
        "制作",
        "整理",
        "输出",
    ]
    pure_chat_keywords = [
        "什么是",
        "介绍",
        "解释",
        "区别",
        "原理",
        "怎么理解",
        "怎么优化",
        "如何优化",
        "为什么",
        "能不能",
        "是否",
        "规则",
        "逻辑",
        "教程",
    ]
    if report_files:
        if not stripped:
            return {
                "intent": "awaiting_report_request",
                "reason": "spreadsheet_uploaded_without_generation_instruction",
                "uses_monthly_program": False,
            }
        if any(keyword in stripped for keyword in pure_chat_keywords):
            return {
                "intent": "chat",
                "reason": "spreadsheet_uploaded_but_question_is_conceptual",
                "uses_monthly_program": False,
            }
        return {
            "intent": "chat",
            "reason": "spreadsheet_uploaded_without_report_intent",
            "uses_monthly_program": False,
        }
    if document_files and any(keyword in stripped for keyword in file_analysis_keywords):
        return {
            "intent": "file_read",
            "reason": "document_uploaded_without_spreadsheet",
            "uses_monthly_program": False,
        }
    return {
        "intent": "chat",
        "reason": "general_wanxiangtai_question",
        "uses_monthly_program": False,
    }


def infer_period_from_text(text: str) -> str:
    return infer_explicit_period_from_text(text) or "month"


def requested_period_from_text_or_settings(text: str, settings: dict[str, Any]) -> str:
    explicit_period = infer_explicit_period_from_text(text)
    raw_period = explicit_period or str(settings.get("period") or "").strip()
    return normalize_period(raw_period) if raw_period else ""


def requested_periods_from_text_or_settings(text: str, settings: dict[str, Any]) -> list[str]:
    raw_setting_value = str(settings.get("period") or "").strip()
    raw_setting = normalize_period(raw_setting_value) if raw_setting_value else ""
    if raw_setting:
        return [raw_setting]
    stripped = text.strip()
    periods: list[str] = []
    text_without_half_month = stripped.replace("半月报", "").replace("半月", "")
    if "月报" in text_without_half_month or "生成月报" in text_without_half_month or "制作月报" in text_without_half_month:
        periods.append("month")
    if "半月报" in stripped or "半月" in stripped:
        periods.append("half-month")
    if "周报" in stripped:
        periods.append("week")
    if not periods:
        periods.append(requested_period_from_text_or_settings(text, settings) or "month")
    unique: list[str] = []
    for period in periods:
        normalized = normalize_period(period)
        if normalized and normalized not in unique:
            unique.append(normalized)
    return unique or ["month"]


def infer_explicit_period_from_text(text: str) -> str | None:
    if "周报" in text:
        return "week"
    if "半月报" in text or "半月" in text:
        return "half-month"
    return None


def parse_document_files(document_files: list[Path], output_root: Path) -> list[dict[str, Any]]:
    if not document_files:
        return []
    file2md_root = Path(os.environ.get("FILE2MD_ROOT", "/Users/gordon/Desktop/file2md")).expanduser()
    script = file2md_root / "scripts" / "parse_document.py"
    audits: list[dict[str, Any]] = []
    for path in document_files:
        item: dict[str, Any] = {"path": str(path), "status": "skipped"}
        if not script.exists():
            item["reason"] = f"file2md not found: {script}"
            audits.append(item)
            continue
        out_dir = output_root / datetime.now().strftime("%Y%m%d_%H%M%S") / safe_stem(path)
        out_dir.mkdir(parents=True, exist_ok=True)
        try:
            completed = subprocess.run(
                [sys.executable, str(script), str(path), "--output", str(out_dir)],
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=90,
            )
            item["status"] = "ok" if completed.returncode == 0 else "failed"
            item["output_dir"] = str(out_dir)
            item["stdout"] = completed.stdout[-1000:]
            item["stderr"] = completed.stderr[-1000:]
        except Exception as exc:  # noqa: BLE001 - document read should not block report generation
            item["status"] = "failed"
            item["reason"] = str(exc)
            item["output_dir"] = str(out_dir)
        audits.append(item)
    return audits


def safe_stem(path: Path) -> str:
    return "".join("_" if char in r'\/:*?"<>|' else char for char in path.stem).strip() or "document"


def qclaw_report_response(result: dict[str, Any], document_files: list[Path], document_audit: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    final_reports = [path for path in result.get("final_reports") or [] if is_allowed_final_report(path)]
    published_reports = [publish_final_report(path) for path in final_reports]
    files = [
        {
            "name": Path(path).name,
            "path": str(path),
            "folder_path": str(Path(path).parent),
            "url": local_file_url(path),
            "folder_url": local_file_url(Path(path).parent),
            "download_url": local_download_url(path),
            "kind": "final_report",
            "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }
        for path in published_reports
    ]
    if files:
        links = "\n".join(f"- [{item['name']}]({item['url']})" for item in files)
        text = "已生成最近可生成周期的最终 Excel 报表，可直接打开：\n" + links
        text += f"\n\n固定保存位置：{FINAL_REPORT_DIR}"
        if result.get("target_period"):
            text += f"\n\n生成周期：{result.get('target_period')}"
        text += "\n\n" + missing_guidance_text(result)
        response_type = "report"
    else:
        blocking = result.get("blocking_missing") or CORE_METRIC_LABELS
        text = result.get("message") or "暂未生成最终报表，请补充以下核心数据：" + "、".join(str(item) for item in blocking)
        response_type = "needs_input"
    return {
        "type": response_type,
        "text": text,
        "files": files,
        "audit": {
            "intent": "report_generation",
            "status": result.get("status"),
            "report_type": result.get("report_type"),
            "target_period": result.get("target_period"),
            "upload_count": result.get("upload_count"),
            "blocked_periods": result.get("blocked_periods", []),
            "recommended_missing": result.get("recommended_missing", []),
            "blocking_missing": result.get("blocking_missing", []),
            "source_files": result.get("source_files", []),
            "csv_classifications": result.get("csv_classifications", []),
            "recommended_missing_is_blocking": False,
            "single_marketing_scene_csv_can_generate": True,
            "missing_groups": missing_groups(result),
            "document_files_received": [str(path) for path in document_files],
            "document_parse": document_audit or [],
            "used_tools": ["monthly_report_agent"],
        },
    }


def combine_period_results(results: list[dict[str, Any]], report_files: list[Path]) -> dict[str, Any]:
    final_reports: list[str] = []
    blocking_missing: list[str] = []
    recommended_missing: list[str] = []
    blocked_periods: list[Any] = []
    target_periods: list[str] = []
    report_types: list[str] = []
    for result in results:
        final_reports.extend(str(path) for path in result.get("final_reports") or [])
        blocking_missing.extend(str(item) for item in result.get("blocking_missing") or [])
        recommended_missing.extend(str(item) for item in result.get("recommended_missing") or [])
        blocked_periods.extend(result.get("blocked_periods") or [])
        if result.get("target_period"):
            target_periods.append(str(result.get("target_period")))
        if result.get("report_type"):
            report_types.append(str(result.get("report_type")))
    return {
        "status": "ok" if final_reports else "blocked",
        "report_type": "、".join(report_types),
        "target_period": "；".join(target_periods),
        "final_reports": final_reports,
        "upload_count": len(report_files),
        "source_files": [str(path) for path in report_files],
        "blocked_periods": blocked_periods,
        "recommended_missing": list(dict.fromkeys(recommended_missing)),
        "blocking_missing": list(dict.fromkeys(blocking_missing)),
        "message": "已按用户要求生成多个周期报表。" if final_reports else "暂未生成最终报表。",
    }


def missing_groups(result: dict[str, Any]) -> dict[str, list[str]]:
    blocking = [str(item) for item in result.get("blocking_missing") or []]
    recommended = [str(item) for item in result.get("recommended_missing") or []]
    grouped = {
        "必须补充": blocking,
        "建议补充": recommended,
    }
    if not blocking:
        grouped["必须补充"] = []
    return grouped


def missing_guidance_text(result: dict[str, Any]) -> str:
    grouped = missing_groups(result)
    blocking = grouped["必须补充"]
    recommended = grouped["建议补充"]
    if blocking:
        return "仍缺少必须数据：" + "、".join(blocking[:5]) + "。补齐后可生成完整店铺收益环比。"
    if recommended:
        preview = "、".join(recommended[:4])
        if len(recommended) > 4:
            preview += f"等 {len(recommended)} 项"
        return f"本次已可交付。建议补充项不影响生成，只会影响部分分析完整度：{preview}。"
    return "本次数据完整度较好，暂无必须补充项。"


def local_file_url(path: str | Path) -> str:
    return Path(path).resolve().as_uri()


def local_download_url(path: str | Path) -> str:
    return "http://127.0.0.1:8799/download?path=" + quote(str(Path(path).resolve()))


def publish_final_report(path: str | Path) -> Path:
    source = Path(path).resolve()
    if not is_allowed_final_report(source):
        raise ValueError(f"只允许发布月报、半月报或周报：{source.name}")
    FINAL_REPORT_DIR.mkdir(parents=True, exist_ok=True)
    target = unique_path(FINAL_REPORT_DIR / source.name)
    shutil.copy2(source, target)
    return target


def is_allowed_final_report(path: str | Path) -> bool:
    name = Path(path).name
    if Path(name).suffix.lower() != ".xlsx":
        return False
    if any(label in name for label in ("推广月报", "推广周报", "推广半月报", "标准经营总表", "数据半成品", "数据缺失", "AI分析", "任务包")):
        return False
    return name.endswith(("月报.xlsx", "半月报.xlsx", "周报.xlsx"))


def qclaw_ai_config(settings: dict[str, Any]) -> AIConfig:
    return AIConfig(
        provider="qclaw",
        api_key="",
        base_url=os.environ.get("QCLAW_LLM_BASE_URL", "http://127.0.0.1:19100/proxy/llm"),
        model=os.environ.get("QCLAW_LLM_MODEL", "modelroute"),
        mode=str(settings.get("ai_mode") or os.environ.get("WXT_AI_MODE", "fast")),
    )


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for index in range(2, 1000):
        candidate = path.with_name(f"{stem}_{index}{suffix}")
        if not candidate.exists():
            return candidate
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{stem}_{stamp}{suffix}")


def open_local_file(path: Path) -> None:
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        elif os.name == "nt":
            os.startfile(str(path))  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


def render_home() -> str:
    return page(
        """
        <main class="shell">
          <section class="panel">
            <div class="title-row">
              <div>
                <h1>万相台企业报表 Agent</h1>
                <p>上传生意参谋和万相台原始报表，自动生成月报、周报、半月报或阶段报表。</p>
              </div>
              <div class="actions">
                <a class="secondary" href="/chat">进入诊断机器人</a>
                <span class="badge">企业服务</span>
              </div>
            </div>
            <form id="reportForm" action="/generate" method="post" enctype="multipart/form-data" class="form">
              <label class="field">
                <span>上传原始报表</span>
                <input id="reportsInput" type="file" name="reports" multiple accept=".xlsx,.xls,.csv" required>
                <small>可以多次选择文件，系统会自动追加到待上传列表。文件名不限，日报/月报都可作为数据来源；新旧版报表任选其一，缺少广告或上期数据时会尽量生成简版。</small>
                <div class="file-toolbar">
                  <span id="reportsCount">已选择 0 个文件</span>
                  <button id="clearReports" class="mini-button" type="button">清空</button>
                </div>
                <ul id="reportsList" class="file-list"></ul>
              </label>
              <label class="field">
                <span>上传模板（可选）</span>
                <input id="templateInput" type="file" name="template" accept=".xlsx,.xls">
                <small>不上传时使用系统默认模板。</small>
                <div id="templateName" class="file-note">未选择模板</div>
              </label>
              <div class="grid">
                <label class="field">
                  <span>报表类型</span>
                  <select name="period">
                    <option value="month">月报</option>
                    <option value="week">周报</option>
                    <option value="half-month">半月报</option>
                  </select>
                </label>
                <label class="field">
                  <span>指定周期（可选）</span>
                  <input type="text" name="target" placeholder="例如 2026-05、2026-05上半月">
                </label>
              </div>
              <details class="advanced">
                <summary>AI文案增强（可选）</summary>
                <div class="grid">
                  <label class="field">
                    <span>AI模型</span>
                    <select id="reportAiProvider" name="ai_provider">
                      <option value="">本地规则（不调用AI）</option>
                      <option value="deepseek">DeepSeek</option>
                    </select>
                  </label>
                  <label class="field">
                    <span>DeepSeek模型</span>
                    <input id="reportDeepseekModel" type="text" name="deepseek_model" value="deepseek-v4-flash">
                  </label>
                </div>
                <label class="field">
                  <span>增强速度</span>
                  <select id="reportAiMode" name="ai_mode">
                    <option value="off">本地模板（最快）</option>
                    <option value="fast">极速增强（推荐）</option>
                    <option value="standard">标准增强</option>
                    <option value="full">完整增强</option>
                  </select>
                  <small>本地模板不调用模型；极速/标准/完整会逐级增加 DeepSeek 参与范围。</small>
                </label>
                <label class="field">
                  <span>DeepSeek API Key</span>
                  <input id="reportDeepseekApiKey" type="password" name="deepseek_api_key" placeholder="sk-...；不填写则不会调用">
                </label>
                <label class="field">
                  <span>接口地址</span>
                  <input id="reportDeepseekBaseUrl" type="text" name="deepseek_base_url" value="https://api.deepseek.com">
                </label>
              </details>
              <button type="submit">生成报表</button>
            </form>
          </section>
          <section class="help">
            <h2>员工使用说明</h2>
            <p>从钉钉工作台打开应用，选择文件后点击生成即可。生成完成后，页面会显示最终 Excel、缺失报告和交付说明下载入口。</p>
          </section>
        </main>
        <script>
          const reportForm = document.getElementById('reportForm');
          const reportProvider = document.getElementById('reportAiProvider');
          const reportKey = document.getElementById('reportDeepseekApiKey');
          const reportModel = document.getElementById('reportDeepseekModel');
          const reportBaseUrl = document.getElementById('reportDeepseekBaseUrl');
          const reportAiMode = document.getElementById('reportAiMode');
          const reportsInput = document.getElementById('reportsInput');
          const reportsList = document.getElementById('reportsList');
          const reportsCount = document.getElementById('reportsCount');
          const clearReports = document.getElementById('clearReports');
          const templateInput = document.getElementById('templateInput');
          const templateName = document.getElementById('templateName');
          let selectedReports = [];

          function fileKey(file) {
            return [file.name, file.size, file.lastModified].join('|');
          }

          function syncReportsInput() {
            const transfer = new DataTransfer();
            selectedReports.forEach(file => transfer.items.add(file));
            reportsInput.files = transfer.files;
            reportsCount.textContent = `已选择 ${selectedReports.length} 个文件`;
            reportsList.innerHTML = '';
            selectedReports.forEach((file, index) => {
              const li = document.createElement('li');
              const name = document.createElement('span');
              const remove = document.createElement('button');
              name.textContent = file.name;
              remove.type = 'button';
              remove.className = 'mini-button';
              remove.textContent = '移除';
              remove.addEventListener('click', () => {
                selectedReports.splice(index, 1);
                syncReportsInput();
              });
              li.append(name, remove);
              reportsList.appendChild(li);
            });
          }

          reportsInput.addEventListener('change', () => {
            const existing = new Set(selectedReports.map(fileKey));
            Array.from(reportsInput.files).forEach(file => {
              const key = fileKey(file);
              if (!existing.has(key)) {
                selectedReports.push(file);
                existing.add(key);
              }
            });
            syncReportsInput();
          });

          clearReports.addEventListener('click', () => {
            selectedReports = [];
            syncReportsInput();
          });

          templateInput.addEventListener('change', () => {
            templateName.textContent = templateInput.files[0] ? `当前模板：${templateInput.files[0].name}` : '未选择模板';
          });

          reportKey.value = localStorage.getItem('wxt_deepseek_api_key') || '';
          reportModel.value = localStorage.getItem('wxt_deepseek_model') || 'deepseek-v4-flash';
          reportBaseUrl.value = localStorage.getItem('wxt_deepseek_base_url') || 'https://api.deepseek.com';
          reportAiMode.value = localStorage.getItem('wxt_ai_mode') || 'fast';
          if (reportKey.value) reportProvider.value = 'deepseek';
          reportKey.addEventListener('input', () => {
            if (reportKey.value.trim()) reportProvider.value = 'deepseek';
          });
          reportForm.addEventListener('submit', () => {
            const key = reportKey.value.trim();
            syncReportsInput();
            if (key && !reportProvider.value) reportProvider.value = 'deepseek';
            localStorage.setItem('wxt_deepseek_api_key', key);
            localStorage.setItem('wxt_deepseek_model', reportModel.value.trim() || 'deepseek-v4-flash');
            localStorage.setItem('wxt_deepseek_base_url', reportBaseUrl.value.trim() || 'https://api.deepseek.com');
            localStorage.setItem('wxt_ai_mode', reportAiMode.value || 'fast');
          });
        </script>
        """
    )


def render_chat() -> str:
    return page(
        """
        <main class="chat-shell">
          <section class="chat-card">
            <header class="chat-header">
              <div>
                <h1>万相台诊断师</h1>
                <p>输入自己的 DeepSeek API Key 后，就可以直接咨询关键词、预算、ROI、商品和报表生成问题。</p>
              </div>
              <a class="secondary" href="/">生成报表</a>
            </header>
            <div class="key-panel">
              <label class="field">
                <span>DeepSeek API Key</span>
                <input id="chatApiKey" type="password" placeholder="sk-...">
              </label>
              <label class="field">
                <span>模型</span>
                <input id="chatModel" type="text" value="deepseek-v4-flash">
              </label>
              <label class="field">
                <span>接口地址</span>
                <input id="chatBaseUrl" type="text" value="https://api.deepseek.com">
              </label>
            </div>
            <div id="chatMessages" class="chat-messages">
              <article class="bot bubble">
                <strong>你好，我是万相台诊断师。</strong>
                <p>你可以问我“关键词推广怎么做”“为什么 ROI 低”“下月预算怎么分”，也可以先生成报表后带着数据来追问。</p>
              </article>
            </div>
            <form id="chatForm" class="chat-input">
              <textarea id="chatText" rows="3" placeholder="请输入消息"></textarea>
              <button type="submit">发送</button>
            </form>
          </section>
        </main>
        <script>
          const apiKey = document.getElementById('chatApiKey');
          const model = document.getElementById('chatModel');
          const baseUrl = document.getElementById('chatBaseUrl');
          const messagesEl = document.getElementById('chatMessages');
          const form = document.getElementById('chatForm');
          const input = document.getElementById('chatText');
          const messages = [];

          apiKey.value = localStorage.getItem('wxt_deepseek_api_key') || '';
          model.value = localStorage.getItem('wxt_deepseek_model') || 'deepseek-v4-flash';
          baseUrl.value = localStorage.getItem('wxt_deepseek_base_url') || 'https://api.deepseek.com';

          function saveSettings() {
            localStorage.setItem('wxt_deepseek_api_key', apiKey.value.trim());
            localStorage.setItem('wxt_deepseek_model', model.value.trim() || 'deepseek-v4-flash');
            localStorage.setItem('wxt_deepseek_base_url', baseUrl.value.trim() || 'https://api.deepseek.com');
          }

          function addMessage(role, text) {
            const item = document.createElement('article');
            item.className = (role === 'user' ? 'user' : 'bot') + ' bubble';
            item.innerHTML = text.replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c])).replace(/\\n/g, '<br>');
            messagesEl.appendChild(item);
            messagesEl.scrollTop = messagesEl.scrollHeight;
          }

          form.addEventListener('submit', async (event) => {
            event.preventDefault();
            const text = input.value.trim();
            if (!text) return;
            saveSettings();
            addMessage('user', text);
            messages.push({role: 'user', content: text});
            input.value = '';
            const thinking = '正在分析...';
            addMessage('assistant', thinking);
            const lastBubble = messagesEl.lastElementChild;
            try {
              const response = await fetch('/api/chat', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                  api_key: apiKey.value.trim(),
                  model: model.value.trim(),
                  base_url: baseUrl.value.trim(),
                  messages
                })
              });
              const data = await response.json();
              if (data.status !== 'ok') throw new Error(data.message || '请求失败');
              lastBubble.innerHTML = data.answer.replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c])).replace(/\\n/g, '<br>');
              messages.push({role: 'assistant', content: data.answer});
            } catch (error) {
              lastBubble.textContent = '连接失败：' + error.message;
            }
          });
        </script>
        """
    )


def render_result(result: dict[str, Any]) -> str:
    status = html.escape(str(result.get("status", "")))
    message = html.escape(str(result.get("message", "")))
    final_reports = result.get("final_reports") or []
    blocked_periods = result.get("blocked_periods") or []
    recommended = result.get("recommended_missing") or []
    output_dir = Path(str(result.get("output_dir", "")))
    handoff = output_dir / "交付说明.md"
    manifest = output_dir / "enterprise_manifest.json"

    final_html = "".join(
        f"<li>{file_link(path, Path(path).name)}<small class=\"path\">{html.escape(str(path))}</small></li>"
        for path in final_reports
    ) or "<li>未生成最终报表</li>"
    audit_html = render_audit_preview(result, handoff, manifest)
    blocked_html = render_blocked_summary(blocked_periods)
    return page(
        f"""
        <main class="shell">
          <section class="panel">
            <div class="title-row">
              <div>
                <h1>生成结果</h1>
                <p class="status {status}">{status}：{message}</p>
              </div>
              <a class="secondary" href="/">继续生成</a>
            </div>
            <h2>最终报表</h2>
            <ul class="downloads">{final_html}</ul>
            <h2>审计与校验</h2>
            {audit_html}
            <h2>未生成周期</h2>
            {blocked_html}
          </section>
        </main>
        """
    )


def render_audit_preview(result: dict[str, Any], handoff: Path, manifest: Path) -> str:
    upload_count = int(result.get("upload_count") or 0)
    recommended = result.get("recommended_missing") or []
    blocked_periods = result.get("blocked_periods") or []
    validation_report = result.get("validation_report")
    standard_total = result.get("standard_total")
    files = []
    if validation_report:
        files.append(file_link(validation_report, "下载识别报告"))
    if standard_total:
        files.append(file_link(standard_total, "下载标准经营总表"))
    if handoff.exists():
        files.append(file_link(str(handoff), "下载交付说明"))
    if manifest.exists():
        files.append(file_link(str(manifest), "下载运行清单"))
    files_html = " ".join(files) or "暂无可下载审计文件"
    recommended_html = render_compact_items(recommended, "无强制补充项，当前数据已可用于生成。", limit=4)
    blocked_note = "无未生成周期" if not blocked_periods else f"{len(blocked_periods)} 个周期因核心数据缺失未生成"
    return f"""
      <div class="audit-grid">
        <article class="audit-card">
          <strong>上传识别</strong>
          <p>已接收 {upload_count} 个文件，系统已完成格式识别和核心数据校验。</p>
        </article>
        <article class="audit-card">
          <strong>生成状态</strong>
          <p>{html.escape(blocked_note)}。</p>
        </article>
        <article class="audit-card">
          <strong>建议补充</strong>
          {recommended_html}
        </article>
      </div>
      <details class="download-panel">
        <summary>需要留档时下载审计文件</summary>
        <div class="download-actions">{files_html}</div>
      </details>
    """


def render_compact_items(items: list[Any], empty: str, limit: int = 4) -> str:
    if not items:
        return f"<p>{html.escape(empty)}</p>"
    visible = items[:limit]
    rows = "".join(f"<li>{html.escape(str(item))}</li>" for item in visible)
    more = len(items) - len(visible)
    if more > 0:
        rows += f"<li>另有 {more} 项，可下载识别报告查看。</li>"
    return f"<ul class=\"compact-list\">{rows}</ul>"


def render_blocked_summary(items: list[dict[str, Any]]) -> str:
    if not items:
        return "<p>无，所有目标周期均已生成。</p>"
    periods = "、".join(str(item.get("period")) for item in items[:6])
    more = len(items) - min(len(items), 6)
    if more > 0:
        periods += f" 等 {len(items)} 个周期"
    rows = []
    for item in items:
        missing = item.get("missing", [])
        missing_text = "、".join(str(value) for value in missing[:3])
        if len(missing) > 3:
            missing_text += f" 等 {len(missing)} 项"
        rows.append(f"<li><strong>{html.escape(str(item.get('period')))}</strong>：{html.escape(missing_text)}</li>")
    return f"""
      <p>{len(items)} 个周期未生成：{html.escape(periods)}。主要原因是核心数据不完整。</p>
      <details class="download-panel">
        <summary>查看具体缺失项</summary>
        <ul class="compact-list">{''.join(rows)}</ul>
      </details>
    """


def render_error(error: dict[str, str]) -> str:
    return page(
        f"""
        <main class="shell">
          <section class="panel">
            <h1>生成失败</h1>
            <p>{html.escape(error.get("message", ""))}</p>
            <details>
              <summary>技术信息</summary>
              <pre>{html.escape(error.get("trace", ""))}</pre>
            </details>
            <a class="secondary" href="/">返回重新上传</a>
          </section>
        </main>
        """
    )


def file_link(path: str, label: str) -> str:
    return f'<a href="/download?path={quote(str(Path(path).resolve()))}">{html.escape(label)}</a>'


def page(content: str) -> str:
    return f"""
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{APP_TITLE}</title>
  <style>
    :root {{
      font-family: "Microsoft YaHei", "PingFang SC", Arial, sans-serif;
      color: #1f2933;
      background: #f6f7f9;
    }}
    body {{ margin: 0; }}
    .shell {{ max-width: 980px; margin: 36px auto; padding: 0 24px; }}
    .panel {{
      background: #fff;
      border: 1px solid #d9dee7;
      border-radius: 8px;
      padding: 28px;
      box-shadow: 0 10px 30px rgba(15, 23, 42, .06);
    }}
    .title-row {{ display: flex; align-items: flex-start; justify-content: space-between; gap: 24px; }}
    .actions {{ display: flex; align-items: center; gap: 12px; flex-wrap: wrap; justify-content: flex-end; }}
    h1 {{ margin: 0 0 8px; font-size: 28px; }}
    h2 {{ margin: 28px 0 12px; font-size: 18px; }}
    p {{ line-height: 1.7; margin: 0; color: #52606d; }}
    .badge {{ border: 1px solid #b7c4d6; border-radius: 999px; padding: 6px 12px; color: #334e68; white-space: nowrap; }}
    .form {{ margin-top: 28px; display: grid; gap: 20px; }}
    .grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
    .field {{ display: grid; gap: 8px; font-weight: 600; }}
    input, select {{
      font: inherit;
      border: 1px solid #cbd2d9;
      border-radius: 6px;
      padding: 11px 12px;
      background: #fff;
    }}
    input[type=file] {{ padding: 9px; }}
    small {{ color: #6b7280; font-weight: 400; }}
    .file-toolbar {{ display: flex; align-items: center; justify-content: space-between; gap: 12px; color: #52606d; font-weight: 500; }}
    .mini-button {{
      border: 1px solid #cbd2d9;
      border-radius: 6px;
      background: #fff;
      color: #334e68;
      padding: 6px 10px;
      font: inherit;
      font-size: 13px;
      font-weight: 700;
      cursor: pointer;
      width: fit-content;
    }}
    .file-list {{ margin: 0; padding: 0; list-style: none; display: grid; gap: 8px; }}
    .file-list li {{ display: flex; align-items: center; justify-content: space-between; gap: 12px; border: 1px solid #e5e7eb; border-radius: 6px; padding: 8px 10px; font-weight: 500; }}
    .file-list span {{ overflow-wrap: anywhere; }}
    .file-note {{ color: #52606d; font-weight: 500; }}
    button, .secondary {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border: 1px solid #1f4e79;
      border-radius: 6px;
      background: #1f4e79;
      color: #fff;
      padding: 12px 18px;
      font: inherit;
      font-weight: 700;
      text-decoration: none;
      cursor: pointer;
      width: fit-content;
    }}
    .secondary {{ background: #fff; color: #1f4e79; }}
    .help {{ margin-top: 18px; padding: 0 4px; }}
    .advanced {{ border: 1px solid #d9dee7; border-radius: 8px; padding: 14px 16px; }}
    .advanced summary {{ cursor: pointer; font-weight: 700; color: #334e68; }}
    ul {{ line-height: 1.8; }}
    .downloads a {{ color: #1f4e79; font-weight: 700; }}
    .path {{ display: block; margin-top: 4px; color: #7b8794; word-break: break-all; }}
    .audit-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; }}
    .audit-card {{ border: 1px solid #d9dee7; border-radius: 8px; padding: 14px; background: #fff; }}
    .audit-card strong {{ display: block; margin-bottom: 8px; color: #1f2933; }}
    .audit-card p {{ font-size: 14px; }}
    .compact-list {{ margin: 0; padding-left: 18px; font-size: 14px; }}
    .download-panel {{ margin-top: 14px; border: 1px solid #d9dee7; border-radius: 8px; padding: 12px 14px; }}
    .download-panel summary {{ cursor: pointer; font-weight: 700; color: #334e68; }}
    .download-actions {{ display: flex; flex-wrap: wrap; gap: 10px; margin-top: 12px; }}
    .download-actions a {{ color: #1f4e79; font-weight: 700; }}
    .status.ok {{ color: #1f7a4d; font-weight: 700; }}
    .status.blocked, .status.error {{ color: #a23b3b; font-weight: 700; }}
    pre {{ white-space: pre-wrap; background: #111827; color: #f9fafb; padding: 16px; border-radius: 6px; overflow: auto; }}
    .chat-shell {{ max-width: 860px; margin: 24px auto; padding: 0 16px; }}
    .chat-card {{ min-height: calc(100vh - 48px); background: #fff; border: 1px solid #d9dee7; border-radius: 8px; display: grid; grid-template-rows: auto auto 1fr auto; overflow: hidden; }}
    .chat-header {{ padding: 22px 24px; border-bottom: 1px solid #e5e7eb; display: flex; align-items: flex-start; justify-content: space-between; gap: 18px; }}
    .key-panel {{ padding: 16px 24px; border-bottom: 1px solid #e5e7eb; display: grid; grid-template-columns: 1.4fr .8fr 1fr; gap: 14px; }}
    .chat-messages {{ padding: 22px 24px; overflow: auto; background: #f7f8fa; display: flex; flex-direction: column; gap: 14px; }}
    .bubble {{ max-width: 78%; padding: 14px 16px; border: 1px solid #e1e5ea; border-radius: 8px; background: #fff; line-height: 1.7; white-space: normal; }}
    .bubble p {{ margin-top: 8px; color: #374151; }}
    .bubble.user {{ align-self: flex-end; background: #dbeafe; border-color: #bfdbfe; }}
    .bubble.bot {{ align-self: flex-start; }}
    .chat-input {{ padding: 16px 24px; border-top: 1px solid #e5e7eb; display: grid; grid-template-columns: 1fr auto; gap: 12px; background: #fff; }}
    textarea {{ font: inherit; border: 1px solid #cbd2d9; border-radius: 6px; padding: 12px; resize: vertical; min-height: 72px; }}
    @media (max-width: 720px) {{
      .grid, .title-row {{ grid-template-columns: 1fr; display: grid; }}
      .shell {{ margin: 18px auto; padding: 0 14px; }}
      .audit-grid {{ grid-template-columns: 1fr; }}
      .key-panel, .chat-header, .chat-input {{ grid-template-columns: 1fr; display: grid; }}
      .bubble {{ max-width: 100%; }}
    }}
  </style>
</head>
<body>
{content}
</body>
</html>
"""


if __name__ == "__main__":
    main()
