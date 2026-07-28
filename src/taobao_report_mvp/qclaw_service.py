from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import time
import traceback
import zipfile
from datetime import datetime
from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import argparse
import threading

from .ai_provider import AIConfig
from .enterprise_agent import build_enterprise_result, normalize_period, write_enterprise_manifest, write_handoff_note
from .qclaw_layers import split_uploaded_files, user_has_confirmed_upload
from .report_workflow_agent import DEFAULT_TEMPLATE_CANDIDATES, build_output_paths, first_existing, run_workflow


APP_TITLE = os.environ.get("WXT_APP_TITLE", "万相台月报 API")
SUPPORTED_SUFFIXES = {".xlsx", ".xls", ".csv", ".zip"}
QCLAW_UPLOAD_SUFFIXES = {".xlsx", ".xls", ".csv", ".doc", ".docx", ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".zip"}
QCLAW_EXTRACTABLE_SUFFIXES = {".xlsx", ".xls", ".csv", ".doc", ".docx", ".png", ".jpg", ".jpeg", ".webp", ".bmp"}
QCLAW_REPORT_SUFFIXES = {".xlsx", ".xls", ".csv"}
DEFAULT_HOST = os.environ.get("WXT_HOST", "127.0.0.1")
DEFAULT_PORT = int(os.environ.get("WXT_PORT", "8788"))
OUTPUT_ROOT = Path(os.environ.get("WXT_OUTPUT_ROOT", "outputs/qclaw_runs"))
FINAL_REPORT_DIR = Path(os.environ.get("WXT_FINAL_REPORT_DIR", "/Users/gordon/Documents/月报/最终月报"))
SOURCE_CACHE_DIR = Path(os.environ.get("WXT_SOURCE_CACHE_DIR", "/Users/gordon/Documents/月报/补充缓存"))
TEMP_OUTPUT_MAX_AGE_DAYS = int(os.environ.get("WXT_TEMP_OUTPUT_MAX_AGE_DAYS", "1"))
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


def main() -> None:
    parser = argparse.ArgumentParser(description="启动万相台月报 QClaw 本地 API 服务")
    parser.add_argument("--host", default=DEFAULT_HOST, help="监听地址，默认 127.0.0.1")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT, help="监听端口，默认 8788")
    args = parser.parse_args()

    cleanup_old_output_runs()
    server, port = create_server(args.host, args.port)
    url = f"http://{args.host}:{port}/qclaw/message"
    print(f"{APP_TITLE} 已启动: {url}")
    print("仅提供 QClaw/OpenClaw API。按 Ctrl+C 停止服务。")
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
    message = "无法启动本地 API 服务，已尝试端口 " + f"{preferred_port}-{preferred_port + 19}。"
    if errors:
        message += "\n" + "\n".join(errors[-5:])
    raise RuntimeError(message)


class ReportAppHandler(BaseHTTPRequestHandler):
    server_version = "WanxiangtaiReportAgent/1.0"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/healthz":
            self.send_json({"status": "ok", "app": APP_TITLE, "service": "qclaw_api_only"})
        else:
            self.send_json(
                {
                    "type": "error",
                    "text": "本服务只提供 QClaw/OpenClaw API，请调用 /qclaw/message。",
                    "files": [],
                },
                status=HTTPStatus.NOT_FOUND,
            )

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/qclaw/message":
            self.handle_qclaw_message()
            return
        self.send_json(
            {
                "type": "error",
                "text": "本服务只提供 QClaw/OpenClaw API，请调用 /qclaw/message。",
                "files": [],
            },
            status=HTTPStatus.NOT_FOUND,
        )

    def log_message(self, format: str, *args: Any) -> None:
        print(f"[qclaw-api] {self.address_string()} - {format % args}")

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
            "ai_mode": field_value(form, "ai_mode", "claw").strip(),
            "period": field_value(form, "period", "").strip(),
            "target": field_value(form, "target", "").strip(),
            "session_key": field_value(form, "session_key", "").strip(),
            "conversation_id": field_value(form, "conversation_id", "").strip(),
            "user_id": field_value(form, "user_id", "").strip(),
            "action": field_value(form, "action", "").strip(),
            "report_mode": field_value(form, "report_mode", "").strip(),
        }
        run_name = datetime.now().strftime("%Y%m%d_%H%M%S")
        upload_dir = OUTPUT_ROOT / f"qclaw_{run_name}" / "uploads"
        files = save_uploads(form, "files", upload_dir, allowed_suffixes=QCLAW_UPLOAD_SUFFIXES)
        files.extend(save_uploads(form, "reports", upload_dir, allowed_suffixes=QCLAW_UPLOAD_SUFFIXES))
        files = expand_archives(files, OUTPUT_ROOT / "qclaw_unzipped" / run_name, QCLAW_EXTRACTABLE_SUFFIXES)
        return self.run_qclaw_request(text, files, settings)

    def run_qclaw_request(self, text: str, files: list[Path], settings: dict[str, Any]) -> dict[str, Any]:
        period = requested_period_from_text_or_settings(text, settings)
        effective_period = period or "month"
        target = str(settings.get("target") or "").strip() or None
        report_files, document_files = split_uploaded_files(files)
        document_audit = parse_document_files(document_files, OUTPUT_ROOT / "qclaw_file2md")
        route = classify_qclaw_intent(text, report_files, document_files)
        wants_report = route["intent"] == "report_generation"
        if wants_report and not report_files:
            session_key = report_session_key(settings, self.client_address[0] if self.client_address else "")
            last_report = get_last_report(session_key)
            if last_report and last_report.source_files and wants_cached_generation(text):
                result = self.generate_report_from_collected_files(
                    list(last_report.source_files),
                    document_files,
                    period or last_report.period or "month",
                    target or last_report.target,
                    settings,
                )
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
                },
            }
        if wants_report and report_files:
            periods = requested_periods_from_text_or_settings(text, settings)
            results: list[dict[str, Any]] = []
            for item_period in periods:
                run_name = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_dir = OUTPUT_ROOT / f"qclaw_{item_period}_{run_name}"
                output_dir.mkdir(parents=True, exist_ok=True)
                template = first_existing(DEFAULT_TEMPLATE_CANDIDATES)
                ai_config = AIConfig(
                    provider="qclaw",
                    api_key="",
                    base_url=os.environ.get("QCLAW_LLM_BASE_URL", "http://127.0.0.1:19100/proxy/llm"),
                    model=os.environ.get("QCLAW_LLM_MODEL", "modelroute"),
                    mode=str(settings.get("ai_mode") or os.environ.get("WXT_AI_MODE", "fast")),
                )
                raw_result = run_workflow(report_files, template, build_output_paths(output_dir), target, item_period, ai_config=ai_config)
                result = build_enterprise_result(raw_result, report_files, output_dir, item_period, target)
                write_enterprise_manifest(output_dir, report_files, item_period, target, result)
                write_handoff_note(output_dir, result)
                store_last_report(report_session_key(settings, self.client_address[0] if self.client_address else ""), result, report_files, item_period, target)
                results.append(result)
            if len(results) == 1:
                return qclaw_report_response(results[0], document_files, document_audit)
            combined = combine_period_results(results, report_files)
            return qclaw_report_response(combined, document_files, document_audit)
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
                    "report_files": [str(path) for path in report_files],
                    "next_step": "upload_report_csv_or_xlsx_for_final_xlsx",
                },
            }
        return {
            "type": "needs_input",
            "text": "这是万相台月报生成 Skill。请进入月报板块并上传 csv/xls/xlsx 报表，确认后我会生成最终 xlsx。普通万相台咨询请交给诊断类 Agent。",
            "files": [],
            "audit": {"intent": "monthly_report_skill_only", "route": route},
        }

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
            result = self.generate_report_from_collected_files(combined_files, document_files, period or last_report.period or "month", last_report.target or target, settings)
            with REPORT_SESSIONS_LOCK:
                REPORT_SESSIONS.pop(session_key, None)
            result.setdefault("audit", {})["supplemented_previous_report"] = True
            return result

        if action == "new":
            with REPORT_SESSIONS_LOCK:
                session = ReportCollectionSession(session_key=session_key, period=period or "month", target=target, files=list(report_files))
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
            if report_files:
                existing = {str(path) for path in session.files or []}
                for path in report_files:
                    if str(path) not in existing:
                        session.files.append(path)
                        existing.add(str(path))
            session.updated_at = datetime.now().isoformat(timespec="seconds")
            collected_files = list(session.files or [])

        if action == "confirm":
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
            return report_collection_response(
                f"已收到 {len(report_files)} 个新文件，当前这个店铺共暂存 {len(collected_files)} 个可用表格。{describe_collected_files(collected_files)}这个店铺还有别的文件吗？有的话继续上传；没有就点“确认生成”。",
                session_key,
                collected_files,
                status="collecting",
            )

        return report_collection_response(
            f"当前月报板块已暂存 {len(collected_files)} 个可用表格。{describe_collected_files(collected_files)}请继续上传，或点“确认生成”。",
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
        result = build_enterprise_result(raw_result, report_files, output_dir, period, target)
        write_enterprise_manifest(output_dir, report_files, period, target, result)
        write_handoff_note(output_dir, result)
        store_last_report(report_session_key(settings, self.client_address[0] if self.client_address else ""), result, report_files, period, target)
        response = qclaw_report_response(result, document_files, document_audit)
        response.setdefault("audit", {})["collection_mode"] = True
        return response

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
    for path in files:
        source = Path(path).expanduser()
        if source.suffix.lower() == ".zip":
            expanded.extend(expand_zip_archive(source, target_dir / source.stem, allowed_suffixes))
        elif source.exists():
            expanded.append(source.resolve())
    return unique_paths(expanded)


def expand_zip_archive(zip_path: Path, target_dir: Path, allowed_suffixes: set[str]) -> list[Path]:
    if not zip_path.exists() or not zipfile.is_zipfile(zip_path):
        return []
    target_dir.mkdir(parents=True, exist_ok=True)
    extracted: list[Path] = []
    allowed = set(allowed_suffixes) - {".zip"}
    with zipfile.ZipFile(zip_path) as archive:
        for info in archive.infolist():
            if info.is_dir():
                continue
            inner = Path(info.filename)
            if inner.name.startswith(".") or "__MACOSX" in inner.parts:
                continue
            if inner.suffix.lower() not in allowed:
                continue
            filename = safe_filename(inner.name)
            if not filename:
                continue
            destination = unique_path(target_dir / filename)
            with archive.open(info) as src, destination.open("wb") as out:
                shutil.copyfileobj(src, out)
            extracted.append(destination.resolve())
    return extracted


def safe_filename(filename: str | None) -> str:
    if not filename:
        return ""
    return Path(filename).name.replace("\x00", "").strip()


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


def cleanup_old_output_runs(max_age_days: int | None = None) -> None:
    max_age_days = TEMP_OUTPUT_MAX_AGE_DAYS if max_age_days is None else max_age_days
    if max_age_days <= 0:
        return
    cutoff = time.time() - max_age_days * 24 * 60 * 60
    cleanup_old_path_children(OUTPUT_ROOT, cutoff)
    cleanup_old_path_children(FINAL_REPORT_DIR, cutoff)


def cleanup_old_path_children(root: Path, cutoff: float) -> None:
    if not root.exists() or not root.is_dir():
        return
    for child in root.iterdir():
        try:
            if child.resolve() == SOURCE_CACHE_DIR.resolve():
                continue
            if child.stat().st_mtime < cutoff:
                if child.is_dir():
                    shutil.rmtree(child)
                else:
                    child.unlink()
        except Exception:
            continue


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
    stripped = text.strip()
    if any(keyword in stripped for keyword in ("进入月报板块", "月报板块", "开始月报", "我要生成月报", "收集月报文件")):
        return "enter"
    if any(keyword in stripped for keyword in ("发完了", "发送完了", "上传完了", "确认生成", "开始生成", "现在生成", "生成吧")):
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
        or normalize_report_action(text, settings) in {"enter", "confirm", "cancel", "supplement", "new"}
    )


def report_collection_response(
    text: str,
    session_key: str,
    files: list[Path],
    status: str = "collecting",
) -> dict[str, Any]:
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
        action_settings = report_action_settings(session_key)
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
            {"id": "supplement_report", "label": "补充上一份月报", "type": "submit", "settings": {"action": "supplement_report", "report_mode": "1", "session_key": session_key, "period": last_report.period or "month"}},
            {"id": "new_report", "label": "生成新月报", "type": "submit", "settings": {"action": "new_report", "report_mode": "1", "session_key": session_key}},
            {"id": "cancel_report", "label": "取消", "type": "submit", "settings": {"action": "cancel_report", "report_mode": "1", "session_key": session_key, "period": last_report.period or "month"}},
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
    lowered = text.lower()
    keywords = [
        "生成月报",
        "生成周报",
        "生成半月报",
        "制作月报",
        "制作周报",
        "制作半月报",
        "最终报表",
        "生成报表",
        "导出报表",
        "输出报表",
        "月报",
        "周报",
        "半月报",
        "xlsx",
        "excel",
    ]
    return any(keyword.lower() in lowered for keyword in keywords)


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
                "intent": "report_generation",
                "reason": "spreadsheet_uploaded_without_question",
                "uses_monthly_program": True,
            }
        if any(keyword in stripped for keyword in file_analysis_keywords):
            return {
                "intent": "report_generation",
                "reason": "spreadsheet_uploaded_with_analysis_or_output_keyword",
                "uses_monthly_program": True,
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


def infer_period_from_text(text: str) -> str:
    explicit_period = infer_explicit_period_from_text(text)
    if explicit_period:
        return explicit_period
    return "month"


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
    final_reports = result.get("final_reports") or []
    published_reports = [publish_final_report(path) for path in final_reports]
    files = [
        {
            "name": Path(path).name,
            "path": str(path),
            "folder_path": str(Path(path).parent),
            "url": local_file_url(path),
            "folder_url": local_file_url(Path(path).parent),
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


def publish_final_report(path: str | Path) -> Path:
    source = Path(path).resolve()
    FINAL_REPORT_DIR.mkdir(parents=True, exist_ok=True)
    target = FINAL_REPORT_DIR / source.name
    if target.exists():
        target.unlink()
    shutil.copy2(source, target)
    return target


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



if __name__ == "__main__":
    main()
