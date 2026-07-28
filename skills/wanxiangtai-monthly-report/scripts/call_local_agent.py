from __future__ import annotations

import argparse
import json
import mimetypes
import os
import platform
import subprocess
import sys
import tempfile
import time
import uuid
from pathlib import Path
from urllib import request
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse


BOUNDARY_PREFIX = "----WanxiangtaiQclaw"


def find_skill_root() -> Path:
    current = Path(__file__).resolve()
    for parent in current.parents:
        if (parent / "SKILL.md").exists():
            return parent
    raise SystemExit("未找到 SKILL.md，请确认 skill 包完整。")


def find_runtime_root(skill_root: Path) -> Path:
    candidates = [
        skill_root.parent / "runtime",
        skill_root.parent.parent / "runtime",
        Path(os.environ.get("QCLAW_RUNTIME_ROOT", "")).expanduser(),
    ]
    for candidate in candidates:
        if candidate and (candidate / "services" / "report.py").exists():
            return candidate.resolve()
    raise SystemExit("未找到共享 runtime，请确认 runtime 与 skill 文件夹在同一总目录下。")


SKILL_ROOT = find_skill_root()


def find_embedded_service_root(skill_root: Path) -> Path | None:
    service_root = skill_root / "assets" / "service"
    package_root = service_root / "src" / "taobao_report_mvp"
    if (package_root / "web_app.py").exists() and (service_root / "start_local_service.py").exists():
        return service_root.resolve()
    return None


EMBEDDED_SERVICE_ROOT = find_embedded_service_root(SKILL_ROOT)
RUNTIME_ROOT = None if EMBEDDED_SERVICE_ROOT else find_runtime_root(SKILL_ROOT)


def python_bin() -> str:
    configured = os.environ.get("QCLAW_PYTHON_BINARY", "").strip()
    if configured:
        return configured
    if platform.system() == "Windows":
        return "python"
    return sys.executable or "python3"


def service_health_url(endpoint: str) -> str:
    parsed = urlparse(endpoint)
    return f"{parsed.scheme}://{parsed.netloc}/healthz"


def service_port(endpoint: str) -> str:
    return str(urlparse(endpoint).port or 8799)


def service_is_alive(endpoint: str) -> bool:
    try:
        with request.urlopen(service_health_url(endpoint), timeout=2) as resp:
            return resp.status == 200
    except Exception:
        return False


def start_service(endpoint: str) -> None:
    if service_is_alive(endpoint):
        return
    env = os.environ.copy()
    env.setdefault("WXT_HOST", "127.0.0.1")
    env.setdefault("WXT_PORT", service_port(endpoint))
    env.setdefault("WXT_AI_MODE", "fast")
    env.setdefault("WXT_TEMP_OUTPUT_MAX_AGE_DAYS", "1")
    existing_pythonpath = env.get("PYTHONPATH", "")
    env.setdefault("WXT_OUTPUT_ROOT", str(SKILL_ROOT / "outputs" / "runs"))
    env.setdefault("WXT_FINAL_REPORT_DIR", str(SKILL_ROOT / "outputs" / "final_reports"))
    if EMBEDDED_SERVICE_ROOT:
        pythonpath_root = str(EMBEDDED_SERVICE_ROOT / "src")
        env.setdefault("WXT_APP_ROOT", str(EMBEDDED_SERVICE_ROOT))
        command = [python_bin(), str(EMBEDDED_SERVICE_ROOT / "start_local_service.py")]
    else:
        assert RUNTIME_ROOT is not None
        pythonpath_root = str(RUNTIME_ROOT.parent)
        command = [
            python_bin(),
            str(RUNTIME_ROOT / "services" / "report.py"),
            "serve",
            "--skill-root",
            str(SKILL_ROOT),
            "--port",
            service_port(endpoint),
        ]
    env["PYTHONPATH"] = pythonpath_root if not existing_pythonpath else pythonpath_root + os.pathsep + existing_pythonpath
    process_dir = Path(tempfile.gettempdir()) / "wanxiangtai-monthly-report"
    process_dir.mkdir(parents=True, exist_ok=True)
    log_dir = process_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    with (log_dir / "local-agent.log").open("a", encoding="utf-8") as log_file:
        subprocess.Popen(
            command,
            cwd=str(process_dir),
            env=env,
            stdout=log_file,
            stderr=subprocess.STDOUT,
            start_new_session=True,
        )
    for _ in range(40):
        if service_is_alive(endpoint):
            return
        time.sleep(0.5)
    raise SystemExit("本地月报服务启动超时，请确认 Python 已安装 pandas/openpyxl。")


def encode_multipart(fields: dict[str, str], files: list[Path]) -> tuple[bytes, str]:
    boundary = BOUNDARY_PREFIX + uuid.uuid4().hex
    chunks: list[bytes] = []
    for name, value in fields.items():
        if value == "":
            continue
        chunks.append(f"--{boundary}\r\n".encode())
        chunks.append(f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode())
        chunks.append(str(value).encode("utf-8"))
        chunks.append(b"\r\n")
    for path in files:
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        chunks.append(f"--{boundary}\r\n".encode())
        chunks.append(f'Content-Disposition: form-data; name="files"; filename="{path.name}"\r\n'.encode("utf-8"))
        chunks.append(f"Content-Type: {mime}\r\n\r\n".encode())
        chunks.append(path.read_bytes())
        chunks.append(b"\r\n")
    chunks.append(f"--{boundary}--\r\n".encode())
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def post_message(endpoint: str, text: str, files: list[Path], fields: dict[str, str]) -> dict:
    start_service(endpoint)
    if files:
        body, content_type = encode_multipart({"text": text, **fields}, files)
        req = request.Request(endpoint, data=body, headers={"Content-Type": content_type}, method="POST")
    else:
        body = json.dumps({"text": text, "files": [], "settings": fields}, ensure_ascii=False).encode("utf-8")
        req = request.Request(endpoint, data=body, headers={"Content-Type": "application/json"}, method="POST")
    try:
        with request.urlopen(req, timeout=900) as resp:
            return json.loads(resp.read().decode("utf-8", errors="replace"))
    except HTTPError as exc:
        content = exc.read().decode("utf-8", errors="replace")
        raise SystemExit(f"月报服务 HTTP {exc.code}: {content}") from exc
    except URLError as exc:
        service_location = EMBEDDED_SERVICE_ROOT or RUNTIME_ROOT or SKILL_ROOT
        raise SystemExit(f"本地月报服务未连接：{service_location}") from exc


def open_final_reports(result: dict) -> None:
    if result.get("type") != "report":
        return
    for item in result.get("files", []):
        if item.get("kind") != "final_report":
            continue
        path = Path(str(item.get("path", ""))).expanduser()
        if not path.exists():
            continue
        try:
            if platform.system() == "Darwin":
                subprocess.Popen(["open", str(path)])
            elif platform.system() == "Windows":
                os.startfile(str(path))  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception:
            pass


def final_report_media_paths(result: dict) -> list[Path]:
    media_paths: list[Path] = []
    seen: set[Path] = set()
    if result.get("type") != "report":
        return media_paths
    for item in result.get("files", []):
        if item.get("kind") != "final_report":
            continue
        path = Path(str(item.get("path", ""))).expanduser().resolve()
        if path.suffix.lower() != ".xlsx" or not path.exists() or path in seen:
            continue
        media_paths.append(path)
        seen.add(path)
    return media_paths


def main() -> None:
    parser = argparse.ArgumentParser(description="Call Wanxiangtai report service.")
    parser.add_argument("--text", required=True, help="用户原话")
    parser.add_argument("--file", action="append", default=[], help="上传文件路径，可重复")
    parser.add_argument("--endpoint", default="http://127.0.0.1:8799/qclaw/message")
    parser.add_argument("--period", default="")
    parser.add_argument("--target", default="")
    parser.add_argument("--ai-mode", default="fast")
    parser.add_argument("--session-key", default="")
    parser.add_argument("--open", action="store_true", help="生成后自动打开最终 Excel；默认不打开")
    args = parser.parse_args()
    files = [Path(item).expanduser().resolve() for item in args.file]
    missing = [str(path) for path in files if not path.exists()]
    if missing:
        raise SystemExit("文件不存在：" + "、".join(missing))
    result = post_message(
        args.endpoint,
        args.text,
        files,
        {
            "period": args.period,
            "target": args.target,
            "ai_mode": args.ai_mode,
            "session_key": args.session_key,
        },
    )
    if args.open:
        open_final_reports(result)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    for path in final_report_media_paths(result):
        print(f"MEDIA:{path}")


if __name__ == "__main__":
    main()
