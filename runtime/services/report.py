from __future__ import annotations

import argparse
import os
import runpy
import sys
from pathlib import Path


RUNTIME_ROOT = Path(__file__).resolve().parents[1]
REPORT_CORE_SRC = Path(__file__).resolve().parent / "report_core"


def serve(skill_root: Path, host: str, port: str) -> None:
    os.environ.setdefault("WXT_APP_ROOT", str(skill_root))
    os.environ.setdefault("WXT_HOST", host)
    os.environ.setdefault("WXT_PORT", port)
    os.environ.setdefault("WXT_OUTPUT_ROOT", str(skill_root / "outputs" / "runs"))
    os.environ.setdefault("WXT_FINAL_REPORT_DIR", str(skill_root / "outputs" / "final_reports"))
    sys.path.insert(0, str(REPORT_CORE_SRC))
    sys.argv = ["taobao_report_mvp.web_app", *sys.argv[1:]]
    runpy.run_module("taobao_report_mvp.web_app", run_name="__main__")


def main() -> None:
    parser = argparse.ArgumentParser(description="Shared report runtime service.")
    parser.add_argument("command", nargs="?", default="serve", choices=["serve"])
    parser.add_argument("--skill-root", required=True)
    parser.add_argument("--host", default=os.environ.get("WXT_HOST", "127.0.0.1"))
    parser.add_argument("--port", default=os.environ.get("WXT_PORT", "8799"))
    args = parser.parse_args()
    serve(Path(args.skill_root).resolve(), args.host, str(args.port))


if __name__ == "__main__":
    main()
