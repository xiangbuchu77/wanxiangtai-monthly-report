from __future__ import annotations

import argparse
from pathlib import Path

from .classifier import classify
from .exporter import write_workbook
from .io import read_excel_reports
from .metrics import build_report


DEFAULT_INPUTS = [
    "/Users/gordon/Downloads/店铺经营核心月报_20260615_97559a9f08e1bc31f1ba2ca92a36d53c.xlsx",
    "/Users/gordon/Downloads/店铺流量来源构成月报_20260615_5fa7cfab747853b68b31365e109eb17e.xlsx",
    "/Users/gordon/Downloads/店铺流量来源构成月报旧_20260615_668883f8b115b8f55135ea06fe385aa9.xlsx",
    "/Users/gordon/Downloads/商品经营投产比核心日报_20260615_3e2f1bf4e411096d12da4a0fc4d3c9fd.xlsx",
    "/Users/gordon/Downloads/商品流量来源构成月报_20260615_e7ba4d2aea2e888970721e03a67749e5.xlsx",
    "/Users/gordon/Downloads/商品流量来源构成月报旧_20260615_6336d82654e4d8cc11708995679ca89e.xlsx",
    "/Users/gordon/Downloads/商品整体效果月报_20260615_3a00cf07aa8b2d3da0544bba65937ec8.xlsx",
]
DEFAULT_OUTPUT = "outputs/ecommerce-monthly/淘宝天猫运营标准经营总表_MVP.xlsx"


def main() -> None:
    parser = argparse.ArgumentParser(description="淘宝/天猫运营月报自动化系统 MVP")
    parser.add_argument("inputs", nargs="*", help="Excel文件路径；不传时使用本次样例文件")
    parser.add_argument("--input-dir", type=Path, help="从目录读取所有 .xlsx 文件")
    parser.add_argument("--output", type=Path, default=Path(DEFAULT_OUTPUT), help="输出 Excel 路径")
    parser.add_argument("--period", choices=["month", "week", "half-month"], default="month", help="汇总周期")
    args = parser.parse_args()

    paths = collect_paths(args.inputs, args.input_dir)
    loaded = read_excel_reports(paths)
    classified = [classify(report.path, report.sheet_name, report.raw) for report in loaded]
    result = build_report(classified, period=args.period)
    write_workbook(result, args.output)
    print(f"output={args.output}")
    print(f"input_files={len(paths)}")
    print(f"worksheets={len(loaded)}")


def collect_paths(inputs: list[str], input_dir: Path | None) -> list[Path]:
    paths: list[Path] = []
    if input_dir is not None:
        paths.extend(sorted(input_dir.glob("*.xlsx")))
    if inputs:
        paths.extend(Path(item) for item in inputs)
    if not paths:
        paths = [Path(item) for item in DEFAULT_INPUTS]

    unique: list[Path] = []
    seen: set[Path] = set()
    for path in paths:
        resolved = path.expanduser()
        if resolved not in seen:
            unique.append(resolved)
            seen.add(resolved)
    return unique


if __name__ == "__main__":
    main()
