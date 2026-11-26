"""命令行入口。"""

import argparse
import sys
from pathlib import Path

from .generator import generate_all
from .parser import parse_workbook


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="从工价表总文件生成工序、工艺路线、绩效三份 xlsx")
    parser.add_argument("--input", required=True, help="总文件 xlsx 路径")
    parser.add_argument("--output-dir", default=None, help="输出目录，默认与输入文件同目录")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        info = parse_workbook(args.input)
        output_dir = args.output_dir or str(Path(args.input).resolve().parent)
        written = generate_all(info, output_dir)
    except (FileNotFoundError, ValueError, OSError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1
    for path in written:
        print(f"已生成：{path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
