#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""工价表转换工具（单文件版）。

从一份「产品制造工序工价表」总文件生成三份 xlsx：工序（8 列）、
工艺路线（3 列）、绩效（13 列）。不需要 PYTHONPATH，也不需要预先安装依赖；
首次运行时若缺少 openpyxl 会自动尝试安装（需联网）。

用法示例：
    python wage_converter.py --input "9672848#冰包（双肩 条纹）.xlsx"
    python wage_converter.py --input "总文件.xlsx" --output-dir ./output
    python wage_converter.py --input test1.xlsx --allow-missing

--allow-missing 用于手写空白模板：产品号/产品名称/工价可以为空，
缺失字段留空、编号不加多余连字符、文件名回退为源文件名；
不传该参数时为严格模式，缺产品信息会直接报错。

支持 Python 3.7+（推荐 3.10+）；文件按 UTF-8 编码保存。
"""

from __future__ import annotations

import argparse
import importlib
import re
import subprocess
import sys
import warnings
from dataclasses import dataclass, field
from pathlib import Path

load_workbook = None
Workbook = None


def _ensure_openpyxl() -> bool:
    """确保 openpyxl 可用；缺失时用当前 Python 的 pip 自动安装。"""
    global load_workbook, Workbook
    try:
        from openpyxl import Workbook as _Workbook, load_workbook as _load_workbook

        load_workbook, Workbook = _load_workbook, _Workbook
        return True
    except ImportError:
        pass

    print("未检测到 openpyxl，正在自动安装（首次运行需联网）...", file=sys.stderr)
    try:
        subprocess.check_call(
            [
                sys.executable,
                "-m",
                "pip",
                "install",
                "openpyxl",
                "-q",
                "--disable-pip-version-check",
            ]
        )
        importlib.invalidate_caches()
        from openpyxl import Workbook as _Workbook, load_workbook as _load_workbook

        load_workbook, Workbook = _load_workbook, _Workbook
        return True
    except Exception as exc:  # noqa: BLE001 - 需要给出明确的手动安装指引
        print(f"自动安装 openpyxl 失败：{exc}", file=sys.stderr)
        print(f"请手动运行：{sys.executable} -m pip install openpyxl", file=sys.stderr)
        return False


# ---------------------------------------------------------------------------
# 配置常量
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class SectionSpec:
    cn_numeral: str
    key: str
    letter: str
    route_no: int


SECTION_SPECS = (
    SectionSpec("一", "车工", "A", 1),
    SectionSpec("二", "手工", "B", 2),
    SectionSpec("三", "备料", "C", 3),
    SectionSpec("四", "包装", "D", 4),
    SectionSpec("五", "裁剪", "E", 5),
)
SECTION_BY_CN = {spec.cn_numeral: spec for spec in SECTION_SPECS}
SECTION_BY_KEY = {spec.key: spec for spec in SECTION_SPECS}
SECTION_ORDER = tuple(spec.key for spec in SECTION_SPECS)

SECTION_TITLE_RE = re.compile(r"^[一二三四五]、(.*)工序$")
PRODUCT_NAME_RE = re.compile(r"产品名称[:：]\s*(.+)")
PRODUCT_NO_RE = re.compile(r"产品号[:：]\s*(\d+)")

PROCESS_HEADERS = (
    "工序编号",
    "工序名称",
    "报工权限",
    "报工数配比",
    "不良品项列表",
    "加工类型",
    "默认供应商",
    "任务打卡",
)
ROUTE_HEADERS = ("工艺路线编号", "工艺路线名称", "工序列表")
PERFORMANCE_HEADERS = (
    "产品编号",
    "产品名称",
    "产品规格",
    "工序名称",
    "计价方式",
    "工资单价（元）",
    "标准产出",
    "标准工时（秒）",
    "辅助参数01",
    "辅助参数02",
    "辅助参数03",
    "辅助参数04",
    "辅助参数05",
)

REPORT_PERMISSION = "所有人"
REPORT_RATIO = 1
DEFECT_ITEM = "其他"
PAY_METHOD = "计件"

FILE_TEMPLATES = {
    "process": "1.工序-{product_no}#{product_name}.xlsx",
    "route": "2.工艺路线-{product_no}#{product_name}.xlsx",
    "performance": "3.绩效-{product_no}#{product_name}.xlsx",
}
STEM_FALLBACK_TEMPLATES = {
    "process": "1.工序-{stem}.xlsx",
    "route": "2.工艺路线-{stem}.xlsx",
    "performance": "3.绩效-{stem}.xlsx",
}
ILLEGAL_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------
@dataclass
class Operation:
    section_key: str
    index: int
    name: str
    price: float | None = None
    source_row: int = 0

    @property
    def letter(self) -> str:
        return SECTION_BY_KEY[self.section_key].letter

    def code(self, product_no: str) -> str:
        prefix = f"{product_no}-" if product_no else ""
        return f"{prefix}{self.letter}{self.index:02d}"

    def display_name(self, product_name: str) -> str:
        suffix = f"-{product_name}" if product_name else ""
        return f"{self.name}{suffix}"


@dataclass
class Section:
    key: str
    operations: list[Operation] = field(default_factory=list)

    @property
    def spec(self) -> SectionSpec:
        return SECTION_BY_KEY[self.key]


@dataclass
class ProductInfo:
    product_no: str
    product_name: str
    sections: dict[str, Section] = field(default_factory=dict)
    source_stem: str = ""

    @property
    def operations(self) -> list[Operation]:
        result: list[Operation] = []
        for key in SECTION_ORDER:
            section = self.sections.get(key)
            if section:
                result.extend(section.operations)
        return result


# ---------------------------------------------------------------------------
# 解析
# ---------------------------------------------------------------------------
def _text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _find_product_info(ws, *, allow_missing: bool = False) -> tuple[str, str]:
    product_no = ""
    product_name = ""
    for row in ws.iter_rows():
        for cell in row:
            value = cell.value
            if not isinstance(value, str):
                continue
            if not product_no:
                match = PRODUCT_NO_RE.search(value)
                if match:
                    product_no = match.group(1)
            if not product_name:
                match = PRODUCT_NAME_RE.search(value)
                if match:
                    product_name = match.group(1).strip()
            if product_no and product_name:
                return product_no, product_name
    if allow_missing:
        return product_no, product_name
    missing = []
    if not product_name:
        missing.append("产品名称")
    if not product_no:
        missing.append("产品号")
    raise ValueError(f"总文件缺少{'、'.join(missing)}信息（应位于 A4 / D4 附近）")


def _find_titles(ws) -> dict[str, tuple[int, int, SectionSpec]]:
    titles: dict[str, tuple[int, int, SectionSpec]] = {}
    for col in (1, 4):
        for row in range(1, ws.max_row + 1):
            value = ws.cell(row=row, column=col).value
            if not isinstance(value, str):
                continue
            match = SECTION_TITLE_RE.match(value.strip())
            if not match:
                continue
            cn = value.strip()[0]
            spec = SECTION_BY_CN.get(cn)
            if spec is None or spec.key != match.group(1):
                raise ValueError(f"第 {row} 行章节标题无法识别：{value}")
            if spec.key in titles:
                raise ValueError(f"章节「{spec.key}」出现多次")
            titles[spec.key] = (col, row, spec)
    missing = [spec.key for spec in SECTION_SPECS if spec.key not in titles]
    if missing:
        raise ValueError(f"总文件缺少工序章节：{'、'.join(missing)}")
    return titles


def _parse_price(value: object, row: int) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            price = float(value.strip())
        except ValueError:
            warnings.warn(f"第 {row} 行工价「{value}」无法解析，已跳过", stacklevel=2)
            return None
        warnings.warn(f"第 {row} 行工价为文本「{value}」，已转为数值 {price}", stacklevel=2)
        return price
    return None


def parse_sheet(ws, *, allow_missing: bool = False) -> ProductInfo:
    product_no, product_name = _find_product_info(ws, allow_missing=allow_missing)
    titles = _find_titles(ws)

    by_column: dict[int, list[tuple[int, str]]] = {}
    for key, (col, row, _) in titles.items():
        by_column.setdefault(col, []).append((row, key))
    for rows in by_column.values():
        rows.sort()

    sections: dict[str, Section] = {}
    for key in SECTION_ORDER:
        col, title_row, _ = titles[key]
        name_col = col + 1
        price_col = col + 2
        following = [r for r, k in by_column[col] if r > title_row]
        end_row = min(following) if following else ws.max_row + 1

        section = Section(key=key)
        for row in range(title_row + 1, end_row):
            name = _text(ws.cell(row=row, column=name_col).value)
            if not name:
                continue
            price = _parse_price(ws.cell(row=row, column=price_col).value, row)
            if price is None and not allow_missing:
                continue
            section.operations.append(
                Operation(
                    section_key=key,
                    index=len(section.operations) + 1,
                    name=name,
                    price=price,
                    source_row=row,
                )
            )
        sections[key] = section

    return ProductInfo(product_no=product_no, product_name=product_name, sections=sections)


def parse_workbook(path: str | Path, *, allow_missing: bool = False) -> ProductInfo:
    if not _ensure_openpyxl():
        raise RuntimeError("openpyxl 不可用，无法读取 xlsx")
    source = Path(path)
    if not source.is_file():
        raise FileNotFoundError(f"总文件不存在：{source}")
    if source.suffix.lower() != ".xlsx":
        raise ValueError(f"仅支持 .xlsx 文件：{source}")
    # 非只读加载：兼容部分导出文件缺少 dimension 标签的情况
    wb = load_workbook(source, data_only=True)
    try:
        if "Sheet1" not in wb.sheetnames:
            raise ValueError(f"总文件缺少 Sheet1 工作表：{source}")
        info = parse_sheet(wb["Sheet1"], allow_missing=allow_missing)
        info.source_stem = source.stem
        return info
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# 生成
# ---------------------------------------------------------------------------
def _new_sheet(headers: tuple[str, ...]):
    if not _ensure_openpyxl():
        raise RuntimeError("openpyxl 不可用，无法写入 xlsx")
    wb = Workbook()
    ws = wb.active
    ws.append(list(headers))
    return wb, ws


def generate_process_workbook(info: ProductInfo):
    wb, ws = _new_sheet(PROCESS_HEADERS)
    for op in info.operations:
        ws.append(
            [
                op.code(info.product_no),
                op.display_name(info.product_name),
                REPORT_PERMISSION,
                REPORT_RATIO,
                DEFECT_ITEM,
                None,
                None,
                None,
            ]
        )
    return wb


def generate_route_workbook(info: ProductInfo):
    wb, ws = _new_sheet(ROUTE_HEADERS)
    prefix = f"{info.product_no}-" if info.product_no else ""
    suffix = f"-{info.product_name}" if info.product_name else ""
    for route_no, key in enumerate(SECTION_ORDER, start=1):
        section = info.sections[key]
        names = [op.display_name(info.product_name) for op in section.operations]
        ws.append([f"{prefix}{route_no:02d}", f"{key}{suffix}", "，".join(names)])
    return wb


def generate_performance_workbook(info: ProductInfo):
    wb, ws = _new_sheet(PERFORMANCE_HEADERS)
    for op in info.operations:
        ws.append(
            [
                info.product_no or None,
                info.product_name or None,
                None,
                op.display_name(info.product_name),
                PAY_METHOD,
                op.price,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ]
        )
    return wb


def output_filename(kind: str, info: ProductInfo) -> str:
    if not info.product_no and not info.product_name:
        return STEM_FALLBACK_TEMPLATES[kind].format(stem=info.source_stem)
    safe_name = ILLEGAL_FILENAME_CHARS.sub("_", info.product_name)
    return FILE_TEMPLATES[kind].format(product_no=info.product_no, product_name=safe_name)


def generate_all(info: ProductInfo, output_dir: str | Path) -> list[Path]:
    target = Path(output_dir)
    target.mkdir(parents=True, exist_ok=True)
    builders = (
        ("process", generate_process_workbook),
        ("route", generate_route_workbook),
        ("performance", generate_performance_workbook),
    )
    written = []
    for kind, build in builders:
        path = target / output_filename(kind, info)
        build(info).save(path)
        written.append(path)
    return written


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="从工价表总文件生成工序、工艺路线、绩效三份 xlsx",
        epilog=(
            "示例：python wage_converter.py --input 总文件.xlsx "
            "[--output-dir 目录] [--allow-missing]\n"
            "--allow-missing 用于缺产品号/产品名称/工价的手写空白模板"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--input", required=True, help="总文件 xlsx 路径")
    parser.add_argument("--output-dir", default=None, help="输出目录，默认与输入文件同目录")
    parser.add_argument(
        "--allow-missing",
        action="store_true",
        help="允许产品信息或工价缺失（字段留空），用于手写空白模板",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if not _ensure_openpyxl():
        return 1
    try:
        info = parse_workbook(args.input, allow_missing=args.allow_missing)
        output_dir = args.output_dir or str(Path(args.input).resolve().parent)
        written = generate_all(info, output_dir)
    except (FileNotFoundError, ValueError, OSError, RuntimeError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1
    for path in written:
        print(f"已生成：{path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
