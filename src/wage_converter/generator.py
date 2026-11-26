"""由解析结果生成工序、工艺路线、绩效三份工作簿。"""

from pathlib import Path

from openpyxl import Workbook

from .config import (
    DEFECT_ITEM,
    FILE_TEMPLATES,
    ILLEGAL_FILENAME_CHARS,
    PAY_METHOD,
    PERFORMANCE_HEADERS,
    PROCESS_HEADERS,
    REPORT_PERMISSION,
    REPORT_RATIO,
    ROUTE_HEADERS,
    SECTION_ORDER,
)
from .models import ProductInfo


def _new_sheet(headers: tuple[str, ...]) -> tuple[Workbook, object]:
    wb = Workbook()
    ws = wb.active
    ws.append(list(headers))
    return wb, ws


def generate_process_workbook(info: ProductInfo) -> Workbook:
    """工序表：8 列，按工序编号顺序排列。"""
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


def generate_route_workbook(info: ProductInfo) -> Workbook:
    """工艺路线表：01→05 顺序，工序列表以中文逗号连接。"""
    wb, ws = _new_sheet(ROUTE_HEADERS)
    operations = info.operations
    cursor = 0
    for route_no, key in enumerate(SECTION_ORDER, start=1):
        section = info.sections[key]
        names = [op.display_name(info.product_name) for op in section.operations]
        ws.append(
            [
                f"{info.product_no}-{route_no:02d}",
                f"{key}-{info.product_name}",
                "，".join(names),
            ]
        )
        cursor += len(names)
    assert cursor == len(operations)
    return wb


def generate_performance_workbook(info: ProductInfo) -> Workbook:
    """绩效表：13 列，顺序与工序表一致，工价写为数值。"""
    wb, ws = _new_sheet(PERFORMANCE_HEADERS)
    for op in info.operations:
        ws.append(
            [
                info.product_no,
                info.product_name,
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
    """文件名中的非法字符替换为下划线，单元格内容不受影响。"""
    safe_name = ILLEGAL_FILENAME_CHARS.sub("_", info.product_name)
    return FILE_TEMPLATES[kind].format(product_no=info.product_no, product_name=safe_name)


def generate_all(info: ProductInfo, output_dir: str | Path) -> list[Path]:
    """生成三份文件并返回其路径；同名文件会被覆盖。"""
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
