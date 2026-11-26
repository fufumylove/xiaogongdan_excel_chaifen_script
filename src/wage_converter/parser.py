"""从工价表总文件解析产品信息与五个工序章节。"""

import warnings
from pathlib import Path

from openpyxl import load_workbook

from .config import (
    PRODUCT_NAME_RE,
    PRODUCT_NO_RE,
    SECTION_BY_CN,
    SECTION_ORDER,
    SECTION_SPECS,
    SECTION_TITLE_RE,
)
from .models import Operation, ProductInfo, Section


def _text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _find_product_info(ws) -> tuple[str, str]:
    """扫描全部单元格提取产品号与产品名称。"""
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
    missing = []
    if not product_name:
        missing.append("产品名称")
    if not product_no:
        missing.append("产品号")
    raise ValueError(f"总文件缺少{'、'.join(missing)}信息（应位于 A4 / D4 附近）")


def _find_titles(ws) -> dict[str, tuple[int, int, object]]:
    """定位五个章节标题，返回 key -> (列号, 行号, spec)。"""
    titles: dict[str, tuple[int, int, object]] = {}
    for col in (1, 4):  # A 列为左区，D 列为右区
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


def parse_sheet(ws) -> ProductInfo:
    """解析单个工作表。"""
    product_no, product_name = _find_product_info(ws)
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
            if price is None:
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


def parse_workbook(path: str | Path) -> ProductInfo:
    """读取总文件并返回解析结果；文件不可读时抛出明确异常。"""
    source = Path(path)
    if not source.is_file():
        raise FileNotFoundError(f"总文件不存在：{source}")
    if source.suffix.lower() != ".xlsx":
        raise ValueError(f"仅支持 .xlsx 文件：{source}")
    wb = load_workbook(source, data_only=True, read_only=True)
    try:
        if "Sheet1" not in wb.sheetnames:
            raise ValueError(f"总文件缺少 Sheet1 工作表：{source}")
        return parse_sheet(wb["Sheet1"])
    finally:
        wb.close()
