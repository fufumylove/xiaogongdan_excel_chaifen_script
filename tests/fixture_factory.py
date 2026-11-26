"""在临时目录生成合成总文件，避免把二进制样本写入仓库。"""

import re
import zipfile
from pathlib import Path

from openpyxl import Workbook

SECTIONS = (
    ("一", "车工", "A"),
    ("二", "手工", "D"),
    ("三", "备料", "D"),
    ("四", "包装", "D"),
    ("五", "裁剪", "D"),
)


def make_source_workbook(
    path: Path,
    *,
    product_no: str = "12345",
    product_name: str = "测试包（红色）",
    counts: dict[str, int] | None = None,
    blank_rows: bool = True,
    text_price: bool = False,
    omit_section: str | None = None,
) -> Path:
    """按真实版式生成一个总文件：左列为车工，右列依次为其余四章。"""
    counts = counts or {"车工": 3, "手工": 2, "备料": 1, "包装": 2, "裁剪": 1}
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "产品制造工序工价表"
    ws["A4"] = f"产品名称：{product_name}"
    ws["D4"] = f"产品号：{product_no}"

    next_left = 5
    next_right = 5
    for cn, key, col in SECTIONS:
        if key == omit_section:
            continue
        row = next_left if col == "A" else next_right
        ws[f"{col}{row}"] = f"{cn}、{key}工序"
        ws[f"{'C' if col == 'A' else 'F'}{row}"] = "工价"
        name_col = "B" if col == "A" else "E"
        price_col = "C" if col == "A" else "F"
        seq_col = "A" if col == "A" else "D"
        for i in range(1, counts[key] + 1):
            data_row = row + i
            ws[f"{seq_col}{data_row}"] = i
            ws[f"{name_col}{data_row}"] = f"{key}步骤{i}"
            price = 0.1 * i
            ws[f"{price_col}{data_row}"] = str(price) if text_price and i == 1 else price
        if blank_rows:
            # 纯序号行、空内容行：解析时必须跳过
            ghost = row + counts[key] + 1
            ws[f"{seq_col}{ghost}"] = counts[key] + 1
        next_row = row + counts[key] + 2
        if col == "A":
            next_left = next_row
        else:
            next_right = next_row

    wb.save(path)
    return path


def strip_dimension(path: Path) -> Path:
    """复制 xlsx 并删除 sheet1 的 dimension 标签，模拟某些导出文件。"""
    target = path.with_name(f"{path.stem}-nodim.xlsx")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(target, "w") as result:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                data = re.sub(r"<dimension\b[^>]*/>", "", text, count=1).encode("utf-8")
            result.writestr(item, data)
    return target
