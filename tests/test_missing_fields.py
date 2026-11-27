"""宽松模式：手写空白模板拆分成三份表，缺失字段留空、表头不减。"""

import pytest

from fixture_factory import make_blank_template_workbook
from wage_converter.generator import (
    generate_all,
    generate_performance_workbook,
    generate_process_workbook,
    generate_route_workbook,
)
from wage_converter.parser import parse_workbook

EXPECTED_CODES = [
    "A01", "A02", "B01", "B02", "C01", "C02",
    "D01", "D02", "E01", "E02",
]
EXPECTED_NAMES = [
    "车工一", "车工二", "手工一", "手工二", "备料一", "备料二",
    "包装一", "包装二", "裁剪一", "裁剪二",
]


@pytest.fixture
def blank_template(tmp_path):
    return make_blank_template_workbook(tmp_path / "blank_template.xlsx")


def test_lenient_parse_keeps_only_named_operations(blank_template):
    info = parse_workbook(blank_template, allow_missing=True)
    assert (info.product_no, info.product_name) == ("", "")
    assert info.source_stem == "blank_template"
    assert [len(s.operations) for s in info.sections.values()] == [2, 2, 2, 2, 2]
    assert [op.code(info.product_no) for op in info.operations] == EXPECTED_CODES

    with pytest.raises(ValueError):
        parse_workbook(blank_template)


def test_codes_and_names_have_no_empty_placeholders(blank_template):
    info = parse_workbook(blank_template, allow_missing=True)
    assert [op.display_name(info.product_name) for op in info.operations] == EXPECTED_NAMES
    assert all("-" not in code for code in EXPECTED_CODES)


def test_missing_price_is_none(blank_template):
    info = parse_workbook(blank_template, allow_missing=True)
    assert all(op.price is None for op in info.operations)


def test_process_file_keeps_all_fields_and_defaults(blank_template):
    info = parse_workbook(blank_template, allow_missing=True)
    ws = generate_process_workbook(info).active
    assert ws.max_row == 11 and ws.max_column == 8
    assert [cell.value for cell in ws[1]] == [
        "工序编号", "工序名称", "报工权限", "报工数配比",
        "不良品项列表", "加工类型", "默认供应商", "任务打卡",
    ]
    assert [ws["A2"].value, ws["B2"].value, ws["C2"].value, ws["D2"].value, ws["E2"].value] == [
        "A01", "车工一", "所有人", 1, "其他",
    ]
    assert [ws["F2"].value, ws["G2"].value, ws["H2"].value] == [None, None, None]


def test_route_file_codes_and_names(blank_template):
    info = parse_workbook(blank_template, allow_missing=True)
    ws = generate_route_workbook(info).active
    assert ws.max_row == 6 and ws.max_column == 3
    assert [ws[f"A{r}"].value for r in range(2, 7)] == ["01", "02", "03", "04", "05"]
    assert [ws[f"B{r}"].value for r in range(2, 7)] == ["车工", "手工", "备料", "包装", "裁剪"]
    assert ws["C2"].value == "车工一，车工二"


def test_performance_and_filename_fall_back(blank_template, tmp_path):
    info = parse_workbook(blank_template, allow_missing=True)
    ws = generate_performance_workbook(info).active
    assert ws.max_row == 11 and ws.max_column == 13
    assert [ws["A2"].value, ws["B2"].value, ws["D2"].value, ws["E2"].value, ws["F2"].value] == [
        None, None, "车工一", "计件", None,
    ]
    written = generate_all(info, tmp_path / "out")
    assert [p.name for p in written] == [
        "1.工序-blank_template.xlsx",
        "2.工艺路线-blank_template.xlsx",
        "3.绩效-blank_template.xlsx",
    ]
