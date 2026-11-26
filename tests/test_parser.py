"""解析器测试：先于实现编写（TDD 红灯）。"""

import warnings

import pytest

from fixture_factory import make_source_workbook
from wage_converter.parser import parse_workbook


def test_parse_product_info_and_sections(tmp_path):
    path = make_source_workbook(
        tmp_path / "source.xlsx",
        product_no="98765",
        product_name="测试包（红色）",
        counts={"车工": 10, "手工": 2, "备料": 1, "包装": 12, "裁剪": 3},
    )
    info = parse_workbook(path)

    assert info.product_no == "98765"
    assert info.product_name == "测试包（红色）"
    assert [s.key for s in info.sections.values()] == ["车工", "手工", "备料", "包装", "裁剪"]
    assert [len(s.operations) for s in info.sections.values()] == [10, 2, 1, 12, 3]

    operations = info.operations
    assert len(operations) == 28
    assert operations[0].code(info.product_no) == "98765-A01"
    assert operations[8].code(info.product_no) == "98765-A09"
    assert operations[9].code(info.product_no) == "98765-A10"
    assert operations[10].code(info.product_no) == "98765-B01"
    assert operations[12].code(info.product_no) == "98765-C01"
    assert operations[-4].code(info.product_no) == "98765-D12"
    assert operations[-1].code(info.product_no) == "98765-E03"
    assert operations[0].display_name(info.product_name) == "车工步骤1-测试包（红色）"
    assert operations[0].price == pytest.approx(0.1)


def test_text_price_is_converted_with_warning(tmp_path):
    path = make_source_workbook(tmp_path / "text_price.xlsx", text_price=True)
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        info = parse_workbook(path)
    assert info.operations[0].price == pytest.approx(0.1)
    assert any("工价" in str(w.message) for w in caught)


def test_missing_section_raises(tmp_path):
    path = make_source_workbook(tmp_path / "missing.xlsx", omit_section="裁剪")
    with pytest.raises(ValueError, match="裁剪"):
        parse_workbook(path)


def test_missing_product_info_raises(tmp_path):
    path = make_source_workbook(tmp_path / "no_info.xlsx")
    from openpyxl import load_workbook

    wb = load_workbook(path)
    wb.active["A4"] = None
    wb.active["D4"] = None
    wb.save(path)
    with pytest.raises(ValueError, match="产品名称"):
        parse_workbook(path)
