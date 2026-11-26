"""生成器测试：先于实现编写（TDD 红灯）。"""

from fixture_factory import make_source_workbook
from wage_converter.config import (
    PERFORMANCE_HEADERS,
    PROCESS_HEADERS,
    ROUTE_HEADERS,
)
from wage_converter.generator import (
    generate_all,
    generate_performance_workbook,
    generate_process_workbook,
    generate_route_workbook,
)
from wage_converter.parser import parse_workbook


def sample_info(tmp_path, **kwargs):
    options = {
        "product_no": "12345",
        "product_name": "测试包（红色）",
        "counts": {"车工": 10, "手工": 2, "备料": 1, "包装": 2, "裁剪": 1},
    }
    options.update(kwargs)
    path = make_source_workbook(tmp_path / "source.xlsx", **options)
    return parse_workbook(path)


def rows(ws):
    return [list(row) for row in ws.iter_rows(values_only=True)]


def test_process_sheet_headers_and_values(tmp_path):
    info = sample_info(tmp_path)
    data = rows(generate_process_workbook(info).active)
    assert data[0] == list(PROCESS_HEADERS)
    assert data[1] == ["12345-A01", "车工步骤1-测试包（红色）", "所有人", 1, "其他", None, None, None]
    assert data[10][0] == "12345-A10"
    assert data[11][0] == "12345-B01"
    assert len(data) == 17  # 表头 + 16 条工序


def test_route_sheet_order_and_join(tmp_path):
    info = sample_info(tmp_path)
    data = rows(generate_route_workbook(info).active)
    assert data[0] == list(ROUTE_HEADERS)
    assert [r[0] for r in data[1:]] == [f"12345-{i:02d}" for i in range(1, 6)]
    assert [r[1] for r in data[1:]] == [
        "车工-测试包（红色）",
        "手工-测试包（红色）",
        "备料-测试包（红色）",
        "包装-测试包（红色）",
        "裁剪-测试包（红色）",
    ]
    assert data[1][2] == "，".join(f"车工步骤{i}-测试包（红色）" for i in range(1, 11))
    assert data[2][2] == "手工步骤1-测试包（红色），手工步骤2-测试包（红色）"


def test_performance_sheet_values(tmp_path):
    info = sample_info(tmp_path)
    data = rows(generate_performance_workbook(info).active)
    assert data[0] == list(PERFORMANCE_HEADERS)
    assert data[1] == [
        "12345", "测试包（红色）", None, "车工步骤1-测试包（红色）",
        "计件", 0.1, None, None, None, None, None, None, None,
    ]
    assert data[11][0] == "12345"
    assert data[-1][3] == "裁剪步骤1-测试包（红色）"
    assert len(data) == 17


def test_generate_all_writes_three_files_and_sanitizes_names(tmp_path):
    info = sample_info(tmp_path, product_name="测试/包:红色")
    written = generate_all(info, tmp_path / "out")
    assert [p.name for p in written] == [
        "1.工序-12345#测试_包_红色.xlsx",
        "2.工艺路线-12345#测试_包_红色.xlsx",
        "3.绩效-12345#测试_包_红色.xlsx",
    ]
    assert all(p.is_file() for p in written)
    # 文件名清洗不影响单元格内容
    from openpyxl import load_workbook

    ws = load_workbook(written[0], read_only=True).active
    assert ws["B2"].value == "车工步骤1-测试/包:红色"
    # 再次生成覆盖同名文件
    again = generate_all(info, tmp_path / "out")
    assert again == written
