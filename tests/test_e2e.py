"""端到端回归：真实总文件 → 三份输出，并与示例核对共有字段。"""

from pathlib import Path

import pytest
from openpyxl import load_workbook

from wage_converter import cli
from wage_converter.config import SECTION_ORDER
from wage_converter.generator import generate_all
from wage_converter.parser import parse_workbook

ROOT = Path(__file__).resolve().parents[1]
SAMPLES_DIR = ROOT / "samples"

OLD_CASE = {
    "no": "9672848",
    "name": "冰包（双肩 条纹）",
    "counts": [26, 4, 4, 6, 19],
    "first_name": "双针拉拉尾×2",
    "first_price": 0.04,
    "manual_first": "翻包",
    "last_name": "珍珠棉6.0-背带内托",
    "last_price": 0.012,
}

NEW_CASE = {
    "no": "9894712",
    "name": "冰包（格子）",
    "counts": [23, 3, 4, 6, 16],
    "first_name": "肩带车日子扣",
    "first_price": 0.07,
    "manual_first": "穿肩带",
    "last_name": "珍珠棉6.0mm-顶底片内托*2",
    "last_price": 0.04,
}

SAMPLES = [pytest.param(OLD_CASE, id="9672848"), pytest.param(NEW_CASE, id="9894712")]

LETTERS = ("A", "B", "C", "D", "E")


def source_path(case: dict) -> Path:
    return SAMPLES_DIR / f"{case['no']}#{case['name']}.xlsx"


def section_start_rows(case: dict) -> dict[str, int]:
    """各章节第一条数据在工序表中的行号（表头占第 1 行）。"""
    starts = {}
    row = 2
    for letter, count in zip(LETTERS, case["counts"]):
        starts[letter] = row
        row += count
    return starts


@pytest.mark.parametrize("case", SAMPLES)
def test_real_sample_end_to_end(case, tmp_path):
    total = sum(case["counts"])
    info = parse_workbook(source_path(case))
    assert (info.product_no, info.product_name) == (case["no"], case["name"])
    assert [len(s.operations) for s in info.sections.values()] == case["counts"]

    process, route, performance = generate_all(info, tmp_path / "out")
    starts = section_start_rows(case)

    pw = load_workbook(process, read_only=True).active
    assert pw.max_row == total + 1 and pw.max_column == 8
    assert [c.value for c in pw[1]] == [
        "工序编号", "工序名称", "报工权限", "报工数配比",
        "不良品项列表", "加工类型", "默认供应商", "任务打卡",
    ]
    assert [pw["A2"].value, pw["B2"].value, pw["C2"].value, pw["D2"].value, pw["E2"].value] == [
        f"{case['no']}-A01", f"{case['first_name']}-{case['name']}", "所有人", 1, "其他",
    ]
    assert [pw["F2"].value, pw["G2"].value, pw["H2"].value] == [None, None, None]
    for letter in LETTERS:
        assert pw[f"A{starts[letter]}"].value == f"{case['no']}-{letter}01"
    assert pw[f"A{total + 1}"].value == f"{case['no']}-E{case['counts'][-1]:02d}"
    assert pw[f"B{total + 1}"].value == f"{case['last_name']}-{case['name']}"

    rw = load_workbook(route, read_only=True).active
    assert rw.max_row == 6 and rw.max_column == 3
    assert [rw[f"A{r}"].value for r in range(2, 7)] == [
        f"{case['no']}-{i:02d}" for i in range(1, 6)
    ]
    assert [rw[f"B{r}"].value for r in range(2, 7)] == [
        f"{key}-{case['name']}" for key in SECTION_ORDER
    ]
    assert rw["C2"].value.count("，") == case["counts"][0] - 1

    fw = load_workbook(performance, read_only=True).active
    assert fw.max_row == total + 1 and fw.max_column == 13
    assert [fw["A2"].value, fw["B2"].value, fw["C2"].value, fw["E2"].value, fw["F2"].value] == [
        case["no"], case["name"], None, "计件", case["first_price"],
    ]
    assert [fw[f"D{r}"].value for r in range(2, total + 2)] == [
        pw[f"B{r}"].value for r in range(2, total + 2)
    ]
    assert fw[f"D{starts['B']}"].value == f"{case['manual_first']}-{case['name']}"
    assert fw[f"F{total + 1}"].value == case["last_price"]


def test_golden_regression_on_shared_fields(tmp_path):
    """旧示例为 6 列、路线 04/03/02/05/01；只比对双方共有的业务内容。

    新文件夹的三份文件是系统导出模板（工序 14 列/路线 5 列含父级行/
    绩效 18 列且倒序、缺裁剪 E16），用户确认只作语义参考、不复刻，
    因此不纳入逐格回归。
    """
    case = OLD_CASE
    info = parse_workbook(source_path(case))
    process, route, performance = generate_all(info, tmp_path / "out")
    generated_p = load_workbook(process, read_only=True).active
    generated_r = load_workbook(route, read_only=True).active
    generated_f = load_workbook(performance, read_only=True).active

    golden_p = load_workbook(SAMPLES_DIR / f"1.工序-{case['no']}#{case['name']}.xlsx", read_only=True).active
    golden_r = load_workbook(SAMPLES_DIR / f"2.工艺路线-{case['no']}#{case['name']}.xlsx", read_only=True).active
    golden_f = load_workbook(SAMPLES_DIR / f"3.绩效-{case['no']}#{case['name']}.xlsx", read_only=True).active

    for row in range(1, 61):
        assert generated_p.cell(row, 1).value == golden_p.cell(row, 1).value
        assert generated_p.cell(row, 2).value == golden_p.cell(row, 2).value
        assert generated_f.cell(row, 4).value == golden_f.cell(row, 4).value
        assert generated_f.cell(row, 6).value == golden_f.cell(row, 6).value
    assert {generated_r.cell(r, 2).value for r in range(2, 7)} == {
        golden_r.cell(r, 2).value for r in range(2, 7)
    }
    assert {generated_r.cell(r, 3).value for r in range(2, 7)} == {
        golden_r.cell(r, 3).value for r in range(2, 7)
    }


@pytest.mark.parametrize("case", SAMPLES)
def test_cli_success_and_failure(case, tmp_path, capsys):
    out = tmp_path / "cli_out"
    assert cli.main(["--input", str(source_path(case)), "--output-dir", str(out)]) == 0
    assert len(list(out.glob("*.xlsx"))) == 3
    assert "工序" in capsys.readouterr().out

    assert cli.main(["--input", str(tmp_path / "missing.xlsx"), "--output-dir", str(out)]) == 1
    assert "不存在" in capsys.readouterr().err
