"""端到端回归：真实总文件 → 三份输出，并与示例逐项核对。"""

from pathlib import Path

from openpyxl import load_workbook

from wage_converter import cli
from wage_converter.generator import generate_all
from wage_converter.parser import parse_workbook

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "samples" / "9672848#冰包（双肩 条纹）.xlsx"
GOLDEN_PROCESS = ROOT / "samples" / "1.工序-9672848#冰包（双肩 条纹）.xlsx"
GOLDEN_ROUTE = ROOT / "samples" / "2.工艺路线-9672848#冰包（双肩 条纹）.xlsx"
GOLDEN_PERFORMANCE = ROOT / "samples" / "3.绩效-9672848#冰包（双肩 条纹）.xlsx"


def test_real_sample_end_to_end(tmp_path):
    info = parse_workbook(SAMPLE)
    assert (info.product_no, info.product_name) == ("9672848", "冰包（双肩 条纹）")
    assert [len(s.operations) for s in info.sections.values()] == [26, 4, 4, 6, 19]

    process, route, performance = generate_all(info, tmp_path / "out")

    pw = load_workbook(process, read_only=True).active
    assert pw.max_row == 60 and pw.max_column == 8
    assert [c.value for c in pw[1]] == [
        "工序编号", "工序名称", "报工权限", "报工数配比",
        "不良品项列表", "加工类型", "默认供应商", "任务打卡",
    ]
    assert [pw["A2"].value, pw["B2"].value, pw["C2"].value, pw["D2"].value, pw["E2"].value] == [
        "9672848-A01", "双针拉拉尾×2-冰包（双肩 条纹）", "所有人", 1, "其他",
    ]
    assert [pw["F2"].value, pw["G2"].value, pw["H2"].value] == [None, None, None]
    assert pw["A28"].value == "9672848-B01"
    assert pw["A32"].value == "9672848-C01"
    assert pw["A36"].value == "9672848-D01"
    assert pw["A60"].value == "9672848-E19"
    assert pw["B60"].value == "珍珠棉6.0-背带内托-冰包（双肩 条纹）"

    rw = load_workbook(route, read_only=True).active
    assert rw.max_row == 6 and rw.max_column == 3
    assert [rw[f"A{r}"].value for r in range(2, 7)] == [f"9672848-{i:02d}" for i in range(1, 6)]
    assert [rw[f"B{r}"].value for r in range(2, 7)] == [
        "车工-冰包（双肩 条纹）", "手工-冰包（双肩 条纹）", "备料-冰包（双肩 条纹）",
        "包装-冰包（双肩 条纹）", "裁剪-冰包（双肩 条纹）",
    ]
    assert rw["C2"].value.count("，") == 25  # 26 条车工工序

    fw = load_workbook(performance, read_only=True).active
    assert fw.max_row == 60 and fw.max_column == 13
    assert [fw["A2"].value, fw["B2"].value, fw["E2"].value, fw["F2"].value] == [
        "9672848", "冰包（双肩 条纹）", "计件", 0.04,
    ]
    assert fw["F60"].value == 0.012
    assert fw["D28"].value == "翻包-冰包（双肩 条纹）"


def test_golden_regression_on_shared_fields(tmp_path):
    """示例为 6 列、路线 04/03/02/05/01；只比对双方共有的业务内容。"""
    info = parse_workbook(SAMPLE)
    process, route, performance = generate_all(info, tmp_path / "out")
    generated_p = load_workbook(process, read_only=True).active
    generated_r = load_workbook(route, read_only=True).active
    generated_f = load_workbook(performance, read_only=True).active

    golden_p = load_workbook(GOLDEN_PROCESS, read_only=True).active
    golden_r = load_workbook(GOLDEN_ROUTE, read_only=True).active
    golden_f = load_workbook(GOLDEN_PERFORMANCE, read_only=True).active

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


def test_cli_success_and_failure(tmp_path, capsys):
    out = tmp_path / "cli_out"
    assert cli.main(["--input", str(SAMPLE), "--output-dir", str(out)]) == 0
    assert len(list(out.glob("*.xlsx"))) == 3
    assert "工序" in capsys.readouterr().out

    assert cli.main(["--input", str(tmp_path / "missing.xlsx"), "--output-dir", str(out)]) == 1
    assert "不存在" in capsys.readouterr().err
