"""版式常量与输出规则配置。"""

import re
from dataclasses import dataclass


@dataclass(frozen=True)
class SectionSpec:
    """一类工序的固定元数据。"""

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

ILLEGAL_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')
