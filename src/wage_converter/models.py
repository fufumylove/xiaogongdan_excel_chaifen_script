"""解析结果的数据结构。"""

from dataclasses import dataclass, field

from .config import SECTION_BY_KEY, SECTION_ORDER


@dataclass
class Operation:
    """一条有效工序：序号按章节内有效顺序重新分配。"""

    section_key: str
    index: int
    name: str
    price: float
    source_row: int = 0

    @property
    def letter(self) -> str:
        return SECTION_BY_KEY[self.section_key].letter

    def code(self, product_no: str) -> str:
        return f"{product_no}-{self.letter}{self.index:02d}"

    def display_name(self, product_name: str) -> str:
        return f"{self.name}-{product_name}"


@dataclass
class Section:
    """一个工序章节及其条目。"""

    key: str
    operations: list[Operation] = field(default_factory=list)

    @property
    def spec(self):
        return SECTION_BY_KEY[self.key]


@dataclass
class ProductInfo:
    """一份总文件的解析结果。"""

    product_no: str
    product_name: str
    sections: dict[str, Section] = field(default_factory=dict)

    @property
    def operations(self) -> list[Operation]:
        """按工艺路线顺序（车工→裁剪）返回全部工序。"""
        result: list[Operation] = []
        for key in SECTION_ORDER:
            section = self.sections.get(key)
            if section:
                result.extend(section.operations)
        return result
