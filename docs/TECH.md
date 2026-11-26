# 技术文档 + 架构文档

## 技术选型
Python 3.13、openpyxl 3.1.5（xlsx 读写）、pytest 8.4.2（测试）。无数据库、无网络调用。

## 数据流

```
总文件 xlsx
   └─ parser.parse_workbook → ProductInfo(product_no, product_name, sections)
        └─ generator.generate_all
             ├─ 1.工序-…xlsx   （8 列）
             ├─ 2.工艺路线-…xlsx（3 列）
             └─ 3.绩效-…xlsx   （13 列）
```

## 模块职责
- `config.py`：章节映射（车工 A/01、手工 B/02、备料 C/03、包装 D/04、裁剪 E/05）、表头常量、文件名模板、正则。
- `models.py`：`Operation`（章节、有效序号、名称、工价、来源行；派生编号与拼接名）、`Section`、`ProductInfo`（按 SECTION_ORDER 汇总全部工序）。
- `parser.py`：定位产品信息与五个章节；按列区块提取；过滤无效行；工价容错；缺项报错。
- `generator.py`：三份工作簿构建；文件名非法字符清洗；输出目录创建与覆盖写入。
- `cli.py`：参数解析、调用编排、异常转友好信息并返回退出码。

## 解析算法
1. `load_workbook(data_only=True, read_only=True)` 读取 Sheet1。
2. 扫描全部单元格，用 `产品名称[:：]\s*(.+)` 与 `产品号[:：]\s*(\d+)` 提取产品信息；缺失即报错。
3. 扫描 A、D 两列匹配 `^[一二三四五]、(.*)工序$` 定位章节；校验中文序号与章节名一致，重复或缺失即报错。
4. 左区（A 列标题）取 B/C 为内容/工价；右区（D 列标题）取 E/F。扫描至同列下一个章节标题或表尾。
5. 内容为空或工价为空/不可解析的行跳过；文本工价转 float 并告警；有效序号从 1 重排。
6. 返回 ProductInfo，工序顺序固定为车工→裁剪。

## 字段映射
| 输出 | 关键列 | 来源/规则 |
| --- | --- | --- |
| 工序 | 工序编号 | `{产品号}-{字母}{index:02d}` |
| 工序 | 工序名称 | `Operation.display_name()` |
| 工序 | 报工权限/配比/不良品项 | 常量：所有人 / 1 / 其他 |
| 路线 | 编号/名称 | `{产品号}-{01..05}` / `{章节}-{产品名}` |
| 路线 | 工序列表 | 章节内名称以 `，` 连接 |
| 绩效 | 产品编号/名称 | ProductInfo 字段 |
| 绩效 | 工资单价 | `Operation.price`（数值） |

## 异常与边界
- 文件不存在、后缀非 xlsx、无 Sheet1、缺产品信息或章节：抛出明确异常；CLI 打印错误并返回 1。
- 部分系统导出的 xlsx 缺少 `<dimension>` 标签：解析采用非只读模式加载，避免只读模式下 `max_row=None` 导致章节扫描失败。
- 工价 0 保留；文本工价转换并告警；不可解析的工价跳过并告警。
- 文件名非法字符 `[\\/:*?"<>|]` 仅替换为 `_`，单元格内容保持原样。
- 条目超过 99 时序号自然扩位（如 100）；输出目录不存在自动创建；同名文件覆盖。

## 测试策略
- `tests/test_parser.py`：产品信息、章节定位、空行过滤、补零、文本工价、缺失章节/信息。
- `tests/test_generator.py`：三份表头与关键值、路线 01→05、文件名清洗、覆盖写入。
- `tests/test_e2e.py`：真实样本 59/59/5 行与关键单元格、与示例共有字段回归、CLI 成功与失败路径。
- 合成总文件由 `tests/fixture_factory.py` 在 pytest 临时目录生成；真实样本存于 `samples/`。
- 命令：`pytest -q`（当前 14 项全部通过，覆盖 9672848 与 9894712 两个真实样本）。

## 运行
```bash
pip install -r requirements.txt
python -m wage_converter.cli --input "总文件.xlsx" [--output-dir 目录]
```
