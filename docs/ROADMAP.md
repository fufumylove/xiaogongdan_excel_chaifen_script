# Roadmap — 实施路线与验收

| 阶段 | 目标 | 产出 | 验收标准 | 验证命令 | 状态 |
| --- | --- | --- | --- | --- | --- |
| S1 骨架 | 建立包结构与依赖 | `src/wage_converter/{__init__,config,models,cli}.py`、`requirements.txt` | 模块可导入、常量齐全 | `python -c "import wage_converter.config"` | 完成 |
| S2 解析 | 先写失败测试，再实现解析 | `tests/test_parser.py`、`parser.py`、`fixture_factory.py` | 4 项测试通过：章节/过滤/补零/容错/异常 | `pytest -q tests/test_parser.py` | 完成 |
| S3 生成 | 先写失败测试，再实现三份输出 | `tests/test_generator.py`、`generator.py` | 4 项测试通过：表头/列序/编号/拼接/文件名 | `pytest -q tests/test_generator.py` | 完成 |
| S4 端到端 | CLI 串联与真实样本回归 | `tests/test_e2e.py`、`samples/`、完整 CLI | 3 项测试通过；59/59/5 行与示例共有字段一致 | `pytest -q` | 完成 |
| S5 文档 | 交付 PRD/技术/Roadmap/流程记录 | `docs/*.md`、`README.md`、`.gitignore` | 文档与实现一致，可提交 GitHub | 人工审阅 + `pytest -q` | 完成 |

## 阶段依赖
S2 依赖 S1 的常量与模型；S3 依赖 S2 的解析结果；S4 依赖 S3 的生成器；S5 汇总前四阶段产出。

## 验收口径
- 每阶段均遵循「红灯测试 → 最小实现 → 绿灯 → 重构检查」。
- 全量测试 11 项通过，真实样本生成三份文件且关键值与示例一致。
- 与示例的两处有意差异（8 列、路线 01→05）已在 PRD 与 README 中记录。

## 风险与回退
- 输入版式漂移：用户确认版式一致；解析器对缺章节/缺产品信息直接报错，避免静默错误。
- 依赖版本差异：`requirements.txt` 固定版本。
- 任一阶段验收失败即停留在该阶段修复，不推进后续阶段。
