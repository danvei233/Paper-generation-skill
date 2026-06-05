# AGENTS.md

## 任务目标

继续完成《化工进展》论文制作 skill。这个目录要被完善成一个可靠的论文 DOCX 生成与格式检查工具链，而不是一次性的 prompt。

## 工作顺序

1. 先读：`FORMAT_SPEC.md`、`STYLE_MAP.json`、`reference/《化工进展》论文模板格式（2026版）.docx`。
2. 运行：`scripts/extract_template_style.py`，比较提取结果和 `STYLE_MAP.json`。
3. 补全：`scripts/build_paper_docx.py`，让它能从结构化 YAML/JSON 生成论文 DOCX。
4. 补全：`scripts/check_docx_format.py`，让它能生成可读的 `format_report.md`。
5. 只有当格式工具链可用后，才开始写论文正文。

## 禁止事项

- 不要凭肉眼“差不多”排版。
- 不要直接在最终稿保留 EndNote/Zotero 域代码、超链接、脚注、尾注、批注、文本框。
- 不要把图、公式放进表格或文本框。
- 不要把图片设置成浮动环绕；必须按嵌入型处理。
- 不要让参考文献编号、图表编号、公式编号手工失控。
- 不要使用未在模板允许范围内的中文字体。

## 完成标准

至少实现下面命令可跑通：

```bash
python scripts/extract_template_style.py reference/《化工进展》论文模板格式（2026版）.docx --out output/extracted_style_map.json
python scripts/build_paper_docx.py examples/manuscript.example.yaml --style STYLE_MAP.json --out output/manuscript.docx
python scripts/check_docx_format.py output/manuscript.docx --style STYLE_MAP.json --out output/format_report.md
```

`format_report.md` 必须列出：

- 页面设置是否合格；
- 禁用 Word 特性是否存在；
- 一级/二级/三级标题编号是否合格；
- 图题是否在图下方；
- 表题是否在表上方；
- 图、表、公式编号是否连续；
- 正文中是否有图、表、公式呼应；
- 参考文献是否按引用顺序排列；
- 是否存在域代码/超链接；
- 表格是否疑似三线表。

## 实现偏好

- Python 优先，使用 `python-docx` + `lxml` 直接处理 DOCX/OOXML。
- 所有规则不要写死在脚本里，尽量来自 `STYLE_MAP.json`。
- 生成器和检查器要分离。
- 检查器不要只输出 pass/fail，要给出定位信息和修复建议。
- 对不确定的样式参数写入 `warnings`，不要装作已经确定。
