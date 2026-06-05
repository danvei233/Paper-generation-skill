# 给 Codex 复制的启动提示

继续完成论文制作 skill 工作。

你正在处理一个《化工进展》论文 DOCX 生成与格式检查工具链。请先阅读：

1. `AGENTS.md`
2. `SKILL.md`
3. `FORMAT_SPEC.md`
4. `STYLE_MAP.json`
5. `docs/EXTRACTION_NOTES.md`
6. `reference/《化工进展》论文模板格式（2026版）.docx`

不要直接开始写论文正文。你的第一阶段任务是完善这个 skill：

- 补全 `scripts/extract_template_style.py`，尽可能精确提取模板样式；
- 补全 `scripts/build_paper_docx.py`，从结构化 YAML 生成论文 DOCX；
- 补全 `scripts/check_docx_format.py`，输出 `format_report.md`；
- 把规则尽量放在 `STYLE_MAP.json`，不要散落在脚本里；
- 确保最终稿无 EndNote/Zotero 域代码、无超链接、无文本框、无批注、无页眉页脚；
- 图题必须在图下方，表题必须在表上方；
- 图、表、公式、参考文献编号必须按出现顺序自动检查。

完成后运行：

```bash
pip install -r requirements.txt
python scripts/extract_template_style.py reference/《化工进展》论文模板格式（2026版）.docx --out output/extracted_style_map.json
python scripts/build_paper_docx.py examples/manuscript.example.yaml --style STYLE_MAP.json --out output/manuscript.docx
python scripts/check_docx_format.py output/manuscript.docx --style STYLE_MAP.json --out output/format_report.md
```

如果发现模板样式存在不确定项，不要瞎猜，写入 `docs/EXTRACTION_NOTES.md` 并在 `format_report.md` 中警告。
