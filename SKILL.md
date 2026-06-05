---
name: huagongjinzhan-paper-skill
description: Generate and validate Chinese journal manuscript DOCX files against the 2026 Chemical Industry and Engineering Progress template. Focuses on page setup, headings, captions, references, field-code removal, and figure/table/equation numbering.
---

# 《化工进展》论文制作 Skill

## 适用场景

当用户要写、改、排版、检查《化工进展》或类似中文科技期刊论文时，使用本 skill。

## 核心原则

- 先解析模板，再生成文档。
- 先结构化内容，再统一排版。
- 图、表、公式、参考文献编号必须自动化。
- 默认期刊格式为《化工进展》；未显式指定其他期刊时，一律按 `STYLE_MAP.json` 中的《化工进展》规则生成和检查。
- 页面、段落、run、表格的 raw OOXML 细节必须按 `STYLE_MAP.json.strict_layout` 生成和检查，包括 `pgSz/pgMar/docGrid`、`w:sz/w:szCs`、字符间距、缩进、行距、表宽、单元格边距和三线表边框。
- 参考文献按 `reference/official/hgjz_reference_format_2024.docx` 和 `docs/REFERENCE_FORMAT_HGJZ_2024.md` 执行。
- 正文和附录中的引用使用同一套上标编号；文后参考文献按首次引用顺序列出，序号后空一字格。
- 中文关键词按 `reference/official/hgjz_standard_keywords_2019.json` 检查，默认至少 3 个命中标准关键词库；如需重建词库，运行 `scripts/extract_standard_keywords.py`。
- 语言风格按 `docs/LANGUAGE_STYLE_RULES.md` 执行；规则来源和外部依赖见 `docs/HGJZ_REQUIREMENTS_MATRIX.md`、`docs/OFFICIAL_SOURCES.md`。
- 最终稿必须纯文本引用，不保留 EndNote/Zotero 域代码。
- 生成 DOCX 后必须运行格式检查，并尽可能渲染页面做视觉验收。

## 输入

推荐输入为 YAML/JSON：

- 文章类型；
- DOI；
- 中英文题名；
- 作者和单位；
- 中英文摘要和关键词；
- 正文章节；
- 图、表、公式；
- 符号说明；
- 参考文献。
  推荐参考文献使用结构化对象；如果只提供 `text`，工具只做格式检查，不保证能自动修正所有著录项目。

## 输出

- `manuscript.docx`：投稿/修改稿 DOCX；
- `format_report.md`：格式检查报告；
- `render_report.md`：Word/PDF 渲染复核报告；
- 可选 `plain_text_submission.docx`：去域代码、去超链接后的提交稿。

## 必走流程

1. 读取 `STYLE_MAP.json`、`docs/HGJZ_REQUIREMENTS_MATRIX.md`、`docs/LANGUAGE_STYLE_RULES.md`。
2. 读取结构化论文源文件。
3. 生成 DOCX。
4. 检查禁用 Word 特性、页面设置 raw OOXML、标题、图表公式编号、参考文献编号、图题/表题位置、前置部分 run 级字体/字号/字符间距。
5. 检查语言风格、摘要、关键词、单位符号、结论/结语。
6. 输出 `format_report.md`。
7. 用 Word 导出 PDF，并渲染页面 PNG 做版面验收，输出 `render_report.md`。
8. 用 `scripts/flatten_word_fields.py` 从工作稿生成纯文本提交稿副本，再重新运行格式检查。
9. 对正式交付稿进行人工目检、真实图片内部质量复核和 MathType 兼容确认。
