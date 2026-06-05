# EXTRACTION_NOTES.md

## 已确认

- 页面尺寸和页边距可从 DOCX section 中直接提取。
- 模板 `sectPr` 中 `docGrid` 为 `type=linesAndChars`、`linePitch=324`、`charSpace=4294967090`，已写入 `STYLE_MAP.json`。
- `extract_template_style.py` 现在输出 `strict_ooxml`：包含 raw `pgSz/pgMar/docGrid`、Normal/docDefaults、关键段落前若干 run 的 `rFonts/sz/szCs/spacing/b`。
- 模板中大量说明以红字普通段落存在，不能全部当作正文样式。
- 模板没有稳定命名的 Word 样式；大量格式是直接格式化，因此需要通过段落内容模式识别样式。
- 中文摘要、关键词、中图分类号行存在 run 级差异：标签为黑体，正文/值为楷体，均为 9 pt 并有 `w:spacing=1`。
- 英文摘要/关键词部分 run 未直接写入 `w:rFonts/w:sz`，但写入了 `w:szCs` 和字符间距；生成器按有效 Times New Roman 字体和 `STYLE_MAP.json` 合同显式写入，避免依赖隐式继承。
- 图题、表题、公式编号、参考文献等规则在模板红字中比较明确。
- 模板图片均为 `wp:inline`，未见 `wp:anchor`；生成器按嵌入型插图。
- 模板数据表使用单元格边框模拟三线表：表头顶线、表头下线和表底线，无竖线；已写入 `STYLE_MAP.json` 并由检查器验证。

## 不确定/需要 Codex 继续确认

1. `气泡泵压降模型的分析与优化` 的中文题名字体显示为 `KYGS;Times New Roman`，这可能是转换/缺字造成的，并非可泛化字体。生成器应优先使用宋体/黑体/仿宋/楷体体系。
2. 首页“收稿日期/基金/作者简介/通信作者”模板要求用题目的脚注形式，但最终工具链按提交稿禁用脚注处理，当前生成器使用普通段落模拟首页注释。若期刊系统强制要求 Word 脚注，需要单独生成工作稿变体。
3. 参考文献著录规则模板只给部分示例，没有包含全部文献类型，需补全 CSL 或专用 formatter。
4. 公式要求 MathType；当前生成器可写入 OMML 公式对象并编号，但正式投稿前需要人工或工具确认 MathType/期刊系统兼容。
5. 模板未给出明确附录样式；当前 `STYLE_MAP.json` 和生成器使用“附录A  标题”的 fallback，并在检查器中验证编号。

## 下一步优先级

- 优先完善检查器，而不是先写论文正文。
- 将 `STYLE_MAP.json` 拆成 `hard_rules`、`observed_styles`、`fallback_styles` 三层。
- 加入渲染检查脚本：DOCX -> PDF/PNG -> 人工目检。
