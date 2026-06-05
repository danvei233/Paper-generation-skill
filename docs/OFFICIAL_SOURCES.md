# 官方来源下载记录

本文件记录 `reference/official/` 中保存的《化工进展》官方资料、页面快照、附件下载和未能下载的外部依赖。完整机器清单见 `reference/official/official_manifest.json`。

## 已保存的官网页面

已从《化工进展》官网作者中心/专家中心抓取 31 个页面 HTML 快照，位于 `reference/official/html/`。其中与论文生成和检查直接相关的页面包括：

- `news39_征稿简则.html`
- `news44_稿件修改基本要求.html`
- `news45_绘图体例.html`
- `news46_常见问题.html`
- `news47_标准关键词库.html`
- `news48_综述性文章写法.html`
- `news49_论文摘要写法.html`
- `news54_参考文献著录格式.html`
- `news55_论文模板.html`
- `news59_专家审稿须知.html`
- `news60_综述性文章评审意见表.html`
- `news61_研究性文章评审意见表.html`
- `news656_中图分类号查询.html`
- `news661_化工进展_出版伦理规范.html`

所有 HTML、PDF、DOCX/DOC 附件的抽取文本已写入 `reference/official/extracted_text/`。

## 已下载附件

- `reference/official/hgjz_reference_format_2024.docx`
  - 来源：https://hgjz.cip.com.cn/CN/news/news54.shtml
  - 附件：https://hgjz.cip.com.cn/attached/file/20240206/20240206112245_857.docx
  - 用途：参考文献著录格式、文献类型标识、作者著录、半角标点、中文文献英文对应行等规则。

- `reference/official/hgjz_template_2023.doc`
  - 来源：https://hgjz.cip.com.cn/CN/news/news55.shtml
  - 附件：https://hgjz.cip.com.cn/fileup/1000-6613/NEWS/20230825095921_NewsFile_68.doc
  - 已转换：`reference/official/converted/hgjz_template_2023.docx`
  - 用途：官网 2023 模板备份。当前页面版式仍以用户提供的 2026 版模板为准。

- `reference/official/hgjz_standard_keywords_2019.doc`
  - 来源：https://hgjz.cip.com.cn/CN/news/news47.shtml
  - 旧附件：`http://www.hgjz.com.cn/UserFiles/File/《化工进展》标准关键词库（2019）.doc`
  - 状态：官网旧域名附件自动下载返回 HTTP 403；用户已从浏览器下载并提供本地文件。
  - 已抽取：`reference/official/hgjz_standard_keywords_2019.json`、`reference/official/extracted_text/hgjz_standard_keywords_2019.txt`
  - 用途：中文关键词至少 3 个来自标准关键词库的机器检查。

- `reference/official/news50_论文著作权授权声明书.pdf`
  - 来源：https://hgjz.cip.com.cn/CN/news/news50.shtml
  - 用途：投稿后的版权授权外部材料。

- `reference/official/news710_校样文件修改情况表.docx`
  - 来源：https://hgjz.cip.com.cn/CN/news/news710.shtml
  - 用途：校样阶段外部材料。

- `reference/official/news734_化工进展_内容授权许可协议.pdf`
  - 来源：https://hgjz.cip.com.cn/CN/news/news734.shtml
  - 用途：内容复用授权外部材料。

- `reference/official/news60_综述性文章评审意见表.doc`
  - 已转换：`reference/official/converted/news60_综述性文章评审意见表.docx`
  - 用途：综述文章质量与语言风格评审项。

- `reference/official/news61_研究性文章评审意见表.doc`
  - 已转换：`reference/official/converted/news61_研究性文章评审意见表.docx`
  - 用途：研究性文章质量与语言风格评审项。

## 附件未能下载

以下页面可访问，但附件位于旧域名 `www.hgjz.com.cn/UserFiles/File/`，多种 URL 编码、User-Agent 和 Referer 尝试后仍返回 HTTP 403：

- 论文摘要写法
  - 页面：https://hgjz.cip.com.cn/CN/news/news49.shtml
  - 旧附件：`http://www.hgjz.com.cn/UserFiles/File/《化工进展》论文摘要写法（2019）.doc`
  - 已采用替代规则来源：征稿简则、稿件修改基本要求、评审意见表、2026 模板。

- 综述性文章写法
  - 页面：https://hgjz.cip.com.cn/CN/news/news48.shtml
  - 旧附件：`http://www.hgjz.com.cn/UserFiles/File/《化工进展》综述性文章写法（2019）.doc`
  - 已采用替代规则来源：稿件修改基本要求、综述性文章评审意见表、征稿简则。

## 工具链默认依据

- 默认期刊格式：《化工进展》。
- 页面版式：优先 `reference/《化工进展》论文模板格式（2026版）.docx`。
- 引用著录：优先 `reference/official/hgjz_reference_format_2024.docx`。
- 语言风格：`docs/LANGUAGE_STYLE_RULES.md`。
- 全量规则矩阵：`docs/HGJZ_REQUIREMENTS_MATRIX.md`。
