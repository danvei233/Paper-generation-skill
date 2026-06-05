# DOCX 渲染复核报告

源文件：output\manuscript.docx
PDF：output\rendered\manuscript.pdf
ERROR: 0，WARN: 0

## 复核结果
- PASS：Word 导出 PDF、页面尺寸、空白页、文本抽取和内容边界未发现问题

## 页面渲染
- page 1: 21.0 cm × 29.7 cm, text=2147 chars, nonwhite=0.0609, bbox=[2.0, 2.79, 19.05, 26.99], png=output\rendered\pages\page_001.png
- page 2: 21.0 cm × 29.7 cm, text=776 chars, nonwhite=0.0309, bbox=[2.0, 7.14, 19.04, 20.82], png=output\rendered\pages\page_002.png

## 边界
- 该复核使用 Microsoft Word 16.0 导出 PDF，并用 PyMuPDF 渲染 PNG。
- 它能发现页面尺寸、空白页、可抽取文本、明显越界等渲染层问题；图片内部坐标轴线宽、字体和真实图像质量仍需逐图检查。