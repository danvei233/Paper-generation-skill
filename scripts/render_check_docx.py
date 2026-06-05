#!/usr/bin/env python3
"""Render a DOCX with Microsoft Word and inspect the produced PDF.

This is a higher-level QA pass than OOXML checks. It uses Word's layout engine
to export PDF, renders PDF pages to PNG with PyMuPDF, and reports page size,
blank-page risk, text extraction, and content bounds.
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path
from typing import Any

from PIL import Image

try:
    import fitz
except Exception as exc:  # pragma: no cover
    raise SystemExit("PyMuPDF is required: pip install PyMuPDF") from exc


PT_TO_CM = 2.54 / 72.0


def load_style(path: Path) -> dict[str, Any]:
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return {}


def ps_quote(path: Path) -> str:
    return str(path).replace("'", "''")


def export_pdf_with_word(docx: Path, pdf: Path) -> str:
    docx = docx.resolve()
    pdf = pdf.resolve()
    pdf.parent.mkdir(parents=True, exist_ok=True)
    script = f"""
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {{
  $doc = $word.Documents.Open('{ps_quote(docx)}', $false, $true)
  $doc.ExportAsFixedFormat('{ps_quote(pdf)}', 17)
  $doc.Close($false)
}} finally {{
  $word.Quit()
}}
Write-Output 'exported'
"""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=180,
    )
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout).strip())
    return result.stdout.strip()


def image_nonwhite_ratio(path: Path) -> float:
    with Image.open(path).convert("RGB") as image:
        pixels = image.getdata()
        total = image.width * image.height
        nonwhite = 0
        for r, g, b in pixels:
            if r < 245 or g < 245 or b < 245:
                nonwhite += 1
        return nonwhite / total if total else 0.0


def render_pdf_pages(pdf: Path, out_dir: Path, dpi: int) -> list[dict[str, Any]]:
    out_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf)
    pages = []
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    for index, page in enumerate(doc, start=1):
        png = out_dir / f"page_{index:03d}.png"
        pixmap = page.get_pixmap(matrix=matrix, alpha=False)
        pixmap.save(png)
        text = page.get_text("text").strip()
        blocks = [b for b in page.get_text("blocks") if len(b) >= 5 and str(b[4]).strip()]
        if blocks:
            x0 = min(b[0] for b in blocks)
            y0 = min(b[1] for b in blocks)
            x1 = max(b[2] for b in blocks)
            y1 = max(b[3] for b in blocks)
            bbox_cm = [round(v * PT_TO_CM, 2) for v in (x0, y0, x1, y1)]
        else:
            bbox_cm = None
        pages.append(
            {
                "number": index,
                "png": str(png),
                "width_cm": round(page.rect.width * PT_TO_CM, 2),
                "height_cm": round(page.rect.height * PT_TO_CM, 2),
                "text_chars": len(text),
                "text": text,
                "text_preview": text[:120].replace("\n", " "),
                "nonwhite_ratio": round(image_nonwhite_ratio(png), 4),
                "content_bbox_cm": bbox_cm,
            }
        )
    doc.close()
    return pages


def near(actual: float, expected: float, tol: float = 0.12) -> bool:
    return abs(actual - expected) <= tol


def add(items: list[dict[str, str]], status: str, message: str, location: str = "", fix: str = "") -> None:
    items.append({"status": status, "message": message, "location": location, "fix": fix})


def inspect_pages(pages: list[dict[str, Any]], style: dict[str, Any]) -> list[dict[str, str]]:
    items: list[dict[str, str]] = []
    page_setup = style.get("page_setup", {})
    width = float(page_setup.get("width_cm", 21.0))
    height = float(page_setup.get("height_cm", 29.7))
    left = float(page_setup.get("left_margin_cm", 2.0))
    right = float(page_setup.get("right_margin_cm", 2.0))
    top = float(page_setup.get("top_margin_cm", 2.5))
    bottom = float(page_setup.get("bottom_margin_cm", 2.0))

    if not pages:
        add(items, "ERROR", "PDF 没有页面", fix="检查 Word 导出是否失败")
        return items

    for page in pages:
        loc = f"page {page['number']}"
        if not near(page["width_cm"], width) or not near(page["height_cm"], height):
            add(items, "ERROR", f"页面尺寸不是 A4：{page['width_cm']} cm × {page['height_cm']} cm", loc, "检查 Word 页面设置和导出设置")
        if page["nonwhite_ratio"] < 0.003:
            add(items, "ERROR", f"疑似空白页：非白像素比例 {page['nonwhite_ratio']}", loc, "检查是否误插入空页或分页符")
        if page["text_chars"] == 0:
            add(items, "WARN", "页面没有可抽取文本，可能是整页图片或导出异常", loc, "正文应保持可编辑文本")
        bbox = page["content_bbox_cm"]
        if bbox:
            x0, y0, x1, y1 = bbox
            if x0 < left - 0.45 or x1 > width - right + 0.45:
                add(items, "WARN", f"内容横向边界接近/越过版心：bbox={bbox}", loc, "检查表格、图片或长英文是否越界")
            if y0 < top - 0.75 or y1 > height - bottom + 0.75:
                add(items, "WARN", f"内容纵向边界接近/越过版心：bbox={bbox}", loc, "检查页眉页脚、脚注或页面溢出")

    joined_text = "\n".join(page.get("text", "") for page in pages)
    joined_compact = "".join(joined_text.split())
    for marker in ["气泡泵压降模型的分析与优化", "图1", "表1", "参考文献"]:
        if marker not in joined_text and marker not in joined_compact:
            add(items, "WARN", f"PDF 文本预览未检测到关键标记：{marker}", fix="打开 PDF 确认文本导出是否完整")

    if not any(item["status"] in {"ERROR", "WARN"} for item in items):
        add(items, "PASS", "Word 导出 PDF、页面尺寸、空白页、文本抽取和内容边界未发现问题")
    return items


def write_report(docx: Path, pdf: Path, pages: list[dict[str, Any]], items: list[dict[str, str]], out: Path) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    errors = sum(1 for item in items if item["status"] == "ERROR")
    warnings = sum(1 for item in items if item["status"] == "WARN")
    lines = [
        "# DOCX 渲染复核报告",
        "",
        f"源文件：{docx}",
        f"PDF：{pdf}",
        f"ERROR: {errors}，WARN: {warnings}",
        "",
        "## 复核结果",
    ]
    for item in items:
        lines.append(f"- {item['status']}：{item['message']}")
        if item.get("location"):
            lines.append(f"  - 位置：{item['location']}")
        if item.get("fix"):
            lines.append(f"  - 建议：{item['fix']}")
    lines.extend(["", "## 页面渲染"])
    for page in pages:
        lines.append(
            f"- page {page['number']}: {page['width_cm']} cm × {page['height_cm']} cm, "
            f"text={page['text_chars']} chars, nonwhite={page['nonwhite_ratio']}, "
            f"bbox={page['content_bbox_cm']}, png={page['png']}"
        )
    lines.extend(
        [
            "",
            "## 边界",
            "- 该复核使用 Microsoft Word 16.0 导出 PDF，并用 PyMuPDF 渲染 PNG。",
            "- 它能发现页面尺寸、空白页、可抽取文本、明显越界等渲染层问题；图片内部坐标轴线宽、字体和真实图像质量仍需逐图检查。",
        ]
    )
    out.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--style", type=Path, default=Path("STYLE_MAP.json"))
    parser.add_argument("--pdf", type=Path, default=Path("output/rendered/manuscript.pdf"))
    parser.add_argument("--png-dir", type=Path, default=Path("output/rendered/pages"))
    parser.add_argument("--out", type=Path, default=Path("output/render_report.md"))
    parser.add_argument("--dpi", type=int, default=180)
    args = parser.parse_args()

    style = load_style(args.style)
    export_pdf_with_word(args.docx, args.pdf)
    pages = render_pdf_pages(args.pdf, args.png_dir, args.dpi)
    items = inspect_pages(pages, style)
    write_report(args.docx, args.pdf, pages, items, args.out)
    errors = sum(1 for item in items if item["status"] == "ERROR")
    warnings = sum(1 for item in items if item["status"] == "WARN")
    print(f"Wrote {args.out} with {errors} errors and {warnings} warnings")
    if errors:
        sys.exit(1)


if __name__ == "__main__":
    main()
