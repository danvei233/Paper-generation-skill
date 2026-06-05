#!/usr/bin/env python3
"""Create a submission-safe DOCX copy without hyperlinks or Word field codes.

The script works directly on OOXML so it can be used before the final format
check. It preserves displayed field results for ordinary complex fields and
fldSimple fields, unwraps hyperlinks into normal runs, and removes comment,
footnote and endnote references from the submission copy.
"""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path
from typing import Any
from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

W_NS = NS["w"]
REL_NS = NS["rel"]
CT_NS = NS["ct"]

DROP_PARTS = {
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
}
DROP_REL_TYPES = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
}


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def part_should_transform(name: str) -> bool:
    return (
        name == "word/document.xml"
        or name.startswith("word/header")
        or name.startswith("word/footer")
        or name in {"word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"}
    )


def parse_xml(data: bytes) -> etree._Element:
    return etree.fromstring(data)


def serialize(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def remove_element(element: etree._Element) -> None:
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def unwrap(element: etree._Element) -> None:
    parent = element.getparent()
    if parent is None:
        return
    index = parent.index(element)
    for child in list(element):
        parent.insert(index, child)
        index += 1
    parent.remove(element)


def cleanup_empty_runs(root: etree._Element) -> None:
    for run in list(root.xpath(".//w:r", namespaces=NS)):
        if run.xpath(".//w:t|.//w:drawing|.//w:pict|.//m:oMath|.//m:oMathPara", namespaces={**NS, "m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}):
            continue
        remove_element(run)


def unwrap_hyperlinks(root: etree._Element) -> int:
    count = 0
    for hyperlink in list(root.xpath(".//w:hyperlink", namespaces=NS)):
        unwrap(hyperlink)
        count += 1
    return count


def unwrap_simple_fields(root: etree._Element) -> int:
    count = 0
    for field in list(root.xpath(".//w:fldSimple", namespaces=NS)):
        unwrap(field)
        count += 1
    return count


def flatten_complex_fields(root: etree._Element) -> int:
    flattened = 0
    for paragraph in root.xpath(".//w:p", namespaces=NS):
        depth = 0
        in_result = False
        for child in list(paragraph):
            if child.tag != w("r"):
                continue
            field_chars = child.xpath(".//w:fldChar", namespaces=NS)
            has_instr = bool(child.xpath(".//w:instrText", namespaces=NS))
            remove_run = False

            if field_chars:
                remove_run = True
                flattened += 1
                for field_char in field_chars:
                    field_type = field_char.get(w("fldCharType"))
                    if field_type == "begin":
                        depth += 1
                        in_result = False
                    elif field_type == "separate" and depth > 0:
                        in_result = True
                    elif field_type == "end" and depth > 0:
                        depth -= 1
                        if depth == 0:
                            in_result = False
            elif has_instr:
                remove_run = True
            elif depth > 0 and not in_result:
                remove_run = True

            if remove_run:
                remove_element(child)
    return flattened


def remove_note_comment_refs(root: etree._Element) -> int:
    count = 0
    tags = [
        "commentRangeStart",
        "commentRangeEnd",
        "commentReference",
        "footnoteReference",
        "endnoteReference",
    ]
    for tag in tags:
        for element in list(root.xpath(f".//w:{tag}", namespaces=NS)):
            remove_element(element)
            count += 1
    return count


def transform_word_xml(data: bytes) -> tuple[bytes, dict[str, int]]:
    root = parse_xml(data)
    stats = {
        "hyperlinks_unwrapped": unwrap_hyperlinks(root),
        "simple_fields_unwrapped": unwrap_simple_fields(root),
        "complex_field_runs_removed": flatten_complex_fields(root),
        "note_comment_refs_removed": remove_note_comment_refs(root),
    }
    cleanup_empty_runs(root)
    return serialize(root), stats


def transform_rels(data: bytes) -> tuple[bytes, int]:
    root = parse_xml(data)
    removed = 0
    for rel in list(root.xpath(".//rel:Relationship", namespaces=NS)):
        rel_type = rel.get("Type")
        if rel_type in DROP_REL_TYPES:
            remove_element(rel)
            removed += 1
    return serialize(root), removed


def transform_content_types(data: bytes) -> tuple[bytes, int]:
    root = parse_xml(data)
    removed = 0
    for override in list(root.xpath(".//ct:Override", namespaces=NS)):
        part_name = override.get("PartName", "").lstrip("/")
        if part_name in DROP_PARTS:
            remove_element(override)
            removed += 1
    return serialize(root), removed


def real_note_ids(xml: str, tag: str) -> list[int]:
    ids: list[int] = []
    for match in re.finditer(rf"<w:{tag}\b[^>]*w:id=\"(-?\d+)\"", xml):
        note_id = int(match.group(1))
        if note_id > 0:
            ids.append(note_id)
    return ids


def scan(docx: Path) -> dict[str, Any]:
    with zipfile.ZipFile(docx) as package:
        names = set(package.namelist())
        xml_parts = {
            name: package.read(name).decode("utf-8", errors="ignore")
            for name in names
            if name.endswith(".xml") or name.endswith(".rels")
        }
    document_xml = xml_parts.get("word/document.xml", "")
    footnote_ids = real_note_ids(xml_parts.get("word/footnotes.xml", ""), "footnote")
    endnote_ids = real_note_ids(xml_parts.get("word/endnotes.xml", ""), "endnote")
    return {
        "has_hyperlinks": any("<w:hyperlink" in xml or "relationships/hyperlink" in xml for xml in xml_parts.values()),
        "has_field_codes": any(token in document_xml for token in ("w:fldChar", "w:instrText", "<w:fldSimple")),
        "has_comments": "word/comments.xml" in names or "w:commentReference" in document_xml,
        "has_footnotes": bool(footnote_ids) or "w:footnoteReference" in document_xml,
        "has_endnotes": bool(endnote_ids) or "w:endnoteReference" in document_xml,
        "footnote_ids": footnote_ids,
        "endnote_ids": endnote_ids,
    }


def flatten_docx(source: Path, out: Path) -> dict[str, Any]:
    out.parent.mkdir(parents=True, exist_ok=True)
    stats: dict[str, Any] = {
        "word_xml_parts_transformed": 0,
        "relationships_removed": 0,
        "content_type_overrides_removed": 0,
        "parts_dropped": [],
        "xml_changes": {
            "hyperlinks_unwrapped": 0,
            "simple_fields_unwrapped": 0,
            "complex_field_runs_removed": 0,
            "note_comment_refs_removed": 0,
        },
    }
    with zipfile.ZipFile(source) as package, zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as new_docx:
        for info in package.infolist():
            name = info.filename
            if name in DROP_PARTS:
                stats["parts_dropped"].append(name)
                continue

            data = package.read(name)
            if part_should_transform(name):
                data, xml_stats = transform_word_xml(data)
                stats["word_xml_parts_transformed"] += 1
                for key, value in xml_stats.items():
                    stats["xml_changes"][key] += value
            elif name.endswith(".rels"):
                data, removed = transform_rels(data)
                stats["relationships_removed"] += removed
            elif name == "[Content_Types].xml":
                data, removed = transform_content_types(data)
                stats["content_type_overrides_removed"] += removed
            new_docx.writestr(info, data)
    return stats


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path, help="write a flattened submission DOCX copy")
    parser.add_argument("--scan-only", action="store_true", help="only print risk scan")
    args = parser.parse_args()

    before = scan(args.docx)
    result: dict[str, Any] = {"source": str(args.docx), "before": before}
    if args.out and not args.scan_only:
        result["flatten"] = flatten_docx(args.docx, args.out)
        result["out"] = str(args.out)
        result["after"] = scan(args.out)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
