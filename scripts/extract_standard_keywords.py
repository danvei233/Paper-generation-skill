#!/usr/bin/env python3
"""Clean the extracted text of the 2019 HGJZ standard keyword library.

The official attachment is an old .doc file. On machines where Word conversion
hangs or is unavailable, the repository keeps a raw text dump and this script
turns it into a conservative JSON/TXT keyword list for the format checker.
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


NOISE_TOKENS = {
    "HYPERLINK",
    "javascript",
    "ViewClassDetails",
    "Root Entry",
    "WordDocument",
    "SummaryInformation",
    "DocumentSummaryInformation",
    "CompObj",
    "橢",
    "坴",
    "阂",
    "鐇",
    "鐂",
    "烿",
    "氃",
    "瑌",
    "朥",
    "誖",
    "畾",
    "纊",
    "奝",
    "畝",
    "汏",
    "摫",
    "尀",
    "繟",
    "怀愁",
    "樀",
    "嘀",
    "猀",
    "挀",
    "瀀",
    "氀",
    "攀",
    "愀",
    "椀",
    "欀",
    "搀",
    "漀",
    "渀",
    "栀",
}

MANUAL_RECOVERIES = {
    "benzene": "苯",
    "pump": "泵",
    "enzyme": "酶",
    "film": "膜",
    "membranes": "膜",
}

MERGE_SUFFIXES = {"作用", "仪表"}
MERGE_PAIRS = {("共沸", "混合")}


def clean_line(line: str) -> str:
    return line.replace("\ufeff", "").replace("\x0c", "").strip()


def is_noise(line: str) -> bool:
    if not line:
        return True
    return any(token in line for token in NOISE_TOKENS)


def is_translation(line: str) -> bool:
    if not line or re.search(r"[\u4e00-\u9fff]", line):
        return False
    return bool(re.search(r"[A-Za-z]", line))


def is_keyword_candidate(line: str) -> bool:
    if is_noise(line):
        return False
    if not re.search(r"[\u4e00-\u9fff]", line):
        return False
    if re.search(r"[A-Za-z0-9]", line):
        return False
    if not re.fullmatch(r"[\u4e00-\u9fff·（）()]+", line):
        return False
    return 1 <= len(line) <= 12


def list_window(lines: list[str]) -> list[str]:
    start = next((i for i, line in enumerate(lines) if clean_line(line) == "安全"), None)
    if start is None:
        raise ValueError("Could not find keyword-list start marker: 安全")

    end = None
    for i in range(start, len(lines)):
        if clean_line(lines[i]) == "tissue engineering":
            end = i + 1
            break
    if end is None:
        end = next((i for i, line in enumerate(lines[start:], start) if "Root Entry" in line), len(lines))
    return lines[start:end]


def append_unique(words: list[str], word: str) -> None:
    if word and word not in words:
        words.append(word)


def flush_pending(pending_zh: list[str], translation: str, keywords: list[str], compounds: list[dict[str, object]]) -> None:
    if not pending_zh:
        return
    if len(pending_zh) == 1:
        append_unique(keywords, pending_zh[0])
        return

    should_merge = pending_zh[-1] in MERGE_SUFFIXES or tuple(pending_zh) in MERGE_PAIRS
    if should_merge:
        compound = "".join(pending_zh)
        append_unique(keywords, compound)
        compounds.append({"keyword": compound, "parts": list(pending_zh), "translation": translation})
        return

    for word in pending_zh:
        append_unique(keywords, word)


def extract_keywords(lines: list[str]) -> dict[str, object]:
    window = list_window(lines)
    keywords: list[str] = []
    compounds: list[dict[str, object]] = []
    recoveries: list[dict[str, str]] = []
    pending_zh: list[str] = []
    english_seen: set[str] = set()

    for raw in window:
        line = clean_line(raw)
        if not line or is_noise(line):
            continue
        if is_translation(line):
            lower = line.strip().lower()
            english_seen.add(lower)
            for english, zh in MANUAL_RECOVERIES.items():
                if lower == english:
                    append_unique(keywords, zh)
                    recoveries.append(
                        {
                            "keyword": zh,
                            "english": line.strip(),
                            "reason": "English keyword was present but the Chinese line was not recoverable from the old .doc text dump.",
                        }
                    )
            flush_pending(pending_zh, line.strip(), keywords, compounds)
            pending_zh = []
            continue
        if is_keyword_candidate(line):
            pending_zh.append(line)

    flush_pending(pending_zh, "", keywords, compounds)

    return {
        "source": "reference/official/hgjz_standard_keywords_2019.doc",
        "raw_text_source": "reference/official/extracted_text/hgjz_standard_keywords_2019_raw.txt",
        "status": "cleaned_from_user_provided_official_attachment",
        "count": len(keywords),
        "keywords": keywords,
        "manual_recoveries": recoveries,
        "compound_recoveries": compounds,
        "notes": [
            "The old .doc file contains hyperlink and OLE noise in raw extraction.",
            "The list is bounded from the first keyword 安全 through the final entry 组织工程学/tissue engineering.",
            "Manual recoveries are included only when the English entry is visible but the adjacent Chinese entry is lost in raw extraction.",
        ],
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("raw_text", type=Path)
    parser.add_argument("--json-out", type=Path, required=True)
    parser.add_argument("--txt-out", type=Path, required=True)
    args = parser.parse_args()

    lines = args.raw_text.read_text(encoding="utf-8", errors="ignore").splitlines()
    data = extract_keywords(lines)

    args.json_out.parent.mkdir(parents=True, exist_ok=True)
    args.txt_out.parent.mkdir(parents=True, exist_ok=True)
    args.json_out.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.txt_out.write_text("\n".join(data["keywords"]) + "\n", encoding="utf-8")
    print(f"Wrote {args.json_out} and {args.txt_out} with {data['count']} keywords")


if __name__ == "__main__":
    main()
