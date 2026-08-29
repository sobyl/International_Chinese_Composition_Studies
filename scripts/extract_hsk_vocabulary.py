#!/usr/bin/env python3
"""Extract the vocabulary table from the 2025 HSK syllabus PDF."""

from __future__ import annotations

import argparse
import csv
import re
from collections import Counter
from pathlib import Path

from pypdf import PdfReader


DEFAULT_INPUT = Path("outputs/新版HSK考试大纲1219.pdf")
DEFAULT_OUTPUT = Path("outputs/新版HSK词汇大纲.csv")
ROW_PATTERN = re.compile(r"^\s*\d+\s{2,}")
LEVEL_PATTERN = re.compile(r"^([^（]+)((?:（[^）]+）)*)$")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract the machine-readable vocabulary table from the HSK syllabus."
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--expected-count", type=int, default=11000)
    return parser.parse_args()


def find_section_page(reader: PdfReader, heading: str) -> int:
    for page_index, page in enumerate(reader.pages):
        lines = [line.replace(" ", "").strip() for line in (page.extract_text() or "").splitlines()]
        if heading in lines:
            return page_index
    raise ValueError(f"PDF 中未找到章节：{heading}")


def parse_level(raw_level: str) -> tuple[str, str, str]:
    match = LEVEL_PATTERN.fullmatch(raw_level)
    if not match:
        raise ValueError(f"无法解析等级：{raw_level}")

    main_level = match.group(1)
    additional_levels = "|".join(re.findall(r"（([^）]+)）", match.group(2)))
    if main_level in {"1", "2", "3"}:
        level_group = "初等"
    elif main_level in {"4", "5", "6"}:
        level_group = "中等"
    elif main_level == "7-9":
        level_group = "高等"
    else:
        level_group = "其他"
    return main_level, additional_levels, level_group


def extract_rows(reader: PdfReader) -> list[dict[str, str | int]]:
    vocabulary_title_page = find_section_page(reader, "词汇大纲")
    character_title_page = find_section_page(reader, "汉字大纲")
    rows: list[dict[str, str | int]] = []

    # The title page is followed by the table; the next section title ends it.
    for page_index in range(vocabulary_title_page + 1, character_title_page):
        text = reader.pages[page_index].extract_text(extraction_mode="layout") or ""
        for line in text.splitlines():
            if not ROW_PATTERN.match(line):
                continue

            parts = re.split(r"\s{2,}", line.strip())
            if len(parts) not in {4, 5} or not parts[0].isdigit():
                raise ValueError(f"PDF 第 {page_index + 1} 页存在无法解析的行：{line!r}")
            if len(parts) == 4:
                parts.append("")

            serial, raw_level, word, pinyin, part_of_speech = parts
            main_level, additional_levels, level_group = parse_level(raw_level)
            rows.append(
                {
                    "序号": int(serial),
                    "等级原文": raw_level,
                    "主等级": main_level,
                    "兼属等级": additional_levels,
                    "等级组": level_group,
                    "词语": word,
                    "拼音": pinyin,
                    "词性": part_of_speech,
                    "PDF页码": page_index + 1,
                }
            )
    return rows


def validate_rows(rows: list[dict[str, str | int]], expected_count: int) -> None:
    serials = [int(row["序号"]) for row in rows]
    expected_serials = list(range(1, expected_count + 1))
    if len(rows) != expected_count:
        raise ValueError(f"词条数量异常：期望 {expected_count}，实际 {len(rows)}")
    if serials != expected_serials:
        counts = Counter(serials)
        missing = [number for number in expected_serials if number not in counts]
        duplicates = [number for number, count in counts.items() if count > 1]
        raise ValueError(f"序号不连续；缺失={missing[:20]}，重复={duplicates[:20]}")


def write_csv(rows: list[dict[str, str | int]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "序号",
        "等级原文",
        "主等级",
        "兼属等级",
        "等级组",
        "词语",
        "拼音",
        "词性",
        "PDF页码",
    ]
    temporary_path = output_path.with_suffix(output_path.suffix + ".tmp")
    with temporary_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    temporary_path.replace(output_path)


def main() -> None:
    args = parse_args()
    if not args.input.is_file():
        raise FileNotFoundError(f"输入 PDF 不存在：{args.input}")

    reader = PdfReader(args.input)
    if reader.is_encrypted and reader.decrypt("") == 0:
        raise ValueError("PDF 已加密，无法使用空密码读取")

    rows = extract_rows(reader)
    validate_rows(rows, args.expected_count)
    write_csv(rows, args.output)

    group_counts = Counter(str(row["等级组"]) for row in rows)
    print(f"已生成：{args.output}")
    print(f"词条总数：{len(rows)}")
    print(
        "等级组："
        + "，".join(f"{name} {group_counts[name]}" for name in ("初等", "中等", "高等"))
    )


if __name__ == "__main__":
    main()
