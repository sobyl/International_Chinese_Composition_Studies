#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import platform
import re
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any


DEFAULT_WORKBOOK = "作文样本主表.xlsx"
DEFAULT_INPUT_DIR = "clean_text"
DEFAULT_SEG_OUTPUT_DIR = "seg_text"
DEFAULT_STATS_OUTPUT = "outputs/作文词性统计宽表.xlsx"
DEFAULT_HSK_VOCAB = "outputs/新版HSK词汇大纲.csv"
SHEET_NAMES = ("J1", "J2", "Y1", "Y2")
REQUIRED_COLUMNS = ("篇名代码", "篇名", "作文编码", "国籍", "作文题目", "作文分数", "体裁", "作文文件名")
EXPECTED_TOTAL = 620
EXPECTED_HSK_VOCAB_COUNT = 11000
PUNCT_POS = "punctuation mark"
HAN_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
HSK_POS_RE = re.compile(r"前缀|后缀|拟声|数量|名|动|形|副|代|介|连|助|数|量|叹")

HSK_LEVELS = ("1", "2", "3", "4", "5", "6", "7-9")
HSK_LEVEL_RANK = {level: index for index, level in enumerate(HSK_LEVELS)}
HSK_LEVEL_GROUPS = {
    "初等": ("1", "2", "3"),
    "中等": ("4", "5", "6"),
    "高等": ("7-9",),
}
HSK_REQUIRED_COLUMNS = ("词语", "主等级", "词性")

PY_NLPIR_TO_HSK_POS = {
    "noun": frozenset({"名"}),
    "pronoun": frozenset({"代"}),
    "verb": frozenset({"动"}),
    "adjective": frozenset({"形"}),
    "adverb": frozenset({"副"}),
    "preposition": frozenset({"介"}),
    "conjunction": frozenset({"连"}),
    "particle": frozenset({"助"}),
    "numeral": frozenset({"数", "数量"}),
    "classifier": frozenset({"量", "数量"}),
    "time word": frozenset({"名"}),
    "noun of locality": frozenset({"名"}),
    "locative word": frozenset({"名"}),
    "suffix": frozenset({"后缀"}),
    "prefix": frozenset({"前缀"}),
    "modal particle": frozenset({"助"}),
    "distinguishing word": frozenset({"形"}),
    "status word": frozenset({"形"}),
    "onomatopoeia": frozenset({"拟声"}),
    "interjection": frozenset({"叹"}),
}

POS_COLUMN_MAP = {
    "noun": "名词数",
    "pronoun": "代词数",
    "verb": "动词数",
    "adjective": "形容词数",
    "adverb": "副词数",
    "preposition": "介词数",
    "conjunction": "连词数",
    "particle": "助词数",
    "numeral": "数词数",
    "classifier": "量词数",
    "time word": "时间词数",
    "noun of locality": "方位词数",
    "locative word": "方位词数",
    "suffix": "后缀数",
    "modal particle": "语气词数",
    "distinguishing word": "区别词数",
    "status word": "状态词数",
    "onomatopoeia": "拟声词数",
    "interjection": "叹词数",
    "multiword expression": "熟语数",
    PUNCT_POS: "标点数",
}

POS_COLUMN_ORDER = [
    "名词数",
    "代词数",
    "动词数",
    "形容词数",
    "副词数",
    "介词数",
    "连词数",
    "助词数",
    "数词数",
    "量词数",
    "时间词数",
    "方位词数",
    "后缀数",
    "语气词数",
    "区别词数",
    "状态词数",
    "拟声词数",
    "叹词数",
    "熟语数",
    "标点数",
]

BASIC_HEADERS = ["篇名代码", "篇名", "作文编码", "国籍", "作文题目", "作文分数", "体裁", "作文文件名"]
TEXT_STAT_HEADERS = ["字数", "纯文本字数", "分词数", "非标点分词数", "去重词数"]
HSK_STAT_HEADERS = [
    header
    for level in HSK_LEVELS
    for header in (f"{level}级词汇次数", f"{level}级词汇占比")
] + [
    header
    for group in HSK_LEVEL_GROUPS
    for header in (f"{group}词汇次数", f"{group}词汇占比")
]


@dataclass(frozen=True)
class EssayRecord:
    sheet_name: str
    code: str
    title_code_name: str
    essay_id: str
    nationality: str
    essay_topic: str
    score: int | str
    genre: str
    filename: str


@dataclass(frozen=True)
class HskVocabularyEntry:
    level: str
    pos_categories: frozenset[str]


@dataclass
class EssayStats:
    record: EssayRecord
    char_count: int
    han_char_count: int
    token_count: int
    non_punct_token_count: int
    unique_word_count: int
    pos_counts: Counter[str]
    hsk_level_counts: Counter[str]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "使用 PyNLPIR 对 clean_text 作文分词，并生成词性统计宽表。"
            "请用：arch -x86_64 /usr/bin/python3 segment_hsk_clean_texts.py"
        ),
    )
    parser.add_argument("--workbook", default=DEFAULT_WORKBOOK, help=f"作文主表，默认：{DEFAULT_WORKBOOK}")
    parser.add_argument("--input-dir", default=DEFAULT_INPUT_DIR, help=f"清洗文本目录，默认：{DEFAULT_INPUT_DIR}")
    parser.add_argument("--seg-output-dir", default=DEFAULT_SEG_OUTPUT_DIR, help=f"分词输出目录，默认：{DEFAULT_SEG_OUTPUT_DIR}")
    parser.add_argument("--stats-output", default=DEFAULT_STATS_OUTPUT, help=f"统计宽表输出，默认：{DEFAULT_STATS_OUTPUT}")
    parser.add_argument("--hsk-vocab", default=DEFAULT_HSK_VOCAB, help=f"HSK 词汇表 CSV，默认：{DEFAULT_HSK_VOCAB}")
    parser.add_argument("--limit", type=int, default=None, help="只处理前 N 篇，用于小样本测试")
    parser.add_argument("--dry-run", action="store_true", help="只校验主表和文本文件，不写输出")
    return parser.parse_args()


def require_x86_64() -> None:
    if platform.machine() != "x86_64":
        raise RuntimeError(
            "PyNLPIR 当前只能在 x86_64 Python 下运行。请使用：\n"
            "  arch -x86_64 /usr/bin/python3 segment_hsk_clean_texts.py"
        )


def validate_args(args: argparse.Namespace) -> None:
    if args.limit is not None and args.limit <= 0:
        raise ValueError("--limit 必须大于 0")
    if not str(args.stats_output).lower().endswith(".xlsx"):
        raise ValueError("--stats-output 必须是 .xlsx 文件")


def import_runtime_dependencies():
    try:
        import pynlpir  # type: ignore[import-not-found]
        from openpyxl import Workbook, load_workbook  # type: ignore[import-not-found]
        from openpyxl.styles import Alignment, Font, PatternFill  # type: ignore[import-not-found]
        from openpyxl.utils import get_column_letter  # type: ignore[import-not-found]
    except Exception as exc:  # pragma: no cover - user-facing environment guard
        raise RuntimeError(
            "缺少运行依赖。请确认已在 x86_64 Python 环境安装 pynlpir 和 openpyxl：\n"
            "  arch -x86_64 /usr/bin/python3 -m pip install --user pynlpir openpyxl"
        ) from exc
    return pynlpir, Workbook, load_workbook, Alignment, Font, PatternFill, get_column_letter


def normalize_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value).strip()
    return str(value).strip()


def normalize_score(value: Any) -> int | str:
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    text = normalize_cell(value)
    return int(text) if text.isdigit() else text


def read_hsk_vocabulary(path: Path) -> dict[str, tuple[HskVocabularyEntry, ...]]:
    if not path.is_file():
        raise FileNotFoundError(f"找不到 HSK 词汇表：{path}")

    entries_by_word: dict[str, list[HskVocabularyEntry]] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        headers = reader.fieldnames or []
        missing_columns = [column for column in HSK_REQUIRED_COLUMNS if column not in headers]
        if missing_columns:
            raise ValueError(f"HSK 词汇表缺少必要列：{missing_columns}；实际表头：{headers}")

        row_count = 0
        for row_number, row in enumerate(reader, start=2):
            row_count += 1
            word = normalize_cell(row.get("词语"))
            level = normalize_cell(row.get("主等级"))
            raw_pos = normalize_cell(row.get("词性"))
            if not word:
                raise ValueError(f"HSK 词汇表第 {row_number} 行缺少词语")
            if level not in HSK_LEVEL_RANK:
                raise ValueError(f"HSK 词汇表第 {row_number} 行主等级无效：{level!r}")

            entries_by_word.setdefault(word, []).append(
                HskVocabularyEntry(
                    level=level,
                    pos_categories=frozenset(HSK_POS_RE.findall(raw_pos.replace(" ", ""))),
                )
            )

    if row_count != EXPECTED_HSK_VOCAB_COUNT:
        raise ValueError(
            f"HSK 词汇表应有 {EXPECTED_HSK_VOCAB_COUNT} 条记录，实际读取 {row_count} 条"
        )
    return {word: tuple(entries) for word, entries in entries_by_word.items()}


def select_hsk_level(
    word: str,
    pos: str | None,
    vocabulary: dict[str, tuple[HskVocabularyEntry, ...]],
) -> str | None:
    entries = vocabulary.get(word)
    if not entries:
        return None

    levels = {entry.level for entry in entries}
    if len(levels) == 1:
        return next(iter(levels))

    target_pos = PY_NLPIR_TO_HSK_POS.get((pos or "").strip(), frozenset())
    matched_levels = {
        entry.level
        for entry in entries
        if target_pos and entry.pos_categories.intersection(target_pos)
    }
    if len(matched_levels) == 1:
        return next(iter(matched_levels))

    return min(levels, key=HSK_LEVEL_RANK.__getitem__)


def read_records(workbook_path: Path, limit: int | None, load_workbook: Any) -> list[EssayRecord]:
    if not workbook_path.exists():
        raise FileNotFoundError(f"找不到作文主表：{workbook_path}")

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    records: list[EssayRecord] = []
    try:
        missing_sheets = [sheet_name for sheet_name in SHEET_NAMES if sheet_name not in workbook.sheetnames]
        if missing_sheets:
            raise ValueError(f"作文主表缺少 sheet：{missing_sheets}")

        for sheet_name in SHEET_NAMES:
            sheet = workbook[sheet_name]
            headers = [normalize_cell(cell.value) for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
            missing_columns = [column for column in REQUIRED_COLUMNS if column not in headers]
            if missing_columns:
                raise ValueError(f"{sheet_name} 缺少必要列：{missing_columns}；实际表头：{headers}")

            index = {column: headers.index(column) for column in REQUIRED_COLUMNS}
            for row_number, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                if not row or all(value is None for value in row):
                    continue

                code = normalize_cell(row[index["篇名代码"]])
                filename = normalize_cell(row[index["作文文件名"]])
                if code != sheet_name:
                    raise ValueError(f"{sheet_name}!A{row_number} 篇名代码应为 {sheet_name}，实际为 {code!r}")
                if not filename:
                    raise ValueError(f"{sheet_name}!{row_number} 缺少作文文件名")
                if filename.endswith(".txt"):
                    filename = filename[:-4]

                records.append(
                    EssayRecord(
                        sheet_name=sheet_name,
                        code=code,
                        title_code_name=normalize_cell(row[index["篇名"]]),
                        essay_id=normalize_cell(row[index["作文编码"]]),
                        nationality=normalize_cell(row[index["国籍"]]),
                        essay_topic=normalize_cell(row[index["作文题目"]]),
                        score=normalize_score(row[index["作文分数"]]),
                        genre=normalize_cell(row[index["体裁"]]),
                        filename=filename,
                    )
                )
    finally:
        workbook.close()

    if len(records) != EXPECTED_TOTAL:
        raise ValueError(f"作文主表应有 {EXPECTED_TOTAL} 条记录，实际读取 {len(records)} 条")
    return records[:limit] if limit is not None else records


def source_path_for(input_dir: Path, record: EssayRecord) -> Path:
    return input_dir / record.code / f"{record.filename}.txt"


def seg_path_for(output_dir: Path, record: EssayRecord) -> Path:
    return output_dir / record.code / f"{record.filename}.txt"


def validate_sources(records: list[EssayRecord], input_dir: Path) -> None:
    missing = [source_path_for(input_dir, record) for record in records if not source_path_for(input_dir, record).exists()]
    if missing:
        preview = "\n".join(str(path) for path in missing[:20])
        raise FileNotFoundError(f"缺少 clean_text 源文件，共 {len(missing)} 个；前几个：\n{preview}")

    duplicate_files = [item for item, count in Counter((record.code, record.filename) for record in records).items() if count > 1]
    if duplicate_files:
        raise ValueError(f"作文文件名重复：{duplicate_files[:20]}")


def pos_column_name(pos: str | None) -> str:
    normalized = (pos or "").strip()
    if normalized in POS_COLUMN_MAP:
        return POS_COLUMN_MAP[normalized]
    safe_pos = re.sub(r"[^0-9A-Za-z]+", "_", normalized).strip("_") or "空"
    return f"其他词性_{safe_pos}数"


def pos_label(pos: str | None) -> str:
    column_name = pos_column_name(pos)
    return column_name[:-1] if column_name.endswith("数") else column_name


def count_chars(text: str) -> tuple[int, int]:
    no_space = re.sub(r"\s+", "", text)
    return len(no_space), len(HAN_RE.findall(no_space))


def segment_lines(
    text: str,
    pynlpir: Any,
    hsk_vocabulary: dict[str, tuple[HskVocabularyEntry, ...]],
) -> tuple[str, Counter[str], Counter[str], int, int]:
    output_lines: list[str] = []
    pos_counts: Counter[str] = Counter()
    hsk_level_counts: Counter[str] = Counter()
    unique_words: set[str] = set()
    token_count = 0

    for line in text.splitlines():
        if not line.strip():
            output_lines.append("")
            continue

        segmented = pynlpir.segment(line)
        token_count += len(segmented)
        tokens: list[str] = []
        for word, pos in segmented:
            word_text = str(word).replace("\n", " ").strip()
            label = pos_label(pos)
            pos_counts[pos_column_name(pos)] += 1
            if (pos or "").strip() != PUNCT_POS:
                if word_text:
                    unique_words.add(word_text)
                hsk_level = select_hsk_level(word_text, pos, hsk_vocabulary)
                if hsk_level is not None:
                    hsk_level_counts[hsk_level] += 1
            tokens.append(f"{word_text}/{label}")
        output_lines.append(" ".join(tokens))

    return (
        "\n".join(output_lines).rstrip() + "\n",
        pos_counts,
        hsk_level_counts,
        token_count,
        len(unique_words),
    )


def atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile("w", encoding="utf-8", dir=path.parent, prefix=f".{path.name}.", suffix=".tmp", delete=False) as tmp:
        tmp.write(text)
        tmp_path = Path(tmp.name)
    tmp_path.replace(path)


def segment_record(
    record: EssayRecord,
    input_dir: Path,
    seg_output_dir: Path,
    pynlpir: Any,
    hsk_vocabulary: dict[str, tuple[HskVocabularyEntry, ...]],
) -> EssayStats:
    source_path = source_path_for(input_dir, record)
    text = source_path.read_text(encoding="utf-8")
    segmented_text, pos_counts, hsk_level_counts, token_count, unique_word_count = segment_lines(
        text,
        pynlpir,
        hsk_vocabulary,
    )
    char_count, han_char_count = count_chars(text)
    non_punct_token_count = token_count - pos_counts.get(POS_COLUMN_MAP[PUNCT_POS], 0)
    atomic_write_text(seg_path_for(seg_output_dir, record), segmented_text)
    return EssayStats(
        record=record,
        char_count=char_count,
        han_char_count=han_char_count,
        token_count=token_count,
        non_punct_token_count=non_punct_token_count,
        unique_word_count=unique_word_count,
        pos_counts=pos_counts,
        hsk_level_counts=hsk_level_counts,
    )


def hsk_stat_values(item: EssayStats) -> list[int | float]:
    denominator = item.non_punct_token_count
    values: list[int | float] = []
    for level in HSK_LEVELS:
        count = item.hsk_level_counts.get(level, 0)
        values.extend((count, count / denominator if denominator else 0.0))
    for levels in HSK_LEVEL_GROUPS.values():
        count = sum(item.hsk_level_counts.get(level, 0) for level in levels)
        values.extend((count, count / denominator if denominator else 0.0))
    return values


def build_stats_rows(stats: list[EssayStats]) -> tuple[list[str], list[list[Any]]]:
    observed_pos_columns = set()
    for item in stats:
        observed_pos_columns.update(item.pos_counts.keys())

    other_pos_columns = sorted(column for column in observed_pos_columns if column not in POS_COLUMN_ORDER)
    pos_columns = [column for column in POS_COLUMN_ORDER if column in observed_pos_columns or column == "标点数"] + other_pos_columns
    headers = BASIC_HEADERS + TEXT_STAT_HEADERS + pos_columns + HSK_STAT_HEADERS

    rows: list[list[Any]] = []
    for item in stats:
        record = item.record
        rows.append(
            [
                record.code,
                record.title_code_name,
                record.essay_id,
                record.nationality,
                record.essay_topic,
                record.score,
                record.genre,
                record.filename,
                item.char_count,
                item.han_char_count,
                item.token_count,
                item.non_punct_token_count,
                item.unique_word_count,
                *[item.pos_counts.get(column, 0) for column in pos_columns],
                *hsk_stat_values(item),
            ]
        )

    return headers, rows


def save_stats_workbook(
    path: Path,
    stats: list[EssayStats],
    Workbook: Any,
    Alignment: Any,
    Font: Any,
    PatternFill: Any,
    get_column_letter: Any,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    headers, rows = build_stats_rows(stats)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "词性统计"
    sheet.append(headers)
    for row in rows:
        sheet.append(row)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    widths = {
        "A": 10,
        "B": 24,
        "C": 20,
        "D": 14,
        "E": 28,
        "F": 10,
        "G": 10,
        "H": 16,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for column_index in range(9, len(headers) + 1):
        sheet.column_dimensions[get_column_letter(column_index)].width = 12
    for column_index, header in enumerate(headers, start=1):
        if header in HSK_STAT_HEADERS:
            sheet.column_dimensions[get_column_letter(column_index)].width = 16

    percentage_columns = {
        column_index
        for column_index, header in enumerate(headers, start=1)
        if header.endswith("词汇占比")
    }
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top")
            if cell.column in percentage_columns:
                cell.number_format = "0.00%"

    with NamedTemporaryFile("wb", dir=path.parent, prefix=f".{path.name}.", suffix=".tmp", delete=False) as tmp:
        tmp_path = Path(tmp.name)
    workbook.save(tmp_path)
    tmp_path.replace(path)


def validate_stats(stats: list[EssayStats]) -> None:
    for item in stats:
        summed = sum(item.pos_counts.values())
        punct_count = item.pos_counts.get(POS_COLUMN_MAP[PUNCT_POS], 0)
        if summed != item.token_count:
            raise ValueError(f"{item.record.filename} 分词数不等于词性计数总和：{item.token_count} != {summed}")
        if item.non_punct_token_count != item.token_count - punct_count:
            raise ValueError(f"{item.record.filename} 非标点分词数不等于分词数减标点数")
        if not 0 <= item.unique_word_count <= item.non_punct_token_count:
            raise ValueError(
                f"{item.record.filename} 去重词数超出有效范围："
                f"{item.unique_word_count} > {item.non_punct_token_count}"
            )
        unknown_levels = set(item.hsk_level_counts).difference(HSK_LEVELS)
        if unknown_levels:
            raise ValueError(f"{item.record.filename} 存在无效 HSK 主等级：{sorted(unknown_levels)}")
        hsk_count = sum(item.hsk_level_counts.values())
        if hsk_count > item.non_punct_token_count:
            raise ValueError(
                f"{item.record.filename} HSK 等级词汇次数超过非标点分词数："
                f"{hsk_count} > {item.non_punct_token_count}"
            )


def main() -> int:
    args = parse_args()
    validate_args(args)
    require_x86_64()
    pynlpir, Workbook, load_workbook, Alignment, Font, PatternFill, get_column_letter = import_runtime_dependencies()

    workbook_path = Path(args.workbook)
    input_dir = Path(args.input_dir)
    seg_output_dir = Path(args.seg_output_dir)
    stats_output = Path(args.stats_output)
    hsk_vocab_path = Path(args.hsk_vocab)

    records = read_records(workbook_path, args.limit, load_workbook)
    validate_sources(records, input_dir)
    hsk_vocabulary = read_hsk_vocabulary(hsk_vocab_path)

    print(f"作文主表：{workbook_path}")
    print(f"清洗文本：{input_dir}")
    print(f"分词输出：{seg_output_dir}")
    print(f"统计宽表：{stats_output}")
    print(f"HSK 词汇表：{hsk_vocab_path}（{sum(len(entries) for entries in hsk_vocabulary.values())} 条）")
    print(f"本次处理：{len(records)} 篇")

    if args.dry_run:
        for record in records[:5]:
            print(f"dry-run: {source_path_for(input_dir, record)} -> {seg_path_for(seg_output_dir, record)}")
        print("dry-run ok")
        return 0

    stats: list[EssayStats] = []
    pynlpir.open()
    try:
        for index, record in enumerate(records, start=1):
            item = segment_record(record, input_dir, seg_output_dir, pynlpir, hsk_vocabulary)
            stats.append(item)
            hsk_count = sum(item.hsk_level_counts.values())
            print(
                f"[{index}/{len(records)}] {record.code}/{record.filename}: "
                f"{item.token_count} tokens, {item.non_punct_token_count} non-punct, "
                f"{item.unique_word_count} unique, {hsk_count} HSK",
                flush=True,
            )
    finally:
        pynlpir.close()

    validate_stats(stats)
    save_stats_workbook(stats_output, stats, Workbook, Alignment, Font, PatternFill, get_column_letter)
    print(f"完成：分词文件 {len(stats)} 个；统计宽表 {stats_output}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        raise SystemExit(1)
