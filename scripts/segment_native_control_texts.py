#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import platform
import re
from pathlib import Path
from typing import Any, Sequence

try:
    from .segment_hsk_clean_texts import (
        DEFAULT_FEATURE_LEXICON,
        DEFAULT_HSK_VOCAB,
        DEFAULT_LONG_SENTENCE_THRESHOLD,
        DEFAULT_MATTR_WINDOW,
        EssayRecord,
        atomic_write_text,
        build_stats_rows,
        compile_feature_lexicon,
        import_runtime_dependencies,
        read_feature_lexicon,
        read_hsk_vocabulary,
        segment_record,
        source_path_for,
        validate_stats,
    )
except ImportError:
    from segment_hsk_clean_texts import (
        DEFAULT_FEATURE_LEXICON,
        DEFAULT_HSK_VOCAB,
        DEFAULT_LONG_SENTENCE_THRESHOLD,
        DEFAULT_MATTR_WINDOW,
        EssayRecord,
        atomic_write_text,
        build_stats_rows,
        compile_feature_lexicon,
        import_runtime_dependencies,
        read_feature_lexicon,
        read_hsk_vocabulary,
        segment_record,
        source_path_for,
        validate_stats,
    )


DEFAULT_SELECTED_JSON = "outputs/native_control/selected_samples.json"
DEFAULT_INPUT_DIR = "native_clean_text"
DEFAULT_SEG_OUTPUT_DIR = "native_seg_text"
DEFAULT_PAYLOAD = "outputs/native_control/native_stats_payload.json"
DEFAULT_LEARNER_WORKBOOK = "作文词性统计宽表.xlsx"
EXPECTED_NATIVE_CODES = ("NJ1", "NJ2", "NY1", "NY2")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="为作文网母语参照语料生成PyNLPIR分词和301项统计。")
    parser.add_argument("--selected-json", default=DEFAULT_SELECTED_JSON)
    parser.add_argument("--input-dir", default=DEFAULT_INPUT_DIR)
    parser.add_argument("--seg-output-dir", default=DEFAULT_SEG_OUTPUT_DIR)
    parser.add_argument("--stats-payload", default=DEFAULT_PAYLOAD)
    parser.add_argument("--learner-workbook", default=DEFAULT_LEARNER_WORKBOOK)
    parser.add_argument("--hsk-vocab", default=DEFAULT_HSK_VOCAB)
    parser.add_argument("--feature-lexicon", default=DEFAULT_FEATURE_LEXICON)
    parser.add_argument("--mattr-window", type=int, default=DEFAULT_MATTR_WINDOW)
    parser.add_argument("--long-sentence-threshold", type=int, default=DEFAULT_LONG_SENTENCE_THRESHOLD)
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--dry-run", action="store_true")
    return parser.parse_args()


def require_x86_64() -> None:
    if platform.machine() != "x86_64":
        raise RuntimeError(
            "PyNLPIR当前只能在x86_64 Python下运行。请使用：\n"
            "  arch -x86_64 /usr/bin/python3 scripts/segment_native_control_texts.py"
        )


def normalize_score(value: Any) -> int | str:
    return "" if value in {None, ""} else int(value)


def load_native_records(path: Path, limit: int | None = None) -> list[EssayRecord]:
    document = json.loads(path.read_text(encoding="utf-8"))
    rows = document.get("selected", [])
    records: list[EssayRecord] = []
    seen_files: set[str] = set()
    seen_ids: set[str] = set()
    for row in rows:
        code = str(row.get("母语代码", ""))
        if code not in EXPECTED_NATIVE_CODES:
            raise ValueError(f"无效母语代码：{code!r}")
        filename = str(row.get("作文文件名", ""))
        essay_id = str(row.get("网页文章ID", ""))
        if not filename or not essay_id:
            raise ValueError(f"母语记录缺少文件名或网页文章ID：{row}")
        if filename in seen_files:
            raise ValueError(f"作文文件名重复：{filename}")
        if essay_id in seen_ids:
            raise ValueError(f"网页文章ID重复：{essay_id}")
        seen_files.add(filename)
        seen_ids.add(essay_id)
        records.append(
            EssayRecord(
                sheet_name=code,
                code=code,
                title_code_name=str(row.get("对应学习者篇名", "")),
                essay_id=essay_id,
                nationality="中国（公开网络样本）",
                essay_topic=str(row.get("作文题目", "")),
                score=normalize_score(row.get("作文分数")),
                genre=str(row.get("体裁", "")),
                filename=filename,
            )
        )
    records.sort(key=lambda item: (item.code, item.filename))
    if limit is not None:
        if limit <= 0:
            raise ValueError("--limit必须大于0")
        records = records[:limit]
    if not records:
        raise ValueError("母语样本JSON中没有入选记录")
    return records


def read_learner_schema(path: Path, load_workbook: Any) -> tuple[list[str], list[dict[str, Any]]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    stats_sheet = workbook["词性统计"]
    headers = [str(cell.value or "") for cell in next(stats_sheet.iter_rows(min_row=1, max_row=1))]
    dictionary_sheet = workbook["字段说明"]
    dictionary_headers = [str(cell.value or "") for cell in next(dictionary_sheet.iter_rows(min_row=1, max_row=1))]
    dictionary_rows = [
        {dictionary_headers[index]: value for index, value in enumerate(row)}
        for row in dictionary_sheet.iter_rows(min_row=2, values_only=True)
    ]
    workbook.close()
    if len(headers) != 301:
        raise ValueError(f"学习者主宽表应有301列，实际{len(headers)}")
    return headers, dictionary_rows


def align_rows_to_learner_schema(
    fields: Sequence[Any],
    rows: Sequence[Sequence[Any]],
    learner_headers: Sequence[str],
) -> list[list[Any]]:
    source_names = [field.name for field in fields]
    extras = sorted(set(source_names).difference(learner_headers))
    missing = sorted(set(learner_headers).difference(source_names))
    fillable_missing = {
        name for name in missing if re.fullmatch(r"其他词性_.+(?:数|每千字)", name)
    }
    unfillable_missing = sorted(set(missing).difference(fillable_missing))
    if extras or unfillable_missing:
        raise ValueError(
            "母语统计字段与学习者301列不一致；"
            f"新增={extras}，缺少={unfillable_missing}"
        )
    index = {name: position for position, name in enumerate(source_names)}
    return [
        [row[index[name]] if name in index else 0 for name in learner_headers]
        for row in rows
    ]


def json_value(value: Any) -> Any:
    if isinstance(value, float):
        if value != value or value in {float("inf"), float("-inf")}:
            return None
        return value
    if isinstance(value, (str, int, bool)) or value is None:
        return value
    return str(value)


def prune_stale_segmented_texts(base_dir: Path, records: Sequence[EssayRecord]) -> list[Path]:
    expected = {
        code: {record.filename for record in records if record.code == code}
        for code in EXPECTED_NATIVE_CODES
    }
    removed: list[Path] = []
    for code, allowed in expected.items():
        code_dir = base_dir / code
        if not code_dir.is_dir():
            continue
        for path in code_dir.glob("*.txt"):
            if path.stem not in allowed:
                path.unlink()
                removed.append(path)
    return removed


def main() -> int:
    args = parse_args()
    require_x86_64()
    selected_json = Path(args.selected_json)
    input_dir = Path(args.input_dir)
    seg_output_dir = Path(args.seg_output_dir)
    stats_payload = Path(args.stats_payload)
    learner_workbook = Path(args.learner_workbook)
    hsk_vocab_path = Path(args.hsk_vocab)
    feature_lexicon_path = Path(args.feature_lexicon)

    pynlpir, _, load_workbook, _, _, _, _ = import_runtime_dependencies()
    records = load_native_records(selected_json, args.limit)
    missing_sources = [str(source_path_for(input_dir, record)) for record in records if not source_path_for(input_dir, record).is_file()]
    if missing_sources:
        raise FileNotFoundError(f"缺少母语清洗文本：{missing_sources[:20]}")
    learner_headers, learner_dictionary = read_learner_schema(learner_workbook, load_workbook)
    hsk_vocabulary = read_hsk_vocabulary(hsk_vocab_path)
    feature_lexicon_specs = read_feature_lexicon(feature_lexicon_path)
    print(f"母语样本：{selected_json}（{len(records)}篇）")
    print(f"清洗文本：{input_dir}")
    print(f"分词输出：{seg_output_dir}")
    print(f"统计载荷：{stats_payload}")
    if args.dry_run:
        for record in records[:8]:
            print(f"dry-run: {source_path_for(input_dir, record)}")
        return 0

    removed = prune_stale_segmented_texts(seg_output_dir, records)
    if removed:
        print(f"已清理旧版分词文本：{len(removed)} 个")
    stats = []
    pynlpir.open()
    try:
        feature_lexicon = compile_feature_lexicon(
            feature_lexicon_specs,
            lambda term: pynlpir.segment(term, pos_names=None),
        )
        for index, record in enumerate(records, start=1):
            stats.append(
                segment_record(
                    record=record,
                    input_dir=input_dir,
                    seg_output_dir=seg_output_dir,
                    pynlpir=pynlpir,
                    hsk_vocabulary=hsk_vocabulary,
                    feature_lexicon=feature_lexicon,
                    mattr_window=args.mattr_window,
                    long_sentence_threshold=args.long_sentence_threshold,
                )
            )
            print(f"[{index:03d}/{len(records):03d}] {record.code}/{record.filename}.txt", flush=True)
    finally:
        pynlpir.close()

    validate_stats(stats)
    fields, rows = build_stats_rows(stats)
    aligned_rows = align_rows_to_learner_schema(fields, rows, learner_headers)
    payload = {
        "generated_at": __import__("datetime").datetime.now().astimezone().isoformat(timespec="seconds"),
        "headers": list(learner_headers),
        "rows": [[json_value(value) for value in row] for row in aligned_rows],
        "field_dictionary": [
            {str(key): json_value(value) for key, value in row.items()} for row in learner_dictionary
        ],
        "validation": {
            "rows": len(aligned_rows),
            "columns": len(learner_headers),
            "code_counts": {
                code: sum(record.code == code for record in records) for code in EXPECTED_NATIVE_CODES
            },
        },
    }
    atomic_write_text(stats_payload, json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
    print(f"完成：{len(aligned_rows)}行 x {len(learner_headers)}列")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
