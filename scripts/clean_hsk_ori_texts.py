#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from tempfile import NamedTemporaryFile


DEFAULT_INPUT_DIR = "ori_text"
DEFAULT_OUTPUT_DIR = "clean_text"
ALLOWED_CODES = ("J1", "J2", "Y1", "Y2")

FULLWIDTH_MARKS = str.maketrans(
    {
        "｛": "{",
        "｝": "}",
        "［": "[",
        "］": "]",
    }
)

FOOTER_RE = re.compile(r"^(?:[（(][A-Za-z0-9]+[）)]\s*)+$")
META_RE = re.compile(r"^\s*\d{6,7}[^。\n]{0,80}\.\d+.*$")
CJ_PAYLOAD_RE = re.compile(r"^CJ([+-])([A-Za-z]*)(.*)$", re.DOTALL)
LEFTOVER_MARK_RE = re.compile(r"\[[A-Za-z#][^\]]{0,40}\]|\{[A-Za-z][^{}]{0,80}\}")


@dataclass(frozen=True)
class CleanTarget:
    code: str
    source_path: Path
    output_path: Path


@dataclass
class CleanResult:
    target: CleanTarget
    status: str
    leftover_markers: list[str]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="清洗 HSK 作文标注文本，保留正确正文并写入 ./clean_text/{代码}/{作文编码}.txt。",
    )
    parser.add_argument("--input-dir", default=DEFAULT_INPUT_DIR, help=f"输入目录，默认：{DEFAULT_INPUT_DIR}")
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR, help=f"输出目录，默认：{DEFAULT_OUTPUT_DIR}")
    parser.add_argument("--essay-id", default=None, help="只清洗指定作文编码，用于生成单篇 demo")
    parser.add_argument("--limit", type=int, default=None, help="只处理前 N 篇，用于小样本验证")
    parser.add_argument("--force", action="store_true", help="覆盖已存在的清洗文本；默认跳过")
    return parser.parse_args()


def validate_args(args: argparse.Namespace) -> None:
    if args.limit is not None and args.limit <= 0:
        raise ValueError("--limit 必须大于 0")


def discover_targets(input_dir: Path, output_dir: Path, essay_id: str | None, limit: int | None) -> list[CleanTarget]:
    if not input_dir.exists():
        raise FileNotFoundError(f"找不到输入目录：{input_dir}")

    targets: list[CleanTarget] = []
    for code in ALLOWED_CODES:
        code_dir = input_dir / code
        if not code_dir.exists():
            continue
        for source_path in sorted(code_dir.glob("*.txt")):
            if essay_id is not None and source_path.stem != essay_id:
                continue
            targets.append(
                CleanTarget(
                    code=code,
                    source_path=source_path,
                    output_path=output_dir / code / source_path.name,
                )
            )

    if essay_id is not None and not targets:
        raise FileNotFoundError(f"在 {input_dir} 下找不到作文编码：{essay_id}")

    if limit is not None:
        return targets[:limit]
    return targets


def remove_non_body_lines(raw_text: str) -> str:
    lines = raw_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    kept: list[str] = []
    skipped_meta = False

    for line in lines:
        stripped = line.strip()
        if not stripped:
            kept.append("")
            continue
        if not skipped_meta and META_RE.match(line):
            skipped_meta = True
            continue
        if stripped.lower().startswith("title"):
            continue
        if FOOTER_RE.fullmatch(stripped):
            continue
        kept.append(line)

    return "\n".join(kept)


def normalize_body_whitespace(text: str) -> str:
    normalized_lines: list[str] = []
    previous_blank = False

    for line in text.splitlines():
        cleaned_line = re.sub(r"[ \t\u3000]+", " ", line).strip()
        if not cleaned_line:
            if normalized_lines and not previous_blank:
                normalized_lines.append("")
            previous_blank = True
            continue
        normalized_lines.append(cleaned_line)
        previous_blank = False

    while normalized_lines and normalized_lines[-1] == "":
        normalized_lines.pop()
    return "\n".join(normalized_lines).strip() + "\n"


def clean_hsk_text(raw_text: str) -> str:
    body_text = remove_non_body_lines(raw_text).translate(FULLWIDTH_MARKS)
    cleaned = clean_annotations(body_text)
    return normalize_body_whitespace(cleaned)


def clean_annotations(text: str) -> str:
    output: list[str] = []
    index = 0
    while index < len(text):
        char = text[index]
        if char in "[{":
            end_index = find_matching_marker(text, index)
            if end_index == -1:
                output.append(char)
                index += 1
                continue

            raw_content = text[index + 1 : end_index]
            if char == "[":
                output.append(clean_square_marker(raw_content))
            else:
                output.append(clean_curly_marker(raw_content))
            index = end_index + 1
            continue

        output.append(char)
        index += 1

    return "".join(output)


def find_matching_marker(text: str, start_index: int) -> int:
    opening = text[start_index]
    closing_for = {"[": "]", "{": "}"}
    opening_for = {"]": "[", "}": "{"}
    if opening not in closing_for:
        return -1

    stack = [opening]
    for index in range(start_index + 1, len(text)):
        char = text[index]
        if char in closing_for:
            stack.append(char)
            continue
        if char in opening_for and stack and stack[-1] == opening_for[char]:
            stack.pop()
            if not stack:
                return index
    return -1


def clean_square_marker(raw_content: str) -> str:
    if raw_content.startswith("BQ"):
        return clean_annotations(raw_content[2:])
    if raw_content.startswith(("BC", "BD")):
        return ""
    if raw_content.startswith("#"):
        return ""
    if raw_content[:1] in {"C", "B", "L", "D", "F", "Y", "P"}:
        return ""
    if raw_content[:1] in {"c", "b", "l", "d", "f", "y", "p"}:
        return ""

    if looks_like_annotation(raw_content):
        return ""
    return f"[{clean_annotations(raw_content)}]"


def clean_curly_marker(raw_content: str) -> str:
    if raw_content.startswith("CP"):
        payload = raw_content[2:]
        if payload.endswith("P"):
            payload = payload[:-1]
        return clean_annotations(payload)

    if raw_content.startswith("CQ"):
        return clean_annotations(raw_content[2:])
    if raw_content.startswith(("CC", "CLH", "CD")):
        return ""
    if raw_content.startswith("W") and (len(raw_content) == 1 or raw_content[1].isdigit() or not raw_content[1].islower()):
        return ""

    cj_match = CJ_PAYLOAD_RE.match(raw_content)
    if cj_match:
        sign = cj_match.group(1)
        payload = cj_match.group(3)
        if sign == "-":
            return clean_annotations(payload)
        return ""
    if raw_content.startswith(("CJ", "WWJ")):
        return ""

    # The help page mostly uses square brackets for these, but a few examples use braces.
    if raw_content.startswith("BQ"):
        return clean_annotations(raw_content[2:])
    if raw_content.startswith(("BC", "BD")):
        return ""
    if raw_content[:1] in {"C", "B", "L", "D", "F", "Y", "P"}:
        return ""

    if looks_like_annotation(raw_content):
        return ""
    return "{" + clean_annotations(raw_content) + "}"


def looks_like_annotation(raw_content: str) -> bool:
    return bool(raw_content) and (raw_content[0].isupper() or raw_content[0] == "#")


def find_leftover_markers(text: str) -> list[str]:
    markers = LEFTOVER_MARK_RE.findall(text)
    seen: set[str] = set()
    unique_markers: list[str] = []
    for marker in markers:
        if marker not in seen:
            seen.add(marker)
            unique_markers.append(marker)
    return unique_markers[:10]


def atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(
        "w",
        encoding="utf-8",
        dir=path.parent,
        prefix=f".{path.name}.",
        suffix=".tmp",
        delete=False,
    ) as tmp:
        tmp.write(text)
        tmp_path = Path(tmp.name)
    tmp_path.replace(path)


def clean_target(target: CleanTarget, force: bool) -> CleanResult:
    if target.output_path.exists() and not force:
        text = target.output_path.read_text(encoding="utf-8")
        return CleanResult(target=target, status="skipped", leftover_markers=find_leftover_markers(text))

    raw_text = target.source_path.read_text(encoding="utf-8")
    cleaned_text = clean_hsk_text(raw_text)
    atomic_write_text(target.output_path, cleaned_text)
    return CleanResult(target=target, status="saved", leftover_markers=find_leftover_markers(cleaned_text))


def main() -> int:
    args = parse_args()
    validate_args(args)

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    targets = discover_targets(input_dir, output_dir, args.essay_id, args.limit)
    if not targets:
        print("没有找到需要清洗的 txt 文件", file=sys.stderr)
        return 1

    print(f"输入目录：{input_dir}")
    print(f"输出目录：{output_dir}")
    print(f"本次处理：{len(targets)} 篇")

    saved = 0
    skipped = 0
    warned = 0
    for index, target in enumerate(targets, start=1):
        result = clean_target(target, force=args.force)
        if result.status == "saved":
            saved += 1
        elif result.status == "skipped":
            skipped += 1

        print(
            f"[{index}/{len(targets)}] {result.status} "
            f"{target.code}/{target.source_path.stem} -> {target.output_path}",
            flush=True,
        )
        if result.leftover_markers:
            warned += 1
            print(f"  注意：疑似残留标注 {result.leftover_markers}", flush=True)

    print(f"完成：保存 {saved}，跳过 {skipped}，疑似残留 {warned}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        raise SystemExit(1)
