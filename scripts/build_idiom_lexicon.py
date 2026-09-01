#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import urllib.request
from pathlib import Path


IDIOM_URL = "https://raw.githubusercontent.com/pwxcoo/chinese-xinhua/master/data/idiom.json"
XIEHOUYU_URL = "https://raw.githubusercontent.com/pwxcoo/chinese-xinhua/master/data/xiehouyu.json"
DEFAULT_CACHE_DIR = Path("tmp/idiom_lexicon")
DEFAULT_OUTPUT = Path("resources/熟语词表.csv")

AUDITED_PROVERBS = (
    "言教不如身教",
    "望子成龙、望女成凤",
    "孩子是父母的镜子",
    "孩子看着父母而长大",
    "人之初，性本善",
    "吸烟有百害而无一利",
    "桂林山水甲天下",
    "夏天南京像个火锅",
    "家和万事兴",
    "上有天堂，下有苏杭",
    "退一步海阔天空，忍一时风平浪静",
    "当局者迷，旁观者清",
    "天下兴亡，匹夫有责",
    "人要脸树要皮",
    "你敬我一尺，我还你一丈",
    "远亲不如近邻",
    "自立人生少年始",
    "自食其力，生活是甘甜的；卑躬屈膝，生活是酸苦的",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="从公开词典数据生成当前项目语料命中的熟语词表。")
    parser.add_argument("--cache-dir", type=Path, default=DEFAULT_CACHE_DIR)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--clean-dir", action="append", type=Path, default=[])
    return parser.parse_args()


def download_if_missing(url: str, path: Path) -> None:
    if path.is_file():
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 HSK-corpus-research/1.0"})
    with urllib.request.urlopen(request, timeout=60) as response:
        path.write_bytes(response.read())


def load_corpus(clean_dirs: list[Path]) -> list[str]:
    texts: list[str] = []
    for base_dir in clean_dirs:
        if not base_dir.is_dir():
            continue
        texts.extend(path.read_text(encoding="utf-8") for path in sorted(base_dir.glob("*/*.txt")))
    if not texts:
        raise FileNotFoundError("没有找到清洗后的作文文本")
    return texts


def main() -> int:
    args = parse_args()
    clean_dirs = args.clean_dir or [Path("clean_text"), Path("native_clean_text")]
    idiom_json = args.cache_dir / "idiom.json"
    xiehouyu_json = args.cache_dir / "xiehouyu.json"
    download_if_missing(IDIOM_URL, idiom_json)
    download_if_missing(XIEHOUYU_URL, xiehouyu_json)

    texts = load_corpus(clean_dirs)
    corpus = "\n".join(texts)
    proverb_set = {term for term in AUDITED_PROVERBS if term in corpus}
    rows: set[tuple[str, str, str, str]] = set()

    idioms = json.loads(idiom_json.read_text(encoding="utf-8"))
    for item in idioms:
        term = str(item.get("word", "")).strip()
        if len(term) < 3 or term not in corpus or term in proverb_set:
            continue
        rows.add(("成语", term, "", "chinese-xinhua成语数据；按当前800篇语料命中筛选"))

    xiehouyu = json.loads(xiehouyu_json.read_text(encoding="utf-8"))
    for item in xiehouyu:
        first = str(item.get("riddle", "")).strip()
        second = str(item.get("answer", "")).strip()
        if not first or not second:
            continue
        if any(first in text and second in text for text in texts):
            rows.add(("歇后语", first, second, "chinese-xinhua歇后语数据；同篇作文前后项共现"))

    for term in proverb_set:
        rows.add(("谚语", term, "", "当前语料中由“俗话说/常言道/谚语”等显式引介语人工复核"))

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, lineterminator="\n")
        writer.writerow(("类型", "前项", "后项", "来源说明"))
        writer.writerows(sorted(rows, key=lambda row: (("成语", "歇后语", "惯用语", "谚语").index(row[0]), row[1], row[2])))

    counts = {idiom_type: sum(row[0] == idiom_type for row in rows) for idiom_type in ("成语", "歇后语", "惯用语", "谚语")}
    print(f"写入 {args.output}：{len(rows)} 条；{counts}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
