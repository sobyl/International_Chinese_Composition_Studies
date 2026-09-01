#!/usr/bin/env python3
from __future__ import annotations

import csv
import math
import re
import statistics
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence


HAN_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
HAN_TOKEN_RE = re.compile(r"^[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]+$")
ASCII_LETTER_RE = re.compile(r"[A-Za-z]")
DIGIT_RE = re.compile(r"\d")
SENTENCE_BOUNDARY_RE = re.compile(r"[。！？!?]+")
CLAUSE_BOUNDARY_RE = re.compile(r"[，,；;：:。！？!?]+")
QUOTE_PATTERNS = (
    re.compile(r"“[^”]*”", re.DOTALL),
    re.compile(r"「[^」]*」", re.DOTALL),
    re.compile(r"『[^』]*』", re.DOTALL),
    re.compile(r'"[^"\n]*"'),
)

LEXICON_REQUIRED_COLUMNS = ("大类", "特征名", "词项", "允许词性前缀", "来源说明")

CATEGORY_LEXICAL_DIVERSITY = "词汇丰富度"
CATEGORY_LEXICAL_DENSITY = "词汇密度与词长"
CATEGORY_STRUCTURE = "句段结构"
CATEGORY_GRAMMAR = "语法标记"
CATEGORY_IDIOM = "熟语"
CATEGORY_COHESION = "复句关系"
CATEGORY_NARRATIVE = "记叙描写"
CATEGORY_HSK = "HSK"

GRAMMAR_LEXICON_FEATURES = (
    "第一人称单数代词",
    "第一人称复数代词",
    "其他第一人称代词",
    "第二人称代词",
    "第三人称代词",
    "非人第三人称代词",
    "指示代词",
    "不定代词",
    "疑问代词",
    "情态动词",
    "否定词",
)

PAPER_GRAMMAR_LEXICON_FEATURES = (
    "疑问副词",
    "限定词",
    "低调词与模糊语",
    "夸张与加强语",
    "表可能性情态词",
    "表必要性情态词",
    "表意愿性情态词",
    "表语性情态动词",
)

TIME_NOUN_LEXICON_FEATURES = (
    "过去时间名词",
    "现在时间名词",
    "未来时间名词",
)

TIME_ADVERB_LEXICON_FEATURES = (
    "表完成时态时间副词",
    "表过去时间副词",
    "表正在进行时间副词",
    "表将来时态时间副词",
)

VERB_CLASS_LEXICON_FEATURES = (
    "私人性动词",
    "建议要求类动词",
    "公共性动词",
)

COMPOUND_SENTENCE_FEATURES = (
    "因果复句",
    "转折复句",
    "条件复句",
    "假设复句",
    "目的复句",
    "递进复句",
    "并列复句",
    "承接复句",
    "解说复句",
)

COMPOUND_MARKER_BY_FEATURE = {
    feature: f"{feature}标记" for feature in COMPOUND_SENTENCE_FEATURES
}

COMPOUND_MARKER_FEATURES = tuple(COMPOUND_MARKER_BY_FEATURE.values())

LEGACY_CONNECTIVE_FEATURES = (
    "因果连接词", "转折连接词", "条件连接词", "递进连接词",
    "并列连接词", "顺序连接词", "总结连接词", "举例连接词",
)

NARRATIVE_LEXICON_FEATURES = (
    "时间副词",
    "动作动词",
    "心理动词",
    "言说动词",
    "程度副词",
    "正面评价词",
    "负面评价词",
)

LEXICON_FEATURES = frozenset(
    (
        *GRAMMAR_LEXICON_FEATURES,
        *PAPER_GRAMMAR_LEXICON_FEATURES,
        *TIME_NOUN_LEXICON_FEATURES,
        *TIME_ADVERB_LEXICON_FEATURES,
        *VERB_CLASS_LEXICON_FEATURES,
        *COMPOUND_MARKER_FEATURES,
        *LEGACY_CONNECTIVE_FEATURES,
        *NARRATIVE_LEXICON_FEATURES,
    )
)

IDIOM_TYPES = ("成语", "歇后语", "惯用语", "谚语")
IDIOM_REQUIRED_COLUMNS = ("类型", "前项", "后项", "来源说明")

RAW_POS_PARENT = {
    "n": "noun",
    "t": "time word",
    "s": "locative word",
    "f": "noun of locality",
    "v": "verb",
    "a": "adjective",
    "b": "distinguishing word",
    "z": "status word",
    "r": "pronoun",
    "m": "numeral",
    "q": "classifier",
    "d": "adverb",
    "p": "preposition",
    "c": "conjunction",
    "u": "particle",
    "e": "interjection",
    "y": "modal particle",
    "o": "onomatopoeia",
    "h": "prefix",
    "k": "suffix",
    "x": "string",
    "w": "punctuation mark",
    "g": "multiword expression",
    "j": "abbreviation",
}

RAW_POS_EXACT_PARENT = {
    "happ": "adjective",
    "gjtgj": "noun",
    "gms": "noun",
    "gwqz": "verb",
}

RAW_POS_EXACT_ROOT = {
    "happ": "a",
    "gjtgj": "n",
    "gms": "n",
    "gwqz": "v",
}

CONTENT_POS_ROOTS = frozenset({"n", "t", "s", "f", "v", "a", "b", "z", "d"})
REAL_WORD_POS_ROOTS = frozenset({"n", "t", "s", "f", "v", "a", "b", "z", "r", "m", "q", "o"})
FUNCTION_WORD_POS_ROOTS = frozenset({"d", "p", "c", "u", "y"})
PROPER_NOUN_PREFIXES = ("nr", "ns", "nt", "nz")
END_SENTENCE_POS = frozenset({"wj", "ww", "wt"})


@dataclass(frozen=True)
class Token:
    word: str
    raw_pos: str
    parent_pos: str
    paragraph_index: int
    hsk_level: str | None = None


@dataclass(frozen=True)
class LexiconSpec:
    category: str
    feature_name: str
    term: str
    allowed_pos_prefixes: tuple[str, ...]
    source_note: str


@dataclass(frozen=True)
class CompiledLexiconEntry:
    category: str
    feature_name: str
    term: str
    words: tuple[str, ...]
    allowed_pos_prefixes: tuple[str, ...]
    source_note: str


@dataclass(frozen=True)
class LexiconMatch:
    start: int
    end: int
    term: str


@dataclass(frozen=True)
class IdiomSpec:
    idiom_type: str
    first: str
    second: str
    source_note: str


@dataclass(frozen=True)
class CompiledIdiomEntry:
    idiom_type: str
    first: str
    second: str
    first_words: tuple[str, ...]
    second_words: tuple[str, ...]
    source_note: str


@dataclass(frozen=True)
class FieldSpec:
    name: str
    category: str
    definition: str
    formula: str
    unit: str
    number_format: str


@dataclass(frozen=True)
class FeatureResult:
    values: dict[str, int | float]
    fields: tuple[FieldSpec, ...]


class FeatureBuilder:
    def __init__(self, han_char_count: int) -> None:
        self.han_char_count = han_char_count
        self.values: dict[str, int | float] = {}
        self.fields: list[FieldSpec] = []

    def add(
        self,
        name: str,
        value: int | float,
        category: str,
        definition: str,
        formula: str,
        unit: str,
        number_format: str,
    ) -> None:
        if name in self.values:
            raise ValueError(f"语言特征字段重复：{name}")
        self.values[name] = value
        self.fields.append(
            FieldSpec(
                name=name,
                category=category,
                definition=definition,
                formula=formula,
                unit=unit,
                number_format=number_format,
            )
        )

    def add_count(
        self,
        name: str,
        count: int,
        category: str,
        definition: str,
        *,
        add_per_thousand: bool = True,
    ) -> None:
        self.add(name, count, category, definition, "直接计数", "次/个", "0")
        if add_per_thousand:
            if name.endswith("次数"):
                rate_name = f"{name[:-2]}每千字"
            elif name.endswith("数"):
                rate_name = f"{name[:-1]}每千字"
            else:
                rate_name = f"{name}每千字"
            rate = count * 1000 / self.han_char_count if self.han_char_count else 0.0
            self.add(
                rate_name,
                rate,
                category,
                f"{definition}的每千汉字标准化频率",
                f"{name} ÷ 纯文本字数 × 1000",
                "次/千汉字",
                "0.00",
            )

    def result(self) -> FeatureResult:
        return FeatureResult(values=dict(self.values), fields=tuple(self.fields))


def raw_pos_parent(raw_pos: str | None) -> str:
    normalized = (raw_pos or "").strip().lower()
    if not normalized:
        return ""
    return RAW_POS_EXACT_PARENT.get(normalized, RAW_POS_PARENT.get(normalized[:1], ""))


def raw_pos_root(raw_pos: str | None) -> str:
    normalized = (raw_pos or "").strip().lower()
    return RAW_POS_EXACT_ROOT.get(normalized, normalized[:1]) if normalized else ""


def is_punctuation(token: Token) -> bool:
    return token.raw_pos.lower().startswith("w")


def safe_ratio(numerator: int | float, denominator: int | float) -> float:
    return numerator / denominator if denominator else 0.0


def mattr(words: Sequence[str], window_size: int) -> float:
    if window_size <= 0:
        raise ValueError("MATTR 窗口必须大于 0")
    if not words:
        return 0.0
    if len(words) <= window_size:
        return len(set(words)) / len(words)
    windows = len(words) - window_size + 1
    return sum(len(set(words[index : index + window_size])) / window_size for index in range(windows)) / windows


def read_feature_lexicon(path: Path) -> list[LexiconSpec]:
    if not path.is_file():
        raise FileNotFoundError(f"找不到语言特征词表：{path}")

    specs: list[LexiconSpec] = []
    seen: set[tuple[str, str]] = set()
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        headers = reader.fieldnames or []
        missing = [column for column in LEXICON_REQUIRED_COLUMNS if column not in headers]
        if missing:
            raise ValueError(f"语言特征词表缺少列：{missing}；实际表头：{headers}")
        for row_number, row in enumerate(reader, start=2):
            category = (row.get("大类") or "").strip()
            feature_name = (row.get("特征名") or "").strip()
            term = (row.get("词项") or "").strip()
            raw_prefixes = (row.get("允许词性前缀") or "").strip()
            source_note = (row.get("来源说明") or "").strip()
            if not category or not feature_name or not term:
                raise ValueError(f"语言特征词表第 {row_number} 行大类、特征名和词项不能为空")
            if feature_name not in LEXICON_FEATURES:
                raise ValueError(f"语言特征词表第 {row_number} 行含未知特征：{feature_name}")
            key = (feature_name, term)
            if key in seen:
                raise ValueError(f"语言特征词表重复词项：{feature_name}/{term}")
            seen.add(key)
            prefixes = tuple(item.strip() for item in raw_prefixes.split("|") if item.strip())
            specs.append(
                LexiconSpec(
                    category=category,
                    feature_name=feature_name,
                    term=term,
                    allowed_pos_prefixes=prefixes,
                    source_note=source_note,
                )
            )

    return specs


def validate_feature_lexicon_specs(specs: Sequence[LexiconSpec]) -> None:
    missing_features = sorted(LEXICON_FEATURES.difference(spec.feature_name for spec in specs))
    if missing_features:
        raise ValueError(f"合并后的语言特征词表缺少特征：{missing_features}")

    duplicate_keys = [
        key
        for key, count in Counter((spec.feature_name, spec.term) for spec in specs).items()
        if count > 1
    ]
    if duplicate_keys:
        raise ValueError(f"合并后的语言特征词表存在重复项：{duplicate_keys[:20]}")


def read_idiom_lexicon(path: Path) -> list[IdiomSpec]:
    if not path.is_file():
        raise FileNotFoundError(f"找不到熟语词表：{path}")

    specs: list[IdiomSpec] = []
    seen: set[tuple[str, str, str]] = set()
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        headers = reader.fieldnames or []
        missing = [column for column in IDIOM_REQUIRED_COLUMNS if column not in headers]
        if missing:
            raise ValueError(f"熟语词表缺少列：{missing}；实际表头：{headers}")
        for row_number, row in enumerate(reader, start=2):
            idiom_type = (row.get("类型") or "").strip()
            first = (row.get("前项") or "").strip()
            second = (row.get("后项") or "").strip()
            source_note = (row.get("来源说明") or "").strip()
            if idiom_type not in IDIOM_TYPES:
                raise ValueError(f"熟语词表第 {row_number} 行类型无效：{idiom_type!r}")
            if not first:
                raise ValueError(f"熟语词表第 {row_number} 行前项不能为空")
            if idiom_type == "歇后语" and not second:
                raise ValueError(f"熟语词表第 {row_number} 行歇后语必须有后项")
            key = (idiom_type, first, second)
            if key in seen:
                raise ValueError(f"熟语词表重复：{key}")
            seen.add(key)
            specs.append(IdiomSpec(idiom_type, first, second, source_note))
    return specs


def compile_feature_lexicon(
    specs: Iterable[LexiconSpec],
    segmenter: Callable[[str], Sequence[tuple[str, str | None]]],
) -> tuple[CompiledLexiconEntry, ...]:
    compiled: list[CompiledLexiconEntry] = []
    for spec in specs:
        words = tuple(str(word).strip() for word, _ in segmenter(spec.term) if str(word).strip())
        if not words:
            raise ValueError(f"语言特征词项分词后为空：{spec.feature_name}/{spec.term}")
        compiled.append(
            CompiledLexiconEntry(
                category=spec.category,
                feature_name=spec.feature_name,
                term=spec.term,
                words=words,
                allowed_pos_prefixes=spec.allowed_pos_prefixes,
                source_note=spec.source_note,
            )
        )
    return tuple(compiled)


def compile_idiom_lexicon(
    specs: Iterable[IdiomSpec],
    segmenter: Callable[[str], Sequence[tuple[str, str | None]]],
) -> tuple[CompiledIdiomEntry, ...]:
    compiled: list[CompiledIdiomEntry] = []
    for spec in specs:
        first_words = tuple(str(word).strip() for word, _ in segmenter(spec.first) if str(word).strip())
        second_words = tuple(str(word).strip() for word, _ in segmenter(spec.second) if str(word).strip())
        if not first_words or (spec.idiom_type == "歇后语" and not second_words):
            raise ValueError(f"熟语分词后为空：{spec.idiom_type}/{spec.first}/{spec.second}")
        compiled.append(
            CompiledIdiomEntry(
                idiom_type=spec.idiom_type,
                first=spec.first,
                second=spec.second,
                first_words=first_words,
                second_words=second_words,
                source_note=spec.source_note,
            )
        )
    return tuple(compiled)


def match_lexicon_feature_spans(
    tokens: Sequence[Token],
    entries: Sequence[CompiledLexiconEntry],
) -> dict[str, list[LexiconMatch]]:
    entries_by_feature: dict[str, list[CompiledLexiconEntry]] = defaultdict(list)
    for entry in entries:
        entries_by_feature[entry.feature_name].append(entry)
    for feature_entries in entries_by_feature.values():
        feature_entries.sort(key=lambda entry: (-len(entry.words), entry.term))

    matches: dict[str, list[LexiconMatch]] = defaultdict(list)
    words = [token.word for token in tokens]
    for feature_name in LEXICON_FEATURES:
        feature_entries = entries_by_feature.get(feature_name, [])
        index = 0
        while index < len(tokens):
            matched: CompiledLexiconEntry | None = None
            for entry in feature_entries:
                end = index + len(entry.words)
                if end > len(tokens) or tuple(words[index:end]) != entry.words:
                    continue
                if len({token.paragraph_index for token in tokens[index:end]}) != 1:
                    continue
                if entry.allowed_pos_prefixes and not any(
                    tokens[index].raw_pos.startswith(prefix) for prefix in entry.allowed_pos_prefixes
                ):
                    continue
                matched = entry
                break
            if matched is None:
                index += 1
                continue
            end = index + len(matched.words)
            matches[feature_name].append(LexiconMatch(index, end, matched.term))
            index = end
    return matches


def match_lexicon_features(
    tokens: Sequence[Token],
    entries: Sequence[CompiledLexiconEntry],
) -> tuple[Counter[str], dict[str, set[str]]]:
    spans = match_lexicon_feature_spans(tokens, entries)
    counts = Counter({feature: len(feature_spans) for feature, feature_spans in spans.items()})
    matched_terms = {
        feature: {match.term for match in feature_spans}
        for feature, feature_spans in spans.items()
    }
    return counts, matched_terms


def sentence_token_ranges(tokens: Sequence[Token]) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    start = 0
    for index, token in enumerate(tokens):
        if token.raw_pos in END_SENTENCE_POS:
            end = index + 1
            if any(not is_punctuation(item) for item in tokens[start:end]):
                ranges.append((start, end))
            start = end
    if start < len(tokens) and any(not is_punctuation(item) for item in tokens[start:]):
        ranges.append((start, len(tokens)))
    return ranges


def _contains_sequence(words: Sequence[str], sequence: Sequence[str]) -> bool:
    if not sequence or len(sequence) > len(words):
        return False
    return any(tuple(words[index : index + len(sequence)]) == tuple(sequence) for index in range(len(words) - len(sequence) + 1))


def match_idiom_features(
    tokens: Sequence[Token],
    entries: Sequence[CompiledIdiomEntry],
) -> tuple[Counter[str], dict[str, set[str]]]:
    counts: Counter[str] = Counter()
    terms: dict[str, set[str]] = defaultdict(set)
    covered: set[int] = set()
    words = [token.word for token in tokens]

    for start, end in sentence_token_ranges(tokens):
        sentence_words = words[start:end]
        for entry in entries:
            if entry.idiom_type != "歇后语":
                continue
            if _contains_sequence(sentence_words, entry.first_words) and _contains_sequence(sentence_words, entry.second_words):
                counts["歇后语"] += 1
                terms["歇后语"].add(f"{entry.first}—{entry.second}")
                covered.update(range(start, end))

    sequence_entries = sorted(
        (entry for entry in entries if entry.idiom_type != "歇后语"),
        key=lambda entry: (-len(entry.first_words), entry.first),
    )
    for index in range(len(tokens)):
        if index in covered:
            continue
        for entry in sequence_entries:
            end = index + len(entry.first_words)
            if end > len(tokens) or tuple(words[index:end]) != entry.first_words:
                continue
            if len({token.paragraph_index for token in tokens[index:end]}) != 1:
                continue
            counts[entry.idiom_type] += 1
            terms[entry.idiom_type].add(entry.first)
            covered.update(range(index, end))
            break

    phrase_positions = {
        index
        for index, token in enumerate(tokens)
        if token.raw_pos.lower() in {"l", "nl", "vl", "al", "bl"} and index not in covered
    }
    counts["惯用语"] += len(phrase_positions)
    terms["惯用语"].update(tokens[index].word for index in phrase_positions)
    return counts, terms


def sentence_token_lengths(tokens: Sequence[Token]) -> list[int]:
    lengths: list[int] = []
    current = 0
    previous_was_boundary = False
    for token in tokens:
        if token.raw_pos in END_SENTENCE_POS:
            if current or not previous_was_boundary:
                if current:
                    lengths.append(current)
                current = 0
            previous_was_boundary = True
            continue
        if not is_punctuation(token):
            current += 1
        previous_was_boundary = False
    if current:
        lengths.append(current)
    return lengths


def count_verb_runs(tokens: Sequence[Token]) -> tuple[int, int]:
    run_count = 0
    longest = 0
    current = 0
    for token in tokens:
        if token.raw_pos in END_SENTENCE_POS:
            if current >= 2:
                run_count += 1
            longest = max(longest, current)
            current = 0
            continue
        if token.raw_pos.startswith("v"):
            current += 1
            continue
        if current >= 2:
            run_count += 1
        longest = max(longest, current)
        current = 0
    if current >= 2:
        run_count += 1
    return run_count, max(longest, current)


def count_direct_quotes(text: str) -> int:
    return sum(len(pattern.findall(text)) for pattern in QUOTE_PATTERNS)


def _add_lexical_diversity(
    builder: FeatureBuilder,
    non_punct_tokens: Sequence[Token],
    mattr_window: int,
) -> None:
    words = [token.word for token in non_punct_tokens]
    counts = Counter(words)
    token_count = len(words)
    type_count = len(counts)
    hapax_count = sum(count == 1 for count in counts.values())

    builder.add(
        "词形丰富度TTR",
        safe_ratio(type_count, token_count),
        CATEGORY_LEXICAL_DIVERSITY,
        "全文非标点词形的类型符号比",
        "去重词数 ÷ 非标点分词数",
        "比例",
        "0.0000",
    )
    builder.add(
        "Guiraud值",
        type_count / math.sqrt(token_count) if token_count else 0.0,
        CATEGORY_LEXICAL_DIVERSITY,
        "对篇幅影响进行修正的词汇丰富度",
        "去重词数 ÷ √非标点分词数",
        "指数",
        "0.0000",
    )
    builder.add(
        f"MATTR-{mattr_window}",
        mattr(words, mattr_window),
        CATEGORY_LEXICAL_DIVERSITY,
        f"长度为{mattr_window}个token的移动窗口平均TTR",
        f"所有{mattr_window}-token窗口TTR的平均值；短文本使用全文TTR",
        "比例",
        "0.0000",
    )
    builder.add_count(
        "仅出现一次词数",
        hapax_count,
        CATEGORY_LEXICAL_DIVERSITY,
        "全文中恰好出现一次的非标点词种数",
    )
    builder.add(
        "仅出现一次词占比",
        safe_ratio(hapax_count, type_count),
        CATEGORY_LEXICAL_DIVERSITY,
        "仅出现一次词种占全部去重词种的比例",
        "仅出现一次词数 ÷ 去重词数",
        "百分比",
        "0.00%",
    )

    top_ten_count = sum(count for _, count in counts.most_common(10))
    builder.add_count(
        "篇内高频词前10位次数",
        top_ten_count,
        CATEGORY_LEXICAL_DIVERSITY,
        "篇内出现频率最高的10个非标点词形的token次数之和",
    )
    builder.add(
        "篇内高频词前10位占比",
        safe_ratio(top_ten_count, token_count),
        CATEGORY_LEXICAL_DIVERSITY,
        "篇内最高频10个词形覆盖全部非标点token的比例",
        "篇内高频词前10位次数 ÷ 非标点分词数",
        "百分比",
        "0.00%",
    )

    pos_groups = (
        ("名词", "noun"),
        ("动词", "verb"),
        ("形容词", ("adjective", "distinguishing word", "status word")),
        ("副词", "adverb"),
    )
    for label, parent_pos in pos_groups:
        parents = (parent_pos,) if isinstance(parent_pos, str) else parent_pos
        pos_words = [token.word for token in non_punct_tokens if token.parent_pos in parents]
        pos_type_count = len(set(pos_words))
        builder.add_count(
            f"{label}去重词数",
            pos_type_count,
            CATEGORY_LEXICAL_DIVERSITY,
            f"{label}词形的去重数量",
        )
        builder.add(
            f"{label}TTR",
            safe_ratio(pos_type_count, len(pos_words)),
            CATEGORY_LEXICAL_DIVERSITY,
            f"{label}内部的类型符号比",
            f"{label}去重词数 ÷ {label}token数",
            "比例",
            "0.0000",
        )


def _add_lexical_density_and_length(builder: FeatureBuilder, non_punct_tokens: Sequence[Token]) -> None:
    roots = [raw_pos_root(token.raw_pos) for token in non_punct_tokens]
    content_count = sum(root in CONTENT_POS_ROOTS for root in roots)
    real_count = sum(root in REAL_WORD_POS_ROOTS for root in roots)
    function_count = sum(root in FUNCTION_WORD_POS_ROOTS for root in roots)
    other_count = len(non_punct_tokens) - real_count - function_count

    builder.add_count("内容词数", content_count, CATEGORY_LEXICAL_DENSITY, "名词性、动词性、形容词性及副词token数")
    builder.add(
        "词汇密度",
        safe_ratio(content_count, len(non_punct_tokens)),
        CATEGORY_LEXICAL_DENSITY,
        "内容词在全部非标点token中的比例",
        "内容词数 ÷ 非标点分词数",
        "百分比",
        "0.00%",
    )
    builder.add_count("实词数", real_count, CATEGORY_LEXICAL_DENSITY, "现代汉语口径的实词token数")
    builder.add_count("虚词数", function_count, CATEGORY_LEXICAL_DENSITY, "现代汉语口径的虚词token数")
    builder.add_count("其他词数", other_count, CATEGORY_LEXICAL_DENSITY, "未归入实词或虚词的非标点token数")
    builder.add(
        "实词率",
        safe_ratio(real_count, len(non_punct_tokens)),
        CATEGORY_LEXICAL_DENSITY,
        "实词在全部非标点token中的比例",
        "实词数 ÷ 非标点分词数",
        "百分比",
        "0.00%",
    )
    builder.add(
        "实虚词比",
        safe_ratio(real_count, function_count),
        CATEGORY_LEXICAL_DENSITY,
        "实词数量与虚词数量之比",
        "实词数 ÷ 虚词数",
        "比值",
        "0.0000",
    )

    lengths = [len(token.word) for token in non_punct_tokens]
    builder.add(
        "平均词长",
        statistics.fmean(lengths) if lengths else 0.0,
        CATEGORY_LEXICAL_DENSITY,
        "非标点token的平均Unicode字符长度",
        "非标点token字符长度之和 ÷ 非标点分词数",
        "字符/token",
        "0.00",
    )
    han_token_lengths = [len(token.word) for token in non_punct_tokens if HAN_TOKEN_RE.fullmatch(token.word)]
    han_token_count = len(han_token_lengths)
    builder.add_count(
        "汉字词分词数",
        han_token_count,
        CATEGORY_LEXICAL_DENSITY,
        "词形完全由汉字组成的非标点token数量，作为音节词长分布的分母",
    )
    length_buckets = (
        ("单音节词", sum(length == 1 for length in han_token_lengths)),
        ("双音节词", sum(length == 2 for length in han_token_lengths)),
        ("三音节及以上词", sum(length >= 3 for length in han_token_lengths)),
    )
    for label, count in length_buckets:
        builder.add_count(
            f"{label}数",
            count,
            CATEGORY_LEXICAL_DENSITY,
            f"按一个汉字对应一个音节近似，属于{label}的全汉字非标点token数",
        )
        builder.add(
            f"{label}占比",
            safe_ratio(count, han_token_count),
            CATEGORY_LEXICAL_DENSITY,
            f"{label}占全部汉字词token的比例",
            f"{label}数 ÷ 汉字词分词数",
            "百分比",
            "0.00%",
        )


def _add_structure(
    builder: FeatureBuilder,
    text: str,
    tokens: Sequence[Token],
    long_sentence_threshold: int,
) -> None:
    paragraphs = [line for line in text.splitlines() if line.strip()]
    sentence_parts = [part for part in SENTENCE_BOUNDARY_RE.split(text) if part.strip() and HAN_RE.search(part)]
    sentence_char_lengths = [len(HAN_RE.findall(part)) for part in sentence_parts]
    sentence_word_lengths = sentence_token_lengths(tokens)
    clause_parts = [part for part in CLAUSE_BOUNDARY_RE.split(text) if part.strip() and HAN_RE.search(part)]
    sentence_count = len(sentence_char_lengths)
    paragraph_count = len(paragraphs)
    long_sentence_count = sum(length > long_sentence_threshold for length in sentence_char_lengths)
    comma_count = text.count("，") + text.count(",")
    clause_count = len(clause_parts)

    builder.add_count("句子数", sentence_count, CATEGORY_STRUCTURE, "按句号、问号和感叹号识别的句子数量")
    builder.add_count("段落数", paragraph_count, CATEGORY_STRUCTURE, "非空文本行数量")

    def add_length_stats(label: str, lengths: Sequence[int], unit: str) -> None:
        builder.add(
            f"平均句长_{label}",
            statistics.fmean(lengths) if lengths else 0.0,
            CATEGORY_STRUCTURE,
            f"每个句子的平均{unit}长度",
            f"句子{unit}长度之和 ÷ 句子数",
            unit,
            "0.00",
        )
        builder.add(
            f"句长中位数_{label}",
            statistics.median(lengths) if lengths else 0.0,
            CATEGORY_STRUCTURE,
            f"句子{unit}长度的中位数",
            "句长序列中位数",
            unit,
            "0.00",
        )
        builder.add(
            f"句长标准差_{label}",
            statistics.pstdev(lengths) if len(lengths) > 1 else 0.0,
            CATEGORY_STRUCTURE,
            f"句子{unit}长度的总体标准差",
            "句长序列总体标准差",
            unit,
            "0.00",
        )
        builder.add(
            f"最长句长度_{label}",
            max(lengths, default=0),
            CATEGORY_STRUCTURE,
            f"最长句子的{unit}长度",
            "句长序列最大值",
            unit,
            "0",
        )

    add_length_stats("字", sentence_char_lengths, "汉字")
    add_length_stats("词", sentence_word_lengths, "token")
    builder.add_count(
        f"超过{long_sentence_threshold}字长句数",
        long_sentence_count,
        CATEGORY_STRUCTURE,
        f"汉字数大于{long_sentence_threshold}的句子数量",
    )
    builder.add(
        f"超过{long_sentence_threshold}字长句占比",
        safe_ratio(long_sentence_count, sentence_count),
        CATEGORY_STRUCTURE,
        f"超过{long_sentence_threshold}字长句占全部句子的比例",
        f"超过{long_sentence_threshold}字长句数 ÷ 句子数",
        "百分比",
        "0.00%",
    )
    paragraph_lengths = [len(HAN_RE.findall(paragraph)) for paragraph in paragraphs]
    builder.add(
        "平均段落长度_字",
        statistics.fmean(paragraph_lengths) if paragraph_lengths else 0.0,
        CATEGORY_STRUCTURE,
        "每个非空段落的平均汉字数",
        "段落汉字数之和 ÷ 段落数",
        "汉字",
        "0.00",
    )
    paragraph_word_counts = Counter(
        token.paragraph_index for token in tokens if not is_punctuation(token)
    )
    paragraph_word_lengths = [paragraph_word_counts.get(index, 0) for index in range(paragraph_count)]
    builder.add(
        "平均段落长度_词",
        statistics.fmean(paragraph_word_lengths) if paragraph_word_lengths else 0.0,
        CATEGORY_STRUCTURE,
        "每个非空段落的平均非标点token数",
        "各段非标点token数之和 ÷ 段落数",
        "token",
        "0.00",
    )
    builder.add_count("逗号数", comma_count, CATEGORY_STRUCTURE, "中文或英文逗号数量")
    builder.add_count("分句数", clause_count, CATEGORY_STRUCTURE, "以逗号、分号、冒号及句末标点划分的非空分句数量")
    builder.add(
        "平均每句逗号数",
        safe_ratio(comma_count, sentence_count),
        CATEGORY_STRUCTURE,
        "平均每个句子包含的逗号数量",
        "逗号数 ÷ 句子数",
        "个/句",
        "0.00",
    )
    punct_count = sum(is_punctuation(token) for token in tokens)
    builder.add(
        "标点占分词比例",
        safe_ratio(punct_count, len(tokens)),
        CATEGORY_STRUCTURE,
        "标点token占全部PyNLPIR token的比例",
        "标点数 ÷ 分词数",
        "百分比",
        "0.00%",
    )
    builder.add(
        "平均每句分句数",
        safe_ratio(clause_count, sentence_count),
        CATEGORY_STRUCTURE,
        "平均每个句子包含的分句数量",
        "分句数 ÷ 句子数",
        "个/句",
        "0.00",
    )


def _covered_token_indices(
    lexicon_spans: dict[str, list[LexiconMatch]],
    feature_names: Sequence[str],
) -> set[int]:
    return {
        index
        for feature_name in feature_names
        for match in lexicon_spans.get(feature_name, [])
        for index in range(match.start, match.end)
    }


def _sentence_class_counts(tokens: Sequence[Token]) -> tuple[int, int, int, int, int]:
    special_words = {
        "谁", "什么", "哪", "哪个", "哪些", "哪里", "哪儿", "怎么", "怎样", "怎么样",
        "为什么", "为何", "如何", "何时", "多少", "几", "多大", "多长", "多久",
    }
    yes_no_patterns = ("是不是", "有没有", "能不能", "可不可以", "是否")
    special_count = 0
    yes_no_count = 0
    exclamation_count = 0
    ba_count = 0
    bei_count = 0
    for start, end in sentence_token_ranges(tokens):
        sentence = tokens[start:end]
        sentence_text = "".join(token.word for token in sentence)
        is_question = any(token.raw_pos == "ww" or token.word in {"？", "?"} for token in sentence)
        if is_question:
            if any(token.word in special_words for token in sentence):
                special_count += 1
            elif any(pattern in sentence_text for pattern in yes_no_patterns) or any(
                token.word in {"吗", "呢", "吧"} and token.raw_pos.startswith("y") for token in sentence
            ):
                yes_no_count += 1
        if any(token.raw_pos == "wt" or token.word in {"！", "!"} for token in sentence):
            exclamation_count += 1
        if any(token.raw_pos.startswith("pba") for token in sentence):
            ba_count += 1
        if any(token.raw_pos.startswith("pbei") for token in sentence):
            bei_count += 1
    return special_count, yes_no_count, exclamation_count, ba_count, bei_count


def _add_grammar(
    builder: FeatureBuilder,
    tokens: Sequence[Token],
    lexicon_counts: Counter[str],
    lexicon_spans: dict[str, list[LexiconMatch]],
) -> dict[str, int]:
    grammar_counts: dict[str, int] = {}
    for feature_name in GRAMMAR_LEXICON_FEATURES:
        count = lexicon_counts.get(feature_name, 0)
        grammar_counts[feature_name] = count
        builder.add_count(f"{feature_name}数", count, CATEGORY_GRAMMAR, f"由可审查词表识别的{feature_name}数量")

    for feature_name in PAPER_GRAMMAR_LEXICON_FEATURES:
        count = lexicon_counts.get(feature_name, 0)
        builder.add_count(f"{feature_name}数", count, CATEGORY_GRAMMAR, f"由论文补充词表识别的{feature_name}数量")

    raw_features = (
        ("时态助词了", "ule"),
        ("时态助词着", "uzhe"),
        ("时态助词过", "uguo"),
        ("结构助词的", "ude1"),
        ("结构助词地", "ude2"),
        ("结构助词得", "ude3"),
    )
    for label, raw_pos in raw_features:
        count = sum(token.raw_pos == raw_pos for token in tokens)
        builder.add_count(f"{label}数", count, CATEGORY_GRAMMAR, f"PyNLPIR细粒度词性{raw_pos}识别的{label}数量")

    for word in ("把", "被", "对", "给", "从", "向"):
        count = sum(token.word == word and token.raw_pos.startswith("p") for token in tokens)
        builder.add_count(f"介词{word}数", count, CATEGORY_GRAMMAR, f"词形为“{word}”且词性为介词的数量")
    for word in ("吧", "呢", "吗", "啊"):
        count = sum(token.word == word and token.raw_pos.startswith("y") for token in tokens)
        builder.add_count(f"语气词{word}数", count, CATEGORY_GRAMMAR, f"词形为“{word}”且词性为语气词的数量")

    pronoun_features = (
        "第一人称单数代词", "第一人称复数代词", "其他第一人称代词", "第二人称代词",
        "第三人称代词", "非人第三人称代词", "指示代词", "不定代词", "疑问代词",
    )
    covered_pronouns = _covered_token_indices(lexicon_spans, pronoun_features)
    other_pronoun_count = sum(
        raw_pos_root(token.raw_pos) == "r" and index not in covered_pronouns
        for index, token in enumerate(tokens)
    )
    builder.add_count("其他代词数", other_pronoun_count, CATEGORY_GRAMMAR, "未归入既有人称、指示、不定或疑问类别的代词token数")

    tracked_particles = {"ule", "uzhe", "uguo", "ude1", "ude2", "ude3"}
    other_particle_count = sum(
        raw_pos_root(token.raw_pos) == "u" and token.raw_pos not in tracked_particles
        for token in tokens
    )
    builder.add_count("其他助词数", other_particle_count, CATEGORY_GRAMMAR, "除的、地、得、了、着、过之外的u*助词token数")

    adjective_adverbial_count = sum(
        raw_pos_root(tokens[index].raw_pos) == "a"
        and index + 1 < len(tokens)
        and tokens[index].paragraph_index == tokens[index + 1].paragraph_index
        and tokens[index + 1].raw_pos.startswith("v")
        for index in range(len(tokens))
    )
    builder.add_count(
        "形容词作状语数",
        adjective_adverbial_count,
        CATEGORY_GRAMMAR,
        "启发式识别性质形容词直接位于动词前且中间无“地”或标点的次数",
    )

    special_questions, yes_no_questions, exclamations, ba_sentences, bei_sentences = _sentence_class_counts(tokens)
    builder.add_count("特殊疑问句数", special_questions, CATEGORY_GRAMMAR, "以问号结尾且含特殊疑问词的句子数")
    builder.add_count("是非问句数", yes_no_questions, CATEGORY_GRAMMAR, "以问号结尾且含吗、呢、是否或正反问形式的句子数")
    builder.add_count("感叹句数", exclamations, CATEGORY_GRAMMAR, "以感叹号结束的句子数")
    builder.add_count("把字句数", ba_sentences, CATEGORY_GRAMMAR, "含PyNLPIR细粒度词性pba的句子数")
    builder.add_count("被字句数", bei_sentences, CATEGORY_GRAMMAR, "含PyNLPIR细粒度词性pbei的句子数")

    personal_total = sum(
        grammar_counts.get(feature, 0)
        for feature in (
            "第一人称单数代词",
            "第一人称复数代词",
            "其他第一人称代词",
            "第二人称代词",
            "第三人称代词",
            "非人第三人称代词",
        )
    )
    builder.add_count("人称代词总数", personal_total, CATEGORY_GRAMMAR, "第一、第二及第三人称代词数量之和")
    grammar_counts["人称代词总数"] = personal_total
    return grammar_counts


def _add_idioms(
    builder: FeatureBuilder,
    tokens: Sequence[Token],
    idiom_lexicon: Sequence[CompiledIdiomEntry],
) -> None:
    counts, terms = match_idiom_features(tokens, idiom_lexicon)
    total = 0
    unique_total = 0
    for idiom_type in IDIOM_TYPES:
        count = counts.get(idiom_type, 0)
        total += count
        unique_total += len(terms.get(idiom_type, set()))
        source = "熟语词表匹配" if idiom_type != "惯用语" else "熟语词表及PyNLPIR l/nl/vl/al/bl标签识别"
        builder.add_count(f"{idiom_type}数", count, CATEGORY_IDIOM, f"由{source}的{idiom_type}出现次数")
    builder.add_count("熟语数", total, CATEGORY_IDIOM, "成语、歇后语、惯用语和谚语次数之和")
    builder.add_count("熟语去重数", unique_total, CATEGORY_IDIOM, "四类熟语实际使用的不同词项数量")
    builder.add(
        "熟语多样性",
        safe_ratio(unique_total, total),
        CATEGORY_IDIOM,
        "熟语去重词项数与熟语总次数之比",
        "熟语去重数 ÷ 熟语数",
        "比例",
        "0.0000",
    )


def _add_cohesion(
    builder: FeatureBuilder,
    tokens: Sequence[Token],
    lexicon_spans: dict[str, list[LexiconMatch]],
) -> None:
    sentence_ranges = sentence_token_ranges(tokens)
    sentence_by_token: dict[int, int] = {}
    for sentence_id, (start, end) in enumerate(sentence_ranges):
        sentence_by_token.update({index: sentence_id for index in range(start, end)})

    relation_sentences: dict[str, set[int]] = {}
    all_sentences: set[int] = set()
    for relation in COMPOUND_SENTENCE_FEATURES:
        marker_feature = COMPOUND_MARKER_BY_FEATURE[relation]
        sentence_ids = {
            sentence_by_token[match.start]
            for match in lexicon_spans.get(marker_feature, [])
            if match.start in sentence_by_token
        }
        relation_sentences[relation] = sentence_ids
        all_sentences.update(sentence_ids)
        builder.add_count(
            f"{relation}数",
            len(sentence_ids),
            CATEGORY_COHESION,
            f"含{relation}关系标记的句子数；同一句同类标记只计一次",
        )
    relation_assignment_total = sum(len(sentence_ids) for sentence_ids in relation_sentences.values())
    relation_type_count = sum(bool(sentence_ids) for sentence_ids in relation_sentences.values())
    builder.add_count("复句句次总数", relation_assignment_total, CATEGORY_COHESION, "九类复句句数之和；同一句可同时归入多类")
    builder.add_count("含关系标记复句数", len(all_sentences), CATEGORY_COHESION, "至少命中一类复句关系标记的不同句子数")
    builder.add_count("复句类型数", relation_type_count, CATEGORY_COHESION, "正文实际出现的不同复句关系类型数量", add_per_thousand=False)
    builder.add(
        "复句类型多样性",
        safe_ratio(relation_type_count, len(COMPOUND_SENTENCE_FEATURES)),
        CATEGORY_COHESION,
        "实际出现的复句关系类型占预设九类关系的比例",
        "复句类型数 ÷ 9",
        "比例",
        "0.0000",
    )


def _add_narrative(
    builder: FeatureBuilder,
    text: str,
    tokens: Sequence[Token],
    lexicon_counts: Counter[str],
    lexicon_spans: dict[str, list[LexiconMatch]],
    grammar_counts: dict[str, int],
) -> None:
    for feature_name in NARRATIVE_LEXICON_FEATURES:
        count = lexicon_counts.get(feature_name, 0)
        builder.add_count(f"{feature_name}数", count, CATEGORY_NARRATIVE, f"由可审查词表识别的{feature_name}数量")

    evaluation_count = lexicon_counts.get("正面评价词", 0) + lexicon_counts.get("负面评价词", 0)
    builder.add_count("评价词总数", evaluation_count, CATEGORY_NARRATIVE, "正面与负面评价词数量之和")
    locative_count = sum(raw_pos_root(token.raw_pos) in {"s", "f"} for token in tokens)
    directional_count = sum(token.raw_pos == "vf" for token in tokens)
    builder.add_count("处所方位词数", locative_count, CATEGORY_NARRATIVE, "PyNLPIR处所词与方位词数量之和")
    builder.add_count("趋向动词数", directional_count, CATEGORY_NARRATIVE, "PyNLPIR细粒度词性vf识别的趋向动词数量")

    for feature_name in TIME_NOUN_LEXICON_FEATURES:
        builder.add_count(f"{feature_name}数", lexicon_counts.get(feature_name, 0), CATEGORY_NARRATIVE, f"由论文补充词表识别的{feature_name}数量")
    covered_time_nouns = _covered_token_indices(lexicon_spans, TIME_NOUN_LEXICON_FEATURES)
    other_time_nouns = sum(
        raw_pos_root(token.raw_pos) == "t" and index not in covered_time_nouns
        for index, token in enumerate(tokens)
    )
    builder.add_count("其他时间名词数", other_time_nouns, CATEGORY_NARRATIVE, "未归入过去、现在或未来类别的t*时间名词token数")

    for feature_name in TIME_ADVERB_LEXICON_FEATURES:
        builder.add_count(f"{feature_name}数", lexicon_counts.get(feature_name, 0), CATEGORY_NARRATIVE, f"由论文补充词表识别的{feature_name}数量")
    covered_adverbs = _covered_token_indices(
        lexicon_spans,
        (
            "疑问副词", "低调词与模糊语", "夸张与加强语", "时间副词", "程度副词",
            *TIME_ADVERB_LEXICON_FEATURES,
        ),
    )
    other_adverbs = sum(
        raw_pos_root(token.raw_pos) == "d" and index not in covered_adverbs
        for index, token in enumerate(tokens)
    )
    builder.add_count("其他副词数", other_adverbs, CATEGORY_NARRATIVE, "未归入疑问、时间、程度、模糊或加强类别的d*副词token数")

    for feature_name in VERB_CLASS_LEXICON_FEATURES:
        builder.add_count(f"{feature_name}数", lexicon_counts.get(feature_name, 0), CATEGORY_NARRATIVE, f"由论文补充词表识别的{feature_name}数量")
    verb_is_count = sum(token.raw_pos == "vshi" or (token.word == "是" and token.raw_pos.startswith("v")) for token in tokens)
    builder.add_count("动词是数", verb_is_count, CATEGORY_NARRATIVE, "PyNLPIR细粒度词性vshi或动词词形“是”的数量")
    covered_verbs = _covered_token_indices(
        lexicon_spans,
        ("情态动词", *PAPER_GRAMMAR_LEXICON_FEATURES[-4:], *VERB_CLASS_LEXICON_FEATURES),
    )
    other_verbs = sum(
        token.raw_pos.startswith("v")
        and token.raw_pos not in {"vf", "vshi"}
        and index not in covered_verbs
        for index, token in enumerate(tokens)
    )
    builder.add_count("其他动词数", other_verbs, CATEGORY_NARRATIVE, "未归入趋向、是、情态、私人、建议要求或公共类别的v*动词token数")

    quote_count = count_direct_quotes(text)
    builder.add_count("直接引语数", quote_count, CATEGORY_NARRATIVE, "成对中文或英文双引号标记的直接引语片段数量")
    first_person = sum(
        grammar_counts.get(feature, 0)
        for feature in ("第一人称单数代词", "第一人称复数代词", "其他第一人称代词")
    )
    builder.add(
        "第一人称代词占人称代词比例",
        safe_ratio(first_person, grammar_counts.get("人称代词总数", 0)),
        CATEGORY_NARRATIVE,
        "第一人称代词在人称代词中的比例",
        "第一人称代词数 ÷ 人称代词总数",
        "百分比",
        "0.00%",
    )
    run_count, longest_run = count_verb_runs(tokens)
    builder.add_count("连续动词结构数", run_count, CATEGORY_NARRATIVE, "同一句内长度至少为2的连续v*词性序列数量")
    builder.add(
        "最长连续动词序列",
        longest_run,
        CATEGORY_NARRATIVE,
        "同一句内最长连续v*词性序列长度",
        "连续动词序列长度最大值",
        "token",
        "0",
    )


def _non_hsk_partition(token: Token) -> str:
    if token.raw_pos.startswith(PROPER_NOUN_PREFIXES):
        return "专名"
    if token.raw_pos.startswith("m") or DIGIT_RE.search(token.word):
        return "数字"
    if token.raw_pos.startswith("x") or ASCII_LETTER_RE.search(token.word):
        return "字母串"
    return "其他"


def _add_hsk(
    builder: FeatureBuilder,
    non_punct_tokens: Sequence[Token],
    hsk_levels: Sequence[str],
    hsk_groups: dict[str, Sequence[str]],
) -> None:
    level_tokens: dict[str, list[Token]] = {level: [] for level in hsk_levels}
    non_hsk_tokens: list[Token] = []
    for token in non_punct_tokens:
        if token.hsk_level in level_tokens:
            level_tokens[token.hsk_level].append(token)
        else:
            non_hsk_tokens.append(token)

    non_punct_count = len(non_punct_tokens)
    hsk_count = sum(len(tokens) for tokens in level_tokens.values())
    for level in hsk_levels:
        tokens = level_tokens[level]
        count = len(tokens)
        type_count = len({token.word for token in tokens})
        builder.add_count(
            f"{level}级词汇次数",
            count,
            CATEGORY_HSK,
            f"主等级为{level}的HSK词汇token次数",
        )
        builder.add(
            f"{level}级词汇占比",
            safe_ratio(count, non_punct_count),
            CATEGORY_HSK,
            f"{level}级HSK词汇占全部非标点token的比例",
            f"{level}级词汇次数 ÷ 非标点分词数",
            "百分比",
            "0.00%",
        )
        builder.add_count(f"{level}级词汇种类数", type_count, CATEGORY_HSK, f"实际使用的{level}级HSK不同词形数量")
        builder.add(
            f"{level}级词汇内部TTR",
            safe_ratio(type_count, count),
            CATEGORY_HSK,
            f"{level}级HSK词汇内部的类型符号比",
            f"{level}级词汇种类数 ÷ {level}级词汇次数",
            "比例",
            "0.0000",
        )

    for group, levels in hsk_groups.items():
        tokens = [token for level in levels for token in level_tokens[level]]
        count = len(tokens)
        type_count = len({token.word for token in tokens})
        builder.add_count(
            f"{group}词汇次数",
            count,
            CATEGORY_HSK,
            f"HSK{group}等级组词汇token次数",
        )
        builder.add(
            f"{group}词汇占比",
            safe_ratio(count, non_punct_count),
            CATEGORY_HSK,
            f"HSK{group}等级组词汇占全部非标点token的比例",
            f"{group}词汇次数 ÷ 非标点分词数",
            "百分比",
            "0.00%",
        )
        builder.add_count(f"{group}词汇种类数", type_count, CATEGORY_HSK, f"实际使用的HSK{group}等级组不同词形数量")
        builder.add(
            f"{group}词汇内部TTR",
            safe_ratio(type_count, count),
            CATEGORY_HSK,
            f"HSK{group}等级组词汇内部的类型符号比",
            f"{group}词汇种类数 ÷ {group}词汇次数",
            "比例",
            "0.0000",
        )

    builder.add_count(
        "HSK词汇次数",
        hsk_count,
        CATEGORY_HSK,
        "命中新版HSK词表的非标点token数量",
    )
    builder.add(
        "HSK词汇覆盖率",
        safe_ratio(hsk_count, non_punct_count),
        CATEGORY_HSK,
        "HSK词汇占全部非标点token的比例",
        "HSK词汇次数 ÷ 非标点分词数",
        "百分比",
        "0.00%",
    )
    non_hsk_count = len(non_hsk_tokens)
    non_hsk_types = len({token.word for token in non_hsk_tokens})
    builder.add_count(
        "非HSK词汇次数",
        non_hsk_count,
        CATEGORY_HSK,
        "未命中新版HSK词表的非标点token数量",
    )
    builder.add(
        "非HSK词汇占比",
        safe_ratio(non_hsk_count, non_punct_count),
        CATEGORY_HSK,
        "非HSK词汇占全部非标点token的比例",
        "非HSK词汇次数 ÷ 非标点分词数",
        "百分比",
        "0.00%",
    )
    builder.add_count("非HSK词汇种类数", non_hsk_types, CATEGORY_HSK, "未命中HSK词表的不同词形数量")
    builder.add(
        "非HSK词汇TTR",
        safe_ratio(non_hsk_types, non_hsk_count),
        CATEGORY_HSK,
        "非HSK词汇内部的类型符号比",
        "非HSK词汇种类数 ÷ 非HSK词汇次数",
        "比例",
        "0.0000",
    )
    high_levels = hsk_groups.get("高等", ())
    high_count = sum(len(level_tokens[level]) for level in high_levels)
    builder.add(
        "高等词汇占HSK词汇比例",
        safe_ratio(high_count, hsk_count),
        CATEGORY_HSK,
        "高等HSK词汇在全部HSK词汇中的比例",
        "高等词汇次数 ÷ HSK词汇次数",
        "百分比",
        "0.00%",
    )

    partitions = Counter(_non_hsk_partition(token) for token in non_hsk_tokens)
    for label in ("专名", "数字", "字母串", "其他"):
        builder.add_count(
            f"非HSK{label}数",
            partitions.get(label, 0),
            CATEGORY_HSK,
            f"非HSK词汇中归为{label}的token数量",
        )


def compute_linguistic_features(
    text: str,
    tokens: Sequence[Token],
    lexicon: Sequence[CompiledLexiconEntry],
    *,
    mattr_window: int,
    long_sentence_threshold: int,
    hsk_levels: Sequence[str],
    hsk_groups: dict[str, Sequence[str]],
    idiom_lexicon: Sequence[CompiledIdiomEntry] = (),
) -> FeatureResult:
    if mattr_window <= 0:
        raise ValueError("MATTR 窗口必须大于 0")
    if long_sentence_threshold <= 0:
        raise ValueError("长句阈值必须大于 0")

    han_char_count = len(HAN_RE.findall(re.sub(r"\s+", "", text)))
    non_punct_tokens = [token for token in tokens if not is_punctuation(token)]
    builder = FeatureBuilder(han_char_count)
    lexicon_spans = match_lexicon_feature_spans(tokens, lexicon)
    lexicon_counts = Counter({feature: len(matches) for feature, matches in lexicon_spans.items()})

    _add_lexical_diversity(builder, non_punct_tokens, mattr_window)
    _add_lexical_density_and_length(builder, non_punct_tokens)
    _add_structure(builder, text, tokens, long_sentence_threshold)
    _add_idioms(builder, tokens, idiom_lexicon)
    grammar_counts = _add_grammar(builder, tokens, lexicon_counts, lexicon_spans)
    _add_cohesion(builder, tokens, lexicon_spans)
    _add_narrative(builder, text, tokens, lexicon_counts, lexicon_spans, grammar_counts)
    _add_hsk(builder, non_punct_tokens, hsk_levels, hsk_groups)
    return builder.result()
