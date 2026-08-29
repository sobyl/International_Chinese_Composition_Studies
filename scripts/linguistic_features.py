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
CATEGORY_COHESION = "篇章连接"
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

CONNECTIVE_FEATURES = (
    "因果连接词",
    "转折连接词",
    "条件连接词",
    "递进连接词",
    "并列连接词",
    "顺序连接词",
    "总结连接词",
    "举例连接词",
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
    (*GRAMMAR_LEXICON_FEATURES, *CONNECTIVE_FEATURES, *NARRATIVE_LEXICON_FEATURES)
)

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
    return RAW_POS_PARENT.get(normalized[:1], "") if normalized else ""


def raw_pos_root(raw_pos: str | None) -> str:
    normalized = (raw_pos or "").strip().lower()
    return normalized[:1] if normalized else ""


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

    missing_features = sorted(LEXICON_FEATURES.difference(spec.feature_name for spec in specs))
    if missing_features:
        raise ValueError(f"语言特征词表缺少特征：{missing_features}")
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


def match_lexicon_features(
    tokens: Sequence[Token],
    entries: Sequence[CompiledLexiconEntry],
) -> tuple[Counter[str], dict[str, set[str]]]:
    entries_by_feature: dict[str, list[CompiledLexiconEntry]] = defaultdict(list)
    for entry in entries:
        entries_by_feature[entry.feature_name].append(entry)
    for feature_entries in entries_by_feature.values():
        feature_entries.sort(key=lambda entry: (-len(entry.words), entry.term))

    counts: Counter[str] = Counter()
    matched_terms: dict[str, set[str]] = defaultdict(set)
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
            counts[feature_name] += 1
            matched_terms[feature_name].add(matched.term)
            index += len(matched.words)
    return counts, matched_terms


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

    pos_groups = (
        ("名词", "noun"),
        ("动词", "verb"),
        ("形容词", "adjective"),
        ("副词", "adverb"),
    )
    for label, parent_pos in pos_groups:
        pos_words = [token.word for token in non_punct_tokens if token.parent_pos == parent_pos]
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
    length_buckets = (
        ("单字词", sum(length == 1 for length in lengths)),
        ("双字词", sum(length == 2 for length in lengths)),
        ("三字及以上词", sum(length >= 3 for length in lengths)),
    )
    for label, count in length_buckets:
        builder.add_count(f"{label}数", count, CATEGORY_LEXICAL_DENSITY, f"长度属于{label}的非标点token数")
        builder.add(
            f"{label}占比",
            safe_ratio(count, len(non_punct_tokens)),
            CATEGORY_LEXICAL_DENSITY,
            f"{label}占全部非标点token的比例",
            f"{label}数 ÷ 非标点分词数",
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
    builder.add(
        "平均每句分句数",
        safe_ratio(clause_count, sentence_count),
        CATEGORY_STRUCTURE,
        "平均每个句子包含的分句数量",
        "分句数 ÷ 句子数",
        "个/句",
        "0.00",
    )


def _add_grammar(
    builder: FeatureBuilder,
    tokens: Sequence[Token],
    lexicon_counts: Counter[str],
) -> dict[str, int]:
    grammar_counts: dict[str, int] = {}
    for feature_name in GRAMMAR_LEXICON_FEATURES:
        count = lexicon_counts.get(feature_name, 0)
        grammar_counts[feature_name] = count
        builder.add_count(f"{feature_name}数", count, CATEGORY_GRAMMAR, f"由可审查词表识别的{feature_name}数量")

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


def _add_cohesion(
    builder: FeatureBuilder,
    lexicon_counts: Counter[str],
    matched_terms: dict[str, set[str]],
) -> None:
    total = 0
    unique_terms: set[str] = set()
    for feature_name in CONNECTIVE_FEATURES:
        count = lexicon_counts.get(feature_name, 0)
        total += count
        unique_terms.update(matched_terms.get(feature_name, set()))
        builder.add_count(f"{feature_name}数", count, CATEGORY_COHESION, f"由可审查词表识别的{feature_name}数量")
    builder.add_count("连接词总数", total, CATEGORY_COHESION, "八类连接标记数量之和")
    builder.add_count("连接词去重数", len(unique_terms), CATEGORY_COHESION, "实际使用过的不同连接词项数量")
    builder.add(
        "连接词多样性",
        safe_ratio(len(unique_terms), total),
        CATEGORY_COHESION,
        "不同连接词项数量与连接词总次数之比",
        "连接词去重数 ÷ 连接词总数",
        "比例",
        "0.0000",
    )


def _add_narrative(
    builder: FeatureBuilder,
    text: str,
    tokens: Sequence[Token],
    lexicon_counts: Counter[str],
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
) -> FeatureResult:
    if mattr_window <= 0:
        raise ValueError("MATTR 窗口必须大于 0")
    if long_sentence_threshold <= 0:
        raise ValueError("长句阈值必须大于 0")

    han_char_count = len(HAN_RE.findall(re.sub(r"\s+", "", text)))
    non_punct_tokens = [token for token in tokens if not is_punctuation(token)]
    builder = FeatureBuilder(han_char_count)
    lexicon_counts, matched_terms = match_lexicon_features(tokens, lexicon)

    _add_lexical_diversity(builder, non_punct_tokens, mattr_window)
    _add_lexical_density_and_length(builder, non_punct_tokens)
    _add_structure(builder, text, tokens, long_sentence_threshold)
    grammar_counts = _add_grammar(builder, tokens, lexicon_counts)
    _add_cohesion(builder, lexicon_counts, matched_terms)
    _add_narrative(builder, text, tokens, lexicon_counts, grammar_counts)
    _add_hsk(builder, non_punct_tokens, hsk_levels, hsk_groups)
    return builder.result()
