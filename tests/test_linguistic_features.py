from __future__ import annotations

import math
import unittest

from linguistic_features import (
    CompiledLexiconEntry,
    Token,
    compute_linguistic_features,
    count_verb_runs,
    match_lexicon_features,
    mattr,
    raw_pos_parent,
    sentence_token_lengths,
)


def token(
    word: str,
    raw_pos: str,
    *,
    paragraph: int = 0,
    hsk_level: str | None = None,
) -> Token:
    return Token(
        word=word,
        raw_pos=raw_pos,
        parent_pos=raw_pos_parent(raw_pos),
        paragraph_index=paragraph,
        hsk_level=hsk_level,
    )


class RawPosTests(unittest.TestCase):
    def test_raw_pos_maps_to_existing_parent_categories(self) -> None:
        expected = {
            "ule": "particle",
            "uzhe": "particle",
            "ude1": "particle",
            "pba": "preposition",
            "pbei": "preposition",
            "vf": "verb",
            "nr": "noun",
            "ns": "noun",
            "Mg": "numeral",
            "Rg": "pronoun",
            "wj": "punctuation mark",
        }
        for raw_pos, parent in expected.items():
            with self.subTest(raw_pos=raw_pos):
                self.assertEqual(raw_pos_parent(raw_pos), parent)


class LexicalDiversityTests(unittest.TestCase):
    def test_mattr_uses_full_ttr_for_short_text(self) -> None:
        self.assertEqual(mattr(["我", "爱", "我", "家"], 50), 0.75)

    def test_mattr_averages_sliding_windows(self) -> None:
        self.assertTrue(math.isclose(mattr(["甲", "乙", "甲", "丙"], 3), 5 / 6))

    def test_mattr_and_ratios_handle_empty_text(self) -> None:
        result = compute_linguistic_features(
            "",
            [],
            [],
            mattr_window=50,
            long_sentence_threshold=30,
            hsk_levels=("1", "2", "3", "4", "5", "6", "7-9"),
            hsk_groups={"初等": ("1", "2", "3"), "中等": ("4", "5", "6"), "高等": ("7-9",)},
        )
        self.assertEqual(result.values["词形丰富度TTR"], 0)
        self.assertEqual(result.values["词汇密度"], 0)
        self.assertEqual(result.values["HSK词汇覆盖率"], 0)


class StructureTests(unittest.TestCase):
    def test_sentence_token_lengths_ignore_punctuation(self) -> None:
        tokens = [
            token("我", "rr"),
            token("喜欢", "v"),
            token("，", "wd"),
            token("中文", "n"),
            token("。", "wj"),
            token("你", "rr"),
            token("呢", "y"),
            token("？", "ww"),
        ]
        self.assertEqual(sentence_token_lengths(tokens), [3, 2])

    def test_continuous_verb_runs_are_sentence_bounded(self) -> None:
        tokens = [
            token("走", "v"),
            token("进", "vf"),
            token("学校", "n"),
            token("开始", "v"),
            token("认真", "d"),
            token("学习", "v"),
            token("写", "v"),
            token("作业", "n"),
            token("。", "wj"),
            token("说", "v"),
            token("完", "vf"),
            token("离开", "v"),
        ]
        self.assertEqual(count_verb_runs(tokens), (3, 3))


class LexiconMatcherTests(unittest.TestCase):
    def test_longest_sequence_wins_within_feature(self) -> None:
        entries = (
            CompiledLexiconEntry("篇章连接", "因果连接词", "因", ("因",), ("c",), "test"),
            CompiledLexiconEntry("篇章连接", "因果连接词", "因为", ("因", "为"), ("c",), "test"),
        )
        counts, terms = match_lexicon_features(
            [token("因", "c"), token("为", "c"), token("天气", "n")],
            entries,
        )
        self.assertEqual(counts["因果连接词"], 1)
        self.assertEqual(terms["因果连接词"], {"因为"})


class HskDerivedFeatureTests(unittest.TestCase):
    def test_hsk_types_groups_and_non_hsk_partitions(self) -> None:
        tokens = [
            token("我", "rr", hsk_level="1"),
            token("不", "d", hsk_level="1"),
            token("喜欢", "v", hsk_level="1"),
            token("学习", "v", hsk_level="2"),
            token("。", "wj"),
            token("我们", "rr", paragraph=1, hsk_level="1"),
            token("去", "vf", paragraph=1),
            token("北京", "ns", paragraph=1),
            token("。", "wj", paragraph=1),
        ]
        result = compute_linguistic_features(
            "我不喜欢学习。\n我们去北京。",
            tokens,
            [],
            mattr_window=50,
            long_sentence_threshold=30,
            hsk_levels=("1", "2", "3", "4", "5", "6", "7-9"),
            hsk_groups={"初等": ("1", "2", "3"), "中等": ("4", "5", "6"), "高等": ("7-9",)},
        )
        values = result.values
        self.assertEqual(values["1级词汇次数"], 4)
        self.assertEqual(values["1级词汇种类数"], 4)
        self.assertEqual(values["2级词汇次数"], 1)
        self.assertEqual(values["初等词汇次数"], 5)
        self.assertEqual(values["HSK词汇次数"], 5)
        self.assertEqual(values["非HSK词汇次数"], 2)
        self.assertEqual(values["非HSK专名数"], 1)
        self.assertEqual(values["非HSK其他数"], 1)
        self.assertEqual(values["句子数"], 2)
        self.assertEqual(values["段落数"], 2)


if __name__ == "__main__":
    unittest.main()
