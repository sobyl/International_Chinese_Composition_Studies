from __future__ import annotations

import json
import unittest
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from urllib.error import HTTPError

from scripts.collect_zuowen_native_controls import (
    CandidateAudit,
    ParsedArticle,
    SearchCandidate,
    TOPIC_RULES,
    SlowHttpClient,
    build_eligible_pools,
    candidate_sort_key,
    character_ngrams,
    classify_topic,
    clean_body,
    clean_source_label,
    decode_page,
    deduplicate_across_groups,
    jaccard,
    parse_article,
    parse_library_listing_page,
    parse_search_payload,
    parse_listing_page,
    prune_generated_texts,
    quantile_targets,
    select_group,
    validate_article,
)
from scripts.segment_native_control_texts import (
    align_rows_to_learner_schema,
    load_native_records,
    prune_stale_segmented_texts,
)
from scripts.linguistic_features import FieldSpec


ARTICLE_HTML = """<!doctype html>
<html><head><meta charset="{charset}"><title>难忘的寒假_500字_作文网</title>
<script>var artID = '123456';</script></head><body>
<div class="path"><a>作文</a> &gt; <a>高中作文</a> &gt; <a>高二</a> &gt; <a>叙事作文</a></div>
<h1 class="h_title">难忘的寒假_500字</h1>
<p style="text-align:center;padding:10px">2010-02-03 来源：作文网原创</p>
<div class="con_content"><p>　　寒假里的一天，我们乘车来到郊外。那次经历让我十分难忘。<br/>
　　一路上发生了许多事情，我也从中懂得了坚持的意义。{padding}</p></div>
</body></html>"""

FIXTURES = Path(__file__).parent / "fixtures"


class DecodeAndArticleParserTests(unittest.TestCase):
    def test_decodes_utf8_gb2312_and_gb18030(self) -> None:
        for declared, actual in (("utf-8", "utf-8"), ("gb2312", "gb18030"), ("gbk", "gb18030")):
            source = ARTICLE_HTML.format(charset=declared, padding="成长" * 120)
            text, encoding = decode_page(source.encode(actual))
            self.assertIn("难忘的寒假", text)
            self.assertTrue(encoding.startswith("utf") or encoding == "gb18030")

    def test_parses_and_cleans_article(self) -> None:
        source = ARTICLE_HTML.format(charset="gb2312", padding="成长" * 120).encode("gb18030")
        article = parse_article(source, "https://www.zuowen.com/e/20100203/abc123.shtml")
        self.assertEqual(article.article_id, "123456")
        self.assertEqual(article.grade, "高二")
        self.assertEqual(article.genre_label, "叙事作文")
        self.assertEqual(article.published_date, "2010-02-03")
        self.assertNotIn("难忘的寒假_500字", article.clean_text)
        self.assertGreater(article.han_count, 200)
        valid, reason = validate_article(article, TOPIC_RULES["NJ2"], 200, 750)
        self.assertTrue(valid, reason)

    def test_parses_local_article_fixture(self) -> None:
        source = (FIXTURES / "zuowen_article.html").read_bytes()
        article = parse_article(source, "https://www.zuowen.com/e/20100203/abc123.shtml")
        self.assertEqual(article.article_id, "123456")
        self.assertNotIn("责任编辑", article.clean_text)
        self.assertNotIn("未经允许不得转载", article.clean_text)

    def test_clean_body_removes_boilerplate(self) -> None:
        value = clean_body(
            "标题\n正文第一段。\n作文网专稿 未经允许不得转载\n责任编辑：某某",
            "标题",
        )
        self.assertEqual(value, "正文第一段。")

    def test_clean_body_removes_site_footer_and_tail_author(self) -> None:
        value = clean_body(
            "标题\n正文第一段。\n包头六中高一三班 任雪薇\n"
            "本文系本站用户原创文章，未经允许禁止转载！\n"
            "中学生写作指导、写作素材、优秀作文以及有奖活动\n"
            "尽在“作文网”微信公众号",
            "标题",
        )
        self.assertEqual(value, "正文第一段。")

    def test_clean_source_label_removes_copyright_tail(self) -> None:
        self.assertEqual(
            clean_source_label("作文网原创作文网专稿 未经允许不得转载"),
            "作文网原创",
        )
        self.assertEqual(clean_source_label("转载作文网专稿 未经允许不得转载"), "转载")
        self.assertEqual(
            clean_source_label("网络资源中学生写作指导、写作素材、优秀作文以及有奖活动尽在微信公众号"),
            "网络资源",
        )
        self.assertEqual(clean_source_label("本站原创本文系本站用户原创文章，"), "本站原创")

    def test_high_school_topic_essay_can_validate_as_argumentative(self) -> None:
        text = ("我认为父母应该适时放手，因此孩子才能独立成长。" * 20)
        article = ParsedArticle(
            url="https://www.zuowen.com/e/20100101/abc123.shtml",
            article_id="abc123",
            title="想对父母说",
            breadcrumb="作文 > 高中作文 > 高一 > 话题作文",
            grade="高一",
            genre_label="话题作文",
            published_date="2010-01-01",
            source_label="作文网原创",
            raw_text=text,
            clean_text=text,
            han_count=300,
            text_hash="hash",
        )
        valid, reason = validate_article(article, TOPIC_RULES["NY2"], 200, 750)
        self.assertTrue(valid, reason)

    def test_middle_school_requires_explicit_backup_permission(self) -> None:
        text = "我认为家庭教育会影响孩子成长，因此父母应该言传身教。" * 20
        article = ParsedArticle(
            url="https://www.zuowen.com/e/20100101/abc124.shtml",
            article_id="abc124",
            title="家庭教育",
            breadcrumb="作文 > 初中作文 > 初二 > 议论文",
            grade="初二",
            genre_label="议论文",
            published_date="2010-01-01",
            source_label="作文网原创",
            raw_text=text,
            clean_text=text,
            han_count=300,
            text_hash="hash",
        )
        self.assertFalse(validate_article(article, TOPIC_RULES["NY2"], 200, 750)[0])
        self.assertTrue(
            validate_article(
                article,
                TOPIC_RULES["NY2"],
                200,
                750,
                allow_middle_school=True,
            )[0]
        )

    def test_complete_high_school_unit_essay_is_allowed(self) -> None:
        text = "那天我们参加社会实践，这段难忘的经历让我成长。" * 25
        article = ParsedArticle(
            url="https://www.zuowen.com/e/20100101/abc125.shtml",
            article_id="abc125",
            title="人教版高中单元作文：难忘的一件事",
            breadcrumb="作文 > 单元作文 > 高中语文 > 人教版高中第二册",
            grade="未标注",
            genre_label="",
            published_date="2010-01-01",
            source_label="作文网",
            raw_text=text,
            clean_text=text,
            han_count=300,
            text_hash="hash",
        )
        valid, reason = validate_article(article, TOPIC_RULES["NJ2"], 200, 750)
        self.assertTrue(valid, reason)


class SearchAndMatchingTests(unittest.TestCase):
    def test_preclassified_topic_mismatches_do_not_fetch_articles(self) -> None:
        class NoFetchClient:
            def fetch(self, url: str, force: bool = False) -> bytes:
                raise AssertionError(f"不应请求主题不匹配页面：{url}")

        discovered = {
            code: [
                SearchCandidate(
                    native_code=code,
                    url=f"https://www.zuowen.com/e/20100101/{code.lower()}.shtml",
                    title="与目标无关的作文",
                    search_date="2010-01-01",
                    query="测试",
                    source_label="作文网",
                    nominal_length=500,
                    pre_tier="不匹配",
                    pre_score=0,
                )
            ]
            for code in TOPIC_RULES
        }
        pools, audit = build_eligible_pools(
            NoFetchClient(),
            discovered,
            manual_reviews={},
            target_per_group=1,
            pool_multiplier=1,
            preferred_start=2005,
            preferred_end=2012,
            min_han=200,
            max_han=750,
            force_fetch=False,
        )
        self.assertTrue(all(not values for values in pools.values()))
        self.assertEqual(len(audit), 4)
        self.assertTrue(all("未请求正文" in row.reason for row in audit))

    def test_listing_page_discovers_topic_candidate(self) -> None:
        page = (FIXTURES / "zuowen_listing.html").read_bytes()
        rows = parse_listing_page(page, ("NY1", "NY2"), "https://www.zuowen.com/gaozhong/gaoer/yilunwen/")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].native_code, "NY1")
        self.assertTrue(rows[0].verified_high_school_hint)

    def test_library_listing_preserves_verified_grade_and_genre(self) -> None:
        page = """
        <div class="list"><div class="list_left"><div class="author">
        <span class="date">投稿时间：2011-05-06</span>
        <b><a href="/e/20110506/abc123.shtml">父母的教育_500字</a></b></div>
        <p class="article">父母的教育影响孩子成长。父母应该言传身教，因此家庭教育十分重要。</p>
        <div class="tags"><a href="/wk/gz-11-0-0-0-0.html">高二</a>
        <a href="/wk/gz-0-9-0-0-0.html">议论文</a></div></div>
        <div class="clear"></div></div>
        """.encode()
        rows = parse_library_listing_page(page, ("NY2",), "https://www.zuowen.com/wk/test.html")
        self.assertEqual(len(rows), 1)
        self.assertTrue(rows[0].library_verified)
        self.assertEqual(rows[0].verified_grade, "高二")
        self.assertEqual(rows[0].verified_genre, "议论文")

    def test_search_payload_keeps_only_zuowen_articles(self) -> None:
        payload = {
            "code": 200,
            "data": {
                "items": [
                    {
                        "data_type": "article",
                        "target_url": "http://www.zuowen.com/e/20100528/abc123.shtml",
                        "title": "<mark>吸烟</mark>的危害_500字",
                        "content": "吸烟危害个人健康和公共利益。",
                        "ctime": "2010-05-28",
                        "copy_from": "作文网原创",
                    },
                    {
                        "data_type": "article",
                        "target_url": "http://www.example.com/e/20100528/abc123.shtml",
                        "title": "吸烟",
                    },
                    {"data_type": "tiku", "target_url": "https://paper.example/test", "title": "广告"},
                ]
            },
        }
        rows = parse_search_payload(json.dumps(payload).encode(), "NY1", "高中 吸烟")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].pre_tier, "精确")
        self.assertEqual(rows[0].nominal_length, 500)

    def test_topic_tiers(self) -> None:
        self.assertEqual(classify_topic("吸烟危害个人健康和公众利益", TOPIC_RULES["NY1"])[0], "精确")
        self.assertEqual(classify_topic("父母言传身教影响孩子成长", TOPIC_RULES["NY2"])[0], "近似")
        public_health = (
            "公共场所的文明责任\n"
            "个人行为会影响他人健康，因此社会成员应该维护公共利益和文明环境。" * 4
        )
        family_influence = (
            "家庭环境影响成长\n"
            "父母与孩子需要沟通，家庭环境和父母陪伴会影响孩子独立成长。" * 4
        )
        self.assertEqual(classify_topic(public_health, TOPIC_RULES["NY1"])[0], "扩展")
        self.assertEqual(classify_topic(family_influence, TOPIC_RULES["NY2"])[0], "近似")
        social_behavior = (
            "自律与责任\n"
            "个人遵守规则、保持自律，既体现道德修养，也会影响他人与社会生活。" * 4
        )
        youth_growth = (
            "在独立中成长\n"
            "青春需要独立与责任，教育应帮助年轻人在成长中理解他人。" * 4
        )
        personal_experience = (
            "一次难忘的经历\n"
            "那天我们参加实践活动，这件事让我难忘，也成为成长中的一段回忆。" * 4
        )
        self.assertEqual(classify_topic(social_behavior, TOPIC_RULES["NY1"])[0], "宽泛")
        self.assertEqual(classify_topic(youth_growth, TOPIC_RULES["NY2"])[0], "宽泛")
        self.assertEqual(classify_topic(personal_experience, TOPIC_RULES["NJ2"])[0], "宽泛")
        self.assertEqual(classify_topic("校园里的梧桐树", TOPIC_RULES["NY2"])[0], "不匹配")
        incidental = "校园生活\n" + "我热爱校园生活。" * 50 + "假期里我曾经旅行一次。"
        self.assertEqual(classify_topic(incidental, TOPIC_RULES["NJ2"])[0], "不匹配")
        incidental_smoking = "远水能解近渴\n" + "人生需要坚持。" * 30 + "近水如烟草般有害。"
        self.assertEqual(classify_topic(incidental_smoking, TOPIC_RULES["NY1"])[0], "不匹配")
        incidental_parenting = "杯聚人生\n" + "大人用故事教育孩子。" + "乐观使人成功。" * 30
        self.assertEqual(classify_topic(incidental_parenting, TOPIC_RULES["NY2"])[0], "不匹配")
        incidental_family = "留一道缝隙\n家庭教育要给子女留一点空间。" + "人生也要留有余地。" * 30
        self.assertEqual(classify_topic(incidental_family, TOPIC_RULES["NY2"])[0], "不匹配")

    def test_quantile_targets_are_ordered(self) -> None:
        targets = quantile_targets([100, 200, 300, 400, 500], 4)
        self.assertEqual(targets, sorted(targets))
        self.assertEqual(len(targets), 4)

    def test_selection_maximizes_high_school_samples_before_backup(self) -> None:
        rows = [
            CandidateAudit(
                native_code="NY2",
                learner_code="Y2",
                url=f"https://www.zuowen.com/e/20100101/{name}.shtml",
                title=name,
                published_date="2010-01-01",
                grade=grade,
                school_stage=stage,
                han_count=length,
                topic_tier=tier,
                topic_score=score,
            )
            for name, grade, stage, length, tier, score in (
                ("高中扩展题材", "高一", "高中", 700, "扩展", 40),
                ("初中精确题材一", "初一", "初中", 300, "精确", 120),
                ("初中精确题材二", "初二", "初中", 300, "精确", 120),
            )
        ]
        selected = select_group(rows, [300, 300], 2, 2005, 2012)
        self.assertIn("高中扩展题材", {row.title for row in selected})
        self.assertEqual(sum(row.school_stage == "高中" for row in selected), 1)

    def test_cross_group_deduplication_keeps_only_one_copy(self) -> None:
        text = "同一篇正文" * 80
        article = ParsedArticle(
            url="https://www.zuowen.com/e/20100101/shared.shtml",
            article_id="shared",
            title="共享正文",
            breadcrumb="作文 > 高中作文 > 高一 > 议论文",
            grade="高一",
            genre_label="议论文",
            published_date="2010-01-01",
            source_label="作文网",
            raw_text=text,
            clean_text=text,
            han_count=400,
            text_hash="same-hash",
        )
        pools = {code: [] for code in TOPIC_RULES}
        pools["NY1"] = [
            CandidateAudit(
                native_code="NY1",
                learner_code="Y1",
                url=article.url,
                title=article.title,
                published_date=article.published_date,
                school_stage="高中",
                topic_tier="宽泛",
                topic_score=30,
                text_hash=article.text_hash,
                article=article,
            )
        ]
        pools["NY2"] = [
            CandidateAudit(
                native_code="NY2",
                learner_code="Y2",
                url=article.url,
                title=article.title,
                published_date=article.published_date,
                school_stage="高中",
                topic_tier="扩展",
                topic_score=50,
                text_hash=article.text_hash,
                article=article,
            )
        ]
        kept, rejected = deduplicate_across_groups(pools)
        self.assertEqual(sum(len(rows) for rows in kept.values()), 1)
        self.assertEqual(len(rejected), 1)
        self.assertEqual(kept["NY2"][0].title, "共享正文")

    def test_candidate_sort_prioritizes_topic_relevance(self) -> None:
        exact_search = SearchCandidate(
            native_code="NY1",
            url="https://www.zuowen.com/e/20100101/exact.shtml",
            title="吸烟的危害",
            search_date="2010-01-01",
            query="吸烟危害",
            source_label="搜索",
            nominal_length=500,
            pre_tier="精确",
            pre_score=100,
            verified_high_school_hint=False,
        )
        broad_listing = SearchCandidate(
            native_code="NY1",
            url="https://www.zuowen.com/e/20100101/broad.shtml",
            title="生命的思考",
            search_date="2010-01-01",
            query="高中议论文栏目",
            source_label="栏目",
            nominal_length=500,
            pre_tier="扩展",
            pre_score=20,
            verified_high_school_hint=True,
        )
        self.assertLess(
            candidate_sort_key(exact_search, 2005, 2012),
            candidate_sort_key(broad_listing, 2005, 2012),
        )

    def test_character_ngram_duplicate_similarity(self) -> None:
        first = character_ngrams("这是一个用于测试近似重复的正文。" * 20)
        second = character_ngrams("这是一个用于测试近似重复的正文。" * 20 + "结尾")
        other = character_ngrams("内容完全不同。" * 20)
        self.assertGreaterEqual(jaccard(first, second), 0.85)
        self.assertLess(jaccard(first, other), 0.5)


class SlowHttpClientTests(unittest.TestCase):
    def test_cache_supports_resume_without_live_request(self) -> None:
        with TemporaryDirectory() as directory:
            client = SlowHttpClient(Path(directory), 0, 0, 1, 1)
            url = "https://www.zuowen.com/e/20100528/abc123.shtml"
            client._cache_path(url).write_bytes(b"cached")
            with patch("scripts.collect_zuowen_native_controls.urlopen") as mocked:
                self.assertEqual(client.fetch(url), b"cached")
            mocked.assert_not_called()
            self.assertEqual(client.cache_hits, 1)

    def test_retry_then_atomic_cache(self) -> None:
        class Response:
            def __enter__(self):
                return self

            def __exit__(self, *args):
                return False

            def read(self):
                return b"fresh"

        with TemporaryDirectory() as directory:
            client = SlowHttpClient(Path(directory), 0, 0, 1, 1)
            url = "https://www.zuowen.com/e/20100528/abc123.shtml"
            failure = HTTPError(url, 429, "rate limited", {}, BytesIO())
            with patch(
                "scripts.collect_zuowen_native_controls.urlopen",
                side_effect=[failure, Response()],
            ), patch("scripts.collect_zuowen_native_controls.time.sleep"):
                self.assertEqual(client.fetch(url), b"fresh")
            self.assertEqual(client._cache_path(url).read_bytes(), b"fresh")
            self.assertEqual(client.live_requests, 1)


class NativeStatisticsInputTests(unittest.TestCase):
    def test_prunes_only_stale_generated_text_files(self) -> None:
        with TemporaryDirectory() as directory:
            base = Path(directory)
            code_dir = base / "NJ1"
            code_dir.mkdir()
            current = code_dir / "NJ1_001.txt"
            stale = code_dir / "NJ1_002.txt"
            unrelated = code_dir / "notes.csv"
            current.write_text("current", encoding="utf-8")
            stale.write_text("stale", encoding="utf-8")
            unrelated.write_text("keep", encoding="utf-8")

            removed = prune_generated_texts(base, {"NJ1": {"NJ1_001"}})

            self.assertEqual(removed, [stale])
            self.assertTrue(current.exists())
            self.assertFalse(stale.exists())
            self.assertTrue(unrelated.exists())

    def test_loads_native_records_without_inventing_scores(self) -> None:
        document = {
            "selected": [
                {
                    "母语代码": "NJ1",
                    "对应学习者篇名": "记对我影响最大的一个人",
                    "作文文件名": "NJ1_001",
                    "网页文章ID": "12345",
                    "作文题目": "我的老师",
                    "体裁": "记叙文",
                }
            ]
        }
        with TemporaryDirectory() as directory:
            path = Path(directory) / "selected.json"
            path.write_text(json.dumps(document, ensure_ascii=False), encoding="utf-8")
            rows = load_native_records(path)
        self.assertEqual(rows[0].score, "")
        self.assertEqual(rows[0].nationality, "中国（公开网络样本）")

    def test_schema_alignment_reorders_without_changing_values(self) -> None:
        fields = [
            FieldSpec("B", "c", "", "", "", "General"),
            FieldSpec("A", "c", "", "", "", "General"),
        ]
        self.assertEqual(align_rows_to_learner_schema(fields, [[2, 1]], ["A", "B"]), [[1, 2]])

    def test_schema_alignment_zero_fills_absent_dynamic_pos_columns(self) -> None:
        fields = [FieldSpec("A", "c", "", "", "", "General")]
        self.assertEqual(
            align_rows_to_learner_schema(
                fields,
                [[1]],
                ["A", "其他词性_string数", "其他词性_string每千字"],
            ),
            [[1, 0, 0]],
        )

    def test_schema_alignment_rejects_absent_regular_columns(self) -> None:
        fields = [FieldSpec("A", "c", "", "", "", "General")]
        with self.assertRaises(ValueError):
            align_rows_to_learner_schema(fields, [[1]], ["A", "名词数"])

    def test_prunes_stale_segmented_texts(self) -> None:
        document = {
            "selected": [
                {
                    "母语代码": "NJ1",
                    "对应学习者篇名": "记对我影响最大的一个人",
                    "作文文件名": "NJ1_001",
                    "网页文章ID": "12345",
                    "作文题目": "我的老师",
                    "体裁": "记叙文",
                }
            ]
        }
        with TemporaryDirectory() as directory:
            base = Path(directory)
            selected_path = base / "selected.json"
            selected_path.write_text(json.dumps(document, ensure_ascii=False), encoding="utf-8")
            code_dir = base / "seg" / "NJ1"
            code_dir.mkdir(parents=True)
            current = code_dir / "NJ1_001.txt"
            stale = code_dir / "NJ1_002.txt"
            current.write_text("current", encoding="utf-8")
            stale.write_text("stale", encoding="utf-8")

            removed = prune_stale_segmented_texts(base / "seg", load_native_records(selected_path))

            self.assertEqual(removed, [stale])
            self.assertTrue(current.exists())
            self.assertFalse(stale.exists())


if __name__ == "__main__":
    unittest.main()
