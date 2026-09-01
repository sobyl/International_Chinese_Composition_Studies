#!/usr/bin/env python3
"""Compare the public online native reference corpus with the HSK learner corpus."""

from __future__ import annotations

import argparse
import json
import math
import os
import warnings
from pathlib import Path
from typing import Any, Iterable, Sequence

os.environ.setdefault("MPLCONFIGDIR", "/tmp/hsk-native-control-matplotlib")

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import statsmodels.formula.api as smf
from scipy import stats
from scipy.optimize import linear_sum_assignment
from statsmodels.stats.multitest import multipletests

try:
    from .analyze_composition_mfmd import (
        bootstrap_stability,
        configure_plotting,
        fit_factor_model,
        hedges_g,
        name_factors,
        parallel_analysis,
        patch_factor_analyzer_compatibility,
        screen_features,
        tucker_congruence,
    )
except ImportError:
    from analyze_composition_mfmd import (
        bootstrap_stability,
        configure_plotting,
        fit_factor_model,
        hedges_g,
        name_factors,
        parallel_analysis,
        patch_factor_analyzer_compatibility,
        screen_features,
        tucker_congruence,
    )


DEFAULT_LEARNER_WORKBOOK = "作文词性统计宽表.xlsx"
DEFAULT_NATIVE_PAYLOAD = "outputs/native_control/native_stats_payload.json"
DEFAULT_SELECTED_JSON = "outputs/native_control/selected_samples.json"
DEFAULT_MFMD_DIR = "outputs/mfmd_analysis"
DEFAULT_OUTPUT_DIR = "outputs/native_control_analysis"
DEFAULT_SEED = 20260830

GROUP_PAIRS = {
    "J1": "NJ1",
    "J2": "NJ2",
    "Y1": "NY1",
    "Y2": "NY2",
}
DIMENSION_ORDER = [
    "词汇丰富度与词汇扩展",
    "基础词汇与叙事推进",
    "人称指涉与信息密度",
    "句法延展与分句复杂度",
    "动作过程与动词链",
]
COLORS = {
    "J1": "#007C91",
    "NJ1": "#65AEB8",
    "J2": "#2A6F97",
    "NJ2": "#83A9C2",
    "Y1": "#B4473A",
    "NY1": "#D9887D",
    "Y2": "#B37A24",
    "NY2": "#DBB46A",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="公开网络母语参照语料与HSK学习者作文对照分析。")
    parser.add_argument("--learner-workbook", default=DEFAULT_LEARNER_WORKBOOK)
    parser.add_argument("--native-payload", default=DEFAULT_NATIVE_PAYLOAD)
    parser.add_argument("--selected-json", default=DEFAULT_SELECTED_JSON)
    parser.add_argument("--mfmd-dir", default=DEFAULT_MFMD_DIR)
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--seed", type=int, default=DEFAULT_SEED)
    parser.add_argument("--bootstrap", type=int, default=2000)
    parser.add_argument("--length-resamples", type=int, default=1000)
    parser.add_argument("--parallel-iterations", type=int, default=1000)
    parser.add_argument("--joint-bootstrap", type=int, default=200)
    parser.add_argument("--skip-joint-bootstrap", action="store_true")
    return parser.parse_args()


def json_value(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_value(item) for item in value]
    if isinstance(value, np.ndarray):
        return value.tolist()
    if isinstance(value, np.integer):
        return int(value)
    if isinstance(value, np.floating):
        return None if not np.isfinite(value) else float(value)
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    if pd.isna(value):
        return None
    return value


def read_inputs(
    learner_workbook: Path,
    native_payload_path: Path,
    selected_json_path: Path,
    mfmd_dir: Path,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict[str, str], pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    learner = pd.read_excel(learner_workbook, sheet_name="词性统计", dtype={"作文编码": str})
    dictionary = pd.read_excel(learner_workbook, sheet_name="字段说明")
    category_map = dict(zip(dictionary["字段名"], dictionary["类别"], strict=True))
    native_document = json.loads(native_payload_path.read_text(encoding="utf-8"))
    native = pd.DataFrame(native_document["rows"], columns=native_document["headers"])
    selected_document = json.loads(selected_json_path.read_text(encoding="utf-8"))
    selected = pd.DataFrame(selected_document["selected"])
    assignments = pd.read_csv(mfmd_dir / "tables" / "指标维度归属.csv", encoding="utf-8-sig")
    loadings = pd.read_csv(mfmd_dir / "tables" / "因子载荷.csv", encoding="utf-8-sig")
    return learner, native, selected, category_map, assignments, loadings, selected_document


def validate_inputs(
    learner: pd.DataFrame,
    native: pd.DataFrame,
    selected: pd.DataFrame,
    assignments: pd.DataFrame,
) -> dict[str, Any]:
    errors: list[str] = []
    if learner.shape[0] != 620 or learner.shape[1] < 301:
        errors.append(f"学习者宽表应为620行且至少301列，实际{learner.shape}")
    if native.shape[1] != learner.shape[1]:
        errors.append(f"母语宽表应与学习者宽表同列数，实际{native.shape[1]}与{learner.shape[1]}")
    if list(learner.columns) != list(native.columns):
        errors.append("母语宽表字段顺序与学习者宽表不一致")
    if len(native) != len(selected):
        errors.append(f"母语统计行数{len(native)}与样本主表{len(selected)}不一致")
    native_counts = native["篇名代码"].value_counts().to_dict()
    if len(set(native_counts.values())) != 1 or set(native_counts) != set(GROUP_PAIRS.values()):
        errors.append(f"母语四组样本数不等或代码不完整：{native_counts}")
    if learner["篇名代码"].value_counts().to_dict() != {code: 155 for code in GROUP_PAIRS}:
        errors.append("学习者四组不再是各155篇")
    if native["作文文件名"].duplicated().any() or selected["作文文件名"].duplicated().any():
        errors.append("母语作文文件名重复")
    if set(native["作文文件名"]) != set(selected["作文文件名"]):
        errors.append("母语统计与样本主表的作文文件名无法一一映射")
    if set(assignments["所属维度"]) != set(DIMENSION_ORDER) or len(assignments) != 39:
        errors.append("原五维39指标归属表不符合预期")
    numeric_native = native.select_dtypes(include=[np.number])
    if np.isinf(numeric_native.to_numpy(dtype=float)).any():
        errors.append("母语统计存在无穷值")
    if errors:
        raise ValueError("；".join(errors))
    return {
        "learner_rows": len(learner),
        "native_rows": len(native),
        "columns": learner.shape[1],
        "native_group_counts": {key: int(value) for key, value in native_counts.items()},
        "selected_features": len(assignments),
    }


def build_projection(
    learner: pd.DataFrame,
    native: pd.DataFrame,
    assignments: pd.DataFrame,
    published_scores_path: Path,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    feature_names = assignments["字段名"].tolist()
    means = learner[feature_names].astype(float).mean(axis=0)
    standard_deviations = learner[feature_names].astype(float).std(axis=0, ddof=1).replace(0, 1.0)
    learner_z = (learner[feature_names].astype(float) - means) / standard_deviations
    native_z = (native[feature_names].astype(float) - means) / standard_deviations

    learner_scores = learner[["篇名代码", "作文文件名", "纯文本字数", "国籍", "作文分数"]].copy()
    native_scores = native[["篇名代码", "作文文件名", "纯文本字数", "国籍", "作文分数"]].copy()
    parameter_rows: list[dict[str, Any]] = []
    raw_parameters: dict[str, tuple[float, float]] = {}
    for dimension in DIMENSION_ORDER:
        rows = assignments.loc[assignments["所属维度"] == dimension]
        features = rows["字段名"].tolist()
        signs = np.sign(rows["主要载荷"].to_numpy(dtype=float))
        learner_raw = (learner_z[features].to_numpy(dtype=float) * signs).sum(axis=1)
        native_raw = (native_z[features].to_numpy(dtype=float) * signs).sum(axis=1)
        raw_mean = float(np.mean(learner_raw))
        raw_sd = float(np.std(learner_raw, ddof=1)) or 1.0
        learner_scores[dimension] = (learner_raw - raw_mean) / raw_sd
        native_scores[dimension] = (native_raw - raw_mean) / raw_sd
        raw_parameters[dimension] = (raw_mean, raw_sd)
        for row in rows.itertuples(index=False):
            parameter_rows.append(
                {
                    "维度": dimension,
                    "字段名": row.字段名,
                    "载荷方向": 1 if float(row.主要载荷) >= 0 else -1,
                    "主要载荷": float(row.主要载荷),
                    "学习者均值": float(means[row.字段名]),
                    "学习者标准差": float(standard_deviations[row.字段名]),
                    "维度原始和均值": raw_mean,
                    "维度原始和标准差": raw_sd,
                }
            )

    published = pd.read_csv(published_scores_path, encoding="utf-8-sig")
    merged = learner_scores.merge(
        published[["作文文件名", *DIMENSION_ORDER]],
        on="作文文件名",
        suffixes=("_重投影", "_原结果"),
        validate="one_to_one",
    )
    maximum_error = max(
        float(np.max(np.abs(merged[f"{dimension}_重投影"] - merged[f"{dimension}_原结果"])))
        for dimension in DIMENSION_ORDER
    )
    if maximum_error > 1e-9:
        raise ValueError(f"学习者重投影未通过回归校验，最大差异={maximum_error:.3e}")
    diagnostics = {
        "feature_count": len(feature_names),
        "learner_reprojection_max_abs_error": maximum_error,
        "raw_dimension_parameters": {
            dimension: {"mean": values[0], "sd": values[1]}
            for dimension, values in raw_parameters.items()
        },
    }
    return learner_scores, native_scores, pd.DataFrame(parameter_rows), diagnostics


def bootstrap_comparison_intervals(
    native: np.ndarray,
    learner: np.ndarray,
    iterations: int,
    rng: np.random.Generator,
) -> tuple[float, float, float, float]:
    differences = np.empty(iterations, dtype=float)
    effect_sizes = np.empty(iterations, dtype=float)
    for index in range(iterations):
        native_sample = rng.choice(native, len(native), replace=True)
        learner_sample = rng.choice(learner, len(learner), replace=True)
        differences[index] = float(native_sample.mean() - learner_sample.mean())
        effect_sizes[index] = float(hedges_g(native_sample, learner_sample))
    return (
        float(np.quantile(differences, 0.025)),
        float(np.quantile(differences, 0.975)),
        float(np.quantile(effect_sizes, 0.025)),
        float(np.quantile(effect_sizes, 0.975)),
    )


def dimension_comparisons(
    learner_scores: pd.DataFrame,
    native_scores: pd.DataFrame,
    iterations: int,
    seed: int,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    rng = np.random.default_rng(seed)
    comparison_records: list[dict[str, Any]] = []
    descriptive_records: list[dict[str, Any]] = []
    for learner_code, native_code in GROUP_PAIRS.items():
        for source, code, frame in (
            ("学习者", learner_code, learner_scores),
            ("公开网络母语参照", native_code, native_scores),
        ):
            subset = frame.loc[frame["篇名代码"] == code]
            for dimension in DIMENSION_ORDER:
                values = subset[dimension].to_numpy(dtype=float)
                descriptive_records.append(
                    {
                        "对应题目": learner_code,
                        "语言来源": source,
                        "代码": code,
                        "维度": dimension,
                        "样本量": len(values),
                        "均值": float(np.mean(values)),
                        "标准差": float(np.std(values, ddof=1)),
                        "中位数": float(np.median(values)),
                    }
                )
        for dimension in DIMENSION_ORDER:
            learner_values = learner_scores.loc[
                learner_scores["篇名代码"] == learner_code, dimension
            ].to_numpy(dtype=float)
            native_values = native_scores.loc[
                native_scores["篇名代码"] == native_code, dimension
            ].to_numpy(dtype=float)
            test = stats.ttest_ind(native_values, learner_values, equal_var=False)
            low, high, effect_low, effect_high = bootstrap_comparison_intervals(
                native_values, learner_values, iterations, rng
            )
            comparison_records.append(
                {
                    "对应题目": learner_code,
                    "学习者代码": learner_code,
                    "母语代码": native_code,
                    "维度": dimension,
                    "学习者样本量": len(learner_values),
                    "母语样本量": len(native_values),
                    "学习者均值": float(np.mean(learner_values)),
                    "母语均值": float(np.mean(native_values)),
                    "均值差_母语减学习者": float(np.mean(native_values) - np.mean(learner_values)),
                    "Bootstrap95%CI下限": low,
                    "Bootstrap95%CI上限": high,
                    "Welch_t": float(test.statistic),
                    "Welch_p值": float(test.pvalue),
                    "Hedges_g_母语减学习者": float(hedges_g(native_values, learner_values)),
                    "Hedges_g_Bootstrap95%CI下限": effect_low,
                    "Hedges_g_Bootstrap95%CI上限": effect_high,
                }
            )
    comparisons = pd.DataFrame(comparison_records)
    comparisons["Holm校正p值"] = multipletests(comparisons["Welch_p值"], method="holm")[1]
    comparisons["Holm显著"] = comparisons["Holm校正p值"] < 0.05
    return comparisons, pd.DataFrame(descriptive_records)


def feature_comparisons(
    learner: pd.DataFrame,
    native: pd.DataFrame,
    assignments: pd.DataFrame,
) -> pd.DataFrame:
    records: list[dict[str, Any]] = []
    for learner_code, native_code in GROUP_PAIRS.items():
        left = learner.loc[learner["篇名代码"] == learner_code]
        right = native.loc[native["篇名代码"] == native_code]
        for feature in assignments["字段名"]:
            learner_values = left[feature].to_numpy(dtype=float)
            native_values = right[feature].to_numpy(dtype=float)
            test = stats.ttest_ind(native_values, learner_values, equal_var=False)
            records.append(
                {
                    "对应题目": learner_code,
                    "字段名": feature,
                    "所属维度": assignments.loc[assignments["字段名"] == feature, "所属维度"].iloc[0],
                    "学习者均值": float(np.mean(learner_values)),
                    "母语均值": float(np.mean(native_values)),
                    "均值差_母语减学习者": float(np.mean(native_values) - np.mean(learner_values)),
                    "Hedges_g_母语减学习者": float(hedges_g(native_values, learner_values)),
                    "Welch_p值": float(test.pvalue),
                }
            )
    frame = pd.DataFrame(records)
    frame["BH校正p值"] = multipletests(frame["Welch_p值"], method="fdr_bh")[1]
    frame["BH显著"] = frame["BH校正p值"] < 0.05
    return frame.sort_values(["对应题目", "BH校正p值", "字段名"])


def robust_regressions(
    learner_scores: pd.DataFrame,
    native_scores: pd.DataFrame,
) -> pd.DataFrame:
    learner = learner_scores.copy()
    learner["语言来源"] = "学习者"
    learner["对应题目"] = learner["篇名代码"]
    native = native_scores.copy()
    native["语言来源"] = "公开网络母语参照"
    native["对应题目"] = native["篇名代码"].map({value: key for key, value in GROUP_PAIRS.items()})
    combined = pd.concat([learner, native], ignore_index=True)
    combined["log纯文本字数"] = np.log(combined["纯文本字数"].astype(float).clip(lower=1))
    records: list[dict[str, Any]] = []
    for dimension in DIMENSION_ORDER:
        local = combined.rename(columns={dimension: "outcome"})
        model = smf.ols(
            "outcome ~ C(语言来源, Treatment(reference='学习者')) * C(对应题目, Treatment(reference='J1')) + log纯文本字数",
            data=local,
        ).fit(cov_type="HC3")
        intervals = model.conf_int()
        for term in model.params.index:
            records.append(
                {
                    "维度": dimension,
                    "项": term,
                    "系数": float(model.params[term]),
                    "HC3标准误": float(model.bse[term]),
                    "t值": float(model.tvalues[term]),
                    "p值": float(model.pvalues[term]),
                    "95%CI下限": float(intervals.loc[term, 0]),
                    "95%CI上限": float(intervals.loc[term, 1]),
                    "样本量": int(model.nobs),
                    "调整R平方": float(model.rsquared_adj),
                }
            )
    frame = pd.DataFrame(records)
    frame["BH校正p值"] = multipletests(frame["p值"], method="fdr_bh")[1]
    return frame


def greedy_length_match(
    learner: pd.DataFrame,
    native: pd.DataFrame,
    rng: np.random.Generator,
) -> pd.DataFrame:
    remaining = learner.copy()
    order = rng.permutation(len(native))
    selected_indices: list[int] = []
    for position in order:
        target = float(native.iloc[position]["纯文本字数"])
        differences = np.abs(remaining["纯文本字数"].to_numpy(dtype=float) - target)
        jitter = rng.uniform(0, 1e-6, size=len(differences))
        chosen_position = int(np.argmin(differences + jitter))
        selected_indices.append(remaining.index[chosen_position])
        remaining = remaining.drop(remaining.index[chosen_position])
    return learner.loc[selected_indices]


def length_matched_resampling(
    learner_scores: pd.DataFrame,
    native_scores: pd.DataFrame,
    iterations: int,
    seed: int,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    rng = np.random.default_rng(seed + 99)
    raw_records: list[dict[str, Any]] = []
    for learner_code, native_code in GROUP_PAIRS.items():
        learner = learner_scores.loc[learner_scores["篇名代码"] == learner_code].copy()
        native = native_scores.loc[native_scores["篇名代码"] == native_code].copy()
        for iteration in range(1, iterations + 1):
            matched = greedy_length_match(learner, native, rng)
            for dimension in DIMENSION_ORDER:
                raw_records.append(
                    {
                        "迭代": iteration,
                        "对应题目": learner_code,
                        "维度": dimension,
                        "均值差_母语减匹配学习者": float(native[dimension].mean() - matched[dimension].mean()),
                        "匹配后篇幅均值差": float(native["纯文本字数"].mean() - matched["纯文本字数"].mean()),
                    }
                )
    raw = pd.DataFrame(raw_records)
    summary = (
        raw.groupby(["对应题目", "维度"], as_index=False)
        .agg(
            重抽样均值差=("均值差_母语减匹配学习者", "mean"),
            重抽样标准差=("均值差_母语减匹配学习者", "std"),
            篇幅均值差=("匹配后篇幅均值差", "mean"),
        )
    )
    quantiles = (
        raw.groupby(["对应题目", "维度"])["均值差_母语减匹配学习者"]
        .quantile([0.025, 0.975])
        .unstack()
        .reset_index()
        .rename(columns={0.025: "重抽样95%CI下限", 0.975: "重抽样95%CI上限"})
    )
    summary = summary.merge(quantiles, on=["对应题目", "维度"], validate="one_to_one")
    return raw, summary


def sensitivity_analysis(
    learner_scores: pd.DataFrame,
    native_scores: pd.DataFrame,
    selected: pd.DataFrame,
) -> pd.DataFrame:
    enriched = native_scores.merge(
        selected[["作文文件名", "主题匹配层级", "发布日期", "年代扩展样本"]],
        on="作文文件名",
        validate="one_to_one",
    )
    enriched["发布日期"] = enriched["发布日期"].astype(str)
    subsets = {
        "全部母语样本": pd.Series(True, index=enriched.index),
        "精确或近似主题": enriched["主题匹配层级"].isin(["精确", "近似"]),
        "2005至2012年": ~enriched["年代扩展样本"].astype(bool),
    }
    records: list[dict[str, Any]] = []
    for subset_name, mask in subsets.items():
        local = enriched.loc[mask]
        for learner_code, native_code in GROUP_PAIRS.items():
            learner_group = learner_scores.loc[learner_scores["篇名代码"] == learner_code]
            native_group = local.loc[local["篇名代码"] == native_code]
            for dimension in DIMENSION_ORDER:
                records.append(
                    {
                        "敏感性样本": subset_name,
                        "对应题目": learner_code,
                        "维度": dimension,
                        "母语样本量": len(native_group),
                        "均值差_母语减学习者": (
                            float(native_group[dimension].mean() - learner_group[dimension].mean())
                            if len(native_group)
                            else None
                        ),
                    }
                )
    return pd.DataFrame(records)


def sampling_audit(selected: pd.DataFrame, selected_document: dict[str, Any]) -> tuple[pd.DataFrame, pd.DataFrame]:
    summary_rows: list[dict[str, Any]] = []
    for code in GROUP_PAIRS.values():
        group = selected.loc[selected["母语代码"] == code]
        summary_rows.append(
            {
                "母语代码": code,
                "样本量": len(group),
                "平均汉字数": float(group["正文汉字数"].mean()),
                "中位汉字数": float(group["正文汉字数"].median()),
                "精确主题": int((group["主题匹配层级"] == "精确").sum()),
                "近似主题": int((group["主题匹配层级"] == "近似").sum()),
                "扩展主题": int((group["主题匹配层级"] == "扩展").sum()),
                "宽泛主题": int((group["主题匹配层级"] == "宽泛").sum()),
                "高一": int((group["年级"] == "高一").sum()),
                "高二": int((group["年级"] == "高二").sum()),
                "高三": int((group["年级"] == "高三").sum()),
                "年级未标注": int((group["年级"] == "未标注").sum()),
                "年代扩展样本": int(group["年代扩展样本"].astype(bool).sum()),
                "平均目标篇幅相对偏差": float(group["目标篇幅相对偏差"].mean()),
            }
        )
    source_rows = [
        {"项目": "语料定位", "说明": selected_document.get("corpus_label", "")},
        {"项目": "来源栏目", "说明": selected_document.get("source_site", "")},
        {"项目": "服务协议", "说明": selected_document.get("source_agreement", "")},
        {"项目": "全文处理", "说明": "完整网页正文仅保存于本地并排除Git；报告只使用必要的极短例句。"},
        {"项目": "身份边界", "说明": "高中栏目归类不等同于作者身份认证，且无法独立排除网站编辑。"},
    ]
    return pd.DataFrame(summary_rows), pd.DataFrame(source_rows)


def joint_factor_sensitivity(
    combined: pd.DataFrame,
    category_map: dict[str, str],
    original_loadings: pd.DataFrame,
    output_dir: Path,
    seed: int,
    parallel_iterations: int,
    bootstrap_iterations: int,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    screening, features, matrix, screening_diag = screen_features(
        combined,
        category_map,
        zero_threshold=0.95,
        corr_threshold=0.85,
        msa_threshold=0.50,
    )
    observed, random_threshold, suggested = parallel_analysis(matrix, parallel_iterations, seed)
    candidates: list[Any] = []
    maximum = min(10, max(2, suggested))
    for factor_count in range(2, maximum + 1):
        try:
            model = fit_factor_model(matrix, features, factor_count, 0.30, 0.30)
        except Exception as error:
            print(f"联合模型{factor_count}因子失败：{error}")
            continue
        if model.diagnostic_valid and bootstrap_iterations > 0:
            final_matrix = combined[model.feature_names].to_numpy(dtype=float)
            medians, q05, successes = bootstrap_stability(
                final_matrix, model, bootstrap_iterations, seed + 7000
            )
            model.stability_medians = medians
            model.stability_q05 = q05
            model.bootstrap_successes = successes
            model.stable = successes >= max(20, int(bootstrap_iterations * 0.80)) and min(medians) >= 0.85
        elif model.diagnostic_valid:
            model.stability_medians = [1.0] * model.k
            model.stability_q05 = [1.0] * model.k
            model.stable = True
        candidates.append(model)
        print(
            f"联合{factor_count}因子：变量{len(model.feature_names)}，KMO={model.kmo:.3f}，"
            f"诊断={model.diagnostic_valid}，稳定={model.stable}",
            flush=True,
        )
    stable = [model for model in candidates if model.diagnostic_valid and model.stable]
    diagnostic = [model for model in candidates if model.diagnostic_valid]
    selected_model = (
        min(stable, key=lambda model: abs(model.k - suggested))
        if stable
        else (min(diagnostic, key=lambda model: abs(model.k - suggested)) if diagnostic else None)
    )
    model_rows = []
    for model in candidates:
        model_rows.append(
            {
                "因子数": model.k,
                "变量数": len(model.feature_names),
                "KMO": model.kmo,
                "累计解释方差": float(model.factor_variance[2, -1]),
                "诊断通过": model.diagnostic_valid,
                "稳定性通过": model.stable,
                "Bootstrap成功次数": model.bootstrap_successes,
                "最小Tucker中位数": min(model.stability_medians or [0]),
                "是否选定": model is selected_model,
                "诊断说明": model.diagnostic_note,
            }
        )
    parallel = pd.DataFrame(
        {
            "因子序号": np.arange(1, len(observed) + 1),
            "实测特征根": observed,
            "随机95%阈值": random_threshold,
            "实测大于随机阈值": observed > random_threshold,
        }
    )
    if selected_model is None:
        empty = pd.DataFrame(columns=["原维度", "联合维度", "Tucker一致性", "绝对Tucker一致性"])
        return screening, pd.DataFrame(model_rows), empty, {
            "suggested_factors": suggested,
            "selected_factors": None,
            "screening": screening_diag,
            "parallel": parallel.to_dict(orient="records"),
        }

    joint_labels, _ = name_factors(selected_model)
    original_dimension_columns = [column for column in original_loadings if column.startswith("D") and "_" in column]
    original_names = [column.split("_", 1)[1] for column in original_dimension_columns]
    all_features = sorted(set(original_loadings["字段名"]) | set(selected_model.feature_names))
    reference = np.zeros((len(all_features), len(original_dimension_columns)))
    candidate = np.zeros((len(all_features), selected_model.k))
    feature_index = {name: index for index, name in enumerate(all_features)}
    for _, row in original_loadings.iterrows():
        index = feature_index[str(row["字段名"])]
        for factor, column in enumerate(original_dimension_columns):
            reference[index, factor] = float(row[column])
    for index, feature in enumerate(selected_model.feature_names):
        candidate[feature_index[feature], :] = selected_model.loadings[index, :]
    congruence = tucker_congruence(reference, candidate)
    rows, columns = linear_sum_assignment(-np.abs(congruence))
    comparison_records: list[dict[str, Any]] = []
    for row, column in zip(rows, columns, strict=True):
        comparison_records.append(
            {
                "原维度": original_names[row],
                "联合维度": joint_labels[column],
                "原维度序号": row + 1,
                "联合维度序号": column + 1,
                "Tucker一致性": float(congruence[row, column]),
                "绝对Tucker一致性": abs(float(congruence[row, column])),
            }
        )
    joint_loading_rows: list[dict[str, Any]] = []
    for index, feature in enumerate(selected_model.feature_names):
        record: dict[str, Any] = {"字段名": feature, "类别": category_map.get(feature, "")}
        for factor, label in enumerate(joint_labels):
            record[f"J{factor + 1}_{label}"] = float(selected_model.loadings[index, factor])
        joint_loading_rows.append(record)
    pd.DataFrame(joint_loading_rows).to_csv(
        output_dir / "tables" / "联合因子载荷.csv", index=False, encoding="utf-8-sig"
    )
    parallel.to_csv(output_dir / "tables" / "联合平行分析.csv", index=False, encoding="utf-8-sig")
    metadata = {
        "suggested_factors": suggested,
        "selected_factors": selected_model.k,
        "selected_labels": joint_labels,
        "selected_features": len(selected_model.feature_names),
        "selected_kmo": selected_model.kmo,
        "screening": screening_diag,
        "parallel": parallel.to_dict(orient="records"),
    }
    return screening, pd.DataFrame(model_rows), pd.DataFrame(comparison_records), metadata


def save_figure(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    plt.tight_layout()
    plt.savefig(path, dpi=300, bbox_inches="tight")
    plt.close()


def portable_path(path: Path) -> str:
    try:
        return str(path.relative_to(Path.cwd()))
    except ValueError:
        return str(path)


def create_figures(
    learner: pd.DataFrame,
    native: pd.DataFrame,
    selected: pd.DataFrame,
    learner_scores: pd.DataFrame,
    native_scores: pd.DataFrame,
    comparisons: pd.DataFrame,
    feature_results: pd.DataFrame,
    resampling_raw: pd.DataFrame,
    congruence: pd.DataFrame,
    figures_dir: Path,
) -> pd.DataFrame:
    figures_dir.mkdir(parents=True, exist_ok=True)
    records: list[dict[str, Any]] = []

    tier = selected.groupby(["母语代码", "主题匹配层级"]).size().unstack(fill_value=0).reindex(GROUP_PAIRS.values())
    tier = tier.reindex(
        columns=[value for value in ("精确", "近似", "扩展", "宽泛") if value in tier],
        fill_value=0,
    )
    tier.plot(
        kind="bar",
        stacked=True,
        color=["#2F7D6D", "#E0A72E", "#B4473A", "#6D7480"],
        figsize=(8.0, 4.6),
    )
    plt.title("母语参照样本的主题匹配层级")
    plt.xlabel("母语代码")
    plt.ylabel("作文篇数")
    plt.xticks(rotation=0)
    plt.legend(title="匹配层级", frameon=False)
    path = figures_dir / "01_采样主题层级.png"
    save_figure(path)
    records.append({"序号": 1, "图名": "采样主题层级", "文件": portable_path(path)})

    length_frame = pd.concat(
        [
            learner[["篇名代码", "纯文本字数"]].assign(语言来源="学习者", 对应题目=lambda x: x["篇名代码"]),
            native[["篇名代码", "纯文本字数"]].assign(
                语言来源="公开网络母语参照",
                对应题目=lambda x: x["篇名代码"].map({value: key for key, value in GROUP_PAIRS.items()}),
            ),
        ],
        ignore_index=True,
    )
    plt.figure(figsize=(9.2, 4.8))
    sns.violinplot(data=length_frame, x="对应题目", y="纯文本字数", hue="语言来源", split=True, inner="quart", palette=["#526D82", "#C26D4A"])
    plt.title("四个题目中学习者与母语参照作文的篇幅分布")
    plt.xlabel("对应题目")
    plt.ylabel("纯文本汉字数")
    plt.legend(title="", frameon=False)
    path = figures_dir / "02_篇幅匹配诊断.png"
    save_figure(path)
    records.append({"序号": 2, "图名": "篇幅匹配诊断", "文件": portable_path(path)})

    score_frame = pd.concat(
        [
            learner_scores.assign(语言来源="学习者", 对应题目=lambda x: x["篇名代码"]),
            native_scores.assign(
                语言来源="公开网络母语参照",
                对应题目=lambda x: x["篇名代码"].map({value: key for key, value in GROUP_PAIRS.items()}),
            ),
        ],
        ignore_index=True,
    )
    means = score_frame.groupby(["对应题目", "语言来源"])[DIMENSION_ORDER].mean()
    profile = means.reset_index().melt(["对应题目", "语言来源"], var_name="维度", value_name="均值")
    matrix = profile.pivot(index=["对应题目", "语言来源"], columns="维度", values="均值").reindex(columns=DIMENSION_ORDER)
    plt.figure(figsize=(11, 5.4))
    sns.heatmap(matrix, cmap="vlag", center=0, annot=True, fmt=".2f", linewidths=0.5, cbar_kws={"label": "五维得分均值"})
    plt.title("学习者与母语参照语料的五维语言画像")
    plt.xlabel("")
    plt.ylabel("")
    path = figures_dir / "03_五维画像热图.png"
    save_figure(path)
    records.append({"序号": 3, "图名": "五维画像热图", "文件": portable_path(path)})

    long_scores = score_frame.melt(
        id_vars=["对应题目", "语言来源"], value_vars=DIMENSION_ORDER, var_name="维度", value_name="维度得分"
    )
    grid = sns.catplot(
        data=long_scores,
        x="对应题目",
        y="维度得分",
        hue="语言来源",
        col="维度",
        col_wrap=2,
        kind="box",
        sharey=False,
        height=3.2,
        aspect=1.35,
        palette=["#526D82", "#C26D4A"],
        showfliers=False,
    )
    grid.set_titles("{col_name}")
    grid.set_axis_labels("对应题目", "维度得分")
    grid._legend.set_title("")
    path = figures_dir / "04_五维得分分布.png"
    grid.fig.savefig(path, dpi=300, bbox_inches="tight")
    plt.close(grid.fig)
    records.append({"序号": 4, "图名": "五维得分分布", "文件": portable_path(path)})

    plot = comparisons.copy()
    plot["标签"] = plot["对应题目"] + " | " + plot["维度"]
    plot = plot.sort_values("Hedges_g_母语减学习者")
    plt.figure(figsize=(9.0, 8.0))
    colors = [COLORS[GROUP_PAIRS[row.对应题目]] for row in plot.itertuples()]
    positions = np.arange(len(plot))
    plt.hlines(
        positions,
        plot["Hedges_g_Bootstrap95%CI下限"],
        plot["Hedges_g_Bootstrap95%CI上限"],
        colors=colors,
        linewidth=1.4,
        alpha=0.85,
    )
    plt.scatter(plot["Hedges_g_母语减学习者"], positions, c=colors, s=34, zorder=3)
    plt.axvline(0, color="#59636A", linewidth=1)
    plt.yticks(np.arange(len(plot)), plot["标签"], fontsize=8)
    plt.xlabel("Hedges g（正值表示母语参照更高）")
    plt.ylabel("")
    plt.title("五维差异的标准化效应量")
    path = figures_dir / "05_五维效应量森林图.png"
    save_figure(path)
    records.append({"序号": 5, "图名": "五维效应量森林图", "文件": portable_path(path)})

    top_features = (
        feature_results.assign(abs_g=lambda x: x["Hedges_g_母语减学习者"].abs())
        .sort_values("abs_g", ascending=False)
        .groupby("对应题目", as_index=False)
        .head(6)
        .sort_values("Hedges_g_母语减学习者")
    )
    top_features["标签"] = top_features["对应题目"] + " | " + top_features["字段名"]
    plt.figure(figsize=(10.2, 9.0))
    plt.barh(
        np.arange(len(top_features)),
        top_features["Hedges_g_母语减学习者"],
        color=[COLORS[GROUP_PAIRS[code]] for code in top_features["对应题目"]],
    )
    plt.axvline(0, color="#59636A", linewidth=1)
    plt.yticks(np.arange(len(top_features)), top_features["标签"], fontsize=8)
    plt.xlabel("Hedges g（正值表示母语参照更高）")
    plt.title("各题目差异最大的语言特征")
    path = figures_dir / "06_主要语言特征效应量.png"
    save_figure(path)
    records.append({"序号": 6, "图名": "主要语言特征效应量", "文件": portable_path(path)})

    hsk_columns = ["初等词汇占比", "中等词汇占比", "高等词汇占比", "非HSK词汇占比"]
    hsk = pd.concat(
        [
            learner[["篇名代码", *hsk_columns]].copy().assign(语言来源="学习者"),
            native[["篇名代码", *hsk_columns]].copy().assign(语言来源="公开网络母语参照"),
        ],
        ignore_index=True,
    )
    hsk["对应题目"] = hsk["篇名代码"].replace({value: key for key, value in GROUP_PAIRS.items()})
    hsk_means = hsk.groupby(["对应题目", "语言来源"])[hsk_columns].mean()
    hsk_means.plot(kind="bar", stacked=True, figsize=(10, 5.2), color=["#C4493D", "#3D8B5A", "#439BD3", "#E3B832"])
    plt.title("HSK等级词汇与非HSK词汇构成")
    plt.xlabel("")
    plt.ylabel("占非标点分词比例")
    plt.xticks(rotation=30, ha="right")
    plt.legend(title="", frameon=False, ncol=2)
    path = figures_dir / "07_HSK词汇构成.png"
    save_figure(path)
    records.append({"序号": 7, "图名": "HSK词汇构成", "文件": portable_path(path)})

    sampled = resampling_raw.copy()
    sampled["标签"] = sampled["对应题目"] + " | " + sampled["维度"]
    order = sampled.groupby("标签")["均值差_母语减匹配学习者"].median().sort_values().index
    plt.figure(figsize=(10.0, 8.5))
    sns.boxplot(data=sampled, y="标签", x="均值差_母语减匹配学习者", order=order, color="#78A5B5", showfliers=False)
    plt.axvline(0, color="#59636A", linewidth=1)
    plt.xlabel("母语参照减篇幅匹配学习者的维度均值差")
    plt.ylabel("")
    plt.title("1000次篇幅匹配重抽样的稳健性")
    path = figures_dir / "08_篇幅匹配重抽样.png"
    save_figure(path)
    records.append({"序号": 8, "图名": "篇幅匹配重抽样", "文件": portable_path(path)})

    year = pd.to_numeric(selected["发布日期"].astype(str).str[:4], errors="coerce")
    plt.figure(figsize=(8.5, 4.6))
    sns.histplot(data=selected.assign(发布年份=year), x="发布年份", hue="母语代码", multiple="stack", discrete=True, palette=COLORS)
    plt.axvspan(2005, 2012, color="#DDE8EE", alpha=0.35, label="优先年份")
    plt.title("母语参照作文的发布年份分布")
    plt.xlabel("发布年份")
    plt.ylabel("篇数")
    path = figures_dir / "09_发布年份分布.png"
    save_figure(path)
    records.append({"序号": 9, "图名": "发布年份分布", "文件": portable_path(path)})

    if len(congruence):
        matrix = congruence.pivot(index="原维度", columns="联合维度", values="绝对Tucker一致性")
        plt.figure(figsize=(9.5, 5.2))
        sns.heatmap(matrix, cmap="YlGnBu", vmin=0, vmax=1, annot=True, fmt=".2f", linewidths=0.5)
        plt.title(f"原五维结构与{len(learner) + len(native)}篇联合因子模型的Tucker一致性")
        plt.xlabel("联合模型维度")
        plt.ylabel("原学习者五维结构")
    else:
        plt.figure(figsize=(8, 3.5))
        plt.text(0.5, 0.5, "联合因子模型未通过预设诊断，未计算Tucker一致性", ha="center", va="center")
        plt.axis("off")
    path = figures_dir / "10_联合模型Tucker一致性.png"
    save_figure(path)
    records.append({"序号": 10, "图名": "联合模型Tucker一致性", "文件": portable_path(path)})
    return pd.DataFrame(records)


def write_tables(output_dir: Path, tables: dict[str, pd.DataFrame]) -> None:
    tables_dir = output_dir / "tables"
    tables_dir.mkdir(parents=True, exist_ok=True)
    for name, frame in tables.items():
        frame.to_csv(tables_dir / f"{name}.csv", index=False, encoding="utf-8-sig")


def main() -> int:
    args = parse_args()
    warnings.filterwarnings("ignore", category=RuntimeWarning)
    patch_factor_analyzer_compatibility()
    configure_plotting()
    learner_workbook = Path(args.learner_workbook).resolve()
    native_payload = Path(args.native_payload).resolve()
    selected_json = Path(args.selected_json).resolve()
    mfmd_dir = Path(args.mfmd_dir).resolve()
    output_dir = Path(args.output_dir).resolve()
    (output_dir / "tables").mkdir(parents=True, exist_ok=True)

    learner, native, selected, category_map, assignments, original_loadings, selected_document = read_inputs(
        learner_workbook, native_payload, selected_json, mfmd_dir
    )
    validation = validate_inputs(learner, native, selected, assignments)
    print(f"输入校验通过：学习者{len(learner)}篇，母语参照{len(native)}篇")
    learner_scores, native_scores, projection_parameters, projection_diag = build_projection(
        learner,
        native,
        assignments,
        mfmd_dir / "tables" / "维度得分.csv",
    )
    comparisons, dimension_descriptives = dimension_comparisons(
        learner_scores, native_scores, args.bootstrap, args.seed
    )
    feature_results = feature_comparisons(learner, native, assignments)
    regressions = robust_regressions(learner_scores, native_scores)
    resampling_raw, resampling_summary = length_matched_resampling(
        learner_scores, native_scores, args.length_resamples, args.seed
    )
    sensitivity = sensitivity_analysis(learner_scores, native_scores, selected)
    sampling_summary, source_notes = sampling_audit(selected, selected_document)

    combined = pd.concat([learner, native], ignore_index=True)
    joint_screening, joint_models, congruence, joint_metadata = joint_factor_sensitivity(
        combined,
        category_map,
        original_loadings,
        output_dir,
        args.seed,
        args.parallel_iterations,
        0 if args.skip_joint_bootstrap else args.joint_bootstrap,
    )

    figures = create_figures(
        learner,
        native,
        selected,
        learner_scores,
        native_scores,
        comparisons,
        feature_results,
        resampling_raw,
        congruence,
        output_dir / "figures",
    )
    combined_scores = pd.concat(
        [learner_scores.assign(语言来源="学习者"), native_scores.assign(语言来源="公开网络母语参照")],
        ignore_index=True,
    )
    tables = {
        "来源与版权说明": source_notes,
        "采样审计概况": sampling_summary,
        "母语样本来源": selected,
        "投影参数": projection_parameters,
        "五维得分": combined_scores,
        "维度描述统计": dimension_descriptives,
        "五维Welch比较": comparisons,
        "39项特征比较": feature_results,
        "HC3稳健回归": regressions,
        "篇幅匹配重抽样": resampling_summary,
        "主题年份敏感性": sensitivity,
        "联合变量筛选": joint_screening,
        "联合模型诊断": joint_models,
        "Tucker一致性": congruence,
        "图表索引": figures,
    }
    write_tables(output_dir, tables)
    metadata = {
        "generated_at": pd.Timestamp.now(tz="Asia/Shanghai").isoformat(),
        "inputs": {
            "learner_workbook": str(learner_workbook),
            "native_payload": str(native_payload),
            "selected_json": str(selected_json),
            "mfmd_dir": str(mfmd_dir),
        },
        "validation": validation,
        "projection": projection_diag,
        "joint_factor_sensitivity": joint_metadata,
        "random_seed": args.seed,
        "bootstrap_iterations": args.bootstrap,
        "length_resamples": args.length_resamples,
        "parallel_iterations": args.parallel_iterations,
        "joint_bootstrap": 0 if args.skip_joint_bootstrap else args.joint_bootstrap,
        "figures": figures.to_dict(orient="records"),
    }
    (output_dir / "analysis_metadata.json").write_text(
        json.dumps(json_value(metadata), ensure_ascii=False, indent=2), encoding="utf-8"
    )
    payload = {
        "metadata": json_value(metadata),
        "sheets": {
            name: {
                "columns": list(frame.columns),
                "rows": json_value(frame.to_numpy(dtype=object).tolist()),
            }
            for name, frame in tables.items()
        },
    }
    (output_dir / "workbook_payload.json").write_text(
        json.dumps(payload, ensure_ascii=False), encoding="utf-8"
    )
    print(f"母语对照分析完成：{output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
