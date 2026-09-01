#!/usr/bin/env python3
"""Run a reproducible MF/MD analysis for the HSK composition corpus."""

from __future__ import annotations

import argparse
import inspect
import json
import math
import os
import re
import warnings
from dataclasses import asdict, dataclass
from itertools import combinations
from pathlib import Path
from typing import Any, Iterable, Sequence

os.environ.setdefault("MPLCONFIGDIR", "/tmp/hsk-mfmd-matplotlib")

import matplotlib

matplotlib.use("Agg")

import matplotlib.font_manager as font_manager
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import statsmodels.formula.api as smf
from factor_analyzer import FactorAnalyzer
import factor_analyzer.factor_analyzer as factor_analyzer_module
from scipy import stats
from scipy.optimize import linear_sum_assignment
from statsmodels.stats.multitest import multipletests
from statsmodels.stats.oneway import anova_oneway


DEFAULT_WORKBOOK = "作文词性统计宽表.xlsx"
DEFAULT_CLEAN_TEXT_DIR = "clean_text"
DEFAULT_OUTPUT_DIR = "outputs/mfmd_analysis"
DEFAULT_SEED = 20260829

EXPECTED_ROWS = 620
MIN_EXPECTED_COLUMNS = 301
EXPECTED_CODES = ("J1", "J2", "Y1", "Y2")
EXPECTED_CODE_COUNT = 155
EXPECTED_SCORES = (55, 60, 65, 70, 80, 85, 90, 95)

FEATURE_CATEGORIES = {
    "词汇丰富度",
    "词汇密度与词长",
    "句段结构",
    "词性",
    "语法标记",
    "复句关系",
    "熟语",
    "记叙描写",
    "HSK",
}

DIRECT_METRIC_MARKERS = (
    "TTR",
    "Guiraud",
    "MATTR",
    "占比",
    "密度",
    "覆盖率",
    "比例",
    "多样性",
    "平均",
    "中位数",
    "标准差",
    "最长",
    "实词率",
    "实虚词比",
)

DROP_EXACT = {
    "人称代词总每千字",
    "复句句次总每千字",
    "评价词总每千字",
    "HSK词汇每千字",
    "非HSK词汇每千字",
    "初等词汇每千字",
    "中等词汇每千字",
    "高等词汇每千字",
    "1级词汇每千字",
    "2级词汇每千字",
    "3级词汇每千字",
    "4级词汇每千字",
    "5级词汇每千字",
    "6级词汇每千字",
    "7-9级词汇每千字",
    "内容词每千字",
    "实词每千字",
    "虚词每千字",
    "其他词每千字",
    "单音节词每千字",
    "双音节词每千字",
    "三音节及以上词每千字",
    "单音节词占比",
    "初等词汇占比",
    "HSK词汇覆盖率",
}

PREFERRED_FEATURES = {
    "词形丰富度TTR",
    "Guiraud值",
    "MATTR-50",
    "仅出现一次词占比",
    "词汇密度",
    "实词率",
    "平均词长",
    "平均句长_字",
    "句长标准差_字",
    "超过30字长句占比",
    "平均每句分句数",
    "复句类型多样性",
    "第一人称代词占人称代词比例",
    "最长连续动词序列",
    "中等词汇占比",
    "高等词汇占比",
    "非HSK词汇占比",
    "高等词汇占HSK词汇比例",
}

THEME_KEYWORDS = {
    "词汇丰富度与词汇扩展": ("TTR", "Guiraud", "MATTR", "去重", "仅出现一次"),
    "词汇等级与词形复杂度": ("级词汇", "非HSK", "平均词长", "三音节及以上"),
    "句法延展与分句复杂度": ("句长", "长句", "分句", "逗号", "段落长度"),
    "人称指涉与叙述参与": ("第一人称", "第二人称", "第三人称", "代词"),
    "动作过程与动词链": ("动作动词", "趋向动词", "连续动词", "动词每千字", "情态动词"),
    "复句关系与逻辑组织": ("复句", "因果", "转折", "条件", "假设", "目的", "递进", "并列", "承接", "解说"),
    "时间叙述与体标记": ("时间", "时态助词", "了每千字", "着每千字", "过每千字"),
    "评价立场与副词修饰": ("评价词", "程度副词", "副词每千字", "否定词", "心理动词"),
    "信息密度与实词使用": ("词汇密度", "实词率", "名词每千字", "内容词", "形容词"),
    "口语互动与语气表达": ("语气词", "疑问代词", "直接引语", "叹词"),
}

CODE_LABELS = {
    "J1": "J1 记对我影响最大的一个人",
    "J2": "J2 我的一个假期",
    "Y1": "Y1 吸烟的影响",
    "Y2": "Y2 父母是孩子的第一任老师",
}

COLORS = {
    "J1": "#007C91",
    "J2": "#4AA8B8",
    "Y1": "#C4493D",
    "Y2": "#E6A23C",
    "accent": "#365B8C",
    "dark": "#253238",
    "muted": "#708090",
}


@dataclass
class FactorModelResult:
    k: int
    feature_names: list[str]
    removed_low_communality: list[dict[str, float]]
    loadings: np.ndarray
    communalities: np.ndarray
    uniquenesses: np.ndarray
    phi: np.ndarray
    factor_variance: np.ndarray
    kmo: float
    min_msa: float
    bartlett_chi2: float
    bartlett_df: int
    bartlett_p: float
    n_per_variable: float
    strong_assigned_counts: list[int]
    cross_loading_count: int
    heywood: bool
    diagnostic_valid: bool
    diagnostic_note: str
    analyzer: FactorAnalyzer
    stability_medians: list[float] | None = None
    stability_q05: list[float] | None = None
    bootstrap_successes: int = 0
    stable: bool = False


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="对作文词性统计宽表执行MF/MD多维分析。")
    parser.add_argument("--workbook", default=DEFAULT_WORKBOOK)
    parser.add_argument("--clean-text-dir", default=DEFAULT_CLEAN_TEXT_DIR)
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--seed", type=int, default=DEFAULT_SEED)
    parser.add_argument("--parallel-iterations", type=int, default=1000)
    parser.add_argument("--bootstrap", type=int, default=200)
    parser.add_argument("--zero-threshold", type=float, default=0.95)
    parser.add_argument("--corr-threshold", type=float, default=0.85)
    parser.add_argument("--msa-threshold", type=float, default=0.50)
    parser.add_argument("--communality-threshold", type=float, default=0.30)
    parser.add_argument("--loading-threshold", type=float, default=0.30)
    parser.add_argument("--max-factors", type=int, default=10)
    parser.add_argument("--skip-bootstrap", action="store_true")
    return parser.parse_args()


def patch_factor_analyzer_compatibility() -> None:
    signature = inspect.signature(factor_analyzer_module.check_array)
    if "force_all_finite" in signature.parameters:
        return
    if "ensure_all_finite" not in signature.parameters:
        return
    original = factor_analyzer_module.check_array

    def compatible_check_array(*args: Any, **kwargs: Any) -> Any:
        if "force_all_finite" in kwargs:
            kwargs["ensure_all_finite"] = kwargs.pop("force_all_finite")
        return original(*args, **kwargs)

    factor_analyzer_module.check_array = compatible_check_array


def configure_plotting() -> None:
    sns.set_theme(style="whitegrid", rc={"grid.color": "#E5EAED", "grid.linewidth": 0.7})
    candidates = (
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/System/Library/Fonts/Supplemental/Songti.ttc",
    )
    for path in candidates:
        if Path(path).exists():
            font_manager.fontManager.addfont(path)
            family = font_manager.FontProperties(fname=path).get_name()
            plt.rcParams["font.family"] = family
            plt.rcParams["font.sans-serif"] = [family]
            break
    plt.rcParams.update(
        {
            "axes.unicode_minus": False,
            "axes.edgecolor": "#AAB4BB",
            "axes.labelcolor": COLORS["dark"],
            "axes.titlecolor": COLORS["dark"],
            "xtick.color": COLORS["dark"],
            "ytick.color": COLORS["dark"],
            "figure.facecolor": "white",
            "savefig.facecolor": "white",
            "font.size": 10,
            "axes.titlesize": 12,
            "axes.labelsize": 10,
        }
    )


def json_value(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_value(item) for item in value]
    if isinstance(value, np.ndarray):
        return value.tolist()
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating,)):
        return None if not np.isfinite(value) else float(value)
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    if pd.isna(value):
        return None
    return value


def load_inputs(workbook_path: Path) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, str]]:
    data = pd.read_excel(workbook_path, sheet_name="词性统计", dtype={"作文编码": str})
    dictionary = pd.read_excel(workbook_path, sheet_name="字段说明")
    category_map = dict(zip(dictionary["字段名"], dictionary["类别"], strict=True))
    return data, dictionary, category_map


def validate_inputs(data: pd.DataFrame, clean_text_dir: Path) -> dict[str, Any]:
    errors: list[str] = []
    if data.shape[0] != EXPECTED_ROWS or data.shape[1] < MIN_EXPECTED_COLUMNS:
        errors.append(f"宽表应为{EXPECTED_ROWS}行且至少{MIN_EXPECTED_COLUMNS}列，实际为{data.shape}")
    if data.isna().any().any():
        missing = int(data.isna().sum().sum())
        errors.append(f"宽表存在{missing}个缺失值")
    code_counts = data["篇名代码"].value_counts().to_dict()
    for code in EXPECTED_CODES:
        if code_counts.get(code) != EXPECTED_CODE_COUNT:
            errors.append(f"{code}应有{EXPECTED_CODE_COUNT}篇，实际为{code_counts.get(code)}")
    scores = tuple(sorted(int(value) for value in data["作文分数"].unique()))
    if scores != EXPECTED_SCORES:
        errors.append(f"作文分数集合不符合预期：{scores}")
    score_tables = []
    for code in EXPECTED_CODES:
        counts = data.loc[data["篇名代码"] == code, "作文分数"].value_counts().sort_index()
        score_tables.append(counts)
    if any(not score_tables[0].equals(table) for table in score_tables[1:]):
        errors.append("四个篇名代码的分数分层不一致")
    missing_texts = []
    for row in data[["篇名代码", "作文文件名"]].itertuples(index=False):
        path = clean_text_dir / str(row.篇名代码) / f"{row.作文文件名}.txt"
        if not path.is_file():
            missing_texts.append(str(path))
    if missing_texts:
        errors.append(f"有{len(missing_texts)}篇清洗文本缺失")
    if errors:
        raise ValueError("；".join(errors))
    return {
        "rows": int(data.shape[0]),
        "columns": int(data.shape[1]),
        "code_counts": {key: int(value) for key, value in code_counts.items()},
        "scores": list(scores),
        "score_counts_per_code": {
            str(score): int(count) for score, count in score_tables[0].items()
        },
        "missing_values": 0,
        "mapped_clean_texts": len(data),
    }


def is_direct_metric(name: str) -> bool:
    return name.endswith("每千字") or any(marker in name for marker in DIRECT_METRIC_MARKERS)


def feature_priority(name: str) -> tuple[int, int]:
    score = 0
    if name in PREFERRED_FEATURES:
        score += 8
    if "平均句长" in name or "平均每句" in name:
        score += 4
    if name.endswith("每千字"):
        score += 2
    if "种类每千字" in name:
        score -= 3
    if "内部TTR" in name:
        score -= 2
    if name in {"句子每千字", "段落每千字"}:
        score -= 2
    return score, -len(name)


def kmo_and_msa(matrix: np.ndarray) -> tuple[float, np.ndarray, np.ndarray]:
    correlation = np.corrcoef(matrix, rowvar=False)
    inverse = np.linalg.pinv(correlation)
    diagonal = np.sqrt(np.clip(np.diag(inverse), 1e-12, None))
    partial = -inverse / np.outer(diagonal, diagonal)
    np.fill_diagonal(partial, 0.0)
    correlation_no_diag = correlation.copy()
    np.fill_diagonal(correlation_no_diag, 0.0)
    r2 = (correlation_no_diag**2).sum(axis=0)
    p2 = (partial**2).sum(axis=0)
    denominator = r2 + p2
    msa = np.divide(r2, denominator, out=np.zeros_like(r2), where=denominator > 0)
    overall = float(r2.sum() / (r2.sum() + p2.sum()))
    return overall, msa, correlation


def bartlett_sphericity(correlation: np.ndarray, sample_size: int) -> tuple[float, int, float]:
    variables = correlation.shape[0]
    sign, log_determinant = np.linalg.slogdet(correlation)
    if sign <= 0:
        eigenvalues = np.linalg.eigvalsh(correlation)
        log_determinant = float(np.log(np.clip(eigenvalues, 1e-12, None)).sum())
    chi2 = -(sample_size - 1 - (2 * variables + 5) / 6) * log_determinant
    degrees = variables * (variables - 1) // 2
    p_value = float(stats.chi2.sf(chi2, degrees))
    return float(chi2), int(degrees), p_value


def screen_features(
    data: pd.DataFrame,
    category_map: dict[str, str],
    *,
    zero_threshold: float,
    corr_threshold: float,
    msa_threshold: float,
) -> tuple[pd.DataFrame, list[str], np.ndarray, dict[str, Any]]:
    records: dict[str, dict[str, Any]] = {}
    candidates: list[str] = []
    for index, name in enumerate(data.columns):
        category = category_map.get(name, "")
        if category not in FEATURE_CATEGORIES or not is_direct_metric(name) or name in DROP_EXACT:
            continue
        candidates.append(name)
        records[name] = {
            "字段名": name,
            "类别": category,
            "初始顺序": index + 1,
            "筛选状态": "候选",
            "筛选原因": "",
            "零值比例": float((data[name] == 0).mean()),
            "标准差": float(data[name].std(ddof=1)),
            "相关变量": "",
            "相关系数": None,
            "MSA": None,
            "共同度": None,
        }

    active: list[str] = []
    for name in candidates:
        record = records[name]
        if record["标准差"] <= 1e-12:
            record["筛选状态"] = "删除"
            record["筛选原因"] = "零方差"
        elif record["零值比例"] >= zero_threshold:
            record["筛选状态"] = "删除"
            record["筛选原因"] = f"零值比例≥{zero_threshold:.0%}"
        else:
            active.append(name)

    correlation = data[active].corr(method="pearson")
    order = sorted(active, key=lambda name: feature_priority(name), reverse=True)
    correlation_kept: list[str] = []
    for name in order:
        conflicts = [
            kept for kept in correlation_kept if abs(float(correlation.loc[name, kept])) >= corr_threshold
        ]
        if not conflicts:
            correlation_kept.append(name)
            continue
        strongest = max(conflicts, key=lambda kept: abs(float(correlation.loc[name, kept])))
        records[name]["筛选状态"] = "删除"
        records[name]["筛选原因"] = f"与保留变量绝对相关系数≥{corr_threshold:.2f}"
        records[name]["相关变量"] = strongest
        records[name]["相关系数"] = float(correlation.loc[name, strongest])

    original_order = {name: index for index, name in enumerate(candidates)}
    active = sorted(correlation_kept, key=original_order.__getitem__)
    matrix = data[active].to_numpy(dtype=float)
    msa_removed: list[str] = []
    while len(active) > 18:
        _, msa, _ = kmo_and_msa(matrix)
        min_index = int(np.argmin(msa))
        for index, name in enumerate(active):
            records[name]["MSA"] = float(msa[index])
        if float(msa[min_index]) >= msa_threshold:
            break
        name = active.pop(min_index)
        msa_removed.append(name)
        records[name]["筛选状态"] = "删除"
        records[name]["筛选原因"] = f"变量级MSA<{msa_threshold:.2f}"
        matrix = np.delete(matrix, min_index, axis=1)

    overall_kmo, final_msa, final_correlation = kmo_and_msa(matrix)
    for index, name in enumerate(active):
        records[name]["MSA"] = float(final_msa[index])
        records[name]["筛选状态"] = "进入候选模型"
        records[name]["筛选原因"] = "通过稀疏、共线和MSA筛选"

    screening = pd.DataFrame([records[name] for name in candidates])
    diagnostics = {
        "initial_candidates": len(candidates),
        "after_variance_sparsity": int(
            sum(record["筛选原因"] not in {"零方差", f"零值比例≥{zero_threshold:.0%}"} for record in records.values())
        ),
        "after_correlation": len(correlation_kept),
        "after_msa": len(active),
        "overall_kmo": overall_kmo,
        "minimum_msa": float(final_msa.min()),
        "msa_removed": msa_removed,
    }
    return screening, active, matrix, diagnostics


def parallel_analysis(
    matrix: np.ndarray, iterations: int, seed: int
) -> tuple[np.ndarray, np.ndarray, int]:
    correlation = np.corrcoef(matrix, rowvar=False)
    observed = np.linalg.eigvalsh(correlation)[::-1]
    rng = np.random.default_rng(seed)
    random_eigenvalues = np.empty((iterations, matrix.shape[1]), dtype=float)
    for index in range(iterations):
        random_matrix = rng.normal(size=matrix.shape)
        random_correlation = np.corrcoef(random_matrix, rowvar=False)
        random_eigenvalues[index] = np.linalg.eigvalsh(random_correlation)[::-1]
    threshold = np.quantile(random_eigenvalues, 0.95, axis=0)
    count = int(np.sum(observed > threshold))
    return observed, threshold, count


def fit_factor_model(
    matrix: np.ndarray,
    feature_names: Sequence[str],
    factor_count: int,
    communality_threshold: float,
    loading_threshold: float,
) -> FactorModelResult:
    working_matrix = matrix.copy()
    working_names = list(feature_names)
    removed: list[dict[str, float]] = []
    analyzer: FactorAnalyzer | None = None

    for _ in range(len(feature_names)):
        if len(working_names) < factor_count * 3:
            break
        analyzer = FactorAnalyzer(n_factors=factor_count, method="minres", rotation="promax")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            analyzer.fit(working_matrix)
        communalities = analyzer.get_communalities()
        low = np.flatnonzero(communalities < communality_threshold)
        if len(low) == 0:
            break
        removable = len(working_names) - factor_count * 3
        if removable <= 0:
            break
        batch_size = 1 if len(low) <= 8 else min(max(2, len(low) // 3), removable)
        remove_indices = sorted(low, key=lambda index: communalities[index])[:batch_size]
        for index in sorted(remove_indices, reverse=True):
            removed.append({"字段名": working_names[index], "共同度": float(communalities[index])})
            working_names.pop(index)
            working_matrix = np.delete(working_matrix, index, axis=1)

    analyzer = FactorAnalyzer(n_factors=factor_count, method="minres", rotation="promax")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        analyzer.fit(working_matrix)
    loadings = np.asarray(analyzer.loadings_, dtype=float)
    communalities = np.asarray(analyzer.get_communalities(), dtype=float)
    uniquenesses = np.asarray(analyzer.get_uniquenesses(), dtype=float)
    phi = np.asarray(getattr(analyzer, "phi_", np.eye(factor_count)), dtype=float)
    if phi.shape != (factor_count, factor_count):
        phi = np.eye(factor_count)
    factor_variance = np.asarray(analyzer.get_factor_variance(), dtype=float)
    kmo, msa, correlation = kmo_and_msa(working_matrix)
    bartlett_chi2, bartlett_df, bartlett_p = bartlett_sphericity(
        correlation, working_matrix.shape[0]
    )
    absolute = np.abs(loadings)
    primary = np.argmax(absolute, axis=1)
    maximum = absolute.max(axis=1)
    assigned_counts = [
        int(np.sum((primary == factor) & (maximum >= loading_threshold)))
        for factor in range(factor_count)
    ]
    cross_loading_count = int(np.sum((absolute >= loading_threshold).sum(axis=1) >= 2))
    heywood = bool(np.any(communalities > 1.01) or np.any(uniquenesses < -0.01))
    reasons: list[str] = []
    if kmo < 0.60:
        reasons.append("KMO<0.60")
    if bartlett_p >= 0.05:
        reasons.append("Bartlett检验不显著")
    if working_matrix.shape[0] / working_matrix.shape[1] < 5:
        reasons.append("样本变量比<5")
    if min(assigned_counts, default=0) < 3:
        reasons.append("存在少于3个主要载荷的因子")
    if heywood:
        reasons.append("出现Heywood异常")
    diagnostic_valid = not reasons
    note = "通过统计诊断" if diagnostic_valid else "；".join(reasons)
    return FactorModelResult(
        k=factor_count,
        feature_names=working_names,
        removed_low_communality=removed,
        loadings=loadings,
        communalities=communalities,
        uniquenesses=uniquenesses,
        phi=phi,
        factor_variance=factor_variance,
        kmo=kmo,
        min_msa=float(msa.min()),
        bartlett_chi2=bartlett_chi2,
        bartlett_df=bartlett_df,
        bartlett_p=bartlett_p,
        n_per_variable=working_matrix.shape[0] / working_matrix.shape[1],
        strong_assigned_counts=assigned_counts,
        cross_loading_count=cross_loading_count,
        heywood=heywood,
        diagnostic_valid=diagnostic_valid,
        diagnostic_note=note,
        analyzer=analyzer,
    )


def tucker_congruence(reference: np.ndarray, candidate: np.ndarray) -> np.ndarray:
    result = np.zeros((reference.shape[1], candidate.shape[1]), dtype=float)
    for row in range(reference.shape[1]):
        for column in range(candidate.shape[1]):
            numerator = float(np.dot(reference[:, row], candidate[:, column]))
            denominator = math.sqrt(
                float(np.dot(reference[:, row], reference[:, row]))
                * float(np.dot(candidate[:, column], candidate[:, column]))
            )
            result[row, column] = numerator / denominator if denominator else 0.0
    return result


def bootstrap_stability(
    matrix: np.ndarray,
    model: FactorModelResult,
    iterations: int,
    seed: int,
) -> tuple[list[float], list[float], int]:
    if iterations <= 0:
        return [1.0] * model.k, [1.0] * model.k, 0
    rng = np.random.default_rng(seed + model.k * 1000)
    reference = model.loadings
    values: list[list[float]] = [[] for _ in range(model.k)]
    successes = 0
    for _ in range(iterations):
        sample = rng.integers(0, matrix.shape[0], size=matrix.shape[0])
        sampled = matrix[sample]
        try:
            analyzer = FactorAnalyzer(n_factors=model.k, method="minres", rotation="promax")
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                analyzer.fit(sampled)
            candidate = np.asarray(analyzer.loadings_, dtype=float)
            congruence = tucker_congruence(reference, candidate)
            rows, columns = linear_sum_assignment(-np.abs(congruence))
            aligned = np.zeros(model.k, dtype=float)
            for row, column in zip(rows, columns, strict=True):
                aligned[row] = abs(congruence[row, column])
            if np.all(np.isfinite(aligned)):
                for factor, value in enumerate(aligned):
                    values[factor].append(float(value))
                successes += 1
        except Exception:
            continue
    medians = [float(np.median(item)) if item else 0.0 for item in values]
    q05 = [float(np.quantile(item, 0.05)) if item else 0.0 for item in values]
    return medians, q05, successes


def choose_model(models: list[FactorModelResult]) -> tuple[FactorModelResult | None, str]:
    stable = [model for model in models if model.diagnostic_valid and model.stable]
    if stable:
        return max(stable, key=lambda model: model.k), "选取通过诊断与Bootstrap稳定性检验的最高维度解"
    diagnostic = [model for model in models if model.diagnostic_valid]
    if diagnostic:
        best = max(diagnostic, key=lambda model: float(np.median(model.stability_medians or [0])))
        return best, "未达到预设稳定性阈值，保留最佳诊断模型并标记为探索性"
    return None, "没有候选模型通过统计诊断"


def name_factors(model: FactorModelResult) -> tuple[list[str], list[dict[str, Any]]]:
    labels: list[str] = []
    details: list[dict[str, Any]] = []
    used: set[str] = set()
    for factor in range(model.k):
        loadings = model.loadings[:, factor]
        order = np.argsort(-np.abs(loadings))[:12]
        theme_scores: dict[str, float] = {theme: 0.0 for theme in THEME_KEYWORDS}
        for index in order:
            name = model.feature_names[index]
            weight = abs(float(loadings[index]))
            for theme, keywords in THEME_KEYWORDS.items():
                if any(keyword in name for keyword in keywords):
                    theme_scores[theme] += weight
        ranked_themes = sorted(theme_scores, key=theme_scores.get, reverse=True)
        label = ranked_themes[0]
        if label in used and len(ranked_themes) > 1:
            label = f"{label}与{ranked_themes[1]}"
        if label in used:
            label = f"{label}（维度{factor + 1}）"
        used.add(label)
        labels.append(label)
        positive = [
            {"字段名": model.feature_names[index], "载荷": float(loadings[index])}
            for index in np.argsort(-loadings)
            if loadings[index] >= 0.30
        ][:6]
        negative = [
            {"字段名": model.feature_names[index], "载荷": float(loadings[index])}
            for index in np.argsort(loadings)
            if loadings[index] <= -0.30
        ][:6]
        details.append(
            {
                "维度序号": factor + 1,
                "维度名称": label,
                "正向主要指标": positive,
                "负向主要指标": negative,
            }
        )
    if model.k == 5:
        top_sets = []
        for factor in range(model.k):
            order = np.argsort(-np.abs(model.loadings[:, factor]))[:6]
            top_sets.append({model.feature_names[index] for index in order})
        expected_signals = (
            "词形丰富度TTR" in top_sets[0],
            "中等词汇占比" in top_sets[1],
            "代词每千字" in top_sets[2],
            "平均每句分句数" in top_sets[3],
            "动词每千字" in top_sets[4],
        )
        if all(expected_signals):
            labels = [
                "词汇丰富度与词汇扩展",
                "基础词汇与叙事推进",
                "人称指涉与信息密度",
                "句法延展与分句复杂度",
                "动作过程与动词链",
            ]
            for detail, label in zip(details, labels, strict=True):
                detail["维度名称"] = label
    return labels, details


def standardize(values: np.ndarray) -> np.ndarray:
    mean = np.mean(values, axis=0)
    standard_deviation = np.std(values, axis=0, ddof=1)
    standard_deviation = np.where(standard_deviation <= 1e-12, 1.0, standard_deviation)
    return (values - mean) / standard_deviation


def calculate_dimension_scores(
    data: pd.DataFrame,
    model: FactorModelResult,
    labels: Sequence[str],
    loading_threshold: float,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    matrix = data[model.feature_names].to_numpy(dtype=float)
    z_matrix = standardize(matrix)
    absolute = np.abs(model.loadings)
    primary = np.argmax(absolute, axis=1)
    maximum = absolute.max(axis=1)
    score_data: dict[str, np.ndarray] = {}
    assignment_records: list[dict[str, Any]] = []
    for factor, label in enumerate(labels):
        indices = np.flatnonzero((primary == factor) & (maximum >= loading_threshold))
        signs = np.sign(model.loadings[indices, factor])
        raw = np.sum(z_matrix[:, indices] * signs, axis=1)
        score_data[label] = standardize(raw.reshape(-1, 1)).ravel()
        for index in indices:
            assignment_records.append(
                {
                    "字段名": model.feature_names[index],
                    "所属维度": label,
                    "主要载荷": float(model.loadings[index, factor]),
                    "共同度": float(model.communalities[index]),
                    "正负极": "正向" if model.loadings[index, factor] > 0 else "负向",
                }
            )

    regression_scores = standardize(np.asarray(model.analyzer.transform(matrix), dtype=float))
    sensitivity = pd.DataFrame(
        {
            "维度名称": labels,
            "符号求和与回归得分相关": [
                float(np.corrcoef(score_data[label], regression_scores[:, factor])[0, 1])
                for factor, label in enumerate(labels)
            ],
        }
    )
    metadata_columns = [
        "篇名代码",
        "篇名",
        "作文编码",
        "国籍",
        "作文题目",
        "作文分数",
        "体裁",
        "作文文件名",
    ]
    scores = data[metadata_columns].copy()
    for factor, label in enumerate(labels):
        scores[label] = score_data[label]
        scores[f"{label}_回归得分"] = regression_scores[:, factor]
    return scores, pd.DataFrame(assignment_records), sensitivity


def omega_squared(groups: Sequence[np.ndarray]) -> float:
    values = np.concatenate(groups)
    grand = float(np.mean(values))
    ss_between = sum(len(group) * (float(np.mean(group)) - grand) ** 2 for group in groups)
    ss_within = sum(float(np.sum((group - np.mean(group)) ** 2)) for group in groups)
    df_between = len(groups) - 1
    df_within = len(values) - len(groups)
    ms_within = ss_within / df_within if df_within else 0.0
    denominator = ss_between + ss_within + ms_within
    return max(0.0, (ss_between - df_between * ms_within) / denominator) if denominator else 0.0


def hedges_g(first: np.ndarray, second: np.ndarray) -> float:
    n1, n2 = len(first), len(second)
    variance = ((n1 - 1) * np.var(first, ddof=1) + (n2 - 1) * np.var(second, ddof=1))
    variance /= max(1, n1 + n2 - 2)
    if variance <= 0:
        return 0.0
    d = (float(np.mean(first)) - float(np.mean(second))) / math.sqrt(variance)
    correction = 1 - 3 / max(1, 4 * (n1 + n2) - 9)
    return d * correction


def group_comparisons(
    scores: pd.DataFrame, labels: Sequence[str]
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    omnibus_records: list[dict[str, Any]] = []
    pairwise_records: list[dict[str, Any]] = []
    descriptive_records: list[dict[str, Any]] = []
    for label in labels:
        groups = [
            scores.loc[scores["篇名代码"] == code, label].to_numpy(dtype=float)
            for code in EXPECTED_CODES
        ]
        result = anova_oneway(groups, use_var="unequal", welch_correction=True)
        omnibus_records.append(
            {
                "维度名称": label,
                "Welch_F": float(result.statistic),
                "分子自由度": float(result.df_num),
                "分母自由度": float(result.df_denom),
                "p值": float(result.pvalue),
                "Omega平方": omega_squared(groups),
            }
        )
        for code, values in zip(EXPECTED_CODES, groups, strict=True):
            descriptive_records.append(
                {
                    "维度名称": label,
                    "篇名代码": code,
                    "样本量": len(values),
                    "均值": float(np.mean(values)),
                    "标准差": float(np.std(values, ddof=1)),
                    "中位数": float(np.median(values)),
                    "四分位距": float(np.quantile(values, 0.75) - np.quantile(values, 0.25)),
                }
            )
        local: list[dict[str, Any]] = []
        for first_code, second_code in combinations(EXPECTED_CODES, 2):
            first = scores.loc[scores["篇名代码"] == first_code, label].to_numpy(dtype=float)
            second = scores.loc[scores["篇名代码"] == second_code, label].to_numpy(dtype=float)
            test = stats.ttest_ind(first, second, equal_var=False)
            local.append(
                {
                    "维度名称": label,
                    "组1": first_code,
                    "组2": second_code,
                    "均值差_组1减组2": float(np.mean(first) - np.mean(second)),
                    "Welch_t": float(test.statistic),
                    "自由度": float(test.df),
                    "p值": float(test.pvalue),
                    "Hedges_g": hedges_g(first, second),
                }
            )
        adjusted = multipletests([record["p值"] for record in local], method="holm")[1]
        for record, p_adjusted in zip(local, adjusted, strict=True):
            record["Holm校正p值"] = float(p_adjusted)
            pairwise_records.append(record)

    omnibus = pd.DataFrame(omnibus_records)
    omnibus["BH校正p值"] = multipletests(omnibus["p值"], method="fdr_bh")[1]
    return omnibus, pd.DataFrame(pairwise_records), pd.DataFrame(descriptive_records)


def score_and_regression_analysis(
    scores: pd.DataFrame, labels: Sequence[str]
) -> tuple[pd.DataFrame, pd.DataFrame]:
    model_data = scores.copy()
    model_data["score_z"] = standardize(model_data[["作文分数"]].to_numpy(dtype=float)).ravel()
    model_data["code"] = model_data["篇名代码"].astype(str)
    model_data["is_japan"] = (model_data["国籍"] == "日本").astype(int)
    correlation_records: list[dict[str, Any]] = []
    regression_records: list[dict[str, Any]] = []
    for index, label in enumerate(labels):
        rho, p_value = stats.spearmanr(model_data["作文分数"], model_data[label])
        correlation_records.append(
            {"维度名称": label, "Spearman_rho": float(rho), "p值": float(p_value)}
        )
        local = model_data[[label, "score_z", "code", "is_japan"]].rename(columns={label: "outcome"})
        result = smf.ols("outcome ~ score_z + C(code) + is_japan", data=local).fit(cov_type="HC3")
        confidence = result.conf_int(alpha=0.05)
        for term in result.params.index:
            regression_records.append(
                {
                    "维度名称": label,
                    "模型项": term,
                    "系数": float(result.params[term]),
                    "稳健标准误": float(result.bse[term]),
                    "t值": float(result.tvalues[term]),
                    "p值": float(result.pvalues[term]),
                    "95%CI下限": float(confidence.loc[term, 0]),
                    "95%CI上限": float(confidence.loc[term, 1]),
                    "样本量": int(result.nobs),
                    "调整R平方": float(result.rsquared_adj),
                }
            )
    correlations = pd.DataFrame(correlation_records)
    correlations["BH校正p值"] = multipletests(correlations["p值"], method="fdr_bh")[1]
    regressions = pd.DataFrame(regression_records)
    for term in ("score_z", "is_japan"):
        mask = regressions["模型项"] == term
        if mask.any():
            regressions.loc[mask, "BH校正p值"] = multipletests(
                regressions.loc[mask, "p值"], method="fdr_bh"
            )[1]
    return correlations, regressions


def sample_tables(data: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    summary = (
        data.groupby(["篇名代码", "篇名", "体裁"], as_index=False)
        .agg(
            样本量=("作文文件名", "count"),
            平均分=("作文分数", "mean"),
            分数标准差=("作文分数", "std"),
            平均纯文本字数=("纯文本字数", "mean"),
            纯文本字数中位数=("纯文本字数", "median"),
            国籍数=("国籍", "nunique"),
        )
    )
    score_distribution = (
        data.groupby(["作文分数", "篇名代码"]).size().unstack(fill_value=0).reset_index()
    )
    nationality = (
        data.groupby(["篇名代码", "国籍"]).size().rename("样本量").reset_index()
        .sort_values(["篇名代码", "样本量"], ascending=[True, False])
    )
    return summary, score_distribution, nationality


def select_excerpt(text: str, target_length: int = 130) -> str:
    compact = re.sub(r"\s+", "", text.strip())
    if not compact:
        return ""
    sentences = [item.strip() for item in re.split(r"(?<=[。！？!?])", compact) if item.strip()]
    eligible = [item for item in sentences if 45 <= len(item) <= 180]
    if eligible:
        excerpt = min(eligible, key=lambda item: abs(len(item) - target_length))
    else:
        excerpt = compact[:target_length]
    if len(excerpt) > 160:
        excerpt = excerpt[:158] + "……"
    return excerpt


def representative_examples(
    scores: pd.DataFrame,
    labels: Sequence[str],
    clean_text_dir: Path,
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    used: set[str] = set()
    for label in labels:
        for pole, quantile in (("低端", 0.10), ("高端", 0.90)):
            target = float(scores[label].quantile(quantile))
            candidates = scores.assign(distance=(scores[label] - target).abs()).sort_values("distance")
            selected = None
            for row in candidates.itertuples(index=False):
                if row.作文文件名 not in used:
                    selected = row
                    break
            if selected is None:
                selected = candidates.iloc[0]
                filename = str(selected["作文文件名"])
                code = str(selected["篇名代码"])
                score = int(selected["作文分数"])
                dimension_score = float(selected[label])
            else:
                filename = str(selected.作文文件名)
                code = str(selected.篇名代码)
                score = int(selected.作文分数)
                dimension_score = float(getattr(selected, label))
            used.add(filename)
            text = (clean_text_dir / code / f"{filename}.txt").read_text(encoding="utf-8")
            records.append(
                {
                    "维度名称": label,
                    "维度位置": pole,
                    "作文文件名": filename,
                    "篇名代码": code,
                    "作文分数": score,
                    "维度得分": dimension_score,
                    "作文片段": select_excerpt(text),
                }
            )
    return records


def save_figure(path: Path) -> None:
    plt.tight_layout()
    plt.savefig(path, dpi=300, bbox_inches="tight")
    plt.close()


def create_figures(
    data: pd.DataFrame,
    screening_diagnostics: dict[str, Any],
    observed_eigenvalues: np.ndarray,
    parallel_thresholds: np.ndarray,
    selected_model: FactorModelResult,
    labels: Sequence[str],
    scores: pd.DataFrame,
    pairwise: pd.DataFrame,
    regressions: pd.DataFrame,
    figures_dir: Path,
) -> pd.DataFrame:
    figures_dir.mkdir(parents=True, exist_ok=True)
    records: list[dict[str, Any]] = []

    fig, axes = plt.subplots(1, 2, figsize=(12, 4.5))
    code_counts = data["篇名代码"].value_counts().reindex(EXPECTED_CODES)
    axes[0].bar(code_counts.index, code_counts.values, color=[COLORS[code] for code in EXPECTED_CODES])
    axes[0].set_title("四个题目组样本量")
    axes[0].set_ylabel("作文篇数")
    axes[0].set_ylim(0, max(code_counts.values) * 1.18)
    for index, value in enumerate(code_counts.values):
        axes[0].text(index, value + 3, str(value), ha="center", color=COLORS["dark"])
    score_matrix = data.groupby(["篇名代码", "作文分数"]).size().unstack(fill_value=0)
    sns.heatmap(score_matrix, annot=True, fmt="d", cmap="YlGnBu", cbar=False, ax=axes[1])
    axes[1].set_title("各组分数分层（75分已排除）")
    axes[1].set_xlabel("作文分数")
    axes[1].set_ylabel("篇名代码")
    path = figures_dir / "01_样本与分数结构.png"
    save_figure(path)
    records.append({"序号": 1, "图名": "样本与分数结构", "文件": str(path)})

    stages = [
        ("初始候选", screening_diagnostics["initial_candidates"]),
        ("方差/稀疏筛选后", screening_diagnostics["after_variance_sparsity"]),
        ("共线筛选后", screening_diagnostics["after_correlation"]),
        ("MSA筛选后", screening_diagnostics["after_msa"]),
        ("最终模型", len(selected_model.feature_names)),
    ]
    fig, ax = plt.subplots(figsize=(10, 3.8))
    x = np.arange(len(stages))
    values = [value for _, value in stages]
    ax.plot(x, values, marker="o", linewidth=2.5, color=COLORS["accent"])
    ax.fill_between(x, values, alpha=0.12, color=COLORS["accent"])
    ax.set_xticks(x, [label for label, _ in stages])
    ax.set_ylabel("变量数")
    ax.set_title("MF/MD变量筛选流程")
    for index, value in enumerate(values):
        ax.text(index, value + max(values) * 0.035, str(value), ha="center")
    path = figures_dir / "02_变量筛选流程.png"
    save_figure(path)
    records.append({"序号": 2, "图名": "变量筛选流程", "文件": str(path)})

    fig, ax = plt.subplots(figsize=(9, 5))
    count = min(25, len(observed_eigenvalues))
    factors = np.arange(1, count + 1)
    ax.plot(factors, observed_eigenvalues[:count], marker="o", label="实测特征根", color=COLORS["accent"])
    ax.plot(factors, parallel_thresholds[:count], marker="s", label="随机数据95%阈值", color=COLORS["Y1"])
    ax.axvline(selected_model.k, color=COLORS["J1"], linestyle="--", label=f"选定{selected_model.k}维")
    ax.axhline(1, color="#8D99A6", linestyle=":", linewidth=1)
    ax.set_xlabel("因子序号")
    ax.set_ylabel("特征根")
    ax.set_title("碎石图与Horn平行分析")
    ax.legend(frameon=False)
    path = figures_dir / "03_平行分析与碎石图.png"
    save_figure(path)
    records.append({"序号": 3, "图名": "平行分析与碎石图", "文件": str(path)})

    loading_frame = pd.DataFrame(
        selected_model.loadings,
        index=selected_model.feature_names,
        columns=[f"D{index + 1} {label}" for index, label in enumerate(labels)],
    )
    primary = np.argmax(np.abs(loading_frame.to_numpy()), axis=1)
    magnitude = np.max(np.abs(loading_frame.to_numpy()), axis=1)
    order = np.lexsort((-magnitude, primary))
    loading_frame = loading_frame.iloc[order]
    fig_height = max(9, 0.24 * len(loading_frame))
    plt.figure(figsize=(11, fig_height))
    sns.heatmap(loading_frame, cmap="vlag", center=0, vmin=-1, vmax=1, linewidths=0.25, linecolor="#F3F5F6")
    plt.title("Promax旋转后的因子载荷矩阵")
    plt.xlabel("语言维度")
    plt.ylabel("保留语言指标")
    path = figures_dir / "04_因子载荷热图.png"
    save_figure(path)
    records.append({"序号": 4, "图名": "因子载荷热图", "文件": str(path)})

    columns = 2
    rows = math.ceil(len(labels) / columns)
    fig, axes = plt.subplots(rows, columns, figsize=(12, 3.5 * rows), squeeze=False)
    for factor, label in enumerate(labels):
        ax = axes[factor // columns][factor % columns]
        sns.boxplot(
            data=scores,
            x="篇名代码",
            y=label,
            order=EXPECTED_CODES,
            palette=[COLORS[code] for code in EXPECTED_CODES],
            hue="篇名代码",
            legend=False,
            showfliers=False,
            ax=ax,
        )
        sns.stripplot(
            data=scores.sample(frac=0.28, random_state=DEFAULT_SEED),
            x="篇名代码",
            y=label,
            order=EXPECTED_CODES,
            jitter=False,
            color="#29343A",
            alpha=0.25,
            size=2,
            ax=ax,
        )
        ax.axhline(0, color="#AAB4BB", linewidth=0.8)
        ax.set_title(f"D{factor + 1} {label}")
        ax.set_xlabel("")
        ax.set_ylabel("标准化维度得分")
    for empty in range(len(labels), rows * columns):
        axes[empty // columns][empty % columns].axis("off")
    path = figures_dir / "05_四组维度得分分布.png"
    save_figure(path)
    records.append({"序号": 5, "图名": "四组维度得分分布", "文件": str(path)})

    profile = scores.groupby("篇名代码")[list(labels)].mean().reindex(EXPECTED_CODES)
    plt.figure(figsize=(max(9, 1.5 * len(labels)), 4.5))
    sns.heatmap(profile, annot=True, fmt=".2f", cmap="vlag", center=0, linewidths=0.8, cbar_kws={"label": "组均值"})
    plt.title("J1/J2/Y1/Y2多维语言画像")
    plt.xlabel("语言维度")
    plt.ylabel("篇名代码")
    path = figures_dir / "06_四组多维画像热图.png"
    save_figure(path)
    records.append({"序号": 6, "图名": "四组多维画像热图", "文件": str(path)})

    fig, axes = plt.subplots(rows, columns, figsize=(12, 3.6 * rows), squeeze=False)
    mean_scores = scores.groupby(["作文分数", "篇名代码"])[list(labels)].mean().reset_index()
    for factor, label in enumerate(labels):
        ax = axes[factor // columns][factor % columns]
        for code in EXPECTED_CODES:
            subset = mean_scores[mean_scores["篇名代码"] == code]
            ax.plot(subset["作文分数"], subset[label], marker="o", label=code, color=COLORS[code])
        ax.axhline(0, color="#AAB4BB", linewidth=0.8)
        ax.set_title(f"D{factor + 1} {label}")
        ax.set_xlabel("作文分数")
        ax.set_ylabel("组内均值")
    for empty in range(len(labels), rows * columns):
        axes[empty // columns][empty % columns].axis("off")
    handles, legend_labels = axes[0][0].get_legend_handles_labels()
    fig.legend(handles, legend_labels, loc="upper center", ncol=4, frameon=False)
    fig.subplots_adjust(top=0.94)
    path = figures_dir / "07_分数与维度得分趋势.png"
    save_figure(path)
    records.append({"序号": 7, "图名": "分数与维度得分趋势", "文件": str(path)})

    hsk_columns = ["初等词汇占比", "中等词汇占比", "高等词汇占比", "非HSK词汇占比"]
    hsk_profile = data.groupby("篇名代码")[hsk_columns].mean().reindex(EXPECTED_CODES)
    fig, ax = plt.subplots(figsize=(10, 5))
    left = np.zeros(len(hsk_profile))
    hsk_colors = ["#B13A3A", "#3F8732", "#3E9BD1", "#E6BE32"]
    for column, color in zip(hsk_columns, hsk_colors, strict=True):
        values = hsk_profile[column].to_numpy()
        ax.barh(hsk_profile.index, values, left=left, label=column.replace("词汇占比", ""), color=color)
        left += values
    ax.set_xlim(0, 1)
    ax.xaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1))
    ax.set_xlabel("全部非标点分词中的比例")
    ax.set_title("四组作文的HSK词汇等级构成")
    ax.legend(ncol=4, loc="lower center", bbox_to_anchor=(0.5, -0.24), frameon=False)
    path = figures_dir / "08_HSK等级构成.png"
    save_figure(path)
    records.append({"序号": 8, "图名": "HSK等级构成", "文件": str(path)})

    effect_data = pairwise.assign(abs_g=pairwise["Hedges_g"].abs()).nlargest(24, "abs_g").copy()
    effect_data["比较"] = effect_data["维度名称"] + " | " + effect_data["组1"] + "-" + effect_data["组2"]
    effect_data = effect_data.sort_values("Hedges_g")
    fig, ax = plt.subplots(figsize=(10, max(6, 0.32 * len(effect_data))))
    colors = [COLORS["J1"] if value >= 0 else COLORS["Y1"] for value in effect_data["Hedges_g"]]
    ax.scatter(effect_data["Hedges_g"], effect_data["比较"], c=colors, s=42)
    ax.axvline(0, color="#88949C", linewidth=1)
    ax.axvline(0.5, color="#C5CDD2", linewidth=0.8, linestyle=":")
    ax.axvline(-0.5, color="#C5CDD2", linewidth=0.8, linestyle=":")
    ax.set_xlabel("Hedges g（正值表示组1更高）")
    ax.set_ylabel("")
    ax.set_title("四组差异较大的维度比较")
    path = figures_dir / "09_组间效应量森林图.png"
    save_figure(path)
    records.append({"序号": 9, "图名": "组间效应量森林图", "文件": str(path)})

    regression_plot = regressions[regressions["模型项"].isin(["score_z", "is_japan"])].copy()
    regression_plot["预测变量"] = regression_plot["模型项"].map(
        {"score_z": "作文分数（标准化）", "is_japan": "日本籍（控制题目与分数）"}
    )
    regression_plot["标签"] = regression_plot["维度名称"] + " | " + regression_plot["预测变量"]
    regression_plot = regression_plot.sort_values("系数")
    fig, ax = plt.subplots(figsize=(10, max(6, 0.34 * len(regression_plot))))
    xerr = np.vstack(
        [
            regression_plot["系数"] - regression_plot["95%CI下限"],
            regression_plot["95%CI上限"] - regression_plot["系数"],
        ]
    )
    color_values = [COLORS["accent"] if item == "score_z" else COLORS["Y1"] for item in regression_plot["模型项"]]
    ax.errorbar(
        regression_plot["系数"],
        np.arange(len(regression_plot)),
        xerr=xerr,
        fmt="none",
        ecolor="#8B979E",
        capsize=3,
    )
    ax.scatter(regression_plot["系数"], np.arange(len(regression_plot)), c=color_values, s=42)
    ax.set_yticks(np.arange(len(regression_plot)), regression_plot["标签"])
    ax.axvline(0, color="#88949C", linewidth=1)
    ax.set_xlabel("标准化回归系数及95%置信区间")
    ax.set_title("作文分数与日本籍的调整后维度效应")
    path = figures_dir / "10_回归系数森林图.png"
    save_figure(path)
    records.append({"序号": 10, "图名": "回归系数森林图", "文件": str(path)})

    plt.figure(figsize=(max(7, 1.1 * len(labels)), max(6, 0.9 * len(labels))))
    phi_frame = pd.DataFrame(selected_model.phi, index=labels, columns=labels)
    sns.heatmap(phi_frame, annot=True, fmt=".2f", cmap="vlag", center=0, vmin=-1, vmax=1, square=True)
    plt.title("Promax斜交旋转后的维度相关")
    path = figures_dir / "11_维度相关热图.png"
    save_figure(path)
    records.append({"序号": 11, "图名": "维度相关热图", "文件": str(path)})
    return pd.DataFrame(records)


def write_tables(output_dir: Path, tables: dict[str, pd.DataFrame]) -> dict[str, str]:
    tables_dir = output_dir / "tables"
    tables_dir.mkdir(parents=True, exist_ok=True)
    paths: dict[str, str] = {}
    for name, frame in tables.items():
        path = tables_dir / f"{name}.csv"
        frame.to_csv(path, index=False, encoding="utf-8-sig")
        paths[name] = str(path)
    return paths


def main() -> int:
    args = parse_args()
    warnings.filterwarnings("ignore", message=".*encountered in slogdet", category=RuntimeWarning)
    patch_factor_analyzer_compatibility()
    configure_plotting()
    workbook_path = Path(args.workbook).resolve()
    clean_text_dir = Path(args.clean_text_dir).resolve()
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    data, dictionary, category_map = load_inputs(workbook_path)
    validation = validate_inputs(data, clean_text_dir)
    print(f"输入校验通过：{validation['rows']}篇，{validation['columns']}列")

    screening, screened_features, screened_matrix, screening_diagnostics = screen_features(
        data,
        category_map,
        zero_threshold=args.zero_threshold,
        corr_threshold=args.corr_threshold,
        msa_threshold=args.msa_threshold,
    )
    observed, parallel_threshold, parallel_count = parallel_analysis(
        screened_matrix, args.parallel_iterations, args.seed
    )
    print(
        f"候选变量：{screening_diagnostics['initial_candidates']} -> "
        f"{screening_diagnostics['after_msa']}；KMO={screening_diagnostics['overall_kmo']:.3f}；"
        f"平行分析建议{parallel_count}个因子"
    )

    maximum_factors = min(args.max_factors, max(2, parallel_count))
    models: list[FactorModelResult] = []
    for factor_count in range(2, maximum_factors + 1):
        print(f"拟合{factor_count}因子候选模型……")
        try:
            model = fit_factor_model(
                screened_matrix,
                screened_features,
                factor_count,
                args.communality_threshold,
                args.loading_threshold,
            )
        except Exception as error:
            print(f"  {factor_count}因子模型失败：{error}")
            continue
        if model.diagnostic_valid and not args.skip_bootstrap:
            final_matrix = data[model.feature_names].to_numpy(dtype=float)
            medians, q05, successes = bootstrap_stability(
                final_matrix, model, args.bootstrap, args.seed
            )
            model.stability_medians = medians
            model.stability_q05 = q05
            model.bootstrap_successes = successes
            model.stable = successes >= max(20, int(args.bootstrap * 0.80)) and min(medians) >= 0.85
        elif model.diagnostic_valid:
            model.stability_medians = [1.0] * model.k
            model.stability_q05 = [1.0] * model.k
            model.stable = True
        print(
            f"  保留{len(model.feature_names)}变量，KMO={model.kmo:.3f}，"
            f"累计方差={model.factor_variance[2, -1]:.3f}，诊断={model.diagnostic_note}，"
            f"稳定={model.stable}"
        )
        models.append(model)

    selected_model, selection_note = choose_model(models)
    if selected_model is None:
        raise RuntimeError(f"MF/MD因子分析未得到可用模型：{selection_note}")
    labels, factor_details = name_factors(selected_model)
    print(f"选定{selected_model.k}维模型：{selection_note}")
    for index, label in enumerate(labels, start=1):
        print(f"  D{index} {label}")

    low_communality_map = {
        item["字段名"]: item["共同度"] for item in selected_model.removed_low_communality
    }
    for index, row in screening.iterrows():
        name = row["字段名"]
        if name in low_communality_map:
            screening.loc[index, "筛选状态"] = "删除"
            screening.loc[index, "筛选原因"] = "选定模型中共同度<0.30"
            screening.loc[index, "共同度"] = low_communality_map[name]
        elif name in selected_model.feature_names:
            position = selected_model.feature_names.index(name)
            screening.loc[index, "筛选状态"] = "最终保留"
            screening.loc[index, "筛选原因"] = "进入最终MF/MD模型"
            screening.loc[index, "共同度"] = float(selected_model.communalities[position])

    scores, assignments, score_sensitivity = calculate_dimension_scores(
        data, selected_model, labels, args.loading_threshold
    )
    omnibus, pairwise, dimension_descriptive = group_comparisons(scores, labels)
    score_correlations, regressions = score_and_regression_analysis(scores, labels)
    sample_summary, score_distribution, nationality_distribution = sample_tables(data)
    examples = pd.DataFrame(representative_examples(scores, labels, clean_text_dir))

    loading_records: list[dict[str, Any]] = []
    for index, name in enumerate(selected_model.feature_names):
        record: dict[str, Any] = {
            "字段名": name,
            "类别": category_map.get(name, ""),
            "共同度": float(selected_model.communalities[index]),
            "独特性": float(selected_model.uniquenesses[index]),
        }
        for factor, label in enumerate(labels):
            record[f"D{factor + 1}_{label}"] = float(selected_model.loadings[index, factor])
        record["主要维度"] = labels[int(np.argmax(np.abs(selected_model.loadings[index])))]
        record["主要载荷"] = float(selected_model.loadings[index, np.argmax(np.abs(selected_model.loadings[index]))])
        loading_records.append(record)
    loadings = pd.DataFrame(loading_records)

    model_records: list[dict[str, Any]] = []
    stability_records: list[dict[str, Any]] = []
    for model in models:
        model_records.append(
            {
                "因子数": model.k,
                "最终变量数": len(model.feature_names),
                "删除低共同度变量数": len(model.removed_low_communality),
                "KMO": model.kmo,
                "最小MSA": model.min_msa,
                "Bartlett卡方": model.bartlett_chi2,
                "Bartlett自由度": model.bartlett_df,
                "Bartlett_p值": model.bartlett_p,
                "样本变量比": model.n_per_variable,
                "累计解释方差": float(model.factor_variance[2, -1]),
                "每因子主要指标数": "/".join(str(value) for value in model.strong_assigned_counts),
                "交叉载荷变量数": model.cross_loading_count,
                "Heywood异常": model.heywood,
                "统计诊断通过": model.diagnostic_valid,
                "Bootstrap成功次数": model.bootstrap_successes,
                "最小稳定性中位数": min(model.stability_medians or [0]),
                "稳定性通过": model.stable,
                "是否选定": model is selected_model,
                "诊断说明": model.diagnostic_note,
            }
        )
        if model.stability_medians:
            for factor, median in enumerate(model.stability_medians):
                stability_records.append(
                    {
                        "因子数": model.k,
                        "因子序号": factor + 1,
                        "Tucker一致性中位数": median,
                        "Tucker一致性5%分位": (model.stability_q05 or [0] * model.k)[factor],
                        "Bootstrap成功次数": model.bootstrap_successes,
                    }
                )

    factor_diagnostics = pd.DataFrame(model_records)
    factor_stability = pd.DataFrame(stability_records)
    factor_correlation = pd.DataFrame(selected_model.phi, index=labels, columns=labels).reset_index()
    factor_correlation = factor_correlation.rename(columns={"index": "维度名称"})
    parallel_frame = pd.DataFrame(
        {
            "因子序号": np.arange(1, len(observed) + 1),
            "实测特征根": observed,
            "随机95%阈值": parallel_threshold,
            "实测大于随机阈值": observed > parallel_threshold,
        }
    )
    figure_index = create_figures(
        data,
        screening_diagnostics,
        observed,
        parallel_threshold,
        selected_model,
        labels,
        scores,
        pairwise,
        regressions,
        output_dir / "figures",
    )

    tables = {
        "样本概况": sample_summary,
        "分数分布": score_distribution,
        "国籍分布": nationality_distribution,
        "变量筛选": screening,
        "平行分析": parallel_frame,
        "候选模型诊断": factor_diagnostics,
        "因子稳定性": factor_stability,
        "因子载荷": loadings,
        "指标维度归属": assignments,
        "维度相关": factor_correlation,
        "维度得分": scores,
        "得分敏感性": score_sensitivity,
        "维度描述统计": dimension_descriptive,
        "Welch检验": omnibus,
        "两两比较": pairwise,
        "分数相关": score_correlations,
        "稳健回归": regressions,
        "匿名文本例证": examples,
        "图表索引": figure_index,
    }
    table_paths = write_tables(output_dir, tables)

    metadata = {
        "input_workbook": str(workbook_path),
        "clean_text_dir": str(clean_text_dir),
        "output_dir": str(output_dir),
        "seed": args.seed,
        "validation": validation,
        "screening": screening_diagnostics,
        "parallel_analysis": {
            "iterations": args.parallel_iterations,
            "suggested_factor_count": parallel_count,
        },
        "selected_model": {
            "factor_count": selected_model.k,
            "feature_count": len(selected_model.feature_names),
            "kmo": selected_model.kmo,
            "bartlett_chi2": selected_model.bartlett_chi2,
            "bartlett_df": selected_model.bartlett_df,
            "bartlett_p": selected_model.bartlett_p,
            "n_per_variable": selected_model.n_per_variable,
            "cumulative_variance": float(selected_model.factor_variance[2, -1]),
            "stability_medians": selected_model.stability_medians,
            "stability_q05": selected_model.stability_q05,
            "bootstrap_successes": selected_model.bootstrap_successes,
            "selection_note": selection_note,
            "dimension_labels": labels,
            "dimension_details": factor_details,
        },
        "thresholds": {
            "zero": args.zero_threshold,
            "correlation": args.corr_threshold,
            "msa": args.msa_threshold,
            "communality": args.communality_threshold,
            "loading": args.loading_threshold,
        },
        "tables": table_paths,
        "figures": figure_index.to_dict(orient="records"),
    }
    (output_dir / "analysis_metadata.json").write_text(
        json.dumps(json_value(metadata), ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (output_dir / "workbook_payload.json").write_text(
        json.dumps(
            {
                "metadata": json_value(metadata),
                "sheets": {
                    name: {
                        "columns": list(frame.columns),
                        "rows": json_value(frame.to_numpy(dtype=object).tolist()),
                    }
                    for name, frame in tables.items()
                },
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    print(f"分析完成：{output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
