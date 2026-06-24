"""
RQ1 Step 5: Final quantification of the relationship between perceived SDG relevance and direct business connection.

This script supports RQ1: Which SDGs dominate the operations and sustainability efforts of Croatian SMEs?

The script performs the final SDG-level analysis using the coded survey dataset. It compares perceived SDG relevance from q36 with direct business connection from q37 across the 17 SDGs. It computes SDG-level summaries, relevance and direct-connection ranks, correlations, rank differences, and an exploratory RQ1 dominance index. It also performs inferential checks using the Friedman test for q36 relevance, Cochran's Q test for q37 direct connection, and post-hoc Wilcoxon and McNemar tests with Benjamini-Hochberg FDR correction.

Input:
    Survey_Results.xlsx, sheet: encoded_responses_corrected

Output:
    RQ1_step5_dominance_inference.xlsx
"""

import pandas as pd
import numpy as np
import scipy.stats as stats
from itertools import combinations
from statsmodels.stats.contingency_tables import cochrans_q, mcnemar

file_path = "/content/Survey_Results.xlsx"
output_path = "/content/RQ1_step5_dominance_inference.xlsx"

coded_data = pd.read_excel(file_path, sheet_name="encoded_responses_corrected")

sdg_relevance_cols = [
    col for col in coded_data.columns
    if col.startswith("q36_relevantnost_sdg__")
]

sdg_direct_cols = [
    col for col in coded_data.columns
    if col.startswith("q37_sdg_vezani_uz_poslovanje__")
    and col != "q37_sdg_vezani_uz_poslovanje__none_directly_related"
]

sdg_short_names = {
    "sdg01_no_poverty": "SDG 1",
    "sdg02_zero_hunger": "SDG 2",
    "sdg03_good_health": "SDG 3",
    "sdg04_quality_education": "SDG 4",
    "sdg05_gender_equality": "SDG 5",
    "sdg06_clean_water": "SDG 6",
    "sdg07_clean_energy": "SDG 7",
    "sdg08_decent_work": "SDG 8",
    "sdg09_industry_innovation": "SDG 9",
    "sdg10_reduced_inequalities": "SDG 10",
    "sdg11_sustainable_cities": "SDG 11",
    "sdg12_responsible_consumption": "SDG 12",
    "sdg13_climate_action": "SDG 13",
    "sdg14_life_below_water": "SDG 14",
    "sdg15_life_on_land": "SDG 15",
    "sdg16_peace_justice": "SDG 16",
    "sdg17_partnerships": "SDG 17",
}

sdg_full_names = {
    "sdg01_no_poverty": "No poverty",
    "sdg02_zero_hunger": "Zero hunger",
    "sdg03_good_health": "Good health and well-being",
    "sdg04_quality_education": "Quality education",
    "sdg05_gender_equality": "Gender equality",
    "sdg06_clean_water": "Clean water and sanitation",
    "sdg07_clean_energy": "Affordable and clean energy",
    "sdg08_decent_work": "Decent work and economic growth",
    "sdg09_industry_innovation": "Industry, innovation and infrastructure",
    "sdg10_reduced_inequalities": "Reduced inequalities",
    "sdg11_sustainable_cities": "Sustainable cities and communities",
    "sdg12_responsible_consumption": "Responsible consumption and production",
    "sdg13_climate_action": "Climate action",
    "sdg14_life_below_water": "Life below water",
    "sdg15_life_on_land": "Life on land",
    "sdg16_peace_justice": "Peace, justice and strong institutions",
    "sdg17_partnerships": "Partnerships for the goals",
}


def extract_sdg_key(column_name: str) -> str:
    """Extract the SDG key from a coded q36 or q37 column name."""
    return column_name.split("__", 1)[1]


def benjamini_hochberg(p_values):
    """Apply Benjamini-Hochberg FDR correction to a sequence of p-values."""
    p_values = np.asarray(p_values, dtype=float)
    n_values = len(p_values)
    order = np.argsort(p_values)
    ranked_p_values = p_values[order]
    adjusted = np.empty(n_values, dtype=float)
    cumulative_minimum = 1.0

    for i in range(n_values - 1, -1, -1):
        rank = i + 1
        adjusted_value = ranked_p_values[i] * n_values / rank
        cumulative_minimum = min(cumulative_minimum, adjusted_value)
        adjusted[order[i]] = cumulative_minimum

    return np.minimum(adjusted, 1.0)


print("RQ1 - Step 5")
print("Final quantification, dominance index and inferential checks")
print("=" * 100)

if len(sdg_relevance_cols) != 17:
    raise ValueError(f"Expected 17 q36 SDG relevance columns, but found {len(sdg_relevance_cols)}.")

if len(sdg_direct_cols) != 17:
    raise ValueError(f"Expected 17 q37 direct SDG connection columns, but found {len(sdg_direct_cols)}.")

q36_sdg_keys = sorted([extract_sdg_key(col) for col in sdg_relevance_cols])
q37_sdg_keys = sorted([extract_sdg_key(col) for col in sdg_direct_cols])

if q36_sdg_keys != q37_sdg_keys:
    raise ValueError("The q36 and q37 SDG keys do not match.")

rows = []

for relevance_col in sdg_relevance_cols:
    sdg_key = extract_sdg_key(relevance_col)
    direct_col = "q37_sdg_vezani_uz_poslovanje__" + sdg_key
    relevance_series = coded_data[relevance_col]
    direct_series = coded_data[direct_col]

    rows.append({
        "sdg_key": sdg_key,
        "SDG": sdg_short_names.get(sdg_key, sdg_key),
        "SDG_full": sdg_full_names.get(sdg_key, sdg_key),
        "mean_relevance": relevance_series.mean(),
        "median_relevance": relevance_series.median(),
        "std_relevance": relevance_series.std(ddof=1),
        "high_relevance_count_4_or_5": int((relevance_series >= 4).sum()),
        "high_relevance_percent_4_or_5": (relevance_series >= 4).mean() * 100,
        "direct_count": int(direct_series.sum()),
        "direct_percent": direct_series.mean() * 100,
    })

rq1_sdg_level_summary = pd.DataFrame(rows)

rq1_sdg_level_summary["rank_mean_relevance"] = (
    rq1_sdg_level_summary["mean_relevance"].rank(ascending=False, method="min").astype(int)
)

rq1_sdg_level_summary["rank_direct_percent"] = (
    rq1_sdg_level_summary["direct_percent"].rank(ascending=False, method="min").astype(int)
)

rq1_sdg_level_summary["rank_difference_direct_minus_relevance"] = (
    rq1_sdg_level_summary["rank_direct_percent"]
    - rq1_sdg_level_summary["rank_mean_relevance"]
)

pearson_r, pearson_p = stats.pearsonr(
    rq1_sdg_level_summary["mean_relevance"],
    rq1_sdg_level_summary["direct_percent"],
)

spearman_rho, spearman_p = stats.spearmanr(
    rq1_sdg_level_summary["mean_relevance"],
    rq1_sdg_level_summary["direct_percent"],
)

kendall_tau, kendall_p = stats.kendalltau(
    rq1_sdg_level_summary["mean_relevance"],
    rq1_sdg_level_summary["direct_percent"],
)

correlation_summary = pd.DataFrame({
    "correlation_type": ["Pearson r", "Spearman rho", "Kendall tau"],
    "what_it_measures": [
        "Linear association between mean relevance and direct percentage",
        "Monotonic/rank association between relevance and direct connection",
        "Rank agreement between relevance and direct connection",
    ],
    "coefficient": [pearson_r, spearman_rho, kendall_tau],
    "p_value": [pearson_p, spearman_p, kendall_p],
    "n_SDG_goals": [
        len(rq1_sdg_level_summary),
        len(rq1_sdg_level_summary),
        len(rq1_sdg_level_summary),
    ],
})

rq1_sdg_level_summary["z_mean_relevance"] = (
    rq1_sdg_level_summary["mean_relevance"]
    - rq1_sdg_level_summary["mean_relevance"].mean()
) / rq1_sdg_level_summary["mean_relevance"].std(ddof=1)

rq1_sdg_level_summary["z_direct_percent"] = (
    rq1_sdg_level_summary["direct_percent"]
    - rq1_sdg_level_summary["direct_percent"].mean()
) / rq1_sdg_level_summary["direct_percent"].std(ddof=1)

rq1_sdg_level_summary["RQ1_dominance_index"] = (
    rq1_sdg_level_summary["z_mean_relevance"]
    + rq1_sdg_level_summary["z_direct_percent"]
) / 2

rq1_sdg_level_summary["rank_RQ1_dominance"] = (
    rq1_sdg_level_summary["RQ1_dominance_index"]
    .rank(ascending=False, method="min")
    .astype(int)
)

q36_matrix = coded_data[sdg_relevance_cols]

friedman_stat, friedman_p = stats.friedmanchisquare(
    *[q36_matrix[col] for col in sdg_relevance_cols]
)

n_subjects = q36_matrix.shape[0]
k_goals = q36_matrix.shape[1]
kendalls_w = friedman_stat / (n_subjects * (k_goals - 1))

friedman_summary = pd.DataFrame([{
    "test": "Friedman test",
    "outcome": "q36 SDG relevance",
    "n_subjects": n_subjects,
    "k_SDG_goals": k_goals,
    "statistic": friedman_stat,
    "p_value": friedman_p,
    "effect_size_Kendalls_W": kendalls_w,
}])

q37_matrix = coded_data[sdg_direct_cols]
cochran_result = cochrans_q(q37_matrix)

cochran_summary = pd.DataFrame([{
    "test": "Cochran Q test",
    "outcome": "q37 direct SDG connection",
    "n_subjects": q37_matrix.shape[0],
    "k_SDG_goals": q37_matrix.shape[1],
    "statistic": cochran_result.statistic,
    "p_value": cochran_result.pvalue,
}])

wilcoxon_rows = []

for col_a, col_b in combinations(sdg_relevance_cols, 2):
    key_a = extract_sdg_key(col_a)
    key_b = extract_sdg_key(col_b)
    x = coded_data[col_a]
    y = coded_data[col_b]

    try:
        statistic, p_value = stats.wilcoxon(x, y, zero_method="wilcox", alternative="two-sided")
    except ValueError:
        statistic, p_value = np.nan, np.nan

    wilcoxon_rows.append({
        "SDG_A": sdg_short_names.get(key_a, key_a),
        "SDG_B": sdg_short_names.get(key_b, key_b),
        "mean_A": x.mean(),
        "mean_B": y.mean(),
        "mean_difference_A_minus_B": x.mean() - y.mean(),
        "wilcoxon_statistic": statistic,
        "p_value": p_value,
    })

wilcoxon_posthoc = pd.DataFrame(wilcoxon_rows)
valid_wilcoxon = wilcoxon_posthoc["p_value"].notna()

wilcoxon_posthoc.loc[valid_wilcoxon, "p_value_fdr"] = benjamini_hochberg(
    wilcoxon_posthoc.loc[valid_wilcoxon, "p_value"].values
)

wilcoxon_posthoc["significant_fdr_0_05"] = wilcoxon_posthoc["p_value_fdr"] < 0.05

mcnemar_rows = []

for col_a, col_b in combinations(sdg_direct_cols, 2):
    key_a = extract_sdg_key(col_a)
    key_b = extract_sdg_key(col_b)
    a = coded_data[col_a].astype(int)
    b = coded_data[col_b].astype(int)

    both_1 = int(((a == 1) & (b == 1)).sum())
    a1_b0 = int(((a == 1) & (b == 0)).sum())
    a0_b1 = int(((a == 0) & (b == 1)).sum())
    both_0 = int(((a == 0) & (b == 0)).sum())

    table = [[both_1, a1_b0], [a0_b1, both_0]]
    result = mcnemar(table, exact=True)

    mcnemar_rows.append({
        "SDG_A": sdg_short_names.get(key_a, key_a),
        "SDG_B": sdg_short_names.get(key_b, key_b),
        "direct_percent_A": a.mean() * 100,
        "direct_percent_B": b.mean() * 100,
        "percent_difference_A_minus_B": (a.mean() - b.mean()) * 100,
        "both_1": both_1,
        "A1_B0": a1_b0,
        "A0_B1": a0_b1,
        "both_0": both_0,
        "mcnemar_statistic": result.statistic,
        "p_value": result.pvalue,
    })

mcnemar_posthoc = pd.DataFrame(mcnemar_rows)
mcnemar_posthoc["p_value_fdr"] = benjamini_hochberg(mcnemar_posthoc["p_value"].values)
mcnemar_posthoc["significant_fdr_0_05"] = mcnemar_posthoc["p_value_fdr"] < 0.05

correlation_summary_display = correlation_summary.copy()
correlation_summary_display["coefficient"] = correlation_summary_display["coefficient"].round(3)
correlation_summary_display["p_value"] = correlation_summary_display["p_value"].round(4)

dominance_display = (
    rq1_sdg_level_summary
    .sort_values("rank_RQ1_dominance")
    .reset_index(drop=True)[[
        "rank_RQ1_dominance",
        "SDG",
        "SDG_full",
        "mean_relevance",
        "direct_percent",
        "rank_mean_relevance",
        "rank_direct_percent",
        "rank_difference_direct_minus_relevance",
        "RQ1_dominance_index",
    ]]
)

dominance_display_rounded = dominance_display.copy()
for col in ["mean_relevance", "direct_percent", "RQ1_dominance_index"]:
    dominance_display_rounded[col] = dominance_display_rounded[col].round(3)

rank_difference_display = (
    rq1_sdg_level_summary
    .sort_values("rank_difference_direct_minus_relevance")
    .reset_index(drop=True)[[
        "SDG",
        "SDG_full",
        "rank_mean_relevance",
        "rank_direct_percent",
        "rank_difference_direct_minus_relevance",
        "mean_relevance",
        "direct_percent",
    ]]
)

rank_difference_display_rounded = rank_difference_display.copy()
for col in ["mean_relevance", "direct_percent"]:
    rank_difference_display_rounded[col] = rank_difference_display_rounded[col].round(3)

friedman_summary_display = friedman_summary.copy()
for col in ["statistic", "p_value", "effect_size_Kendalls_W"]:
    friedman_summary_display[col] = friedman_summary_display[col].round(4)

cochran_summary_display = cochran_summary.copy()
for col in ["statistic", "p_value"]:
    cochran_summary_display[col] = cochran_summary_display[col].round(4)

wilcoxon_display = wilcoxon_posthoc.copy()
for col in ["mean_A", "mean_B", "mean_difference_A_minus_B", "p_value", "p_value_fdr"]:
    wilcoxon_display[col] = wilcoxon_display[col].round(4)

mcnemar_display = mcnemar_posthoc.copy()
for col in ["direct_percent_A", "direct_percent_B", "percent_difference_A_minus_B", "p_value", "p_value_fdr"]:
    mcnemar_display[col] = mcnemar_display[col].round(4)

print("\nCorrelation between q36 mean relevance and q37 direct connection")
print(correlation_summary_display.to_string(index=False))

print("\nRQ1 combined SDG dominance index")
print(dominance_display_rounded.to_string(index=False))

print("\nRank differences: direct-connection rank minus relevance rank")
print(rank_difference_display_rounded.to_string(index=False))

print("\nInferential analysis for q36: Friedman test")
print(friedman_summary_display.to_string(index=False))

print("\nInferential analysis for q37: Cochran Q test")
print(cochran_summary_display.to_string(index=False))

print("\nPost-hoc q36: Wilcoxon tests significant after FDR correction")
wilcoxon_sig = wilcoxon_display[wilcoxon_display["significant_fdr_0_05"]]
if len(wilcoxon_sig) > 0:
    print(
        wilcoxon_sig.sort_values("p_value_fdr")[[
            "SDG_A",
            "SDG_B",
            "mean_A",
            "mean_B",
            "mean_difference_A_minus_B",
            "p_value",
            "p_value_fdr",
        ]].to_string(index=False)
    )
else:
    print("No q36 pairwise differences remained significant after FDR correction.")

print("\nPost-hoc q37: McNemar tests significant after FDR correction")
mcnemar_sig = mcnemar_display[mcnemar_display["significant_fdr_0_05"]]
if len(mcnemar_sig) > 0:
    print(
        mcnemar_sig.sort_values("p_value_fdr")[[
            "SDG_A",
            "SDG_B",
            "direct_percent_A",
            "direct_percent_B",
            "percent_difference_A_minus_B",
            "p_value",
            "p_value_fdr",
        ]].to_string(index=False)
    )
else:
    print("No q37 pairwise differences remained significant after FDR correction.")

print("\nSummary")
print(f"Pearson r = {pearson_r:.3f}, p = {pearson_p:.4f}")
print(f"Spearman rho = {spearman_rho:.3f}, p = {spearman_p:.4f}")
print(f"Kendall tau = {kendall_tau:.3f}, p = {kendall_p:.4f}")
print(f"Friedman q36: statistic = {friedman_stat:.3f}, p = {friedman_p:.4f}, Kendall's W = {kendalls_w:.3f}")
print(f"Cochran Q q37: statistic = {cochran_result.statistic:.3f}, p = {cochran_result.pvalue:.4f}")

print("\nTop 5 SDGs according to the combined RQ1 dominance index")
for _, row in dominance_display_rounded.head(5).iterrows():
    print(
        f"{row['rank_RQ1_dominance']}. {row['SDG']} - {row['SDG_full']} "
        f"(mean relevance = {row['mean_relevance']}, direct = {row['direct_percent']}%)"
    )

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    rq1_sdg_level_summary.to_excel(writer, sheet_name="sdg_level_summary", index=False)
    correlation_summary.to_excel(writer, sheet_name="correlations", index=False)
    dominance_display.to_excel(writer, sheet_name="dominance_ranking", index=False)
    rank_difference_display.to_excel(writer, sheet_name="rank_differences", index=False)
    friedman_summary.to_excel(writer, sheet_name="friedman_q36", index=False)
    cochran_summary.to_excel(writer, sheet_name="cochran_q37", index=False)
    wilcoxon_posthoc.to_excel(writer, sheet_name="posthoc_wilcoxon_q36", index=False)
    mcnemar_posthoc.to_excel(writer, sheet_name="posthoc_mcnemar_q37", index=False)

print("\nRQ1 Step 5 results saved to:")
print(output_path)
