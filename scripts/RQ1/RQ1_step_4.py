"""
RQ1 Step 4: Visualize the SDG ranking results.

This script uses the Excel output from RQ1 Step 3
(`RQ1_step3_SDG_rankings.xlsx`) and creates visualizations for
Research Question 1: Which SDGs dominate the operations and sustainability
efforts of Croatian SMEs?

The script loads the validated ranking tables from Step 3, checks that the
required columns and 17 SDG rows are present, computes correlation statistics
for visualization, and saves five figures:
1. mean perceived relevance of SDGs,
2. direct connection of SDGs to SME operations,
3. rank comparison between perceived relevance and direct connection,
4. scatter plot of perceived relevance and direct connection,
5. largest rank gaps between the two dimensions.

Input:
    /content/RQ1_step3_SDG_rankings.xlsx

Output:
    /content/RQ1_graphs/
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy.stats import spearmanr


print("RQ1 - STEP 4")
print("Visualization of dominant SDGs")
print("=" * 90)

input_path = "/content/RQ1_step3_SDG_rankings.xlsx"

rq1_relevance_ranking = pd.read_excel(
    input_path,
    sheet_name="relevance_ranking"
)

rq1_direct_ranking = pd.read_excel(
    input_path,
    sheet_name="direct_ranking"
)

rq1_combined_sdg_summary = pd.read_excel(
    input_path,
    sheet_name="combined_summary"
)

rq1_largest_rank_gaps = pd.read_excel(
    input_path,
    sheet_name="largest_rank_gaps"
)

print("\n1) Tables loaded from Step 3:")
print("rq1_relevance_ranking:", rq1_relevance_ranking.shape)
print("rq1_direct_ranking:", rq1_direct_ranking.shape)
print("rq1_combined_sdg_summary:", rq1_combined_sdg_summary.shape)
print("rq1_largest_rank_gaps:", rq1_largest_rank_gaps.shape)

required_relevance_cols = [
    "rank_by_mean_relevance",
    "sdg_key",
    "SDG",
    "mean_relevance",
    "median_relevance",
    "std_relevance",
    "high_relevance_count_4_or_5",
    "high_relevance_percent_4_or_5",
    "very_high_relevance_count_5",
    "very_high_relevance_percent_5"
]

required_direct_cols = [
    "rank_by_direct_connection",
    "sdg_key",
    "SDG",
    "direct_count",
    "direct_percent"
]

required_combined_cols = [
    "rank_by_mean_relevance",
    "sdg_key",
    "SDG",
    "mean_relevance",
    "median_relevance",
    "std_relevance",
    "high_relevance_count_4_or_5",
    "high_relevance_percent_4_or_5",
    "very_high_relevance_count_5",
    "very_high_relevance_percent_5",
    "rank_by_direct_connection",
    "direct_count",
    "direct_percent",
    "rank_difference_direct_minus_relevance",
    "absolute_rank_difference"
]

missing_relevance = [
    col for col in required_relevance_cols
    if col not in rq1_relevance_ranking.columns
]

missing_direct = [
    col for col in required_direct_cols
    if col not in rq1_direct_ranking.columns
]

missing_combined = [
    col for col in required_combined_cols
    if col not in rq1_combined_sdg_summary.columns
]

if missing_relevance:
    raise ValueError(f"Missing columns in relevance_ranking: {missing_relevance}")

if missing_direct:
    raise ValueError(f"Missing columns in direct_ranking: {missing_direct}")

if missing_combined:
    raise ValueError(f"Missing columns in combined_summary: {missing_combined}")

relevance_has_17 = len(rq1_relevance_ranking) == 17
direct_has_17 = len(rq1_direct_ranking) == 17
combined_has_17 = len(rq1_combined_sdg_summary) == 17

print("\n2) Validation of loaded tables:")
print("Relevance ranking table has 17 rows:", relevance_has_17)
print("Direct connection ranking table has 17 rows:", direct_has_17)
print("Combined table has 17 rows:", combined_has_17)

if not (relevance_has_17 and direct_has_17 and combined_has_17):
    raise ValueError("One or more RQ1 tables do not contain the expected 17 SDGs.")

output_dir = Path("/content/RQ1_graphs")
output_dir.mkdir(exist_ok=True)

print("\nFigures will be saved in:")
print(output_dir)

plt.rcParams.update({
    "font.size": 11,
    "axes.titlesize": 13,
    "axes.labelsize": 11,
    "xtick.labelsize": 10,
    "ytick.labelsize": 10,
    "legend.fontsize": 10
})

rank_corr, rank_corr_p = spearmanr(
    rq1_combined_sdg_summary["rank_by_mean_relevance"],
    rq1_combined_sdg_summary["rank_by_direct_connection"]
)

measure_corr, measure_corr_p = spearmanr(
    rq1_combined_sdg_summary["mean_relevance"],
    rq1_combined_sdg_summary["direct_percent"]
)

print("\n3) Spearman correlations used in the figures:")
print(
    "Relevance rank vs direct-connection rank: "
    f"rho = {rank_corr:.4f}, p = {rank_corr_p:.4f}"
)
print(
    "Mean relevance vs direct percentage: "
    f"rho = {measure_corr:.4f}, p = {measure_corr_p:.4f}"
)

plot_data = rq1_relevance_ranking.sort_values(
    "mean_relevance",
    ascending=True
)

plt.figure(figsize=(11, 7))
plt.barh(plot_data["SDG"], plot_data["mean_relevance"])
plt.xlabel("Mean relevance score, Likert scale 1–5")
plt.ylabel("UN Sustainable Development Goal")
plt.title("RQ1: Mean relevance of SDGs for Croatian SMEs")
plt.xlim(0, 5)

for index, value in enumerate(plot_data["mean_relevance"]):
    plt.text(value + 0.03, index, f"{value:.2f}", va="center")

plt.tight_layout()

fig1_path = output_dir / "RQ1_01_mean_relevance_SDGs.png"
plt.savefig(fig1_path, dpi=300, bbox_inches="tight")
plt.show()

print("Saved Figure 1:", fig1_path)

plot_data = rq1_direct_ranking.sort_values(
    "direct_percent",
    ascending=True
)

plt.figure(figsize=(11, 7))
plt.barh(plot_data["SDG"], plot_data["direct_percent"])
plt.xlabel("Share of SMEs directly linking the SDG to business operations (%)")
plt.ylabel("UN Sustainable Development Goal")
plt.title("RQ1: Direct connection of SDGs to SME operations")
plt.xlim(0, max(plot_data["direct_percent"]) + 10)

for index, value in enumerate(plot_data["direct_percent"]):
    plt.text(value + 0.3, index, f"{value:.1f}%", va="center")

plt.tight_layout()

fig2_path = output_dir / "RQ1_02_direct_connection_SDGs.png"
plt.savefig(fig2_path, dpi=300, bbox_inches="tight")
plt.show()

print("Saved Figure 2:", fig2_path)

plot_data = rq1_combined_sdg_summary.copy()
plot_data = plot_data.sort_values(
    "rank_by_mean_relevance",
    ascending=False
).reset_index(drop=True)

y_positions = np.arange(len(plot_data))

plt.figure(figsize=(11, 7))

plt.scatter(
    plot_data["rank_by_mean_relevance"],
    y_positions,
    label="Rank by mean relevance"
)

plt.scatter(
    plot_data["rank_by_direct_connection"],
    y_positions,
    label="Rank by direct connection"
)

for y, (_, row) in zip(y_positions, plot_data.iterrows()):
    plt.plot(
        [row["rank_by_mean_relevance"], row["rank_by_direct_connection"]],
        [y, y]
    )

plt.yticks(y_positions, plot_data["SDG"])
plt.xlabel("Rank position, 1 = highest dominance")
plt.ylabel("UN Sustainable Development Goal")
plt.title(
    "RQ1: Difference between perceived relevance and direct business connection\n"
    f"Spearman rank correlation: rho = {rank_corr:.3f}, p = {rank_corr_p:.4f}"
)

plt.xlim(0.5, 17.5)
plt.xticks(range(1, 18))
plt.legend()
plt.tight_layout()

fig3_path = output_dir / "RQ1_03_rank_comparison_relevance_vs_direct.png"
plt.savefig(fig3_path, dpi=300, bbox_inches="tight")
plt.show()

print("Saved Figure 3:", fig3_path)

plt.figure(figsize=(9, 7))

plt.scatter(
    rq1_combined_sdg_summary["mean_relevance"],
    rq1_combined_sdg_summary["direct_percent"]
)

label_offsets = {
    "SDG 3": (0.015, 0.65),
    "SDG 5": (0.015, -0.75),
    "SDG 12": (0.015, 0.15),
    "SDG 9": (0.015, -0.85),
    "SDG 6": (0.015, 0.55),
    "SDG 2": (0.015, -0.55),
    "SDG 15": (0.015, 0.55),
    "SDG 16": (0.015, -0.55),
    "SDG 17": (0.015, -0.30),
    "SDG 10": (0.015, -0.17),
    "SDG 13": (0.015, -0.15)
}

for _, row in rq1_combined_sdg_summary.iterrows():
    short_label = str(row["SDG"]).split(" - ")[0]
    dx, dy = label_offsets.get(short_label, (0.015, 0.25))

    plt.text(
        row["mean_relevance"] + dx,
        row["direct_percent"] + dy,
        short_label,
        fontsize=9
    )

plt.xlabel("Mean relevance score, Likert scale 1–5")
plt.ylabel("Direct connection to business operations (%)")
plt.title(
    "RQ1: Perceived SDG relevance vs. direct business connection\n"
    f"Spearman rho = {measure_corr:.3f}, p = {measure_corr_p:.4f}"
)

plt.xlim(
    rq1_combined_sdg_summary["mean_relevance"].min() - 0.1,
    rq1_combined_sdg_summary["mean_relevance"].max() + 0.15
)

plt.ylim(
    0,
    rq1_combined_sdg_summary["direct_percent"].max() + 8
)

plt.tight_layout()

fig4_path = output_dir / "RQ1_04_scatter_relevance_vs_direct.png"
plt.savefig(fig4_path, dpi=300, bbox_inches="tight")
plt.show()

print("Saved Figure 4:", fig4_path)

plot_data = rq1_largest_rank_gaps.copy()
plot_data = plot_data.head(10)

plot_data = plot_data.sort_values(
    "absolute_rank_difference",
    ascending=True
)

plt.figure(figsize=(11, 6))
plt.barh(
    plot_data["SDG"],
    plot_data["rank_difference_direct_minus_relevance"]
)

plt.axvline(0, linewidth=1)
plt.xlabel("Rank difference: direct-connection rank minus relevance rank")
plt.ylabel("UN Sustainable Development Goal")
plt.title("")

for index, value in enumerate(plot_data["rank_difference_direct_minus_relevance"]):
    x_position = value + 0.15
    plt.text(
        x_position,
        index,
        f"{value:.0f}",
        va="center",
        ha="left"
    )

plt.tight_layout()

fig5_path = output_dir / "RQ1_05_largest_rank_gaps.png"
plt.savefig(fig5_path, dpi=300, bbox_inches="tight")
plt.show()

print("Saved Figure 5:", fig5_path)

print("\n" + "=" * 90)
print("RQ1 - STEP 4 COMPLETED")
print("=" * 90)

print("Five visualizations were created for RQ1:")
print("1. Mean relevance of SDGs")
print("2. Direct connection of SDGs to business operations")
print("3. Rank comparison between relevance and direct connection")
print("4. Scatter plot of relevance and direct connection")
print("5. Largest gaps between relevance and direct connection")

print("\nSaved files:")
print(fig1_path)
print(fig2_path)
print(fig3_path)
print(fig4_path)
print(fig5_path)
