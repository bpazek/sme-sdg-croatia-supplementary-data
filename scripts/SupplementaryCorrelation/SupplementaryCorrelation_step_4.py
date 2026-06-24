"""
Supplementary Correlation Step 4: Graphs and final synthesis.

This script supports the supplementary analysis of the relationship between
external funding sources used during the previous three years and motives for
implemented improvements.

The script uses the index file created in Step 2 and the association results
created in Step 3. It generates the main figures for the supplementary analysis,
optional appendix figures, a final summary table, and the data behind the main
bubble plot.

Input files:
    Additional_Analysis_3_q31_q34_indices.xlsx
    Additional_Analysis_3_q31_q34_results_with_kendall.xlsx

Output folder:
    Additional_Analysis_3_q31_q34_graphs/

Output file:
    Additional_Analysis_3_q31_q34_step4_graph_summaries.xlsx
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


INDICES_PATH = "/content/Additional_Analysis_3_q31_q34_indices.xlsx"
RESULTS_PATH = "/content/Additional_Analysis_3_q31_q34_results_with_kendall.xlsx"
OUTPUT_DIR = Path("/content/Additional_Analysis_3_q31_q34_graphs")
SUMMARY_OUTPUT_PATH = "/content/Additional_Analysis_3_q31_q34_step4_graph_summaries.xlsx"
GENERATE_APPENDIX_GRAPHS = True

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

print("SUPPLEMENTARY CORRELATION - STEP 4")
print("Graphs and final synthesis")
print("=" * 100)


df = pd.read_excel(
    INDICES_PATH,
    sheet_name="encoded_with_q31_q34"
)

main_correlation = pd.read_excel(
    RESULTS_PATH,
    sheet_name="main_correlation"
)

aggregate_correlations = pd.read_excel(
    RESULTS_PATH,
    sheet_name="aggregate_correlations"
)

funding_group_compare = pd.read_excel(
    RESULTS_PATH,
    sheet_name="funding_group_compare"
)

pairwise_q34_q31 = pd.read_excel(
    RESULTS_PATH,
    sheet_name="pairwise_q34_q31"
)

top_abs_phi = pd.read_excel(
    RESULTS_PATH,
    sheet_name="top_abs_phi"
)

aggregate_binary_pairs = pd.read_excel(
    RESULTS_PATH,
    sheet_name="aggregate_binary_pairs"
)

phi_matrix = pd.read_excel(
    RESULTS_PATH,
    sheet_name="phi_matrix",
    index_col=0
)

fdr_p_matrix = pd.read_excel(
    RESULTS_PATH,
    sheet_name="fdr_p_matrix",
    index_col=0
)

risk_diff_matrix = pd.read_excel(
    RESULTS_PATH,
    sheet_name="risk_diff_matrix",
    index_col=0
)

print("Main dataset dimensions:", df.shape)
print("Number of firms:", len(df))

required_df_columns = [
    "External_funding_count",
    "Improvement_motivators_count",
    "Any_external_funding",
    "Any_EU_or_public_funding",
    "Any_bank_or_credit_support",
]

missing_df_columns = [
    column for column in required_df_columns
    if column not in df.columns
]

if missing_df_columns:
    raise ValueError(
        "The main table is missing required columns: "
        + ", ".join(missing_df_columns)
    )

required_main_columns = [
    "spearman_rho",
    "spearman_p",
    "kendall_tau_b",
    "kendall_p",
    "n",
]

missing_main_columns = [
    column for column in required_main_columns
    if column not in main_correlation.columns
]

if missing_main_columns:
    raise ValueError(
        "The main_correlation table is missing required columns: "
        + ", ".join(missing_main_columns)
    )

required_group_columns = [
    "funding_indicator",
    "n_selected",
    "n_not_selected",
    "mean_motivators_selected",
    "mean_motivators_not_selected",
    "median_motivators_selected",
    "median_motivators_not_selected",
    "mean_difference_selected_minus_not",
    "mannwhitney_u",
    "mannwhitney_p",
    "mannwhitney_p_fdr",
    "cliffs_delta",
    "cliffs_delta_magnitude",
    "significant_fdr_0_05",
]

missing_group_columns = [
    column for column in required_group_columns
    if column not in funding_group_compare.columns
]

if missing_group_columns:
    raise ValueError(
        "The funding_group_compare table is missing required columns: "
        + ", ".join(missing_group_columns)
    )

main_result = main_correlation.iloc[0]

spearman_rho = float(main_result["spearman_rho"])
spearman_p = float(main_result["spearman_p"])
kendall_tau_b = float(main_result["kendall_tau_b"])
kendall_p = float(main_result["kendall_p"])
sample_n = int(main_result["n"])

print("\n1) Main correlation result:")
print(f"Spearman rho = {spearman_rho:.4f}, p = {spearman_p:.4f}, n = {sample_n}")
print(f"Kendall tau-b = {kendall_tau_b:.4f}, p = {kendall_p:.4f}")

funding_indicator_labels = {
    "Any_external_funding": "Any external funding",
    "Any_EU_or_public_funding": "Any EU or public funding",
    "Any_bank_or_credit_support": "Any bank or credit support",
}

selected_funding_indicators = [
    "Any_external_funding",
    "Any_EU_or_public_funding",
    "Any_bank_or_credit_support",
]

bubble_data = (
    df
    .groupby(
        ["External_funding_count", "Improvement_motivators_count"],
        dropna=False
    )
    .size()
    .reset_index(name="firm_count")
    .sort_values(["External_funding_count", "Improvement_motivators_count"])
)

print("\n2) Observed index combinations used in the bubble plot:")
print(bubble_data.to_string(index=False))

plt.figure(figsize=(11, 7.5))
plt.scatter(
    bubble_data["External_funding_count"],
    bubble_data["Improvement_motivators_count"],
    s=bubble_data["firm_count"] * 55,
    alpha=0.65,
    edgecolors="black",
    linewidths=0.8
)

for row in bubble_data.itertuples():
    plt.text(
        row.External_funding_count,
        row.Improvement_motivators_count,
        str(row.firm_count),
        ha="center",
        va="center",
        fontsize=9
    )

plt.xlabel("Number of external funding sources used", fontsize=12)
plt.ylabel("Number of positive improvement motives reported", fontsize=12)
plt.title(
    "Relationship between external funding sources and positive improvement motives\n"
    rf"Spearman $\rho_S$ = {spearman_rho:.3f}, $p$ = {spearman_p:.4f}; "
    rf"Kendall $\tau_b$ = {kendall_tau_b:.3f}, $p$ = {kendall_p:.4f}",
    fontsize=14
)
plt.xticks(
    range(
        int(df["External_funding_count"].min()),
        int(df["External_funding_count"].max()) + 1
    )
)
plt.yticks(
    range(
        int(df["Improvement_motivators_count"].min()),
        int(df["Improvement_motivators_count"].max()) + 1
    )
)
plt.grid(alpha=0.25)
plt.tight_layout()

figure1_path = OUTPUT_DIR / "AA3_Figure_1_main_bubble_funding_vs_motivators.png"
plt.savefig(figure1_path, dpi=300, bbox_inches="tight")
plt.show()

print("\nFigure 1 saved to:")
print(figure1_path)

group_plot = funding_group_compare[
    funding_group_compare["funding_indicator"].isin(selected_funding_indicators)
].copy()

group_plot["funding_label"] = group_plot["funding_indicator"].map(
    funding_indicator_labels
)

group_plot = group_plot.sort_values(
    "mean_difference_selected_minus_not",
    ascending=True
)

print("\n3) Supplementary group comparisons included in Figure 2:")
print(
    group_plot[
        [
            "funding_indicator",
            "n_selected",
            "n_not_selected",
            "mean_motivators_selected",
            "mean_motivators_not_selected",
            "mean_difference_selected_minus_not",
            "cliffs_delta",
            "mannwhitney_p_fdr",
            "significant_fdr_0_05",
        ]
    ].to_string(index=False)
)

y_positions = np.arange(len(group_plot))
bar_height = 0.34

plt.figure(figsize=(11, 6))
plt.barh(
    y_positions - bar_height / 2,
    group_plot["mean_motivators_selected"],
    bar_height,
    label="Funding category reported"
)
plt.barh(
    y_positions + bar_height / 2,
    group_plot["mean_motivators_not_selected"],
    bar_height,
    label="Funding category not reported"
)
plt.yticks(y_positions, group_plot["funding_label"], fontsize=11)
plt.xlabel("Mean number of positive improvement motives", fontsize=12)
plt.ylabel("External funding category", fontsize=12)
plt.title("Positive improvement motives by external funding category", fontsize=14)
plt.legend(loc="upper right")

for i, row in enumerate(group_plot.itertuples()):
    selected_text = f"{row.mean_motivators_selected:.2f}"
    not_selected_text = f"{row.mean_motivators_not_selected:.2f}"

    if bool(row.significant_fdr_0_05):
        selected_text += "*"

    plt.text(
        row.mean_motivators_selected + 0.04,
        i - bar_height / 2,
        selected_text,
        va="center",
        fontsize=10
    )
    plt.text(
        row.mean_motivators_not_selected + 0.04,
        i + bar_height / 2,
        not_selected_text,
        va="center",
        fontsize=10
    )

plt.figtext(
    0.01,
    -0.01,
    "* FDR-adjusted Mann-Whitney p < 0.05",
    ha="left",
    fontsize=10
)
plt.tight_layout()

figure2_path = OUTPUT_DIR / "AA3_Figure_2_mean_motivators_by_funding_category.png"
plt.savefig(figure2_path, dpi=300, bbox_inches="tight")
plt.show()

print("\nFigure 2 saved to:")
print(figure2_path)

appendix_figure1_path = None
appendix_figure2_path = None

if GENERATE_APPENDIX_GRAPHS:
    heatmap_data = phi_matrix.values.astype(float)
    plt.figure(figsize=(16, 10))
    image = plt.imshow(heatmap_data, aspect="auto")
    plt.colorbar(image, label=r"Phi coefficient ($\phi$)")
    plt.xticks(
        ticks=np.arange(len(phi_matrix.columns)),
        labels=phi_matrix.columns,
        rotation=45,
        ha="right",
        fontsize=9
    )
    plt.yticks(
        ticks=np.arange(len(phi_matrix.index)),
        labels=phi_matrix.index,
        fontsize=9
    )
    plt.title(
        "Exploratory pairwise associations: external funding sources x improvement motives",
        fontsize=14
    )

    for i in range(heatmap_data.shape[0]):
        for j in range(heatmap_data.shape[1]):
            value = heatmap_data[i, j]
            if np.isfinite(value):
                marker = ""
                try:
                    if float(fdr_p_matrix.iloc[i, j]) < 0.05:
                        marker = "*"
                except Exception:
                    marker = ""

                plt.text(
                    j,
                    i,
                    f"{value:.2f}{marker}",
                    ha="center",
                    va="center",
                    fontsize=8
                )

    plt.figtext(
        0.01,
        -0.01,
        "* Fisher exact test remained significant after FDR correction; "
        "no individual pair met this criterion in the current analysis.",
        ha="left",
        fontsize=9
    )
    plt.tight_layout()

    appendix_figure1_path = OUTPUT_DIR / "AA3_Appendix_Figure_A1_phi_heatmap_q34_q31.png"
    plt.savefig(appendix_figure1_path, dpi=300, bbox_inches="tight")
    plt.show()

    print("\nAppendix Figure A1 saved to:")
    print(appendix_figure1_path)

if GENERATE_APPENDIX_GRAPHS:
    top_pairs_plot = top_abs_phi.copy()
    top_pairs_plot["pair_label"] = (
        top_pairs_plot["funding_label"]
        + " -> "
        + top_pairs_plot["motive_label"]
    )
    top_pairs_plot = top_pairs_plot.head(12).sort_values("phi", ascending=True)

    plt.figure(figsize=(13, 8))
    plt.barh(top_pairs_plot["pair_label"], top_pairs_plot["phi"])
    plt.axvline(0, linewidth=1)
    plt.xlabel(r"Phi coefficient ($\phi$)", fontsize=12)
    plt.ylabel("Funding source -> improvement motive", fontsize=12)
    plt.title("Largest exploratory individual funding-motive associations", fontsize=14)

    for i, row in enumerate(top_pairs_plot.itertuples()):
        label = rf"$\phi$={row.phi:.3f}, FDR $p$={row.fisher_p_fdr:.3f}"
        plt.text(row.phi + 0.01, i, label, va="center", fontsize=9)

    plt.figtext(
        0.01,
        -0.01,
        "Note: None of the displayed individual associations remained "
        "statistically significant after FDR correction.",
        ha="left",
        fontsize=9
    )
    plt.tight_layout()

    appendix_figure2_path = OUTPUT_DIR / "AA3_Appendix_Figure_A2_top_exploratory_phi_pairs.png"
    plt.savefig(appendix_figure2_path, dpi=300, bbox_inches="tight")
    plt.show()

    print("\nAppendix Figure A2 saved to:")
    print(appendix_figure2_path)

summary_rows = []

summary_rows.append({
    "analysis": "External_funding_count with Improvement_motivators_count",
    "analysis_role": "Primary supplementary correlation",
    "sample_information": f"n={sample_n}",
    "estimate": (
        f"Spearman rho={spearman_rho:.3f}; "
        f"Kendall tau-b={kendall_tau_b:.3f}"
    ),
    "statistical_evidence": (
        f"p_rho={spearman_p:.4f}; "
        f"p_tau={kendall_p:.4f}"
    ),
    "interpretation": (
        "More external funding sources were associated with more "
        "positive improvement motives."
    )
})

for indicator in selected_funding_indicators:
    row = funding_group_compare[
        funding_group_compare["funding_indicator"] == indicator
    ].iloc[0]
    label = funding_indicator_labels[indicator]

    summary_rows.append({
        "analysis": f"{label}: selected vs not selected",
        "analysis_role": "Supplementary group comparison",
        "sample_information": (
            f"selected n={int(row['n_selected'])}; "
            f"not selected n={int(row['n_not_selected'])}"
        ),
        "estimate": (
            f"mean difference={row['mean_difference_selected_minus_not']:.3f}; "
            f"Cliff's delta={row['cliffs_delta']:.3f}"
        ),
        "statistical_evidence": f"FDR p={row['mannwhitney_p_fdr']:.4f}",
        "interpretation": (
            "Statistically supported difference."
            if bool(row["significant_fdr_0_05"])
            else "Difference was not statistically supported."
        )
    })

n_pairwise_tests = len(pairwise_q34_q31)
n_pairwise_significant_fdr = int(pairwise_q34_q31["significant_fdr_0_05"].sum())
strongest_pair = (
    pairwise_q34_q31
    .assign(abs_phi=pairwise_q34_q31["phi"].abs())
    .sort_values(["abs_phi", "fisher_p"], ascending=[False, True])
    .iloc[0]
)

summary_rows.append({
    "analysis": "Individual funding-source x motive associations",
    "analysis_role": "Exploratory disaggregated analysis",
    "sample_information": f"{n_pairwise_tests} tested pairs",
    "estimate": (
        f"Strongest pair: {strongest_pair['funding_label']} x "
        f"{strongest_pair['motive_label']}; "
        f"phi={strongest_pair['phi']:.3f}"
    ),
    "statistical_evidence": (
        f"raw p={strongest_pair['fisher_p']:.4f}; "
        f"FDR p={strongest_pair['fisher_p_fdr']:.4f}"
    ),
    "interpretation": (
        f"{n_pairwise_significant_fdr} individual pairs remained "
        "significant after FDR correction."
    )
})

final_summary = pd.DataFrame(summary_rows)

print("\n4) Final summary table for manuscript text:")
print(final_summary.to_string(index=False))

bubble_data_export = bubble_data.copy()
bubble_data_export["percentage_of_sample"] = (
    bubble_data_export["firm_count"] / len(df) * 100
).round(2)
bubble_data_export = bubble_data_export.rename(
    columns={
        "External_funding_count": "external_funding_count",
        "Improvement_motivators_count": "improvement_motivators_count",
        "firm_count": "number_of_firms"
    }
)

print("\n5) Data used in the main bubble plot:")
print(bubble_data_export.to_string(index=False))

with pd.ExcelWriter(SUMMARY_OUTPUT_PATH, engine="openpyxl") as writer:
    final_summary.to_excel(writer, sheet_name="final_summary", index=False)
    bubble_data_export.to_excel(writer, sheet_name="main_figure_bubble_data", index=False)
    main_correlation.to_excel(writer, sheet_name="main_correlation", index=False)
    group_plot.to_excel(writer, sheet_name="main_group_comparisons", index=False)
    aggregate_correlations.to_excel(writer, sheet_name="aggregate_correlations", index=False)
    pairwise_q34_q31.to_excel(writer, sheet_name="pairwise_q34_q31", index=False)
    top_abs_phi.to_excel(writer, sheet_name="top_abs_phi", index=False)
    phi_matrix.to_excel(writer, sheet_name="phi_matrix")
    fdr_p_matrix.to_excel(writer, sheet_name="fdr_p_matrix")
    risk_diff_matrix.to_excel(writer, sheet_name="risk_diff_matrix")

generated_figures = [
    {
        "figure": "Figure 1",
        "role": "Main text - primary association",
        "path": str(figure1_path)
    },
    {
        "figure": "Figure 2",
        "role": "Main text - supplementary comparison",
        "path": str(figure2_path)
    },
]

if GENERATE_APPENDIX_GRAPHS:
    generated_figures.extend([
        {
            "figure": "Appendix Figure A1",
            "role": "Appendix - exploratory phi heatmap",
            "path": str(appendix_figure1_path)
        },
        {
            "figure": "Appendix Figure A2",
            "role": "Appendix - strongest exploratory pairs",
            "path": str(appendix_figure2_path)
        }
    ])

generated_figures_df = pd.DataFrame(generated_figures)

print("\n6) Summary tables saved to:")
print(SUMMARY_OUTPUT_PATH)

print("\n7) Generated figures:")
print(generated_figures_df.to_string(index=False))

figure1_caption = (
    "Figure X. Relationship between the number of external funding sources "
    "used and the number of positive improvement motives reported. "
    "Note: Bubble size and the value within each bubble represent the number "
    "of firms reporting each observed combination of the two count-based "
    "indices. The association was positive and statistically significant "
    f"according to Spearman's rank correlation "
    f"(rho_S={spearman_rho:.3f}, p={spearman_p:.4f}) and Kendall's tau-b "
    f"(tau_b={kendall_tau_b:.3f}, p={kendall_p:.4f})."
)

figure2_caption = (
    "Figure Y. Mean number of positive improvement motives by selected "
    "external funding categories. Note: Bars compare firms reporting each "
    "funding category with firms not reporting it. An asterisk indicates "
    "a statistically significant Mann-Whitney comparison after FDR correction."
)

appendix_caption_a1 = (
    "Appendix Figure A1. Exploratory pairwise associations between specific "
    "external funding sources and specific positive improvement motives. "
    "Note: Cells represent phi coefficients. No individual association "
    "remained statistically significant after FDR correction."
)

appendix_caption_a2 = (
    "Appendix Figure A2. Largest exploratory individual funding-source-motive "
    "associations according to the absolute phi coefficient. "
    "Note: None of the displayed individual associations remained "
    "statistically significant after FDR correction."
)

print("\n8) Suggested caption for Figure 1:")
print(figure1_caption)
print("\n9) Suggested caption for Figure 2:")
print(figure2_caption)

if GENERATE_APPENDIX_GRAPHS:
    print("\n10) Suggested caption for Appendix Figure A1:")
    print(appendix_caption_a1)
    print("\n11) Suggested caption for Appendix Figure A2:")
    print(appendix_caption_a2)

print("\n" + "=" * 100)
print("Supplementary Correlation - Step 4 completed.")
print("Figures 1 and 2 are recommended for the main text.")
print("Appendix Figures A1 and A2 should be used only as supplementary documentation.")
print("=" * 100)
