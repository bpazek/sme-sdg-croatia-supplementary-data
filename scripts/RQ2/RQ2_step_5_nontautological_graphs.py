"""
RQ2 Step 5: Nontautological visualization of RQ2 results.

This script generates the figure outputs for Research Question 2:
What are the most important motivators for SDG involvement in SME operations?

The script is aligned with the nontautological multivariate modelling approach
used in Step 4. The variable Sustainable_success_orientation is not presented as
a main motivator because it directly reflects whether the firm already includes
sustainable goals in measuring business success.

The script creates six figures:
1. Top active nontautological factors by frequency.
2. Binary factors associated with higher SDG_count.
3. Numeric or scalar predictors ranked by Spearman correlation with SDG_count.
4. Main nontautological Negative Binomial model, shown as IRR with 95% CI.
5. Nontautological logistic sensitivity model for SDG_count >= 2, shown as OR with 95% CI.
6. Exploratory Negative Binomial model of specific pressure sources, shown as IRR with 95% CI.

Input files:
    RQ2_step3a_frequency_rankings.xlsx
    RQ2_step3b_SDG_involvement_associations.xlsx
    RQ2_step4_nontautological_multivariate_models.xlsx

Output folder:
    RQ2_graphs_nontautological/
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


STEP3A_PATH = "/content/RQ2_step3a_frequency_rankings.xlsx"
STEP3B_PATH = "/content/RQ2_step3b_SDG_involvement_associations.xlsx"
STEP4_PATH = "/content/RQ2_step4_nontautological_multivariate_models.xlsx"
OUTPUT_DIR = Path("/content/RQ2_graphs_nontautological")


FIGURE_FILENAMES = {
    "frequency": "RQ2_01_top_nontautological_factors_frequency.png",
    "binary_difference": "RQ2_02_nontautological_binary_SDG_count_difference.png",
    "numeric_spearman": "RQ2_03_numeric_predictors_spearman_SDG_count.png",
    "negative_binomial": "RQ2_04_nontautological_negative_binomial_IRR.png",
    "logistic": "RQ2_05_nontautological_logistic_OR_SDG_count_ge_2.png",
    "pressure_sources": "RQ2_06_pressure_sources_negative_binomial_IRR.png",
}


EXCLUDED_NONT_AUTOLOGICAL_LABELS = [
    "Profit and sustainable goals",
]


def wrap_label(label, width=38):
    """Wrap long labels to improve readability in horizontal plots."""
    words = str(label).split()
    lines = []
    current_line = ""

    for word in words:
        if len(current_line) + len(word) + 1 <= width:
            current_line = (current_line + " " + word).strip()
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)

    return "\n".join(lines)


def significance_star(p_value):
    """Return an asterisk for p < 0.05."""
    if pd.notna(p_value) and p_value < 0.05:
        return "*"
    return ""


def safe_xlim_for_ratio(ci_low, ci_high, lower_margin=0.2, upper_margin=0.5):
    """Create safe x-axis limits for IRR and OR confidence interval plots."""
    xmin = max(0, ci_low.min() - lower_margin)
    xmax = ci_high.max() + upper_margin
    return xmin, xmax


def validate_columns(data, required_columns, table_name):
    """Raise an error if required columns are missing from a table."""
    missing_columns = [column for column in required_columns if column not in data.columns]
    if missing_columns:
        raise ValueError(f"Missing columns in {table_name}: {missing_columns}")


def configure_plot_style():
    """Set common Matplotlib parameters for the generated figures."""
    plt.rcParams.update({
        "font.size": 11,
        "axes.titlesize": 13,
        "axes.labelsize": 11,
        "xtick.labelsize": 10,
        "ytick.labelsize": 10,
        "legend.fontsize": 10,
    })


def load_inputs():
    """Load all Excel inputs needed for the RQ2 visualizations."""
    rq2_positive_indicators = pd.read_excel(
        STEP3A_PATH,
        sheet_name="positive_indicators",
    )
    binary_sdg_count = pd.read_excel(
        STEP3B_PATH,
        sheet_name="binary_SDG_count_reliable",
    )
    numeric_sdg_count = pd.read_excel(
        STEP3B_PATH,
        sheet_name="numeric_SDG_count_ranked",
    )
    count_model_coefficients = pd.read_excel(
        STEP4_PATH,
        sheet_name="count_model_coefficients",
    )
    logit_coefficients = pd.read_excel(
        STEP4_PATH,
        sheet_name="logit_coefficients",
    )
    excluded_tautological = pd.read_excel(
        STEP4_PATH,
        sheet_name="excluded_tautological",
    )

    validate_columns(
        rq2_positive_indicators,
        ["label", "selected_percent"],
        "positive_indicators",
    )
    validate_columns(
        binary_sdg_count,
        ["label", "mean_difference_selected_minus_not", "mannwhitney_p_fdr_by_outcome"],
        "binary_SDG_count_reliable",
    )
    validate_columns(
        numeric_sdg_count,
        ["label", "spearman_rho", "spearman_p_fdr_by_outcome"],
        "numeric_SDG_count_ranked",
    )
    validate_columns(
        count_model_coefficients,
        ["model", "term", "label", "IRR", "IRR_CI_low", "IRR_CI_high", "p_value"],
        "count_model_coefficients",
    )
    validate_columns(
        logit_coefficients,
        ["term", "label", "OR", "OR_CI_low", "OR_CI_high", "p_value"],
        "logit_coefficients",
    )

    return {
        "rq2_positive_indicators": rq2_positive_indicators,
        "binary_sdg_count": binary_sdg_count,
        "numeric_sdg_count": numeric_sdg_count,
        "count_model_coefficients": count_model_coefficients,
        "logit_coefficients": logit_coefficients,
        "excluded_tautological": excluded_tautological,
    }


def plot_top_active_factors(rq2_positive_indicators):
    """Create Figure 1: most frequent active nontautological factors."""
    plot_data = rq2_positive_indicators.copy()
    plot_data = plot_data[
        ~plot_data["label"].isin(EXCLUDED_NONT_AUTOLOGICAL_LABELS)
    ].copy()
    plot_data = plot_data.head(10).sort_values("selected_percent", ascending=True)
    plot_data["label_wrapped"] = plot_data["label"].apply(lambda x: wrap_label(x, width=34))

    plt.figure(figsize=(11, 7))
    plt.barh(plot_data["label_wrapped"], plot_data["selected_percent"])
    plt.xlabel("Share of SMEs selecting the indicator (%)")
    plt.ylabel("Motivational / institutional factor")
    plt.title("RQ2: Most frequent active factors")

    for index, value in enumerate(plot_data["selected_percent"]):
        plt.text(value + 0.8, index, f"{value:.1f}%", va="center")

    plt.xlim(0, max(plot_data["selected_percent"]) + 10)
    plt.tight_layout()

    figure_path = OUTPUT_DIR / FIGURE_FILENAMES["frequency"]
    plt.savefig(figure_path, dpi=300, bbox_inches="tight")
    plt.show()
    return figure_path


def plot_binary_sdg_count_difference(binary_sdg_count):
    """Create Figure 2: binary indicators with the largest positive SDG_count differences."""
    plot_data = binary_sdg_count.copy()
    plot_data = plot_data[
        ~plot_data["label"].isin(EXCLUDED_NONT_AUTOLOGICAL_LABELS)
    ].copy()
    plot_data = plot_data[plot_data["mean_difference_selected_minus_not"] > 0].copy()
    plot_data = plot_data.sort_values(
        "mean_difference_selected_minus_not",
        ascending=False,
    ).head(12)
    plot_data = plot_data.sort_values("mean_difference_selected_minus_not", ascending=True)
    plot_data["label_wrapped"] = plot_data["label"].apply(lambda x: wrap_label(x, width=36))

    plt.figure(figsize=(11, 7))
    plt.barh(plot_data["label_wrapped"], plot_data["mean_difference_selected_minus_not"])
    plt.xlabel("Difference in mean SDG_count: selected minus not selected")
    plt.ylabel("Binary factor")
    plt.title("RQ2: Binary factors associated with higher SDG_count")

    for position, (_, row) in enumerate(plot_data.iterrows()):
        value = row["mean_difference_selected_minus_not"]
        fdr_p = row["mannwhitney_p_fdr_by_outcome"]
        annotation = f"{value:.2f}{significance_star(fdr_p)}"
        plt.text(value + 0.05, position, annotation, va="center")

    plt.axvline(0, linewidth=1)
    plt.xlim(0, max(plot_data["mean_difference_selected_minus_not"]) + 0.6)
    plt.text(0, -1.0, "* FDR-adjusted Mann-Whitney p < 0.05", fontsize=9)
    plt.tight_layout()

    figure_path = OUTPUT_DIR / FIGURE_FILENAMES["binary_difference"]
    plt.savefig(figure_path, dpi=300, bbox_inches="tight")
    plt.show()
    return figure_path


def plot_numeric_spearman(numeric_sdg_count):
    """Create Figure 3: numeric predictors ranked by Spearman correlation with SDG_count."""
    plot_data = numeric_sdg_count.copy()
    plot_data = plot_data.sort_values("spearman_rho", ascending=False).head(10)
    plot_data = plot_data.sort_values("spearman_rho", ascending=True)
    plot_data["label_wrapped"] = plot_data["label"].apply(lambda x: wrap_label(x, width=36))

    plt.figure(figsize=(11, 7))
    plt.barh(plot_data["label_wrapped"], plot_data["spearman_rho"])
    plt.xlabel("Spearman correlation with SDG_count")
    plt.ylabel("Numeric / scalar predictor")
    plt.title("RQ2: Bivariate associations with SDG_count")

    for position, (_, row) in enumerate(plot_data.iterrows()):
        value = row["spearman_rho"]
        fdr_p = row["spearman_p_fdr_by_outcome"]
        annotation = f"{value:.3f}{significance_star(fdr_p)}"
        plt.text(value + 0.01, position, annotation, va="center")

    plt.axvline(0, linewidth=1)
    plt.xlim(
        min(0, plot_data["spearman_rho"].min() - 0.05),
        plot_data["spearman_rho"].max() + 0.08,
    )
    plt.text(0, -0.9, "* FDR-adjusted Spearman p < 0.05", fontsize=9)
    plt.tight_layout()

    figure_path = OUTPUT_DIR / FIGURE_FILENAMES["numeric_spearman"]
    plt.savefig(figure_path, dpi=300, bbox_inches="tight")
    plt.show()
    return figure_path


def plot_negative_binomial_irr(count_model_coefficients):
    """Create Figure 4: main nontautological Negative Binomial model with IRR and 95% CI."""
    plot_data = count_model_coefficients.copy()
    plot_data = plot_data[
        (plot_data["model"] == "Negative Binomial main, non-tautological")
        & (~plot_data["term"].isin(["const", "alpha", "Intercept"]))
    ].copy()
    plot_data = plot_data.sort_values("IRR", ascending=True)
    plot_data["label_wrapped"] = plot_data["label"].apply(lambda x: wrap_label(x, width=34))

    plt.figure(figsize=(10.5, 6.2))
    y_positions = np.arange(len(plot_data))
    plt.errorbar(
        plot_data["IRR"],
        y_positions,
        xerr=[
            plot_data["IRR"] - plot_data["IRR_CI_low"],
            plot_data["IRR_CI_high"] - plot_data["IRR"],
        ],
        fmt="o",
        capsize=4,
    )
    plt.axvline(1, linewidth=1)
    plt.yticks(y_positions, plot_data["label_wrapped"])
    plt.xlabel("Incidence Rate Ratio (IRR), 95% CI")
    plt.ylabel("Predictor")
    plt.title("RQ2: Negative Binomial model of SDG_count")

    for i, row in enumerate(plot_data.itertuples()):
        annotation = f"IRR={row.IRR:.2f}{significance_star(row.p_value)}"
        plt.text(row.IRR_CI_high + 0.04, i, annotation, va="center")

    xmin, xmax = safe_xlim_for_ratio(
        plot_data["IRR_CI_low"],
        plot_data["IRR_CI_high"],
        lower_margin=0.15,
        upper_margin=0.45,
    )
    plt.xlim(xmin, xmax)
    plt.text(xmin + 0.01, -0.13, "* p < 0.05", fontsize=9)
    plt.tight_layout()

    figure_path = OUTPUT_DIR / FIGURE_FILENAMES["negative_binomial"]
    plt.savefig(figure_path, dpi=300, bbox_inches="tight")
    plt.show()
    return figure_path


def plot_logistic_or(logit_coefficients):
    """Create Figure 5: nontautological logistic sensitivity model with OR and 95% CI."""
    plot_data = logit_coefficients.copy()
    plot_data = plot_data[~plot_data["term"].isin(["Intercept", "const"])].copy()
    plot_data = plot_data.sort_values("OR", ascending=True)
    plot_data["label_wrapped"] = plot_data["label"].apply(lambda x: wrap_label(x, width=34))

    plt.figure(figsize=(10.5, 6.2))
    y_positions = np.arange(len(plot_data))
    plt.errorbar(
        plot_data["OR"],
        y_positions,
        xerr=[
            plot_data["OR"] - plot_data["OR_CI_low"],
            plot_data["OR_CI_high"] - plot_data["OR"],
        ],
        fmt="o",
        capsize=4,
    )
    plt.axvline(1, linewidth=1)
    plt.yticks(y_positions, plot_data["label_wrapped"])
    plt.xlabel("Odds Ratio (OR), 95% CI")
    plt.ylabel("Predictor")
    plt.title("RQ2: Logistic model for SDG_count >= 2")

    for i, row in enumerate(plot_data.itertuples()):
        annotation = f"OR={row.OR:.2f}{significance_star(row.p_value)}"
        plt.text(row.OR_CI_high + 0.08, i, annotation, va="center")

    xmin, xmax = safe_xlim_for_ratio(
        plot_data["OR_CI_low"],
        plot_data["OR_CI_high"],
        lower_margin=0.20,
        upper_margin=0.70,
    )
    plt.xlim(xmin, xmax)
    plt.text(xmin + 0.01, -0.13, "* p < 0.05", fontsize=9)
    plt.tight_layout()

    figure_path = OUTPUT_DIR / FIGURE_FILENAMES["logistic"]
    plt.savefig(figure_path, dpi=300, bbox_inches="tight")
    plt.show()
    return figure_path


def plot_pressure_sources_irr(count_model_coefficients):
    """Create Figure 6: exploratory model of specific pressure sources with IRR and 95% CI."""
    plot_data = count_model_coefficients.copy()
    plot_data = plot_data[
        (plot_data["model"] == "Negative Binomial pressure sources")
        & (~plot_data["term"].isin(["const", "alpha", "Intercept"]))
    ].copy()
    plot_data = plot_data.sort_values("IRR", ascending=True)
    plot_data["label_wrapped"] = plot_data["label"].apply(lambda x: wrap_label(x, width=34))

    plt.figure(figsize=(10.5, 7))
    y_positions = np.arange(len(plot_data))
    plt.errorbar(
        plot_data["IRR"],
        y_positions,
        xerr=[
            plot_data["IRR"] - plot_data["IRR_CI_low"],
            plot_data["IRR_CI_high"] - plot_data["IRR"],
        ],
        fmt="o",
        capsize=4,
    )
    plt.axvline(1, linewidth=1)
    plt.yticks(y_positions, plot_data["label_wrapped"])
    plt.xlabel("Incidence Rate Ratio (IRR), 95% CI")
    plt.ylabel("Predictor")
    plt.title("RQ2: Exploratory model of specific pressure sources")

    for i, row in enumerate(plot_data.itertuples()):
        annotation = f"IRR={row.IRR:.2f}{significance_star(row.p_value)}"
        plt.text(row.IRR_CI_high + 0.05, i, annotation, va="center")

    xmin, xmax = safe_xlim_for_ratio(
        plot_data["IRR_CI_low"],
        plot_data["IRR_CI_high"],
        lower_margin=0.15,
        upper_margin=0.75,
    )
    plt.xlim(xmin, xmax)
    plt.text(xmin, -0.25, "* p < 0.05", fontsize=9)
    plt.tight_layout()

    figure_path = OUTPUT_DIR / FIGURE_FILENAMES["pressure_sources"]
    plt.savefig(figure_path, dpi=300, bbox_inches="tight")
    plt.show()
    return figure_path


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    configure_plot_style()

    print("RQ2 - STEP 5")
    print("Visualization of nontautological RQ2 results")
    print("=" * 100)

    inputs = load_inputs()

    print("\nLoaded tables:")
    print("positive_indicators:", inputs["rq2_positive_indicators"].shape)
    print("binary_SDG_count_reliable:", inputs["binary_sdg_count"].shape)
    print("numeric_SDG_count_ranked:", inputs["numeric_sdg_count"].shape)
    print("count_model_coefficients:", inputs["count_model_coefficients"].shape)
    print("logit_coefficients:", inputs["logit_coefficients"].shape)
    print("excluded_tautological:", inputs["excluded_tautological"].shape)

    print("\nVariables intentionally excluded as conceptually too close:")
    print(inputs["excluded_tautological"].to_string(index=False))

    print("\nFigures will be saved to:")
    print(OUTPUT_DIR)

    figure_paths = [
        plot_top_active_factors(inputs["rq2_positive_indicators"]),
        plot_binary_sdg_count_difference(inputs["binary_sdg_count"]),
        plot_numeric_spearman(inputs["numeric_sdg_count"]),
        plot_negative_binomial_irr(inputs["count_model_coefficients"]),
        plot_logistic_or(inputs["logit_coefficients"]),
        plot_pressure_sources_irr(inputs["count_model_coefficients"]),
    ]

    print("\n" + "=" * 100)
    print("RQ2 - STEP 5 COMPLETED")
    print("=" * 100)
    print("Six figures were generated:")
    print("1. Top active nontautological factors by frequency")
    print("2. Nontautological binary factors associated with higher SDG_count")
    print("3. Numeric or scalar predictors ranked by Spearman correlation with SDG_count")
    print("4. Main nontautological Negative Binomial model: IRR with 95% confidence intervals")
    print("5. Nontautological logistic sensitivity model: OR with 95% confidence intervals")
    print("6. Exploratory Negative Binomial model of specific pressure sources: IRR with 95% confidence intervals")

    print("\nSaved files:")
    for figure_path in figure_paths:
        print(figure_path)


if __name__ == "__main__":
    main()
