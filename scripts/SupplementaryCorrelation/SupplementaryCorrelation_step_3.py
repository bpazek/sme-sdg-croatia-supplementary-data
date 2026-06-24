"""
Supplementary correlation analysis, Step 3: Association analysis between external funding and improvement motives.

This script supports the supplementary extension of RQ2. It examines whether
external funding use in the previous three years is associated with motives for
implemented improvements.

The script uses the derived q31/q34 indices created in Step 2. It computes
aggregate rank correlations, group comparisons of improvement motives by funding
indicators, exploratory pairwise associations between specific funding sources
and specific motives, and aggregate binary associations between funding
indicators and positive improvement motivation.

Spearman correlation is used as the main aggregate rank-based measure. Kendall's
tau-b is reported as a supplementary rank-based coefficient, and Fisher's exact
test is used for pairwise binary associations. Benjamini-Hochberg FDR correction
is applied where multiple tests are conducted.

Input:
    Additional_Analysis_3_q31_q34_indices.xlsx, sheet "encoded_with_q31_q34".

Output:
    Additional_Analysis_3_q31_q34_results_with_kendall.xlsx.
"""

import numpy as np
import pandas as pd
import scipy.stats as stats


INPUT_PATH = "/content/Additional_Analysis_3_q31_q34_indices.xlsx"
INPUT_SHEET = "encoded_with_q31_q34"
OUTPUT_PATH = "/content/Additional_Analysis_3_q31_q34_results_with_kendall.xlsx"

Q31_POSITIVE_MOTIVE_COLS = [
    "q31_motivi_promjena__regulatorni_okviri",
    "q31_motivi_promjena__zbog_konkurencije",
    "q31_motivi_promjena__ideja_od_partnera",
    "q31_motivi_promjena__bespovratna_sredstva",
    "q31_motivi_promjena__briga_za_okolis",
    "q31_motivi_promjena__energetska_ucinkovitost",
    "q31_motivi_promjena__drustveni_ciljevi",
    "q31_motivi_promjena__poslovni_razvoj",
    "q31_motivi_promjena__moderne_tehnologije_inovacije",
    "q31_motivi_promjena__covid_19",
    "q31_motivi_promjena__other_selected",
]

Q31_LABELS = {
    "q31_motivi_promjena__nista_poboljsano": "Nothing improved",
    "q31_motivi_promjena__regulatorni_okviri": "Regulatory frameworks",
    "q31_motivi_promjena__zbog_konkurencije": "Competition",
    "q31_motivi_promjena__ideja_od_partnera": "Partner idea",
    "q31_motivi_promjena__bespovratna_sredstva": "Grants",
    "q31_motivi_promjena__briga_za_okolis": "Environmental concern",
    "q31_motivi_promjena__energetska_ucinkovitost": "Energy efficiency",
    "q31_motivi_promjena__drustveni_ciljevi": "Social goals",
    "q31_motivi_promjena__poslovni_razvoj": "Business development",
    "q31_motivi_promjena__moderne_tehnologije_inovacije": "Modern technologies and innovation",
    "q31_motivi_promjena__covid_19": "COVID-19",
    "q31_motivi_promjena__other_selected": "Other motive",
}

Q34_POSITIVE_FUNDING_COLS = [
    "q34_eksterna_sredstva_3_godine__eu_konkurentnost",
    "q34_eksterna_sredstva_3_godine__eu_energetska_ucinkovitost",
    "q34_eksterna_sredstva_3_godine__npoo",
    "q34_eksterna_sredstva_3_godine__lokalni_poticaji",
    "q34_eksterna_sredstva_3_godine__banke_sufinancirana_kamata",
    "q34_eksterna_sredstva_3_godine__hbor",
    "q34_eksterna_sredstva_3_godine__ruralni_razvoj",
    "q34_eksterna_sredstva_3_godine__hzz_zeleno_digitalno",
    "q34_eksterna_sredstva_3_godine__hzz_zaposljavanje",
    "q34_eksterna_sredstva_3_godine__zaposljavanje_osoba_s_invaliditetom",
    "q34_eksterna_sredstva_3_godine__efop",
    "q34_eksterna_sredstva_3_godine__drugi_eu_fondovi",
    "q34_eksterna_sredstva_3_godine__other_selected",
]

Q34_LABELS = {
    "q34_eksterna_sredstva_3_godine__bez_eksternih_sredstava": "No external funding",
    "q34_eksterna_sredstva_3_godine__eu_konkurentnost": "EU competitiveness funds",
    "q34_eksterna_sredstva_3_godine__eu_energetska_ucinkovitost": "EU energy efficiency funds",
    "q34_eksterna_sredstva_3_godine__npoo": "NPOO",
    "q34_eksterna_sredstva_3_godine__lokalni_poticaji": "Local incentives",
    "q34_eksterna_sredstva_3_godine__banke_sufinancirana_kamata": "Bank co-financed interest",
    "q34_eksterna_sredstva_3_godine__hbor": "HBOR",
    "q34_eksterna_sredstva_3_godine__ruralni_razvoj": "Rural development",
    "q34_eksterna_sredstva_3_godine__hzz_zeleno_digitalno": "HZZ green/digital",
    "q34_eksterna_sredstva_3_godine__hzz_zaposljavanje": "HZZ employment",
    "q34_eksterna_sredstva_3_godine__zaposljavanje_osoba_s_invaliditetom": "Employment of persons with disabilities",
    "q34_eksterna_sredstva_3_godine__efop": "EFOP",
    "q34_eksterna_sredstva_3_godine__drugi_eu_fondovi": "Other EU funds",
    "q34_eksterna_sredstva_3_godine__other_selected": "Other external funding",
}


def benjamini_hochberg(p_values):
    """Apply Benjamini-Hochberg FDR correction while preserving NaN values."""
    p_values = np.asarray(p_values, dtype=float)
    adjusted = np.full(len(p_values), np.nan)
    valid_mask = ~np.isnan(p_values)
    valid_p_values = p_values[valid_mask]

    if len(valid_p_values) == 0:
        return adjusted

    n_values = len(valid_p_values)
    order = np.argsort(valid_p_values)
    ranked_p_values = valid_p_values[order]
    adjusted_valid = np.empty(n_values, dtype=float)
    cumulative_minimum = 1.0

    for i in range(n_values - 1, -1, -1):
        rank = i + 1
        adjusted_value = ranked_p_values[i] * n_values / rank
        cumulative_minimum = min(cumulative_minimum, adjusted_value)
        adjusted_valid[order[i]] = cumulative_minimum

    adjusted[valid_mask] = np.minimum(adjusted_valid, 1.0)
    return adjusted


def phi_coefficient_from_table(a, b, c, d):
    """Calculate the phi coefficient from a 2x2 contingency table."""
    denominator = np.sqrt((a + b) * (c + d) * (a + c) * (b + d))
    if denominator == 0:
        return np.nan
    return (a * d - b * c) / denominator


def cliffs_delta(x, y):
    """Calculate Cliff's delta for two independent groups."""
    x = np.asarray(x)
    y = np.asarray(y)
    if len(x) == 0 or len(y) == 0:
        return np.nan
    greater = 0
    less = 0
    for value in x:
        greater += np.sum(value > y)
        less += np.sum(value < y)
    return (greater - less) / (len(x) * len(y))


def cliffs_delta_magnitude(delta):
    """Classify the magnitude of Cliff's delta."""
    if pd.isna(delta):
        return np.nan
    absolute_delta = abs(delta)
    if absolute_delta < 0.147:
        return "negligible"
    if absolute_delta < 0.330:
        return "small"
    if absolute_delta < 0.474:
        return "medium"
    return "large"


def round_dataframe(data):
    """Round selected numeric columns for readable output."""
    output = data.copy()
    columns_to_round = [
        "spearman_rho", "spearman_p", "spearman_p_fdr",
        "kendall_tau_b", "kendall_p", "kendall_p_fdr",
        "pearson_r", "pearson_p", "mean_motivators_selected",
        "mean_motivators_not_selected", "median_motivators_selected",
        "median_motivators_not_selected", "mean_difference_selected_minus_not",
        "mannwhitney_p", "mannwhitney_p_fdr", "cliffs_delta",
        "p_motive_given_funding", "p_motive_given_no_funding",
        "risk_difference", "phi", "odds_ratio", "fisher_p",
        "fisher_p_fdr", "chi2", "chi2_p", "min_expected",
    ]
    for column in columns_to_round:
        if column in output.columns:
            output[column] = output[column].round(4)
    return output


def validate_required_columns(data, columns):
    """Raise an error if expected columns are missing."""
    missing_columns = [column for column in columns if column not in data.columns]
    if missing_columns:
        raise ValueError("Missing required columns: " + ", ".join(missing_columns))


def main():
    df = pd.read_excel(INPUT_PATH, sheet_name=INPUT_SHEET)

    print("Supplementary correlation analysis - Step 3")
    print("Correlations and associations between external funding and improvement motives")
    print("=" * 110)
    print("Dataset dimensions:", df.shape)
    print("Number of respondents / firms:", len(df))

    required_derived_columns = [
        "External_funding_count", "Any_external_funding",
        "No_external_funding_indicator", "Any_EU_or_public_funding",
        "Any_bank_or_credit_support", "Improvement_motivators_count",
        "Any_positive_improvement_motivation", "No_improvement_indicator",
    ]
    validate_required_columns(df, required_derived_columns)

    q34_pairwise_cols = [
        column for column in Q34_POSITIVE_FUNDING_COLS
        if column in df.columns and df[column].nunique(dropna=True) > 1
    ]
    q31_pairwise_cols = [
        column for column in Q31_POSITIVE_MOTIVE_COLS
        if column in df.columns and df[column].nunique(dropna=True) > 1
    ]

    print("\n1) q34 funding variables included in pairwise analysis:")
    for column in q34_pairwise_cols:
        print(f" - {Q34_LABELS[column]}: n={int(df[column].sum())}")

    print("\n2) q31 motive variables included in pairwise analysis:")
    for column in q31_pairwise_cols:
        print(f" - {Q31_LABELS[column]}: n={int(df[column].sum())}")

    excluded_q34 = [
        column for column in Q34_POSITIVE_FUNDING_COLS
        if column in df.columns and column not in q34_pairwise_cols
    ]
    excluded_q31 = [
        column for column in Q31_POSITIVE_MOTIVE_COLS
        if column in df.columns and column not in q31_pairwise_cols
    ]

    print("\nExcluded q34 variables due to zero variance:")
    print([Q34_LABELS[column] for column in excluded_q34])
    print("\nExcluded q31 variables due to zero variance:")
    print([Q31_LABELS[column] for column in excluded_q31])

    aggregate_pairs = [
        ("External_funding_count", "Improvement_motivators_count"),
        ("Any_external_funding", "Improvement_motivators_count"),
        ("No_external_funding_indicator", "Improvement_motivators_count"),
        ("Any_EU_or_public_funding", "Improvement_motivators_count"),
        ("Any_bank_or_credit_support", "Improvement_motivators_count"),
        ("External_funding_count", "Any_positive_improvement_motivation"),
        ("External_funding_count", "No_improvement_indicator"),
        ("Any_external_funding", "Any_positive_improvement_motivation"),
        ("Any_external_funding", "No_improvement_indicator"),
    ]

    aggregate_rows = []
    for x_column, y_column in aggregate_pairs:
        x = df[x_column]
        y = df[y_column]
        if x.nunique(dropna=True) < 2 or y.nunique(dropna=True) < 2:
            spearman_rho = spearman_p = kendall_tau_b = kendall_p = pearson_r = pearson_p = np.nan
        else:
            spearman_rho, spearman_p = stats.spearmanr(x, y, nan_policy="omit")
            kendall_tau_b, kendall_p = stats.kendalltau(x, y, nan_policy="omit")
            pearson_r, pearson_p = stats.pearsonr(x, y)
        aggregate_rows.append({
            "x_variable": x_column,
            "y_variable": y_column,
            "n": len(df),
            "spearman_rho": spearman_rho,
            "spearman_p": spearman_p,
            "kendall_tau_b": kendall_tau_b,
            "kendall_p": kendall_p,
            "pearson_r": pearson_r,
            "pearson_p": pearson_p,
        })

    aggregate_correlations = pd.DataFrame(aggregate_rows)
    aggregate_correlations["spearman_p_fdr"] = benjamini_hochberg(aggregate_correlations["spearman_p"].values)
    aggregate_correlations["spearman_significant_fdr_0_05"] = aggregate_correlations["spearman_p_fdr"] < 0.05
    aggregate_correlations["kendall_p_fdr"] = benjamini_hochberg(aggregate_correlations["kendall_p"].values)
    aggregate_correlations["kendall_significant_fdr_0_05"] = aggregate_correlations["kendall_p_fdr"] < 0.05

    main_correlation = aggregate_correlations[
        (aggregate_correlations["x_variable"] == "External_funding_count")
        & (aggregate_correlations["y_variable"] == "Improvement_motivators_count")
    ].copy()
    main_correlation["analysis_role"] = "Primary supplementary correlation specified for Section 3.9"

    funding_binary_predictors = [
        "Any_external_funding",
        "No_external_funding_indicator",
        "Any_EU_or_public_funding",
        "Any_bank_or_credit_support",
    ]

    group_rows = []
    for predictor in funding_binary_predictors:
        selected = df.loc[df[predictor] == 1, "Improvement_motivators_count"]
        not_selected = df.loc[df[predictor] == 0, "Improvement_motivators_count"]
        mannwhitney_u, mannwhitney_p = stats.mannwhitneyu(selected, not_selected, alternative="two-sided")
        delta = cliffs_delta(selected, not_selected)
        group_rows.append({
            "funding_indicator": predictor,
            "n_selected": len(selected),
            "n_not_selected": len(not_selected),
            "mean_motivators_selected": selected.mean(),
            "mean_motivators_not_selected": not_selected.mean(),
            "median_motivators_selected": selected.median(),
            "median_motivators_not_selected": not_selected.median(),
            "mean_difference_selected_minus_not": selected.mean() - not_selected.mean(),
            "mannwhitney_u": mannwhitney_u,
            "mannwhitney_p": mannwhitney_p,
            "cliffs_delta": delta,
            "cliffs_delta_magnitude": cliffs_delta_magnitude(delta),
        })

    funding_group_comparisons = pd.DataFrame(group_rows)
    funding_group_comparisons["mannwhitney_p_fdr"] = benjamini_hochberg(funding_group_comparisons["mannwhitney_p"].values)
    funding_group_comparisons["significant_fdr_0_05"] = funding_group_comparisons["mannwhitney_p_fdr"] < 0.05

    pairwise_rows = []
    for funding_column in q34_pairwise_cols:
        for motive_column in q31_pairwise_cols:
            funding = df[funding_column].astype(int)
            motive = df[motive_column].astype(int)
            a = int(((funding == 1) & (motive == 1)).sum())
            b = int(((funding == 1) & (motive == 0)).sum())
            c = int(((funding == 0) & (motive == 1)).sum())
            d = int(((funding == 0) & (motive == 0)).sum())
            contingency_table = np.array([[a, b], [c, d]])
            phi = phi_coefficient_from_table(a, b, c, d)
            odds_ratio, fisher_p = stats.fisher_exact(contingency_table, alternative="two-sided")
            try:
                chi2, chi2_p, _, expected = stats.chi2_contingency(contingency_table, correction=False)
                min_expected = expected.min()
            except Exception:
                chi2 = chi2_p = min_expected = np.nan
            p_motive_given_funding = a / (a + b) if (a + b) > 0 else np.nan
            p_motive_given_no_funding = c / (c + d) if (c + d) > 0 else np.nan
            pairwise_rows.append({
                "funding_variable": funding_column,
                "funding_label": Q34_LABELS[funding_column],
                "motive_variable": motive_column,
                "motive_label": Q31_LABELS[motive_column],
                "a_both_1": a,
                "b_funding1_motive0": b,
                "c_funding0_motive1": c,
                "d_both_0": d,
                "funding_n": int(funding.sum()),
                "motive_n": int(motive.sum()),
                "p_motive_given_funding": p_motive_given_funding,
                "p_motive_given_no_funding": p_motive_given_no_funding,
                "risk_difference": p_motive_given_funding - p_motive_given_no_funding,
                "phi": phi,
                "odds_ratio": odds_ratio,
                "fisher_p": fisher_p,
                "chi2": chi2,
                "chi2_p": chi2_p,
                "min_expected": min_expected,
            })

    pairwise_associations = pd.DataFrame(pairwise_rows)
    pairwise_associations["fisher_p_fdr"] = benjamini_hochberg(pairwise_associations["fisher_p"].values)
    pairwise_associations["significant_fdr_0_05"] = pairwise_associations["fisher_p_fdr"] < 0.05

    none_pair_rows = []
    aggregate_funding_cols = [
        "Any_external_funding",
        "No_external_funding_indicator",
        "Any_EU_or_public_funding",
        "Any_bank_or_credit_support",
    ]
    aggregate_motive_cols = [
        "Any_positive_improvement_motivation",
        "No_improvement_indicator",
    ]

    for funding_column in aggregate_funding_cols:
        for motive_column in aggregate_motive_cols:
            funding = df[funding_column].astype(int)
            motive = df[motive_column].astype(int)
            a = int(((funding == 1) & (motive == 1)).sum())
            b = int(((funding == 1) & (motive == 0)).sum())
            c = int(((funding == 0) & (motive == 1)).sum())
            d = int(((funding == 0) & (motive == 0)).sum())
            contingency_table = np.array([[a, b], [c, d]])
            phi = phi_coefficient_from_table(a, b, c, d)
            odds_ratio, fisher_p = stats.fisher_exact(contingency_table, alternative="two-sided")
            p_motive_given_funding = a / (a + b) if (a + b) > 0 else np.nan
            p_motive_given_no_funding = c / (c + d) if (c + d) > 0 else np.nan
            none_pair_rows.append({
                "funding_indicator": funding_column,
                "motive_indicator": motive_column,
                "a_both_1": a,
                "b_funding1_motive0": b,
                "c_funding0_motive1": c,
                "d_both_0": d,
                "p_motive_given_funding": p_motive_given_funding,
                "p_motive_given_no_funding": p_motive_given_no_funding,
                "risk_difference": p_motive_given_funding - p_motive_given_no_funding,
                "phi": phi,
                "odds_ratio": odds_ratio,
                "fisher_p": fisher_p,
            })

    none_pair_associations = pd.DataFrame(none_pair_rows)
    none_pair_associations["fisher_p_fdr"] = benjamini_hochberg(none_pair_associations["fisher_p"].values)
    none_pair_associations["significant_fdr_0_05"] = none_pair_associations["fisher_p_fdr"] < 0.05

    phi_matrix = pairwise_associations.pivot(index="funding_label", columns="motive_label", values="phi")
    fdr_p_matrix = pairwise_associations.pivot(index="funding_label", columns="motive_label", values="fisher_p_fdr")
    risk_difference_matrix = pairwise_associations.pivot(index="funding_label", columns="motive_label", values="risk_difference")

    aggregate_display = round_dataframe(aggregate_correlations)
    main_correlation_display = round_dataframe(main_correlation)
    group_display = round_dataframe(funding_group_comparisons)
    pairwise_display = round_dataframe(
        pairwise_associations.sort_values(["fisher_p_fdr", "fisher_p", "phi"], ascending=[True, True, False])
    )
    none_pair_display = round_dataframe(
        none_pair_associations.sort_values(["fisher_p_fdr", "fisher_p"], ascending=[True, True])
    )
    significant_pairwise = pairwise_display[pairwise_display["significant_fdr_0_05"]].copy()
    top_pairwise_by_abs_phi = round_dataframe(
        pairwise_associations.assign(abs_phi=lambda data: data["phi"].abs())
        .sort_values(["abs_phi", "fisher_p"], ascending=[False, True])
        .head(20)
    )

    print("\n3) Main predefined supplementary correlation:")
    print(main_correlation_display.to_string(index=False))
    print("\n4) All aggregate rank correlations:")
    print(aggregate_display.to_string(index=False))
    print("\n5) Group comparisons of positive improvement motives by funding indicators:")
    print(group_display.to_string(index=False))
    print("\n6) q34 x q31 pairwise associations significant after FDR correction:")
    if len(significant_pairwise) > 0:
        print(significant_pairwise.to_string(index=False))
    else:
        print("No q34 x q31 pair remained significant after FDR correction.")
    print("\n7) Top 20 pairs by absolute phi coefficient:")
    print(top_pairwise_by_abs_phi.to_string(index=False))
    print("\n8) Additional aggregate funding associations with positive motivation or no improvement:")
    print(none_pair_display.to_string(index=False))

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        main_correlation_display.to_excel(writer, sheet_name="main_correlation", index=False)
        aggregate_display.to_excel(writer, sheet_name="aggregate_correlations", index=False)
        group_display.to_excel(writer, sheet_name="funding_group_compare", index=False)
        pairwise_display.to_excel(writer, sheet_name="pairwise_q34_q31", index=False)
        significant_pairwise.to_excel(writer, sheet_name="significant_pairwise", index=False)
        top_pairwise_by_abs_phi.to_excel(writer, sheet_name="top_abs_phi", index=False)
        none_pair_display.to_excel(writer, sheet_name="aggregate_binary_pairs", index=False)
        phi_matrix.round(4).to_excel(writer, sheet_name="phi_matrix")
        fdr_p_matrix.round(4).to_excel(writer, sheet_name="fdr_p_matrix")
        risk_difference_matrix.round(4).to_excel(writer, sheet_name="risk_diff_matrix")

    print("\n9) Step 3 results saved to:")
    print(OUTPUT_PATH)
    print("\nSupplementary correlation analysis - Step 3 completed.")


if __name__ == "__main__":
    main()
