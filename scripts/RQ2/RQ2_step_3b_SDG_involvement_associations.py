"""
RQ2 Step 3b: Association analysis between RQ2 factors and SDG involvement.

This script examines how motivational, institutional, resource-related and
value-oriented factors are associated with SDG involvement among Croatian SMEs.
It uses the validated RQ2 index file created in Step 2 and analyses three SDG
outcomes: SDG_count as the primary outcome, and SDG_mean_relevance and
SDG_high_relevance_count as secondary outcomes.

For binary indicators, the script compares SDG outcomes between firms that
selected and did not select each indicator using Mann-Whitney U tests, Spearman
correlations and Cliff's delta. For numeric and ordinal predictors, the script
computes Spearman and Kendall rank correlations. FDR correction is applied
separately by outcome. The script exports all association results and ranked
SDG_count summaries to an Excel file.

Input:
    RQ2_step2_indices.xlsx, sheet "encoded_with_RQ2_indices".

Output:
    RQ2_step3b_SDG_involvement_associations.xlsx.
"""

import pandas as pd
import numpy as np
import scipy.stats as stats

print("RQ2 - STEP 3b")
print("Associations between RQ2 factors and SDG involvement")
print("=" * 100)


input_path = "/content/RQ2_step2_indices.xlsx"

coded_data = pd.read_excel(
    input_path,
    sheet_name="encoded_with_RQ2_indices"
)

print("Loaded table:")
print(input_path)
print("Table dimensions:", coded_data.shape)
print("Number of respondents / firms:", len(coded_data))


sdg_outcome_cols = [
    "SDG_count",
    "SDG_mean_relevance",
    "SDG_high_relevance_count"
]

missing_sdg_outcomes = [
    col for col in sdg_outcome_cols
    if col not in coded_data.columns
]

if missing_sdg_outcomes:
    raise ValueError(
        f"Missing SDG outcomes from previous steps: {missing_sdg_outcomes}"
    )

primary_outcome = "SDG_count"


q15_cols = [
    col for col in coded_data.columns
    if col.startswith("q15_pritisak_na_odrzivost__")
]

q27_cols = [
    col for col in coded_data.columns
    if col.startswith("q27_pomoc_za_odrzivost__")
    and not col.endswith("__other_text")
]

q31_cols = [
    col for col in coded_data.columns
    if col.startswith("q31_motivi_promjena__")
    and not col.endswith("__other_text")
]

q34_cols = [
    col for col in coded_data.columns
    if col.startswith("q34_eksterna_sredstva_3_godine__")
    and not col.endswith("__other_text")
]

q16_cols = [
    col for col in coded_data.columns
    if col.startswith("q16_mjerenje_uspjeha__")
]

q17_cols = [
    col for col in coded_data.columns
    if col.startswith("q17_izvjestaj_o_odrzivosti__")
]

q18_cols = [
    col for col in coded_data.columns
    if col.startswith("q18_zakonodavni_utjecaj_sdg__")
]

q31_no_improvement_col = "q31_motivi_promjena__nista_poboljsano"
q34_no_external_funding_col = "q34_eksterna_sredstva_3_godine__bez_eksternih_sredstava"
q17_no_report_col = "q17_izvjestaj_o_odrzivosti__ne"

q31_positive_motivation_cols = [
    col for col in q31_cols
    if col != q31_no_improvement_col
]

q34_external_funding_source_cols = [
    col for col in q34_cols
    if col != q34_no_external_funding_col
]

q17_positive_awareness_cols = [
    col for col in q17_cols
    if col != q17_no_report_col
]


variable_labels = {
    "q15_pritisak_na_odrzivost__pravni": "Legal pressure",
    "q15_pritisak_na_odrzivost__potrosaci": "Consumers",
    "q15_pritisak_na_odrzivost__konkurencija": "Competition pressure",
    "q15_pritisak_na_odrzivost__okruzenje": "Business environment pressure",
    "q15_pritisak_na_odrzivost__zaposlenici": "Employees",
    "q15_pritisak_na_odrzivost__banka": "Bank",

    "q27_pomoc_za_odrzivost__nepovratni_poticaji": "Non-refundable incentives",
    "q27_pomoc_za_odrzivost__beskamatni_zajmovi": "Interest-free loans",
    "q27_pomoc_za_odrzivost__savjeti_strucnjaka": "Expert advice",
    "q27_pomoc_za_odrzivost__jasan_zakonodavni_okvir": "Clear legislative framework",
    "q27_pomoc_za_odrzivost__digitalna_rjesenja": "Digital solutions",
    "q27_pomoc_za_odrzivost__other_selected": "Other support",

    "q31_motivi_promjena__nista_poboljsano": "No improvements made",
    "q31_motivi_promjena__regulatorni_okviri": "Regulatory frameworks",
    "q31_motivi_promjena__zbog_konkurencije": "Competition as improvement motive",
    "q31_motivi_promjena__ideja_od_partnera": "Idea from partners",
    "q31_motivi_promjena__bespovratna_sredstva": "Grants / non-refundable funds",
    "q31_motivi_promjena__briga_za_okolis": "Environmental concern",
    "q31_motivi_promjena__energetska_ucinkovitost": "Energy efficiency",
    "q31_motivi_promjena__drustveni_ciljevi": "Social goals",
    "q31_motivi_promjena__poslovni_razvoj": "Business development",
    "q31_motivi_promjena__moderne_tehnologije_inovacije": "Modern technologies and innovation",
    "q31_motivi_promjena__covid_19": "COVID-19",
    "q31_motivi_promjena__other_selected": "Other improvement motive",

    "q34_eksterna_sredstva_3_godine__eu_konkurentnost": "EU competitiveness funds",
    "q34_eksterna_sredstva_3_godine__eu_energetska_ucinkovitost": "EU energy efficiency funds",
    "q34_eksterna_sredstva_3_godine__npoo": "NPOO",
    "q34_eksterna_sredstva_3_godine__lokalni_poticaji": "Local incentives",
    "q34_eksterna_sredstva_3_godine__banke_sufinancirana_kamata": "Bank loans with subsidized interest",
    "q34_eksterna_sredstva_3_godine__hbor": "HBOR",
    "q34_eksterna_sredstva_3_godine__ruralni_razvoj": "Rural development programmes",
    "q34_eksterna_sredstva_3_godine__hzz_zeleno_digitalno": "HZZ green/digital employment support",
    "q34_eksterna_sredstva_3_godine__hzz_zaposljavanje": "HZZ employment support",
    "q34_eksterna_sredstva_3_godine__zaposljavanje_osoba_s_invaliditetom": "Employment support for persons with disabilities",
    "q34_eksterna_sredstva_3_godine__efop": "EFOP",
    "q34_eksterna_sredstva_3_godine__drugi_eu_fondovi": "Other EU funds",
    "q34_eksterna_sredstva_3_godine__bez_eksternih_sredstava": "No external funding",
    "q34_eksterna_sredstva_3_godine__other_selected": "Other external funding",

    "q16_mjerenje_uspjeha__neto_profit": "Net profit only",
    "q16_mjerenje_uspjeha__profit_i_odrzivi_ciljevi": "Profit and sustainable goals",

    "q17_izvjestaj_o_odrzivosti__ne": "No awareness of sustainability report",
    "q17_izvjestaj_o_odrzivosti__da_zakonodavac_regulira": "Aware: legislator regulates",
    "q17_izvjestaj_o_odrzivosti__da_sankcije": "Aware: sanctions",
    "q17_izvjestaj_o_odrzivosti__da_doprinos_konkurentnosti": "Aware: contribution to competitiveness",
    "q17_izvjestaj_o_odrzivosti__da_nije_primjenjivo": "Aware: not applicable",

    "q18_zakonodavni_utjecaj_sdg__da": "Legislative SDG inclusion would affect results",
    "q18_zakonodavni_utjecaj_sdg__ne": "Legislative SDG inclusion would not affect results",

    "Any_pressure": "Any sustainability pressure",
    "Any_positive_improvement_motivation": "Any positive improvement motivation",
    "Any_external_funding": "Any external funding",
    "Sustainable_success_orientation": "Profit and sustainable goals",
    "Any_sustainability_report_awareness": "Any sustainability report awareness",
    "Regulatory_SDG_impact_yes": "Legislative SDG inclusion would affect results",
    "No_improvement_indicator": "No improvements made",
    "No_external_funding_indicator": "No external funding",
    "No_sustainability_report_awareness": "No awareness of sustainability report",

    "Pressure_count": "Number of sustainability pressure sources",
    "Support_needed_count": "Number of support types needed",
    "Improvement_motivators_count": "Number of positive improvement motives",
    "External_funding_count": "Number of external funding sources",
    "Sustainability_report_awareness_count": "Number of sustainability-report awareness types",
    "Economic_sustainability_importance": "Economic sustainability importance",
    "Environmental_sustainability_importance": "Environmental sustainability importance",
    "Social_sustainability_importance": "Social sustainability importance",
    "Sustainability_importance_mean": "Mean sustainability importance",
    "Sustainability_importance_high_count": "High sustainability-importance count",
}

indicator_type_map = {}

for col in q15_cols:
    indicator_type_map[col] = "pressure"

for col in q27_cols:
    indicator_type_map[col] = "support_need"

for col in q31_positive_motivation_cols:
    indicator_type_map[col] = "improvement_motive"

indicator_type_map[q31_no_improvement_col] = "absence_negative"

for col in q34_external_funding_source_cols:
    indicator_type_map[col] = "external_funding_resource"

indicator_type_map[q34_no_external_funding_col] = "absence_negative"

indicator_type_map["q16_mjerenje_uspjeha__profit_i_odrzivi_ciljevi"] = "success_orientation"
indicator_type_map["q16_mjerenje_uspjeha__neto_profit"] = "absence_negative"

for col in q17_positive_awareness_cols:
    indicator_type_map[col] = "report_awareness"

indicator_type_map[q17_no_report_col] = "absence_negative"

indicator_type_map["q18_zakonodavni_utjecaj_sdg__da"] = "regulatory_perception"
indicator_type_map["q18_zakonodavni_utjecaj_sdg__ne"] = "absence_negative"

indicator_type_map["Any_pressure"] = "synthetic_binary"
indicator_type_map["Any_positive_improvement_motivation"] = "synthetic_binary"
indicator_type_map["Any_external_funding"] = "synthetic_binary"
indicator_type_map["Any_sustainability_report_awareness"] = "synthetic_binary"
indicator_type_map["Sustainable_success_orientation"] = "success_orientation"
indicator_type_map["Regulatory_SDG_impact_yes"] = "regulatory_perception"
indicator_type_map["No_improvement_indicator"] = "absence_negative"
indicator_type_map["No_external_funding_indicator"] = "absence_negative"
indicator_type_map["No_sustainability_report_awareness"] = "absence_negative"

numeric_indicator_type_map = {
    "Pressure_count": "pressure_index",
    "Support_needed_count": "support_need_index",
    "Improvement_motivators_count": "improvement_motive_index",
    "External_funding_count": "external_funding_index",
    "Sustainability_report_awareness_count": "report_awareness_index",
    "Economic_sustainability_importance": "sustainability_value",
    "Environmental_sustainability_importance": "sustainability_value",
    "Social_sustainability_importance": "sustainability_value",
    "Sustainability_importance_mean": "sustainability_value_index",
    "Sustainability_importance_high_count": "sustainability_value_index",
}


binary_indicator_groups = {
    "Sustainability pressure": q15_cols,
    "Support needed": q27_cols,
    "Positive improvement motives": q31_positive_motivation_cols,
    "External funding resources": q34_external_funding_source_cols,
    "Success orientation": ["Sustainable_success_orientation"],
    "Sustainability report awareness": q17_positive_awareness_cols,
    "Regulatory SDG impact": ["Regulatory_SDG_impact_yes"],
    "Synthetic positive indicators": [
        "Any_pressure",
        "Any_positive_improvement_motivation",
        "Any_external_funding",
        "Any_sustainability_report_awareness",
    ],
    "Absence / negative indicators": [
        "No_improvement_indicator",
        "No_external_funding_indicator",
        "No_sustainability_report_awareness",
    ],
}

numeric_predictors = [
    "Pressure_count",
    "Support_needed_count",
    "Improvement_motivators_count",
    "External_funding_count",
    "Sustainability_report_awareness_count",
    "Economic_sustainability_importance",
    "Environmental_sustainability_importance",
    "Social_sustainability_importance",
    "Sustainability_importance_mean",
    "Sustainability_importance_high_count",
]

constant_binary_indicators = []

for group_name, indicators in binary_indicator_groups.items():
    for indicator in indicators:
        if indicator not in coded_data.columns:
            raise ValueError(f"Missing indicator: {indicator}")

        unique_values = sorted(coded_data[indicator].dropna().unique())

        if len(unique_values) < 2:
            constant_binary_indicators.append(indicator)

constant_numeric_predictors = []

for predictor in numeric_predictors:
    if predictor not in coded_data.columns:
        raise ValueError(f"Missing numeric predictor: {predictor}")

    if coded_data[predictor].nunique(dropna=True) < 2:
        constant_numeric_predictors.append(predictor)

print("\n1) Constant binary indicators that will be skipped:")
if constant_binary_indicators:
    print(constant_binary_indicators)
else:
    print("No constant binary indicators in the analysed set.")

print("\n2) Constant numeric predictors that will be skipped:")
if constant_numeric_predictors:
    print(constant_numeric_predictors)
else:
    print("No constant numeric predictors in the analysed set.")


def cliffs_delta(x, y):
    """Calculate Cliff's delta for two independent groups."""
    x = np.asarray(x)
    y = np.asarray(y)

    n_x = len(x)
    n_y = len(y)

    if n_x == 0 or n_y == 0:
        return np.nan

    greater = 0
    less = 0

    for xi in x:
        greater += np.sum(xi > y)
        less += np.sum(xi < y)

    return (greater - less) / (n_x * n_y)


def cliffs_delta_magnitude(delta):
    """Classify the magnitude of Cliff's delta."""
    if pd.isna(delta):
        return np.nan

    abs_delta = abs(delta)

    if abs_delta < 0.147:
        return "negligible"
    elif abs_delta < 0.33:
        return "small"
    elif abs_delta < 0.474:
        return "medium"
    else:
        return "large"


def benjamini_hochberg(p_values):
    """Apply Benjamini-Hochberg FDR correction."""
    p_values = np.asarray(p_values, dtype=float)
    n = len(p_values)

    order = np.argsort(p_values)
    ranked_p = p_values[order]

    adjusted = np.empty(n, dtype=float)

    cumulative_min = 1.0

    for i in range(n - 1, -1, -1):
        rank = i + 1
        value = ranked_p[i] * n / rank
        cumulative_min = min(cumulative_min, value)
        adjusted[order[i]] = cumulative_min

    return np.minimum(adjusted, 1.0)


def compare_outcome_by_binary_indicator(
    df,
    indicator,
    outcome,
    construct=None
):
    """Compare an SDG outcome between indicator=1 and indicator=0 groups."""

    selected = df.loc[df[indicator] == 1, outcome].dropna()
    not_selected = df.loc[df[indicator] == 0, outcome].dropna()

    n_selected = len(selected)
    n_not_selected = len(not_selected)

    if n_selected == 0 or n_not_selected == 0:
        mannwhitney_u = np.nan
        mannwhitney_p = np.nan
        spearman_rho = np.nan
        spearman_p = np.nan
        delta = np.nan
        delta_magnitude = np.nan
    else:
        mannwhitney_u, mannwhitney_p = stats.mannwhitneyu(
            selected,
            not_selected,
            alternative="two-sided"
        )

        spearman_rho, spearman_p = stats.spearmanr(
            df[indicator],
            df[outcome],
            nan_policy="omit"
        )

        delta = cliffs_delta(selected, not_selected)
        delta_magnitude = cliffs_delta_magnitude(delta)

    return {
        "construct": construct,
        "indicator_type": indicator_type_map.get(indicator, "unknown"),
        "indicator": indicator,
        "label": variable_labels.get(indicator, indicator),
        "outcome": outcome,
        "n_selected": n_selected,
        "n_not_selected": n_not_selected,
        "selected_percent": n_selected / len(df) * 100,
        "mean_selected": selected.mean() if n_selected > 0 else np.nan,
        "mean_not_selected": not_selected.mean() if n_not_selected > 0 else np.nan,
        "median_selected": selected.median() if n_selected > 0 else np.nan,
        "median_not_selected": not_selected.median() if n_not_selected > 0 else np.nan,
        "mean_difference_selected_minus_not": (
            selected.mean() - not_selected.mean()
            if n_selected > 0 and n_not_selected > 0 else np.nan
        ),
        "median_difference_selected_minus_not": (
            selected.median() - not_selected.median()
            if n_selected > 0 and n_not_selected > 0 else np.nan
        ),
        "mannwhitney_u": mannwhitney_u,
        "mannwhitney_p": mannwhitney_p,
        "spearman_rho": spearman_rho,
        "spearman_p": spearman_p,
        "cliffs_delta": delta,
        "cliffs_delta_magnitude": delta_magnitude
    }


def correlate_numeric_predictor_with_outcome(
    df,
    predictor,
    outcome
):
    """Correlate a numeric or ordinal predictor with an SDG outcome."""

    x = df[predictor]
    y = df[outcome]

    valid = x.notna() & y.notna()

    if valid.sum() < 3 or x[valid].nunique() < 2 or y[valid].nunique() < 2:
        spearman_rho = np.nan
        spearman_p = np.nan
        kendall_tau = np.nan
        kendall_p = np.nan
    else:
        spearman_rho, spearman_p = stats.spearmanr(
            x[valid],
            y[valid]
        )

        kendall_tau, kendall_p = stats.kendalltau(
            x[valid],
            y[valid]
        )

    return {
        "predictor": predictor,
        "label": variable_labels.get(predictor, predictor),
        "predictor_type": numeric_indicator_type_map.get(predictor, "unknown"),
        "outcome": outcome,
        "n": int(valid.sum()),
        "predictor_mean": x[valid].mean(),
        "predictor_median": x[valid].median(),
        "outcome_mean": y[valid].mean(),
        "outcome_median": y[valid].median(),
        "spearman_rho": spearman_rho,
        "spearman_p": spearman_p,
        "kendall_tau": kendall_tau,
        "kendall_p": kendall_p
    }


binary_rows = []

for construct, indicators in binary_indicator_groups.items():
    for indicator in indicators:

        if indicator in constant_binary_indicators:
            continue

        for outcome in sdg_outcome_cols:
            binary_rows.append(
                compare_outcome_by_binary_indicator(
                    coded_data,
                    indicator=indicator,
                    outcome=outcome,
                    construct=construct
                )
            )

rq2_3b_binary_results = pd.DataFrame(binary_rows)

rq2_3b_binary_results["mannwhitney_p_fdr_by_outcome"] = np.nan
rq2_3b_binary_results["spearman_p_fdr_by_outcome"] = np.nan

for outcome in sdg_outcome_cols:
    outcome_mask = rq2_3b_binary_results["outcome"] == outcome

    mw_valid = outcome_mask & rq2_3b_binary_results["mannwhitney_p"].notna()
    sp_valid = outcome_mask & rq2_3b_binary_results["spearman_p"].notna()

    if mw_valid.sum() > 0:
        rq2_3b_binary_results.loc[mw_valid, "mannwhitney_p_fdr_by_outcome"] = (
            benjamini_hochberg(
                rq2_3b_binary_results.loc[mw_valid, "mannwhitney_p"].values
            )
        )

    if sp_valid.sum() > 0:
        rq2_3b_binary_results.loc[sp_valid, "spearman_p_fdr_by_outcome"] = (
            benjamini_hochberg(
                rq2_3b_binary_results.loc[sp_valid, "spearman_p"].values
            )
        )

rq2_3b_binary_results["mw_significant_raw_0_05"] = (
    rq2_3b_binary_results["mannwhitney_p"] < 0.05
)

rq2_3b_binary_results["mw_significant_fdr_0_05"] = (
    rq2_3b_binary_results["mannwhitney_p_fdr_by_outcome"] < 0.05
)

rq2_3b_binary_results["spearman_significant_raw_0_05"] = (
    rq2_3b_binary_results["spearman_p"] < 0.05
)

rq2_3b_binary_results["spearman_significant_fdr_0_05"] = (
    rq2_3b_binary_results["spearman_p_fdr_by_outcome"] < 0.05
)


numeric_rows = []

for predictor in numeric_predictors:

    if predictor in constant_numeric_predictors:
        continue

    for outcome in sdg_outcome_cols:
        numeric_rows.append(
            correlate_numeric_predictor_with_outcome(
                coded_data,
                predictor=predictor,
                outcome=outcome
            )
        )

rq2_3b_numeric_results = pd.DataFrame(numeric_rows)

rq2_3b_numeric_results["spearman_p_fdr_by_outcome"] = np.nan
rq2_3b_numeric_results["kendall_p_fdr_by_outcome"] = np.nan

for outcome in sdg_outcome_cols:
    outcome_mask = rq2_3b_numeric_results["outcome"] == outcome

    sp_valid = outcome_mask & rq2_3b_numeric_results["spearman_p"].notna()
    kt_valid = outcome_mask & rq2_3b_numeric_results["kendall_p"].notna()

    if sp_valid.sum() > 0:
        rq2_3b_numeric_results.loc[sp_valid, "spearman_p_fdr_by_outcome"] = (
            benjamini_hochberg(
                rq2_3b_numeric_results.loc[sp_valid, "spearman_p"].values
            )
        )

    if kt_valid.sum() > 0:
        rq2_3b_numeric_results.loc[kt_valid, "kendall_p_fdr_by_outcome"] = (
            benjamini_hochberg(
                rq2_3b_numeric_results.loc[kt_valid, "kendall_p"].values
            )
        )

rq2_3b_numeric_results["spearman_significant_raw_0_05"] = (
    rq2_3b_numeric_results["spearman_p"] < 0.05
)

rq2_3b_numeric_results["spearman_significant_fdr_0_05"] = (
    rq2_3b_numeric_results["spearman_p_fdr_by_outcome"] < 0.05
)

rq2_3b_numeric_results["kendall_significant_raw_0_05"] = (
    rq2_3b_numeric_results["kendall_p"] < 0.05
)

rq2_3b_numeric_results["kendall_significant_fdr_0_05"] = (
    rq2_3b_numeric_results["kendall_p_fdr_by_outcome"] < 0.05
)


binary_sdg_count = (
    rq2_3b_binary_results[
        rq2_3b_binary_results["outcome"] == "SDG_count"
    ]
    .copy()
    .sort_values(
        ["mean_difference_selected_minus_not", "spearman_rho"],
        ascending=False
    )
    .reset_index(drop=True)
)

binary_sdg_count.insert(
    0,
    "rank_by_mean_difference",
    np.arange(1, len(binary_sdg_count) + 1)
)

binary_sdg_count_reliable = binary_sdg_count[
    (binary_sdg_count["n_selected"] >= 5)
    & (binary_sdg_count["n_not_selected"] >= 5)
].copy()

numeric_sdg_count = (
    rq2_3b_numeric_results[
        rq2_3b_numeric_results["outcome"] == "SDG_count"
    ]
    .copy()
    .sort_values(
        ["spearman_rho", "kendall_tau"],
        ascending=False
    )
    .reset_index(drop=True)
)

numeric_sdg_count.insert(
    0,
    "rank_by_spearman_sdg_count",
    np.arange(1, len(numeric_sdg_count) + 1)
)


def round_numeric_columns(df, decimals=4):
    out = df.copy()
    for col in out.select_dtypes(include=[np.number]).columns:
        out[col] = out[col].round(decimals)
    return out

binary_sdg_count_display = round_numeric_columns(binary_sdg_count)
binary_sdg_count_reliable_display = round_numeric_columns(binary_sdg_count_reliable)
numeric_sdg_count_display = round_numeric_columns(numeric_sdg_count)

binary_results_display = round_numeric_columns(rq2_3b_binary_results)
numeric_results_display = round_numeric_columns(rq2_3b_numeric_results)


binary_sdg_count_raw_sig = binary_sdg_count[
    binary_sdg_count["mw_significant_raw_0_05"]
].copy()

binary_sdg_count_fdr_sig = binary_sdg_count[
    binary_sdg_count["mw_significant_fdr_0_05"]
].copy()

numeric_sdg_count_raw_sig = numeric_sdg_count[
    numeric_sdg_count["spearman_significant_raw_0_05"]
].copy()

numeric_sdg_count_fdr_sig = numeric_sdg_count[
    numeric_sdg_count["spearman_significant_fdr_0_05"]
].copy()


print("\n" + "=" * 100)
print("RQ2 - STEP 3b: ASSOCIATIONS BETWEEN FACTORS AND SDG INVOLVEMENT")
print("=" * 100)

print("\n3) Top 15 binary indicators by positive difference in SDG_count:")
print(
    binary_sdg_count_display[
        [
            "rank_by_mean_difference",
            "construct",
            "indicator_type",
            "label",
            "n_selected",
            "n_not_selected",
            "selected_percent",
            "mean_selected",
            "mean_not_selected",
            "mean_difference_selected_minus_not",
            "median_selected",
            "median_not_selected",
            "mannwhitney_p",
            "mannwhitney_p_fdr_by_outcome",
            "cliffs_delta",
            "cliffs_delta_magnitude"
        ]
    ].head(15).to_string(index=False)
)

print("\n4) Top 15 binary indicators with sufficiently large groups, n_selected >= 5 and n_not_selected >= 5:")
print(
    binary_sdg_count_reliable_display[
        [
            "rank_by_mean_difference",
            "construct",
            "indicator_type",
            "label",
            "n_selected",
            "n_not_selected",
            "selected_percent",
            "mean_selected",
            "mean_not_selected",
            "mean_difference_selected_minus_not",
            "median_selected",
            "median_not_selected",
            "mannwhitney_p",
            "mannwhitney_p_fdr_by_outcome",
            "cliffs_delta",
            "cliffs_delta_magnitude"
        ]
    ].head(15).to_string(index=False)
)

print("\n5) Binary indicators associated with SDG_count: raw p < 0.05")
if len(binary_sdg_count_raw_sig) > 0:
    print(
        round_numeric_columns(binary_sdg_count_raw_sig)[
            [
                "construct",
                "indicator_type",
                "label",
                "n_selected",
                "n_not_selected",
                "mean_selected",
                "mean_not_selected",
                "mean_difference_selected_minus_not",
                "mannwhitney_p",
                "mannwhitney_p_fdr_by_outcome",
                "cliffs_delta",
                "cliffs_delta_magnitude"
            ]
        ].to_string(index=False)
    )
else:
    print("No binary indicators had an unadjusted p-value < 0.05.")

print("\n6) Binary indicators associated with SDG_count: FDR p < 0.05")
if len(binary_sdg_count_fdr_sig) > 0:
    print(
        round_numeric_columns(binary_sdg_count_fdr_sig)[
            [
                "construct",
                "indicator_type",
                "label",
                "n_selected",
                "n_not_selected",
                "mean_selected",
                "mean_not_selected",
                "mean_difference_selected_minus_not",
                "mannwhitney_p",
                "mannwhitney_p_fdr_by_outcome",
                "cliffs_delta",
                "cliffs_delta_magnitude"
            ]
        ].to_string(index=False)
    )
else:
    print("No binary indicators remained significant after FDR correction.")

print("\n7) Bottom 10 binary indicators associated with lower SDG_count:")
bottom10_binary = binary_sdg_count_display.tail(10).sort_values(
    "mean_difference_selected_minus_not",
    ascending=True
)
print(
    bottom10_binary[
        [
            "construct",
            "indicator_type",
            "label",
            "n_selected",
            "n_not_selected",
            "mean_selected",
            "mean_not_selected",
            "mean_difference_selected_minus_not",
            "mannwhitney_p",
            "mannwhitney_p_fdr_by_outcome",
            "cliffs_delta",
            "cliffs_delta_magnitude"
        ]
    ].to_string(index=False)
)

print("\n8) Numeric/scalar predictors ranked by Spearman rho with SDG_count:")
print(
    numeric_sdg_count_display[
        [
            "rank_by_spearman_sdg_count",
            "predictor_type",
            "label",
            "n",
            "predictor_mean",
            "predictor_median",
            "outcome_mean",
            "outcome_median",
            "spearman_rho",
            "spearman_p",
            "spearman_p_fdr_by_outcome",
            "kendall_tau",
            "kendall_p",
            "kendall_p_fdr_by_outcome"
        ]
    ].to_string(index=False)
)

print("\n9) Numeric/scalar predictors associated with SDG_count: raw Spearman p < 0.05")
if len(numeric_sdg_count_raw_sig) > 0:
    print(
        round_numeric_columns(numeric_sdg_count_raw_sig)[
            [
                "predictor_type",
                "label",
                "spearman_rho",
                "spearman_p",
                "spearman_p_fdr_by_outcome",
                "kendall_tau",
                "kendall_p",
                "kendall_p_fdr_by_outcome"
            ]
        ].to_string(index=False)
    )
else:
    print("No numeric/scalar predictors had an unadjusted Spearman p-value < 0.05.")

print("\n10) Numeric/scalar predictors associated with SDG_count: FDR-adjusted Spearman p < 0.05")
if len(numeric_sdg_count_fdr_sig) > 0:
    print(
        round_numeric_columns(numeric_sdg_count_fdr_sig)[
            [
                "predictor_type",
                "label",
                "spearman_rho",
                "spearman_p",
                "spearman_p_fdr_by_outcome",
                "kendall_tau",
                "kendall_p",
                "kendall_p_fdr_by_outcome"
            ]
        ].to_string(index=False)
    )
else:
    print("No numeric/scalar predictors remained significant after FDR correction.")


print("\n11) Short interpretive summary for SDG_count:")

top_binary = binary_sdg_count_reliable_display.head(5)
top_numeric = numeric_sdg_count_display.head(5)

print("\nLargest positive differences in SDG_count among binary indicators:")
for _, row in top_binary.iterrows():
    print(
        f"- {row['label']} ({row['construct']}): "
        f"mean selected = {row['mean_selected']:.2f}, "
        f"mean not selected = {row['mean_not_selected']:.2f}, "
        f"difference = {row['mean_difference_selected_minus_not']:.2f}, "
        f"Cliff's delta = {row['cliffs_delta']:.2f} ({row['cliffs_delta_magnitude']}), "
        f"FDR p = {row['mannwhitney_p_fdr_by_outcome']:.4f}"
    )

print("\nStrongest positive Spearman correlations with SDG_count among numeric/scalar predictors:")
for _, row in top_numeric.iterrows():
    print(
        f"- {row['label']}: "
        f"rho = {row['spearman_rho']:.3f}, "
        f"FDR p = {row['spearman_p_fdr_by_outcome']:.4f}"
    )

print(
    "\nNote: The results show associations, not causality. "
    "For RQ2, they are interpreted as motivational, institutional and value-related "
    "factors that accompany higher or lower SDG involvement."
)


output_path = "/content/RQ2_step3b_SDG_involvement_associations.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    rq2_3b_binary_results.to_excel(
        writer,
        sheet_name="binary_all_outcomes",
        index=False
    )

    rq2_3b_numeric_results.to_excel(
        writer,
        sheet_name="numeric_all_outcomes",
        index=False
    )

    binary_sdg_count.to_excel(
        writer,
        sheet_name="binary_SDG_count_ranked",
        index=False
    )

    binary_sdg_count_reliable.to_excel(
        writer,
        sheet_name="binary_SDG_count_reliable",
        index=False
    )

    numeric_sdg_count.to_excel(
        writer,
        sheet_name="numeric_SDG_count_ranked",
        index=False
    )

    binary_sdg_count_raw_sig.to_excel(
        writer,
        sheet_name="binary_raw_sig",
        index=False
    )

    binary_sdg_count_fdr_sig.to_excel(
        writer,
        sheet_name="binary_fdr_sig",
        index=False
    )

    numeric_sdg_count_raw_sig.to_excel(
        writer,
        sheet_name="numeric_raw_sig",
        index=False
    )

    numeric_sdg_count_fdr_sig.to_excel(
        writer,
        sheet_name="numeric_fdr_sig",
        index=False
    )

print("\nRQ2 Step 3b results saved to:")
print(output_path)

print("\nRQ2 Step 3b completed.")
