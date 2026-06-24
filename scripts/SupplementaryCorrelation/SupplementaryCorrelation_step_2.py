"""
Supplementary Correlation Analysis Step 2: Construction of q31 and q34 derived indices.

This script constructs derived indicators for the supplementary analysis of the
relationship between external funding use and motives for implemented improvements.

Question q31 refers to motives for implemented changes or improvements. Question
q34 refers to external funding sources used during the previous three years.
The script creates count-based and binary indicators for positive improvement
motives, external funding use, EU or public funding, and bank or credit-related
support. It also validates the derived indicators, produces frequency tables and
descriptive summaries, and saves the results to an Excel file.

Input:
    Survey_Results.xlsx, sheet: encoded_responses_corrected

Output:
    Additional_Analysis_3_q31_q34_indices.xlsx
"""

import pandas as pd


FILE_PATH = "/content/Survey_Results.xlsx"
SHEET_NAME = "encoded_responses_corrected"
OUTPUT_PATH = "/content/Additional_Analysis_3_q31_q34_indices.xlsx"

q31_no_improvement_col = "q31_motivi_promjena__nista_poboljsano"

q31_positive_motive_cols = [
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

q31_all_binary_cols = [q31_no_improvement_col] + q31_positive_motive_cols

q31_labels = {
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

q34_no_external_col = "q34_eksterna_sredstva_3_godine__bez_eksternih_sredstava"

q34_positive_funding_cols = [
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

q34_all_binary_cols = [q34_no_external_col] + q34_positive_funding_cols

q34_labels = {
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


def frequency_table(data, column):
    """Create a frequency table for a categorical or count variable."""
    table = (
        data[column]
        .value_counts(dropna=False)
        .sort_index()
        .rename_axis(column)
        .reset_index(name="n")
    )
    table["percent"] = (table["n"] / len(data) * 100).round(2)
    return table


def indicator_frequency_table(data, columns, labels, negative_column, positive_type):
    """Create a frequency table for a list of binary indicators."""
    rows = []
    for column in columns:
        selected_n = int(data[column].sum())
        rows.append({
            "variable": column,
            "label": labels[column],
            "selected_n": selected_n,
            "selected_percent": round(selected_n / len(data) * 100, 2),
            "type": "negative_none" if column == negative_column else positive_type,
        })
    return pd.DataFrame(rows).sort_values(
        "selected_percent",
        ascending=False,
    ).reset_index(drop=True)


def main():
    data = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)

    print("SUPPLEMENTARY CORRELATION ANALYSIS - STEP 2")
    print("Derived indices: external funding and improvement motives")
    print("=" * 92)
    print("Dataset dimensions:", data.shape)
    print("Number of respondents / firms:", len(data))

    missing_q31 = [column for column in q31_all_binary_cols if column not in data.columns]
    missing_q34 = [column for column in q34_all_binary_cols if column not in data.columns]

    if missing_q31:
        raise ValueError(f"Missing q31 columns: {missing_q31}")

    if missing_q34:
        raise ValueError(f"Missing q34 columns: {missing_q34}")

    print("\n1) All required q31 and q34 columns are present.")

    data["Improvement_motivators_count"] = data[q31_positive_motive_cols].sum(axis=1)
    data["Any_positive_improvement_motivation"] = (
        data["Improvement_motivators_count"] > 0
    ).astype(int)
    data["No_improvement_indicator"] = data[q31_no_improvement_col]
    data["q31_total_selected_count"] = data[q31_all_binary_cols].sum(axis=1)

    data["External_funding_count"] = data[q34_positive_funding_cols].sum(axis=1)
    data["Any_external_funding"] = (data["External_funding_count"] > 0).astype(int)
    data["No_external_funding_indicator"] = data[q34_no_external_col]
    data["q34_total_selected_count"] = data[q34_all_binary_cols].sum(axis=1)

    data["Any_EU_or_public_funding"] = (
        data[[
            "q34_eksterna_sredstva_3_godine__eu_konkurentnost",
            "q34_eksterna_sredstva_3_godine__eu_energetska_ucinkovitost",
            "q34_eksterna_sredstva_3_godine__npoo",
            "q34_eksterna_sredstva_3_godine__lokalni_poticaji",
            "q34_eksterna_sredstva_3_godine__hzz_zeleno_digitalno",
            "q34_eksterna_sredstva_3_godine__hzz_zaposljavanje",
            "q34_eksterna_sredstva_3_godine__ruralni_razvoj",
            "q34_eksterna_sredstva_3_godine__drugi_eu_fondovi",
        ]].sum(axis=1) > 0
    ).astype(int)

    data["Any_bank_or_credit_support"] = (
        data[[
            "q34_eksterna_sredstva_3_godine__banke_sufinancirana_kamata",
            "q34_eksterna_sredstva_3_godine__hbor",
        ]].sum(axis=1) > 0
    ).astype(int)

    improvement_count_freq = frequency_table(data, "Improvement_motivators_count")
    external_funding_count_freq = frequency_table(data, "External_funding_count")

    print("\n2) Distribution of Improvement_motivators_count:")
    print(improvement_count_freq.to_string(index=False))

    print("\n3) Distribution of External_funding_count:")
    print(external_funding_count_freq.to_string(index=False))

    q31_indicator_freq = indicator_frequency_table(
        data,
        q31_all_binary_cols,
        q31_labels,
        q31_no_improvement_col,
        "positive_motive",
    )

    q34_indicator_freq = indicator_frequency_table(
        data,
        q34_all_binary_cols,
        q34_labels,
        q34_no_external_col,
        "positive_funding",
    )

    print("\n4) Frequencies of q31 improvement motives:")
    print(q31_indicator_freq.to_string(index=False))

    print("\n5) Frequencies of q34 external funding sources:")
    print(q34_indicator_freq.to_string(index=False))

    derived_cols = [
        "Improvement_motivators_count",
        "Any_positive_improvement_motivation",
        "No_improvement_indicator",
        "q31_total_selected_count",
        "External_funding_count",
        "Any_external_funding",
        "No_external_funding_indicator",
        "q34_total_selected_count",
        "Any_EU_or_public_funding",
        "Any_bank_or_credit_support",
    ]

    derived_summary = data[derived_cols].describe().T
    derived_summary["median"] = data[derived_cols].median()
    derived_summary = derived_summary[
        ["count", "mean", "std", "min", "25%", "median", "50%", "75%", "max"]
    ]
    derived_summary_display = derived_summary.round(3)

    print("\n6) Descriptive statistics for derived q31/q34 indices:")
    print(derived_summary_display.to_string())

    q31_no_contradiction = (
        ((data["No_improvement_indicator"] == 1) & (data["Improvement_motivators_count"] > 0)).sum() == 0
    )
    q34_no_contradiction = (
        ((data["No_external_funding_indicator"] == 1) & (data["External_funding_count"] > 0)).sum() == 0
    )
    q31_any_consistent = (
        data["Any_positive_improvement_motivation"]
        == (data["Improvement_motivators_count"] > 0).astype(int)
    ).all()
    q34_any_consistent = (
        data["Any_external_funding"]
        == (data["External_funding_count"] > 0).astype(int)
    ).all()
    q31_none_complement = (
        data["No_improvement_indicator"]
        == 1 - data["Any_positive_improvement_motivation"]
    ).all()
    q34_none_complement = (
        data["No_external_funding_indicator"]
        == 1 - data["Any_external_funding"]
    ).all()
    q31_count_range_valid = data["Improvement_motivators_count"].between(
        0,
        len(q31_positive_motive_cols),
    ).all()
    q34_count_range_valid = data["External_funding_count"].between(
        0,
        len(q34_positive_funding_cols),
    ).all()
    derived_no_missing = data[derived_cols].isna().sum().sum() == 0

    step2_valid = (
        q31_no_contradiction
        and q34_no_contradiction
        and q31_any_consistent
        and q34_any_consistent
        and q31_none_complement
        and q34_none_complement
        and q31_count_range_valid
        and q34_count_range_valid
        and derived_no_missing
    )

    print("\n7) Validation of derived indices:")
    print("q31 has no contradiction between no-improvement and positive motives:", bool(q31_no_contradiction))
    print("q34 has no contradiction between no-external-funding and positive funding:", bool(q34_no_contradiction))
    print("Any_positive_improvement_motivation is consistent:", bool(q31_any_consistent))
    print("Any_external_funding is consistent:", bool(q34_any_consistent))
    print("No_improvement_indicator is the complement of positive improvement motivation:", bool(q31_none_complement))
    print("No_external_funding_indicator is the complement of external funding use:", bool(q34_none_complement))
    print("Improvement_motivators_count is within the expected range:", bool(q31_count_range_valid))
    print("External_funding_count is within the expected range:", bool(q34_count_range_valid))
    print("Derived indices have no missing values:", bool(derived_no_missing))
    print("\nSupplementary Correlation Analysis - Step 2 fully consistent:", bool(step2_valid))

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        data.to_excel(writer, sheet_name="encoded_with_q31_q34", index=False)
        improvement_count_freq.to_excel(writer, sheet_name="improvement_count_freq", index=False)
        external_funding_count_freq.to_excel(writer, sheet_name="external_funding_freq", index=False)
        q31_indicator_freq.to_excel(writer, sheet_name="q31_indicator_freq", index=False)
        q34_indicator_freq.to_excel(writer, sheet_name="q34_indicator_freq", index=False)
        derived_summary_display.to_excel(writer, sheet_name="derived_summary")

    print("\nFile with derived q31/q34 indices saved to:")
    print(OUTPUT_PATH)
    print("\nSupplementary Correlation Analysis - Step 2 completed.")


if __name__ == "__main__":
    main()
