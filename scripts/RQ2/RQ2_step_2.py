"""
RQ2 Step 2: Construction of motivational, institutional and value-based indices.

This script constructs derived variables used to answer RQ2:
"What are the most important motivators for SDG involvement in SME operations?"

The script uses the coded survey dataset, reconstructs the SDG outcome variables
from RQ1, creates RQ2 indices related to sustainability pressure, institutional
support needs, improvement motives, external funding, sustainability-report
awareness, regulatory SDG impact, and sustainability values. It validates ranges,
missing values, binary indicators, and logical consistency, then saves the
extended dataset and validation summaries to an Excel file.

Input:
    Survey_Results.xlsx, sheet "encoded_responses_corrected".

Output:
    RQ2_step2_indices.xlsx.
"""

import pandas as pd


FILE_PATH = "/content/Survey_Results.xlsx"
SHEET_NAME = "encoded_responses_corrected"
OUTPUT_PATH = "/content/RQ2_step2_indices.xlsx"
NONE_SDG_COL = "q37_sdg_vezani_uz_poslovanje__none_directly_related"

Q20_COL = "q20_vaznost_ekonomske_odrzivosti"
Q21_COL = "q21_vaznost_okolisne_odrzivosti"
Q22_COL = "q22_vaznost_drustvene_odrzivosti"

Q31_NO_IMPROVEMENT_COL = "q31_motivi_promjena__nista_poboljsano"
Q34_NO_EXTERNAL_FUNDING_COL = "q34_eksterna_sredstva_3_godine__bez_eksternih_sredstava"
Q17_NO_REPORT_COL = "q17_izvjestaj_o_odrzivosti__ne"
Q16_SUSTAINABLE_SUCCESS_COL = "q16_mjerenje_uspjeha__profit_i_odrzivi_ciljevi"
Q18_REGULATORY_YES_COL = "q18_zakonodavni_utjecaj_sdg__da"


def get_columns(data, prefix, exclude_other_text=False, exclude_column=None):
    columns = [column for column in data.columns if column.startswith(prefix)]
    if exclude_other_text:
        columns = [column for column in columns if not column.endswith("__other_text")]
    if exclude_column is not None:
        columns = [column for column in columns if column != exclude_column]
    return columns


def validate_columns_present(data, columns, label):
    missing_columns = [column for column in columns if column not in data.columns]
    if missing_columns:
        raise ValueError(f"Missing {label} columns: {missing_columns}")


def frequency_table(data, column):
    table = (
        data[column]
        .value_counts()
        .sort_index()
        .rename_axis(column)
        .reset_index(name="number_of_SMEs")
    )
    table["percent_of_SMEs"] = (table["number_of_SMEs"] / len(data) * 100).round(2)
    return table


def add_logic_check(rows, name, condition):
    rows.append({"logic_check": name, "valid": bool(condition)})


def main():
    coded_data = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)

    print("RQ2 - STEP 2")
    print("Construction of motivational, institutional and value-based indices")
    print("=" * 100)
    print("Input table dimensions:", coded_data.shape)
    print("Number of respondents / firms:", len(coded_data))

    sdg_relevance_cols = get_columns(coded_data, "q36_relevantnost_sdg__")
    sdg_direct_cols = get_columns(
        coded_data,
        "q37_sdg_vezani_uz_poslovanje__",
        exclude_column=NONE_SDG_COL,
    )

    if len(sdg_relevance_cols) != 17:
        raise ValueError(
            f"Expected 17 q36 SDG relevance columns, but found {len(sdg_relevance_cols)}."
        )
    if len(sdg_direct_cols) != 17:
        raise ValueError(
            f"Expected 17 q37 direct SDG columns, but found {len(sdg_direct_cols)}."
        )
    if NONE_SDG_COL not in coded_data.columns:
        raise ValueError(f"Missing expected column: {NONE_SDG_COL}")

    coded_data["SDG_count"] = coded_data[sdg_direct_cols].sum(axis=1)
    coded_data["SDG_mean_relevance"] = coded_data[sdg_relevance_cols].mean(axis=1)
    coded_data["SDG_high_relevance_count"] = (coded_data[sdg_relevance_cols] >= 4).sum(axis=1)

    q15_cols = get_columns(coded_data, "q15_pritisak_na_odrzivost__")
    q27_cols = get_columns(coded_data, "q27_pomoc_za_odrzivost__", exclude_other_text=True)
    q31_cols = get_columns(coded_data, "q31_motivi_promjena__", exclude_other_text=True)
    q34_cols = get_columns(coded_data, "q34_eksterna_sredstva_3_godine__", exclude_other_text=True)
    q16_cols = get_columns(coded_data, "q16_mjerenje_uspjeha__")
    q17_cols = get_columns(coded_data, "q17_izvjestaj_o_odrzivosti__")
    q18_cols = get_columns(coded_data, "q18_zakonodavni_utjecaj_sdg__")

    q20_q22_cols = [Q20_COL, Q21_COL, Q22_COL]
    validate_columns_present(coded_data, q20_q22_cols, "q20-q22 Likert")

    rq2_group_sizes = {
        "q15_pritisak_na_odrzivost": len(q15_cols),
        "q27_pomoc_za_odrzivost": len(q27_cols),
        "q31_motivi_promjena": len(q31_cols),
        "q34_eksterna_sredstva_3_godine": len(q34_cols),
        "q16_mjerenje_uspjeha": len(q16_cols),
        "q17_izvjestaj_o_odrzivosti": len(q17_cols),
        "q18_zakonodavni_utjecaj_sdg": len(q18_cols),
        "q20_q22_likert": len(q20_q22_cols),
    }

    print("\n1) Number of columns by RQ2 group:")
    for group_name, number_of_columns in rq2_group_sizes.items():
        print(f"{group_name}: {number_of_columns}")

    empty_groups = [group_name for group_name, number_of_columns in rq2_group_sizes.items() if number_of_columns == 0]
    if empty_groups:
        raise ValueError(f"The following groups have no detected columns: {empty_groups}")

    required_single_columns = [
        Q31_NO_IMPROVEMENT_COL,
        Q34_NO_EXTERNAL_FUNDING_COL,
        Q17_NO_REPORT_COL,
        Q16_SUSTAINABLE_SUCCESS_COL,
        Q18_REGULATORY_YES_COL,
    ]
    validate_columns_present(coded_data, required_single_columns, "required RQ2")

    q31_positive_motivation_cols = [column for column in q31_cols if column != Q31_NO_IMPROVEMENT_COL]
    q34_external_funding_source_cols = [column for column in q34_cols if column != Q34_NO_EXTERNAL_FUNDING_COL]
    q17_positive_awareness_cols = [column for column in q17_cols if column != Q17_NO_REPORT_COL]

    coded_data["Pressure_count"] = coded_data[q15_cols].sum(axis=1)
    coded_data["Any_pressure"] = (coded_data["Pressure_count"] > 0).astype(int)
    coded_data["Support_needed_count"] = coded_data[q27_cols].sum(axis=1)
    coded_data["Any_support_needed"] = (coded_data["Support_needed_count"] > 0).astype(int)
    coded_data["Improvement_motivators_count"] = coded_data[q31_positive_motivation_cols].sum(axis=1)
    coded_data["No_improvement_indicator"] = coded_data[Q31_NO_IMPROVEMENT_COL]
    coded_data["Any_positive_improvement_motivation"] = (
        coded_data["Improvement_motivators_count"] > 0
    ).astype(int)
    coded_data["External_funding_count"] = coded_data[q34_external_funding_source_cols].sum(axis=1)
    coded_data["No_external_funding_indicator"] = coded_data[Q34_NO_EXTERNAL_FUNDING_COL]
    coded_data["Any_external_funding"] = (coded_data["External_funding_count"] > 0).astype(int)
    coded_data["Sustainable_success_orientation"] = coded_data[Q16_SUSTAINABLE_SUCCESS_COL]
    coded_data["Sustainability_report_awareness_count"] = coded_data[q17_positive_awareness_cols].sum(axis=1)
    coded_data["No_sustainability_report_awareness"] = coded_data[Q17_NO_REPORT_COL]
    coded_data["Any_sustainability_report_awareness"] = (
        coded_data["Sustainability_report_awareness_count"] > 0
    ).astype(int)
    coded_data["Regulatory_SDG_impact_yes"] = coded_data[Q18_REGULATORY_YES_COL]
    coded_data["Economic_sustainability_importance"] = coded_data[Q20_COL]
    coded_data["Environmental_sustainability_importance"] = coded_data[Q21_COL]
    coded_data["Social_sustainability_importance"] = coded_data[Q22_COL]
    coded_data["Sustainability_importance_mean"] = coded_data[
        [
            "Economic_sustainability_importance",
            "Environmental_sustainability_importance",
            "Social_sustainability_importance",
        ]
    ].mean(axis=1)
    coded_data["Sustainability_importance_high_count"] = (
        coded_data[
            [
                "Economic_sustainability_importance",
                "Environmental_sustainability_importance",
                "Social_sustainability_importance",
            ]
        ] >= 4
    ).sum(axis=1)

    rq2_index_cols = [
        "Pressure_count",
        "Any_pressure",
        "Support_needed_count",
        "Any_support_needed",
        "Improvement_motivators_count",
        "No_improvement_indicator",
        "Any_positive_improvement_motivation",
        "External_funding_count",
        "No_external_funding_indicator",
        "Any_external_funding",
        "Sustainable_success_orientation",
        "Sustainability_report_awareness_count",
        "No_sustainability_report_awareness",
        "Any_sustainability_report_awareness",
        "Regulatory_SDG_impact_yes",
        "Economic_sustainability_importance",
        "Environmental_sustainability_importance",
        "Social_sustainability_importance",
        "Sustainability_importance_mean",
        "Sustainability_importance_high_count",
    ]

    rq2_index_summary = coded_data[rq2_index_cols].describe().T
    rq2_index_summary["median"] = coded_data[rq2_index_cols].median()
    rq2_index_summary = rq2_index_summary[
        ["count", "mean", "std", "min", "25%", "median", "50%", "75%", "max"]
    ]
    rq2_index_summary_display = rq2_index_summary.round(3)

    count_index_cols = [
        "Pressure_count",
        "Support_needed_count",
        "Improvement_motivators_count",
        "External_funding_count",
        "Sustainability_report_awareness_count",
        "Sustainability_importance_high_count",
    ]

    distribution_tables = {
        column: frequency_table(coded_data, column)
        for column in count_index_cols
    }

    expected_ranges = {
        "Pressure_count": (0, len(q15_cols)),
        "Support_needed_count": (0, len(q27_cols)),
        "Improvement_motivators_count": (0, len(q31_positive_motivation_cols)),
        "External_funding_count": (0, len(q34_external_funding_source_cols)),
        "Sustainability_report_awareness_count": (0, len(q17_positive_awareness_cols)),
        "Sustainability_importance_high_count": (0, 3),
        "Economic_sustainability_importance": (1, 5),
        "Environmental_sustainability_importance": (1, 5),
        "Social_sustainability_importance": (1, 5),
        "Sustainability_importance_mean": (1, 5),
    }

    range_validation_rows = []
    for column, (expected_minimum, expected_maximum) in expected_ranges.items():
        observed_minimum = coded_data[column].min()
        observed_maximum = coded_data[column].max()
        range_validation_rows.append({
            "index": column,
            "expected_min": expected_minimum,
            "expected_max": expected_maximum,
            "observed_min": observed_minimum,
            "observed_max": observed_maximum,
            "range_valid": bool(observed_minimum >= expected_minimum and observed_maximum <= expected_maximum),
        })

    rq2_index_range_validation = pd.DataFrame(range_validation_rows)

    rq2_index_missing = (
        coded_data[rq2_index_cols]
        .isna()
        .sum()
        .rename_axis("index")
        .reset_index(name="missing_n")
    )
    rq2_no_missing = rq2_index_missing["missing_n"].sum() == 0

    binary_derived_cols = [
        "Any_pressure",
        "Any_support_needed",
        "No_improvement_indicator",
        "Any_positive_improvement_motivation",
        "No_external_funding_indicator",
        "Any_external_funding",
        "Sustainable_success_orientation",
        "No_sustainability_report_awareness",
        "Any_sustainability_report_awareness",
        "Regulatory_SDG_impact_yes",
    ]

    binary_validation_rows = []
    for column in binary_derived_cols:
        values = sorted(coded_data[column].dropna().unique())
        binary_validation_rows.append({
            "index": column,
            "unique_values": values,
            "is_binary_0_1": set(values).issubset({0, 1}),
            "missing_n": int(coded_data[column].isna().sum()),
        })

    rq2_binary_index_validation = pd.DataFrame(binary_validation_rows)
    binary_indices_valid = (
        rq2_binary_index_validation["is_binary_0_1"].all()
        and rq2_binary_index_validation["missing_n"].sum() == 0
    )

    logic_validation_rows = []
    add_logic_check(
        logic_validation_rows,
        "Any_pressure equals Pressure_count > 0",
        (coded_data["Any_pressure"] == (coded_data["Pressure_count"] > 0).astype(int)).all(),
    )
    add_logic_check(
        logic_validation_rows,
        "Any_support_needed equals Support_needed_count > 0",
        (coded_data["Any_support_needed"] == (coded_data["Support_needed_count"] > 0).astype(int)).all(),
    )
    add_logic_check(
        logic_validation_rows,
        "Any_positive_improvement_motivation equals Improvement_motivators_count > 0",
        (
            coded_data["Any_positive_improvement_motivation"]
            == (coded_data["Improvement_motivators_count"] > 0).astype(int)
        ).all(),
    )
    add_logic_check(
        logic_validation_rows,
        "No_improvement_indicator complements Any_positive_improvement_motivation",
        (
            coded_data["No_improvement_indicator"]
            == 1 - coded_data["Any_positive_improvement_motivation"]
        ).all(),
    )
    add_logic_check(
        logic_validation_rows,
        "Any_external_funding equals External_funding_count > 0",
        (coded_data["Any_external_funding"] == (coded_data["External_funding_count"] > 0).astype(int)).all(),
    )
    add_logic_check(
        logic_validation_rows,
        "No_external_funding_indicator complements Any_external_funding",
        (
            coded_data["No_external_funding_indicator"]
            == 1 - coded_data["Any_external_funding"]
        ).all(),
    )
    add_logic_check(
        logic_validation_rows,
        "Any_sustainability_report_awareness equals Sustainability_report_awareness_count > 0",
        (
            coded_data["Any_sustainability_report_awareness"]
            == (coded_data["Sustainability_report_awareness_count"] > 0).astype(int)
        ).all(),
    )
    add_logic_check(
        logic_validation_rows,
        "No_sustainability_report_awareness complements Any_sustainability_report_awareness",
        (
            coded_data["No_sustainability_report_awareness"]
            == 1 - coded_data["Any_sustainability_report_awareness"]
        ).all(),
    )

    rq2_logic_validation = pd.DataFrame(logic_validation_rows)
    logic_valid = rq2_logic_validation["valid"].all()
    range_valid = rq2_index_range_validation["range_valid"].all()
    rq2_step2_valid = rq2_no_missing and range_valid and binary_indices_valid and logic_valid

    print("\n2) New RQ2 derived indices:")
    for column in rq2_index_cols:
        print(" -", column)

    print("\n3) Descriptive statistics for derived RQ2 indices:")
    print(rq2_index_summary_display.to_string())

    print("\n4) Distributions of key count indices:")
    for column, table in distribution_tables.items():
        print(f"\n{column}")
        print(table.to_string(index=False))

    print("\n5) Range validation for derived indices:")
    print(rq2_index_range_validation.to_string(index=False))

    print("\n6) Missing values in derived indices:")
    print(rq2_index_missing.to_string(index=False))

    print("\n7) Validation of binary derived indicators:")
    print(rq2_binary_index_validation.to_string(index=False))

    print("\n8) Logical consistency validation for derived indices:")
    print(rq2_logic_validation.to_string(index=False))

    print("\n9) Summary of key values:")
    print("Mean number of sustainability pressure sources:", round(coded_data["Pressure_count"].mean(), 3))
    print("Mean number of support types needed:", round(coded_data["Support_needed_count"].mean(), 3))
    print("Mean number of positive improvement motives:", round(coded_data["Improvement_motivators_count"].mean(), 3))
    print("Mean number of external funding sources used:", round(coded_data["External_funding_count"].mean(), 3))
    print("Mean sustainability importance across q20-q22:", round(coded_data["Sustainability_importance_mean"].mean(), 3))
    print("Share of firms with at least one external funding source:", round(coded_data["Any_external_funding"].mean() * 100, 2), "%")
    print("Share of firms measuring success through both profit and sustainable goals:", round(coded_data["Sustainable_success_orientation"].mean() * 100, 2), "%")
    print("Share of firms considering that legal inclusion of SDGs would affect results:", round(coded_data["Regulatory_SDG_impact_yes"].mean() * 100, 2), "%")

    print("\n" + "=" * 100)
    print("RQ2 - STEP 2 VALIDATION SUMMARY")
    print("=" * 100)
    print("Derived indices have no missing values:", bool(rq2_no_missing))
    print("All derived-index ranges are valid:", bool(range_valid))
    print("All binary derived indicators are valid:", bool(binary_indices_valid))
    print("Logical consistency of derived indices is valid:", bool(logic_valid))
    print("\nRQ2 Step 2 fully consistent:", bool(rq2_step2_valid))

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        coded_data.to_excel(writer, sheet_name="encoded_with_RQ2_indices", index=False)
        rq2_index_summary_display.to_excel(writer, sheet_name="index_summary")
        rq2_index_range_validation.to_excel(writer, sheet_name="range_validation", index=False)
        rq2_index_missing.to_excel(writer, sheet_name="missing_validation", index=False)
        rq2_binary_index_validation.to_excel(writer, sheet_name="binary_validation", index=False)
        rq2_logic_validation.to_excel(writer, sheet_name="logic_validation", index=False)
        for column, table in distribution_tables.items():
            table.to_excel(writer, sheet_name=column[:31], index=False)

    print("\nRQ2 Step 2 results saved to:")
    print(OUTPUT_PATH)


if __name__ == "__main__":
    main()
