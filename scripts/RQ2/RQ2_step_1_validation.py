"""
RQ2 Step 1: Validation of motivational, institutional and business-context variables.

This script validates the variables used to answer RQ2:
"What are the most important motivators for SDG involvement in SME operations?"

The script loads the coded survey dataset, reconstructs the SDG outcome variables
from RQ1, validates the RQ2 binary variable groups, checks Likert-scale variables,
verifies single-select logic, detects logical contradictions, checks empty
multi-select groups, verifies other_selected and other_text alignment, and saves
validation outputs to an Excel file.

Input:
    Survey_Results.xlsx, sheet "encoded_responses_corrected".

Output:
    RQ2_step1_validation.xlsx.
"""

import numpy as np
import pandas as pd


FILE_PATH = "/content/Survey_Results.xlsx"
SHEET_NAME = "encoded_responses_corrected"
OUTPUT_PATH = "/content/RQ2_step1_validation.xlsx"
NONE_SDG_COL = "q37_sdg_vezani_uz_poslovanje__none_directly_related"


def unique_values_report(data, columns):
    """Validate binary columns and summarize unique values, missingness and frequencies."""
    rows = []

    for column in columns:
        values = sorted(data[column].dropna().unique())
        is_binary = set(values).issubset({0, 1})
        rows.append({
            "variable": column,
            "unique_values": values,
            "is_binary_0_1": is_binary,
            "missing_n": int(data[column].isna().sum()),
            "selected_n": int(data[column].sum()) if is_binary else np.nan,
            "selected_percent": float(data[column].mean() * 100) if is_binary else np.nan,
        })

    return pd.DataFrame(rows)


def get_columns(data, prefix, exclude_other_text=False, exclude_column=None):
    """Return columns that start with a prefix, with optional exclusions."""
    columns = [column for column in data.columns if column.startswith(prefix)]

    if exclude_other_text:
        columns = [column for column in columns if not column.endswith("__other_text")]

    if exclude_column is not None:
        columns = [column for column in columns if column != exclude_column]

    return columns


def check_required_columns(data, columns, label):
    """Raise an error if a required group has no detected columns."""
    if len(columns) == 0:
        raise ValueError(f"No columns were detected for {label}.")


def make_single_select_checks(data, single_select_groups):
    """Check whether single-select blocks contain exactly one selected option per row."""
    rows = []

    for group_name, columns in single_select_groups.items():
        row_sums = data[columns].sum(axis=1)
        rows.append({
            "group": group_name,
            "n_columns": len(columns),
            "min_sum": int(row_sums.min()),
            "max_sum": int(row_sums.max()),
            "rows_with_sum_0": int((row_sums == 0).sum()),
            "rows_with_sum_1": int((row_sums == 1).sum()),
            "rows_with_sum_greater_than_1": int((row_sums > 1).sum()),
            "valid_exactly_one": bool((row_sums == 1).all()),
        })

    return pd.DataFrame(rows)


def detect_exclusive_option_contradictions(data, columns, exclusive_column):
    """Detect rows where an exclusive option is selected together with positive options."""
    positive_columns = [column for column in columns if column != exclusive_column]
    return data[
        (data[exclusive_column] == 1)
        & (data[positive_columns].sum(axis=1) > 0)
    ].copy()


def make_empty_group_summary(data, groups):
    """Summarize rows with no selected option in each multi-select group."""
    rows = []

    for group_name, columns in groups.items():
        row_sums = data[columns].sum(axis=1)
        rows.append({
            "group": group_name,
            "rows_with_no_selected_option": int((row_sums == 0).sum()),
            "percent_with_no_selected_option": float((row_sums == 0).mean() * 100),
        })

    summary = pd.DataFrame(rows)
    summary["percent_with_no_selected_option"] = summary["percent_with_no_selected_option"].round(2)
    return summary


def check_other_alignment(data, other_pairs):
    """Check alignment between other_selected indicators and other_text variables."""
    rows = []
    tables = {}

    for pair in other_pairs:
        block = pair["block"]
        indicator_column = pair["indicator"]
        text_column = pair["text"]

        if indicator_column not in data.columns or text_column not in data.columns:
            rows.append({
                "block": block,
                "indicator_col": indicator_column,
                "text_col": text_column,
                "indicator_exists": indicator_column in data.columns,
                "text_exists": text_column in data.columns,
                "nonempty_text_n": np.nan,
                "selected_with_text_n": np.nan,
                "selected_without_text_n": np.nan,
                "text_without_selected_n": np.nan,
            })
            continue

        clean_text = data[text_column].astype("string").str.strip()
        nonempty_text_mask = clean_text.notna() & (clean_text != "")
        selected_mask = data[indicator_column] == 1

        selected_with_text = data[selected_mask & nonempty_text_mask].copy()
        selected_without_text = data[selected_mask & ~nonempty_text_mask].copy()
        text_without_selected = data[(~selected_mask) & nonempty_text_mask].copy()

        rows.append({
            "block": block,
            "indicator_col": indicator_column,
            "text_col": text_column,
            "indicator_exists": True,
            "text_exists": True,
            "nonempty_text_n": int(nonempty_text_mask.sum()),
            "selected_with_text_n": len(selected_with_text),
            "selected_without_text_n": len(selected_without_text),
            "text_without_selected_n": len(text_without_selected),
        })

        tables[f"{block}_selected_without_text"] = selected_without_text
        tables[f"{block}_text_without_selected"] = text_without_selected

    return pd.DataFrame(rows), tables


def print_problem_rows(title, data, columns):
    """Print selected columns for problem rows, or state that no rows were found."""
    if len(data) > 0:
        print(f"\n{title}:")
        print(data[columns].to_string(index=False))
    else:
        print(f"\n{title}: no contradictions detected.")


def main():
    coded_data = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)

    print("RQ2 - STEP 1")
    print("Validation of motivational and institutional variables")
    print("=" * 90)
    print("Initial dataset dimensions:", coded_data.shape)
    print("Number of respondents / firms:", len(coded_data))

    sdg_relevance_cols = get_columns(coded_data, "q36_relevantnost_sdg__")
    sdg_direct_cols = get_columns(
        coded_data,
        "q37_sdg_vezani_uz_poslovanje__",
        exclude_column=NONE_SDG_COL,
    )

    if NONE_SDG_COL not in coded_data.columns:
        raise ValueError(f"Expected column is missing: {NONE_SDG_COL}")

    print("\n1) SDG blocks used for RQ2 outcomes:")
    print("Number of q36 SDG relevance columns:", len(sdg_relevance_cols))
    print("Number of q37 direct SDG columns:", len(sdg_direct_cols))
    print("The none_directly_related column exists:", NONE_SDG_COL in coded_data.columns)

    if len(sdg_relevance_cols) != 17:
        raise ValueError(f"Expected 17 q36 SDG relevance columns, found {len(sdg_relevance_cols)}.")

    if len(sdg_direct_cols) != 17:
        raise ValueError(f"Expected 17 q37 direct SDG columns, found {len(sdg_direct_cols)}.")

    coded_data["SDG_count"] = coded_data[sdg_direct_cols].sum(axis=1)
    coded_data["SDG_mean_relevance"] = coded_data[sdg_relevance_cols].mean(axis=1)
    coded_data["SDG_high_relevance_count"] = (coded_data[sdg_relevance_cols] >= 4).sum(axis=1)

    sdg_derived_cols = [
        "SDG_count",
        "SDG_mean_relevance",
        "SDG_high_relevance_count",
    ]

    sdg_derived_missing = coded_data[sdg_derived_cols].isna().sum()
    sdg_count_valid = coded_data["SDG_count"].between(0, 17).all()
    sdg_mean_relevance_valid = coded_data["SDG_mean_relevance"].between(1, 5).all()
    sdg_high_relevance_count_valid = coded_data["SDG_high_relevance_count"].between(0, 17).all()
    sdg_none_consistent = (
        ((coded_data["SDG_count"] == 0) & (coded_data[NONE_SDG_COL] == 1))
        | ((coded_data["SDG_count"] > 0) & (coded_data[NONE_SDG_COL] == 0))
    ).all()

    print("\n2) Validation of derived SDG measures:")
    print("Missing values:")
    print(sdg_derived_missing.to_string())
    print("SDG_count is within the range 0-17:", bool(sdg_count_valid))
    print("SDG_mean_relevance is within the range 1-5:", bool(sdg_mean_relevance_valid))
    print("SDG_high_relevance_count is within the range 0-17:", bool(sdg_high_relevance_count_valid))
    print("SDG_count is consistent with none_directly_related:", bool(sdg_none_consistent))

    q15_cols = get_columns(coded_data, "q15_pritisak_na_odrzivost__")
    q27_binary_cols = get_columns(coded_data, "q27_pomoc_za_odrzivost__", exclude_other_text=True)
    q31_binary_cols = get_columns(coded_data, "q31_motivi_promjena__", exclude_other_text=True)
    q34_binary_cols = get_columns(coded_data, "q34_eksterna_sredstva_3_godine__", exclude_other_text=True)
    q16_cols = get_columns(coded_data, "q16_mjerenje_uspjeha__")
    q17_cols = get_columns(coded_data, "q17_izvjestaj_o_odrzivosti__")
    q18_cols = get_columns(coded_data, "q18_zakonodavni_utjecaj_sdg__")

    q20_q22_cols = [
        column for column in [
            "q20_vaznost_ekonomske_odrzivosti",
            "q21_vaznost_okolisne_odrzivosti",
            "q22_vaznost_drustvene_odrzivosti",
        ]
        if column in coded_data.columns
    ]

    rq2_binary_groups = {
        "q15_pritisak_na_odrzivost": q15_cols,
        "q27_pomoc_za_odrzivost": q27_binary_cols,
        "q31_motivi_promjena": q31_binary_cols,
        "q34_eksterna_sredstva_3_godine": q34_binary_cols,
        "q16_mjerenje_uspjeha": q16_cols,
        "q17_izvjestaj_o_odrzivosti": q17_cols,
        "q18_zakonodavni_utjecaj_sdg": q18_cols,
    }

    print("\n3) Number of columns by RQ2 binary group:")
    for group_name, columns in rq2_binary_groups.items():
        print(f"{group_name}: {len(columns)} columns")

    print("\n4) Likert value variables for RQ2:")
    for column in q20_q22_cols:
        print(" -", column)

    empty_defined_groups = [
        group_name for group_name, columns in rq2_binary_groups.items()
        if len(columns) == 0
    ]

    if empty_defined_groups:
        raise ValueError(f"The following RQ2 groups have no detected columns: {empty_defined_groups}")

    all_rq2_binary_cols = []
    for columns in rq2_binary_groups.values():
        all_rq2_binary_cols.extend(columns)
    all_rq2_binary_cols = list(dict.fromkeys(all_rq2_binary_cols))

    rq2_binary_validation = unique_values_report(coded_data, all_rq2_binary_cols)
    all_binary_valid = rq2_binary_validation["is_binary_0_1"].all()
    binary_missing_total = rq2_binary_validation["missing_n"].sum()

    print("\n5) Binary validation summary for RQ2 variables:")
    print("All RQ2 binary variables contain only 0/1 values:", bool(all_binary_valid))
    print("Total missing values in RQ2 binary variables:", int(binary_missing_total))

    if not all_binary_valid or binary_missing_total > 0:
        print("\nVariables with problems:")
        print(
            rq2_binary_validation[
                (~rq2_binary_validation["is_binary_0_1"])
                | (rq2_binary_validation["missing_n"] > 0)
            ].to_string(index=False)
        )

    print("\n6) Frequencies of selected options by RQ2 group:")
    rq2_group_frequencies = {}

    for group_name, columns in rq2_binary_groups.items():
        frequency = (
            rq2_binary_validation[
                rq2_binary_validation["variable"].isin(columns)
            ][["variable", "selected_n", "selected_percent"]]
            .sort_values("selected_percent", ascending=False)
            .reset_index(drop=True)
        )
        frequency["selected_percent"] = frequency["selected_percent"].round(2)
        rq2_group_frequencies[group_name] = frequency

        print(f"\n{group_name}")
        print(frequency.to_string(index=False))

    likert_rows = []

    for column in q20_q22_cols:
        values = sorted(coded_data[column].dropna().unique())
        valid_likert = set(values).issubset({1, 2, 3, 4, 5})
        likert_rows.append({
            "variable": column,
            "unique_values": values,
            "valid_likert_1_5": valid_likert,
            "missing_n": int(coded_data[column].isna().sum()),
            "mean": coded_data[column].mean(),
            "median": coded_data[column].median(),
            "std": coded_data[column].std(ddof=1),
        })

    rq2_likert_validation = pd.DataFrame(likert_rows)

    if len(rq2_likert_validation) > 0:
        print("\n7) Validation of q20-q22 Likert variables:")
        likert_display = rq2_likert_validation.copy()
        for column in ["mean", "median", "std"]:
            likert_display[column] = likert_display[column].round(3)
        print(likert_display.to_string(index=False))
        likert_valid = rq2_likert_validation["valid_likert_1_5"].all()
        likert_missing_total = rq2_likert_validation["missing_n"].sum()
    else:
        print("\n7) No q20-q22 Likert variables were detected.")
        likert_valid = True
        likert_missing_total = 0

    print("\n8) Single-select logic check for q16 and q18:")
    single_select_groups = {
        "q16_mjerenje_uspjeha": q16_cols,
        "q18_zakonodavni_utjecaj_sdg": q18_cols,
    }
    single_select_checks_df = make_single_select_checks(coded_data, single_select_groups)
    print(single_select_checks_df.to_string(index=False))

    q16_valid = single_select_checks_df.loc[
        single_select_checks_df["group"] == "q16_mjerenje_uspjeha",
        "valid_exactly_one",
    ].iloc[0]
    q18_valid = single_select_checks_df.loc[
        single_select_checks_df["group"] == "q18_zakonodavni_utjecaj_sdg",
        "valid_exactly_one",
    ].iloc[0]

    print("\n9) Logical contradiction check for exclusive options:")
    contradiction_tables = {}

    q31_none_col = "q31_motivi_promjena__nista_poboljsano"
    q34_none_col = "q34_eksterna_sredstva_3_godine__bez_eksternih_sredstava"
    q17_no_col = "q17_izvjestaj_o_odrzivosti__ne"

    for required_column in [q31_none_col, q34_none_col, q17_no_col]:
        if required_column not in coded_data.columns:
            raise ValueError(f"Expected column is missing: {required_column}")

    q31_contradictions = detect_exclusive_option_contradictions(
        coded_data,
        q31_binary_cols,
        q31_none_col,
    )
    q34_contradictions = detect_exclusive_option_contradictions(
        coded_data,
        q34_binary_cols,
        q34_none_col,
    )
    q17_contradictions = detect_exclusive_option_contradictions(
        coded_data,
        q17_cols,
        q17_no_col,
    )

    contradiction_tables["q31_no_improvement_with_other_motives"] = q31_contradictions
    contradiction_tables["q34_no_external_funding_with_other_sources"] = q34_contradictions
    contradiction_tables["q17_no_report_with_yes_options"] = q17_contradictions

    contradiction_summary = pd.DataFrame([
        {"contradiction_type": name, "number_of_rows": len(table)}
        for name, table in contradiction_tables.items()
    ])
    print(contradiction_summary.to_string(index=False))

    print("\n10) Rows with potential contradictions:")
    print_problem_rows("q31", q31_contradictions, q31_binary_cols)
    print_problem_rows("q34", q34_contradictions, q34_binary_cols)
    print_problem_rows("q17", q17_contradictions, q17_cols)

    print("\n11) Rows with fully empty multi-select groups:")
    multi_select_groups = {
        "q15_pritisak_na_odrzivost": q15_cols,
        "q27_pomoc_za_odrzivost": q27_binary_cols,
        "q31_motivi_promjena": q31_binary_cols,
        "q34_eksterna_sredstva_3_godine": q34_binary_cols,
        "q17_izvjestaj_o_odrzivosti": q17_cols,
    }
    empty_group_rows_df = make_empty_group_summary(coded_data, multi_select_groups)
    print(empty_group_rows_df.to_string(index=False))
    print(
        "\nNote: Empty rows in q15 or q27 are not automatically errors. "
        "They may indicate that the respondent did not select any pressure source "
        "or any type of support."
    )

    print("\n12) Alignment check for other_selected and other_text:")
    other_pairs = [
        {
            "block": "q27_pomoc_za_odrzivost",
            "indicator": "q27_pomoc_za_odrzivost__other_selected",
            "text": "q27_pomoc_za_odrzivost__other_text",
        },
        {
            "block": "q31_motivi_promjena",
            "indicator": "q31_motivi_promjena__other_selected",
            "text": "q31_motivi_promjena__other_text",
        },
        {
            "block": "q34_eksterna_sredstva_3_godine",
            "indicator": "q34_eksterna_sredstva_3_godine__other_selected",
            "text": "q34_eksterna_sredstva_3_godine__other_text",
        },
    ]

    other_alignment_df, other_alignment_tables = check_other_alignment(coded_data, other_pairs)
    print(other_alignment_df.to_string(index=False))

    other_alignment_ok = (
        other_alignment_df["indicator_exists"].fillna(False).all()
        and other_alignment_df["text_exists"].fillna(False).all()
        and other_alignment_df["selected_without_text_n"].fillna(0).sum() == 0
        and other_alignment_df["text_without_selected_n"].fillna(0).sum() == 0
    )
    print("Other_selected / other_text alignment is valid:", bool(other_alignment_ok))

    no_contradictions = (
        len(q31_contradictions) == 0
        and len(q34_contradictions) == 0
        and len(q17_contradictions) == 0
    )
    sdg_derived_valid = (
        sdg_derived_missing.sum() == 0
        and sdg_count_valid
        and sdg_mean_relevance_valid
        and sdg_high_relevance_count_valid
        and sdg_none_consistent
    )
    rq2_step1_valid = (
        sdg_derived_valid
        and all_binary_valid
        and binary_missing_total == 0
        and likert_valid
        and likert_missing_total == 0
        and q16_valid
        and q18_valid
        and no_contradictions
        and other_alignment_ok
    )

    print("\n" + "=" * 90)
    print("RQ2 STEP 1 VALIDATION SUMMARY")
    print("=" * 90)
    print("Derived SDG measures are valid:", bool(sdg_derived_valid))
    print("All RQ2 binary variables are binary:", bool(all_binary_valid))
    print("RQ2 binary variables have no missing values:", bool(binary_missing_total == 0))
    print("q20-q22 Likert variables are valid:", bool(likert_valid))
    print("q20-q22 have no missing values:", bool(likert_missing_total == 0))
    print("q16 single-select logic is valid:", bool(q16_valid))
    print("q18 single-select logic is valid:", bool(q18_valid))
    print("No q31/q34/q17 contradictions detected:", bool(no_contradictions))
    print("Other_selected / other_text alignment is valid:", bool(other_alignment_ok))
    print("\nRQ2 Step 1 fully consistent:", bool(rq2_step1_valid))

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        rq2_binary_validation.to_excel(writer, sheet_name="binary_validation", index=False)

        for group_name, frequency in rq2_group_frequencies.items():
            sheet_name = group_name[:31]
            frequency.to_excel(writer, sheet_name=sheet_name, index=False)

        rq2_likert_validation.to_excel(writer, sheet_name="likert_q20_q22", index=False)
        single_select_checks_df.to_excel(writer, sheet_name="single_select_checks", index=False)
        contradiction_summary.to_excel(writer, sheet_name="contradiction_summary", index=False)
        empty_group_rows_df.to_excel(writer, sheet_name="empty_multiselect_rows", index=False)
        other_alignment_df.to_excel(writer, sheet_name="other_alignment", index=False)
        coded_data[sdg_derived_cols].describe().T.to_excel(writer, sheet_name="sdg_derived_summary")

    print("\nRQ2 Step 1 validation results saved to:")
    print(OUTPUT_PATH)


if __name__ == "__main__":
    main()
