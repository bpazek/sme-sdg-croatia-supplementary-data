"""
Supplementary Correlation Analysis Step 1: Validation of q31 and q34 variables.

This script validates the coded survey variables used in the supplementary
analysis of the relationship between external funding sources and improvement
motives. It identifies q31 improvement-motive variables and q34 external-funding
variables, separates binary and text columns, validates binary coding, detects
none-like response options, checks row-sum distributions, tests logical
contradictions, verifies alignment between "other selected" indicators and
free-text responses, and saves the validation outputs to an Excel file.

Input:
    Survey_Results.xlsx, sheet "encoded_responses_corrected".

Output:
    Additional_Analysis_3_q31_q34_step1_validation_corrected.xlsx.
"""

import numpy as np
import pandas as pd


FILE_PATH = "/content/Survey_Results.xlsx"
SHEET_NAME = "encoded_responses_corrected"
OUTPUT_PATH = "/content/Additional_Analysis_3_q31_q34_step1_validation_corrected.xlsx"


def validate_binary_columns(data, columns):
    """Validate whether selected columns are binary and summarize their frequencies."""
    rows = []

    for column in columns:
        unique_values = sorted(data[column].dropna().unique())
        is_binary = set(unique_values).issubset({0, 1})
        rows.append({
            "variable": column,
            "unique_values": unique_values,
            "is_binary_0_1": is_binary,
            "missing_n": int(data[column].isna().sum()),
            "selected_n": int(data[column].sum()) if is_binary else np.nan,
            "selected_percent": float(data[column].mean() * 100) if is_binary else np.nan,
        })

    return pd.DataFrame(rows)


def identify_none_like_columns(columns):
    """Identify none-like binary columns based on their names."""
    return [
        column for column in columns
        if any(keyword in column.lower() for keyword in ["nista", "ništa", "bez", "none"])
    ]


def make_row_sum_distribution(data, row_sum_column, label):
    """Create a frequency distribution for the number of selected options per row."""
    distribution = (
        data[row_sum_column]
        .value_counts()
        .sort_index()
        .rename_axis(label)
        .reset_index(name="n")
    )
    distribution["percent"] = (distribution["n"] / len(data) * 100).round(2)
    return distribution


def check_none_contradictions(data, binary_columns, none_columns, block_name):
    """Check whether none-like options are selected together with positive options."""
    if len(none_columns) == 0:
        print(f"\n{block_name}: no none-like columns detected.")
        return pd.DataFrame()

    positive_columns = [column for column in binary_columns if column not in none_columns]
    contradiction_mask = (
        data[none_columns].sum(axis=1) > 0
    ) & (
        data[positive_columns].sum(axis=1) > 0
    )

    contradictions = data.loc[
        contradiction_mask,
        none_columns + positive_columns
    ].copy()

    print(f"\n11) {block_name}: none-like contradiction check")
    print("Rows with a none-like option and at least one positive option:", len(contradictions))

    if len(contradictions) > 0:
        print(contradictions.to_string(index=False))

    return contradictions


def check_other_text_alignment(data, binary_columns, text_columns, block_name):
    """Check whether other-text responses align with the corresponding other-selected indicator."""
    rows = []

    for text_column in text_columns:
        selected_column = text_column.replace("__other_text", "__other_selected")
        if selected_column not in binary_columns:
            selected_column = None

        clean_text = data[text_column].astype("string").str.strip()
        nonempty_text_n = int((clean_text.notna() & (clean_text != "")).sum())

        if selected_column is not None:
            selected_with_text_n = int(((data[selected_column] == 1) & clean_text.notna() & (clean_text != "")).sum())
            selected_without_text_n = int(((data[selected_column] == 1) & (clean_text.isna() | (clean_text == ""))).sum())
            text_without_selected_n = int(((data[selected_column] == 0) & clean_text.notna() & (clean_text != "")).sum())
        else:
            selected_with_text_n = np.nan
            selected_without_text_n = np.nan
            text_without_selected_n = np.nan

        rows.append({
            "block": block_name,
            "text_col": text_column,
            "selected_col": selected_column,
            "nonempty_text_n": nonempty_text_n,
            "selected_with_text_n": selected_with_text_n,
            "selected_without_text_n": selected_without_text_n,
            "text_without_selected_n": text_without_selected_n,
        })

    return pd.DataFrame(rows)


def print_columns(title, columns):
    """Print a list of detected columns."""
    print(f"\n{title}")
    if len(columns) == 0:
        print("No columns detected.")
    else:
        for column in columns:
            print(" -", column)


def main():
    coded_data = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)

    print("SUPPLEMENTARY CORRELATION ANALYSIS - STEP 1")
    print("Validation of external funding and improvement-motive variables")
    print("=" * 90)
    print("Dataset dimensions:", coded_data.shape)
    print("Number of respondents / firms:", len(coded_data))

    q31_cols_all = [
        column for column in coded_data.columns
        if column.startswith("q31_motivi_promjena__")
    ]
    q34_cols_all = [
        column for column in coded_data.columns
        if column.startswith("q34_eksterna_sredstva_3_godine__")
    ]

    print_columns("1) Detected q31 columns: improvement motives", q31_cols_all)
    print("Number of q31 columns:", len(q31_cols_all))
    print_columns("2) Detected q34 columns: external funding sources", q34_cols_all)
    print("Number of q34 columns:", len(q34_cols_all))

    q31_text_cols = [column for column in q31_cols_all if column.endswith("__other_text")]
    q34_text_cols = [column for column in q34_cols_all if column.endswith("__other_text")]
    q31_binary_cols = [column for column in q31_cols_all if not column.endswith("__other_text")]
    q34_binary_cols = [column for column in q34_cols_all if not column.endswith("__other_text")]

    print_columns("3) q31 text columns", q31_text_cols)
    print_columns("4) q34 text columns", q34_text_cols)
    print("\nNumber of q31 binary columns:", len(q31_binary_cols))
    print("Number of q34 binary columns:", len(q34_binary_cols))

    q31_validation = validate_binary_columns(coded_data, q31_binary_cols)
    q34_validation = validate_binary_columns(coded_data, q34_binary_cols)

    q31_validation_display = q31_validation.copy()
    q34_validation_display = q34_validation.copy()
    q31_validation_display["selected_percent"] = q31_validation_display["selected_percent"].round(2)
    q34_validation_display["selected_percent"] = q34_validation_display["selected_percent"].round(2)

    print("\n5) Binary validation of q31 variables:")
    print(q31_validation_display.to_string(index=False))
    print("\n6) Binary validation of q34 variables:")
    print(q34_validation_display.to_string(index=False))

    q31_none_like_cols = identify_none_like_columns(q31_binary_cols)
    q34_none_like_cols = identify_none_like_columns(q34_binary_cols)

    print_columns("7) q31 none-like columns", q31_none_like_cols)
    print_columns("8) q34 none-like columns", q34_none_like_cols)

    coded_data["q31_binary_sum_check"] = coded_data[q31_binary_cols].sum(axis=1)
    coded_data["q34_binary_sum_check"] = coded_data[q34_binary_cols].sum(axis=1)

    q31_row_sum_freq = make_row_sum_distribution(coded_data, "q31_binary_sum_check", "q31_binary_sum")
    q34_row_sum_freq = make_row_sum_distribution(coded_data, "q34_binary_sum_check", "q34_binary_sum")

    print("\n9) Distribution of selected q31 options per respondent:")
    print(q31_row_sum_freq.to_string(index=False))
    print("\n10) Distribution of selected q34 options per respondent:")
    print(q34_row_sum_freq.to_string(index=False))

    q31_contradictions = check_none_contradictions(
        coded_data,
        q31_binary_cols,
        q31_none_like_cols,
        "q31 improvement motives"
    )
    q34_contradictions = check_none_contradictions(
        coded_data,
        q34_binary_cols,
        q34_none_like_cols,
        "q34 external funding sources"
    )

    q31_other_alignment = check_other_text_alignment(
        coded_data,
        q31_binary_cols,
        q31_text_cols,
        "q31"
    )
    q34_other_alignment = check_other_text_alignment(
        coded_data,
        q34_binary_cols,
        q34_text_cols,
        "q34"
    )

    print("\n12) Alignment of q31 other_text fields:")
    if len(q31_other_alignment) > 0:
        print(q31_other_alignment.to_string(index=False))
    else:
        print("No q31 other_text columns detected.")

    print("\n13) Alignment of q34 other_text fields:")
    if len(q34_other_alignment) > 0:
        print(q34_other_alignment.to_string(index=False))
    else:
        print("No q34 other_text columns detected.")

    q31_all_binary = q31_validation["is_binary_0_1"].all() if len(q31_validation) > 0 else False
    q34_all_binary = q34_validation["is_binary_0_1"].all() if len(q34_validation) > 0 else False
    q31_no_missing = q31_validation["missing_n"].sum() == 0 if len(q31_validation) > 0 else False
    q34_no_missing = q34_validation["missing_n"].sum() == 0 if len(q34_validation) > 0 else False
    q31_has_cols = len(q31_binary_cols) > 0
    q34_has_cols = len(q34_binary_cols) > 0
    q31_no_empty_rows = (coded_data["q31_binary_sum_check"] > 0).all()
    q34_no_empty_rows = (coded_data["q34_binary_sum_check"] > 0).all()
    q31_no_contradictions = len(q31_contradictions) == 0
    q34_no_contradictions = len(q34_contradictions) == 0

    q31_other_ok = (
        len(q31_other_alignment) == 0
        or (
            (q31_other_alignment["selected_without_text_n"] == 0).all()
            and (q31_other_alignment["text_without_selected_n"] == 0).all()
        )
    )
    q34_other_ok = (
        len(q34_other_alignment) == 0
        or (
            (q34_other_alignment["selected_without_text_n"] == 0).all()
            and (q34_other_alignment["text_without_selected_n"] == 0).all()
        )
    )

    step1_valid_full = (
        q31_has_cols
        and q34_has_cols
        and q31_all_binary
        and q34_all_binary
        and q31_no_missing
        and q34_no_missing
        and q31_no_empty_rows
        and q34_no_empty_rows
        and q31_no_contradictions
        and q34_no_contradictions
        and q31_other_ok
        and q34_other_ok
    )

    print("\n14) STEP 1 VALIDATION SUMMARY")
    print("q31 binary columns detected:", bool(q31_has_cols))
    print("q34 binary columns detected:", bool(q34_has_cols))
    print("q31 columns are binary:", bool(q31_all_binary))
    print("q34 columns are binary:", bool(q34_all_binary))
    print("q31 has no missing values:", bool(q31_no_missing))
    print("q34 has no missing values:", bool(q34_no_missing))
    print("q31 has no empty rows:", bool(q31_no_empty_rows))
    print("q34 has no empty rows:", bool(q34_no_empty_rows))
    print("q31 has no none-like contradictions:", bool(q31_no_contradictions))
    print("q34 has no none-like contradictions:", bool(q34_no_contradictions))
    print("q31 other_text is aligned:", bool(q31_other_ok))
    print("q34 other_text is aligned:", bool(q34_other_ok))
    print("\nSupplementary Correlation Analysis - Step 1 fully consistent:", bool(step1_valid_full))

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        q31_validation_display.to_excel(writer, sheet_name="q31_validation", index=False)
        q34_validation_display.to_excel(writer, sheet_name="q34_validation", index=False)
        q31_row_sum_freq.to_excel(writer, sheet_name="q31_row_sum", index=False)
        q34_row_sum_freq.to_excel(writer, sheet_name="q34_row_sum", index=False)
        q31_contradictions.to_excel(writer, sheet_name="q31_contradictions", index=False)
        q34_contradictions.to_excel(writer, sheet_name="q34_contradictions", index=False)
        q31_other_alignment.to_excel(writer, sheet_name="q31_other_alignment", index=False)
        q34_other_alignment.to_excel(writer, sheet_name="q34_other_alignment", index=False)

    print("\nValidation results saved to:")
    print(OUTPUT_PATH)
    print("\nSupplementary Correlation Analysis - Step 1 completed.")


if __name__ == "__main__":
    main()
