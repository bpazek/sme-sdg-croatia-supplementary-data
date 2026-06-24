"""
RQ1 Step 1: Validation of q36 and q37 SDG variables.

This script validates the coded survey variables used to answer RQ1:
"Which SDGs dominate the operations and sustainability efforts of Croatian SMEs?"

The script checks whether the dataset contains the expected 17 SDG relevance
variables from q36 and the expected 17 direct SDG connection variables from q37.
It also verifies the value ranges, missing values, logical consistency of the
"none directly related" option, empty q37 rows, and the distribution of the
number of directly connected SDGs.

Input:
    Survey_Results.xlsx, sheet "encoded_responses_corrected".

Output:
    Printed validation summary for RQ1 Step 1.
"""

import pandas as pd


file_path = "/content/Survey_Results.xlsx"

coded_data = pd.read_excel(
    file_path,
    sheet_name="encoded_responses_corrected"
)

print("RQ1 - STEP 1")
print("Validation of q36 and q37 SDG variables")
print("=" * 80)
print("Dataset dimensions:", coded_data.shape)
print("Number of respondents / firms:", len(coded_data))

sdg_relevance_cols = [
    col for col in coded_data.columns
    if col.startswith("q36_relevantnost_sdg__")
]

sdg_direct_cols = [
    col for col in coded_data.columns
    if col.startswith("q37_sdg_vezani_uz_poslovanje__")
    and col != "q37_sdg_vezani_uz_poslovanje__none_directly_related"
]

none_sdg_col = "q37_sdg_vezani_uz_poslovanje__none_directly_related"

if none_sdg_col not in coded_data.columns:
    raise ValueError(f"Expected q37 column is missing: {none_sdg_col}")

print("\n1) Detected SDG columns:")
print("Number of q36 SDG relevance columns:", len(sdg_relevance_cols))
print("Number of q37 direct SDG connection columns:", len(sdg_direct_cols))
print("'None directly related' variable exists:", none_sdg_col in coded_data.columns)

q36_has_17_cols = len(sdg_relevance_cols) == 17
q37_has_17_cols = len(sdg_direct_cols) == 17

print("\n2) Number of SDG columns:")
print("q36 has exactly 17 SDG columns:", q36_has_17_cols)
print("q37 has exactly 17 SDG columns:", q37_has_17_cols)

if not q36_has_17_cols:
    print("\nWARNING: q36 does not have exactly 17 SDG columns.")
    print("Detected q36 columns:")
    for col in sdg_relevance_cols:
        print(" -", col)

if not q37_has_17_cols:
    print("\nWARNING: q37 does not have exactly 17 SDG columns.")
    print("Detected q37 columns:")
    for col in sdg_direct_cols:
        print(" -", col)

print("\n3) Values in q36 SDG relevance variables:")

q36_value_check = {}

for col in sdg_relevance_cols:
    values = sorted(coded_data[col].dropna().unique())
    q36_value_check[col] = values
    print(col, values)

q36_valid_values = all(
    set(values).issubset({1, 2, 3, 4, 5})
    for values in q36_value_check.values()
)

print("\n4) Values in q37 direct SDG connection variables:")

q37_value_check = {}

for col in sdg_direct_cols + [none_sdg_col]:
    values = sorted(coded_data[col].dropna().unique())
    q37_value_check[col] = values
    print(col, values)

q37_valid_values = all(
    set(values).issubset({0, 1})
    for values in q37_value_check.values()
)

q36_missing_by_col = coded_data[sdg_relevance_cols].isna().sum()
q37_missing_by_col = coded_data[sdg_direct_cols + [none_sdg_col]].isna().sum()

q36_missing_total = int(q36_missing_by_col.sum())
q37_missing_total = int(q37_missing_by_col.sum())

print("\n5) Missing values:")
print("Total missing values in q36:", q36_missing_total)
print("Total missing values in q37:", q37_missing_total)

if q36_missing_total > 0:
    print("\nq36 columns with missing values:")
    print(q36_missing_by_col[q36_missing_by_col > 0])

if q37_missing_total > 0:
    print("\nq37 columns with missing values:")
    print(q37_missing_by_col[q37_missing_by_col > 0])

q36_no_missing = q36_missing_total == 0
q37_no_missing = q37_missing_total == 0

coded_data["q37_direct_sdg_count_check"] = coded_data[sdg_direct_cols].sum(axis=1)

q37_contradictions = coded_data[
    (coded_data[none_sdg_col] == 1)
    & (coded_data["q37_direct_sdg_count_check"] > 0)
].copy()

print("\n6) Logical consistency check for q37:")
print(
    "Number of contradictory q37 rows "
    "('none directly related' = 1 and at least one SDG = 1):",
    len(q37_contradictions)
)

if len(q37_contradictions) > 0:
    print("\nContradictory q37 rows:")
    print(
        q37_contradictions[
            sdg_direct_cols + [none_sdg_col, "q37_direct_sdg_count_check"]
        ].to_string(index=False)
    )

q37_no_contradictions = len(q37_contradictions) == 0

q37_empty_rows = coded_data[
    (coded_data["q37_direct_sdg_count_check"] == 0)
    & (coded_data[none_sdg_col] == 0)
].copy()

print("\n7) Empty q37 row check:")
print(
    "Number of rows with no selected SDG and no 'none directly related' option:",
    len(q37_empty_rows)
)

if len(q37_empty_rows) > 0:
    print("\nEmpty q37 rows:")
    print(
        q37_empty_rows[
            sdg_direct_cols + [none_sdg_col, "q37_direct_sdg_count_check"]
        ].to_string(index=False)
    )

q37_no_empty_rows = len(q37_empty_rows) == 0

q37_direct_count_distribution = (
    coded_data["q37_direct_sdg_count_check"]
    .value_counts()
    .sort_index()
    .rename_axis("number_of_directly_related_SDGs")
    .reset_index(name="n")
)

q37_direct_count_distribution["percent"] = (
    q37_direct_count_distribution["n"] / len(coded_data) * 100
).round(2)

print("\n8) Distribution of the number of directly connected SDGs:")
print(q37_direct_count_distribution.to_string(index=False))

rq1_step1_valid = (
    q36_has_17_cols
    and q37_has_17_cols
    and q36_valid_values
    and q37_valid_values
    and q36_no_missing
    and q37_no_missing
    and q37_no_contradictions
    and q37_no_empty_rows
)

print("\n" + "=" * 80)
print("RQ1 - STEP 1 VALIDATION SUMMARY")
print("=" * 80)
print("q36 has exactly 17 SDG columns:", q36_has_17_cols)
print("q37 has exactly 17 SDG columns:", q37_has_17_cols)
print("q36 values are within {1, 2, 3, 4, 5}:", q36_valid_values)
print("q37 values are binary {0, 1}:", q37_valid_values)
print("q36 has no missing values:", q36_no_missing)
print("q37 has no missing values:", q37_no_missing)
print("q37 has no contradictions with 'none directly related':", q37_no_contradictions)
print("q37 has no empty rows:", q37_no_empty_rows)
print("\nRQ1 - Step 1 fully consistent:", rq1_step1_valid)
