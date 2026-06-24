"""
RQ1 Step 2: Construction of derived SDG variables.

This script creates derived variables used to answer RQ1:
"Which SDGs dominate the operations and sustainability efforts of Croatian SMEs?"

The script uses the coded survey dataset and constructs three firm-level SDG
indicators: the number of SDGs directly connected to business operations,
the mean perceived relevance of all 17 SDGs, and the number of SDGs rated as
highly relevant. It then validates these variables, checks their ranges and
logical consistency, produces descriptive summaries and distributions, and
saves an Excel file with the derived variables and summary tables.

Input:
    Survey_Results.xlsx, sheet "encoded_responses_corrected".

Output:
    RQ1_with_SDG_derived_variables.xlsx.
"""

import pandas as pd


file_path = "/content/Survey_Results.xlsx"
output_path = "/content/RQ1_with_SDG_derived_variables.xlsx"

coded_data = pd.read_excel(
    file_path,
    sheet_name="encoded_responses_corrected"
)

print("RQ1 - STEP 2")
print("Construction of derived SDG variables")
print("=" * 80)

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

if len(sdg_relevance_cols) != 17:
    raise ValueError(
        f"Expected 17 q36 SDG relevance columns, "
        f"but found {len(sdg_relevance_cols)}."
    )

if len(sdg_direct_cols) != 17:
    raise ValueError(
        f"Expected 17 q37 direct SDG columns, "
        f"but found {len(sdg_direct_cols)}."
    )

if none_sdg_col not in coded_data.columns:
    raise ValueError(f"Expected q37 none column is missing: {none_sdg_col}")

print("Number of q36 SDG relevance columns:", len(sdg_relevance_cols))
print("Number of q37 direct SDG columns:", len(sdg_direct_cols))
print("Number of respondents / firms:", len(coded_data))

coded_data["SDG_count"] = coded_data[sdg_direct_cols].sum(axis=1)
coded_data["SDG_mean_relevance"] = coded_data[sdg_relevance_cols].mean(axis=1)
coded_data["SDG_high_relevance_count"] = (coded_data[sdg_relevance_cols] >= 4).sum(axis=1)

derived_sdg_cols = [
    "SDG_count",
    "SDG_mean_relevance",
    "SDG_high_relevance_count",
]

print("\n1) Derived SDG variables:")
print(coded_data[derived_sdg_cols].to_string(index=False))

derived_missing = coded_data[derived_sdg_cols].isna().sum()

print("\n2) Missing values in derived SDG variables:")
print(derived_missing.to_string())

derived_no_missing = derived_missing.sum() == 0

sdg_count_valid = coded_data["SDG_count"].between(0, 17).all()
sdg_mean_relevance_valid = coded_data["SDG_mean_relevance"].between(1, 5).all()
sdg_high_relevance_count_valid = coded_data["SDG_high_relevance_count"].between(0, 17).all()

print("\n3) Range checks for derived SDG variables:")
print("SDG_count is within the range 0-17:", bool(sdg_count_valid))
print("SDG_mean_relevance is within the range 1-5:", bool(sdg_mean_relevance_valid))
print("SDG_high_relevance_count is within the range 0-17:", bool(sdg_high_relevance_count_valid))

sdg_count_none_consistent = (
    (
        (coded_data["SDG_count"] == 0)
        & (coded_data[none_sdg_col] == 1)
    )
    | (
        (coded_data["SDG_count"] > 0)
        & (coded_data[none_sdg_col] == 0)
    )
).all()

sdg_count_none_inconsistent_rows = coded_data[
    ~(
        (
            (coded_data["SDG_count"] == 0)
            & (coded_data[none_sdg_col] == 1)
        )
        | (
            (coded_data["SDG_count"] > 0)
            & (coded_data[none_sdg_col] == 0)
        )
    )
].copy()

print("\n4) Logical consistency check for SDG_count and none_directly_related:")
print("SDG_count is consistent with none_directly_related:", bool(sdg_count_none_consistent))
print("Number of inconsistent rows:", len(sdg_count_none_inconsistent_rows))

if len(sdg_count_none_inconsistent_rows) > 0:
    print(
        sdg_count_none_inconsistent_rows[
            ["SDG_count", none_sdg_col] + sdg_direct_cols
        ].to_string(index=False)
    )

sdg_derived_summary = coded_data[derived_sdg_cols].describe().T
sdg_derived_summary["median"] = coded_data[derived_sdg_cols].median()
sdg_derived_summary = sdg_derived_summary[
    ["count", "mean", "std", "min", "25%", "median", "50%", "75%", "max"]
]
sdg_derived_summary_display = sdg_derived_summary.round(3)

print("\n5) Descriptive statistics for derived SDG variables:")
print(sdg_derived_summary_display.to_string())

sdg_count_distribution = (
    coded_data["SDG_count"]
    .value_counts()
    .sort_index()
    .rename_axis("SDG_count")
    .reset_index(name="number_of_firms")
)

sdg_count_distribution["percent"] = (
    sdg_count_distribution["number_of_firms"] / len(coded_data) * 100
).round(2)

print("\n6) Distribution of SDG_count:")
print(sdg_count_distribution.to_string(index=False))

sdg_high_relevance_distribution = (
    coded_data["SDG_high_relevance_count"]
    .value_counts()
    .sort_index()
    .rename_axis("SDG_high_relevance_count")
    .reset_index(name="number_of_firms")
)

sdg_high_relevance_distribution["percent"] = (
    sdg_high_relevance_distribution["number_of_firms"] / len(coded_data) * 100
).round(2)

print("\n7) Distribution of SDG_high_relevance_count:")
print(sdg_high_relevance_distribution.to_string(index=False))

n_no_direct_sdg = int((coded_data["SDG_count"] == 0).sum())
percent_no_direct_sdg = n_no_direct_sdg / len(coded_data) * 100

n_one_direct_sdg = int((coded_data["SDG_count"] == 1).sum())
percent_one_direct_sdg = n_one_direct_sdg / len(coded_data) * 100

n_zero_or_one_direct_sdg = int((coded_data["SDG_count"] <= 1).sum())
percent_zero_or_one_direct_sdg = n_zero_or_one_direct_sdg / len(coded_data) * 100

n_high_direct_sdg = int((coded_data["SDG_count"] >= 10).sum())
percent_high_direct_sdg = n_high_direct_sdg / len(coded_data) * 100

n_all_or_nearly_all_high_relevance = int(
    (coded_data["SDG_high_relevance_count"] >= 15).sum()
)
percent_all_or_nearly_all_high_relevance = (
    n_all_or_nearly_all_high_relevance / len(coded_data) * 100
)

print("\n8) Key descriptive checks:")
print(f"Number of firms with SDG_count = 0: {n_no_direct_sdg} ({percent_no_direct_sdg:.2f}%)")
print(f"Number of firms with SDG_count = 1: {n_one_direct_sdg} ({percent_one_direct_sdg:.2f}%)")
print(f"Number of firms with SDG_count <= 1: {n_zero_or_one_direct_sdg} ({percent_zero_or_one_direct_sdg:.2f}%)")
print(f"Number of firms with SDG_count >= 10: {n_high_direct_sdg} ({percent_high_direct_sdg:.2f}%)")
print(
    "Number of firms with SDG_high_relevance_count >= 15: "
    f"{n_all_or_nearly_all_high_relevance} "
    f"({percent_all_or_nearly_all_high_relevance:.2f}%)"
)

rq1_step2_valid = (
    derived_no_missing
    and sdg_count_valid
    and sdg_mean_relevance_valid
    and sdg_high_relevance_count_valid
    and sdg_count_none_consistent
)

print("\n" + "=" * 80)
print("RQ1 - STEP 2 VALIDATION SUMMARY")
print("=" * 80)
print("Derived SDG variables have no missing values:", bool(derived_no_missing))
print("SDG_count is within the range 0-17:", bool(sdg_count_valid))
print("SDG_mean_relevance is within the range 1-5:", bool(sdg_mean_relevance_valid))
print("SDG_high_relevance_count is within the range 0-17:", bool(sdg_high_relevance_count_valid))
print("SDG_count is consistent with none_directly_related:", bool(sdg_count_none_consistent))
print("\nRQ1 - Step 2 fully consistent:", bool(rq1_step2_valid))

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    coded_data.to_excel(
        writer,
        sheet_name="encoded_with_RQ1_SDG_vars",
        index=False
    )
    sdg_derived_summary_display.to_excel(
        writer,
        sheet_name="SDG_derived_summary"
    )
    sdg_count_distribution.to_excel(
        writer,
        sheet_name="SDG_count_distribution",
        index=False
    )
    sdg_high_relevance_distribution.to_excel(
        writer,
        sheet_name="high_relevance_distribution",
        index=False
    )

print("\nTable with derived RQ1 SDG variables saved to:")
print(output_path)
