"""
RQ1 Step 3: SDG ranking and comparison analysis.

This script ranks the 17 Sustainable Development Goals according to two survey-based
measures: perceived relevance and direct business connection. It validates that the
q36 and q37 SDG columns refer to the same goals, computes descriptive rankings for
both dimensions, combines the results into one summary table, calculates rank and
measure-level Spearman correlations, identifies the largest rank gaps, and exports
the resulting tables to an Excel file.

Input file:
    /content/Survey_Results.xlsx

Expected input sheet:
    encoded_responses_corrected

Output file:
    /content/RQ1_step3_SDG_rankings.xlsx
"""

import numpy as np
import pandas as pd
from scipy.stats import spearmanr


FILE_PATH = "/content/Survey_Results.xlsx"
SHEET_NAME = "encoded_responses_corrected"
OUTPUT_PATH = "/content/RQ1_step3_SDG_rankings.xlsx"
NONE_SDG_COL = "q37_sdg_vezani_uz_poslovanje__none_directly_related"

SDG_SHORT_NAMES = {
    "sdg01_no_poverty": "SDG 1 - No poverty",
    "sdg02_zero_hunger": "SDG 2 - Zero hunger",
    "sdg03_good_health": "SDG 3 - Good health and well-being",
    "sdg04_quality_education": "SDG 4 - Quality education",
    "sdg05_gender_equality": "SDG 5 - Gender equality",
    "sdg06_clean_water": "SDG 6 - Clean water and sanitation",
    "sdg07_clean_energy": "SDG 7 - Affordable and clean energy",
    "sdg08_decent_work": "SDG 8 - Decent work and economic growth",
    "sdg09_industry_innovation": "SDG 9 - Industry, innovation and infrastructure",
    "sdg10_reduced_inequalities": "SDG 10 - Reduced inequalities",
    "sdg11_sustainable_cities": "SDG 11 - Sustainable cities and communities",
    "sdg12_responsible_consumption": "SDG 12 - Responsible consumption and production",
    "sdg13_climate_action": "SDG 13 - Climate action",
    "sdg14_life_below_water": "SDG 14 - Life below water",
    "sdg15_life_on_land": "SDG 15 - Life on land",
    "sdg16_peace_justice": "SDG 16 - Peace, justice and strong institutions",
    "sdg17_partnerships": "SDG 17 - Partnerships for the goals",
}


def extract_sdg_key(column_name):
    """Extract the SDG key from a coded q36 or q37 column name."""
    return column_name.split("__", 1)[1]


def print_table(title, table):
    """Print a table with a compact title separator."""
    print("\n" + "=" * 90)
    print(title)
    print("=" * 90)
    print(table.to_string(index=False))


def round_float_columns(table, digits=2):
    """Return a copy of a table with all float columns rounded."""
    rounded_table = table.copy()
    for column in rounded_table.select_dtypes(include=[float]).columns:
        rounded_table[column] = rounded_table[column].round(digits)
    return rounded_table


def main():
    coded_table = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)

    sdg_relevance_cols = [
        column for column in coded_table.columns
        if column.startswith("q36_relevantnost_sdg__")
    ]
    sdg_direct_cols = [
        column for column in coded_table.columns
        if column.startswith("q37_sdg_vezani_uz_poslovanje__")
        and column != NONE_SDG_COL
    ]

    print("RQ1 - STEP 3")
    print("Ranking individual SDGs")
    print("=" * 90)
    print("Input table dimensions:", coded_table.shape)
    print("Number of respondents/firms:", len(coded_table))

    if len(sdg_relevance_cols) != 17:
        raise ValueError(
            f"Expected 17 q36 SDG relevance columns, found {len(sdg_relevance_cols)}."
        )

    if len(sdg_direct_cols) != 17:
        raise ValueError(
            f"Expected 17 q37 direct SDG connection columns, found {len(sdg_direct_cols)}."
        )

    if NONE_SDG_COL not in coded_table.columns:
        raise ValueError(f"Missing expected q37 none column: {NONE_SDG_COL}")

    q36_sdg_keys = sorted(extract_sdg_key(column) for column in sdg_relevance_cols)
    q37_sdg_keys = sorted(extract_sdg_key(column) for column in sdg_direct_cols)
    sdg_keys_match = q36_sdg_keys == q37_sdg_keys

    print("\n1) SDG key validation")
    print("q36 and q37 contain the same SDG keys:", bool(sdg_keys_match))

    if not sdg_keys_match:
        print("SDG keys present in q36 but missing in q37:")
        print(sorted(set(q36_sdg_keys) - set(q37_sdg_keys)))
        print("SDG keys present in q37 but missing in q36:")
        print(sorted(set(q37_sdg_keys) - set(q36_sdg_keys)))
        raise ValueError("q36 and q37 SDG keys do not match.")

    missing_short_names = [key for key in q36_sdg_keys if key not in SDG_SHORT_NAMES]
    print("All SDG keys have short names:", len(missing_short_names) == 0)

    if missing_short_names:
        raise ValueError(f"Missing short names for SDG keys: {missing_short_names}")

    relevance_rows = []
    for column in sdg_relevance_cols:
        sdg_key = extract_sdg_key(column)
        series = coded_table[column]
        relevance_rows.append({
            "sdg_key": sdg_key,
            "SDG": SDG_SHORT_NAMES[sdg_key],
            "mean_relevance": series.mean(),
            "median_relevance": series.median(),
            "std_relevance": series.std(ddof=1),
            "min_relevance": series.min(),
            "max_relevance": series.max(),
            "high_relevance_count_4_or_5": int((series >= 4).sum()),
            "high_relevance_percent_4_or_5": (series >= 4).mean() * 100,
            "very_high_relevance_count_5": int((series == 5).sum()),
            "very_high_relevance_percent_5": (series == 5).mean() * 100,
        })

    sdg_relevance_summary = pd.DataFrame(relevance_rows).sort_values(
        by=["mean_relevance", "high_relevance_percent_4_or_5"],
        ascending=False,
    ).reset_index(drop=True)
    sdg_relevance_summary.insert(
        0,
        "rank_by_mean_relevance",
        np.arange(1, len(sdg_relevance_summary) + 1),
    )

    direct_rows = []
    for column in sdg_direct_cols:
        sdg_key = extract_sdg_key(column)
        series = coded_table[column]
        direct_rows.append({
            "sdg_key": sdg_key,
            "SDG": SDG_SHORT_NAMES[sdg_key],
            "direct_count": int(series.sum()),
            "direct_percent": series.mean() * 100,
        })

    sdg_direct_summary = pd.DataFrame(direct_rows).sort_values(
        by=["direct_count", "direct_percent"],
        ascending=False,
    ).reset_index(drop=True)
    sdg_direct_summary.insert(
        0,
        "rank_by_direct_connection",
        np.arange(1, len(sdg_direct_summary) + 1),
    )

    combined_sdg_summary = pd.merge(
        sdg_relevance_summary[[
            "rank_by_mean_relevance",
            "sdg_key",
            "SDG",
            "mean_relevance",
            "median_relevance",
            "std_relevance",
            "high_relevance_count_4_or_5",
            "high_relevance_percent_4_or_5",
            "very_high_relevance_count_5",
            "very_high_relevance_percent_5",
        ]],
        sdg_direct_summary[[
            "rank_by_direct_connection",
            "sdg_key",
            "direct_count",
            "direct_percent",
        ]],
        on="sdg_key",
        how="inner",
    )

    combined_sdg_summary["rank_difference_direct_minus_relevance"] = (
        combined_sdg_summary["rank_by_direct_connection"]
        - combined_sdg_summary["rank_by_mean_relevance"]
    )
    combined_sdg_summary["absolute_rank_difference"] = combined_sdg_summary[
        "rank_difference_direct_minus_relevance"
    ].abs()
    combined_sdg_summary = combined_sdg_summary.sort_values(
        by="rank_by_mean_relevance"
    ).reset_index(drop=True)

    combined_has_17_rows = len(combined_sdg_summary) == 17
    print("\n2) Combined table validation")
    print("Combined table contains 17 SDGs:", bool(combined_has_17_rows))

    if not combined_has_17_rows:
        raise ValueError(
            f"Combined table contains {len(combined_sdg_summary)} rows instead of 17."
        )

    rank_corr, rank_corr_p = spearmanr(
        combined_sdg_summary["rank_by_mean_relevance"],
        combined_sdg_summary["rank_by_direct_connection"],
    )
    measure_corr, measure_corr_p = spearmanr(
        combined_sdg_summary["mean_relevance"],
        combined_sdg_summary["direct_percent"],
    )

    print("\n3) Spearman correlation between relevance rank and direct-connection rank")
    print("Spearman rho:", round(rank_corr, 4))
    print("p-value:", round(rank_corr_p, 4))

    print("\n4) Spearman correlation between mean relevance and direct percentage")
    print("Spearman rho:", round(measure_corr, 4))
    print("p-value:", round(measure_corr_p, 4))

    largest_rank_gaps = combined_sdg_summary.sort_values(
        by="absolute_rank_difference",
        ascending=False,
    ).reset_index(drop=True)

    print_table(
        "5) Largest differences between relevance rank and direct-connection rank",
        round_float_columns(largest_rank_gaps)[[
            "SDG",
            "rank_by_mean_relevance",
            "rank_by_direct_connection",
            "rank_difference_direct_minus_relevance",
            "mean_relevance",
            "direct_percent",
        ]].head(10),
    )

    relevance_display = round_float_columns(sdg_relevance_summary)
    direct_display = round_float_columns(sdg_direct_summary)
    combined_display = round_float_columns(combined_sdg_summary)

    print_table(
        "SDG ranking by mean perceived relevance (q36)",
        relevance_display[[
            "rank_by_mean_relevance",
            "SDG",
            "mean_relevance",
            "median_relevance",
            "std_relevance",
            "high_relevance_count_4_or_5",
            "high_relevance_percent_4_or_5",
            "very_high_relevance_count_5",
            "very_high_relevance_percent_5",
        ]],
    )

    print_table(
        "SDG ranking by direct business connection (q37)",
        direct_display[[
            "rank_by_direct_connection",
            "SDG",
            "direct_count",
            "direct_percent",
        ]],
    )

    print_table(
        "Combined summary: perceived relevance and direct business connection",
        combined_display[[
            "SDG",
            "rank_by_mean_relevance",
            "mean_relevance",
            "high_relevance_percent_4_or_5",
            "rank_by_direct_connection",
            "direct_count",
            "direct_percent",
            "rank_difference_direct_minus_relevance",
        ]],
    )

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        sdg_relevance_summary.to_excel(writer, sheet_name="relevance_ranking", index=False)
        sdg_direct_summary.to_excel(writer, sheet_name="direct_ranking", index=False)
        combined_sdg_summary.to_excel(writer, sheet_name="combined_summary", index=False)
        largest_rank_gaps.to_excel(writer, sheet_name="largest_rank_gaps", index=False)

    rq1_step3_valid = (
        sdg_keys_match
        and combined_has_17_rows
        and len(sdg_relevance_summary) == 17
        and len(sdg_direct_summary) == 17
        and len(combined_sdg_summary) == 17
    )

    print("\n" + "=" * 90)
    print("RQ1 STEP 3 VALIDATION SUMMARY")
    print("=" * 90)
    print("q36 and q37 SDG keys match:", bool(sdg_keys_match))
    print("Relevance ranking table contains 17 rows:", len(sdg_relevance_summary) == 17)
    print("Direct-connection ranking table contains 17 rows:", len(sdg_direct_summary) == 17)
    print("Combined summary table contains 17 rows:", len(combined_sdg_summary) == 17)
    print("\nRQ1 Step 3 fully consistent:", bool(rq1_step3_valid))
    print("\nRQ1 Step 3 output file:")
    print(OUTPUT_PATH)


if __name__ == "__main__":
    main()
