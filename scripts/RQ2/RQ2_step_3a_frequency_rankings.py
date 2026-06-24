"""
RQ2 Step 3a: Descriptive frequency ranking of RQ2 factors.

This script supports Research Question 2:
"What are the most important motivators for SDG involvement in SME operations?"

The script uses the indexed dataset created in RQ2 Step 2 and ranks motivational,
institutional, resource-related and value-related factors by frequency. This is a
descriptive step only: it ranks factors according to how often they were reported,
but it does not yet estimate their statistical association with SDG_count or other
SDG outcomes.

Input:
    RQ2_step2_indices.xlsx, sheet "encoded_with_RQ2_indices".

Output:
    RQ2_step3a_frequency_rankings.xlsx.
"""

import numpy as np
import pandas as pd


INPUT_PATH = "/content/RQ2_step2_indices.xlsx"
INPUT_SHEET = "encoded_with_RQ2_indices"
OUTPUT_PATH = "/content/RQ2_step3a_frequency_rankings.xlsx"


def make_frequency_table(data, columns, construct_name, indicator_type, labels, type_map):
    rows = []

    for column in columns:
        values = sorted(data[column].dropna().unique())
        is_binary = set(values).issubset({0, 1})

        if not is_binary:
            raise ValueError(f"Variable {column} is not binary. Values: {values}")

        selected_n = int(data[column].sum())
        selected_percent = selected_n / len(data) * 100

        rows.append({
            "construct": construct_name,
            "indicator_type": type_map.get(column, indicator_type),
            "variable": column,
            "label": labels.get(column, column),
            "selected_n": selected_n,
            "selected_percent": selected_percent,
        })

    table = pd.DataFrame(rows)
    table = table.sort_values(
        ["selected_percent", "selected_n"],
        ascending=False,
    ).reset_index(drop=True)
    table.insert(0, "rank_within_construct", np.arange(1, len(table) + 1))

    return table


def make_likert_importance_table(data, labels):
    rows = []
    likert_columns = [
        "Economic_sustainability_importance",
        "Environmental_sustainability_importance",
        "Social_sustainability_importance",
    ]

    for column in likert_columns:
        series = data[column]
        values = sorted(series.dropna().unique())
        valid_likert = set(values).issubset({1, 2, 3, 4, 5})

        if not valid_likert:
            raise ValueError(f"Variable {column} is not a valid Likert 1-5 variable.")

        rows.append({
            "construct": "Sustainability value orientation",
            "indicator_type": "sustainability_value",
            "variable": column,
            "label": labels.get(column, column),
            "mean": series.mean(),
            "median": series.median(),
            "std": series.std(ddof=1),
            "high_importance_count_4_or_5": int((series >= 4).sum()),
            "high_importance_percent_4_or_5": (series >= 4).mean() * 100,
        })

    table = pd.DataFrame(rows)
    table = table.sort_values(
        ["mean", "high_importance_percent_4_or_5"],
        ascending=False,
    ).reset_index(drop=True)
    table.insert(0, "rank_within_construct", np.arange(1, len(table) + 1))

    return table


def round_percent_table(table):
    output = table.copy()
    if "selected_percent" in output.columns:
        output["selected_percent"] = output["selected_percent"].round(2)
    return output


def print_selected_table(title, table, columns, head=None):
    print(f"\n{title}")
    display_table = table[columns]
    if head is not None:
        display_table = display_table.head(head)
    print(display_table.to_string(index=False))


def main():
    data = pd.read_excel(INPUT_PATH, sheet_name=INPUT_SHEET)

    print("RQ2 - STEP 3a")
    print("Descriptive analysis and frequency ranking of RQ2 factors")
    print("=" * 100)
    print("Loaded table from Step 2:")
    print(INPUT_PATH)
    print("Dataset dimensions:", data.shape)
    print("Number of respondents / firms:", len(data))

    q15_cols = [column for column in data.columns if column.startswith("q15_pritisak_na_odrzivost__")]
    q27_cols = [
        column for column in data.columns
        if column.startswith("q27_pomoc_za_odrzivost__") and not column.endswith("__other_text")
    ]
    q31_cols = [
        column for column in data.columns
        if column.startswith("q31_motivi_promjena__") and not column.endswith("__other_text")
    ]
    q34_cols = [
        column for column in data.columns
        if column.startswith("q34_eksterna_sredstva_3_godine__") and not column.endswith("__other_text")
    ]
    q16_cols = [column for column in data.columns if column.startswith("q16_mjerenje_uspjeha__")]
    q17_cols = [column for column in data.columns if column.startswith("q17_izvjestaj_o_odrzivosti__")]
    q18_cols = [column for column in data.columns if column.startswith("q18_zakonodavni_utjecaj_sdg__")]

    sustainability_value_cols = [
        "Economic_sustainability_importance",
        "Environmental_sustainability_importance",
        "Social_sustainability_importance",
        "Sustainability_importance_mean",
        "Sustainability_importance_high_count",
    ]

    q31_no_improvement_col = "q31_motivi_promjena__nista_poboljsano"
    q34_no_external_funding_col = "q34_eksterna_sredstva_3_godine__bez_eksternih_sredstava"
    q17_no_report_col = "q17_izvjestaj_o_odrzivosti__ne"

    q31_positive_motivation_cols = [column for column in q31_cols if column != q31_no_improvement_col]
    q34_external_funding_source_cols = [column for column in q34_cols if column != q34_no_external_funding_col]
    q17_positive_awareness_cols = [column for column in q17_cols if column != q17_no_report_col]

    group_size_summary = pd.DataFrame([
        {"group": "q15_pritisak_na_odrzivost", "n_columns": len(q15_cols)},
        {"group": "q27_pomoc_za_odrzivost", "n_columns": len(q27_cols)},
        {"group": "q31_motivi_promjena", "n_columns": len(q31_cols)},
        {"group": "q34_eksterna_sredstva_3_godine", "n_columns": len(q34_cols)},
        {"group": "q16_mjerenje_uspjeha", "n_columns": len(q16_cols)},
        {"group": "q17_izvjestaj_o_odrzivosti", "n_columns": len(q17_cols)},
        {"group": "q18_zakonodavni_utjecaj_sdg", "n_columns": len(q18_cols)},
        {"group": "sustainability_values", "n_columns": len(sustainability_value_cols)},
    ])

    print("\n1) Number of columns by RQ2 group:")
    print(group_size_summary.to_string(index=False))

    required_columns = (
        q15_cols + q27_cols + q31_cols + q34_cols + q16_cols + q17_cols
        + q18_cols + sustainability_value_cols
    )
    missing_columns = [column for column in required_columns if column not in data.columns]

    if missing_columns:
        raise ValueError(f"Missing expected columns: {missing_columns}")

    empty_groups = group_size_summary[group_size_summary["n_columns"] == 0]
    if len(empty_groups) > 0:
        raise ValueError("Some groups have no detected columns:\n" + empty_groups.to_string(index=False))

    variable_labels = {
        "q15_pritisak_na_odrzivost__pravni": "Legal pressure",
        "q15_pritisak_na_odrzivost__potrosaci": "Consumers",
        "q15_pritisak_na_odrzivost__konkurencija": "Competition",
        "q15_pritisak_na_odrzivost__okruzenje": "Business environment",
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
        "q31_motivi_promjena__zbog_konkurencije": "Competition",
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
        "Economic_sustainability_importance": "Economic sustainability importance",
        "Environmental_sustainability_importance": "Environmental sustainability importance",
        "Social_sustainability_importance": "Social sustainability importance",
        "Sustainability_importance_mean": "Mean sustainability importance",
        "Sustainability_importance_high_count": "High sustainability-importance count",
    }

    indicator_type_map = {}
    for column in q15_cols:
        indicator_type_map[column] = "pressure"
    for column in q27_cols:
        indicator_type_map[column] = "support_need"
    for column in q31_positive_motivation_cols:
        indicator_type_map[column] = "improvement_motive"
    indicator_type_map[q31_no_improvement_col] = "absence_negative"
    for column in q34_external_funding_source_cols:
        indicator_type_map[column] = "external_funding_resource"
    indicator_type_map[q34_no_external_funding_col] = "absence_negative"
    indicator_type_map["q16_mjerenje_uspjeha__profit_i_odrzivi_ciljevi"] = "success_orientation"
    indicator_type_map["q16_mjerenje_uspjeha__neto_profit"] = "absence_negative"
    for column in q17_positive_awareness_cols:
        indicator_type_map[column] = "report_awareness"
    indicator_type_map[q17_no_report_col] = "absence_negative"
    indicator_type_map["q18_zakonodavni_utjecaj_sdg__da"] = "regulatory_perception"
    indicator_type_map["q18_zakonodavni_utjecaj_sdg__ne"] = "absence_negative"
    for column in sustainability_value_cols:
        indicator_type_map[column] = "sustainability_value"

    rq2_pressure_frequency = make_frequency_table(data, q15_cols, "Sustainability pressure", "pressure", variable_labels, indicator_type_map)
    rq2_support_frequency = make_frequency_table(data, q27_cols, "Support needed for sustainability", "support_need", variable_labels, indicator_type_map)
    rq2_improvement_motives_frequency = make_frequency_table(data, q31_positive_motivation_cols, "Positive improvement motives", "improvement_motive", variable_labels, indicator_type_map)
    rq2_no_improvement_frequency = make_frequency_table(data, [q31_no_improvement_col], "Absence of improvement", "absence_negative", variable_labels, indicator_type_map)
    rq2_external_funding_frequency = make_frequency_table(data, q34_external_funding_source_cols, "External funding resources", "external_funding_resource", variable_labels, indicator_type_map)
    rq2_no_external_funding_frequency = make_frequency_table(data, [q34_no_external_funding_col], "Absence of external funding", "absence_negative", variable_labels, indicator_type_map)
    rq2_success_orientation_frequency = make_frequency_table(data, q16_cols, "Success measurement orientation", "success_orientation", variable_labels, indicator_type_map)
    rq2_report_awareness_frequency = make_frequency_table(data, q17_positive_awareness_cols, "Sustainability report awareness", "report_awareness", variable_labels, indicator_type_map)
    rq2_no_report_awareness_frequency = make_frequency_table(data, [q17_no_report_col], "Absence of report awareness", "absence_negative", variable_labels, indicator_type_map)
    rq2_regulatory_impact_frequency = make_frequency_table(data, q18_cols, "Perceived regulatory SDG impact", "regulatory_perception", variable_labels, indicator_type_map)
    rq2_sustainability_values = make_likert_importance_table(data, variable_labels)

    frequency_tables = [
        rq2_pressure_frequency,
        rq2_support_frequency,
        rq2_improvement_motives_frequency,
        rq2_no_improvement_frequency,
        rq2_external_funding_frequency,
        rq2_no_external_funding_frequency,
        rq2_success_orientation_frequency,
        rq2_report_awareness_frequency,
        rq2_no_report_awareness_frequency,
        rq2_regulatory_impact_frequency,
    ]

    rq2_all_frequency = pd.concat(frequency_tables, ignore_index=True)
    rq2_all_frequency["selected_percent"] = rq2_all_frequency["selected_percent"].round(2)

    positive_indicator_mask = rq2_all_frequency["indicator_type"] != "absence_negative"
    rq2_positive_indicator_frequency = (
        rq2_all_frequency[positive_indicator_mask]
        .copy()
        .sort_values(["selected_percent", "selected_n"], ascending=False)
        .reset_index(drop=True)
    )
    rq2_positive_indicator_frequency.insert(
        0,
        "overall_frequency_rank",
        np.arange(1, len(rq2_positive_indicator_frequency) + 1),
    )

    rq2_negative_or_absence_frequency = (
        rq2_all_frequency[rq2_all_frequency["indicator_type"] == "absence_negative"]
        .copy()
        .sort_values(["selected_percent", "selected_n"], ascending=False)
        .reset_index(drop=True)
    )
    rq2_negative_or_absence_frequency.insert(
        0,
        "overall_frequency_rank",
        np.arange(1, len(rq2_negative_or_absence_frequency) + 1),
    )

    rq2_top_positive_10 = rq2_positive_indicator_frequency.head(10).copy()
    rq2_top_pressure = rq2_pressure_frequency.copy()
    rq2_top_support = rq2_support_frequency.copy()
    rq2_top_improvement_motives = rq2_improvement_motives_frequency.head(10).copy()
    rq2_top_external_funding = rq2_external_funding_frequency.head(10).copy()

    rq2_pressure_display = round_percent_table(rq2_pressure_frequency)
    rq2_support_display = round_percent_table(rq2_support_frequency)
    rq2_improvement_display = round_percent_table(rq2_improvement_motives_frequency)
    rq2_external_funding_display = round_percent_table(rq2_external_funding_frequency)
    rq2_success_display = round_percent_table(rq2_success_orientation_frequency)
    rq2_report_display = round_percent_table(rq2_report_awareness_frequency)
    rq2_regulatory_display = round_percent_table(rq2_regulatory_impact_frequency)
    rq2_positive_display = round_percent_table(rq2_positive_indicator_frequency)
    rq2_negative_display = round_percent_table(rq2_negative_or_absence_frequency)

    rq2_sustainability_values_display = rq2_sustainability_values.copy()
    for column in ["mean", "median", "std", "high_importance_percent_4_or_5"]:
        rq2_sustainability_values_display[column] = rq2_sustainability_values_display[column].round(2)

    print("\n" + "=" * 100)
    print("RQ2 - STEP 3a: DESCRIPTIVE ANALYSIS AND FREQUENCY RANKING")
    print("=" * 100)

    print_selected_table("2) Sustainability pressures, q15:", rq2_pressure_display, ["rank_within_construct", "label", "selected_n", "selected_percent"])
    print_selected_table("3) Support needed for sustainability, q27:", rq2_support_display, ["rank_within_construct", "label", "selected_n", "selected_percent"])
    print_selected_table("4) Positive improvement motives, q31:", rq2_improvement_display, ["rank_within_construct", "label", "selected_n", "selected_percent"])
    print_selected_table("5) External funding sources, q34, excluding the no-external-funding option:", rq2_external_funding_display, ["rank_within_construct", "label", "selected_n", "selected_percent"])
    print_selected_table("6) Success measurement orientation, q16:", rq2_success_display, ["rank_within_construct", "label", "selected_n", "selected_percent"])
    print_selected_table("7) Sustainability report awareness, q17, excluding the no-awareness option:", rq2_report_display, ["rank_within_construct", "label", "selected_n", "selected_percent"])
    print_selected_table("8) Perceived regulatory SDG impact, q18:", rq2_regulatory_display, ["rank_within_construct", "label", "selected_n", "selected_percent"])

    print_selected_table(
        "9) Value importance of sustainability dimensions, q20-q22:",
        rq2_sustainability_values_display,
        [
            "rank_within_construct",
            "label",
            "mean",
            "median",
            "high_importance_count_4_or_5",
            "high_importance_percent_4_or_5",
        ],
    )

    print_selected_table(
        "10) Overall descriptive rank of positive or active indicators by frequency:",
        rq2_positive_display,
        ["overall_frequency_rank", "construct", "indicator_type", "label", "selected_n", "selected_percent"],
        head=20,
    )

    print_selected_table(
        "11) Absence or negative-orientation indicators:",
        rq2_negative_display,
        ["overall_frequency_rank", "construct", "indicator_type", "label", "selected_n", "selected_percent"],
    )

    print("\n12) Automatic summary of most frequent factors:")
    top_pressure = rq2_pressure_display.iloc[0]
    top_support = rq2_support_display.iloc[0]
    top_improvement = rq2_improvement_display.iloc[0]
    top_external = rq2_external_funding_display.iloc[0]
    top_success = rq2_success_display.iloc[0]
    top_regulatory_yes = rq2_regulatory_display[
        rq2_regulatory_display["variable"] == "q18_zakonodavni_utjecaj_sdg__da"
    ].iloc[0]
    top_value = rq2_sustainability_values_display.iloc[0]

    print(
        f"The most frequent source of sustainability pressure was '{top_pressure['label']}' "
        f"({int(top_pressure['selected_n'])} SMEs; {top_pressure['selected_percent']:.2f}%)."
    )
    print(
        f"The most frequently requested form of support was '{top_support['label']}' "
        f"({int(top_support['selected_n'])} SMEs; {top_support['selected_percent']:.2f}%)."
    )
    print(
        f"The most frequent positive improvement motive was '{top_improvement['label']}' "
        f"({int(top_improvement['selected_n'])} SMEs; {top_improvement['selected_percent']:.2f}%)."
    )
    print(
        f"The most frequent concrete external funding source was '{top_external['label']}' "
        f"({int(top_external['selected_n'])} SMEs; {top_external['selected_percent']:.2f}%)."
    )
    print(
        f"The most frequent success-measurement orientation was '{top_success['label']}' "
        f"({int(top_success['selected_n'])} SMEs; {top_success['selected_percent']:.2f}%)."
    )
    print(
        f"The share of firms that believed legislative SDG inclusion would affect results "
        f"was {top_regulatory_yes['selected_percent']:.2f}%."
    )
    print(
        f"The highest-rated sustainability dimension was '{top_value['label']}' "
        f"(mean = {top_value['mean']:.2f}; "
        f"high importance = {top_value['high_importance_percent_4_or_5']:.2f}%)."
    )
    print(
        "\nNote: This step ranks factors by frequency. The final RQ2 importance ranking is "
        "estimated in the next step by relating these factors to SDG_count and other SDG outcomes."
    )

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        group_size_summary.to_excel(writer, sheet_name="group_sizes", index=False)
        rq2_pressure_frequency.to_excel(writer, sheet_name="pressure_q15", index=False)
        rq2_support_frequency.to_excel(writer, sheet_name="support_q27", index=False)
        rq2_improvement_motives_frequency.to_excel(writer, sheet_name="improvement_q31", index=False)
        rq2_external_funding_frequency.to_excel(writer, sheet_name="external_funding_q34", index=False)
        rq2_success_orientation_frequency.to_excel(writer, sheet_name="success_orientation_q16", index=False)
        rq2_report_awareness_frequency.to_excel(writer, sheet_name="report_awareness_q17", index=False)
        rq2_regulatory_impact_frequency.to_excel(writer, sheet_name="regulatory_q18", index=False)
        rq2_sustainability_values.to_excel(writer, sheet_name="sustainability_values", index=False)
        rq2_all_frequency.to_excel(writer, sheet_name="all_frequency", index=False)
        rq2_positive_indicator_frequency.to_excel(writer, sheet_name="positive_indicators", index=False)
        rq2_negative_or_absence_frequency.to_excel(writer, sheet_name="negative_absence", index=False)
        rq2_top_positive_10.to_excel(writer, sheet_name="top_positive_10", index=False)
        rq2_top_pressure.to_excel(writer, sheet_name="top_pressure", index=False)
        rq2_top_support.to_excel(writer, sheet_name="top_support", index=False)
        rq2_top_improvement_motives.to_excel(writer, sheet_name="top_improvement", index=False)
        rq2_top_external_funding.to_excel(writer, sheet_name="top_external_funding", index=False)

    print("\nRQ2 Step 3a results saved to:")
    print(OUTPUT_PATH)
    print("\nRQ2 Step 3a completed.")


if __name__ == "__main__":
    main()
