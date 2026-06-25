"""
Descriptive Statistics: Sample characteristics and initial SDG involvement overview.

This script prepares descriptive statistics for the empirical analysis of SDG
involvement among SMEs. It loads the coded survey dataset, produces frequency
tables for demographic, organizational, sectoral and sustainability-related
variables, computes selected derived indicators, summarizes initial SDG_count
and SDG involvement levels, and saves all descriptive outputs to an Excel file.

Input:
    Survey_Results.xlsx, sheet "encoded_responses_corrected".

Output:
    chapter_sample_characteristics.xlsx.
"""

import numpy as np
import pandas as pd


FILE_PATH = "/content/Survey_Results.xlsx"
SHEET_NAME = "encoded_responses_corrected"
OUTPUT_PATH = "/content/chapter_sample_characteristics.xlsx"
REFERENCE_YEAR = 2025


LABELS = {
    "q01_spol__musko": "Male",
    "q01_spol__zensko": "Female",
    "q02_generacija__boomers": "Baby boomers",
    "q02_generacija__gen_x": "Gen X",
    "q02_generacija__gen_y": "Gen Y",
    "q02_generacija__gen_alpha": "Gen Alpha",
    "q02_generacija__other_selected": "Other",
    "q03_pravni_oblik__doo_jdoo": "d.o.o./j.d.o.o.",
    "q03_pravni_oblik__javno_poduzece": "Public company",
    "q03_pravni_oblik__other_selected": "Other legal form",
    "q05_organizacijska_struktura__hrvatsko_vlasnistvo": "Croatian ownership",
    "q05_organizacijska_struktura__uglavnom_hrvatsko": "Mostly Croatian ownership",
    "q05_organizacijska_struktura__mijesano_50_50": "Mixed Croatian/foreign ownership",
    "q05_organizacijska_struktura__uglavnom_strano": "Mostly foreign ownership",
    "q05_organizacijska_struktura__strano_vlasnistvo": "Foreign ownership",
    "q08_velicina_subjekta__mikro": "Micro",
    "q08_velicina_subjekta__mali": "Small",
    "q08_velicina_subjekta__srednji": "Medium",
    "q08_velicina_subjekta__veliki": "Large",
    "q09_sektor__administrativne_i_usluzne_djelatnosti": "Administrative and support service activities",
    "q09_sektor__rudarstvo": "Mining and quarrying",
    "q09_sektor__ostale_usluge": "Other service activities",
    "q09_sektor__gradjevina": "Construction",
    "q09_sektor__proizvodnja": "Manufacturing",
    "q09_sektor__zdravlje_i_socijalni_rad": "Human health and social work",
    "q09_sektor__kucna_radinost_ili_obiteljsko_poljoprivredno": "Cottage industry / family agricultural business",
    "q09_sektor__informacije_i_komunikacije": "Information and communication",
    "q09_sektor__nekretnine": "Real estate activities",
    "q09_sektor__trgovina": "Wholesale and retail trade",
    "q09_sektor__osiguranje": "Insurance",
    "q09_sektor__umjetnost_zabava_slobodno_vrijeme": "Arts, entertainment and recreation",
    "q09_sektor__poljoprivreda_sumarstvo_ribarstvo": "Agriculture, forestry and fishing",
    "q09_sektor__obrazovanje": "Education",
    "q09_sektor__financijske_usluge": "Financial services",
    "q09_sektor__profesionalne_tehnicke_usluge": "Professional and technical services",
    "q09_sektor__usluge_smjestaja": "Accommodation services",
    "q09_sektor__dostava_skladistenje": "Transportation and storage",
    "q09_sektor__extraterritorial_organization": "Extraterritorial organization",
    "q09_sektor__energenti": "Energy",
    "q09_sektor__upravljanje_vodama": "Water management",
    "q09_sektor__other_selected": "Other sector",
    "q10_mjerljivi_indikatori__da": "Yes",
    "q10_mjerljivi_indikatori__ne": "No",
    "q14_odrzivost_u_strategiji__da": "Yes",
    "q14_odrzivost_u_strategiji__ne": "No",
    "q19_elementi_odrzivosti__iso_9001": "ISO 9001",
    "q19_elementi_odrzivosti__iso_14001": "ISO 14001",
    "q19_elementi_odrzivosti__ciljevi_drustvene_odgovornosti": "CSR goals / activities",
    "q19_elementi_odrzivosti__nista_od_navedenog": "None of the listed",
    "q24_mjerenje_ciljeva_odrzivosti__da": "Yes",
    "q24_mjerenje_ciljeva_odrzivosti__ne": "No",
    "q26_najvisa_pozicija_za_odrzivost__top_menadzment": "Top management",
    "q26_najvisa_pozicija_za_odrzivost__srednji_menadzment": "Middle management",
    "q26_najvisa_pozicija_za_odrzivost__odjel_menadzmenta": "Management department",
    "q26_najvisa_pozicija_za_odrzivost__menadzer": "Manager",
    "q26_najvisa_pozicija_za_odrzivost__vlasnik": "Owner",
    "q26_najvisa_pozicija_za_odrzivost__subordinate": "Subordinate",
    "q26_najvisa_pozicija_za_odrzivost__nema_pozicije": "No such position",
    "q35_culi_za_un_sdg__da": "Yes",
    "q35_culi_za_un_sdg__ne": "No",
}


def make_frequency_table(data, columns, labels=None, total_n=None):
    """Create a frequency table for binary one-hot encoded variables."""
    if total_n is None:
        total_n = len(data)

    rows = []
    for column in columns:
        if column in data.columns:
            selected_n = int(data[column].sum())
            rows.append({
                "variable": column,
                "category": labels.get(column, column) if labels else column,
                "n": selected_n,
                "percent": round(selected_n / total_n * 100, 2),
            })

    return (
        pd.DataFrame(rows)
        .sort_values(["n", "category"], ascending=[False, True])
        .reset_index(drop=True)
    )


def print_table(title, table):
    """Print a formatted table title and table contents."""
    print("\n" + "=" * 100)
    print(title)
    print("=" * 100)
    print(table.to_string(index=False))


def get_columns(data, prefix, exclude_other_text=False):
    """Return columns that start with a given prefix."""
    columns = [column for column in data.columns if column.startswith(prefix)]
    if exclude_other_text:
        columns = [column for column in columns if not column.endswith("__other_text")]
    return columns


def make_numeric_summary(data, columns):
    """Create a descriptive summary for numeric columns."""
    summary = data[columns].describe().T
    summary["median"] = data[columns].median()
    return summary[["count", "mean", "std", "min", "25%", "median", "50%", "75%", "max"]].round(3)


def main():
    data = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)

    print("DESCRIPTIVE STATISTICS")
    print("Sample characteristics and initial SDG involvement overview")
    print("=" * 100)
    print("Dataset dimensions:", data.shape)
    print("Number of analysed firms/respondents:", len(data))

    n_firms = len(data)
    n_columns = data.shape[1]

    sample_overview = pd.DataFrame([
        {"measure": "Number of analysed firms/respondents", "value": n_firms},
        {"measure": "Number of encoded variables in the dataset", "value": n_columns},
    ])

    print_table("1. Number of analysed firms/respondents", sample_overview)

    gender_cols = get_columns(data, "q01_spol__")
    generation_cols = get_columns(data, "q02_generacija__", exclude_other_text=True)
    gender_table = make_frequency_table(data, gender_cols, LABELS, n_firms)
    generation_table = make_frequency_table(data, generation_cols, LABELS, n_firms)

    print_table("2.1 Gender of owner/manager", gender_table)
    print_table("2.2 Generation of owner/manager", generation_table)

    legal_form_cols = get_columns(data, "q03_pravni_oblik__", exclude_other_text=True)
    ownership_cols = get_columns(data, "q05_organizacijska_struktura__")
    firm_size_cols = get_columns(data, "q08_velicina_subjekta__")
    sector_cols = get_columns(data, "q09_sektor__", exclude_other_text=True)

    legal_form_table = make_frequency_table(data, legal_form_cols, LABELS, n_firms)
    ownership_table = make_frequency_table(data, ownership_cols, LABELS, n_firms)
    firm_size_table = make_frequency_table(data, firm_size_cols, LABELS, n_firms)
    sector_table = make_frequency_table(data, sector_cols, LABELS, n_firms)

    numeric_structural_cols = ["q06_godina_osnutka", "q07_broj_zaposlenika"]
    numeric_structural_summary = make_numeric_summary(data, numeric_structural_cols)

    data["company_age"] = REFERENCE_YEAR - data["q06_godina_osnutka"]
    data["log_employees"] = np.log1p(data["q07_broj_zaposlenika"])
    numeric_derived_summary = make_numeric_summary(data, ["company_age", "log_employees"])

    print_table("3.1 Legal form", legal_form_table)
    print_table("3.2 Ownership structure", ownership_table)
    print_table(
        "3.3 Year of foundation and number of employees",
        numeric_structural_summary.reset_index().rename(columns={"index": "variable"}),
    )
    print_table(
        "3.4 Derived numeric characteristics: company age and log employees",
        numeric_derived_summary.reset_index().rename(columns={"index": "variable"}),
    )
    print_table("3.5 Firm size", firm_size_table)
    print_table("3.6 Sector", sector_table)

    measurable_cols = get_columns(data, "q10_mjerljivi_indikatori__")
    strategy_cols = get_columns(data, "q14_odrzivost_u_strategiji__")
    standards_cols = get_columns(data, "q19_elementi_odrzivosti__")
    measurement_cols = get_columns(data, "q24_mjerenje_ciljeva_odrzivosti__")
    position_cols = get_columns(data, "q26_najvisa_pozicija_za_odrzivost__")
    awareness_cols = get_columns(data, "q35_culi_za_un_sdg__")

    measurable_table = make_frequency_table(data, measurable_cols, LABELS, n_firms)
    strategy_table = make_frequency_table(data, strategy_cols, LABELS, n_firms)
    standards_table = make_frequency_table(data, standards_cols, LABELS, n_firms)
    measurement_table = make_frequency_table(data, measurement_cols, LABELS, n_firms)
    position_table = make_frequency_table(data, position_cols, LABELS, n_firms)
    awareness_table = make_frequency_table(data, awareness_cols, LABELS, n_firms)

    positive_standard_cols = [
        "q19_elementi_odrzivosti__iso_9001",
        "q19_elementi_odrzivosti__iso_14001",
        "q19_elementi_odrzivosti__ciljevi_drustvene_odgovornosti",
    ]

    data["Has_quality_or_sustainability_standard"] = (
        data[positive_standard_cols].sum(axis=1) > 0
    ).astype(int)
    data["Has_sustainability_position"] = (
        1 - data["q26_najvisa_pozicija_za_odrzivost__nema_pozicije"]
    )

    sustainability_key_summary = pd.DataFrame([
        {
            "characteristic": "Has measurable business indicators",
            "n_yes": int(data["q10_mjerljivi_indikatori__da"].sum()),
            "percent_yes": round(data["q10_mjerljivi_indikatori__da"].mean() * 100, 2),
        },
        {
            "characteristic": "Sustainability included in mission/vision/strategy/goals",
            "n_yes": int(data["q14_odrzivost_u_strategiji__da"].sum()),
            "percent_yes": round(data["q14_odrzivost_u_strategiji__da"].mean() * 100, 2),
        },
        {
            "characteristic": "Has at least one quality/sustainability standard or CSR element",
            "n_yes": int(data["Has_quality_or_sustainability_standard"].sum()),
            "percent_yes": round(data["Has_quality_or_sustainability_standard"].mean() * 100, 2),
        },
        {
            "characteristic": "Measures sustainability goals",
            "n_yes": int(data["q24_mjerenje_ciljeva_odrzivosti__da"].sum()),
            "percent_yes": round(data["q24_mjerenje_ciljeva_odrzivosti__da"].mean() * 100, 2),
        },
        {
            "characteristic": "Has sustainability-related position",
            "n_yes": int(data["Has_sustainability_position"].sum()),
            "percent_yes": round(data["Has_sustainability_position"].mean() * 100, 2),
        },
        {
            "characteristic": "Prior awareness of UN SDGs",
            "n_yes": int(data["q35_culi_za_un_sdg__da"].sum()),
            "percent_yes": round(data["q35_culi_za_un_sdg__da"].mean() * 100, 2),
        },
    ])

    print_table("4.1 Measurable business indicators", measurable_table)
    print_table("4.2 Sustainability in strategy", strategy_table)
    print_table("4.3 Sustainability standards / CSR elements", standards_table)
    print_table("4.4 Sustainability goal measurement", measurement_table)
    print_table("4.5 Sustainability-related position", position_table)
    print_table("4.6 Prior awareness of UN SDGs", awareness_table)
    print_table("4.7 Key sustainability-related characteristics summary", sustainability_key_summary)

    sdg_direct_cols = [
        column for column in data.columns
        if column.startswith("q37_sdg_vezani_uz_poslovanje__")
        and column != "q37_sdg_vezani_uz_poslovanje__none_directly_related"
    ]
    none_sdg_col = "q37_sdg_vezani_uz_poslovanje__none_directly_related"

    data["SDG_count"] = data[sdg_direct_cols].sum(axis=1)
    if none_sdg_col in data.columns:
        data.loc[data[none_sdg_col] == 1, "SDG_count"] = 0

    data["SDG_involvement_level"] = pd.cut(
        data["SDG_count"],
        bins=[-1, 1, 4, 17],
        labels=[
            "Low SDG involvement, 0-1",
            "Moderate SDG involvement, 2-4",
            "High SDG involvement, 5+",
        ],
    )

    sdg_count_summary = data["SDG_count"].describe().to_frame().T
    sdg_count_summary["median"] = data["SDG_count"].median()
    sdg_count_summary = sdg_count_summary[
        ["count", "mean", "std", "min", "25%", "median", "50%", "75%", "max"]
    ].round(3)

    sdg_level_table = (
        data["SDG_involvement_level"]
        .value_counts()
        .sort_index()
        .rename_axis("SDG involvement level")
        .reset_index(name="n")
    )
    sdg_level_table["percent"] = (sdg_level_table["n"] / n_firms * 100).round(2)

    print_table("5.1 Initial overview of SDG_count", sdg_count_summary.reset_index(drop=True))
    print_table("5.2 SDG involvement levels", sdg_level_table)

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        sample_overview.to_excel(writer, sheet_name="sample_overview", index=False)
        gender_table.to_excel(writer, sheet_name="gender", index=False)
        generation_table.to_excel(writer, sheet_name="generation", index=False)
        legal_form_table.to_excel(writer, sheet_name="legal_form", index=False)
        ownership_table.to_excel(writer, sheet_name="ownership", index=False)
        numeric_structural_summary.to_excel(writer, sheet_name="numeric_structural")
        numeric_derived_summary.to_excel(writer, sheet_name="numeric_derived")
        firm_size_table.to_excel(writer, sheet_name="firm_size", index=False)
        sector_table.to_excel(writer, sheet_name="sector", index=False)
        measurable_table.to_excel(writer, sheet_name="measurable_indicators", index=False)
        strategy_table.to_excel(writer, sheet_name="sustainability_strategy", index=False)
        standards_table.to_excel(writer, sheet_name="standards_CSR", index=False)
        measurement_table.to_excel(writer, sheet_name="sustainability_measure", index=False)
        position_table.to_excel(writer, sheet_name="sustainability_position", index=False)
        awareness_table.to_excel(writer, sheet_name="SDG_awareness", index=False)
        sustainability_key_summary.to_excel(writer, sheet_name="sustainability_summary", index=False)
        sdg_count_summary.to_excel(writer, sheet_name="SDG_count_summary", index=False)
        sdg_level_table.to_excel(writer, sheet_name="SDG_levels", index=False)

    print("\n" + "=" * 100)
    print("Excel report for sample characteristics was saved to:")
    print(OUTPUT_PATH)


if __name__ == "__main__":
    main()
