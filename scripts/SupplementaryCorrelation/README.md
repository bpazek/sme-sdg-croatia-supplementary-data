# Supplementary Correlation Analysis

This folder contains the Python scripts and output files used for the supplementary analysis of the relationship between external funding sources and motives for implemented improvements.

The analysis connects q34 variables, which describe external funding sources used during the previous three years, with q31 variables, which describe motives for implemented changes or improvements.

## Python scripts

* `SupplementaryCorrelation_step_1_validation.py` – validates q31 improvement-motive variables and q34 external-funding variables, including binary coding, missing values, none-like options, logical contradictions, and `other_text` alignment.
* `SupplementaryCorrelation_step_2.py` – constructs derived q31/q34 indicators, including `Improvement_motivators_count`, `External_funding_count`, `Any_external_funding`, `Any_EU_or_public_funding`, and `Any_bank_or_credit_support`.
* `SupplementaryCorrelation_step_3.py` – performs the supplementary correlation and association analysis, including Spearman correlation, Kendall’s tau-b, Mann–Whitney group comparisons, Fisher exact tests, phi coefficients, and FDR correction.
* `SupplementaryCorrelation_step_4.py` – generates the main and appendix figures and creates final summary tables for the supplementary analysis.

## Output files

* `Additional_Analysis_3_q31_q34_step_1_validation.xlsx` – validation output for q31 and q34 variables.
* `Additional_Analysis_3_q31_q34_step_2_indices.xlsx` – dataset with derived q31/q34 indices and descriptive summaries.
* `Additional_Analysis_3_q31_q34_step_3_results_with_kendall.xlsx` – correlation, group-comparison, pairwise-association, and FDR-corrected results.
* `Additional_Analysis_3_q31_q34_step_4_graph_summaries.xlsx` – final summary tables, graph data, and supporting tables used for the figures.

The scripts are intended to be run sequentially. They were developed and run in Google Colab using the coded survey dataset. Together, they produce the supplementary tables, figures, and statistical outputs used to support the analysis of external funding and improvement motives.
