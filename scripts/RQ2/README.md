# RQ2 Analysis

This folder contains the Python scripts used for Research Question 2:

**RQ2: What are the most important motivators for SDG involvement in SME operations?**

The analysis examines motivational, institutional, resource-related and value-based factors associated with SDG involvement among SMEs. The main SDG outcome is `SDG_count`, while `SDG_mean_relevance` and `SDG_high_relevance_count` are used as secondary SDG outcomes.

## Python scripts

* `RQ2_step_1_validation.py` – validates the motivational, institutional and business-context variables used for RQ2. It reconstructs SDG outcome variables from RQ1, checks binary coding, Likert-scale variables, single-select logic, logical contradictions, empty multi-select groups, and `other_selected` / `other_text` alignment.
* `RQ2_step_2.py` – constructs derived RQ2 indices related to sustainability pressure, support needs, improvement motives, external funding, sustainability-report awareness, regulatory SDG impact, and sustainability values.
* `RQ2_step_3a_frequency_rankings.py` – ranks RQ2 factors descriptively according to frequency. This step shows how often each factor was reported, but it does not yet estimate statistical association with SDG involvement.
* `RQ2_step_3b_SDG_involvement_associations.py` – analyses associations between RQ2 factors and SDG outcomes. It uses Mann–Whitney tests, Spearman correlations, Kendall correlations, Cliff’s delta, and FDR correction.
* `RQ2_step_4_nontautological_multivariate_models.py` – estimates non-tautological multivariate models for `SDG_count`. It intentionally excludes `Sustainable_success_orientation` from the main model because this variable directly reflects internal strategic alignment with sustainability.
* `RQ2_step_5_nontautological_graphs.py` – generates the final RQ2 figures aligned with the non-tautological modelling approach.

## Expected output files (see Excel_Files_Results folder)

Running the scripts sequentially produces the following main output files:

* `RQ2_step1_validation.xlsx` – validation tables for the RQ2 variables.
* `RQ2_step2_indices.xlsx` – coded dataset extended with derived RQ2 indices and validation summaries.
* `RQ2_step3a_frequency_rankings.xlsx` – descriptive frequency rankings of RQ2 factors.
* `RQ2_step3b_SDG_involvement_associations.xlsx` – bivariate association results for binary and numeric predictors across SDG outcomes.
* `RQ2_step4_nontautological_multivariate_models.xlsx` – multivariate model results, including VIF, overdispersion checks, model-fit summaries, IRR tables and OR tables.
* `RQ2_graphs_nontautological/` – folder containing the final RQ2 figures.

The scripts are intended to be run sequentially. They were developed and run in Google Colab using the coded survey dataset. Together, they produce the validation outputs, derived indices, association tables, multivariate models and figures used to answer RQ2.
