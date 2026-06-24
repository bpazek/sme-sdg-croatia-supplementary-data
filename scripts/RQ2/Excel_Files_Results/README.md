# RQ2 Output Files

This folder contains the Excel output files generated during the RQ2 analysis.

**RQ2: What are the most important motivators for SDG involvement in SME operations?**

The files were produced by running the RQ2 Python scripts sequentially. They include validation outputs, derived indices, descriptive rankings, bivariate association results and non-tautological multivariate model results.

## Files

* `RQ2_step1_validation.xlsx` – validation output for the variables used in RQ2. It includes checks of SDG outcome variables, binary coding, Likert-scale variables, single-select logic, logical contradictions, empty multi-select groups, and `other_selected` / `other_text` alignment.
* `RQ2_step2_indices.xlsx` – coded dataset extended with derived RQ2 indices. These include indicators related to sustainability pressure, support needs, improvement motives, external funding, sustainability-report awareness, regulatory SDG impact, and sustainability values.
* `RQ2_step3a_frequency_rankings.xlsx` – descriptive frequency rankings of RQ2 factors. This file shows how often motivational, institutional, resource-related and value-based factors were reported by SMEs.
* `RQ2_step3b_SDG_involvement_associations.xlsx` – bivariate association results between RQ2 factors and SDG outcomes. It includes results for binary and numeric predictors across `SDG_count`, `SDG_mean_relevance`, and `SDG_high_relevance_count`.
* `RQ2_step4_nontautological_multivariate_models.xlsx` – non-tautological multivariate model results for `SDG_count`. It includes VIF diagnostics, overdispersion checks, model-fit summaries, incidence rate ratios (IRR), odds ratios (OR), and information about intentionally excluded conceptually close predictors.

These output files support the RQ2 analysis by documenting each analytical stage, from validation and index construction to descriptive, bivariate and multivariate statistical results.
