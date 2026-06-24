# RQ1 Analysis

This folder contains the Python scripts and output files used for Research Question 1:

**RQ1: Which SDGs dominate the operations and sustainability efforts of Croatian SMEs?**

## Python scripts

* `RQ1_step_1.py` – validates the q36 SDG relevance variables and q37 direct SDG connection variables.
* `RQ1_step_2.py` – constructs derived SDG variables, including `SDG_count`, `SDG_mean_relevance`, and `SDG_high_relevance_count`.
* `RQ1_step_3.py` – ranks the 17 SDGs by perceived relevance, direct business connection, and combined ranking information.
* `RQ1_step_4.py` – generates the RQ1 visualizations based on the ranking results.
* `RQ1_step_5.py` – computes the final RQ1 dominance index and inferential checks.

## Output files

* `RQ1_with_SDG_derived_variables_Step_2.xlsx` – dataset with derived SDG variables and summary tables from Step 2.
* `RQ1_step3_SDG_rankings_Step_3.xlsx` – SDG relevance rankings, direct-connection rankings, combined summary, and rank-gap results from Step 3.
* `RQ1_step5_dominance_inference_Step_5.xlsx` – final SDG-level summary, correlation results, dominance ranking, rank differences, and inferential tests from Step 5.

The scripts are intended to be run sequentially. They were developed and run in Google Colab using the coded survey dataset, and they produce the tables, figures, and statistical outputs reported for RQ1 in the study.
