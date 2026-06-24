"""
RQ2 Step 4: Non-tautological multivariate modelling of SDG involvement.

This script estimates multivariate models for the primary RQ2 outcome,
SDG_count, using non-tautological predictors derived in the previous RQ2 steps.
The model intentionally excludes q16 / Sustainable_success_orientation, because
that variable directly reflects whether a firm already includes sustainable goals
in its measure of business success. It is therefore interpreted as internal
strategic alignment with sustainability rather than as a narrower external or
operational motivator.

The main model includes the number of sustainability pressure sources, number of
external funding sources, number of positive improvement motives, and
environmental sustainability importance. Sensitivity analyses include a model
using mean sustainability importance, a logistic model for SDG_count >= 2, and an
exploratory model with individual pressure sources.

Input files:
    RQ2_step2_indices.xlsx, sheet "encoded_with_RQ2_indices".
    RQ2_step3b_SDG_involvement_associations.xlsx.

Output file:
    RQ2_step4_nontautological_multivariate_models.xlsx.
"""

import pandas as pd
import numpy as np
import statsmodels.api as sm
import statsmodels.formula.api as smf

from pathlib import Path
from statsmodels.stats.outliers_influence import variance_inflation_factor

print("RQ2 - STEP 4")
print("Non-tautological multivariate modelling of SDG involvement")
print("=" * 100)


step2_path = "/content/RQ2_step2_indices.xlsx"
step3b_path = "/content/RQ2_step3b_SDG_involvement_associations.xlsx"

kodirana_tablica = pd.read_excel(
    step2_path,
    sheet_name="encoded_with_RQ2_indices"
)

binary_sdg_count_ranked = pd.read_excel(
    step3b_path,
    sheet_name="binary_SDG_count_reliable"
)

numeric_sdg_count_ranked = pd.read_excel(
    step3b_path,
    sheet_name="numeric_SDG_count_ranked"
)

print("Loaded table from Step 2:")
print(step2_path)
print("Dimensions:", kodirana_tablica.shape)

print("\nLoaded Step 3b results:")
print(step3b_path)
print("binary_SDG_count_reliable:", binary_sdg_count_ranked.shape)
print("numeric_SDG_count_ranked:", numeric_sdg_count_ranked.shape)


outcome = "SDG_count"

if outcome not in kodirana_tablica.columns:
    raise ValueError(f"Missing outcome variable: {outcome}")


excluded_tautological_predictors = [
    {
        "variable": "Sustainable_success_orientation",
        "reason": (
            "Excluded from the main RQ2 model because it directly reflects "
            "whether the firm already includes sustainable goals in measuring "
            "business success. It is better interpreted as internal strategic "
            "alignment with sustainability, not as an external or operational motivator."
        )
    },
    {
        "variable": "q16_mjerenje_uspjeha__profit_i_odrzivi_ciljevi",
        "reason": (
            "Original q16 one-hot indicator corresponding to Sustainable_success_orientation; "
            "excluded for the same conceptual reason."
        )
    }
]

excluded_tautological_predictors_df = pd.DataFrame(excluded_tautological_predictors)

print("\nIntentionally excluded conceptually close / tautological variables:")
print(excluded_tautological_predictors_df.to_string(index=False))



main_predictors_raw = [
    "Pressure_count",
    "External_funding_count",
    "Improvement_motivators_count",
    "Environmental_sustainability_importance"
]

sensitivity_predictors_raw = [
    "Pressure_count",
    "External_funding_count",
    "Improvement_motivators_count",
    "Sustainability_importance_mean"
]

pressure_source_predictors_raw = [
    "q15_pritisak_na_odrzivost__banka",
    "q15_pritisak_na_odrzivost__konkurencija",
    "q15_pritisak_na_odrzivost__potrosaci",
    "q15_pritisak_na_odrzivost__zaposlenici",
    "External_funding_count",
    "Improvement_motivators_count",
    "Environmental_sustainability_importance"
]

logit_predictors_raw = main_predictors_raw.copy()

all_required_cols_raw = (
    [outcome]
    + main_predictors_raw
    + sensitivity_predictors_raw
    + pressure_source_predictors_raw
    + logit_predictors_raw
)

all_required_cols = list(dict.fromkeys(all_required_cols_raw))

missing_cols = [
    col for col in all_required_cols
    if col not in kodirana_tablica.columns
]

if missing_cols:
    raise ValueError(f"Missing required columns: {missing_cols}")

model_df = kodirana_tablica[all_required_cols].copy()

duplicate_columns = model_df.columns[model_df.columns.duplicated()].tolist()

if duplicate_columns:
    raise ValueError(f"model_df has duplicated columns: {duplicate_columns}")

print("\nColumns included in model_df:")
for col in model_df.columns:
    print(" -", col)


print("\nMissing values in model variables:")
print(model_df.isna().sum().to_string())

if model_df.isna().sum().sum() > 0:
    print("\nMissing values are present. Rows with missing values will be removed.")
    model_df = model_df.dropna().copy()

print("\nNumber of rows in model_df:", len(model_df))

print("\nDescriptive statistics for SDG_count:")
print(model_df[outcome].describe().to_string())

sdg_count_mean = model_df[outcome].mean()
sdg_count_var = model_df[outcome].var(ddof=1)

print("\nMean of SDG_count:", round(sdg_count_mean, 4))
print("Variance of SDG_count:", round(sdg_count_var, 4))

if sdg_count_mean > 0:
    print("Variance-to-mean ratio:", round(sdg_count_var / sdg_count_mean, 4))


continuous_predictors_to_standardize = [
    "Pressure_count",
    "External_funding_count",
    "Improvement_motivators_count",
    "Environmental_sustainability_importance",
    "Sustainability_importance_mean"
]

def add_z_scores(df, cols):
    df = df.copy()
    z_cols = []

    for col in cols:
        if col not in df.columns:
            continue

        if isinstance(df[col], pd.DataFrame):
            raise ValueError(f"Column {col} is duplicated.")

        std = df[col].std(ddof=1)

        if pd.isna(std) or std == 0:
            raise ValueError(f"Variable {col} has no variability.")

        z_col = col + "_z"
        df[z_col] = (df[col] - df[col].mean()) / std
        z_cols.append(z_col)

    return df, z_cols

model_df, z_cols_created = add_z_scores(
    model_df,
    continuous_predictors_to_standardize
)

print("\nStandardized predictors:")
for col in z_cols_created:
    print(" -", col)


main_predictors = [
    "Pressure_count_z",
    "External_funding_count_z",
    "Improvement_motivators_count_z",
    "Environmental_sustainability_importance_z"
]

sensitivity_predictors = [
    "Pressure_count_z",
    "External_funding_count_z",
    "Improvement_motivators_count_z",
    "Sustainability_importance_mean_z"
]

pressure_source_predictors = [
    "q15_pritisak_na_odrzivost__banka",
    "q15_pritisak_na_odrzivost__konkurencija",
    "q15_pritisak_na_odrzivost__potrosaci",
    "q15_pritisak_na_odrzivost__zaposlenici",
    "External_funding_count_z",
    "Improvement_motivators_count_z",
    "Environmental_sustainability_importance_z"
]

logit_predictors = main_predictors.copy()

print("\nMain non-tautological model:")
for col in main_predictors:
    print(" -", col)

print("\nSensitivity model:")
for col in sensitivity_predictors:
    print(" -", col)

print("\nExploratory model with individual pressure sources:")
for col in pressure_source_predictors:
    print(" -", col)


term_labels = {
    "const": "Intercept",
    "Intercept": "Intercept",
    "Pressure_count_z": "Number of sustainability pressure sources, z-score",
    "External_funding_count_z": "Number of external funding sources, z-score",
    "Improvement_motivators_count_z": "Number of positive improvement motives, z-score",
    "Environmental_sustainability_importance_z": "Environmental sustainability importance, z-score",
    "Sustainability_importance_mean_z": "Mean sustainability importance, z-score",
    "q15_pritisak_na_odrzivost__banka": "Bank pressure",
    "q15_pritisak_na_odrzivost__konkurencija": "Competition pressure",
    "q15_pritisak_na_odrzivost__potrosaci": "Consumer pressure",
    "q15_pritisak_na_odrzivost__zaposlenici": "Employee pressure",
    "alpha": "Negative binomial alpha"
}


def compute_vif(df, predictors):
    X = df[predictors].copy()
    X = sm.add_constant(X, has_constant="add")

    rows = []

    for i, col in enumerate(X.columns):
        if col == "const":
            continue

        rows.append({
            "variable": col,
            "label": term_labels.get(col, col),
            "VIF": variance_inflation_factor(X.values, i)
        })

    return pd.DataFrame(rows)

vif_main = compute_vif(model_df, main_predictors)
vif_sensitivity = compute_vif(model_df, sensitivity_predictors)
vif_pressure_sources = compute_vif(model_df, pressure_source_predictors)

print("\nVIF - main model:")
print(vif_main.round(3).to_string(index=False))

print("\nVIF - sensitivity model:")
print(vif_sensitivity.round(3).to_string(index=False))

print("\nVIF - individual pressure-source model:")
print(vif_pressure_sources.round(3).to_string(index=False))


def fit_poisson_glm(df, outcome, predictors, robust=False):
    formula = outcome + " ~ " + " + ".join(predictors)

    if robust:
        return smf.glm(
            formula=formula,
            data=df,
            family=sm.families.Poisson()
        ).fit(cov_type="HC3")

    return smf.glm(
        formula=formula,
        data=df,
        family=sm.families.Poisson()
    ).fit()


def fit_negative_binomial(df, outcome, predictors):
    X = df[predictors].copy()
    X = sm.add_constant(X, has_constant="add")
    y = df[outcome]

    model = sm.NegativeBinomial(y, X)
    result = model.fit(disp=False)

    return result


def poisson_overdispersion(model):
    return {
        "pearson_chi2": model.pearson_chi2,
        "df_resid": model.df_resid,
        "dispersion_ratio": model.pearson_chi2 / model.df_resid
    }


def count_model_irr_table(model, model_name, model_type):
    params = model.params
    conf = model.conf_int()
    pvalues = model.pvalues
    bse = model.bse

    rows = []

    for term in params.index:
        if term == "alpha":
            rows.append({
                "model": model_name,
                "model_type": model_type,
                "term": term,
                "label": term_labels.get(term, term),
                "coef": params[term],
                "std_error": bse[term],
                "p_value": pvalues[term],
                "IRR": np.nan,
                "IRR_CI_low": np.nan,
                "IRR_CI_high": np.nan
            })
        else:
            rows.append({
                "model": model_name,
                "model_type": model_type,
                "term": term,
                "label": term_labels.get(term, term),
                "coef": params[term],
                "std_error": bse[term],
                "p_value": pvalues[term],
                "IRR": np.exp(params[term]),
                "IRR_CI_low": np.exp(conf.loc[term, 0]),
                "IRR_CI_high": np.exp(conf.loc[term, 1])
            })

    return pd.DataFrame(rows)


def logit_or_table(model, model_name):
    params = model.params
    conf = model.conf_int()
    pvalues = model.pvalues
    bse = model.bse

    rows = []

    for term in params.index:
        rows.append({
            "model": model_name,
            "term": term,
            "label": term_labels.get(term, term),
            "coef": params[term],
            "std_error": bse[term],
            "p_value": pvalues[term],
            "OR": np.exp(params[term]),
            "OR_CI_low": np.exp(conf.loc[term, 0]),
            "OR_CI_high": np.exp(conf.loc[term, 1])
        })

    return pd.DataFrame(rows)


def mcfadden_pseudo_r2(model_full, model_null):
    return 1 - (model_full.llf / model_null.llf)


def fit_summary_glm(model, model_name, null_model=None):
    out = {
        "model": model_name,
        "n": int(model.nobs),
        "df_model": model.df_model,
        "df_resid": model.df_resid,
        "log_likelihood": model.llf,
        "AIC": model.aic,
        "BIC": getattr(model, "bic", np.nan)
    }

    if null_model is not None:
        out["McFadden_pseudo_R2"] = mcfadden_pseudo_r2(model, null_model)
    else:
        out["McFadden_pseudo_R2"] = np.nan

    return out


def fit_summary_discrete(model, model_name, null_model=None):
    out = {
        "model": model_name,
        "n": int(model.nobs),
        "df_model": model.df_model,
        "df_resid": model.df_resid,
        "log_likelihood": model.llf,
        "AIC": model.aic,
        "BIC": model.bic
    }

    if null_model is not None:
        out["McFadden_pseudo_R2"] = mcfadden_pseudo_r2(model, null_model)
    else:
        out["McFadden_pseudo_R2"] = np.nan

    return out


def round_numeric(df, decimals=4):
    out = df.copy()
    for col in out.select_dtypes(include=[np.number]).columns:
        out[col] = out[col].round(decimals)
    return out


poisson_null = smf.glm(
    formula=outcome + " ~ 1",
    data=model_df,
    family=sm.families.Poisson()
).fit()

nb_null = fit_negative_binomial(
    model_df,
    outcome=outcome,
    predictors=[]
)


print("\n" + "=" * 100)
print("MODEL A: Main non-tautological model")
print("=" * 100)

poisson_main = fit_poisson_glm(
    model_df,
    outcome=outcome,
    predictors=main_predictors,
    robust=False
)

poisson_main_robust = fit_poisson_glm(
    model_df,
    outcome=outcome,
    predictors=main_predictors,
    robust=True
)

nb_main = fit_negative_binomial(
    model_df,
    outcome=outcome,
    predictors=main_predictors
)

overdisp_main = poisson_overdispersion(poisson_main)

print("\nPoisson main:")
print(poisson_main.summary())

print("\nPoisson main robust HC3:")
print(poisson_main_robust.summary())

print("\nNegative Binomial main:")
print(nb_main.summary())

print("\nOverdispersion main:")
for key, value in overdisp_main.items():
    print(key, "=", round(value, 4))


print("\n" + "=" * 100)
print("MODEL B: Sensitivity model with Sustainability_importance_mean")
print("=" * 100)

poisson_sensitivity = fit_poisson_glm(
    model_df,
    outcome=outcome,
    predictors=sensitivity_predictors,
    robust=False
)

poisson_sensitivity_robust = fit_poisson_glm(
    model_df,
    outcome=outcome,
    predictors=sensitivity_predictors,
    robust=True
)

nb_sensitivity = fit_negative_binomial(
    model_df,
    outcome=outcome,
    predictors=sensitivity_predictors
)

overdisp_sensitivity = poisson_overdispersion(poisson_sensitivity)

print("\nNegative Binomial sensitivity:")
print(nb_sensitivity.summary())


print("\n" + "=" * 100)
print("MODEL C: Exploratory model with individual pressure sources")
print("=" * 100)

poisson_pressure_sources = fit_poisson_glm(
    model_df,
    outcome=outcome,
    predictors=pressure_source_predictors,
    robust=True
)

nb_pressure_sources = fit_negative_binomial(
    model_df,
    outcome=outcome,
    predictors=pressure_source_predictors
)

print("\nPoisson pressure sources robust HC3:")
print(poisson_pressure_sources.summary())

print("\nNegative Binomial pressure sources:")
print(nb_pressure_sources.summary())


model_df["SDG_count_ge_2"] = (model_df["SDG_count"] >= 2).astype(int)

print("\nDistribution of SDG_count_ge_2:")
print(model_df["SDG_count_ge_2"].value_counts().sort_index().to_string())

logit_formula = "SDG_count_ge_2 ~ " + " + ".join(logit_predictors)

logit_main = smf.logit(
    formula=logit_formula,
    data=model_df
).fit(disp=False)

print("\n" + "=" * 100)
print("MODEL D: Logistic sensitivity model, SDG_count >= 2")
print("=" * 100)
print(logit_main.summary())


count_coefficients = pd.concat(
    [
        count_model_irr_table(poisson_main, "Poisson main, non-tautological", "Poisson"),
        count_model_irr_table(poisson_main_robust, "Poisson main robust HC3, non-tautological", "Poisson robust HC3"),
        count_model_irr_table(nb_main, "Negative Binomial main, non-tautological", "Negative Binomial"),
        count_model_irr_table(poisson_sensitivity, "Poisson sensitivity", "Poisson"),
        count_model_irr_table(poisson_sensitivity_robust, "Poisson sensitivity robust HC3", "Poisson robust HC3"),
        count_model_irr_table(nb_sensitivity, "Negative Binomial sensitivity", "Negative Binomial"),
        count_model_irr_table(poisson_pressure_sources, "Poisson pressure sources robust HC3", "Poisson robust HC3"),
        count_model_irr_table(nb_pressure_sources, "Negative Binomial pressure sources", "Negative Binomial"),
    ],
    ignore_index=True
)

logit_coefficients = logit_or_table(
    logit_main,
    "Logit SDG_count_ge_2, non-tautological"
)


model_fit_summary = pd.DataFrame(
    [
        fit_summary_glm(poisson_main, "Poisson main, non-tautological", poisson_null),
        fit_summary_glm(poisson_main_robust, "Poisson main robust HC3, non-tautological", poisson_null),
        fit_summary_discrete(nb_main, "Negative Binomial main, non-tautological", nb_null),
        fit_summary_glm(poisson_sensitivity, "Poisson sensitivity", poisson_null),
        fit_summary_glm(poisson_sensitivity_robust, "Poisson sensitivity robust HC3", poisson_null),
        fit_summary_discrete(nb_sensitivity, "Negative Binomial sensitivity", nb_null),
        fit_summary_glm(poisson_pressure_sources, "Poisson pressure sources robust HC3", poisson_null),
        fit_summary_discrete(nb_pressure_sources, "Negative Binomial pressure sources", nb_null),
    ]
)

overdispersion_summary = pd.DataFrame(
    [
        {"model": "Poisson main, non-tautological", **overdisp_main},
        {"model": "Poisson sensitivity", **overdisp_sensitivity},
    ]
)


print("\n" + "=" * 100)
print("SUMMARY OF NON-TAUTOLOGICAL MODELS")
print("=" * 100)

print("\nVIF main:")
print(round_numeric(vif_main).to_string(index=False))

print("\nVIF sensitivity:")
print(round_numeric(vif_sensitivity).to_string(index=False))

print("\nVIF pressure sources:")
print(round_numeric(vif_pressure_sources).to_string(index=False))

print("\nOverdispersion:")
print(round_numeric(overdispersion_summary).to_string(index=False))

print("\nModel fit summary:")
print(round_numeric(model_fit_summary).to_string(index=False))

print("\nCount model coefficients, IRR:")
print(
    round_numeric(count_coefficients)[
        [
            "model",
            "model_type",
            "label",
            "coef",
            "std_error",
            "p_value",
            "IRR",
            "IRR_CI_low",
            "IRR_CI_high"
        ]
    ].to_string(index=False)
)

print("\nLogit model coefficients, OR:")
print(
    round_numeric(logit_coefficients)[
        [
            "model",
            "label",
            "coef",
            "std_error",
            "p_value",
            "OR",
            "OR_CI_low",
            "OR_CI_high"
        ]
    ].to_string(index=False)
)


print("\n" + "=" * 100)
print("AUTOMATSKI SAŽETAK GLAVNOG NETAUTOLOŠKOG NEGATIVE BINOMIAL MODELA")
print("=" * 100)

nb_main_table = count_coefficients[
    count_coefficients["model"] == "Negative Binomial main, non-tautological"
].copy()

nb_main_table = nb_main_table[
    ~nb_main_table["term"].isin(["const", "alpha"])
].copy()

nb_main_table = nb_main_table.sort_values("p_value")

for _, row in nb_main_table.iterrows():
    if row["p_value"] < 0.05:
        sig = "statistically significant"
    elif row["p_value"] < 0.10:
        sig = "borderline / indicative"
    else:
        sig = "not statistically significant"

    print(
        f"- {row['label']}: IRR = {row['IRR']:.3f}, "
        f"95% CI [{row['IRR_CI_low']:.3f}, {row['IRR_CI_high']:.3f}], "
        f"p = {row['p_value']:.4f}; {sig}."
    )

print(
    "\nInterpretation: this model excludes q16 / Sustainable_success_orientation. "
    "Therefore, the results better address which external, resource-related, "
    "change-related, and value-related factors remain associated with SDG_count "
    "when they are examined jointly."
)


output_path = "/content/RQ2_step4_nontautological_multivariate_models.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    model_df.to_excel(
        writer,
        sheet_name="model_data",
        index=False
    )

    excluded_tautological_predictors_df.to_excel(
        writer,
        sheet_name="excluded_tautological",
        index=False
    )

    numeric_sdg_count_ranked.to_excel(
        writer,
        sheet_name="step3b_numeric_ranked",
        index=False
    )

    binary_sdg_count_ranked.to_excel(
        writer,
        sheet_name="step3b_binary_ranked",
        index=False
    )

    vif_main.to_excel(
        writer,
        sheet_name="vif_main",
        index=False
    )

    vif_sensitivity.to_excel(
        writer,
        sheet_name="vif_sensitivity",
        index=False
    )

    vif_pressure_sources.to_excel(
        writer,
        sheet_name="vif_pressure_sources",
        index=False
    )

    overdispersion_summary.to_excel(
        writer,
        sheet_name="overdispersion",
        index=False
    )

    model_fit_summary.to_excel(
        writer,
        sheet_name="model_fit_summary",
        index=False
    )

    count_coefficients.to_excel(
        writer,
        sheet_name="count_model_coefficients",
        index=False
    )

    logit_coefficients.to_excel(
        writer,
        sheet_name="logit_coefficients",
        index=False
    )

print("\nRQ2 Step 4 non-tautological results saved to:")
print(output_path)

print("\nRQ2 Step 4, non-tautological version, completed.")
