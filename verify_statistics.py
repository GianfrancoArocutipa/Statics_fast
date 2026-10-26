import pandas as pd
import numpy as np
from scipy import stats
import statsmodels.api as sm
from statsmodels.formula.api import ols
from statsmodels.stats.multicomp import pairwise_tukeyhsd

# Load dataset
df = pd.read_csv(r'c:\Users\PC\Desktop\Estats\archive\Heterogeneous_Data .csv', low_memory=False)
df_rep = df[df['_type'] == 'RepositoryItem'].copy()

print(f"Total repositories: {len(df_rep)}")

# Create transformed variables ln(x + 1)
df_rep['ln_stars_1'] = np.log(df_rep['stars'] + 1)
df_rep['ln_forks_1'] = np.log(df_rep['forks'] + 1)
df_rep['ln_open_issues_1'] = np.log(df_rep['open_issues'] + 1)
df_rep['ln_size_kb_1'] = np.log(df_rep['size_kb'] + 1)

# Difference for paired test
df_rep['diff_stars_forks'] = df_rep['ln_stars_1'] - df_rep['ln_forks_1']

print("\n--- 1. PAIRED SAMPLES TEST (ln_stars_1 vs ln_forks_1) ---")
mean_diff = df_rep['diff_stars_forks'].mean()
std_diff = df_rep['diff_stars_forks'].std(ddof=1)
n_paired = len(df_rep)
se_diff = std_diff / np.sqrt(n_paired)
t_paired = mean_diff / se_diff
corr = df_rep['ln_stars_1'].corr(df_rep['ln_forks_1'])
sw_diff = stats.shapiro(df_rep['diff_stars_forks'])

print(f"Mean diff: {mean_diff:.10f}")
print(f"Std diff: {std_diff:.10f}")
print(f"SE diff: {se_diff:.10f}")
print(f"95% CI diff: [{mean_diff - 1.9615*se_diff:.10f}, {mean_diff + 1.9615*se_diff:.10f}]")
print(f"t-statistic: {t_paired:.3f}, df: {n_paired - 1}, p-value: {stats.t.sf(t_paired, n_paired - 1)*2}")
print(f"Correlation: {corr:.3f}")
print(f"Shapiro-Wilk diff: W = {sw_diff.statistic:.4f}, p = {sw_diff.pvalue:.4e}")

print("\n--- 2. TWO INDEPENDENT GROUPS (ln_forks_1 by owner_type) ---")
org_forks = df_rep[df_rep['owner_type'] == 'Organization']['ln_forks_1']
user_forks = df_rep[df_rep['owner_type'] == 'User']['ln_forks_1']
levene_forks = stats.levene(org_forks, user_forks)
sw_org_forks = stats.shapiro(org_forks)
sw_user_forks = stats.shapiro(user_forks)
ttest_ind_equal = stats.ttest_ind(org_forks, user_forks, equal_var=True)
ttest_ind_welch = stats.ttest_ind(org_forks, user_forks, equal_var=False)

print(f"Org: n={len(org_forks)}, mean={org_forks.mean():.10f}, std={org_forks.std():.10f}")
print(f"User: n={len(user_forks)}, mean={user_forks.mean():.10f}, std={user_forks.std():.10f}")
print(f"Levene: F = {levene_forks.statistic:.3f}, p = {levene_forks.pvalue:.4f}")
print(f"Shapiro-Wilk Org: W = {sw_org_forks.statistic:.4f}, p = {sw_org_forks.pvalue:.4e}")
print(f"Shapiro-Wilk User: W = {sw_user_forks.statistic:.4f}, p = {sw_user_forks.pvalue:.4e}")
print(f"Welch t: t = {ttest_ind_welch.statistic:.3f}, df = {ttest_ind_welch.df:.3f}, p = {ttest_ind_welch.pvalue:.4f}")

print("\n--- 3. ANOVA (ln_open_issues_1 across 3 licenses) ---")
df_lic3 = df_rep[df_rep['license_spdx_id'].isin(['MIT', 'Apache-2.0', 'GPL-3.0'])].copy()
mit_iss = df_lic3[df_lic3['license_spdx_id'] == 'MIT']['ln_open_issues_1']
apa_iss = df_lic3[df_lic3['license_spdx_id'] == 'Apache-2.0']['ln_open_issues_1']
gpl_iss = df_lic3[df_lic3['license_spdx_id'] == 'GPL-3.0']['ln_open_issues_1']

print(f"MIT: n={len(mit_iss)}, mean={mit_iss.mean():.10f}, std={mit_iss.std():.10f}")
print(f"Apache-2.0: n={len(apa_iss)}, mean={apa_iss.mean():.10f}, std={apa_iss.std():.10f}")
print(f"GPL-3.0: n={len(gpl_iss)}, mean={gpl_iss.mean():.10f}, std={gpl_iss.std():.10f}")

levene_anova = stats.levene(mit_iss, apa_iss, gpl_iss, center='mean')
print(f"Levene (mean): F = {levene_anova.statistic:.3f}, p = {levene_anova.pvalue:.4f}")

sw_mit = stats.shapiro(mit_iss)
sw_apa = stats.shapiro(apa_iss)
sw_gpl = stats.shapiro(gpl_iss)
print(f"Shapiro-Wilk MIT: W = {sw_mit.statistic:.4f}, p = {sw_mit.pvalue:.4e}")
print(f"Shapiro-Wilk Apache: W = {sw_apa.statistic:.4f}, p = {sw_apa.pvalue:.4e}")
print(f"Shapiro-Wilk GPL: W = {sw_gpl.statistic:.4f}, p = {sw_gpl.pvalue:.4e}")

f_oneway = stats.f_oneway(mit_iss, apa_iss, gpl_iss)
print(f"ANOVA F: {f_oneway.statistic:.3f}, p = {f_oneway.pvalue:.4e}")

# Tukey HSD
tukey = pairwise_tukeyhsd(df_lic3['ln_open_issues_1'], df_lic3['license_spdx_id'], alpha=0.05)
print("\nTukey HSD Summary:")
print(tukey)
