import numpy as np
import scipy.stats as stats

print("=== PROBLEMA 1: Chi-cuadrado de bondad de ajuste ===")
# Observados: 42 (autenticación), 25 (conexión), 33 (rendimiento), 20 (interfaz). Total = 120.
obs = np.array([42, 25, 33, 20])
esp = np.array([30, 30, 30, 30])
chi2_stat, p_val1 = stats.chisquare(obs, esp)
print(f"Chi2 = {chi2_stat:.4f}, gl = {len(obs)-1}, p-valor = {p_val1:.4f}")

print("\n=== PROBLEMA 2: Dos algoritmos independientes ===")
alg_A = np.array([120, 135, 128, 142, 130, 125, 138, 132, 145, 127], dtype=float)
alg_B = np.array([110, 118, 115, 122, 108, 120, 112, 116, 125, 119, 114], dtype=float)

print(f"Alg A: n={len(alg_A)}, mean={np.mean(alg_A):.2f}, std={np.std(alg_A, ddof=1):.2f}")
print(f"Alg B: n={len(alg_B)}, mean={np.mean(alg_B):.2f}, std={np.std(alg_B, ddof=1):.2f}")

sw_A = stats.shapiro(alg_A)
sw_B = stats.shapiro(alg_B)
print(f"Shapiro Alg A: W={sw_A.statistic:.4f}, p={sw_A.pvalue:.4f}")
print(f"Shapiro Alg B: W={sw_B.statistic:.4f}, p={sw_B.pvalue:.4f}")

lev2 = stats.levene(alg_A, alg_B, center='mean')
print(f"Levene (mean): F={lev2.statistic:.4f}, p={lev2.pvalue:.4f}")

ttest_ind = stats.ttest_ind(alg_A, alg_B, equal_var=True)
print(f"T-test (equal var): t={ttest_ind.statistic:.4f}, gl={len(alg_A)+len(alg_B)-2}, p={ttest_ind.pvalue:.6f}")

mw = stats.mannwhitneyu(alg_A, alg_B, alternative='two-sided')
print(f"Mann-Whitney: U={mw.statistic:.4f}, p={mw.pvalue:.6f}")

print("\n=== PROBLEMA 3: Prueba de Rachas ===")
seq_str = "E, E, F, E, F, F, E, E, F, E, F, F, E, F, E, E, F, F, E, F, E, E, F, E, F, F, E, E, F, E"
seq = [x.strip() for x in seq_str.split(',')]
n_total = len(seq)
n_E = seq.count('E')
n_F = seq.count('F')

runs = 1
for i in range(1, len(seq)):
    if seq[i] != seq[i-1]:
        runs += 1

mu_R = 1 + (2 * n_E * n_F) / n_total
var_R = (2 * n_E * n_F * (2 * n_E * n_F - n_total)) / ((n_total ** 2) * (n_total - 1))
sd_R = np.sqrt(var_R)
z_R = (runs - mu_R) / sd_R
p_val3 = 2 * (1 - stats.norm.cdf(abs(z_R)))

print(f"Total n={n_total}, n(E)={n_E}, n(F)={n_F}, Rachas (R)={runs}")
print(f"Media esperada={mu_R:.4f}, Varianza={var_R:.4f}, Desv={sd_R:.4f}")
print(f"Z={z_R:.4f}, p-valor asintotico (2-colas)={p_val3:.4f}")

print("\n=== PROBLEMA 4: Prueba Binomial ===")
p_binom_exact = stats.binom.cdf(40, 50, 0.90)
p_binom_2tail = stats.binomtest(40, 50, 0.90, alternative='two-sided').pvalue
p_binom_less = stats.binomtest(40, 50, 0.90, alternative='less').pvalue
print(f"Binomial exacto P(X <= 40): {p_binom_exact:.4f}")
print(f"Binomtest alternative less: {p_binom_less:.4f}")
print(f"Binomtest alternative two-sided: {p_binom_2tail:.4f}")

print("\n=== PROBLEMA 5: Tres servicios de almacenamiento en la nube ===")
serv_A = np.array([95, 102, 98, 105, 100, 97, 103, 99, 106, 101, 96], dtype=float)
serv_B = np.array([115, 122, 118, 125, 120, 117, 123, 119, 126, 121, 116, 124, 128], dtype=float)
serv_C = np.array([135, 142, 138, 145, 140, 137, 143, 139, 146, 141, 136, 144, 148, 134, 147], dtype=float)

print(f"Serv A: n={len(serv_A)}, mean={np.mean(serv_A):.2f}, std={np.std(serv_A, ddof=1):.2f}")
print(f"Serv B: n={len(serv_B)}, mean={np.mean(serv_B):.2f}, std={np.std(serv_B, ddof=1):.2f}")
print(f"Serv C: n={len(serv_C)}, mean={np.mean(serv_C):.2f}, std={np.std(serv_C, ddof=1):.2f}")

sw_sA = stats.shapiro(serv_A)
sw_sB = stats.shapiro(serv_B)
sw_sC = stats.shapiro(serv_C)
print(f"Shapiro Serv A: W={sw_sA.statistic:.4f}, p={sw_sA.pvalue:.4f}")
print(f"Shapiro Serv B: W={sw_sB.statistic:.4f}, p={sw_sB.pvalue:.4f}")
print(f"Shapiro Serv C: W={sw_sC.statistic:.4f}, p={sw_sC.pvalue:.4f}")

lev_serv = stats.levene(serv_A, serv_B, serv_C, center='mean')
print(f"Levene (mean): F={lev_serv.statistic:.4f}, p={lev_serv.pvalue:.4f}")

f_val, p_val5 = stats.f_oneway(serv_A, serv_B, serv_C)
print(f"ANOVA: F={f_val:.4f}, p={p_val5:.6e}")

kw_serv = stats.kruskal(serv_A, serv_B, serv_C)
print(f"Kruskal-Wallis: H={kw_serv.statistic:.4f}, p={kw_serv.pvalue:.6e}")
