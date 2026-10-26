import numpy as np
import pandas as pd
from scipy import stats

def compute_spss_mann_whitney(group1, group2, name1="Grupo 1", name2="Grupo 2"):
    n1 = len(group1)
    n2 = len(group2)
    N = n1 + n2
    
    # Combined data with tags
    data = [(val, 1) for val in group1] + [(val, 2) for val in group2]
    # Sort values
    data_sorted = sorted(data, key=lambda x: x[0])
    
    # Assign ranks with ties (fractional ranks)
    ranks = pd.Series([x[0] for x in data_sorted]).rank(method="average").values
    
    # Sum of ranks for group 1 and group 2
    sum_ranks_1 = sum(r for r, d in zip(ranks, data_sorted) if d[1] == 1)
    sum_ranks_2 = sum(r for r, d in zip(ranks, data_sorted) if d[1] == 2)
    
    mean_rank_1 = sum_ranks_1 / n1
    mean_rank_2 = sum_ranks_2 / n2
    
    # Mann-Whitney U
    u1 = sum_ranks_1 - (n1 * (n1 + 1)) / 2.0
    u2 = sum_ranks_2 - (n2 * (n2 + 1)) / 2.0
    u = min(u1, u2)
    
    # W de Wilcoxon is the sum of ranks of the smaller group or group with smaller U
    # In SPSS, Wilcoxon W is usually the sum of ranks of the group corresponding to min(U)
    # or the second group if specified
    if u == u1:
        w_wilcoxon = sum_ranks_1
    else:
        w_wilcoxon = sum_ranks_2
        
    # Tie correction
    # T = sum(t^3 - t)
    unique_vals, counts = np.unique([x[0] for x in data], return_counts=True)
    tie_sum = sum(t**3 - t for t in counts if t > 1)
    
    sigma_u = np.sqrt((n1 * n2 / 12.0) * ((N + 1) - tie_sum / (N * (N - 1))))
    mean_u = (n1 * n2) / 2.0
    
    # Z with continuity correction (SPSS applies continuity correction)
    # If U > mean_u: z = (U - 0.5 - mean_u) / sigma_u
    # If U < mean_u: z = (U + 0.5 - mean_u) / sigma_u
    if u > mean_u:
        z = (u - 0.5 - mean_u) / sigma_u
    else:
        z = (u + 0.5 - mean_u) / sigma_u
        
    p_asymp = 2 * (1 - stats.norm.cdf(abs(z)))
    
    # Scipy mannwhitneyu for comparison
    scipy_res = stats.mannwhitneyu(group1, group2, alternative="two-sided")
    
    return {
        "n1": n1, "n2": n2, "N": N,
        "name1": name1, "name2": name2,
        "mean_rank_1": mean_rank_1, "sum_ranks_1": sum_ranks_1,
        "mean_rank_2": mean_rank_2, "sum_ranks_2": sum_ranks_2,
        "u1": u1, "u2": u2, "u": u,
        "w_wilcoxon": w_wilcoxon,
        "z": z,
        "p_asymp": p_asymp,
        "scipy_u": scipy_res.statistic,
        "scipy_p": scipy_res.pvalue
    }

ex1_java = [120.0, 115.0, 130.0, 90.0, 125.0, 118.0, 116.0, 117.0]
ex1_python = [90.0, 85.0, 88.0, 92.0, 87.0]

ex2_sist_a = [150.0, 137.0, 148.0, 149.0, 152.0, 153.0, 151.0, 150.0, 154.0, 155.0]
ex2_sist_b = [160.0, 158.0, 159.0, 161.0, 162.0, 160.0, 157.0]

ex3_serv_a = [12.0, 10.0, 11.0, 13.0, 11.0, 12.0, 10.0, 13.0, 12.0, 11.0, 10.0, 12.0, 13.0, 11.0, 10.0, 12.0, 11.0, 10.0, 13.0, 12.0]
ex3_serv_b = [14.0, 15.0, 13.0, 14.0, 15.0, 14.0, 13.0, 15.0, 14.0, 13.0, 15.0, 14.0, 13.0, 15.0, 14.0, 13.0, 15.0, 14.0]

ex4_alg_a = [70.0, 65.0, 68.0, 72.0, 71.0, 69.0, 67.0, 73.0, 70.0, 66.0, 68.0, 71.0, 69.0, 67.0, 72.0, 70.0, 68.0, 65.0, 73.0, 71.0]
ex4_alg_b = [75.0, 78.0, 76.0, 77.0, 79.0, 75.0, 76.0, 78.0, 80.0, 77.0, 76.0, 79.0, 65.0, 77.0, 78.0]

for idx, (g1, g2, n1, n2) in enumerate([
    (ex1_java, ex1_python, "Java", "Python"),
    (ex2_sist_a, ex2_sist_b, "Sistema A", "Sistema B"),
    (ex3_serv_a, ex3_serv_b, "Servidor A", "Servidor B"),
    (ex4_alg_a, ex4_alg_b, "Algoritmo A", "Algoritmo B")
], 1):
    res = compute_spss_mann_whitney(g1, g2, n1, n2)
    print(f"=== EXERCICE {idx}: {n1} vs {n2} ===")
    print(f"  {n1} (n={res['n1']}): Rango promedio = {res['mean_rank_1']:.2f}, Suma de rangos = {res['sum_ranks_1']:.2f}")
    print(f"  {n2} (n={res['n2']}): Rango promedio = {res['mean_rank_2']:.2f}, Suma de rangos = {res['sum_ranks_2']:.2f}")
    print(f"  U1 = {res['u1']}, U2 = {res['u2']} -> U de Mann-Whitney = {res['u']}")
    print(f"  W de Wilcoxon = {res['w_wilcoxon']}")
    print(f"  Z = {res['z']:.3f}")
    print(f"  Sig. asintótica (bilateral) = {res['p_asymp']:.3f} (exact scipy p={res['scipy_p']:.4f})")
    print()
