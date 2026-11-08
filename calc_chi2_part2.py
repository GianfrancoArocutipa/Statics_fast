# -*- coding: utf-8 -*-
import numpy as np
from scipy import stats

def spss_chisq_table(obs):
    obs = np.array(obs, dtype=float)
    N = obs.sum()
    r_sum = obs.sum(axis=1, keepdims=True)
    c_sum = obs.sum(axis=0, keepdims=True)
    exp = (r_sum @ c_sum) / N
    
    # Pearson
    pearson_val = np.sum((obs - exp)**2 / exp)
    df = (obs.shape[0] - 1) * (obs.shape[1] - 1)
    pearson_p = 1 - stats.chi2.cdf(pearson_val, df)
    
    # Likelihood ratio G^2
    with np.errstate(divide='ignore', invalid='ignore'):
        lr_terms = np.where(obs > 0, obs * np.log(obs / exp), 0.0)
    lr_val = 2 * np.sum(lr_terms)
    lr_p = 1 - stats.chi2.cdf(lr_val, df)
    
    # Linear-by-linear association: (N - 1) * r^2 where r is Pearson correlation of row idx (1..R) and col idx (1..C)
    r_idx = np.arange(1, obs.shape[0] + 1).reshape(-1, 1)
    c_idx = np.arange(1, obs.shape[1] + 1).reshape(1, -1)
    mean_r = np.sum(obs * r_idx) / N
    mean_c = np.sum(obs * c_idx) / N
    cov_rc = np.sum(obs * (r_idx - mean_r) * (c_idx - mean_c))
    var_r = np.sum(obs * (r_idx - mean_r)**2)
    var_c = np.sum(obs * (c_idx - mean_c)**2)
    r_corr = cov_rc / np.sqrt(var_r * var_c)
    lbl_val = (N - 1) * (r_corr**2)
    lbl_p = 1 - stats.chi2.cdf(lbl_val, 1)
    
    low_cells = int(np.sum(exp < 5.0))
    total_cells = obs.size
    low_pct = 100.0 * low_cells / total_cells
    min_exp = float(np.min(exp))
    
    res = {
        'N': int(N), 'exp': exp, 'df': df,
        'pearson': (pearson_val, pearson_p),
        'lr': (lr_val, lr_p),
        'lbl': (lbl_val, lbl_p),
        'low_cells': low_cells, 'low_pct': low_pct, 'min_exp': min_exp
    }
    
    if obs.shape == (2, 2):
        # Continuity correction (Yates)
        yates_val = np.sum((np.maximum(0.0, np.abs(obs - exp) - 0.5))**2 / exp)
        yates_p = 1 - stats.chi2.cdf(yates_val, 1)
        # Fisher exact
        _, f_two = stats.fisher_exact(obs, alternative='two-sided')
        _, f_less = stats.fisher_exact(obs, alternative='less')
        _, f_greater = stats.fisher_exact(obs, alternative='greater')
        f_one = min(f_less, f_greater)
        # McNemar exact (binomial on off-diagonal b and c: obs[0,1] and obs[1,0])
        b, c = int(obs[0, 1]), int(obs[1, 0])
        if b + c > 0:
            mcn_p = stats.binomtest(min(b, c), n=b+c, p=0.5, alternative='two-sided').pvalue
        else:
            mcn_p = 1.0
        # Cohen's Kappa
        po = (obs[0,0] + obs[1,1]) / N
        pe = (r_sum[0,0]*c_sum[0,0] + r_sum[1,0]*c_sum[0,1]) / (N**2)
        kappa = (po - pe) / (1 - pe)
        # SPSS Asymptotic SE for Kappa and Approx T (under H0)
        p_mat = obs / N
        r_p = r_sum.flatten() / N
        c_p = c_sum.flatten() / N
        term1 = 0.0
        for i in range(2):
            term1 += p_mat[i,i] * ((1 - (r_p[i] + c_p[i])*(1 - kappa))**2)
        term2 = 0.0
        for i in range(2):
            for j in range(2):
                if i != j:
                    term2 += p_mat[i,j] * ((r_p[j] + c_p[i])**2)
        term2 *= (1 - kappa)**2
        term3 = (kappa - pe * (1 - kappa))**2
        ase = np.sqrt((term1 + term2 - term3) / (N * ((1 - pe)**2)))
        sum_rc_sq = np.sum(r_p * c_p * (r_p + c_p))
        se0 = np.sqrt((pe + pe**2 - sum_rc_sq) / (N * ((1 - pe)**2)))
        t_approx = kappa / se0
        p_kappa = 2 * (1 - stats.norm.cdf(abs(t_approx)))
        res.update({
            'yates': (yates_val, yates_p),
            'fisher': (f_two, f_one),
            'mcnemar_p': mcn_p,
            'kappa': (kappa, ase, t_approx, p_kappa)
        })
    return res

if __name__ == '__main__':
    tables = [
        ('Ex1_Indep1', [[40, 30, 20], [50, 20, 30], [20, 10, 80]]),
        ('Ex2_Indep2', [[50, 25], [30, 40], [35, 20]]),
        ('Ex3_Homog1', [[20, 80], [30, 70], [10, 90]]),
        ('Ex4_Homog2', [[70, 30, 60, 40, 50, 50], [80, 20, 70, 30, 60, 40], [75, 25, 65, 35, 55, 45]]),
        ('Ex5_McNemar1', [[15, 15], [5, 65]]),
        ('Ex6_McNemar2', [[8, 12], [2, 78]]),
        ('Ex7_Kappa1', [[72, 13], [8, 7]]),
        ('Ex8_Kappa2_inner', [[8, 12], [2, 78]]),
        ('Ex9_Fisher1', [[27, 3], [8, 2]]),
        ('Ex10_Fisher2', [[25, 10], [13, 2]])
    ]

    for name, mat in tables:
        r = spss_chisq_table(mat)
        print(f"=== {name} (N={r['N']}) ===")
        print(f"  Pearson={r['pearson'][0]:.3f} (df={r['df']}, p={r['pearson'][1]:.3f}), LR={r['lr'][0]:.3f} (p={r['lr'][1]:.3f}), LBL={r['lbl'][0]:.3f} (p={r['lbl'][1]:.3f})")
        print(f"  Footnote a: {r['low_cells']} casillas ({r['low_pct']:.1f}%) < 5, min_exp={r['min_exp']:.2f}")
        if 'yates' in r:
            print(f"  Yates={r['yates'][0]:.3f} (p={r['yates'][1]:.3f}), Fisher 2-sided={r['fisher'][0]:.3f}, 1-sided={r['fisher'][1]:.3f}, McNemar p={r['mcnemar_p']:.3f}")
            k, ase, t_k, p_k = r['kappa']
            print(f"  Kappa={k:.3f}, ASE={ase:.3f}, T={t_k:.3f}, Sig={p_k:.3f}")
        obs_arr = np.array(mat, dtype=float)
        exp = r['exp']
        N_val = r['N']
        for i in range(obs_arr.shape[0]):
            r_tot = obs_arr[i].sum()
            print(f"    Row {i+1}: obs={obs_arr[i].astype(int).tolist()} (tot={int(r_tot)}), exp={[round(x,1) for x in exp[i]]}, r_pct={[round(100*x/r_tot,1) for x in obs_arr[i]]}, t_pct={[round(100*x/N_val,1) for x in obs_arr[i]]}")
        c_tot = obs_arr.sum(axis=0)
        print(f"    Total: obs={c_tot.astype(int).tolist()} (N={N_val}), pct={[round(100*x/N_val,1) for x in c_tot]}")

