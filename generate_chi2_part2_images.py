# -*- coding: utf-8 -*-
"""
Generador de imágenes SPSS idénticas al estilo de 'Ejercicios Chi cuadrado 2da. parte (1).docx'
para los 10 ejercicios de:
1-2: Chi cuadrado de independencia
3-4: Chi cuadrado de homogeneidad
5-6: Chi cuadrado de McNemar
7-8: Índice de Kappa de Cohen
9-10: Prueba exacta de Fisher
"""

import os
import shutil
from playwright.sync_api import sync_playwright

output_dir = "generated_chi2_part2_images"
os.makedirs(output_dir, exist_ok=True)

SPSS_LIGHT_CSS = """
body {
    font-family: Arial, Helvetica, sans-serif;
    background-color: #ffffff;
    margin: 0;
    padding: 6px;
    display: inline-block;
}

.spss-wrap {
    background-color: #ffffff;
    display: inline-block;
    padding: 4px 6px;
}

.spss-title {
    text-align: center;
    font-size: 12px;
    font-weight: bold;
    color: #000000;
    margin-bottom: 6px;
    letter-spacing: 0.2px;
}

table.spss-tbl {
    border-collapse: collapse;
    font-size: 11px;
    color: #000000;
    margin: 0 auto;
}

table.spss-tbl thead th {
    font-weight: normal;
    color: #264A60;
    padding: 3px 10px;
    text-align: center;
    background-color: #ffffff;
    border-left: 1px solid #D3D3D3;
}

table.spss-tbl thead th.no-bl {
    border-left: none;
}

table.spss-tbl thead tr.bottom-hdr th {
    border-bottom: 1px solid #264A60;
}

table.spss-tbl tbody td.rh {
    background-color: #E0E0E0;
    color: #264A60;
    text-align: left;
    padding: 3px 8px;
    vertical-align: top;
    border-bottom: 1px solid #AEAEAE;
}

table.spss-tbl tbody td.rh-sub {
    background-color: #E0E0E0;
    color: #264A60;
    text-align: left;
    padding: 3px 8px;
    border-bottom: 1px solid #AEAEAE;
}

table.spss-tbl tbody td.dc {
    background-color: #F9F9F9;
    color: #000000;
    text-align: right;
    padding: 3px 10px;
    border-left: 1px solid #D3D3D3;
    border-bottom: 1px solid #AEAEAE;
}

table.spss-tbl tbody tr.last-row td {
    border-bottom: 1px solid #264A60 !important;
}

.spss-fn {
    font-size: 11px;
    color: #000000;
    margin-top: 4px;
    padding-left: 10px;
    text-align: left;
    line-height: 1.35;
}
"""

def make_crosstab_3metric(row_var, col_var, row_labels, col_labels, obs_mat, exp_mat, rpct_mat, col_totals, col_pcts, N_total):
    """
    Genera Tabla cruzada con Recuento, Recuento esperado y % dentro de [row_var]
    idéntica a image1.png del docente.
    """
    ncols = len(col_labels)
    col_hdrs = "".join(f'<th>{c}</th>' for c in col_labels)
    
    tbody = ""
    for i, rlabel in enumerate(row_labels):
        row_span_first = f'<td class="rh" rowspan="{len(row_labels)*3}" style="border-bottom: 1px solid #AEAEAE;">{row_var}</td>' if i == 0 else ""
        r_tot_obs = sum(obs_mat[i])
        r_tot_exp = f"{r_tot_obs:.1f}"
        
        # Subfila 1: Recuento
        c_obs = "".join(f'<td class="dc">{v}</td>' for v in obs_mat[i]) + f'<td class="dc">{r_tot_obs}</td>'
        tbody += f"""
        <tr>
            {row_span_first}
            <td class="rh" rowspan="3">{rlabel}</td>
            <td class="rh-sub">Recuento</td>
            {c_obs}
        </tr>"""
        
        # Subfila 2: Recuento esperado
        c_exp = "".join(f'<td class="dc">{v:.1f}</td>' for v in exp_mat[i]) + f'<td class="dc">{r_tot_exp}</td>'
        tbody += f"""
        <tr>
            <td class="rh-sub">Recuento esperado</td>
            {c_exp}
        </tr>"""
        
        # Subfila 3: % dentro de row_var
        c_pct = "".join(f'<td class="dc">{v:.1f}%</td>' for v in rpct_mat[i]) + '<td class="dc">100.0%</td>'
        tbody += f"""
        <tr>
            <td class="rh-sub">% dentro de {row_var}</td>
            {c_pct}
        </tr>"""

    # Total block (3 subfilas)
    t_obs = "".join(f'<td class="dc">{v}</td>' for v in col_totals) + f'<td class="dc">{N_total}</td>'
    t_exp = "".join(f'<td class="dc">{v:.1f}</td>' for v in col_totals) + f'<td class="dc">{N_total:.1f}</td>'
    t_pct = "".join(f'<td class="dc">{v:.1f}%</td>' for v in col_pcts) + '<td class="dc">100.0%</td>'
    tbody += f"""
        <tr>
            <td class="rh" colspan="2" rowspan="3" style="border-bottom: 1px solid #264A60;">Total</td>
            <td class="rh-sub">Recuento</td>
            {t_obs}
        </tr>
        <tr>
            <td class="rh-sub">Recuento esperado</td>
            {t_exp}
        </tr>
        <tr class="last-row">
            <td class="rh-sub">% dentro de {row_var}</td>
            {t_pct}
        </tr>"""

    return f"""
<div class="spss-wrap">
    <div class="spss-title">Tabla cruzada {row_var}*{col_var}</div>
    <table class="spss-tbl">
        <thead>
            <tr>
                <th colspan="3" rowspan="2" class="no-bl" style="border-bottom: 1px solid #264A60;"></th>
                <th colspan="{ncols}" style="border-left: none;">{col_var}</th>
                <th rowspan="2" style="vertical-align: bottom; border-bottom: 1px solid #264A60;">Total</th>
            </tr>
            <tr class="bottom-hdr">
                {col_hdrs}
            </tr>
        </thead>
        <tbody>
            {tbody}
        </tbody>
    </table>
</div>
"""

def make_crosstab_2metric_total(row_var, col_var, row_labels, col_labels, obs_mat, tpct_mat, col_totals, col_pcts, N_total):
    """
    Genera Tabla cruzada con Recuento y % del total
    idéntica a image3.png del docente.
    """
    ncols = len(col_labels)
    col_hdrs = "".join(f'<th>{c}</th>' for c in col_labels)
    
    tbody = ""
    for i, rlabel in enumerate(row_labels):
        row_span_first = f'<td class="rh" rowspan="{len(row_labels)*2}" style="border-bottom: 1px solid #AEAEAE;">{row_var}</td>' if i == 0 else ""
        r_tot_obs = sum(obs_mat[i])
        r_tot_pct = sum(tpct_mat[i])
        
        c_obs = "".join(f'<td class="dc">{v}</td>' for v in obs_mat[i]) + f'<td class="dc">{r_tot_obs}</td>'
        tbody += f"""
        <tr>
            {row_span_first}
            <td class="rh" rowspan="2">{rlabel}</td>
            <td class="rh-sub">Recuento</td>
            {c_obs}
        </tr>"""
        
        c_pct = "".join(f'<td class="dc">{v:.1f}%</td>' for v in tpct_mat[i]) + f'<td class="dc">{r_tot_pct:.1f}%</td>'
        tbody += f"""
        <tr>
            <td class="rh-sub">% del total</td>
            {c_pct}
        </tr>"""

    t_obs = "".join(f'<td class="dc">{v}</td>' for v in col_totals) + f'<td class="dc">{N_total}</td>'
    t_pct = "".join(f'<td class="dc">{v:.1f}%</td>' for v in col_pcts) + '<td class="dc">100.0%</td>'
    tbody += f"""
        <tr>
            <td class="rh" colspan="2" rowspan="2" style="border-bottom: 1px solid #264A60;">Total</td>
            <td class="rh-sub">Recuento</td>
            {t_obs}
        </tr>
        <tr class="last-row">
            <td class="rh-sub">% del total</td>
            {t_pct}
        </tr>"""

    return f"""
<div class="spss-wrap">
    <div class="spss-title">Tabla cruzada {row_var}*{col_var}</div>
    <table class="spss-tbl">
        <thead>
            <tr>
                <th colspan="3" rowspan="2" class="no-bl" style="border-bottom: 1px solid #264A60;"></th>
                <th colspan="{ncols}" style="border-left: none;">{col_var}</th>
                <th rowspan="2" style="vertical-align: bottom; border-bottom: 1px solid #264A60;">Total</th>
            </tr>
            <tr class="bottom-hdr">
                {col_hdrs}
            </tr>
        </thead>
        <tbody>
            {tbody}
        </tbody>
    </table>
</div>
"""

def fmt_sig(p):
    if p < 0.001:
        return "&lt;.001"
    return f".{int(round(p*1000)):03d}"

def fmt_val(v):
    if v < 1.0:
        return f".{int(round(v*1000)):03d}"
    return f"{v:.3f}"

def make_chisq_standard(pearson_v, pearson_p, lr_v, lr_p, lbl_v, lbl_p, df, N_val, low_cells, low_pct, min_exp):
    """
    Genera Pruebas de chi-cuadrado estándar (Independencia / Homogeneidad)
    idéntica a image2.png del docente.
    """
    return f"""
<div class="spss-wrap">
    <div class="spss-title">Pruebas de chi-cuadrado</div>
    <table class="spss-tbl">
        <thead>
            <tr class="bottom-hdr">
                <th class="no-bl" style="vertical-align: bottom;"></th>
                <th style="vertical-align: bottom;">Valor</th>
                <th style="vertical-align: bottom;">gl</th>
                <th style="vertical-align: bottom;">Significación<br>asintótica<br>(bilateral)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="rh-sub">Chi-cuadrado de Pearson</td>
                <td class="dc">{fmt_val(pearson_v)}<sup><b>a</b></sup></td>
                <td class="dc">{df}</td>
                <td class="dc">{fmt_sig(pearson_p)}</td>
            </tr>
            <tr>
                <td class="rh-sub">Razón de verosimilitud</td>
                <td class="dc">{fmt_val(lr_v)}</td>
                <td class="dc">{df}</td>
                <td class="dc">{fmt_sig(lr_p)}</td>
            </tr>
            <tr>
                <td class="rh-sub">Asociación lineal por lineal</td>
                <td class="dc">{fmt_val(lbl_v)}</td>
                <td class="dc">1</td>
                <td class="dc">{fmt_sig(lbl_p)}</td>
            </tr>
            <tr class="last-row">
                <td class="rh-sub">N de casos válidos</td>
                <td class="dc">{N_val}</td>
                <td class="dc"></td>
                <td class="dc"></td>
            </tr>
        </tbody>
    </table>
    <div class="spss-fn">
        a. {low_cells} casillas ({low_pct:.1f}%) han esperado un recuento menor que 5. El<br>
        &nbsp;&nbsp;&nbsp;&nbsp;recuento mínimo esperado es {min_exp:.2f}.
    </div>
</div>
"""

def make_chisq_2x2(pearson_v, pearson_p, yates_v, yates_p, lr_v, lr_p, f_two, f_one, lbl_v, lbl_p, N_val, low_cells, low_pct, min_exp, mcnemar_p=None):
    """
    Genera Pruebas de chi-cuadrado para tablas 2x2 (McNemar y Prueba exacta de Fisher)
    idéntica a image4.png del docente.
    """
    mcn_row = ""
    fn_c = ""
    if mcnemar_p is not None:
        mcn_row = f"""
            <tr>
                <td class="rh-sub">Prueba de McNemar</td>
                <td class="dc"></td>
                <td class="dc"></td>
                <td class="dc"></td>
                <td class="dc">{fmt_sig(mcnemar_p)}<sup><b>c</b></sup></td>
                <td class="dc"></td>
            </tr>"""
        fn_c = "<br>c. Distribución binomial utilizada."

    return f"""
<div class="spss-wrap">
    <div class="spss-title">Pruebas de chi-cuadrado</div>
    <table class="spss-tbl">
        <thead>
            <tr class="bottom-hdr">
                <th class="no-bl" style="vertical-align: bottom;"></th>
                <th style="vertical-align: bottom;">Valor</th>
                <th style="vertical-align: bottom;">gl</th>
                <th style="vertical-align: bottom;">Significación<br>asintótica<br>(bilateral)</th>
                <th style="vertical-align: bottom;">Significación<br>exacta<br>(bilateral)</th>
                <th style="vertical-align: bottom;">Significación<br>exacta<br>(unilateral)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="rh-sub">Chi-cuadrado de Pearson</td>
                <td class="dc">{fmt_val(pearson_v)}<sup><b>a</b></sup></td>
                <td class="dc">1</td>
                <td class="dc">{fmt_sig(pearson_p)}</td>
                <td class="dc"></td>
                <td class="dc"></td>
            </tr>
            <tr>
                <td class="rh-sub">Corrección de continuidad<sup><b>b</b></sup></td>
                <td class="dc">{fmt_val(yates_v)}</td>
                <td class="dc">1</td>
                <td class="dc">{fmt_sig(yates_p)}</td>
                <td class="dc"></td>
                <td class="dc"></td>
            </tr>
            <tr>
                <td class="rh-sub">Razón de verosimilitud</td>
                <td class="dc">{fmt_val(lr_v)}</td>
                <td class="dc">1</td>
                <td class="dc">{fmt_sig(lr_p)}</td>
                <td class="dc"></td>
                <td class="dc"></td>
            </tr>
            <tr>
                <td class="rh-sub">Prueba exacta de Fisher</td>
                <td class="dc"></td>
                <td class="dc"></td>
                <td class="dc"></td>
                <td class="dc">{fmt_sig(f_two)}</td>
                <td class="dc">{fmt_sig(f_one)}</td>
            </tr>
            <tr>
                <td class="rh-sub">Asociación lineal por lineal</td>
                <td class="dc">{fmt_val(lbl_v)}</td>
                <td class="dc">1</td>
                <td class="dc">{fmt_sig(lbl_p)}</td>
                <td class="dc"></td>
                <td class="dc"></td>
            </tr>{mcn_row}
            <tr class="last-row">
                <td class="rh-sub">N de casos válidos</td>
                <td class="dc">{N_val}</td>
                <td class="dc"></td>
                <td class="dc"></td>
                <td class="dc"></td>
                <td class="dc"></td>
            </tr>
        </tbody>
    </table>
    <div class="spss-fn">
        a. {low_cells} casillas ({low_pct:.1f}%) han esperado un recuento menor que 5. El recuento mínimo esperado es {min_exp:.2f}.<br>
        b. Sólo se ha calculado para una tabla 2x2{fn_c}
    </div>
</div>
"""

def make_kappa_table(kappa_v, ase_v, t_v, sig_v, N_val):
    """
    Genera tabla SPSS 'Medidas simétricas' para Índice de Kappa de Cohen.
    """
    kv_str = f"-{fmt_val(abs(kappa_v))}" if kappa_v < 0 else fmt_val(kappa_v)
    return f"""
<div class="spss-wrap">
    <div class="spss-title">Medidas simétricas</div>
    <table class="spss-tbl">
        <thead>
            <tr class="bottom-hdr">
                <th colspan="2" class="no-bl" style="vertical-align: bottom;"></th>
                <th style="vertical-align: bottom;">Valor</th>
                <th style="vertical-align: bottom;">Error estándar<br>asintótico<sup><b>a</b></sup></th>
                <th style="vertical-align: bottom;">T aproximada<sup><b>b</b></sup></th>
                <th style="vertical-align: bottom;">Significación<br>aproximada</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="rh-sub">Medida de acuerdo</td>
                <td class="rh-sub">Kappa</td>
                <td class="dc">{kv_str}</td>
                <td class="dc">{fmt_val(ase_v)}</td>
                <td class="dc">{t_v:.3f}</td>
                <td class="dc">{fmt_sig(sig_v)}</td>
            </tr>
            <tr class="last-row">
                <td class="rh-sub" colspan="2">N de casos válidos</td>
                <td class="dc">{N_val}</td>
                <td class="dc"></td>
                <td class="dc"></td>
                <td class="dc"></td>
            </tr>
        </tbody>
    </table>
    <div class="spss-fn">
        a. No se presupone la hipótesis nula.<br>
        b. Utilización del error estándar asintótico que presupone la hipótesis nula.
    </div>
</div>
"""

tables = {}

# ==========================================
# EJERCICIO 1: Independencia 1 (Dispositivo vs SO)
# ==========================================
tables["ex1_crosstab.png"] = make_crosstab_3metric(
    "Tipo de dispositivo", "Sistema Operativo",
    ["Laptop", "Desktop", "Servidor"], ["Windows", "MacOS", "Linux"],
    [[40, 30, 20], [50, 20, 30], [20, 10, 80]],
    [[33.0, 18.0, 39.0], [36.7, 20.0, 43.3], [40.3, 22.0, 47.7]],
    [[44.4, 33.3, 22.2], [50.0, 20.0, 30.0], [18.2, 9.1, 72.7]],
    [110, 60, 130], [36.7, 20.0, 43.3], 300
)
tables["ex1_test.png"] = make_chisq_standard(66.421, 0.000, 67.298, 0.000, 39.024, 0.000, 4, 300, 0, 0.0, 18.00)

# ==========================================
# EJERCICIO 2: Independencia 2 (Lenguaje vs Condición)
# ==========================================
tables["ex2_crosstab.png"] = make_crosstab_3metric(
    "Lenguaje de programación", "Condición del proyecto",
    ["Python", "Java", "JavaScript"], ["Éxito", "Fracaso"],
    [[50, 25], [30, 40], [35, 20]],
    [[43.1, 31.9], [40.3, 29.8], [31.6, 23.4]],
    [[66.7, 33.3], [42.9, 57.1], [63.6, 36.4]],
    [115, 85], [57.5, 42.5], 200
)
tables["ex2_test.png"] = make_chisq_standard(9.568, 0.008, 9.555, 0.008, 0.390, 0.532, 2, 200, 0, 0.0, 23.38)

# ==========================================
# EJERCICIO 3: Homogeneidad 1 (Servidores vs Condición de fallos)
# ==========================================
tables["ex3_crosstab.png"] = make_crosstab_3metric(
    "Tipos de servidores", "Condición de fallos",
    ["Linux", "Windows", "MacOS"], ["Críticas", "No críticas"],
    [[20, 80], [30, 70], [10, 90]],
    [[20.0, 80.0], [20.0, 80.0], [20.0, 80.0]],
    [[20.0, 80.0], [30.0, 70.0], [10.0, 90.0]],
    [60, 240], [20.0, 80.0], 300
)
tables["ex3_test.png"] = make_chisq_standard(12.500, 0.002, 12.972, 0.002, 3.115, 0.078, 2, 300, 0, 0.0, 20.00)

# ==========================================
# EJERCICIO 5: McNemar 1 (Algoritmo de compresión Antes vs Después)
# ==========================================
tables["ex5_crosstab.png"] = make_crosstab_2metric_total(
    "Antes de la intervención", "Después de la intervención",
    ["Fracaso", "Éxito"], ["Fracaso", "Éxito"],
    [[15, 15], [5, 65]],
    [[15.0, 15.0], [5.0, 65.0]],
    [20, 80], [20.0, 80.0], 100
)
tables["ex5_test.png"] = make_chisq_2x2(24.107, 0.000, 21.503, 0.000, 22.467, 0.000, 0.000, 0.000, 23.866, 0.000, 100, 0, 0.0, 6.00, mcnemar_p=0.041)

# ==========================================
# EJERCICIO 7: Kappa 1 (Experto 1 vs Experto 2)
# ==========================================
tables["ex7_crosstab.png"] = make_crosstab_2metric_total(
    "Experto 1", "Experto 2",
    ["Intrusión detectada", "Intrusión no detectada"], ["Intrusión detectada", "Intrusión no detectada"],
    [[72, 13], [8, 7]],
    [[72.0, 13.0], [8.0, 7.0]],
    [80, 20], [80.0, 20.0], 100
)
tables["ex7_test.png"] = make_kappa_table(0.276, 0.118, 2.801, 0.005, 100)

# ==========================================
# EJERCICIO 8: Kappa 2 (Algoritmo A vs Algoritmo B)
# ==========================================
tables["ex8_crosstab.png"] = make_crosstab_2metric_total(
    "Algoritmo A", "Algoritmo B",
    ["Fallo identificado", "Fallo no identificado"], ["Fallo identificado", "Fallo no identificado"],
    [[8, 12], [2, 78]],
    [[8.0, 12.0], [2.0, 78.0]],
    [10, 90], [10.0, 90.0], 100
)
tables["ex8_test.png"] = make_kappa_table(0.462, 0.117, 5.000, 0.000, 100)

# ==========================================
# EJERCICIO 9: Fisher 1 (Sistema Antiguo vs Sistema Nuevo)
# ==========================================
tables["ex9_crosstab.png"] = make_crosstab_3metric(
    "Sistema Antiguo", "Sistema Nuevo",
    ["Éxito", "Fracaso"], ["Éxito", "Fracaso"],
    [[27, 3], [8, 2]],
    [[26.3, 3.8], [8.8, 1.3]],
    [[90.0, 10.0], [80.0, 20.0]],
    [35, 5], [87.5, 12.5], 40
)
tables["ex9_test.png"] = make_chisq_2x2(0.686, 0.408, 0.076, 0.783, 0.629, 0.428, 0.584, 0.367, 0.669, 0.414, 40, 2, 50.0, 1.25, mcnemar_p=None)

# ==========================================
# EJERCICIO 10: Fisher 2 (Sistema A vs Sistema B)
# ==========================================
tables["ex10_crosstab.png"] = make_crosstab_3metric(
    "Sistema A", "Sistema B",
    ["Contraseñas seguras", "Contraseñas inseguras"], ["Contraseñas seguras", "Contraseñas inseguras"],
    [[25, 10], [13, 2]],
    [[26.6, 8.4], [11.4, 3.6]],
    [[71.4, 28.6], [86.7, 13.3]],
    [38, 12], [76.0, 24.0], 50
)
tables["ex10_test.png"] = make_chisq_2x2(1.337, 0.248, 0.632, 0.427, 1.449, 0.229, 0.304, 0.217, 1.310, 0.252, 50, 1, 25.0, 3.60, mcnemar_p=None)

def render_all():
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge")
        page = browser.new_page(device_scale_factor=2)
        for fname, html_body in tables.items():
            full_html = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><style>{SPSS_LIGHT_CSS}</style></head>
<body>{html_body}</body>
</html>"""
            page.set_content(full_html)
            el = page.locator(".spss-wrap")
            el.screenshot(path=os.path.join(output_dir, fname))
            print(f"Generada imagen: {fname}")
        browser.close()

    # Copiar las 4 imágenes originales del docente para Ejercicio 4 y Ejercicio 6
    if os.path.exists("extracted_chi2_media/image1.png"):
        shutil.copy2("extracted_chi2_media/image1.png", os.path.join(output_dir, "ex4_crosstab.png"))
        shutil.copy2("extracted_chi2_media/image2.png", os.path.join(output_dir, "ex4_test.png"))
        shutil.copy2("extracted_chi2_media/image3.png", os.path.join(output_dir, "ex6_crosstab.png"))
        shutil.copy2("extracted_chi2_media/image4.png", os.path.join(output_dir, "ex6_test.png"))
        print("Copiadas imágenes originales del docente para Ejercicio 4 y Ejercicio 6.")

if __name__ == "__main__":
    render_all()

