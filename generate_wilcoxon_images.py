# -*- coding: utf-8 -*-
"""
Generador de imágenes SPSS oficiales para Ejercicios de Prueba de Wilcoxon
Utiliza Playwright y CSS fiel a IBM SPSS Statistics 26-29.
"""

import os
from playwright.sync_api import sync_playwright

output_dir = "generated_wilcoxon_images"
os.makedirs(output_dir, exist_ok=True)

SPSS_CSS = """
body {
    font-family: Arial, "Helvetica Neue", Helvetica, sans-serif;
    background-color: transparent;
    margin: 0;
    padding: 10px;
    display: inline-block;
}

.spss-container {
    background-color: #ffffff;
    display: inline-block;
    padding: 2px;
}

.spss-nav-header {
    display: inline-flex;
    align-items: center;
    border: 1px solid #555555;
    padding: 1px 6px;
    font-size: 11px;
    font-weight: bold;
    color: #000000;
    margin-bottom: 6px;
    background: #ffffff;
}

.red-arrow {
    color: #c00000;
    margin-right: 5px;
    font-size: 11px;
}

.spss-box {
    border: 1px solid #999999;
    padding: 8px 12px;
    background-color: #ffffff;
    display: inline-block;
}

.spss-title {
    text-align: center;
    font-size: 12px;
    font-weight: bold;
    color: #000000;
    margin-bottom: 8px;
}

table.spss-table {
    border-collapse: collapse;
    font-size: 11px;
    color: #000000;
    margin: 0 auto;
}

table.spss-table th {
    font-weight: normal;
    color: #103B60;
    padding: 3px 10px;
    text-align: right;
    border-top: 1px solid #103B60;
    border-bottom: 1px solid #103B60;
    border-left: 1px solid #E0E4E8;
    background-color: #ffffff;
}

table.spss-table th:first-child {
    border-left: none;
    text-align: left;
}

table.spss-table td {
    padding: 3px 10px;
    border-bottom: 1px solid #DDE2E6;
    text-align: right;
}

table.spss-table tr.total-row td {
    border-bottom: 1px solid #103B60;
}

table.spss-table td.row-header {
    background-color: #E8ECEF;
    color: #103B60;
    text-align: left;
    border-right: 1px solid #DDE2E6;
}

.spss-footnote {
    font-size: 10px;
    color: #000000;
    margin-top: 5px;
    text-align: left;
    line-height: 1.35;
}
"""

def render_all(tables_dict):
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge")
        page = browser.new_page(device_scale_factor=2)
        for fname, html_body in tables_dict.items():
            full_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>{SPSS_CSS}</style>
</head>
<body>
{html_body}
</body>
</html>"""
            page.set_content(full_html)
            container = page.locator(".spss-container")
            container.screenshot(path=os.path.join(output_dir, fname))
            print(f"Generada imagen: {fname}")
        browser.close()

def make_informe(group_col_title, rows):
    """
    rows: list of (label, media, n, sd) including Total row at the end
    """
    tr_html = ""
    for i, (label, media, n, sd) in enumerate(rows):
        cls = ' class="total-row"' if i == len(rows) - 1 else ""
        tr_html += f"""
                <tr{cls}>
                    <td class="row-header">{label}</td>
                    <td>{media}</td>
                    <td>{n}</td>
                    <td>{sd}</td>
                </tr>"""
    return f"""
<div class="spss-container">
    <div class="spss-box">
        <div class="spss-title">Informe</div>
        <div style="font-size: 10px; color: #103B60; margin-bottom: 4px; text-align: left;">Valores</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;">{group_col_title}</th>
                    <th>Media</th>
                    <th>N</th>
                    <th>Desv. estándar</th>
                </tr>
            </thead>
            <tbody>{tr_html}
            </tbody>
        </table>
    </div>
</div>
"""

def make_normalidad(rows):
    """
    rows: list of (label, sw_stat, gl, sw_sig)
    """
    tr_html = ""
    for i, (label, sw_stat, gl, sw_sig) in enumerate(rows):
        cls = ' class="total-row"' if i == len(rows) - 1 else ""
        tr_html += f"""
                    <tr{cls}>
                        <td class="row-header">{label}</td>
                        <td>{sw_stat}</td>
                        <td>{gl}</td>
                        <td>{sw_sig}</td>
                    </tr>"""
    return f"""
<div class="spss-container">
    <div style="display: flex; align-items: center;">
        <span class="red-arrow" style="margin-right: 6px;">&#9658;</span>
        <div class="spss-box">
            <div class="spss-title">Pruebas de normalidad</div>
            <table class="spss-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="background-color: #ffffff; vertical-align: bottom;"></th>
                        <th colspan="3" style="text-align: center;">Shapiro-Wilk</th>
                    </tr>
                    <tr>
                        <th>Estadístico</th>
                        <th>gl</th>
                        <th>Sig.</th>
                    </tr>
                </thead>
                <tbody>{tr_html}
                </tbody>
            </table>
            <div class="spss-footnote">
                *. Esto es un límite inferior de la significación verdadera.<br>
                a. Corrección de significación de Lilliefors
            </div>
        </div>
    </div>
</div>
"""

def make_wilcoxon_rangos(pair_label, var_before, var_after, n_neg, mr_neg, sr_neg, n_pos, mr_pos, sr_pos, n_ties, n_tot):
    return f"""
<div class="spss-container">
    <div style="font-size: 11px; font-weight: bold; margin-bottom: 4px;">Prueba de rangos con signo de Wilcoxon</div>
    <div class="spss-box">
        <div class="spss-title">Rangos</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th colspan="2" style="background-color: #ffffff;"></th>
                    <th>N</th>
                    <th>Rango promedio</th>
                    <th>Suma de rangos</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td rowspan="4" class="row-header" style="vertical-align: top; border-bottom: 1px solid #103B60;">{pair_label}</td>
                    <td class="row-header">Rangos negativos</td>
                    <td>{n_neg}<sup>a</sup></td>
                    <td>{mr_neg}</td>
                    <td>{sr_neg}</td>
                </tr>
                <tr>
                    <td class="row-header">Rangos positivos</td>
                    <td>{n_pos}<sup>b</sup></td>
                    <td>{mr_pos}</td>
                    <td>{sr_pos}</td>
                </tr>
                <tr>
                    <td class="row-header">Empates</td>
                    <td>{n_ties}<sup>c</sup></td>
                    <td></td>
                    <td></td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Total</td>
                    <td>{n_tot}</td>
                    <td></td>
                    <td></td>
                </tr>
            </tbody>
        </table>
        <div class="spss-footnote">
            a. {var_after} &lt; {var_before}<br>
            b. {var_after} &gt; {var_before}<br>
            c. {var_after} = {var_before}
        </div>
    </div>
</div>
"""

def make_wilcoxon_test(pair_label, z_val, sig_val, basis_note):
    return f"""
<div class="spss-container">
    <div class="spss-nav-header">
        <span class="red-arrow">&#9658;</span> Pruebas NPar
    </div>
    <div style="clear:both;"></div>
    <div class="spss-box">
        <div class="spss-title">Estadísticos de prueba<sup>a</sup></div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th style="text-align: center;">{pair_label}</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Z</td>
                    <td>{z_val}<sup>b</sup></td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Sig. asintótica(bilateral)</td>
                    <td>{sig_val}</td>
                </tr>
            </tbody>
        </table>
        <div class="spss-footnote">
            a. Prueba de rangos con signo de Wilcoxon<br>
            b. {basis_note}
        </div>
    </div>
</div>
"""

tables = {}

# ==========================================
# EJERCICIO 1
# ==========================================
tables["ex1_informe.png"] = make_informe("OPTIMIZACIÓN", [
    ("Antes de la optimización", "16.30", "10", "2.669"),
    ("Después de la optimización", "11.20", "10", "1.033"),
    ("Total", "13.75", "20", "3.275")
])
tables["ex1_normalidad.png"] = make_normalidad([
    ("Tiempo de ejecución antes de la optimización", ".780", "10", ".008"),
    ("Tiempo de ejecución después de la optimización", ".895", "10", ".191"),
    ("Diferencia (Después - Antes)", ".824", "10", ".028")
])
tables["ex1_rangos.png"] = make_wilcoxon_rangos(
    "Tiempo de ejecución después de la optimización -<br>Tiempo de ejecución antes de la optimización",
    "Tiempo de ejecución antes de la optimización",
    "Tiempo de ejecución después de la optimización",
    "10", "5.50", "55.00",
    "0", ".00", ".00",
    "0", "10"
)
tables["ex1_test.png"] = make_wilcoxon_test(
    "Tiempo de ejecución después de la optimización -<br>Tiempo de ejecución antes de la optimización",
    "-2.814", ".005",
    "Se basa en rangos positivos."
)

# ==========================================
# EJERCICIO 2
# ==========================================
tables["ex2_informe.png"] = make_informe("ACTUALIZACIÓN", [
    ("Antes de la Actualización", "78.10", "10", "3.213"),
    ("Después de la Actualización", "69.10", "10", "1.197"),
    ("Total", "73.60", "20", "5.185")
])
tables["ex2_normalidad.png"] = make_normalidad([
    ("Tiempos de Respuesta Antes de la Actualización", ".914", "10", ".308"),
    ("Tiempos de Respuesta Después de la Actualización", ".952", "10", ".691"),
    ("Diferencia (Después - Antes)", ".812", "10", ".020")
])
tables["ex2_rangos.png"] = make_wilcoxon_rangos(
    "Tiempos de Respuesta Después de la Actualización -<br>Tiempos de Respuesta Antes de la Actualización",
    "Tiempos de Respuesta Antes de la Actualización",
    "Tiempos de Respuesta Después de la Actualización",
    "10", "5.50", "55.00",
    "0", ".00", ".00",
    "0", "10"
)
tables["ex2_test.png"] = make_wilcoxon_test(
    "Tiempos de Respuesta Después de la Actualización -<br>Tiempos de Respuesta Antes de la Actualización",
    "-2.814", ".005",
    "Se basa en rangos positivos."
)

# ==========================================
# EJERCICIO 3
# ==========================================
tables["ex3_informe.png"] = make_informe("PROTOCOLO", [
    ("Antes de la implementación", "155.86", "7", "3.716"),
    ("Después de la implementación", "144.00", "7", "4.320"),
    ("Total", "149.93", "14", "7.269")
])
tables["ex3_normalidad.png"] = make_normalidad([
    ("Antes de la implementación", ".929", "7", ".546"),
    ("Después de la implementación", ".822", "7", ".067"),
    ("Diferencia (Después - Antes)", ".803", "7", ".044")
])
tables["ex3_rangos.png"] = make_wilcoxon_rangos(
    "Después de la implementación -<br>Antes de la implementación",
    "Antes de la implementación",
    "Después de la implementación",
    "6", "4.50", "27.00",
    "1", "1.00", "1.00",
    "0", "7"
)
tables["ex3_test.png"] = make_wilcoxon_test(
    "Después de la implementación -<br>Antes de la implementación",
    "-2.201", ".028",
    "Se basa en rangos positivos."
)

# ==========================================
# EJERCICIO 4
# ==========================================
tables["ex4_informe.png"] = make_informe("MEJORA", [
    ("Antes de la mejora", "79.43", "7", "3.505"),
    ("Después de la mejora", "90.86", "7", "1.345"),
    ("Total", "85.14", "14", "6.455")
])
tables["ex4_normalidad.png"] = make_normalidad([
    ("Antes de la mejora", ".784", "7", ".028"),
    ("Después de la mejora", ".967", "7", ".873"),
    ("Diferencia (Después - Antes)", ".639", "7", ".001")
])
tables["ex4_rangos.png"] = make_wilcoxon_rangos(
    "Después de la mejora -<br>Antes de la mejora",
    "Antes de la mejora",
    "Después de la mejora",
    "0", ".00", ".00",
    "7", "4.00", "28.00",
    "0", "7"
)
tables["ex4_test.png"] = make_wilcoxon_test(
    "Después de la mejora -<br>Antes de la mejora",
    "-2.388", ".017",
    "Se basa en rangos negativos."
)

if __name__ == "__main__":
    render_all(tables)
    print("Todas las imágenes SPSS de Wilcoxon fueron generadas exitosamente.")
