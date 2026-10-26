import os
from playwright.sync_api import sync_playwright

output_dir = "generated_spss_mann_images"
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
    font-size: 12px;
    font-weight: bold;
    color: #000000;
    margin-bottom: 8px;
    background: #ffffff;
}

.red-arrow {
    color: #c00000;
    margin-right: 5px;
    font-size: 10px;
}

.spss-box {
    border: 1px solid #999999;
    padding: 10px 14px;
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
}
"""

def render_html(html_content, output_path):
    full_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>{SPSS_CSS}</style>
</head>
<body>
{html_content}
</body>
</html>"""
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge")
        page = browser.new_page(device_scale_factor=2)
        page.set_content(full_html)
        container = page.locator(".spss-container")
        container.screenshot(path=output_path)
        browser.close()

# --- EXERCISE 1 (Java vs Python) ---
html_ex1_rangos = """
<div class="spss-container">
    <div class="spss-box">
        <div class="spss-title">Rangos</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;">Lenguaje</th>
                    <th>N</th>
                    <th>Rango promedio</th>
                    <th>Suma de rangos</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Java</td>
                    <td>8</td>
                    <td>9.31</td>
                    <td>74.50</td>
                </tr>
                <tr>
                    <td class="row-header">Python</td>
                    <td>5</td>
                    <td>3.30</td>
                    <td>16.50</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Total</td>
                    <td>13</td>
                    <td></td>
                    <td></td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
"""

html_ex1_test = """
<div class="spss-container">
    <div class="spss-nav-header">
        <span class="red-arrow">&#9658;</span> Pruebas NPar
    </div>
    <div style="clear:both;"></div>
    <div class="spss-box">
        <div class="spss-title">Estad&iacute;sticos de prueba<sup>a</sup></div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th>tiempo_ejecucion</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">U de Mann-Whitney</td>
                    <td>1.500</td>
                </tr>
                <tr>
                    <td class="row-header">W de Wilcoxon</td>
                    <td>16.500</td>
                </tr>
                <tr>
                    <td class="row-header">Z</td>
                    <td>-2.639</td>
                </tr>
                <tr>
                    <td class="row-header">Sig. asint&oacute;t. (bilateral)</td>
                    <td>.008</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Sig. exacta [2*(Sig. unilateral)]</td>
                    <td>.006<sup>b</sup></td>
                </tr>
            </tbody>
        </table>
        <div class="spss-footnote">
            a. Variable de agrupaci&oacute;n: Lenguaje<br>
            b. No corregido para empates.
        </div>
    </div>
</div>
"""

# --- EXERCISE 2 (Sistema A vs Sistema B) ---
html_ex2_rangos = """
<div class="spss-container">
    <div class="spss-box">
        <div class="spss-title">Rangos</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;">Sistema</th>
                    <th>N</th>
                    <th>Rango promedio</th>
                    <th>Suma de rangos</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Sistema A</td>
                    <td>10</td>
                    <td>5.50</td>
                    <td>55.00</td>
                </tr>
                <tr>
                    <td class="row-header">Sistema B</td>
                    <td>7</td>
                    <td>14.00</td>
                    <td>98.00</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Total</td>
                    <td>17</td>
                    <td></td>
                    <td></td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
"""

html_ex2_test = """
<div class="spss-container">
    <div class="spss-nav-header">
        <span class="red-arrow">&#9658;</span> Pruebas NPar
    </div>
    <div style="clear:both;"></div>
    <div class="spss-box">
        <div class="spss-title">Estad&iacute;sticos de prueba<sup>a</sup></div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th>tiempo_acceso</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">U de Mann-Whitney</td>
                    <td>.000</td>
                </tr>
                <tr>
                    <td class="row-header">W de Wilcoxon</td>
                    <td>55.000</td>
                </tr>
                <tr>
                    <td class="row-header">Z</td>
                    <td>-3.371</td>
                </tr>
                <tr>
                    <td class="row-header">Sig. asint&oacute;t. (bilateral)</td>
                    <td>.001</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Sig. exacta [2*(Sig. unilateral)]</td>
                    <td>.000<sup>b</sup></td>
                </tr>
            </tbody>
        </table>
        <div class="spss-footnote">
            a. Variable de agrupaci&oacute;n: Sistema<br>
            b. No corregido para empates.
        </div>
    </div>
</div>
"""

# --- EXERCISE 3 (Servidor A vs Servidor B) ---
html_ex3_rangos = """
<div class="spss-container">
    <div class="spss-box">
        <div class="spss-title">Rangos</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;">Servidor</th>
                    <th>N</th>
                    <th>Rango promedio</th>
                    <th>Suma de rangos</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Servidor A</td>
                    <td>20</td>
                    <td>11.00</td>
                    <td>220.00</td>
                </tr>
                <tr>
                    <td class="row-header">Servidor B</td>
                    <td>18</td>
                    <td>28.94</td>
                    <td>521.00</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Total</td>
                    <td>38</td>
                    <td></td>
                    <td></td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
"""

html_ex3_test = """
<div class="spss-container">
    <div class="spss-nav-header">
        <span class="red-arrow">&#9658;</span> Pruebas NPar
    </div>
    <div style="clear:both;"></div>
    <div class="spss-box">
        <div class="spss-title">Estad&iacute;sticos de prueba<sup>a</sup></div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th>tiempo_procesamiento</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">U de Mann-Whitney</td>
                    <td>10.000</td>
                </tr>
                <tr>
                    <td class="row-header">W de Wilcoxon</td>
                    <td>220.000</td>
                </tr>
                <tr>
                    <td class="row-header">Z</td>
                    <td>-5.035</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Sig. asint&oacute;t. (bilateral)</td>
                    <td>.000</td>
                </tr>
            </tbody>
        </table>
        <div class="spss-footnote">
            a. Variable de agrupaci&oacute;n: Servidor
        </div>
    </div>
</div>
"""

# --- EXERCISE 4 (Algoritmo A vs Algoritmo B) ---
html_ex4_rangos = """
<div class="spss-container">
    <div class="spss-box">
        <div class="spss-title">Rangos</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;">Algoritmo</th>
                    <th>N</th>
                    <th>Rango promedio</th>
                    <th>Suma de rangos</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Algoritmo A</td>
                    <td>20</td>
                    <td>11.45</td>
                    <td>229.00</td>
                </tr>
                <tr>
                    <td class="row-header">Algoritmo B</td>
                    <td>15</td>
                    <td>26.73</td>
                    <td>401.00</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Total</td>
                    <td>35</td>
                    <td></td>
                    <td></td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
"""

html_ex4_test = """
<div class="spss-container">
    <div class="spss-nav-header">
        <span class="red-arrow">&#9658;</span> Pruebas NPar
    </div>
    <div style="clear:both;"></div>
    <div class="spss-box">
        <div class="spss-title">Estad&iacute;sticos de prueba<sup>a</sup></div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th>tasa_compresion</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">U de Mann-Whitney</td>
                    <td>19.000</td>
                </tr>
                <tr>
                    <td class="row-header">W de Wilcoxon</td>
                    <td>229.000</td>
                </tr>
                <tr>
                    <td class="row-header">Z</td>
                    <td>-4.360</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Sig. asint&oacute;t. (bilateral)</td>
                    <td>.000</td>
                </tr>
            </tbody>
        </table>
        <div class="spss-footnote">
            a. Variable de agrupaci&oacute;n: Algoritmo
        </div>
    </div>
</div>
"""

render_html(html_ex1_rangos, os.path.join(output_dir, "spss_mann_rangos_ex1.png"))
render_html(html_ex1_test, os.path.join(output_dir, "spss_mann_test_ex1.png"))
render_html(html_ex2_rangos, os.path.join(output_dir, "spss_mann_rangos_ex2.png"))
render_html(html_ex2_test, os.path.join(output_dir, "spss_mann_test_ex2.png"))
render_html(html_ex3_rangos, os.path.join(output_dir, "spss_mann_rangos_ex3.png"))
render_html(html_ex3_test, os.path.join(output_dir, "spss_mann_test_ex3.png"))
render_html(html_ex4_rangos, os.path.join(output_dir, "spss_mann_rangos_ex4.png"))
render_html(html_ex4_test, os.path.join(output_dir, "spss_mann_test_ex4.png"))

print("All Mann-Whitney SPSS images rendered successfully!")
