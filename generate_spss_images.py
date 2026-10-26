import os
from playwright.sync_api import sync_playwright

output_dir = "generated_spss_images"
os.makedirs(output_dir, exist_ok=True)

# CSS for authentic IBM SPSS 26-29 table style
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

table.spss-table td.sub-header {
    background-color: #E8ECEF;
    color: #103B60;
    text-align: left;
}

.spss-footnote {
    font-size: 10px;
    color: #000000;
    margin-top: 5px;
    text-align: left;
}
"""

def render_html_to_image(html_content, output_path):
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
        page = browser.new_page(device_scale_factor=2) # 2x scale for crisp image
        page.set_content(full_html)
        # Select the container element
        container = page.locator(".spss-container")
        container.screenshot(path=output_path)
        browser.close()

# 1A: Chi-Square Frequencies - Sistemas Operativos
html_1a = """
<div class="spss-container">
    <div class="spss-box">
        <div class="spss-title">sistemas_operativos</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th>N observado</th>
                    <th>N esperado</th>
                    <th>Residual</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Windows</td>
                    <td>210</td>
                    <td>200.0</td>
                    <td>10.0</td>
                </tr>
                <tr>
                    <td class="row-header">Linux</td>
                    <td>170</td>
                    <td>175.0</td>
                    <td>-5.0</td>
                </tr>
                <tr>
                    <td class="row-header">macOS</td>
                    <td>120</td>
                    <td>125.0</td>
                    <td>-5.0</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Total</td>
                    <td>500</td>
                    <td></td>
                    <td></td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
"""

# 1B: Chi-Square Test Statistics - Sistemas Operativos
html_1b = """
<div class="spss-container">
    <div class="spss-nav-header">
        <span class="red-arrow">&#9658;</span> Pruebas NPar
    </div>
    <div style="clear:both;"></div>
    <div class="spss-box">
        <div class="spss-title">Estad&iacute;sticos de prueba</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th>sistemas_operativos</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Chi-cuadrado<sup>a</sup></td>
                    <td>.843</td>
                </tr>
                <tr>
                    <td class="row-header">gl</td>
                    <td>2</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Sig. asint&oacute;t.</td>
                    <td>.656</td>
                </tr>
            </tbody>
        </table>
        <div class="spss-footnote">
            a. 0 casillas (0.0%) tienen frecuencias esperadas menores que 5. La frecuencia de casilla esperada m&iacute;nima es 125.0.
        </div>
    </div>
</div>
"""

# 2A: Chi-Square Frequencies - Lenguajes de Programacion
html_2a = """
<div class="spss-container">
    <div class="spss-box">
        <div class="spss-title">lenguajes_programacion</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th>N observado</th>
                    <th>N esperado</th>
                    <th>Residual</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Python</td>
                    <td>200</td>
                    <td>200.0</td>
                    <td>.0</td>
                </tr>
                <tr>
                    <td class="row-header">Java</td>
                    <td>140</td>
                    <td>140.0</td>
                    <td>.0</td>
                </tr>
                <tr>
                    <td class="row-header">JavaScript</td>
                    <td>60</td>
                    <td>60.0</td>
                    <td>.0</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Total</td>
                    <td>400</td>
                    <td></td>
                    <td></td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
"""

# 2B: Chi-Square Test Statistics - Lenguajes de Programacion
html_2b = """
<div class="spss-container">
    <div class="spss-nav-header">
        <span class="red-arrow">&#9658;</span> Pruebas NPar
    </div>
    <div style="clear:both;"></div>
    <div class="spss-box">
        <div class="spss-title">Estad&iacute;sticos de prueba</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th>lenguajes_programacion</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Chi-cuadrado<sup>a</sup></td>
                    <td>.000</td>
                </tr>
                <tr>
                    <td class="row-header">gl</td>
                    <td>2</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Sig. asint&oacute;t.</td>
                    <td>1.000</td>
                </tr>
            </tbody>
        </table>
        <div class="spss-footnote">
            a. 0 casillas (0.0%) tienen frecuencias esperadas menores que 5. La frecuencia de casilla esperada m&iacute;nima es 60.0.
        </div>
    </div>
</div>
"""

# 6: Runs Test - Servidor de Correo (VAR00001 / spam_umbral)
html_6 = """
<div class="spss-container">
    <div class="spss-nav-header">
        <span class="red-arrow">&#9658;</span> Pruebas NPar
    </div>
    <div style="clear:both;"></div>
    <div class="spss-box">
        <div class="spss-title">Prueba de rachas</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th>VAR00001</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Valor de prueba<sup>a</sup></td>
                    <td>.63</td>
                </tr>
                <tr>
                    <td class="row-header">Casos &lt; Valor de prueba</td>
                    <td>11</td>
                </tr>
                <tr>
                    <td class="row-header">Casos &gt;= Valor de prueba</td>
                    <td>19</td>
                </tr>
                <tr>
                    <td class="row-header">Casos totales</td>
                    <td>30</td>
                </tr>
                <tr>
                    <td class="row-header">N&uacute;mero de rachas</td>
                    <td>19</td>
                </tr>
                <tr>
                    <td class="row-header">Z</td>
                    <td>1.431</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Sig. asin. (bilateral)</td>
                    <td>.152</td>
                </tr>
            </tbody>
        </table>
        <div class="spss-footnote">
            a. Media
        </div>
    </div>
</div>
"""

render_html_to_image(html_1a, os.path.join(output_dir, "spss_chi_freq_so.png"))
render_html_to_image(html_1b, os.path.join(output_dir, "spss_chi_test_so.png"))
render_html_to_image(html_2a, os.path.join(output_dir, "spss_chi_freq_lenguajes.png"))
render_html_to_image(html_2b, os.path.join(output_dir, "spss_chi_test_lenguajes.png"))
render_html_to_image(html_6, os.path.join(output_dir, "spss_rachas_correo.png"))

print("All SPSS images rendered successfully!")
