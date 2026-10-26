import os
from playwright.sync_api import sync_playwright

output_dir = "generated_kruskal_images"
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
    padding: 3px 12px;
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
    padding: 3px 12px;
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
    margin-top: 6px;
    text-align: left;
    line-height: 1.3;
}
"""

def render(html_content, fname):
    full_html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>{SPSS_CSS}</style></head>
<body>{html_content}</body></html>"""
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge")
        page = browser.new_page(device_scale_factor=2)
        page.set_content(full_html)
        container = page.locator(".spss-container")
        container.screenshot(path=os.path.join(output_dir, fname))
        browser.close()
    print("Generated", fname)

# --- EXERCISE 3 IMAGES ---
# 1. Informe descriptivos
h_ex3_desc = """
<div class="spss-container">
    <div class="spss-box">
        <div class="spss-title">Informe</div>
        <div style="font-size: 10px; color: #103B60; margin-bottom: 4px; text-align: left;">Valores</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;">APLICACIONES</th>
                    <th>Media</th>
                    <th>N</th>
                    <th>Desv. estándar</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Aplicación 1</td>
                    <td>321.43</td>
                    <td>7</td>
                    <td>42.790</td>
                </tr>
                <tr>
                    <td class="row-header">Aplicación 2</td>
                    <td>412.50</td>
                    <td>4</td>
                    <td>6.455</td>
                </tr>
                <tr>
                    <td class="row-header">Aplicación 3</td>
                    <td>286.33</td>
                    <td>6</td>
                    <td>31.443</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Total</td>
                    <td>330.47</td>
                    <td>17</td>
                    <td>58.753</td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
"""
render(h_ex3_desc, "ex3_informe.png")

# 2. Normalidad Shapiro-Wilk
h_ex3_norm = """
<div class="spss-container">
    <div style="display: flex; align-items: center;">
        <span class="red-arrow" style="margin-right: 6px;">&#9658;</span>
        <div class="spss-box">
            <div class="spss-title">Pruebas de normalidad</div>
            <table class="spss-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="background-color: #ffffff; vertical-align: bottom;">APLICACIONES</th>
                        <th colspan="3" style="text-align: center;">Shapiro-Wilk</th>
                    </tr>
                    <tr>
                        <th>Estadístico</th>
                        <th>gl</th>
                        <th>Sig.</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="row-header">Aplicación 1</td>
                        <td>.871</td>
                        <td>7</td>
                        <td>.189</td>
                    </tr>
                    <tr>
                        <td class="row-header">Aplicación 2</td>
                        <td>.993</td>
                        <td>4</td>
                        <td>.972</td>
                    </tr>
                    <tr class="total-row">
                        <td class="row-header">Aplicación 3</td>
                        <td>.759</td>
                        <td>6</td>
                        <td>.025</td>
                    </tr>
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
render(h_ex3_norm, "ex3_normalidad.png")

# 3. Rangos Kruskal-Wallis
h_ex3_rangos = """
<div class="spss-container">
    <div style="font-size: 11px; font-weight: bold; margin-bottom: 4px;">Prueba de Kruskal-Wallis</div>
    <div class="spss-box">
        <div class="spss-title">Rangos</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;">APLICACIONES</th>
                    <th>N</th>
                    <th>Rango promedio</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Aplicación 1</td>
                    <td>7</td>
                    <td>8.71</td>
                </tr>
                <tr>
                    <td class="row-header">Aplicación 2</td>
                    <td>4</td>
                    <td>15.50</td>
                </tr>
                <tr>
                    <td class="row-header">Aplicación 3</td>
                    <td>6</td>
                    <td>5.00</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Total</td>
                    <td>17</td>
                    <td></td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
"""
render(h_ex3_rangos, "ex3_rangos.png")

# 4. Estadisticos de prueba Kruskal-Wallis
h_ex3_test = """
<div class="spss-container">
    <div class="spss-box">
        <div class="spss-title">Estadísticos de prueba<sup>a,b</sup></div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th>Valores</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">H de Kruskal-Wallis</td>
                    <td>10.415</td>
                </tr>
                <tr>
                    <td class="row-header">gl</td>
                    <td>2</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Sig. asin.</td>
                    <td>.005</td>
                </tr>
            </tbody>
        </table>
        <div class="spss-footnote">
            a. Prueba de Kruskal Wallis<br>
            b. Variable de agrupación: APLICACIONES
        </div>
    </div>
</div>
"""
render(h_ex3_test, "ex3_test.png")

# --- EXERCISE 4 IMAGES ---
# 1. Informe descriptivos
h_ex4_desc = """
<div class="spss-container">
    <div class="spss-box">
        <div class="spss-title">Informe</div>
        <div style="font-size: 10px; color: #103B60; margin-bottom: 4px; text-align: left;">Valores</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;">REDES</th>
                    <th>Media</th>
                    <th>N</th>
                    <th>Desv. estándar</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Red X</td>
                    <td>31.60</td>
                    <td>5</td>
                    <td>2.702</td>
                </tr>
                <tr>
                    <td class="row-header">Red Y</td>
                    <td>47.50</td>
                    <td>6</td>
                    <td>1.871</td>
                </tr>
                <tr>
                    <td class="row-header">Red Z</td>
                    <td>25.20</td>
                    <td>5</td>
                    <td>7.396</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Total</td>
                    <td>35.56</td>
                    <td>16</td>
                    <td>10.758</td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
"""
render(h_ex4_desc, "ex4_informe.png")

# 2. Normalidad Shapiro-Wilk
h_ex4_norm = """
<div class="spss-container">
    <div style="display: flex; align-items: center;">
        <span class="red-arrow" style="margin-right: 6px;">&#9658;</span>
        <div class="spss-box">
            <div class="spss-title">Pruebas de normalidad</div>
            <table class="spss-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="background-color: #ffffff; vertical-align: bottom;">REDES</th>
                        <th colspan="3" style="text-align: center;">Shapiro-Wilk</th>
                    </tr>
                    <tr>
                        <th>Estadístico</th>
                        <th>gl</th>
                        <th>Sig.</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="row-header">Red X</td>
                        <td>.990</td>
                        <td>5</td>
                        <td>.980</td>
                    </tr>
                    <tr>
                        <td class="row-header">Red Y</td>
                        <td>.982</td>
                        <td>6</td>
                        <td>.961</td>
                    </tr>
                    <tr class="total-row">
                        <td class="row-header">Red Z</td>
                        <td>.764</td>
                        <td>5</td>
                        <td>.040</td>
                    </tr>
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
render(h_ex4_norm, "ex4_normalidad.png")

# 3. Rangos Kruskal-Wallis
h_ex4_rangos = """
<div class="spss-container">
    <div style="font-size: 11px; font-weight: bold; margin-bottom: 4px;">Prueba de Kruskal-Wallis</div>
    <div class="spss-box">
        <div class="spss-title">Rangos</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;">REDES</th>
                    <th>N</th>
                    <th>Rango promedio</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Red X</td>
                    <td>5</td>
                    <td>7.00</td>
                </tr>
                <tr>
                    <td class="row-header">Red Y</td>
                    <td>6</td>
                    <td>13.50</td>
                </tr>
                <tr>
                    <td class="row-header">Red Z</td>
                    <td>5</td>
                    <td>4.00</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Total</td>
                    <td>16</td>
                    <td></td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
"""
render(h_ex4_rangos, "ex4_rangos.png")

# 4. Estadisticos de prueba Kruskal-Wallis
h_ex4_test = """
<div class="spss-container">
    <div class="spss-box">
        <div class="spss-title">Estadísticos de prueba<sup>a,b</sup></div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th>Valores</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">H de Kruskal-Wallis</td>
                    <td>11.581</td>
                </tr>
                <tr>
                    <td class="row-header">gl</td>
                    <td>2</td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Sig. asin.</td>
                    <td>.003</td>
                </tr>
            </tbody>
        </table>
        <div class="spss-footnote">
            a. Prueba de Kruskal Wallis<br>
            b. Variable de agrupación: REDES
        </div>
    </div>
</div>
"""
render(h_ex4_test, "ex4_test.png")

# 5. ANOVA and Tukey for Ex4 (optional alternative)
h_ex4_levene = """
<div class="spss-container">
    <div style="display: flex; align-items: center;">
        <span class="red-arrow" style="margin-right: 6px;">&#9658;</span>
        <div class="spss-box">
            <div class="spss-title">Prueba de homogeneidad de varianza</div>
            <table class="spss-table">
                <thead>
                    <tr>
                        <th style="background-color: #ffffff;"></th>
                        <th>Estadístico de Levene</th>
                        <th>gl1</th>
                        <th>gl2</th>
                        <th>Sig.</th>
                    </tr>
                </thead>
                <tbody>
                    <tr class="total-row">
                        <td class="row-header">Latencia Se basa en la media</td>
                        <td>2.595</td>
                        <td>2</td>
                        <td>13</td>
                        <td>.113</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>
</div>
"""
render(h_ex4_levene, "ex4_levene.png")

h_ex4_anova = """
<div class="spss-container">
    <div class="spss-box">
        <div class="spss-title">ANOVA</div>
        <div style="font-size: 10px; color: #103B60; margin-bottom: 4px; text-align: left;">Latencia</div>
        <table class="spss-table">
            <thead>
                <tr>
                    <th style="background-color: #ffffff;"></th>
                    <th>Suma de cuadrados</th>
                    <th>gl</th>
                    <th>Media cuadrática</th>
                    <th>F</th>
                    <th>Sig.</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="row-header">Entre grupos</td>
                    <td>1470.438</td>
                    <td>2</td>
                    <td>735.219</td>
                    <td>35.999</td>
                    <td>.000</td>
                </tr>
                <tr>
                    <td class="row-header">Dentro de grupos</td>
                    <td>265.500</td>
                    <td>13</td>
                    <td>20.423</td>
                    <td></td>
                    <td></td>
                </tr>
                <tr class="total-row">
                    <td class="row-header">Total</td>
                    <td>1735.938</td>
                    <td>15</td>
                    <td></td>
                    <td></td>
                    <td></td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
"""
render(h_ex4_anova, "ex4_anova.png")
