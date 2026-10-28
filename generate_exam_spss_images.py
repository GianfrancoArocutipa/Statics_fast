# -*- coding: utf-8 -*-
"""
Generador de imágenes SPSS oficiales para la Segunda Práctica Calificada
Utiliza Playwright y CSS fiel a IBM SPSS Statistics 26-29.
"""

import os
from playwright.sync_api import sync_playwright

output_dir = os.path.join("prueba", "spss_images")
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

def wrap_html(nav_title, title, table_html, footnote=""):
    fn_html = f'<div class="spss-footnote">{footnote}</div>' if footnote else ''
    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>{SPSS_CSS}</style>
</head>
<body>
<div class="spss-container" id="target">
    <div class="spss-nav-header"><span class="red-arrow">&#9658;</span>{nav_title}</div>
    <div class="spss-box">
        <div class="spss-title">{title}</div>
        {table_html}
        {fn_html}
    </div>
</div>
</body>
</html>"""

def render_images():
    tables = {}

    # ==========================================
    # PROBLEMA 1: CHI-CUADRADO DE BONDAD DE AJUSTE
    # ==========================================
    # Tabla 1.1: Frecuencias observadas y esperadas
    t1_frec = """
    <table class="spss-table">
        <thead>
            <tr>
                <th style="text-align:left;">Categoría de incidente</th>
                <th>N observado</th>
                <th>N esperado</th>
                <th>Residual</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header">Errores de autenticación</td>
                <td>42</td>
                <td>30,0</td>
                <td>12,0</td>
            </tr>
            <tr>
                <td class="row-header">Fallas de conexión</td>
                <td>25</td>
                <td>30,0</td>
                <td>-5,0</td>
            </tr>
            <tr>
                <td class="row-header">Problemas de rendimiento</td>
                <td>33</td>
                <td>30,0</td>
                <td>3,0</td>
            </tr>
            <tr>
                <td class="row-header">Errores de interfaz</td>
                <td>20</td>
                <td>30,0</td>
                <td>-10,0</td>
            </tr>
            <tr class="total-row">
                <td class="row-header">Total</td>
                <td>120</td>
                <td></td>
                <td></td>
            </tr>
        </tbody>
    </table>
    """
    tables["p1_frecuencias.png"] = wrap_html(
        "Prueba de chi-cuadrado",
        "Categorías de incidentes reportados",
        t1_frec
    )

    # Tabla 1.2: Estadísticos de prueba
    t1_stat = """
    <table class="spss-table">
        <thead>
            <tr>
                <th style="text-align:left;"></th>
                <th>Categoría de incidente</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header">Chi-cuadrado</td>
                <td>9,267<sup>a</sup></td>
            </tr>
            <tr>
                <td class="row-header">gl</td>
                <td>3</td>
            </tr>
            <tr class="total-row">
                <td class="row-header">Sig. asintótica</td>
                <td>,026</td>
            </tr>
        </tbody>
    </table>
    """
    tables["p1_estadisticos_prueba.png"] = wrap_html(
        "Prueba de chi-cuadrado",
        "Estadísticos de prueba",
        t1_stat,
        "a. 0 casillas (0,0%) han esperado frecuencias menores que 5. La frecuencia de casilla esperada mínima es 30,0."
    )

    # ==========================================
    # PROBLEMA 2: COMPARACIÓN DE 2 ALGORITMOS
    # ==========================================
    # Tabla 2.1: Informe descriptivo
    t2_inf = """
    <table class="spss-table">
        <thead>
            <tr>
                <th style="text-align:left;">Algoritmo</th>
                <th>Media</th>
                <th>N</th>
                <th>Desv. Desviación</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header">Algoritmo A</td>
                <td>132,20</td>
                <td>10</td>
                <td>7,829</td>
            </tr>
            <tr>
                <td class="row-header">Algoritmo B</td>
                <td>116,27</td>
                <td>11</td>
                <td>5,159</td>
            </tr>
            <tr class="total-row">
                <td class="row-header">Total</td>
                <td>123,86</td>
                <td>21</td>
                <td>10,370</td>
            </tr>
        </tbody>
    </table>
    """
    tables["p2_informe.png"] = wrap_html(
        "Medias",
        "Informe<br><span style='font-weight:normal; font-size:11px;'>Tiempo de respuesta (ms)</span>",
        t2_inf
    )

    # Tabla 2.2: Normalidad
    t2_norm = """
    <table class="spss-table">
        <thead>
            <tr>
                <th rowspan="2" style="text-align:left; vertical-align:middle;">Algoritmo</th>
                <th colspan="3" style="text-align:center; border-left:1px solid #E0E4E8;">Kolmogorov-Smirnov<sup>a</sup></th>
                <th colspan="3" style="text-align:center; border-left:1px solid #E0E4E8;">Shapiro-Wilk</th>
            </tr>
            <tr>
                <th style="border-left:1px solid #E0E4E8;">Estadístico</th>
                <th>gl</th>
                <th>Sig.</th>
                <th style="border-left:1px solid #E0E4E8;">Estadístico</th>
                <th>gl</th>
                <th>Sig.</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header">Algoritmo A</td>
                <td>,147</td>
                <td>10</td>
                <td>,200<sup>*</sup></td>
                <td>,978</td>
                <td>10</td>
                <td>,957</td>
            </tr>
            <tr class="total-row">
                <td class="row-header">Algoritmo B</td>
                <td>,119</td>
                <td>11</td>
                <td>,200<sup>*</sup></td>
                <td>,990</td>
                <td>11</td>
                <td>,997</td>
            </tr>
        </tbody>
    </table>
    """
    tables["p2_normalidad.png"] = wrap_html(
        "Explorar",
        "Pruebas de normalidad",
        t2_norm,
        "*. Esto es un límite inferior de la significación verdadera.<br>a. Corrección de significación de Lilliefors"
    )

    # Tabla 2.3: Prueba T para muestras independientes
    t2_ttest = """
    <table class="spss-table">
        <thead>
            <tr>
                <th rowspan="3" style="text-align:left; vertical-align:middle;"></th>
                <th colspan="2" style="text-align:center; border-left:1px solid #E0E4E8;">Prueba de Levene para la igualdad de varianzas</th>
                <th colspan="5" style="text-align:center; border-left:1px solid #E0E4E8;">Prueba T para la igualdad de medias</th>
            </tr>
            <tr>
                <th rowspan="2" style="border-left:1px solid #E0E4E8;">F</th>
                <th rowspan="2">Sig.</th>
                <th rowspan="2" style="border-left:1px solid #E0E4E8;">t</th>
                <th rowspan="2">gl</th>
                <th rowspan="2">Sig. (bilateral)</th>
                <th rowspan="2">Diferencia de medias</th>
                <th rowspan="2">Error estándar de la diferencia</th>
            </tr>
            <tr></tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header">Se asumen varianzas iguales</td>
                <td>1,855</td>
                <td>,189</td>
                <td>5,556</td>
                <td>19</td>
                <td>,000</td>
                <td>15,927</td>
                <td>2,867</td>
            </tr>
            <tr class="total-row">
                <td class="row-header">No se asumen varianzas iguales</td>
                <td></td>
                <td></td>
                <td>5,364</td>
                <td>15,165</td>
                <td>,000</td>
                <td>15,927</td>
                <td>2,969</td>
            </tr>
        </tbody>
    </table>
    """
    tables["p2_t_test.png"] = wrap_html(
        "Prueba T",
        "Prueba de muestras independientes",
        t2_ttest
    )

    # Tabla 2.4: Rangos y Estadísticos de Mann-Whitney (complementaria)
    t2_mw = """
    <table class="spss-table">
        <thead>
            <tr>
                <th style="text-align:left;"></th>
                <th>Tiempo de respuesta (ms)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header">U de Mann-Whitney</td>
                <td>3,000</td>
            </tr>
            <tr>
                <td class="row-header">W de Wilcoxon</td>
                <td>69,000</td>
            </tr>
            <tr>
                <td class="row-header">Z</td>
                <td>-3,535</td>
            </tr>
            <tr class="total-row">
                <td class="row-header">Sig. asintótica(bilateral)</td>
                <td>,000</td>
            </tr>
        </tbody>
    </table>
    """
    tables["p2_mann_whitney.png"] = wrap_html(
        "Pruebas no paramétricas",
        "Estadísticos de prueba<sup>a</sup>",
        t2_mw,
        "a. Variable de agrupación: Algoritmo"
    )

    # ==========================================
    # PROBLEMA 3: PRUEBA DE RACHAS
    # ==========================================
    t3_rachas = """
    <table class="spss-table">
        <thead>
            <tr>
                <th style="text-align:left;"></th>
                <th>Resultado de acceso</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header">Valor de prueba<sup>a</sup></td>
                <td>,50</td>
            </tr>
            <tr>
                <td class="row-header">Casos &lt; Valor de prueba</td>
                <td>14</td>
            </tr>
            <tr>
                <td class="row-header">Casos &gt;= Valor de prueba</td>
                <td>16</td>
            </tr>
            <tr>
                <td class="row-header">Total de casos</td>
                <td>30</td>
            </tr>
            <tr>
                <td class="row-header">Número de rachas</td>
                <td>21</td>
            </tr>
            <tr>
                <td class="row-header">Z</td>
                <td>1,891</td>
            </tr>
            <tr class="total-row">
                <td class="row-header">Sig. asintótica (bilateral)</td>
                <td>,059</td>
            </tr>
        </tbody>
    </table>
    """
    tables["p3_rachas.png"] = wrap_html(
        "Pruebas no paramétricas",
        "Prueba de rachas",
        t3_rachas,
        "a. Mediana / Punto de corte definido para la secuencia dicotómica."
    )

    # ==========================================
    # PROBLEMA 4: PRUEBA BINOMIAL
    # ==========================================
    t4_bin = """
    <table class="spss-table">
        <thead>
            <tr>
                <th style="text-align:left;"></th>
                <th style="text-align:left;">Categoría</th>
                <th>N</th>
                <th>Proporción observada</th>
                <th>Prop. de prueba</th>
                <th>Sig. exacta (unilateral)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header">Grupo 1</td>
                <td style="text-align:left;">Procesada con éxito</td>
                <td>40</td>
                <td>,80</td>
                <td>,90</td>
                <td>,025</td>
            </tr>
            <tr>
                <td class="row-header">Grupo 2</td>
                <td style="text-align:left;">Presentó error</td>
                <td>10</td>
                <td>,20</td>
                <td></td>
                <td></td>
            </tr>
            <tr class="total-row">
                <td class="row-header">Total</td>
                <td style="text-align:left;"></td>
                <td>50</td>
                <td>1,00</td>
                <td></td>
                <td></td>
            </tr>
        </tbody>
    </table>
    """
    tables["p4_binomial.png"] = wrap_html(
        "Pruebas no paramétricas",
        "Prueba binomial",
        t4_bin,
        "a. Basado en la aproximación y cálculo de distribución binomial exacta."
    )

    # ==========================================
    # PROBLEMA 5: TRES SERVICIOS DE ALMACENAMIENTO EN LA NUBE
    # ==========================================
    # Tabla 5.1: Descriptivos (Informe)
    t5_inf = """
    <table class="spss-table">
        <thead>
            <tr>
                <th style="text-align:left;">Servicio en la nube</th>
                <th>Media</th>
                <th>N</th>
                <th>Desv. Desviación</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header">Servicio A</td>
                <td>100,18</td>
                <td>11</td>
                <td>3,601</td>
            </tr>
            <tr>
                <td class="row-header">Servicio B</td>
                <td>121,08</td>
                <td>13</td>
                <td>4,030</td>
            </tr>
            <tr>
                <td class="row-header">Servicio C</td>
                <td>141,00</td>
                <td>15</td>
                <td>4,472</td>
            </tr>
            <tr class="total-row">
                <td class="row-header">Total</td>
                <td>122,82</td>
                <td>39</td>
                <td>17,078</td>
            </tr>
        </tbody>
    </table>
    """
    tables["p5_informe.png"] = wrap_html(
        "Medias",
        "Informe<br><span style='font-weight:normal; font-size:11px;'>Tiempo de respuesta (ms)</span>",
        t5_inf
    )

    # Tabla 5.2: Normalidad
    t5_norm = """
    <table class="spss-table">
        <thead>
            <tr>
                <th rowspan="2" style="text-align:left; vertical-align:middle;">Servicio en la nube</th>
                <th colspan="3" style="text-align:center; border-left:1px solid #E0E4E8;">Kolmogorov-Smirnov<sup>a</sup></th>
                <th colspan="3" style="text-align:center; border-left:1px solid #E0E4E8;">Shapiro-Wilk</th>
            </tr>
            <tr>
                <th style="border-left:1px solid #E0E4E8;">Estadístico</th>
                <th>gl</th>
                <th>Sig.</th>
                <th style="border-left:1px solid #E0E4E8;">Estadístico</th>
                <th>gl</th>
                <th>Sig.</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header">Servicio A</td>
                <td>,152</td>
                <td>11</td>
                <td>,200<sup>*</sup></td>
                <td>,970</td>
                <td>11</td>
                <td>,885</td>
            </tr>
            <tr>
                <td class="row-header">Servicio B</td>
                <td>,134</td>
                <td>13</td>
                <td>,200<sup>*</sup></td>
                <td>,976</td>
                <td>13</td>
                <td>,954</td>
            </tr>
            <tr class="total-row">
                <td class="row-header">Servicio C</td>
                <td>,131</td>
                <td>15</td>
                <td>,200<sup>*</sup></td>
                <td>,964</td>
                <td>15</td>
                <td>,755</td>
            </tr>
        </tbody>
    </table>
    """
    tables["p5_normalidad.png"] = wrap_html(
        "Explorar",
        "Pruebas de normalidad",
        t5_norm,
        "*. Esto es un límite inferior de la significación verdadera.<br>a. Corrección de significación de Lilliefors"
    )

    # Tabla 5.3: Homogeneidad de varianza (Levene)
    t5_lev = """
    <table class="spss-table">
        <thead>
            <tr>
                <th style="text-align:left;"></th>
                <th>Estadístico de Levene</th>
                <th>gl1</th>
                <th>gl2</th>
                <th>Sig.</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header">Se basa en la media</td>
                <td>,476</td>
                <td>2</td>
                <td>36</td>
                <td>,625</td>
            </tr>
            <tr>
                <td class="row-header">Se basa en la mediana</td>
                <td>,461</td>
                <td>2</td>
                <td>36</td>
                <td>,634</td>
            </tr>
            <tr class="total-row">
                <td class="row-header">Se basa en la media recortada</td>
                <td>,474</td>
                <td>2</td>
                <td>36</td>
                <td>,626</td>
            </tr>
        </tbody>
    </table>
    """
    tables["p5_levene.png"] = wrap_html(
        "ANOVA de un factor",
        "Prueba de homogeneidad de varianzas",
        t5_lev
    )

    # Tabla 5.4: ANOVA
    t5_anova = """
    <table class="spss-table">
        <thead>
            <tr>
                <th style="text-align:left;"></th>
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
                <td>10477,530</td>
                <td>2</td>
                <td>5238,765</td>
                <td>316,629</td>
                <td>,000</td>
            </tr>
            <tr>
                <td class="row-header">Dentro de grupos</td>
                <td>595,649</td>
                <td>36</td>
                <td>16,546</td>
                <td></td>
                <td></td>
            </tr>
            <tr class="total-row">
                <td class="row-header">Total</td>
                <td>11073,179</td>
                <td>38</td>
                <td></td>
                <td></td>
                <td></td>
            </tr>
        </tbody>
    </table>
    """
    tables["p5_anova.png"] = wrap_html(
        "ANOVA de un factor",
        "ANOVA<br><span style='font-weight:normal; font-size:11px;'>Tiempo de respuesta (ms)</span>",
        t5_anova
    )

    # Tabla 5.5: Post Hoc Tukey
    t5_posthoc = """
    <table class="spss-table">
        <thead>
            <tr>
                <th style="text-align:left;">(I) Servicio</th>
                <th style="text-align:left;">(J) Servicio</th>
                <th>Diferencia de medias (I-J)</th>
                <th>Error estándar</th>
                <th>Sig.</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header" rowspan="2">Servicio A</td>
                <td style="text-align:left;">Servicio B</td>
                <td>-20,895<sup>*</sup></td>
                <td>1,673</td>
                <td>,000</td>
            </tr>
            <tr>
                <td style="text-align:left;">Servicio C</td>
                <td>-40,818<sup>*</sup></td>
                <td>1,617</td>
                <td>,000</td>
            </tr>
            <tr>
                <td class="row-header" rowspan="2">Servicio B</td>
                <td style="text-align:left;">Servicio A</td>
                <td>20,895<sup>*</sup></td>
                <td>1,673</td>
                <td>,000</td>
            </tr>
            <tr>
                <td style="text-align:left;">Servicio C</td>
                <td>-19,923<sup>*</sup></td>
                <td>1,544</td>
                <td>,000</td>
            </tr>
            <tr>
                <td class="row-header" rowspan="2">Servicio C</td>
                <td style="text-align:left;">Servicio A</td>
                <td>40,818<sup>*</sup></td>
                <td>1,617</td>
                <td>,000</td>
            </tr>
            <tr class="total-row">
                <td style="text-align:left;">Servicio B</td>
                <td>19,923<sup>*</sup></td>
                <td>1,544</td>
                <td>,000</td>
            </tr>
        </tbody>
    </table>
    """
    tables["p5_posthoc.png"] = wrap_html(
        "ANOVA de un factor",
        "Comparaciones múltiples<br><span style='font-weight:normal; font-size:11px;'>HSD Tukey</span>",
        t5_posthoc,
        "*. La diferencia de medias es significativa en el nivel 0.05."
    )

    # Tabla 5.6: Rangos y Kruskal-Wallis (alternativa no paramétrica)
    t5_kw = """
    <table class="spss-table">
        <thead>
            <tr>
                <th style="text-align:left;"></th>
                <th>Tiempo de respuesta (ms)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="row-header">H de Kruskal-Wallis</td>
                <td>33,600</td>
            </tr>
            <tr>
                <td class="row-header">gl</td>
                <td>2</td>
            </tr>
            <tr class="total-row">
                <td class="row-header">Sig. asintótica</td>
                <td>,000</td>
            </tr>
        </tbody>
    </table>
    """
    tables["p5_kruskal.png"] = wrap_html(
        "Pruebas no paramétricas",
        "Estadísticos de prueba<sup>a,b</sup>",
        t5_kw,
        "a. Prueba de Kruskal-Wallis<br>b. Variable de agrupación: Servicio en la nube"
    )

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(device_scale_factor=2)

        for filename, html_content in tables.items():
            filepath = os.path.join(output_dir, filename)
            page.set_content(html_content)
            element = page.locator("#target")
            element.screenshot(path=filepath, omit_background=True)
            print(f"Generado: {filepath}")

        browser.close()

if __name__ == "__main__":
    render_images()
