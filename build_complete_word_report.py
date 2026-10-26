# -*- coding: utf-8 -*-
import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def build_report():
    doc = docx.Document()

    # 1. Page Margins & Setup (APA 7th Edition: 1 inch / 2.54 cm all margins)
    for sec in doc.sections:
        sec.top_margin = Inches(1.0)
        sec.bottom_margin = Inches(1.0)
        sec.left_margin = Inches(1.0)
        sec.right_margin = Inches(1.0)
        sec.page_width = Inches(8.5)
        sec.page_height = Inches(11.0)

    # Styles
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(6)
    style_normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Color palette
    NAVY = RGBColor(0x1B, 0x36, 0x5D)      # Primary institucional
    DARK_BLUE = RGBColor(0x2C, 0x3E, 0x50) # H2
    SLATE = RGBColor(0x34, 0x49, 0x5E)     # H3
    BODY = RGBColor(0x22, 0x22, 0x22)
    GRAY = RGBColor(0x55, 0x55, 0x55)

    # Helper functions
    def add_p(text="", bold=False, italic=False, size=12, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=6, color=BODY):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.italic = italic
            run.font.color.rgb = color
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = NAVY
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12.5)
        run.font.bold = True
        run.font.color.rgb = DARK_BLUE
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.italic = True
        run.font.color.rgb = SLATE
        return p

    def add_img(img_path, caption, width_in=5.8):
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(2)
            p_img.paragraph_format.keep_with_next = True
            run_img = p_img.add_run()
            run_img.add_picture(img_path, width=Inches(width_in))

            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(10)
            run_cap = p_cap.add_run(caption)
            run_cap.font.name = 'Times New Roman'
            run_cap.font.size = Pt(10)
            run_cap.font.bold = True
            run_cap.font.italic = True
            run_cap.font.color.rgb = GRAY
        else:
            print(f"Warning: Image {img_path} not found.")

    def add_placeholder(title, description, detail_text=""):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        cell.width = Inches(6.5)

        shd = parse_xml(r'<w:shd {} w:fill="F8F9F9"/>'.format(nsdecls('w')))
        borders = parse_xml(
            r'<w:tcBorders {} >'
            r'  <w:top w:val="dashed" w:sz="8" w:space="0" w:color="3498DB"/>'
            r'  <w:left w:val="dashed" w:sz="8" w:space="0" w:color="3498DB"/>'
            r'  <w:bottom w:val="dashed" w:sz="8" w:space="0" w:color="3498DB"/>'
            r'  <w:right w:val="dashed" w:sz="8" w:space="0" w:color="3498DB"/>'
            r'</w:tcBorders>'.format(nsdecls('w'))
        )
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(shd)
        tcPr.append(borders)

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run("📷 [ESPACIO RESERVADO PARA IMAGEN POSTERIOR]\n")
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(0x29, 0x80, 0xB9)

        r2 = p.add_run(f"{title}\n")
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(11)
        r2.font.bold = True
        r2.font.color.rgb = NAVY

        r3 = p.add_run(f"{description}")
        r3.font.name = 'Times New Roman'
        r3.font.size = Pt(10)
        r3.font.italic = True
        r3.font.color.rgb = GRAY

        if detail_text:
            p2 = cell.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p2.paragraph_format.space_before = Pt(2)
            p2.paragraph_format.space_after = Pt(6)
            r4 = p2.add_run(f"Instrucción para el usuario: {detail_text}")
            r4.font.name = 'Times New Roman'
            r4.font.size = Pt(9.5)
            r4.font.italic = True
            r4.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)

        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(6)

    def create_custom_table(data, col_widths=None, header_bg="1B365D", alt_bg="F8F9F9"):
        table = doc.add_table(rows=len(data), cols=len(data[0]))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        for r_idx, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(r'<w:cantSplit {}/>'.format(nsdecls('w'))))
            if r_idx == 0:
                trPr.append(parse_xml(r'<w:tblHeader {}/>'.format(nsdecls('w'))))

            for c_idx, cell in enumerate(row.cells):
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                cell.text = str(data[r_idx][c_idx])
                if col_widths and c_idx < len(col_widths):
                    cell.width = col_widths[c_idx]

                tcPr = cell._tc.get_or_add_tcPr()
                if r_idx == 0:
                    shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), header_bg))
                    tcPr.append(shd)
                    for cp in cell.paragraphs:
                        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in cp.runs:
                            run.font.name = 'Times New Roman'
                            run.font.bold = True
                            run.font.size = Pt(10)
                            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                else:
                    if r_idx % 2 == 1:
                        shd = parse_xml(r'<w:shd {} w:fill="FFFFFF"/>'.format(nsdecls('w')))
                    else:
                        shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), alt_bg))
                    tcPr.append(shd)
                    for cp in cell.paragraphs:
                        # left align first col, right align numeric cols
                        if c_idx == 0:
                            cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        else:
                            cp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        for run in cp.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(9.5)
                            run.font.color.rgb = BODY

                borders = parse_xml(
                    r'<w:tcBorders {} >'
                    r'  <w:top w:val="single" w:sz="4" w:space="0" w:color="BDC3C7"/>'
                    r'  <w:left w:val="none"/>'
                    r'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="BDC3C7"/>'
                    r'  <w:right w:val="none"/>'
                    r'</w:tcBorders>'.format(nsdecls('w'))
                )
                tcPr.append(borders)

        p_after = doc.add_paragraph()
        p_after.paragraph_format.space_before = Pt(0)
        p_after.paragraph_format.space_after = Pt(6)
        return table

    # =========================================================================
    # CARÁTULA (Cover Page)
    # =========================================================================
    add_p("UNIVERSIDAD PRIVADA DE TACNA", bold=True, size=16, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=4, color=NAVY)
    add_p("FACULTAD DE INGENIERÍA", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=3, color=NAVY)
    add_p("ESCUELA DE INGENIERÍA DE SISTEMAS", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18, color=NAVY)

    # University Logo
    logo_path = 'extracted_model_media/image8.png'
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_after = Pt(18)
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Inches(2.5))

    add_p("Trabajo Práctico", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4, color=DARK_BLUE)
    add_p("“APLICACIÓN DE ESTADÍSTICA INFERENCIAL CON DATOS REALES”", bold=True, size=15, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24, color=NAVY)

    add_p("Que se presenta para el curso:", italic=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_p("ESTADÍSTICA INFERENCIAL", bold=True, size=12.5, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18, color=DARK_BLUE)

    add_p("Docente:", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_p("LUIS ALFREDO FERNANDEZ VIZCARRA", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)

    add_p("Integrantes:", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_p("Arocutipa Arocutipa, Gian Franco  (2023076790)", size=11.5, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_p("Perez Peralta, Fabrizio Salvador Elias  (2023077476)", size=11.5, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_p("LLica Mamani, Jimmy Mijair  (2023076789)", size=11.5, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=26)

    add_p("TACNA – PERÚ", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_p("2026", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)

    doc.add_page_break()

    # =========================================================================
    # ÍNDICE GENERAL
    # =========================================================================
    add_p("Índice General", bold=True, size=16, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10, space_after=16, color=NAVY)

    toc_items = [
        ("1. Descripción del dataset", "1"),
        ("2. Planteamiento del análisis", "2"),
        ("   2.1 Contexto y planteamiento del problema", "2"),
        ("   2.2 Preguntas de investigación (Q1 - Q5)", "2"),
        ("   2.3 Identificación y justificación de variables", "3"),
        ("3. Análisis descriptivo y estimación", "3"),
        ("   3.1 Estadísticos descriptivos generales", "3"),
        ("   3.2 Asimetría, curtosis y forma de las distribuciones", "4"),
        ("   3.3 Estimación puntual de parámetros poblacionales", "4"),
        ("   3.4 Intervalos de confianza al 95 % para la media", "4"),
        ("   3.5 Intervalos de confianza para proporciones", "5"),
        ("4. Verificación de supuestos", "5"),
        ("   4.1 El supuesto de distribución normal y el Teorema del Límite Central", "5"),
        ("   4.2 Pruebas de normalidad mediante Shapiro–Wilk", "5"),
        ("   4.3 Evaluación de homogeneidad de varianzas mediante la prueba de Levene", "6"),
        ("5. Prueba de hipótesis para una muestra", "7"),
        ("   5.1 Pregunta de investigación y formulación de hipótesis", "7"),
        ("   5.2 Nivel de significancia y verificación de supuestos", "7"),
        ("   5.3 Selección de la prueba estadística", "7"),
        ("   5.4 Presentación de resultados y p-valor en SPSS", "8"),
        ("   5.5 Decisión estadística e interpretación contextual", "8"),
        ("   5.6 Tamaño del efecto (d de Cohen y corrección de Hedges)", "9"),
        ("   5.7 Análisis complementario con valor de referencia 5000 estrellas", "9"),
        ("6. Comparación de promedios entre dos grupos independientes", "10"),
        ("   6.1 Comparación principal de ln_forks_1 según Tipo de Propietario (Organization vs. User)", "10"),
        ("   6.2 Comparación multidimensional complementaria: Organization vs. User", "12"),
        ("   6.3 Comparación secundaria según Tipo de Licencia: MIT vs. Apache-2.0", "14"),
        ("   6.4 Significancia estadística frente a significancia práctica y limitaciones", "15"),
        ("7. Comparación de promedios en muestras relacionadas (antes–después / mediciones sobre las mismas unidades)", "16"),
        ("   7.1 Justificación metodológica y emparejamiento de las unidades", "16"),
        ("   7.2 Pregunta de investigación y formulación de hipótesis", "16"),
        ("   7.3 Cálculo de diferencias emparejadas y verificación de normalidad", "17"),
        ("   7.4 Aplicación de la prueba t para muestras relacionadas en SPSS", "17"),
        ("   7.5 Decisión estadística, tamaño del efecto e interpretación en Ingeniería de Sistemas", "18"),
        ("8. Análisis de Varianza (ANOVA) para tres o más grupos independientes", "20"),
        ("   8.1 Planteamiento y definición de variables dependiente y factor", "20"),
        ("   8.2 Pregunta de investigación y formulación de hipótesis", "20"),
        ("   8.3 Estadísticos descriptivos por grupo de licencia (MIT, Apache-2.0, GPL-3.0)", "20"),
        ("   8.4 Verificación de supuestos: Normalidad y Homogeneidad de Varianzas (Levene)", "21"),
        ("   8.5 Aplicación del ANOVA de un factor y prueba robusta de Welch", "21"),
        ("   8.6 Estadístico F, p-valor, tamaños de efecto y decisión estadística", "22"),
        ("   8.7 Interpretación del resultado en el contexto del software", "23"),
        ("9. Pruebas Post Hoc: Tukey o Games-Howell según corresponda", "24"),
        ("   9.1 Criterio de selección y justificación de la prueba Post Hoc", "24"),
        ("   9.2 Comparaciones múltiples por pares (HSD Tukey)", "24"),
        ("   9.3 Diferencias de medias, errores estándar y p-valores ajustados", "25"),
        ("   9.4 Análisis de subconjuntos homogéneos", "25"),
        ("   9.5 Interpretación contextual en Ingeniería de Software", "26"),
        ("10. Conclusiones", "27"),
        ("11. Referencias y fuentes de datos", "29")
    ]

    for item, page in toc_items:
        p_toc = doc.add_paragraph()
        p_toc.paragraph_format.space_before = Pt(1)
        p_toc.paragraph_format.space_after = Pt(2)
        p_toc.paragraph_format.line_spacing = 1.15
        r_item = p_toc.add_run(item)
        r_item.font.name = 'Times New Roman'
        r_item.font.size = Pt(10.5)
        if item.startswith("1.") or item.startswith("2.") or item.startswith("3.") or item.startswith("4.") or item.startswith("5.") or item.startswith("6.") or item.startswith("7.") or item.startswith("8.") or item.startswith("9.") or item.startswith("10.") or item.startswith("11."):
            r_item.font.bold = True
            r_item.font.color.rgb = NAVY
        else:
            r_item.font.color.rgb = BODY

    doc.add_page_break()

    # =========================================================================
    # 1. DESCRIPCIÓN DEL DATASET
    # =========================================================================
    add_h1("1. Descripción del dataset")

    add_p(
        "El presente informe práctico de estadística inferencial se desarrolla a partir del conjunto de datos real titulado "
        "“GitHub Repositories – Heterogeneous Data”, extraído de la plataforma abierta Kaggle Datasets (disponible públicamente en "
        "https://www.kaggle.com/datasets). Este repositorio alberga información cuantitativa y cualitativa exhaustiva sobre el "
        "ecosistema de desarrollo de software libre a nivel global, constituyendo una fuente fidedigna y de alta relevancia empírica "
        "para el campo de la Ingeniería de Sistemas y las Ciencias de la Computación."
    )

    add_p(
        "El dataset bruto original contiene 66 641 registros heterogéneos estructurados en múltiples entidades relacionales "
        "(elementos de repositorio, contribuidores, lanzamientos de versiones, lenguajes de programación, temas y perfiles de usuarios). "
        "Para los fines del presente estudio y en concordancia estricta con las directivas metodológicas del curso, se seleccionó el "
        "subconjunto homogéneo de registros de tipo RepositoryItem correspondientes a proyectos de código abierto cuyo lenguaje "
        "principal es Python, totalizando una muestra representativa de n = 1567 repositorios de software independientes."
    )

    add_h2("Variables seleccionadas para el estudio:")

    vars_info = [
        ("stars (Cuantitativa discreta / continua):", " Representa el número total de estrellas (stars) recibidas por el repositorio en GitHub. Constituye el indicador por excelencia de popularidad, visibilidad general y respaldo comunitario en el ecosistema open source."),
        ("forks (Cuantitativa discreta / continua):", " Indica el número de bifurcaciones o copias derivadas creadas a partir del repositorio original. Refleja de manera fidedigna la tasa de adopción técnica, colaboración distribuida y reutilización directa del código fuente."),
        ("open_issues (Cuantitativa discreta):", " Cantidad de incidencias, tickets de errores (bugs) y solicitudes de mejora actualmente abiertas y no resueltas. Sirve como métrica clave de la carga de mantenimiento pendiente, actividad de depuración y soporte comunitario."),
        ("watchers y size_kb (Cuantitativas de soporte):", " watchers mide el número de desarrolladores suscritos a las notificaciones en tiempo real del repositorio (cuyos valores en este dataset guardan identidad matemática con stars), mientras que size_kb refleja la magnitud física del repositorio en kilobytes, permitiendo dimensionar la escala del código."),
        ("ln_stars_1, ln_forks_1, ln_open_issues_1 (Cuantitativas transformadas):", " Corresponden a las transformaciones logarítmicas naturales estandarizadas ln(X + 1), aplicadas rigurosamente en el software IBM SPSS Statistics para estabilizar la heterocedasticidad y reducir la asimetría positiva intrínseca de los datos de conteo en repositorios."),
        ("owner_type / owner2 (Cualitativa nominal dicotómica):", " Criterio de clasificación del propietario del proyecto en dos categorías mutuamente excluyentes: Organization (entidades corporativas, comunidades o fundaciones; n = 878) y User (desarrolladores o usuarios individuales; n = 689). Define el diseño de dos grupos independientes."),
        ("license_spdx_id / lic3 (Cualitativa nominal politómica):", " Tipo de licencia de software de código abierto bajo la cual se distribuye el proyecto. Para el Análisis de Varianza (ANOVA) y las pruebas Post Hoc se seleccionaron los tres grupos predominantes: MIT (lic3 = 1.00; n = 546), Apache-2.0 (lic3 = 2.00; n = 405) y GPL-3.0 (lic3 = 3.00; n = 112), sumando N = 1063 repositorios.")
    ]

    for title_v, desc_v in vars_info:
        p_v = doc.add_paragraph()
        p_v.paragraph_format.space_before = Pt(2)
        p_v.paragraph_format.space_after = Pt(4)
        p_v.paragraph_format.line_spacing = 1.15
        r_v1 = p_v.add_run(f"• {title_v}")
        r_v1.font.name = 'Times New Roman'
        r_v1.font.size = Pt(11)
        r_v1.font.bold = True
        r_v1.font.color.rgb = DARK_BLUE
        r_v2 = p_v.add_run(desc_v)
        r_v2.font.name = 'Times New Roman'
        r_v2.font.size = Pt(11)

    # =========================================================================
    # 2. PLANTEAMIENTO DEL ANÁLISIS
    # =========================================================================
    add_h1("2. Planteamiento del análisis")

    add_h2("2.1 Contexto y planteamiento del problema")
    add_p(
        "En la ingeniería de software moderna, las plataformas colaborativas como GitHub representan el núcleo operativo de la "
        "innovación tecnológica. Sin embargo, los proyectos que conviven en este ecosistema difieren sustancialmente en sus modelos "
        "de gobernanza, estructura de licenciamiento y niveles de respaldo comunitario. Mientras que algunas iniciativas son impulsadas "
        "por grandes empresas y consorcios tecnológicos (Organizaciones), otras nacen de programadores independientes (Usuarios). "
        "Asimismo, la elección entre licencias permisivas (MIT, Apache-2.0) y licencias recíprocas o copyleft fuerte (GPL-3.0) define "
        "el marco de derechos y obligaciones bajo el cual la comunidad puede interactuar con el código."
    )
    add_p(
        "Frente a este escenario, resulta crucial determinar mediante modelos de estadística inferencial si los parámetros poblacionales "
        "de popularidad, reutilización de código y mantenimiento presentan patrones sistemáticos diferenciables, o si por el contrario "
        "la dispersión observada obedece al azar muestral. Con este propósito, se estructuran cinco procedimientos inferenciales centrales:"
    )

    add_h2("2.2 Preguntas de investigación")
    questions = [
        ("Q1 (Prueba de hipótesis para una muestra):", " ¿El promedio poblacional de popularidad (medido a través del número de estrellas, stars) de los repositorios de Python de código abierto difiere estadísticamente del estándar de referencia de 9000 estrellas (y complementariamente del umbral de 5000 estrellas)?"),
        ("Q2 (Comparación de dos grupos independientes):", " ¿Existe una diferencia estadísticamente significativa en el nivel promedio de bifurcaciones transformadas (ln_forks_1) entre los repositorios gestionados por Organizaciones y los administrados por Usuarios individuales?"),
        ("Q3 (Comparación de muestras relacionadas):", " ¿Existe una diferencia estadísticamente significativa entre el promedio de estrellas (ln_stars_1) y el promedio de bifurcaciones (ln_forks_1) medidos sobre las mismas unidades muestrales (mismos repositorios de software)? ¿Supera la popularidad pasiva (stars) a la adopción colaborativa activa (forks)?"),
        ("Q4 (Análisis de Varianza - ANOVA de un factor):", " ¿Existen diferencias estadísticamente significativas en el promedio de incidencias abiertas transformadas (ln_open_issues_1) entre los repositorios de software según su tipo de licencia (MIT, Apache-2.0 y GPL-3.0)?"),
        ("Q5 (Pruebas Post Hoc de comparaciones múltiples):", " Habiéndose determinado la significancia global del ANOVA, ¿entre qué pares específicos de licencias se localizan las diferencias significativas en la carga de issues abiertos y qué implicancias metodológicas revisten?")
    ]

    for q_t, q_d in questions:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(2)
        p_q.paragraph_format.space_after = Pt(4)
        p_q.paragraph_format.line_spacing = 1.15
        r_q1 = p_q.add_run(f"• {q_t}")
        r_q1.font.name = 'Times New Roman'
        r_q1.font.size = Pt(11)
        r_q1.font.bold = True
        r_q1.font.color.rgb = NAVY
        r_q2 = p_q.add_run(q_d)
        r_q2.font.name = 'Times New Roman'
        r_q2.font.size = Pt(11)

    add_h2("2.3 Justificación metodológica de las variables")
    add_p(
        "La selección de stars, forks y open_issues obedece a que constituyen la tríada métrica universal del modelo de desarrollo open source. "
        "Las estrellas cuantifican el aprecio e impacto externo del producto; las bifurcaciones reflejan el compromiso de la comunidad técnica para "
        "extender o auditar el software; y las incidencias abiertas reflejan la deuda técnica acumulada y la dinámica de mantenimiento. "
        "Por su parte, la inclusión de las transformaciones ln(X + 1) responde a las mejores prácticas de la inferencia paramétrica ante distribuciones "
        "con marcada asimetría positiva, permitiendo reducir el sesgo de escala sin alterar el orden monotónico de los datos."
    )

    # =========================================================================
    # 3. ANÁLISIS DESCRIPTIVO Y ESTIMACIÓN
    # =========================================================================
    add_h1("3. Análisis descriptivo y estimación")

    add_h2("3.1 Estadísticos descriptivos generales")
    add_p(
        "A continuación se presentan los estadísticos descriptivos fundamentales obtenidos a partir de la muestra completa de n = 1567 "
        "repositorios de software en Python. La Tabla 1 consolida las medidas de tendencia central, dispersión y posición para las variables "
        "originales bajo estudio."
    )

    add_p("Tabla 1. Estadísticos descriptivos generales de la muestra completa (n = 1567).", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)

    t1_data = [
        ["Estadístico", "Stars", "Forks", "Watchers", "Open Issues", "Size (KB)"],
        ["N", "1567", "1567", "1567", "1567", "1567"],
        ["Media", "8963.88", "1318.58", "8963.88", "179.91", "136916.26"],
        ["Desv. estándar", "4411.51", "1211.51", "4411.51", "337.52", "769547.30"],
        ["Mínimo", "4001", "8", "4001", "0", "3"],
        ["Q1 (25 %)", "4591.50", "592.50", "4591.50", "29.00", "3486.00"],
        ["Mediana", "8539.00", "998.00", "8539.00", "78.00", "16923.00"],
        ["Q3 (75 %)", "11960.00", "1649.50", "11960.00", "186.00", "71656.50"],
        ["Máximo", "19931", "19556", "19931", "5959", "20737671"]
    ]
    create_custom_table(t1_data)

    add_h2("3.2 Asimetría, curtosis y forma de las distribuciones")
    add_p(
        "La evaluación de la forma distributiva revela una asimetría positiva consistente en todas las variables analizadas, "
        "donde la media aritmética supera sensiblemente a la mediana. Este comportamiento es característico de las métricas de software, "
        "en las que la gran mayoría de proyectos exhibe niveles moderados de actividad mientras que un reducido grupo de proyectos "
        "estelares acumula volúmenes desproporcionados de bifurcaciones, incidencias y tamaño de almacenamiento."
    )

    add_p("Tabla 2. Coeficientes de asimetría y curtosis (n = 1567).", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t2_data = [
        ["Variable", "Stars", "Forks", "Open Issues", "Size (KB)"],
        ["Asimetría", "0.60", "4.74", "6.71", "16.94"],
        ["Curtosis", "-0.66", "48.69", "76.95", "373.41"]
    ]
    create_custom_table(t2_data)

    add_p(
        "Interpretación: La variable stars exhibe una asimetría leve a moderada (0.60) y curtosis platicúrtica (-0.66), facilitando "
        "su análisis en pruebas de una muestra. En marcado contraste, forks, open_issues y size_kb presentan coeficientes de asimetría "
        "y curtosis leptocúrtica extremadamente elevados (curtosis de 48.69, 76.95 y 373.41 respectivamente), justificando plenamente "
        "el uso de la transformación logarítmica natural ln(X + 1) aplicada en las secciones subsecuentes."
    )

    add_h2("3.3 Estimación puntual de parámetros poblacionales")
    add_p(
        "A partir de los estimadores insesgados calculados sobre la muestra representativa (n = 1567), las estimaciones puntuales de los "
        "parámetros de tendencia central y variabilidad poblacional son:"
    )
    add_p("• Media poblacional de Stars: μ̂ = 8963.88 estrellas (Varianza estimada: s² = 19 461 426.70; Desv. Est.: s = 4411.51).")
    add_p("• Media poblacional de Forks: μ̂ = 1318.58 bifurcaciones (Varianza estimada: s² = 1 467 748.40; Desv. Est.: s = 1211.51).")
    add_p("• Media poblacional de Open Issues: μ̂ = 179.91 incidencias (Varianza estimada: s² = 113 922.21; Desv. Est.: s = 337.52).")
    add_p("• Media poblacional de Size (KB): μ̂ = 136 916.26 KB (Desv. Est.: s = 769 547.30 KB).")

    add_h2("3.4 Intervalos de confianza al 95 % para la media")
    add_p(
        "Utilizando la distribución t de Student (con n - 1 = 1566 grados de libertad y valor crítico t_0.025 ≈ 1.9615), se determinaron "
        "los intervalos de confianza al 95 % para la media poblacional de cada variable métrica:"
    )

    add_p("Tabla 3. Intervalos de confianza al 95 % para la media poblacional (n = 1567).", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t3_data = [
        ["Variable", "Media", "Error estándar (SE)", "Margen de error", "IC 95 %"],
        ["Stars", "8963.88", "111.44", "218.59", "[8745.29 ; 9182.48]"],
        ["Forks", "1318.58", "30.60", "60.03", "[1258.55 ; 1378.61]"],
        ["Open Issues", "179.91", "8.53", "16.72", "[163.18 ; 196.63]"],
        ["Size (KB)", "136916.26", "19440.00", "38131.50", "[98784.76 ; 175047.76]"]
    ]
    create_custom_table(t3_data)

    add_p(
        "Interpretación: Con un 95 % de nivel de confianza, se tiene la certeza estadística de que el verdadero promedio poblacional "
        "de estrellas de los repositorios de Python se sitúa entre 8745.29 y 9182.48. Es notable resaltar que el valor de prueba μ₀ = 9000 "
        "queda contenido dentro de este intervalo, lo cual anticipa formalmente la no significancia de la prueba t en la Sección 5."
    )

    add_h2("3.5 Intervalos de confianza para proporciones")
    add_p(
        "Complementariamente, se evaluaron proporciones de hitos críticos en la comunidad open source: de los 1567 proyectos, 573 repositorios "
        "(36.57 %) superan las 10 000 estrellas, estableciendo un IC 95 % = [34.18 % ; 38.95 %]. Asimismo, 130 repositorios (8.30 %) superan "
        "las 500 incidencias abiertas, con un IC 95 % = [6.93 % ; 9.67 %], demostrando que una proporción sustancial del ecosistema analizado "
        "ostenta estatus de alta masividad comunitaria."
    )

    # =========================================================================
    # 4. VERIFICACIÓN DE SUPUESTOS
    # =========================================================================
    add_h1("4. Verificación de supuestos")

    add_h2("4.1 El supuesto de distribución normal y el Teorema del Límite Central")
    add_p(
        "La aplicación rigurosa de procedimientos inferenciales paramétricos descansa sobre dos pilares teóricos fundamentales: "
        "la distribución normal de las variables y la homogeneidad de varianzas (homocedasticidad). En la práctica del desarrollo de software, "
        "las variables de conteo nunca siguen una distribución gaussiana perfecta debido a la naturaleza discreta de los eventos y la presencia "
        "de proyectos de impacto extraordinario (efecto 'red')."
    )
    add_p(
        "No obstante, el Teorema del Límite Central (TLC) establece formalmente que, para tamaños muestrales suficientemente grandes "
        "(típicamente n ≥ 30, y en este estudio n = 1567, con subgrupos de n = 405 a 878), la distribución muestral de las medias muestrales "
        "aproxima asintóticamente a una distribución normal con media μ y varianza σ²/n, con independencia de la forma distributiva "
        "de la población original. Esto confiere una solidez matemática incuestionable a las pruebas t y al modelo ANOVA implementados."
    )

    add_h2("4.2 Pruebas de normalidad mediante Shapiro–Wilk")
    add_p(
        "Se contrastaron formalmente las hipótesis de normalidad (H₀: los datos proceden de una población con distribución normal vs. "
        "H₁: los datos no proceden de una población normal) fijando un nivel de significancia de α = 0.05. La Tabla 4 resume las pruebas "
        "de Shapiro–Wilk para la muestra global y la Tabla 5 por grupos de propietarios."
    )

    add_p("Tabla 4. Prueba de normalidad Shapiro–Wilk — Muestra completa (n = 1567).", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t4_data = [
        ["Variable", "Estadístico W", "Grados de libertad (gl)", "Valor p", "¿Sigue distribución normal?"],
        ["Stars", "0.8989", "1567", "< 0.0001", "No (Rechaza H₀)"],
        ["Forks", "0.6878", "1567", "< 0.0001", "No (Rechaza H₀)"],
        ["Open Issues", "0.4757", "1567", "< 0.0001", "No (Rechaza H₀)"],
        ["Size (KB)", "0.1316", "1567", "< 0.0001", "No (Rechaza H₀)"]
    ]
    create_custom_table(t4_data)

    add_p("Tabla 5. Prueba de normalidad Shapiro–Wilk por Tipo de Propietario (Organization vs. User).", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t5_data = [
        ["Variable", "Grupo de Propietario", "N", "Estadístico W", "Valor p", "¿Normal?"],
        ["Stars", "Organization", "878", "0.9076", "< 0.0001", "No"],
        ["Stars", "User", "689", "0.8861", "< 0.0001", "No"],
        ["Forks", "Organization", "878", "0.8103", "< 0.0001", "No"],
        ["Forks", "User", "689", "0.6040", "< 0.0001", "No"],
        ["Open Issues", "Organization", "878", "0.5043", "< 0.0001", "No"],
        ["Open Issues", "User", "689", "0.5647", "< 0.0001", "No"]
    ]
    create_custom_table(t5_data)

    add_p(
        "Decisión sobre normalidad: Al arrojar p < 0.001 en todas las variables originales, se rechaza formalmente H₀. Esto ratifica que ninguna "
        "variable bruta es estrictamente normal. Sin embargo, amparados en el Teorema del Límite Central y robustecidos mediante el empleo de "
        "transformaciones logarítmicas ln(X + 1), las inferencias sobre la media mantienen completa validez estadística."
    )

    add_h2("4.3 Evaluación de homogeneidad de varianzas mediante la prueba de Levene")
    add_p(
        "El supuesto de homocedasticidad (H₀: σ₁² = σ₂² vs. H₁: σ₁² ≠ σ₂², con α = 0.05) se evaluó mediante el estadístico de Levene basado "
        "en la media y en la mediana. Este supuesto determina la selección entre la prueba t de Student clásica (cuando las varianzas son homogéneas) "
        "y la prueba t de Welch con grados de libertad ajustados de Satterthwaite (cuando existe heterocedasticidad)."
    )

    add_p("Tabla 6. Prueba de Levene para igualdad de varianzas — Organization vs. User.", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t6_data = [
        ["Variable", "Estadístico de Levene (F)", "Valor p", "¿Varianzas homogéneas?", "Prueba inferencial a usar"],
        ["Stars", "3.8845", "0.0489", "No (p < 0.05)", "t de Welch (varianzas desiguales)"],
        ["Forks", "3.0677", "0.0801", "Sí (p ≥ 0.05)", "t de Student (varianzas iguales)"],
        ["ln_forks_1 (SPSS)", "7.3520", "0.0070", "No (p < 0.05)", "t de Welch (varianzas desiguales)"],
        ["Open Issues", "40.1906", "< 0.0001", "No (p < 0.05)", "t de Welch (varianzas desiguales)"],
        ["Size (KB)", "0.0664", "0.7968", "Sí (p ≥ 0.05)", "t de Student (varianzas iguales)"]
    ]
    create_custom_table(t6_data)

    add_p(
        "Interpretación: A diferencia de la normalidad, la homocedasticidad varía según la variable considerada. En el análisis central "
        "de ln_forks_1, Levene arroja F = 7.352, p = 0.007 < 0.05, descartando la igualdad de varianzas y demandando el uso obligado "
        "de la prueba t de Welch, procedimiento que ajusta con exactitud los grados de libertad para garantizar la tasa de error tipo I."
    )

    # =========================================================================
    # 5. PRUEBA DE HIPÓTESIS PARA UNA MUESTRA
    # =========================================================================
    add_h1("5. Prueba de hipótesis para una muestra")

    add_h2("5.1 Pregunta de investigación y formulación de hipótesis")
    add_p(
        "Pregunta de investigación: ¿El promedio poblacional de popularidad (estrellas, stars) de los repositorios de código abierto en Python "
        "difiere significativamente del valor de referencia estándar de 9000 estrellas?"
    )
    add_p("Hipótesis estadística formulada (prueba bilateral o de dos colas):")
    add_p("• Hipótesis Nula (H₀): μ = 9000  (El promedio poblacional de estrellas es igual a 9000).")
    add_p("• Hipótesis Alterna (H₁): μ ≠ 9000  (El promedio poblacional de estrellas es diferente de 9000).")

    add_h2("5.2 Nivel de significancia y verificación de supuestos")
    add_p("• Nivel de significancia establecido: α = 0.05.")
    add_p(
        "• Verificación de normalidad: En la prueba de Shapiro–Wilk ejecutada en SPSS sobre la variable stars (n = 1567), se obtuvo un estadístico "
        "W = 0.899 (gl = 1567, p < 0.001) y en Kolmogorov–Smirnov con corrección Lilliefors D = 0.195 (p < 0.001). Si bien se rechaza la normalidad "
        "estricta, el gran tamaño de la muestra (n = 1567 >> 30) asegura por el Teorema del Límite Central que la media muestral sigue una distribución "
        "normal, posibilitando la aplicación certera de la prueba t."
    )

    add_h2("5.3 Selección de la prueba estadística")
    add_p(
        "Se seleccionó la prueba t de Student para una muestra cuantitativa continua, calculando el estadístico t = (x̄ - μ₀) / (s / √n) con "
        "gl = n - 1 = 1566 grados de libertad."
    )

    add_h2("5.4 Presentación de resultados y p-valor en SPSS")
    add_p(
        "Los análisis procesados en IBM SPSS Statistics arrojaron los siguientes resultados muestrales:"
    )
    add_p("• Tamaño de la muestra: n = 1567 repositorios válidos (100.0 % de casos procesados, 0 % perdidos).")
    add_p("• Media muestral observada: x̄ = 8963.8826 estrellas.")
    add_p("• Desviación estándar muestral: s = 4411.5107 estrellas.")
    add_p("• Error estándar de la media: SE = 111.4430 estrellas.")
    add_p("• Valor de prueba de referencia: μ₀ = 9000 estrellas.")
    add_p("• Diferencia de medias (x̄ - μ₀): -36.1174 estrellas.")
    add_p("• Intervalo de confianza al 95 % de la diferencia: [-254.711 ; 182.476].")
    add_p("• Estadístico t calculado: t = -0.324.")
    add_p("• Grados de libertad: gl = 1566.")
    add_p("• Significancia bilateral (p-valor de dos factores): p = 0.746 (p de un factor = 0.373).")

    add_h2("5.5 Decisión estadística e interpretación contextual")
    add_p(
        "Regla de decisión inferencial:"
    )
    add_p("• Si p-valor ≥ 0.05 → No se rechaza H₀.")
    add_p("• Si p-valor < 0.05 → Se rechaza H₀ a favor de H₁.")
    add_p(
        "Decisión estadística: Dado que el p-valor obtenido (p = 0.746) es ampliamente superior al nivel de significancia α = 0.05, "
        "NO SE RECHAZA LA HIPÓTESIS NULA (H₀)."
    )
    add_p(
        "Interpretación en el contexto del problema: No existe evidencia estadística suficiente para afirmar que la popularidad promedio de los "
        "repositorios de software en Python difiera del estándar de 9000 estrellas. La discrepancia observada de 36.12 estrellas representa una "
        "variación aleatoria atribuible al muestreo y carece de significancia práctica y estadística. En consecuencia, se concluye con un 95 % de "
        "confianza que el nivel de respaldo comunitario de los proyectos analizados coincide estrechamente con el valor de referencia de 9000 estrellas."
    )

    add_h2("5.6 Tamaño del efecto (d de Cohen y corrección de Hedges)")
    add_p(
        "Para complementar el p-valor e ilustrar la magnitud real de la diferencia observada, se calcularon las métricas de tamaño del efecto "
        "estandarizadas en SPSS:"
    )
    add_p("• d de Cohen: d = (x̄ - μ₀) / s = -0.008 (IC 95 %: [-0.058 ; 0.041]). Denominador estandarizador: s = 4411.5107.")
    add_p("• Corrección de Hedges: g = -0.008 (IC 95 %: [-0.058 ; 0.041]). Factor estandarizador corregido: 4413.6249.")
    add_p(
        "Interpretación del tamaño del efecto: Un valor de |d| = 0.008 se clasifica inequívocamente como un efecto nulo o trivial (|d| < 0.20), "
        "corroborando que prácticamente no existe distancia sustantiva entre la realidad observada en el código abierto y el umbral de 9000 estrellas."
    )

    # Insert images from trabajo.docx for Section 5
    add_img("extracted_trabajo_media/image15.png", "Figura 1. Resumen de procesamiento de casos, estadísticos descriptivos y pruebas de normalidad de stars en SPSS.")
    add_img("extracted_trabajo_media/image9.png", "Figura 2. Histograma de frecuencias de la variable stars con curva de ajuste y parámetros muestrales.")
    add_img("extracted_trabajo_media/image7.png", "Figura 3. Gráfico Q-Q normal de la variable stars frente a los cuantiles teóricos esperados.")
    add_img("extracted_trabajo_media/image6.png", "Figura 4. Gráfico Q-Q normal sin tendencia (detrended) para las desviaciones de normalidad de stars.")
    add_img("extracted_trabajo_media/image11.png", "Figura 5. Diagrama de caja y bigotes (Boxplot) de stars mostrando cuartiles y rango intercuartílico.")
    add_img("extracted_trabajo_media/image5.png", "Figura 6. Tabla SPSS de Estadísticas para una muestra, Prueba t (valor de prueba = 9000) y Tamaños de efecto.")

    add_h2("5.7 Análisis complementario con valor de referencia 5000 estrellas")
    add_p(
        "Como contraste metodológico, si se evalúa la pregunta frente al umbral histórico de popularidad mínima comunitaria de μ₀ = 5000 estrellas "
        "(H₁: μ ≠ 5000), la prueba t arroja un estadístico t = 35.57 (gl = 1566, p < 0.0001), rechazándose contundentemente H₀ con un tamaño del "
        "efecto grande (d = 0.90, diferencia de +3963.88 estrellas). Esto demuestra que, mientras el ecosistema Python sobrepasa holgadamente el umbral "
        "de 5000 estrellas, se estabiliza con notable precisión alrededor de la cota de 9000 estrellas."
    )

    # =========================================================================
    # 6. COMPARACIÓN DE PROMEDIOS ENTRE DOS GRUPOS INDEPENDIENTES
    # =========================================================================
    add_h1("6. Comparación de promedios entre dos grupos independientes")

    add_h2("6.1 Comparación principal de ln_forks_1 según Tipo de Propietario (Organization vs. User)")
    add_p(
        "Pregunta de investigación: ¿Existen diferencias estadísticamente significativas en el nivel de bifurcaciones transformadas (ln_forks_1) "
        "entre repositorios gestionados por Organizaciones (empresas, fundaciones) y repositorios mantenidos por Usuarios individuales?"
    )

    add_h3("6.1.1 Definición de grupos y formulación de hipótesis")
    add_p("• Grupo 1: Organization (owner2 = 1.00; n₁ = 878 repositorios).")
    add_p("• Grupo 2: User (owner2 = 2.00; n₂ = 689 repositorios).")
    add_p("Hipótesis formulada (bilateral, nivel de significancia α = 0.05):")
    add_p("• H₀: μ_Org = μ_User  (El promedio de ln_forks_1 es igual entre organizaciones y usuarios individuales).")
    add_p("• H₁: μ_Org ≠ μ_User  (El promedio de ln_forks_1 es diferente entre organizaciones y usuarios individuales).")

    add_h3("6.1.2 Verificación de supuestos: Normalidad y Levene")
    add_p(
        "• Normalidad por grupo: La prueba de Shapiro–Wilk sobre ln_forks_1 arroja W = 0.9774 (p < 0.001) para Organizaciones y W = 0.9972 "
        "(p = 0.2974) para Usuarios individuales. El grupo de usuarios individuales cumple satisfactoriamente la normalidad al 5 %, mientras que "
        "el grupo de organizaciones se ampara en el Teorema del Límite Central (n₁ = 878 >> 30)."
    )
    add_p(
        "• Homogeneidad de varianzas (Prueba de Levene): En la salida de SPSS, la prueba de Levene arroja un estadístico F = 7.352 con una "
        "significancia de p = 0.007. Dado que p = 0.007 < 0.05, se RECHAZA la hipótesis de igualdad de varianzas. Las varianzas entre grupos "
        "son heterogéneas, descartando el uso de la prueba t de Student clásica."
    )

    add_h3("6.1.3 Selección y aplicación de la prueba: t de Welch")
    add_p(
        "Ante la heterocedasticidad comprobada, se selecciona la prueba t de Welch ('No se asumen varianzas iguales'), la cual corrige "
        "los grados de libertad mediante la fórmula de Welch–Satterthwaite:"
    )
    add_p("• Media muestral Organization: x̄₁ = 6.9141 (s₁ = 0.7477; SE₁ = 0.0252).")
    add_p("• Media muestral User: x̄₂ = 6.8701 (s₂ = 0.8030; SE₂ = 0.0306).")
    add_p("• Diferencia de medias (x̄₁ - x̄₂): 0.04391 unidades logarítmicas.")
    add_p("• Error estándar de la diferencia: SE_dif = 0.03966.")
    add_p("• Intervalo de confianza al 95 % de la diferencia: [-0.03388 ; 0.12171].")
    add_p("• Estadístico t de Welch: t = 1.107.")
    add_p("• Grados de libertad ajustados: gl = 1425.171.")
    add_p("• Significancia bilateral: p = 0.268 (p de un factor = 0.134).")

    add_h3("6.1.4 Decisión estadística e interpretación")
    add_p(
        "Decisión: Como el p-valor (p = 0.268) es estrictamente mayor a α = 0.05, NO SE RECHAZA LA HIPÓTESIS NULA (H₀)."
    )
    add_p(
        "Interpretación en el contexto del problema: No existe diferencia estadísticamente significativa en el nivel promedio de bifurcaciones "
        "transformadas entre proyectos de organizaciones y proyectos de usuarios independientes. Esto demuestra un hallazgo de alto valor para "
        "la Ingeniería de Sistemas: la capacidad de un software para ser derivado, clonado y reutilizado por la comunidad depende del interés "
        "técnico de la solución algorítmica y no de si quien administra el repositorio es una corporación o un desarrollador independiente."
    )

    add_h3("6.1.5 Tamaño del efecto para dos grupos independientes")
    add_p("• d de Cohen: d = 0.057 (IC 95 %: [-0.043 ; 0.157]). Estandarizador combinado = 0.7725.")
    add_p("• Corrección de Hedges: g = 0.057 (IC 95 %: [-0.043 ; 0.157]). Estandarizador corregido = 0.7729.")
    add_p("• Delta de Glass: Δ = 0.055 (IC 95 %: [-0.045 ; 0.154]). Estandarizador grupo control (User) = 0.8030.")
    add_p("Interpretación: Los tres índices arrojan magnitudes inferiores a 0.10, tipificando un efecto prácticamente nulo o trivial.")

    # Insert images from trabajo.docx for Section 6
    add_img("extracted_trabajo_media/image2.png", "Figura 7. Prueba de Levene para homogeneidad de varianzas de ln_forks_1 según tipo de propietario (F = 7.352, Sig. = 0.007).")
    add_img("extracted_trabajo_media/image12.png", "Figura 8. Estadísticas de grupo y Prueba de muestras independientes (t de Welch y Student) para ln_forks_1 en SPSS.")
    add_img("extracted_trabajo_media/image14.png", "Figura 9. Nota explicativa de Ayuda Curada generada por SPSS confirmando que la diferencia media no es significativa.")
    add_img("extracted_trabajo_media/image8.png", "Figura 10. Tamaños de efecto de muestras independientes (d de Cohen, g de Hedges y Delta de Glass) para ln_forks_1.")

    # Placeholder for user future images in Section 6
    add_placeholder(
        "Gráficos Exploratorios Adicionales por Tipo de Propietario (Organization vs. User)",
        "Espacio reservado para Gráfico Q-Q Normal / Diagrama de Cajas comparativo de ln_forks_1 por grupo.",
        "Inserte aquí las capturas de gráficos Q-Q o diagramas de cajas generados en SPSS para Organization vs. User."
    )

    add_h2("6.2 Comparación multidimensional complementaria: Organization vs. User")
    add_p(
        "Para ofrecer una visión integral del dataset, se evaluaron las cuatro métricas originales entre Organization y User. "
        "Las Tablas 7, 8 y 9 exponen los estadísticos descriptivos, los contrastes inferenciales y los tamaños del efecto."
    )

    add_p("Tabla 7. Estadísticos descriptivos por grupo — Organization vs. User.", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t7_data = [
        ["Variable", "Grupo", "N", "Media", "Desv. Estándar", "Error Estándar"],
        ["Stars", "Organization", "878", "9243.75", "4496.82", "151.76"],
        ["Stars", "User", "689", "8607.25", "4277.10", "162.94"],
        ["Forks", "Organization", "878", "1305.18", "1023.69", "34.55"],
        ["Forks", "User", "689", "1335.65", "1415.80", "53.94"],
        ["Open Issues", "Organization", "878", "237.30", "414.29", "13.98"],
        ["Open Issues", "User", "689", "106.77", "175.96", "6.70"],
        ["Size (KB)", "Organization", "878", "144030.95", "838704.44", "28304.90"],
        ["Size (KB)", "User", "689", "127849.93", "671670.34", "25588.60"]
    ]
    create_custom_table(t7_data)

    add_p("Tabla 8. Pruebas de hipótesis para dos grupos independientes — Organization vs. User.", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t8_data = [
        ["Variable", "Prueba aplicada", "Estadístico t", "Grados de libertad (gl)", "Sig. (bilateral)", "Decisión inferencial"],
        ["Stars", "t de Welch", "2.859", "1511.45", "0.004", "Diferencia significativa (p < 0.05)"],
        ["Forks", "t de Student", "-0.494", "1565", "0.621", "No significativa (p ≥ 0.05)"],
        ["Open Issues", "t de Welch", "8.418", "1247.38", "< 0.001", "Diferencia significativa (p < 0.05)"],
        ["Size (KB)", "t de Student", "0.413", "1565", "0.680", "No significativa (p ≥ 0.05)"]
    ]
    create_custom_table(t8_data)

    add_p("Tabla 9. Tamaños del efecto e intervalos de confianza — Organization vs. User.", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t9_data = [
        ["Variable", "Diferencia de medias", "IC 95 % de la diferencia", "g de Hedges", "Magnitud del efecto"],
        ["Stars", "636.50", "[200.07 ; 1072.93]", "0.15", "Trivial - pequeño"],
        ["Open Issues", "130.53", "[100.14 ; 160.92]", "0.39", "Pequeño - moderado"],
        ["Forks", "-30.47", "[-151.78 ; 90.84]", "-0.03", "Trivial (no significativo)"],
        ["Size (KB)", "16181.02", "[-60759.50 ; 93121.54]", "0.02", "Trivial (no significativo)"]
    ]
    create_custom_table(t9_data)

    # Insert bar chart figures from model
    add_img("extracted_model_media/image9.png", "Figura 11. Comparación de la media de stars por tipo de propietario (Organization vs. User).", width_in=4.5)
    add_img("extracted_model_media/image5.png", "Figura 12. Comparación de la media de forks por tipo de propietario (Organization vs. User).", width_in=4.5)
    add_img("extracted_model_media/image7.png", "Figura 13. Comparación de la media de open_issues por tipo de propietario (Organization vs. User).", width_in=4.5)
    add_img("extracted_model_media/image6.png", "Figura 14. Comparación de la media de size_kb por tipo de propietario (Organization vs. User).", width_in=4.5)

    add_h2("6.3 Comparación secundaria según Tipo de Licencia: MIT vs. Apache-2.0")
    add_p(
        "Se contrastaron los repositorios bajo las dos licencias más utilizadas (MIT: n₁ = 546 vs. Apache-2.0: n₂ = 405). "
        "Las Tablas 10, 11 y 12 resumen este procedimiento."
    )

    add_p("Tabla 10. Estadísticos descriptivos por grupo — MIT vs. Apache-2.0.", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t10_data = [
        ["Variable", "Grupo de Licencia", "N", "Media", "Desv. Estándar", "Error Estándar"],
        ["Stars", "MIT", "546", "8860.57", "4437.93", "189.93"],
        ["Stars", "Apache-2.0", "405", "9090.93", "4379.47", "217.62"],
        ["Forks", "MIT", "546", "1312.85", "1329.09", "56.88"],
        ["Forks", "Apache-2.0", "405", "1262.27", "1162.93", "57.79"],
        ["Open Issues", "MIT", "546", "138.47", "263.30", "11.27"],
        ["Open Issues", "Apache-2.0", "405", "209.83", "302.11", "15.01"],
        ["Size (KB)", "MIT", "546", "94772.06", "422444.77", "18078.97"],
        ["Size (KB)", "Apache-2.0", "405", "101869.28", "214894.27", "10678.18"]
    ]
    create_custom_table(t10_data)

    add_p("Tabla 11. Pruebas de hipótesis para dos grupos independientes — MIT vs. Apache-2.0.", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t11_data = [
        ["Variable", "Prueba aplicada", "Estadístico t", "Grados de libertad (gl)", "Sig. (bilateral)", "Decisión inferencial"],
        ["Stars", "t de Student", "-0.796", "949", "0.426", "No significativa (p ≥ 0.05)"],
        ["Forks", "t de Student", "0.612", "949", "0.541", "No significativa (p ≥ 0.05)"],
        ["Open Issues", "t de Welch", "-3.802", "804.88", "< 0.001", "Diferencia significativa (p < 0.05)"],
        ["Size (KB)", "t de Student", "-0.310", "949", "0.757", "No significativa (p ≥ 0.05)"]
    ]
    create_custom_table(t11_data)

    add_p("Tabla 12. Tamaños del efecto e intervalos de confianza — MIT vs. Apache-2.0.", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t12_data = [
        ["Variable", "Diferencia de medias", "IC 95 % de la diferencia", "g de Hedges", "Magnitud del efecto"],
        ["Open Issues", "-71.36", "[-108.15 ; -34.57]", "-0.25", "Pequeño (significativo)"],
        ["Stars", "-230.36", "[-797.80 ; 337.08]", "-0.05", "Trivial (no significativo)"],
        ["Forks", "50.58", "[-111.45 ; 212.61]", "0.04", "Trivial (no significativo)"],
        ["Size (KB)", "-7097.22", "[-52026.40 ; 37831.96]", "-0.02", "Trivial (no significativo)"]
    ]
    create_custom_table(t12_data)

    # Insert license bar charts from model
    add_img("extracted_model_media/image1.png", "Figura 15. Comparación de la media de stars por tipo de licencia (MIT vs. Apache-2.0).", width_in=4.5)
    add_img("extracted_model_media/image2.png", "Figura 16. Comparación de la media de forks por tipo de licencia (MIT vs. Apache-2.0).", width_in=4.5)
    add_img("extracted_model_media/image3.png", "Figura 17. Comparación de la media de open_issues por tipo de licencia (MIT vs. Apache-2.0).", width_in=4.5)
    add_img("extracted_model_media/image4.png", "Figura 18. Comparación de la media de size_kb por tipo de licencia (MIT vs. Apache-2.0).", width_in=4.5)

    add_h2("6.4 Significancia estadística frente a significancia práctica y limitaciones")
    add_p(
        "Al contar con un tamaño de muestra masivo (N = 1567), la potencia estadística de los contrastes es sumamente alta, permitiendo "
        "detectar diferencias mínimas como 'estadísticamente significativas'. Por ello, la evaluación del tamaño del efecto (g de Hedges) "
        "resulta indispensable: en Open Issues la ventaja de Organizaciones y Apache-2.0 corresponde a efectos pequeños a moderados (g = 0.39 y -0.25), "
        "revelando que el impacto real, si bien estadísticamente detectable, no representa una brecha infranqueable en la práctica ingenieril."
    )

    # =========================================================================
    # 7. COMPARACIÓN DE PROMEDIOS EN MUESTRAS RELACIONADAS
    # =========================================================================
    add_h1("7. Comparación de promedios en muestras relacionadas (antes–después / mediciones sobre las mismas unidades)")

    add_h2("7.1 Justificación metodológica y emparejamiento de las unidades")
    add_p(
        "Directiva metodológica: El requerimiento de muestras relacionadas exige evaluar dos mediciones cuantitativas realizadas de forma "
        "indiscutible sobre las MISMAS unidades de análisis (diseño intra-sujeto o de medidas repetidas). En este trabajo, cada uno de los "
        "1567 repositorios de software en GitHub constituye una unidad de análisis única, sobre la cual coexisten de manera simultánea dos métricas "
        "críticas de interacción: el logaritmo de popularidad pasiva (ln_stars_1) y el logaritmo de reutilización activa (ln_forks_1)."
    )
    add_p(
        "Existe una justificación teórica y metodológica de emparejamiento rigurosa: ambas mediciones comparten idénticos factores intrínsecos al "
        "repositorio (calidad de arquitectura de software, dominio de aplicación, trayectoria cronológica y prestigio del autor). Por tanto, la "
        "diferencia intra-repositorio D_i = ln_stars_1 - ln_forks_1 anula la variabilidad entre proyectos distintos y permite aislar con máxima "
        "potencia inferencial la discrepancia intrínseca entre el interés meramente declarativo (stars) y el involucramiento técnico operativo (forks)."
    )

    add_h2("7.2 Pregunta de investigación y formulación de hipótesis")
    add_p(
        "Pregunta de investigación: ¿Existe una diferencia estadísticamente significativa entre el promedio de estrellas transformadas (ln_stars_1) "
        "y el promedio de bifurcaciones transformadas (ln_forks_1) evaluadas sobre los mismos repositorios de código abierto en Python?"
    )
    add_p("Hipótesis formulada (prueba bilateral sobre la media de las diferencias poblacionales μ_D, α = 0.05):")
    add_p("• H₀: μ_D = 0  (El promedio de las diferencias entre ln_stars_1 y ln_forks_1 es igual a cero; μ_stars = μ_forks).")
    add_p("• H₁: μ_D ≠ 0  (El promedio de las diferencias es distinto de cero; existe una discrepancia significativa entre estrellas y forks).")

    add_h2("7.3 Cálculo de diferencias emparejadas y verificación de normalidad")
    add_p(
        "Se calculó la variable de diferencia para cada repositorio: D_i = ln(stars_i + 1) - ln(forks_i + 1). "
        "La prueba de normalidad de Shapiro–Wilk sobre las diferencias arroja un estadístico W = 0.9662 (gl = 1567, p < 0.0001). "
        "Conforme a los postulados del Teorema del Límite Central para n = 1567 >> 30, la distribución muestral de la media de las diferencias "
        "D̄ converge a la normalidad de forma asintótica estricta, autorizando plenamente la prueba t para muestras emparejadas."
    )

    add_h2("7.4 Aplicación de la prueba t para muestras relacionadas en SPSS")
    add_p(
        "El procesamiento computacional ejecutado en IBM SPSS Statistics proporcionó los siguientes resultados paramétricos exactos:"
    )
    add_p("1. Estadísticas de muestras emparejadas:")
    add_p("• ln_stars_1: Media = 8.977201 (Desv. Est. = 0.503141; Error estándar = 0.012710; n = 1567).")
    add_p("• ln_forks_1: Media = 6.894743 (Desv. Est. = 0.772577; Error estándar = 0.019517; n = 1567).")
    add_p("2. Correlación de Pearson de muestras emparejadas:")
    add_p("• Correlación lineal r = 0.623 (p < 0.001 tanto para un factor como para dos factores). Existe una correlación positiva moderada-fuerte "
          "estadísticamente significativa que respalda formalmente la idoneidad del emparejamiento.")
    add_p("3. Prueba t para diferencias emparejadas (ln_stars_1 - ln_forks_1):")
    add_p("• Media de las diferencias (D̄): 2.082458 unidades logarítmicas.")
    add_p("• Desviación estándar de las diferencias (s_D): 0.604453.")
    add_p("• Error estándar de la media de diferencias (SE_D): 0.015270.")
    add_p("• Intervalo de confianza al 95 % de la diferencia: [2.052507 ; 2.112409].")
    add_p("• Estadístico t calculado: t = 136.379.")
    add_p("• Grados de libertad: gl = 1566.")
    add_p("• Significancia bilateral: p < 0.001 (p de un factor < 0.001).")

    add_h2("7.5 Decisión estadística, tamaño del efecto e interpretación en Ingeniería de Sistemas")
    add_p(
        "Decisión inferencial: Dado que el p-valor (p < 0.001) es abrumadoramente menor a α = 0.05, SE RECHAZA ROTUNDAMENTE LA HIPÓTESIS NULA (H₀) "
        "y se acepta la hipótesis alterna (H₁)."
    )
    add_p(
        "Tamaño del efecto para muestras emparejadas (calculado en SPSS):"
    )
    add_p("• d de Cohen: d = D̄ / s_D = 3.445 (IC 95 %: [3.315 ; 3.575]). Estandarizador: s_D = 0.604453.")
    add_p("• Corrección de Hedges: g = 3.444 (IC 95 %: [3.313 ; 3.574]). Estandarizador corregido: 0.604743.")
    add_p(
        "Interpretación profunda: Un valor de d de Cohen de 3.445 representa un tamaño del efecto extraordinariamente masivo (d >> 0.80), "
        "rara vez presenciado en estudios de ciencias sociales o ingeniería. En términos prácticos del ecosistema de software, una diferencia logarítmica "
        "media de 2.0825 unidades equivale a un factor multiplicativo de e^(2.0825) ≈ 8.02. Esto significa que, de manera sistemática e invariable "
        "dentro de un mismo proyecto, existen aproximadamente 8 estrellas por cada bifurcación creada."
    )
    add_p(
        "La explicación técnica en Ingeniería de Sistemas es transparente: otorgar una estrella es un acto de interacción pasiva y de costo nulo "
        "(un clic de aprobación o guardado como favorito), mientras que crear un fork demanda un esfuerzo computacional y cognitivo tangible "
        "(clonar el repositorio, configurar entornos, analizar la base de código y proyectar contribuciones). Los datos reales confirman esta asimetría estructural."
    )

    # Insert images from trabajo.docx for Section 7
    add_img("extracted_trabajo_media/image16.png", "Figura 19. Salida de SPSS: Estadísticas, Correlaciones (r = 0.623) y Prueba t de muestras emparejadas (t = 136.379, p < 0.001).")
    add_img("extracted_trabajo_media/image1.png", "Figura 20. Ayuda curada de SPSS y Tabla de Tamaños de Efecto de muestras emparejadas (d de Cohen = 3.445, g de Hedges = 3.444).")

    # Placeholders for future images in Section 7
    add_placeholder(
        "Prueba de Normalidad de las Diferencias Emparejadas (Shapiro–Wilk)",
        "Captura de pantalla de la tabla de Pruebas de Normalidad de SPSS para la variable diff = ln_stars_1 - ln_forks_1.",
        "Inserte aquí la salida de SPSS con el estadístico Shapiro-Wilk de la columna de diferencias calculada."
    )
    add_placeholder(
        "Gráfico Q-Q Normal / Histograma de las Diferencias Emparejadas",
        "Gráfico Q-Q o histograma de la distribución de las diferencias calculadas en SPSS.",
        "Inserte aquí el gráfico de dispersión Q-Q generado en SPSS para las diferencias emparejadas."
    )

    # =========================================================================
    # 8. ANÁLISIS DE VARIANZA (ANOVA) PARA TRES O MÁS GRUPOS
    # =========================================================================
    add_h1("8. Análisis de Varianza (ANOVA) para tres o más grupos independientes")

    add_h2("8.1 Planteamiento y definición de variables")
    add_p(
        "Para dar cumplimiento estricto a la sección de ANOVA de un factor para tres o más grupos independientes, se seleccionaron las siguientes "
        "variables del dataset real:"
    )
    add_p("• Variable Dependiente (Cuantitativa continua): ln_open_issues_1 = ln(open_issues + 1), que mide la tasa de incidencias de mantenimiento pendientes.")
    add_p("• Factor o Variable de Agrupación (Cualitativa policotómica de 3 niveles): Tipo de Licencia de Software (lic3), estructurada en:")
    add_p("   - Grupo 1 (lic3 = 1.00): Licencia MIT (n₁ = 546 repositorios) — Licencia altamente permisiva y minimalista.")
    add_p("   - Grupo 2 (lic3 = 2.00): Licencia Apache-2.0 (n₂ = 405 repositorios) — Licencia permisiva corporativa con cláusulas explícitas de patentes.")
    add_p("   - Grupo 3 (lic3 = 3.00): Licencia GPL-3.0 (n₃ = 112 repositorios) — Licencia de software libre con copyleft recíproco fuerte.")
    add_p("   - Tamaño total de la muestra bajo ANOVA: N = 1063 repositorios.")

    add_h2("8.2 Pregunta de investigación y formulación de hipótesis")
    add_p(
        "Pregunta de investigación: ¿Existen diferencias estadísticamente significativas en el nivel medio de incidencias abiertas transformadas "
        "(ln_open_issues_1) entre los repositorios de software en Python según su tipo de licencia (MIT, Apache-2.0 y GPL-3.0)?"
    )
    add_p("Hipótesis del Análisis de Varianza (ANOVA):")
    add_p("• Hipótesis Nula (H₀): μ_MIT = μ_Apache = μ_GPL  (Los promedios poblacionales de issues abiertos son iguales entre los tres tipos de licencias).")
    add_p("• Hipótesis Alterna (H₁): Al menos un promedio poblacional de issues abiertos es diferente entre los tipos de licencias.")
    add_p("• Nivel de significancia establecido: α = 0.05.")

    add_h2("8.3 Estadísticos descriptivos por grupo de licencia")
    add_p(
        "La Tabla 13 resume los descriptivos detallados por grupo de licencia procesados en SPSS y ratificados en Python:"
    )

    add_p("Tabla 13. Descriptivos de ln_open_issues_1 por tipo de licencia (N = 1063).", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t13_data = [
        ["Licencia (lic3)", "N", "Media", "Desv. Estándar", "Error Estándar", "IC 95 % Lím. Inf.", "IC 95 % Lím. Sup.", "Mín.", "Máx."],
        ["MIT (1.00)", "546", "4.0227", "1.4769", "0.0632", "3.8985", "4.1469", "0.0000", "8.2973"],
        ["Apache-2.0 (2.00)", "405", "4.5735", "1.3876", "0.0689", "4.4379", "4.7090", "0.0000", "7.8590"],
        ["GPL-3.0 (3.00)", "112", "4.3647", "1.3920", "0.1315", "4.1041", "4.6254", "0.0000", "7.1959"],
        ["Total", "1063", "4.2686", "1.4567", "0.0447", "4.1809", "4.3563", "0.0000", "8.2973"]
    ]
    create_custom_table(t13_data)

    add_h2("8.4 Verificación de supuestos: Normalidad y Homogeneidad de Varianzas (Levene)")
    add_p(
        "1. Supuesto de Normalidad por grupo (Shapiro–Wilk):"
    )
    add_p("• Grupo MIT: W = 0.9817 (p < 0.0001; n = 546).")
    add_p("• Grupo Apache-2.0: W = 0.9760 (p < 0.0001; n = 405).")
    add_p("• Grupo GPL-3.0: W = 0.9801 (p = 0.0930; n = 112). Cumple normalidad paramétrica estricta (p ≥ 0.05).")
    add_p(
        "Al amparo del Teorema del Límite Central y la amplia robustez del modelo lineal general para tamaños de celda balanceados y grandes "
        "(n ≥ 112 en todas las celdas), la ligera asimetría residual no compromete la validez del estadístico F."
    )
    add_p(
        "2. Supuesto de Homogeneidad de Varianzas (Prueba de Levene):"
    )
    add_p("En la salida de SPSS se verificó Levene bajo cuatro métricas:")
    add_p("• Se basa en la media: Estadístico de Levene = 1.109, gl1 = 2, gl2 = 1060, Sig. = 0.330.")
    add_p("• Se basa en la mediana: Estadístico = 0.984, gl1 = 2, gl2 = 1060, Sig. = 0.374.")
    add_p("• Se basa en la mediana y con gl ajustado: Estadístico = 0.984, gl1 = 2, gl2 = 1056.894, Sig. = 0.374.")
    add_p("• Se basa en la media recortada: Estadístico = 1.084, gl1 = 2, gl2 = 1060, Sig. = 0.339.")
    add_p(
        "Decisión sobre homocedasticidad: Como el p-valor de Levene basado en la media (p = 0.330) es marcadamente superior a α = 0.05, "
        "NO SE RECHAZA LA HOMOGENEIDAD DE VARIANZAS. Las varianzas de las tres poblaciones son estadísticamente homogéneas (iguales), "
        "satisfaciendo con excelencia el supuesto crítico del ANOVA clásico."
    )

    add_h2("8.5 Aplicación del ANOVA de un factor y prueba robusta de Welch")
    add_p(
        "La descomposición de la suma de cuadrados del ANOVA en IBM SPSS Statistics arrojó los siguientes resultados:"
    )
    add_p("Tabla 14. Tabla de Análisis de Varianza (ANOVA) de un factor para ln_open_issues_1.", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t14_data = [
        ["Fuente de variación", "Suma de cuadrados (SS)", "gl", "Media cuadrática (MS)", "Estadístico F", "Significancia (Sig.)"],
        ["Entre grupos (Tratamiento)", "71.698", "2", "35.849", "17.418", "< 0.001"],
        ["Dentro de grupos (Error)", "2181.693", "1060", "2.058", "—", "—"],
        ["Total", "2253.391", "1062", "—", "—", "—"]
    ]
    create_custom_table(t14_data)

    add_p(
        "Prueba robusta complementaria de Welch: Para garantizar máxima solidez analítica, se calculó el estadístico de Welch asintótico, "
        "arrojando F_Welch = 17.511 (gl1 = 2, gl2 = 314.565, p < 0.001), coincidiendo plenamente con el ANOVA ordinario."
    )

    add_h2("8.6 Estadístico F, p-valor, tamaños de efecto y decisión estadística")
    add_p(
        "Resultados inferenciales clave:"
    )
    add_p("• Estadístico F = 17.418 (gl_entre = 2, gl_dentro = 1060).")
    add_p("• p-valor asociado: p < 0.001.")
    add_p("• Tamaños del efecto calculados en SPSS:")
    add_p("   - Eta cuadrado (η²): η² = SS_entre / SS_total = 71.698 / 2253.391 = 0.032 (IC 95 %: [0.014 ; 0.054]).")
    add_p("   - Epsilon cuadrado (ε²): ε² = 0.030 (IC 95 %: [0.012 ; 0.052]).")
    add_p("   - Omega cuadrado efecto fijo (ω²): ω² = 0.030 (IC 95 %: [0.012 ; 0.052]).")
    add_p("   - Omega cuadrado efecto aleatorio: 0.015 (IC 95 %: [0.006 ; 0.027]).")
    add_p(
        "Decisión estadística: Dado que p-valor < 0.001 < 0.05, SE RECHAZA CONTUNDENTEMENTE LA HIPÓTESIS NULA (H₀) y se acepta la hipótesis alterna (H₁)."
    )

    add_h2("8.7 Interpretación del resultado en el contexto del software")
    add_p(
        "Interpretación: Existen diferencias altamente significativas entre los promedios de incidencias abiertas transformadas según el marco "
        "de licenciamiento de software. El tipo de licencia explica el 3.2 % de la varianza total en la carga de issues abiertos (η² = 0.032). "
        "Sin embargo, el ANOVA global únicamente señala que al menos un grupo difiere de los demás, pero no especifica qué licencias generan dicha discrepancia. "
        "Por ende, y en estricta coherencia con las pautas de la actividad, es preceptivo continuar con las Pruebas Post Hoc."
    )

    # Insert images from trabajo.docx for Section 8
    add_img("extracted_trabajo_media/image10.png", "Figura 21. Tabla de Descriptivos por tipo de licencia y Prueba de Homogeneidad de Varianzas de Levene (F = 1.109, p = 0.330) en SPSS.")
    add_img("extracted_trabajo_media/image3.png", "Figura 22. Tabla ANOVA de un factor (F = 17.418, p < 0.001), Tamaños de Efecto (Eta cuadrado = 0.032) y Prueba Robusta de Welch en SPSS.")

    # Placeholders for future images in Section 8
    add_placeholder(
        "Pruebas de Normalidad de Shapiro–Wilk por Grupo de Licencia en SPSS",
        "Captura de pantalla de la salida de SPSS con las pruebas de normalidad para cada uno de los 3 grupos de licencias.",
        "Inserte aquí la tabla de normalidad de SPSS generada con Explorar desglosada por factor lic3."
    )
    add_placeholder(
        "Gráfico de Medias del ANOVA de un Factor en SPSS",
        "Gráfico de líneas de medias con barras de error generado en SPSS para ln_open_issues_1 según tipo de licencia.",
        "Inserte aquí el gráfico de medias obtenido de la opción 'Gráfico de medias' del cuadro de diálogo ANOVA."
    )

    # =========================================================================
    # 9. PRUEBAS POST HOC
    # =========================================================================
    add_h1("9. Pruebas Post Hoc: Tukey o Games-Howell según corresponda")

    add_h2("9.1 Criterio de selección y justificación de la prueba Post Hoc")
    add_p(
        "Directiva de selección metodológica:"
    )
    add_p("• Si la prueba de Levene determina Varianzas Homogéneas (p ≥ 0.05) → Corresponde aplicar la prueba HSD de Tukey (Honest Significant Difference).")
    add_p("• Si la prueba de Levene determina Varianzas Heterogéneas (p < 0.05) → Corresponde aplicar la prueba de Games–Howell.")
    add_p(
        "Justificación de la elección: Como se constató en la Sección 8.4, el estadístico de Levene arrojó F = 1.109 con p = 0.330 ≥ 0.05, "
        "confirmando la perfecta homocedasticidad entre los grupos de licencias. Por lo tanto, se justifica plenamente y con rigor matemático el "
        "empleo de la prueba HSD de Tukey como procedimiento Post Hoc óptimo."
    )

    add_h2("9.2 Comparaciones múltiples por pares (HSD Tukey)")
    add_p(
        "La prueba HSD de Tukey ajusta las tasas de error tipo I para todas las comparaciones pareadas posibles. La Tabla 15 detalla la matriz "
        "completa de contrastes extraída de SPSS:"
    )

    add_p("Tabla 15. Comparaciones múltiples Post Hoc mediante HSD de Tukey para ln_open_issues_1.", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t15_data = [
        ["(I) Licencia", "(J) Licencia", "Diferencia de medias (I - J)", "Error estándar", "Sig. ajustada (p)", "IC 95 % Lím. Inferior", "IC 95 % Lím. Superior", "¿Significativo al 5 %?"],
        ["MIT (1.00)", "Apache-2.0 (2.00)", "-0.550792 *", "0.094083", "< 0.001", "-0.771605", "-0.329979", "Sí (Diferencia significativa)"],
        ["MIT (1.00)", "GPL-3.0 (3.00)", "-0.342007", "0.148817", "0.056", "-0.691280", "0.007266", "No (p ≥ 0.05 marginal)"],
        ["Apache-2.0 (2.00)", "MIT (1.00)", "0.550792 *", "0.094083", "< 0.001", "0.329979", "0.771605", "Sí (Diferencia significativa)"],
        ["Apache-2.0 (2.00)", "GPL-3.0 (3.00)", "0.208785", "0.153163", "0.361", "-0.150688", "0.568258", "No (No significativo)"],
        ["GPL-3.0 (3.00)", "MIT (1.00)", "0.342007", "0.148817", "0.056", "-0.007266", "0.691280", "No (p ≥ 0.05 marginal)"],
        ["GPL-3.0 (3.00)", "Apache-2.0 (2.00)", "-0.208785", "0.153163", "0.361", "-0.568258", "0.150688", "No (No significativo)"]
    ]
    create_custom_table(t15_data)
    add_p("*. La diferencia de medias es estadísticamente significativa en el nivel 0.05.", italic=True, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_h2("9.3 Diferencias de medias, errores estándar y p-valores ajustados")
    add_p("Análisis detallado de cada par:")
    add_p("1. Comparación MIT vs. Apache-2.0:")
    add_p("• Diferencia de medias: x̄_MIT - x̄_Apache = -0.550792 (Error estándar = 0.094083; p < 0.001; IC 95 %: [-0.7716 ; -0.3300]).")
    add_p("• Decisión: Altamente significativa. Los repositorios con licencia Apache-2.0 presentan en promedio 0.55 unidades logarítmicas más de incidencias "
          "abiertas que los de licencia MIT.")
    add_p("2. Comparación MIT vs. GPL-3.0:")
    add_p("• Diferencia de medias: x̄_MIT - x̄_GPL = -0.342007 (Error estándar = 0.148817; p = 0.056; IC 95 %: [-0.6913 ; 0.0073]).")
    add_p("• Decisión: No significativa al nivel α = 0.05. Aunque existe una ligera tendencia hacia mayor volumen en GPL-3.0, el p-valor de 0.056 "
          "supera la cota crítica convencional.")
    add_p("3. Comparación Apache-2.0 vs. GPL-3.0:")
    add_p("• Diferencia de medias: x̄_Apache - x̄_GPL = 0.208785 (Error estándar = 0.153163; p = 0.361; IC 95 %: [-0.1507 ; 0.5683]).")
    add_p("• Decisión: No significativa. No existe evidencia de diferencia entre los repositorios Apache-2.0 y GPL-3.0.")

    add_h2("9.4 Análisis de subconjuntos homogéneos")
    add_p(
        "La prueba HSD de Tukey sintetiza los resultados agrupando las categorías en subconjuntos homogéneos cuyas medias no difieren significativamente "
        "entre sí al nivel α = 0.05:"
    )

    add_p("Tabla 16. Subconjuntos homogéneos HSD de Tukey para ln_open_issues_1 (α = 0.05).", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
    t16_data = [
        ["Tipo de Licencia (lic3)", "N", "Subconjunto 1 (Baja carga de issues)", "Subconjunto 2 (Alta carga de issues)"],
        ["MIT (1.00)", "546", "4.022699", "—"],
        ["GPL-3.0 (3.00)", "112", "—", "4.364706"],
        ["Apache-2.0 (2.00)", "405", "—", "4.573491"],
        ["Significancia observada (p)", "—", "1.000", "0.268"]
    ]
    create_custom_table(t16_data)
    add_p("Se visualizan las medias para los grupos en los subconjuntos homogéneos basados en la media armónica de los tamaños muestrales (n_h = 226.771).", italic=True, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_h2("9.5 Interpretación contextual en Ingeniería de Software")
    add_p(
        "El análisis de subconjuntos homogéneos revela una clara bifurcación en la gestión de proyectos de código abierto:"
    )
    add_p(
        "• El Subconjunto 1 está conformado en exclusiva por la Licencia MIT (media = 4.0227). La licencia MIT es el epítome del minimalismo normativo: "
        "otorga permisos irrestrictos de uso, copia y modificación con mínimas obligaciones formales. Esto propicia repositorios livianos, herramientas "
        "concisas y bibliotecas modulares donde el flujo de incidencias pendientes se resuelve o depura con agilidad, manteniendo una baja carga de issues abiertos."
    )
    add_p(
        "• El Subconjunto 2 agrupa a las Licencias GPL-3.0 (media = 4.3647) y Apache-2.0 (media = 4.5735), cuya diferencia mutua no es significativa (p = 0.268). "
        "Tanto Apache-2.0 como GPL-3.0 son marcos de licenciamiento altamente estructurados. Apache-2.0 es adoptada mayoritariamente por consorcios empresariales "
        "y proyectos de gran infraestructura de software (como el ecosistema Apache Foundation), caracterizados por estrictos procesos de gobernanza, "
        "auditoría de código, requerimientos corporativos complejos y prolongados debates técnicos que mantienen abiertas incidencias durante meses o años. "
        "Por su parte, GPL-3.0 exige el cumplimiento de copyleft recíproco y compatibilidad de dependencias, atrayendo discusiones comunitarias extensas sobre "
        "licenciamiento y arquitectura."
    )
    add_p(
        "En conclusión, la diferencia significativa detectada por el ANOVA no es genérica ni azarosa: obedece a la distancia estructural entre la agilidad "
        "minimalista de MIT frente a la gobernanza corporativa e institucional de Apache-2.0."
    )

    # Insert images from trabajo.docx for Section 9
    add_img("extracted_trabajo_media/image13.png", "Figura 23. Tabla de Comparaciones Múltiples Post Hoc HSD de Tukey para ln_open_issues_1 en SPSS.")
    add_img("extracted_trabajo_media/image4.png", "Figura 24. Tabla de Subconjuntos Homogéneos HSD de Tukey clasificando a MIT en el Subconjunto 1 y a GPL-3.0 y Apache-2.0 en el Subconjunto 2.")

    # Placeholders for future images in Section 9
    add_placeholder(
        "Diagrama de Cajas Comparativo Post Hoc por Tipo de Licencia",
        "Boxplot de ln_open_issues_1 desglosado por MIT, Apache-2.0 y GPL-3.0 generado en SPSS.",
        "Inserte aquí el diagrama de cajas que visualice la distribución y medianas de cada subconjunto homogéneo."
    )

    # =========================================================================
    # 10. CONCLUSIONES
    # =========================================================================
    add_h1("10. Conclusiones")

    add_p(
        "Las conclusiones del presente estudio se sustentan estrictamente en la evidencia estadística inferencial obtenida a partir del análisis "
        "de 1567 repositorios reales de software libre en Python extraídos de GitHub:"
    )

    conclusions = [
        ("1. Respecto a la popularidad de los repositorios frente al estándar de una muestra (Q1):",
         " No existe evidencia estadística suficiente para rechazar que el promedio poblacional de estrellas de los repositorios de Python coincida con el valor de referencia de 9000 estrellas (t = -0.324, gl = 1566, p = 0.746). La diferencia de -36.12 estrellas es insignificante y ostenta un tamaño del efecto nulo (d = -0.008), lo que confirma con 95 % de confianza que la muestra analizada representa a un estrato altamente consolidado de la comunidad tecnológica. Asimismo, frente al estándar mínimo de 5000 estrellas, los repositorios superan con creces dicha cota (p < 0.001, d = 0.90)."),

        ("2. Respecto a la comparación de bifurcaciones entre grupos independientes (Q2):",
         " Al comparar las bifurcaciones transformadas (ln_forks_1) según el tipo de propietario (Organizaciones corporativas vs. Usuarios individuales), la prueba t de Welch demostró que no existen diferencias estadísticamente significativas (t = 1.107, gl = 1425.17, p = 0.268, d = 0.057). Esto evidencia que la reutilización técnica de código abierto ocurre a tasas similares con independencia de la naturaleza organizativa del autor, aunque en términos absolutos de difusión masiva (stars y volumen bruto de incidencias), las organizaciones exhiben una ventaja moderada por su infraestructura."),

        ("3. Respecto a la relación intra-repositorio en muestras emparejadas (Q3):",
         " La comparación de muestras relacionadas reveló una diferencia abrumadoramente significativa entre las estrellas (ln_stars_1) y las bifurcaciones (ln_forks_1) evaluadas sobre los mismos 1567 repositorios (t = 136.379, gl = 1566, p < 0.001). Con un tamaño del efecto colosal (d de Cohen = 3.445) y una diferencia media de 2.0825 unidades logarítmicas, se formaliza la existencia de una relación aproximada de 8 estrellas por cada bifurcación, confirmando que el respaldo comunitario pasivo es exponencialmente más frecuente que la reutilización activa."),

        ("4. Respecto a la heterogeneidad de incidencias entre licencias de software en el ANOVA (Q4):",
         " El Análisis de Varianza de un factor demostró la existencia de diferencias estadísticamente significativas en el volumen de incidencias abiertas transformadas (ln_open_issues_1) entre los diferentes modelos de licenciamiento (F = 17.418, gl = (2, 1060), p < 0.001, η² = 0.032). Se rechaza con certeza matemática la hipótesis de igualdad de medias, comprobándose que el marco jurídico de derechos de autor influye de manera apreciable en la dinámica de mantenimiento del proyecto."),

        ("5. Respecto a la localización de las diferencias mediante pruebas Post Hoc (Q5):",
         " Habiéndose constatado la homogeneidad de varianzas mediante la prueba de Levene (F = 1.109, p = 0.330), la prueba HSD de Tukey reveló que la discrepancia significativa se focaliza de forma exclusiva entre la Licencia Apache-2.0 y la Licencia MIT (diferencia de medias = 0.5508, p < 0.001). El análisis de subconjuntos homogéneos diferenció a MIT en un clúster de baja carga de incidencias frente a Apache-2.0 y GPL-3.0, reflejando que los entornos de gobernanza corporativa acumulan una mayor deuda técnica que los proyectos permisivos minimalistas."),

        ("6. Respecto a la validez metodológica y el Teorema del Límite Central:",
         " Si bien los contrastes de Shapiro–Wilk rechazaron la normalidad estricta en las variables originales (p < 0.001), la investigación validó con rigor la pertinencia de las pruebas paramétricas sustentándose en el Teorema del Límite Central (n = 1567), la aplicación de transformaciones estabilizadoras ln(X + 1), y la selección de variantes robustas (t de Welch) ante heterocedasticidad, garantizando conclusiones científicas replicables y sólidas.")
    ]

    for c_title, c_text in conclusions:
        p_c = doc.add_paragraph()
        p_c.paragraph_format.space_before = Pt(3)
        p_c.paragraph_format.space_after = Pt(5)
        p_c.paragraph_format.line_spacing = 1.15
        r_c1 = p_c.add_run(f"• {c_title}")
        r_c1.font.name = 'Times New Roman'
        r_c1.font.size = Pt(11)
        r_c1.font.bold = True
        r_c1.font.color.rgb = NAVY
        r_c2 = p_c.add_run(f" {c_text}")
        r_c2.font.name = 'Times New Roman'
        r_c2.font.size = Pt(11)

    # =========================================================================
    # 11. REFERENCIAS Y FUENTES DE DATOS
    # =========================================================================
    add_h1("11. Referencias y fuentes de datos")

    refs = [
        "Cohen, J. (1988). Statistical Power Analysis for the Behavioral Sciences (2nd ed.). Lawrence Erlbaum Associates.",
        "Hedges, L. V. (1981). Distribution theory for Glass's estimator of effect size and related estimators. Journal of Educational Statistics, 6(2), 107–128. https://doi.org/10.3102/10769986006002107",
        "IBM Corp. (2023). IBM SPSS Statistics for Windows, Version 29.0. Armonk, NY: IBM Corp.",
        "Kaggle Datasets. (2024). GitHub Repositories – Heterogeneous Data. Extraído del repositorio público de datos de Kaggle: https://www.kaggle.com/datasets",
        "Montgomery, D. C., & Runger, G. C. (2018). Probabilidad y estadística aplicadas a la ingeniería (5.ª ed.). Limusa Wiley.",
        "SciPy Community. (2024). Statistical functions (scipy.stats). Documentación oficial de SciPy. https://docs.scipy.org/doc/scipy/reference/stats.html",
        "Statsmodels Developers. (2024). Statsmodels: Econometric and statistical modeling with Python. https://www.statsmodels.org/",
        "Triola, M. F. (2018). Estadística (12.ª ed.). Pearson Educación."
    ]

    for ref in refs:
        p_r = doc.add_paragraph()
        p_r.paragraph_format.space_before = Pt(2)
        p_r.paragraph_format.space_after = Pt(4)
        p_r.paragraph_format.left_indent = Inches(0.5)
        p_r.paragraph_format.first_line_indent = Inches(-0.5)
        p_r.paragraph_format.line_spacing = 1.15
        r_r = p_r.add_run(ref)
        r_r.font.name = 'Times New Roman'
        r_r.font.size = Pt(11)

    # Save document
    output_filename = "INFORME_FINAL_APLICACION_DE_ESTADISTICA_INFERENCIAL.docx"
    doc.save(output_filename)
    print(f"Document successfully created and saved as: {output_filename}")

if __name__ == '__main__':
    build_report()
