# -*- coding: utf-8 -*-
"""
Construye el documento Word de Ejercicios de Prueba de Wilcoxon (Original y Resuelto)
preservando al 100% el formato institucional de la plantilla original del docente
(encabezado, tipografía Calibri 11pt, banner gris #D4D4D4, numeración de lista y tablas con bordes)
y exporta automáticamente a PDF mediante Microsoft Word COM.
"""

import os
import copy
import shutil
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.parts.document import DocumentPart

# Evitar colisión de rId / shape id con encabezados
def _custom_next_id(self):
    used_ids = set()
    for part in self.package.parts:
        if hasattr(part, '_element') and part._element is not None:
            for s in part._element.xpath('//@id'):
                if s.isdigit():
                    val = int(s)
                    if val < 1000000:
                        used_ids.add(val)
    return (max(used_ids) + 1) if used_ids else 2000

DocumentPart.next_id = property(_custom_next_id)

template_file = "Ejercicios U Mann-Whitney_Original.docx"
out_orig_docx = "Ejercicios Wilcoxon_Original.docx"
out_res_docx = "Ejercicios Wilcoxon_Resuelto.docx"
out_main_docx = "Ejercicios Wilcoxon.docx"
out_res_pdf = "Ejercicios Wilcoxon_Resuelto.pdf"
img_dir = "generated_wilcoxon_images"

def create_wilcoxon_document(include_solutions=True):
    doc = docx.Document(template_file)
    
    # Guardamos copias de los elementos XML modelo de la plantilla del docente:
    # P1: Título centrado en negrita 14pt
    # P3: Banner con sombreado gris #D4D4D4
    # P4: Párrafo numerado (numId=29, ilvl=0)
    # P5: Párrafo con sangría izquierda (ind left=360) para preguntas debajo de tablas
    # Table 1: Tabla de 2 filas con bordes exactos y sangría izquierda (tblInd=426)
    p1_pPr_xml = copy.deepcopy(doc.paragraphs[1]._p.pPr)
    p3_pPr_xml = copy.deepcopy(doc.paragraphs[3]._p.pPr)
    p4_pPr_xml = copy.deepcopy(doc.paragraphs[4]._p.pPr)
    p5_pPr_xml = copy.deepcopy(doc.paragraphs[5]._p.pPr)
    
    tbl_template = copy.deepcopy(doc.tables[1]._tbl)
    
    # Limpiamos todos los párrafos y tablas del cuerpo excepto sectPr
    body = doc._body._element
    for child in list(body):
        tag = child.tag.split("}")[-1]
        if tag in ("p", "tbl"):
            body.remove(child)
            
    # Función auxiliar para añadir párrafo con formato exacto
    def add_custom_p(text="", bold_prefix="", pPr_template=None, space_before=4, space_after=4, align=WD_ALIGN_PARAGRAPH.JUSTIFY, font_size=11, bold=False, indent_left=Inches(0.25)):
        p = doc.add_paragraph()
        if pPr_template is not None:
            p._p.remove(p._p.pPr) if p._p.pPr is not None else None
            p._p.insert(0, copy.deepcopy(pPr_template))
        else:
            p.alignment = align
            p.paragraph_format.space_before = Pt(space_before)
            p.paragraph_format.space_after = Pt(space_after)
            p.paragraph_format.line_spacing = 1.15
            if indent_left is not None:
                p.paragraph_format.left_indent = indent_left
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.bold = True
            rb.font.name = "Calibri"
            rb.font.size = Pt(font_size)
        if text:
            rt = p.add_run(text)
            rt.bold = bold
            rt.font.name = "Calibri"
            rt.font.size = Pt(font_size)
        return p

    # Función para construir una tabla de datos idéntica a la del PDF del docente
    def add_data_table(label_row1, vals_row1, label_row2, vals_row2):
        # Construir tabla a partir de python-docx y aplicar propiedades XML exactas del docente
        ncols = 1 + len(vals_row1)
        tbl = doc.add_table(rows=2, cols=ncols)
        
        # Copiar tblPr de la tabla del docente (tblInd=426, tblCellMar left/right=70)
        tbl._tbl.remove(tbl._tbl.tblPr)
        tbl._tbl.insert(0, copy.deepcopy(tbl_template.tblPr))
        if tbl._tbl.tblGrid is not None:
            tbl._tbl.remove(tbl._tbl.tblGrid)
        
        rows_data = [
            [label_row1] + [str(v) for v in vals_row1],
            [label_row2] + [str(v) for v in vals_row2]
        ]
        
        for r_idx, r_vals in enumerate(rows_data):
            row = tbl.rows[r_idx]
            for c_idx, val_str in enumerate(r_vals):
                cell = row.cells[c_idx]
                tcPr = cell._tc.get_or_add_tcPr()
                for old_w in tcPr.xpath("./w:tcW"):
                    tcPr.remove(old_w)
                tcPr.insert(0, parse_xml(f'<w:tcW {nsdecls("w")} w:w="0" w:type="auto"/>'))
                # Bordes simples finos negros como en el PDF original
                borders_xml = parse_xml(
                    f'<w:tcBorders {nsdecls("w")}>'
                    '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                    '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                    '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                    '</w:tcBorders>'
                )
                tcPr.append(borders_xml)
                tcPr.append(parse_xml(f'<w:noWrap {nsdecls("w")}/>'))
                tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.05
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(val_str)
                r.font.name = "Calibri"
                r.font.size = Pt(11)
        return tbl

    def add_img(img_name, width_inches=4.2):
        img_path = os.path.join(img_dir, img_name)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run()
        r.add_picture(img_path, width=Inches(width_inches))
        return p

    # --- ENCABEZADO DEL DOCUMENTO ---
    add_custom_p("", pPr_template=doc.paragraphs[0]._p.pPr if len(doc.paragraphs) > 0 else None, space_before=0, space_after=0, indent_left=None)
    
    p_title1 = add_custom_p("Ejercicios", pPr_template=p1_pPr_xml)
    p_title2 = add_custom_p("Prueba de Wilcoxon", pPr_template=p1_pPr_xml)
    
    add_custom_p("", space_before=0, space_after=4, indent_left=None)
    
    # Banner sombreado gris #D4D4D4
    add_custom_p("Prueba Wilcoxon para muestras relacionadas", pPr_template=p3_pPr_xml)

    # =========================================================================
    # EJERCICIO 1
    # =========================================================================
    add_custom_p(
        "Un equipo de ingeniería de sistemas ha realizado una optimización en un algoritmo de procesamiento de datos para mejorar su eficiencia. Se generan dos muestras relacionadas, una para el tiempo de ejecución antes de la optimización y otra para el tiempo de ejecución después de la optimización. Se quiere determinar si esta optimización ha tenido un impacto significativo en el tiempo de ejecución del algoritmo.",
        pPr_template=p4_pPr_xml
    )
    add_data_table(
        "Tiempo de ejecución antes de la optimización", [15, 14, 16, 17, 18, 15, 16, 23, 15, 14],
        "Tiempo de ejecución después de la optimización", [10, 11, 12, 10, 11, 13, 12, 10, 11, 12]
    )

    if include_solutions:
        add_custom_p(bold_prefix="Solución:", space_before=6, space_after=2)
        add_custom_p(bold_prefix="Paso 1: Formación de hipótesis", space_before=3, space_after=2)
        add_custom_p("Los promedios (medianas) del tiempo de ejecución antes y después de la optimización son iguales (x̄Antes = x̄Después)", bold_prefix="H0: ", space_before=1, space_after=2)
        add_custom_p("Los promedios (medianas) del tiempo de ejecución antes y después de la optimización son diferentes (x̄Antes ≠ x̄Después)", bold_prefix="H1: ", space_before=1, space_after=3)
        
        add_custom_p(bold_prefix="Paso 2: Determinar el nivel de significancia", space_before=3, space_after=2)
        add_custom_p("N.S = 5% (0.05)", space_before=1, space_after=3)
        
        add_custom_p(bold_prefix="Paso 3: Hallar estadísticos descriptivos", space_before=3, space_after=2)
        add_img("ex1_informe.png", width_inches=3.6)
        
        add_custom_p(bold_prefix="Paso 4: Prueba de Distribución normal", space_before=3, space_after=2)
        add_custom_p("Los datos de los grupos Antes y Después de la optimización (y sus diferencias) presentan distribución normal", bold_prefix="H0: ", space_before=1, space_after=2)
        add_custom_p("Por lo menos los datos de un grupo Antes o Después de la optimización (o sus diferencias) no presentan distribución normal", bold_prefix="H1: ", space_before=1, space_after=2)
        add_img("ex1_normalidad.png", width_inches=4.3)
        add_custom_p("Dado que el p-valor del grupo Después de la optimización (0.191) es mayor que el nivel de significancia, pero el grupo Antes de la optimización (0.008) y la diferencia emparejada (0.028) son menores al nivel de significancia (0.05) por lo que es probable que se rechace la hipótesis nula", space_before=2, space_after=2)
        add_custom_p("A un margen de error de 5 % se concluye que , los datos de los grupos Antes y Después de la optimización, no presentan Distribución normal", bold_prefix="Conclusión: ", space_before=2, space_after=3)
        
        add_custom_p(bold_prefix="Paso 5: Elegir la prueba estadística", space_before=3, space_after=2)
        add_custom_p("Dado que los datos del grupo Antes de la optimización y las diferencias no presentan distribución normal se aplicara la prueba de los rangos con signo de Wilcoxon (Prueba no paramétrica)", space_before=1, space_after=2)
        add_img("ex1_rangos.png", width_inches=4.8)
        add_img("ex1_test.png", width_inches=3.5)
        
        add_custom_p(bold_prefix="Paso 6: Contrastación de Hipótesis", space_before=3, space_after=2)
        add_custom_p("Dado que el p -valor de Wilcoxon (0.005) es menor que el nivel de significancia (0.05), se rechaza la hipótesis nula.", space_before=1, space_after=3)
        
        add_custom_p(bold_prefix="Paso 7: Conclusión", space_before=3, space_after=2)
        add_custom_p("A un margen de error del 5% se concluye que, la optimización ha tenido un impacto significativo en el tiempo de ejecución del algoritmo.", space_before=1, space_after=8)

    # =========================================================================
    # EJERCICIO 2
    # =========================================================================
    add_custom_p(
        "Un equipo de ingeniería de sistemas quiere determinar si la actualización de hardware de un servidor ha reducido significativamente el tiempo de respuesta para procesar solicitudes. A un margen de error del 5% ¿Hay diferencias significativas respecto a los tiempos de respuesta (en milisegundos) antes y después de la actualización de hardware para el mismo conjunto de solicitudes?",
        pPr_template=p4_pPr_xml
    )
    add_data_table(
        "Tiempos de Respuesta Antes de la Actualización:", [80, 75, 78, 82, 79, 81, 77, 71, 80, 78],
        "Tiempos de Respuesta Después de la Actualización:", [70, 68, 69, 71, 67, 70, 68, 69, 70, 69]
    )

    if include_solutions:
        add_custom_p(bold_prefix="Solución:", space_before=6, space_after=2)
        add_custom_p(bold_prefix="Paso 1: Formación de hipótesis", space_before=3, space_after=2)
        add_custom_p("Los promedios (medianas) de los tiempos de respuesta antes y después de la actualización de hardware son iguales (x̄Antes = x̄Después)", bold_prefix="H0: ", space_before=1, space_after=2)
        add_custom_p("Los promedios (medianas) de los tiempos de respuesta antes y después de la actualización de hardware son diferentes (x̄Antes ≠ x̄Después)", bold_prefix="H1: ", space_before=1, space_after=3)
        
        add_custom_p(bold_prefix="Paso 2: Determinar el nivel de significancia", space_before=3, space_after=2)
        add_custom_p("N.S = 5% (0.05)", space_before=1, space_after=3)
        
        add_custom_p(bold_prefix="Paso 3: Hallar estadísticos descriptivos", space_before=3, space_after=2)
        add_img("ex2_informe.png", width_inches=3.6)
        
        add_custom_p(bold_prefix="Paso 4: Prueba de Distribución normal", space_before=3, space_after=2)
        add_custom_p("Los datos de los grupos Antes y Después de la actualización (y sus diferencias) presentan distribución normal", bold_prefix="H0: ", space_before=1, space_after=2)
        add_custom_p("Por lo menos los datos de un grupo (o las diferencias emparejadas) no presentan distribución normal", bold_prefix="H1: ", space_before=1, space_after=2)
        add_img("ex2_normalidad.png", width_inches=4.3)
        add_custom_p("Dado que los p-valores de los grupos Antes (0.308) y Después (0.691) son mayores que el nivel de significancia, pero el grupo de Diferencias (0.020) es menor al nivel de significancia (0.05) por lo que es probable que se rechace la hipótesis nula", space_before=2, space_after=2)
        add_custom_p("A un margen de error de 5 % se concluye que , las diferencias de los tiempos de respuesta antes y después de la actualización, no presentan Distribución normal", bold_prefix="Conclusión: ", space_before=2, space_after=3)
        
        add_custom_p(bold_prefix="Paso 5: Elegir la prueba estadística", space_before=3, space_after=2)
        add_custom_p("Dado que las diferencias de los datos no presentan distribución normal se aplicara la prueba de los rangos con signo de Wilcoxon (Prueba no paramétrica)", space_before=1, space_after=2)
        add_img("ex2_rangos.png", width_inches=4.8)
        add_img("ex2_test.png", width_inches=3.5)
        
        add_custom_p(bold_prefix="Paso 6: Contrastación de Hipótesis", space_before=3, space_after=2)
        add_custom_p("Dado que el p -valor de Wilcoxon (0.005) es menor que el nivel de significancia (0.05), se rechaza la hipótesis nula.", space_before=1, space_after=3)
        
        add_custom_p(bold_prefix="Paso 7: Conclusión", space_before=3, space_after=2)
        add_custom_p("A un margen de error del 5% se concluye que, existen diferencias significativas respecto a los tiempos de respuesta antes y después de la actualización de hardware.", space_before=1, space_after=8)

    # =========================================================================
    # EJERCICIO 3
    # =========================================================================
    add_custom_p(
        "Se ha implementado un nuevo protocolo de comunicación en la red y se desea comparar el tiempo de latencia (en milisegundos) antes y después de la implementación del protocolo:",
        pPr_template=p4_pPr_xml
    )
    add_data_table(
        "Antes de la implementación:", [150, 160, 155, 158, 152, 157, 159],
        "Después de la implementación:", [140, 145, 142, 144, 153, 143, 141]
    )
    add_custom_p(
        "¿El tiempo de latencia es menor después de la implementación del nuevo protocolo de comunicación que antes?",
        pPr_template=p5_pPr_xml
    )

    if include_solutions:
        add_custom_p(bold_prefix="Solución:", space_before=6, space_after=2)
        add_custom_p(bold_prefix="Paso 1: Formación de hipótesis", space_before=3, space_after=2)
        add_custom_p("El tiempo de latencia no es menor después de la implementación del nuevo protocolo de comunicación que antes (x̄Después ≥ x̄Antes)", bold_prefix="H0: ", space_before=1, space_after=2)
        add_custom_p("El tiempo de latencia es menor después de la implementación del nuevo protocolo de comunicación que antes (x̄Después < x̄Antes)", bold_prefix="H1: ", space_before=1, space_after=3)
        
        add_custom_p(bold_prefix="Paso 2: Determinar el nivel de significancia", space_before=3, space_after=2)
        add_custom_p("N.S = 5% (0.05)", space_before=1, space_after=3)
        
        add_custom_p(bold_prefix="Paso 3: Hallar estadísticos descriptivos", space_before=3, space_after=2)
        add_img("ex3_informe.png", width_inches=3.6)
        
        add_custom_p(bold_prefix="Paso 4: Prueba de Distribución normal", space_before=3, space_after=2)
        add_custom_p("Los datos de los grupos Antes y Después de la implementación (y sus diferencias) presentan distribución normal", bold_prefix="H0: ", space_before=1, space_after=2)
        add_custom_p("Por lo menos los datos de un grupo (o las diferencias emparejadas) no presentan distribución normal", bold_prefix="H1: ", space_before=1, space_after=2)
        add_img("ex3_normalidad.png", width_inches=4.0)
        add_custom_p("Dado que los p-valores de los grupos Antes (0.546) y Después (0.067) son mayores que el nivel de significancia, pero el grupo de Diferencias (0.044) es menor al nivel de significancia (0.05) por lo que es probable que se rechace la hipótesis nula", space_before=2, space_after=2)
        add_custom_p("A un margen de error de 5 % se concluye que , las diferencias de los tiempos de latencia antes y después de la implementación, no presentan Distribución normal", bold_prefix="Conclusión: ", space_before=2, space_after=3)
        
        add_custom_p(bold_prefix="Paso 5: Elegir la prueba estadística", space_before=3, space_after=2)
        add_custom_p("Dado que las diferencias de los datos no presentan distribución normal se aplicara la prueba de los rangos con signo de Wilcoxon (Prueba no paramétrica)", space_before=1, space_after=2)
        add_img("ex3_rangos.png", width_inches=4.5)
        add_img("ex3_test.png", width_inches=3.3)
        
        add_custom_p(bold_prefix="Paso 6: Contrastación de Hipótesis", space_before=3, space_after=2)
        add_custom_p("Dado que el p -valor de Wilcoxon (0.028 bilateral; 0.014 unilateral) es menor que el nivel de significancia (0.05), se rechaza la hipótesis nula.", space_before=1, space_after=3)
        
        add_custom_p(bold_prefix="Paso 7: Conclusión", space_before=3, space_after=2)
        add_custom_p("A un margen de error del 5% se concluye que, el tiempo de latencia es menor después de la implementación del nuevo protocolo de comunicación que antes.", space_before=1, space_after=8)

    # =========================================================================
    # EJERCICIO 4
    # =========================================================================
    add_custom_p(
        "Se han aplicado mejoras en un algoritmo de inteligencia artificial, y se ha registrado la precisión (%) del algoritmo antes y después de la mejora:",
        pPr_template=p4_pPr_xml
    )
    add_data_table(
        "Antes de la mejora:", [80, 72, 81, 79, 83, 81, 80],
        "Después de la mejora:", [90, 91, 92, 89, 93, 90, 91]
    )
    add_custom_p(
        "¿Es mayor la precisión del algoritmo después de la mejora que antes?",
        pPr_template=p5_pPr_xml
    )

    if include_solutions:
        add_custom_p(bold_prefix="Solución:", space_before=6, space_after=2)
        add_custom_p(bold_prefix="Paso 1: Formación de hipótesis", space_before=3, space_after=2)
        add_custom_p("La precisión del algoritmo no es mayor después de la mejora que antes (x̄Después ≤ x̄Antes)", bold_prefix="H0: ", space_before=1, space_after=2)
        add_custom_p("La precisión del algoritmo es mayor después de la mejora que antes (x̄Después > x̄Antes)", bold_prefix="H1: ", space_before=1, space_after=3)
        
        add_custom_p(bold_prefix="Paso 2: Determinar el nivel de significancia", space_before=3, space_after=2)
        add_custom_p("N.S = 5% (0.05)", space_before=1, space_after=3)
        
        add_custom_p(bold_prefix="Paso 3: Hallar estadísticos descriptivos", space_before=3, space_after=2)
        add_img("ex4_informe.png", width_inches=3.6)
        
        add_custom_p(bold_prefix="Paso 4: Prueba de Distribución normal", space_before=3, space_after=2)
        add_custom_p("Los datos de los grupos Antes y Después de la mejora (y sus diferencias) presentan distribución normal", bold_prefix="H0: ", space_before=1, space_after=2)
        add_custom_p("Por lo menos los datos de un grupo Antes o Después de la mejora (o sus diferencias) no presentan distribución normal", bold_prefix="H1: ", space_before=1, space_after=2)
        add_img("ex4_normalidad.png", width_inches=4.0)
        add_custom_p("Dado que el p-valor del grupo Después de la mejora (0.873) es mayor que el nivel de significancia, pero el grupo Antes de la mejora (0.028) y la diferencia emparejada (0.001) son menores al nivel de significancia (0.05) por lo que es probable que se rechace la hipótesis nula", space_before=2, space_after=2)
        add_custom_p("A un margen de error de 5 % se concluye que , los datos de los grupos Antes y Después de la mejora, no presentan Distribución normal", bold_prefix="Conclusión: ", space_before=2, space_after=3)
        
        add_custom_p(bold_prefix="Paso 5: Elegir la prueba estadística", space_before=3, space_after=2)
        add_custom_p("Dado que los datos del grupo Antes de la mejora y las diferencias no presentan distribución normal se aplicara la prueba de los rangos con signo de Wilcoxon (Prueba no paramétrica)", space_before=1, space_after=2)
        add_img("ex4_rangos.png", width_inches=4.3)
        add_img("ex4_test.png", width_inches=3.2)
        
        add_custom_p(bold_prefix="Paso 6: Contrastación de Hipótesis", space_before=3, space_after=2)
        add_custom_p("Dado que el p -valor de Wilcoxon (0.017 bilateral; 0.009 unilateral) es menor que el nivel de significancia (0.05), se rechaza la hipótesis nula.", space_before=1, space_after=3)
        
        add_custom_p(bold_prefix="Paso 7: Conclusión", space_before=3, space_after=2)
        add_custom_p("A un margen de error del 5% se concluye que, es mayor la precisión del algoritmo después de la mejora que antes.", space_before=1, space_after=8)

    return doc

if __name__ == "__main__":
    doc_orig = create_wilcoxon_document(include_solutions=False)
    doc_orig.save(out_orig_docx)
    print(f"Guardado documento base: {out_orig_docx}")

    doc_res = create_wilcoxon_document(include_solutions=True)
    doc_res.save(out_res_docx)
    shutil.copy2(out_res_docx, out_main_docx)
    print(f"Guardado documento resuelto: {out_res_docx} y {out_main_docx}")

    # Exportar a PDF nativo con Word COM
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        abs_docx = os.path.abspath(out_res_docx)
        abs_pdf = os.path.abspath(out_res_pdf)
        wdoc = word.Documents.Open(abs_docx)
        wdoc.SaveAs(abs_pdf, FileFormat=17)
        wdoc.Close(False)
        print(f"Exportado PDF exitosamente: {out_res_pdf}")
    finally:
        word.Quit()
