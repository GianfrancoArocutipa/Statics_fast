import os
import shutil
import docx
from docx.shared import Inches, Pt, Emu
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
import win32com.client

BASE_DIR = r"c:\Users\PC\Desktop\Estats"
SRC_DOCX = os.path.join(BASE_DIR, "Ejercicios Chi cuadrado 2da. parte (1) (1).docx")
BACKUP_DOCX = os.path.join(BASE_DIR, "Ejercicios Chi cuadrado 2da. parte_Original.docx")
RESUELTO_DOCX = os.path.join(BASE_DIR, "Ejercicios Chi cuadrado 2da. parte_Resuelto.docx")
RESUELTO_PDF = os.path.join(BASE_DIR, "Ejercicios Chi cuadrado 2da. parte_Resuelto.pdf")
IMG_DIR = os.path.join(BASE_DIR, "generated_chi2_part2_images")

if not os.path.exists(BACKUP_DOCX):
    shutil.copy2(SRC_DOCX, BACKUP_DOCX)


def set_run_font(run, bold=False, subscript=False):
    run.font.name = "Calibri"
    run.font.size = Pt(11)
    run.bold = bold
    run.font.subscript = subscript
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), "Calibri")
    rFonts.set(qn("w:hAnsi"), "Calibri")
    rFonts.set(qn("w:cs"), "Calibri")


def format_solution_paragraph(p, bold=False):
    pPr = p._p.get_or_add_pPr()
    # Remove any numPr or pStyle if inherited
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is not None:
        pPr.remove(pStyle)
    numPr = pPr.find(qn("w:numPr"))
    if numPr is not None:
        pPr.remove(numPr)

    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    spacing.set(qn("w:before"), "120")
    spacing.set(qn("w:after"), "120")
    spacing.set(qn("w:line"), "276")
    spacing.set(qn("w:lineRule"), "auto")

    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    ind.set(qn("w:left"), "284")
    for attr in ["w:right", "w:hanging", "w:firstLine"]:
        if ind.get(qn(attr)) is not None:
            del ind.attrib[qn(attr)]

    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), "both")


def insert_paragraph_after(ref_p, text="", bold=False, hypothesis=None, multiline_items=None, img_path=None, exact_emu_width=None):
    new_p = OxmlElement("w:p")
    ref_p._p.addnext(new_p)
    p = docx.text.paragraph.Paragraph(new_p, ref_p._parent)
    format_solution_paragraph(p, bold=bold)

    if hypothesis is not None:
        # hypothesis = ("0" or "1", "text after colon")
        sub_digit, body_text = hypothesis
        r1 = p.add_run("H")
        set_run_font(r1, bold=False, subscript=False)
        r2 = p.add_run(sub_digit)
        set_run_font(r2, bold=False, subscript=True)
        r3 = p.add_run(f": {body_text}")
        set_run_font(r3, bold=False, subscript=False)
    elif multiline_items is not None:
        # First line + <w:br/> + <w:br/> + items separated by <w:br/>
        intro_text, items = multiline_items
        r_intro = p.add_run(intro_text)
        set_run_font(r_intro, bold=False)
        r_br1 = p.add_run()
        set_run_font(r_br1, bold=False)
        r_br1._r.append(OxmlElement("w:br"))
        for item in items:
            r_item = p.add_run()
            set_run_font(r_item, bold=False)
            r_item._r.append(OxmlElement("w:br"))
            t = OxmlElement("w:t")
            t.text = item
            r_item._r.append(t)
    elif img_path is not None:
        r_img = p.add_run()
        set_run_font(r_img, bold=False)
        if exact_emu_width is not None:
            r_img.add_picture(img_path, width=Emu(exact_emu_width))
        else:
            with Image.open(img_path) as im:
                w_px, h_px = im.size
            # Rendered at 2x scale; convert to Word width (capped at 5731510 EMUs = ~6.27 in)
            target_emu = min(int(w_px * 4850), 5731510)
            r_img.add_picture(img_path, width=Emu(target_emu))
    else:
        if text:
            r = p.add_run(text)
            set_run_font(r, bold=bold)

    return p


def update_table_3_labels(table):
    # Match the exact labels added in class by the professor in Table 3 (Exercise 4)
    col_labels = [
        "Sistemas de bases de datos",
        "Rápido (Baja) Y1",
        "Lento (Baja)  Y2",
        "Rápido (Media)\t Y3",
        "Lento (Media) Y4",
        "Rápido (Alta) Y5",
        "Lento (Alta) Y6",
        "Total",
    ]
    for col_idx, label in enumerate(col_labels):
        cell = table.rows[1].cells[col_idx]
        if cell.paragraphs and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].text = label
            for extra_run in cell.paragraphs[0].runs[1:]:
                extra_run.text = ""

    row_labels = {2: "SQL (X1)", 3: "NoSQL (X2)", 4: "NewSQL (X3)"}
    for row_idx, label in row_labels.items():
        cell = table.rows[row_idx].cells[0]
        if cell.paragraphs and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].text = label
            for extra_run in cell.paragraphs[0].runs[1:]:
                extra_run.text = ""


SOLUTIONS = [
    {
        "num": 1,
        "question_substr": "¿Existe una asociación entre el sistema operativo y el tipo de dispositivo?",
        "h0": "No existe una asociación entre el sistema operativo y el tipo de dispositivo.",
        "h1": "Existe una asociación entre el sistema operativo y el tipo de dispositivo.",
        "ns": "5% (0.05)",
        "crosstab_img": "ex1_crosstab.png",
        "test_name": "Chi cuadrado de independencia",
        "test_img": "ex1_test.png",
        "contrast": "Dado que, el p valor (0.000) es menor que nivel de significancia (0.05), es probable que se rechace la hipótesis nula.",
        "stat_conc": "A un margen de error del 5% se concluye que, existe una asociación entre el sistema operativo y el tipo de dispositivo.",
        "admin_conc": "Dado que la elección del sistema operativo está asociada al tipo de dispositivo, se sugiere destinar Windows principalmente para equipos Desktop (50.0%) y Laptop (44.4%), y priorizar el sistema operativo Linux para los Servidores, donde presenta una clara predominancia (72.7%).",
    },
    {
        "num": 2,
        "question_substr": "¿Existe una asociación entre el lenguaje de programación y el éxito del proyecto?",
        "h0": "No existe una asociación entre el lenguaje de programación y el éxito del proyecto.",
        "h1": "Existe una asociación entre el lenguaje de programación y el éxito del proyecto.",
        "ns": "5% (0.05)",
        "crosstab_img": "ex2_crosstab.png",
        "test_name": "Chi cuadrado de independencia",
        "test_img": "ex2_test.png",
        "contrast": "Dado que, el p valor (0.008) es menor que nivel de significancia (0.05), es probable que se rechace la hipótesis nula.",
        "stat_conc": "A un margen de error del 5% se concluye que, existe una asociación entre el lenguaje de programación y el éxito del proyecto.",
        "admin_conc": "Es recomendable priorizar el uso de Python (66.7% de éxito) y JavaScript (63.6% de éxito) en el desarrollo de nuevos proyectos de software, así como reforzar la capacitación técnica y el seguimiento en los proyectos desarrollados en Java, los cuales registraron una mayor tasa de fracaso (57.1%).",
    },
    {
        "num": 3,
        "question_substr": "¿La condición de fallos son homogéneas entre los diferentes tipos de servidores?",
        "h0": "Las condiciones de fallos son homogéneas entre los diferentes tipos de servidores.",
        "h1": "Las condiciones de fallos no son homogéneas entre los diferentes tipos de servidores.",
        "ns": "5% (0.05)",
        "crosstab_img": "ex3_crosstab.png",
        "test_name": "Chi cuadrado de homogenidad",
        "test_img": "ex3_test.png",
        "contrast": "Dado que, el p valor (0.002) es menor que nivel de significancia (0.05), es probable que se rechace la hipótesis nula.",
        "stat_conc": "A un margen de error del 5% se concluye que, las condiciones de fallos no son homogéneas entre los diferentes tipos de servidores.",
        "admin_conc": "Se sugiere priorizar el uso de servidores MacOS y Linux para servicios críticos de hosting debido a que presentan menores porcentajes de fallas críticas (10.0% y 20.0% respectivamente), y establecer planes de mantenimiento preventivo en los servidores Windows, que registraron la mayor incidencia de fallas críticas (30.0%).",
    },
    {
        "num": 4,
        "question_substr": "¿Los tiempos de respuesta son homogéneos entre los diferentes sistemas de bases de datos",
        "h0": "Los tiempos de respuesta son homogéneos entre los diferentes sistemas de bases de datos bajo las distintas cargas de trabajo.",
        "h1": "Los tiempos de respuesta no son homogéneos entre los diferentes sistemas de bases de datos bajo las distintas cargas de trabajo.",
        "ns": "5% (0.05)",
        "crosstab_img": "ex4_crosstab.png",
        "crosstab_emu": 5731510,
        "test_name": "Chi cuadrado de homogenidad",
        "test_img": "ex4_test.png",
        "test_emu": 5731510,
        "contrast": "Dado que, el p valor (0.736) es mayor que nivel de significancia (0.05), es probable que no se rechace la hipótesis nula.",
        "stat_conc": "A un margen de error del 5% se concluye que, los tiempos de respuesta son homogéneos entre los diferentes sistemas de bases de datos bajo las distintas cargas de trabajo.",
        "admin_multiline": (
            "Dado que los 3 sistemas de base de datos, con respecto a los tiempos de respuesta son iguales, se podría sugerir adquirir cualquiera de estos sistemas. El factor diferenciador se podría estar considerando:",
            ["1.Costos", "2.Garantias", "3.Interaccion"],
        ),
    },
    {
        "num": 5,
        "question_substr": "¿Hubo una mejora significativa en el rendimiento del algoritmo de compresión de datos",
        "h0": "No hubo una mejora significativa en el rendimiento del algoritmo de compresión de datos después de la intervención.",
        "h1": "Hubo una mejora significativa en el rendimiento del algoritmo de compresión de datos después de la intervención.",
        "ns": "5% (0.05)",
        "crosstab_img": "ex5_crosstab.png",
        "test_name": "Chi cuadrado de Mc Nemar",
        "test_img": "ex5_test.png",
        "contrast": "Dado que, el p valor (0.041) es menor que nivel de significancia (0.05), es probable que se rechace la hipótesis nula.",
        "stat_conc": "A un margen de error del 5% se concluye que, hubo una mejora significativa en el rendimiento del algoritmo de compresión de datos después de la intervención.",
        "admin_conc": "Es favorable implementar la optimización del código en el algoritmo de compresión de datos, porque ha mostrado efectividad en los datos evaluados y en sus respectivas pruebas, logrando incrementar la tasa de éxito del 70.0% al 80.0% (reducción de errores del 10%).",
    },
    {
        "num": 6,
        "question_substr": "¿Hubo una mejora significativa en la fiabilidad del sistema de respaldo de datos",
        "h0": "No hubo una mejora significativa en la fiabilidad del sistema de respaldo de datos después de los ajustes en la configuración.",
        "h1": "Hubo una mejora significativa en la fiabilidad del sistema de respaldo de datos después de los ajustes en la configuración.",
        "ns": "5% (0.05)",
        "crosstab_img": "ex6_crosstab.png",
        "crosstab_emu": 5731510,
        "test_name": "Chi cuadrado de Mc Nemar",
        "test_img": "ex6_test.png",
        "test_emu": 5731510,
        "contrast": "Dado que, el p valor (0.013) es menor que nivel de significancia (0.05), es probable que se rechace la hipótesis nula.",
        "stat_conc": "A un margen de error del 5% se concluye que, hubo una mejora significativa en la fiabilidad del sistema de respaldo de datos después de los ajustes en la configuración.",
        "admin_conc": "Es favorable implementar el sistema de respaldo de datos, porque a mostrado efectividad en los datos mostrados y en sus respectivas pruebas, uno de los factores fue la reducción de porcentajes de errores (10%).",
    },
    {
        "num": 7,
        "question_substr": "¿Cuál es el índice de Kappa de Cohen para evaluar la confiabilidad del sistema de detección de intrusos",
        "h0": "No existe concordancia significativa entre los dos expertos al evaluar la confiabilidad del sistema de detección de intrusos.",
        "h1": "Existe concordancia significativa entre los dos expertos al evaluar la confiabilidad del sistema de detección de intrusos.",
        "ns": "5% (0.05)",
        "crosstab_img": "ex7_crosstab.png",
        "test_name": "Índice de Kappa de Cohen",
        "test_img": "ex7_test.png",
        "contrast": "Dado que, el p valor (0.005) es menor que nivel de significancia (0.05), es probable que se rechace la hipótesis nula.",
        "stat_conc": "A un margen de error del 5% se concluye que, existe concordancia significativa entre los dos expertos al evaluar la confiabilidad del sistema de detección de intrusos (Kappa = 0.276).",
        "admin_conc": "Si bien existe concordancia estadísticamente significativa entre ambos expertos (coincidiendo en el 79.0% de los casos evaluados), el valor del índice de Kappa (0.276) indica un nivel de acuerdo bajo o regular, por lo que se recomienda unificar los criterios de análisis de registros de red entre los especialistas.",
    },
    {
        "num": 8,
        "question_substr": "¿Cuál es el índice de Kappa de Cohen para evaluar la confiabilidad del sistema de clasificación de fallos",
        "h0": "No existe concordancia significativa entre los dos algoritmos al evaluar la confiabilidad del sistema de clasificación de fallos.",
        "h1": "Existe concordancia significativa entre los dos algoritmos al evaluar la confiabilidad del sistema de clasificación de fallos.",
        "ns": "5% (0.05)",
        "crosstab_img": "ex8_crosstab.png",
        "test_name": "Índice de Kappa de Cohen",
        "test_img": "ex8_test.png",
        "contrast": "Dado que, el p valor (0.000) es menor que nivel de significancia (0.05), es probable que se rechace la hipótesis nula.",
        "stat_conc": "A un margen de error del 5% se concluye que, existe concordancia significativa entre los dos algoritmos al evaluar la confiabilidad del sistema de clasificación de fallos (Kappa = 0.462).",
        "admin_conc": "Los algoritmos A y B presentan un nivel de concordancia moderado (Kappa = 0.462) y coinciden en el 86.0% de los eventos clasificados, por lo que es favorable emplear ambos algoritmos de detección de anomalías de forma conjunta para respaldar el diagnóstico de fallos.",
    },
    {
        "num": 9,
        "question_substr": "¿Hay una diferencia significativa en la eficacia de detección de intrusos entre el nuevo sistema y el antiguo?",
        "h0": "No hay una diferencia significativa en la eficacia de detección de intrusos entre el nuevo sistema y el antiguo.",
        "h1": "Hay una diferencia significativa en la eficacia de detección de intrusos entre el nuevo sistema y el antiguo.",
        "ns": "5% (0.05)",
        "crosstab_img": "ex9_crosstab.png",
        "test_name": "Prueba exacta de Fisher",
        "test_img": "ex9_test.png",
        "contrast": "Dado que, el p valor (0.584) es mayor que nivel de significancia (0.05), es probable que no se rechace la hipótesis nula.",
        "stat_conc": "A un margen de error del 5% se concluye que, no hay una diferencia significativa en la eficacia de detección de intrusos entre el nuevo sistema y el antiguo.",
        "admin_conc": "Dado que el nuevo sistema de detección de intrusos tiene una eficacia estadísticamente similar a la del sistema antiguo, no se justifica realizar una inversión adicional para el cambio de sistema salvo que existan mejoras en costos operativos, soporte o rendimiento.",
    },
    {
        "num": 10,
        "question_substr": "¿Hay una diferencia significativa en la seguridad de las contraseñas generadas entre el Sistema A y el Sistema B?",
        "h0": "No hay una diferencia significativa en la seguridad de las contraseñas generadas entre el Sistema A y el Sistema B.",
        "h1": "Hay una diferencia significativa en la seguridad de las contraseñas generadas entre el Sistema A y el Sistema B.",
        "ns": "5% (0.05)",
        "crosstab_img": "ex10_crosstab.png",
        "test_name": "Prueba exacta de Fisher",
        "test_img": "ex10_test.png",
        "contrast": "Dado que, el p valor (0.304) es mayor que nivel de significancia (0.05), es probable que no se rechace la hipótesis nula.",
        "stat_conc": "A un margen de error del 5% se concluye que, no hay una diferencia significativa en la seguridad de las contraseñas generadas entre el Sistema A y el Sistema B.",
        "admin_conc": "Dado que ambos sistemas de autenticación presentan un nivel de seguridad estadísticamente igual en la generación de contraseñas, se podría sugerir adquirir cualquiera de los dos sistemas considerando como factores diferenciadores el costo de implementación, la compatibilidad y el soporte técnico.",
    },
]


def main():
    doc = docx.Document(BACKUP_DOCX)

    # Update Table 3 with Y1..Y6 and X1..X3 labels as done in class
    update_table_3_labels(doc.tables[3])

    for sol in SOLUTIONS:
        target_p = None
        for p in doc.paragraphs:
            if sol["question_substr"] in p.text:
                target_p = p
                break
        if target_p is None:
            raise RuntimeError(f"Could not find paragraph for Exercise {sol['num']}")

        cur = target_p
        # PASO 1
        cur = insert_paragraph_after(cur, "PASO 1: Planteamiento de hipótesis", bold=True)
        cur = insert_paragraph_after(cur, hypothesis=("0", sol["h0"]))
        cur = insert_paragraph_after(cur, hypothesis=("1", sol["h1"]))

        # PASO 2
        cur = insert_paragraph_after(cur, "PASO 2: Determinar el nivel de significancia", bold=True)
        cur = insert_paragraph_after(cur, sol["ns"], bold=False)

        # PASO 3
        cur = insert_paragraph_after(cur, "PASO 3: Presentación de tablas de contingencia", bold=True)
        cur = insert_paragraph_after(
            cur,
            img_path=os.path.join(IMG_DIR, sol["crosstab_img"]),
            exact_emu_width=sol.get("crosstab_emu"),
        )

        # PASO 4
        cur = insert_paragraph_after(cur, "PASO 4: Determinar la prueba estadística", bold=True)
        cur = insert_paragraph_after(cur, sol["test_name"], bold=False)

        # PASO 5
        cur = insert_paragraph_after(cur, "PASO 5: Contrastación de hipótesis", bold=True)
        cur = insert_paragraph_after(
            cur,
            img_path=os.path.join(IMG_DIR, sol["test_img"]),
            exact_emu_width=sol.get("test_emu"),
        )
        cur = insert_paragraph_after(cur, sol["contrast"], bold=False)

        # PASO 6
        cur = insert_paragraph_after(cur, "PASO 6: Conclusión estadística", bold=True)
        cur = insert_paragraph_after(cur, sol["stat_conc"], bold=False)

        # PASO 7
        cur = insert_paragraph_after(cur, "PASO 7: Conclusión administrativa", bold=True)
        if "admin_multiline" in sol:
            cur = insert_paragraph_after(cur, multiline_items=sol["admin_multiline"])
        else:
            cur = insert_paragraph_after(cur, sol["admin_conc"], bold=False)

    # Remove trailing empty paragraphs at the end of the document to prevent a blank last page
    while doc.paragraphs and not doc.paragraphs[-1].text.strip() and not doc.paragraphs[-1]._p.xpath(".//w:drawing"):
        p_el = doc.paragraphs[-1]._p
        p_el.getparent().remove(p_el)

    doc.save(RESUELTO_DOCX)
    doc.save(SRC_DOCX)
    print(f"Saved: {RESUELTO_DOCX}")
    print(f"Saved: {SRC_DOCX}")

    # Export to PDF via Word COM
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        wdoc = word.Documents.Open(os.path.abspath(RESUELTO_DOCX))
        wdoc.SaveAs(os.path.abspath(RESUELTO_PDF), FileFormat=17)
        wdoc.Close(False)
        print(f"Exported PDF: {RESUELTO_PDF}")
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
