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

DOC_1_2_PATH = os.path.join(BASE_DIR, "Ejercicios Chi cuadrado 2da. parte (1) (2).docx")
DOC_1_2_BACKUP = os.path.join(BASE_DIR, "Ejercicios Chi cuadrado 2da. parte (1) (2)_Original.docx")
DOC_1_2_PDF = os.path.join(BASE_DIR, "Ejercicios Chi cuadrado 2da. parte (1) (2).pdf")

IMG_DIR = os.path.join(BASE_DIR, "generated_chi2_part2_images")

if not os.path.exists(BACKUP_DOCX) and os.path.exists(SRC_DOCX):
    shutil.copy2(SRC_DOCX, BACKUP_DOCX)

if not os.path.exists(DOC_1_2_BACKUP) and os.path.exists(DOC_1_2_PATH):
    shutil.copy2(DOC_1_2_PATH, DOC_1_2_BACKUP)


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
        sub_digit, body_text = hypothesis
        r1 = p.add_run("H")
        set_run_font(r1, bold=False, subscript=False)
        r2 = p.add_run(sub_digit)
        set_run_font(r2, bold=False, subscript=True)
        r3 = p.add_run(f": {body_text}")
        set_run_font(r3, bold=False, subscript=False)
    elif multiline_items is not None:
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
            target_emu = min(int(w_px * 4850), 5731510)
            r_img.add_picture(img_path, width=Emu(target_emu))
    else:
        if text:
            r = p.add_run(text)
            set_run_font(r, bold=bold)

    return p


def update_table_3_labels(table):
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
        "admin_conc": "Es favorable implementar el sistema operativo de acuerdo al tipo de dispositivo (Windows en equipos Desktop y Laptop, y Linux en Servidores), porque a mostrado efectividad en los datos mostrados y en sus respectivas pruebas, uno de los factores fue el mayor porcentaje de predominancia en Servidores con Linux (72.7%) y Desktop con Windows (50.0%).",
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
        "admin_conc": "Es favorable implementar los lenguajes de programación Python y JavaScript en los proyectos de software, porque a mostrado efectividad en los datos mostrados y en sus respectivas pruebas, uno de los factores fue el mayor porcentaje de éxito alcanzado (66.7% y 63.6% respectivamente) y la reducción de porcentajes de fracaso frente a Java (57.1%).",
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
        "admin_conc": "Es favorable implementar los servidores MacOS y Linux en la compañía de hosting, porque a mostrado efectividad en los datos mostrados y en sus respectivas pruebas, uno de los factores fue la reducción de porcentajes de fallas críticas (10.0% y 20.0% respectivamente) frente a Windows (30.0%).",
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
        "admin_conc": "Es favorable implementar el algoritmo de compresión de datos optimizado, porque a mostrado efectividad en los datos mostrados y en sus respectivas pruebas, uno de los factores fue la reducción de porcentajes de errores (10%).",
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
        "admin_conc": "Es favorable implementar el sistema de detección de intrusos evaluado por los dos expertos, porque a mostrado efectividad en los datos mostrados y en sus respectivas pruebas, uno de los factores fue el porcentaje de coincidencia en la detección de intrusiones (79%).",
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
        "admin_conc": "Es favorable implementar el sistema de clasificación de fallos mediante los dos algoritmos de detección de anomalías, porque a mostrado efectividad en los datos mostrados y en sus respectivas pruebas, uno de los factores fue el porcentaje de coincidencia en la identificación de fallos (86%).",
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
        "admin_multiline": (
            "Dado que los 2 sistemas de detección de intrusos, con respecto a la eficacia de detección son iguales, se podría sugerir adquirir cualquiera de estos sistemas. El factor diferenciador se podría estar considerando:",
            ["1.Costos", "2.Garantias", "3.Interaccion"],
        ),
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
        "admin_multiline": (
            "Dado que los 2 sistemas de autenticación, con respecto a la seguridad de las contraseñas generadas son iguales, se podría sugerir adquirir cualquiera de estos sistemas. El factor diferenciador se podría estar considerando:",
            ["1.Costos", "2.Garantias", "3.Interaccion"],
        ),
    },
]


def insert_solution_steps(target_p, sol):
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


def build_full_10_doc():
    doc = docx.Document(BACKUP_DOCX)
    update_table_3_labels(doc.tables[3])

    for sol in SOLUTIONS:
        target_p = None
        for p in doc.paragraphs:
            if sol["question_substr"] in p.text:
                target_p = p
                break
        if target_p is None:
            raise RuntimeError(f"Could not find paragraph for Exercise {sol['num']}")
        insert_solution_steps(target_p, sol)

    while doc.paragraphs and not doc.paragraphs[-1].text.strip() and not doc.paragraphs[-1]._p.xpath(".//w:drawing"):
        p_el = doc.paragraphs[-1]._p
        p_el.getparent().remove(p_el)

    doc.save(RESUELTO_DOCX)
    doc.save(SRC_DOCX)
    print(f"Saved: {RESUELTO_DOCX}")
    print(f"Saved: {SRC_DOCX}")


def build_doc_1_2():
    if not os.path.exists(DOC_1_2_BACKUP):
        return
    doc = docx.Document(DOC_1_2_BACKUP)
    # In DOC_1_2, only problems corresponding to SOLUTIONS[0] (num=1) and SOLUTIONS[6] (num=7) are unsolved
    for sol in [SOLUTIONS[0], SOLUTIONS[6]]:
        target_p = None
        for p in doc.paragraphs:
            if sol["question_substr"] in p.text:
                target_p = p
                break
        if target_p is not None:
            insert_solution_steps(target_p, sol)

    while doc.paragraphs and not doc.paragraphs[-1].text.strip() and not doc.paragraphs[-1]._p.xpath(".//w:drawing"):
        p_el = doc.paragraphs[-1]._p
        p_el.getparent().remove(p_el)

    doc.save(DOC_1_2_PATH)
    print(f"Saved: {DOC_1_2_PATH}")


def main():
    build_full_10_doc()
    build_doc_1_2()

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        wdoc = word.Documents.Open(os.path.abspath(RESUELTO_DOCX))
        wdoc.SaveAs(os.path.abspath(RESUELTO_PDF), FileFormat=17)
        wdoc.Close(False)
        print(f"Exported PDF: {RESUELTO_PDF}")

        if os.path.exists(DOC_1_2_PATH):
            wdoc2 = word.Documents.Open(os.path.abspath(DOC_1_2_PATH))
            wdoc2.SaveAs(os.path.abspath(DOC_1_2_PDF), FileFormat=17)
            wdoc2.Close(False)
            print(f"Exported PDF: {DOC_1_2_PDF}")
    finally:
        word.Quit()


if __name__ == "__main__":
    main()
