import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = docx.Document("Ejercicios Pruebas de chi cuadrado.docx")

# Let's see: if we find P5, we can insert before it
target_p = None
for p in doc.paragraphs:
    if "En una empresa de desarrollo de software" in p.text:
        target_p = p
        break

if target_p:
    new_p = target_p.insert_paragraph_before("TEST INSERTION BEFORE P5", style="List Paragraph")
    new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    new_p.paragraph_format.space_before = Pt(6)
    new_p.paragraph_format.space_after = Pt(6)
    new_p.paragraph_format.line_spacing = 1.15
    for r in new_p.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(11)
    doc.save("test_insert.docx")
    print("Test insert succeeded!")
else:
    print("Target paragraph not found")
