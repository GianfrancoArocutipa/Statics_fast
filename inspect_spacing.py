import docx

doc = docx.Document("Ejercicios Pruebas de chi cuadrado.docx")
for i in [1, 3, 4, 8, 9, 14, 15, 16, 18, 20, 21, 22, 23]:
    p = doc.paragraphs[i]
    pf = p.paragraph_format
    sb = pf.space_before.pt if pf.space_before else None
    sa = pf.space_after.pt if pf.space_after else None
    ls = pf.line_spacing
    print(f"P{i} ['{p.text[:30]}']: align={p.alignment}, style={p.style.name}, sb={sb}, sa={sa}, ls={ls}")
