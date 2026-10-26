import docx
import re

doc = docx.Document('Ejercicios Pruebas de chi cuadrado.docx')
for i, p in enumerate(doc.paragraphs):
    xml = p._p.xml
    cx_matches = re.findall(r'cx="(\d+)"\s+cy="(\d+)"', xml)
    embeds = re.findall(r'r:embed="([^"]+)"', xml)
    if cx_matches:
        for (cx, cy), rId in zip(cx_matches, embeds):
            cx_in = int(cx) / 914400.0
            cy_in = int(cy) / 914400.0
            target = p.part.rels[rId].target_ref
            print(f'P{i}: file={target}, width={cx_in:.2f} in ({cx_in*2.54:.2f} cm), height={cy_in:.2f} in ({cy_in*2.54:.2f} cm)')
