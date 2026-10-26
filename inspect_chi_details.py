import docx
from PIL import Image
import os
import re

doc = docx.Document("Ejercicios Pruebas de chi cuadrado.docx")
for i, p in enumerate(doc.paragraphs):
    xml = p._p.xml
    embeds = re.findall(r'r:embed="([^"]+)"', xml)
    for rId in embeds:
        target = p.part.rels[rId].target_ref
        fname = os.path.basename(target)
        img_path = os.path.join("extracted_chi_media", fname)
        im = Image.open(img_path)
        print(f"P{i} embeds {fname} ({im.size[0]}x{im.size[1]})")
