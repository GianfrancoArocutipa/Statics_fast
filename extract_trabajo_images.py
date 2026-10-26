import docx
import os
import zipfile
from xml.etree import ElementTree as ET

# Extract media from trabajo.docx
doc_path = 'trabajo.docx'
extract_dir = 'extracted_trabajo_media'
os.makedirs(extract_dir, exist_ok=True)

with zipfile.ZipFile(doc_path, 'r') as z:
    for filename in z.namelist():
        if filename.startswith('word/media/'):
            target_name = os.path.basename(filename)
            target_path = os.path.join(extract_dir, target_name)
            with open(target_path, 'wb') as f:
                f.write(z.read(filename))
            print(f'Extracted: {target_name} ({len(z.read(filename))} bytes)')

# Map images to paragraphs in trabajo.docx
doc = docx.Document(doc_path)
for p_idx, p in enumerate(doc.paragraphs):
    print(f"\n--- Paragraph {p_idx}: '{p.text}' ---")
    xml_str = p._p.xml
    # find blip r:embed
    import re
    embeds = re.findall(r'r:embed="(rId\d+)"', xml_str)
    for rId in embeds:
        target = p.part.rels[rId].target_ref
        print(f"  Image target: {target} (rId: {rId})")
