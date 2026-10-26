# -*- coding: utf-8 -*-
import docx

def update_document(filename):
    print(f"Updating {filename}...")
    doc = docx.Document(filename)
    updated = False
    
    for i, p in enumerate(doc.paragraphs):
        if "[COMPLETAR: N.º TOTAL DE VARIABLES]" in p.text or "COMPLETAR" in p.text:
            print(f"Found in paragraph {i}")
            for r in p.runs:
                if "[COMPLETAR: N.º TOTAL DE VARIABLES]" in r.text:
                    r.text = r.text.replace("[COMPLETAR: N.º TOTAL DE VARIABLES]", "65")
                    r.font.highlight_color = None
                    updated = True
                elif "COMPLETAR" in r.text:
                    r.text = r.text.replace("COMPLETAR", "65")
                    r.font.highlight_color = None
                    updated = True
                else:
                    if r.font.highlight_color is not None:
                        r.font.highlight_color = None
                        
            # Also ensure if whole text had it
            if "[COMPLETAR: N.º TOTAL DE VARIABLES]" in p.text:
                # In case it spans multiple runs
                p.text = p.text.replace("[COMPLETAR: N.º TOTAL DE VARIABLES]", "65")
                updated = True

    if updated:
        doc.save(filename)
        print(f"Successfully saved updated {filename}")
    else:
        print(f"Marker not found in {filename}")

if __name__ == '__main__':
    files_to_update = [
        'PEREZ_AROCUTIPA_LLICA_APLICACION_ESTADISTICA_INFERENCIAL_COMPLEMENTARIO_CORREGIDO (1).docx',
        'PEREZ_AROCUTIPA_LLICA_APLICACION_ESTADISTICA_INFERENCIAL_COMPLEMENTARIO_CORREGIDO.docx',
        'PEREZ_AROCUTIPA_LLICA_APLICACION_ESTADISTICA_INFERENCIAL.docx'
    ]
    for f in files_to_update:
        try:
            update_document(f)
        except Exception as e:
            print(f"Error updating {f}: {e}")
