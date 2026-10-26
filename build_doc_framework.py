# -*- coding: utf-8 -*-
"""
Script to generate the comprehensive, professional Word report for:
"APLICACIÓN DE ESTADÍSTICA INFERENCIAL CON DATOS REALES"
Universidad Privada de Tacna - Facultad de Ingeniería - Escuela de Ingeniería de Sistemas
Docente: Luis Alfredo Fernandez Vizcarra
Estudiantes:
- Arocutipa Arocutipa, Gian Franco (2023076790)
- Perez Peralta, Fabrizio Salvador Elias (2023077476)
- LLica Mamani, Jimmy Mijair (2023076789)
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_document():
    doc = docx.Document()
    
    # Configure page setup (A4 / Letter with 2.54cm = 1 inch margins)
    section = doc.sections[0]
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    
    # Configure default style
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(6)
    style_normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Helper functions
    def add_p(text="", bold=False, italic=False, size=12, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=6, color=RGBColor(0x22, 0x22, 0x22)):
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
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D) # Navy institutional
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x2C, 0x3E, 0x50)
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
        run.font.color.rgb = RGBColor(0x34, 0x49, 0x5E)
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
            run_cap.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        else:
            print(f"Warning: Image {img_path} not found.")

    def add_placeholder(title, description, detail_text=""):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        cell.width = Inches(6.5)
        
        # Border & shading
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
        r1 = p.add_run(f"📷 [ESPACIO RESERVADO PARA IMAGEN COMPLEMENTARIA]\n")
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(0x29, 0x80, 0xB9)
        
        r2 = p.add_run(f"{title}\n")
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(11)
        r2.font.bold = True
        r2.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
        
        r3 = p.add_run(f"{description}")
        r3.font.name = 'Times New Roman'
        r3.font.size = Pt(10)
        r3.font.italic = True
        r3.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        
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

    def style_table(table, header_bg="1B365D", header_fg="FFFFFF", alt_bg="F8F9F9"):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_idx, row in enumerate(table.rows):
            # Cant split row across pages
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(r'<w:cantSplit {}/>'.format(nsdecls('w'))))
            if r_idx == 0:
                trPr.append(parse_xml(r'<w:tblHeader {}/>'.format(nsdecls('w'))))
                
            for c_idx, cell in enumerate(row.cells):
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                tcPr = cell._tc.get_or_add_tcPr()
                if r_idx == 0:
                    shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), header_bg))
                    tcPr.append(shd)
                    for cp in cell.paragraphs:
                        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in cp.runs:
                            run.font.name = 'Times New Roman'
                            run.font.bold = True
                            run.font.size = Pt(10.5)
                            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                else:
                    if r_idx % 2 == 1:
                        shd = parse_xml(r'<w:shd {} w:fill="FFFFFF"/>'.format(nsdecls('w')))
                    else:
                        shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), alt_bg))
                    tcPr.append(shd)
                    for cp in cell.paragraphs:
                        for run in cp.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(10)
                            run.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
                
                # borders: APA style subtle borders
                borders = parse_xml(
                    r'<w:tcBorders {} >'
                    r'  <w:top w:val="single" w:sz="4" w:space="0" w:color="BDC3C7"/>'
                    r'  <w:left w:val="none"/>'
                    r'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="BDC3C7"/>'
                    r'  <w:right w:val="none"/>'
                    r'</w:tcBorders>'.format(nsdecls('w'))
                )
                tcPr.append(borders)

    print("Framework initialized. Building sections...")
    return doc, add_p, add_h1, add_h2, add_h3, add_img, add_placeholder, style_table

if __name__ == '__main__':
    doc, add_p, add_h1, add_h2, add_h3, add_img, add_placeholder, style_table = create_document()
    print("Document structure ready.")
