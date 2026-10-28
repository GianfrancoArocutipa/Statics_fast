import os
import win32com.client

docx_path = os.path.abspath(os.path.join("prueba", "Segunda_Practica_Calificada_Inferencia_Estadistica_2026-II_Resuelto.docx"))
pdf_path = os.path.abspath(os.path.join("prueba", "Segunda_Practica_Calificada_Inferencia_Estadistica_2026-II_Resuelto.pdf"))

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False

try:
    doc = word.Documents.Open(docx_path)
    # wdFormatPDF = 17
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    print(f"PDF generado exitosamente en: {pdf_path}")
except Exception as e:
    print(f"Error al generar PDF: {e}")
finally:
    word.Quit()
