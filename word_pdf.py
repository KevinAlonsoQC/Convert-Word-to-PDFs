import os
import win32com.client

wdFormatPDF = 17

# Ruta de la carpeta con archivos .docx
# Ejemplo
carpeta_docx = r"C:\Users\usuario\Desktop\Documentos\PRUEBAS"

# Ruta de la carpeta de salida para archivos .pdf
# Ejemplo
carpeta_pdf = r"C:\Users\usuario\Desktop\Documentos\PRUEBAS\PDFs"

# Asegúrate de que la carpeta de salida exista, si no, créala
if not os.path.exists(carpeta_pdf):
    os.makedirs(carpeta_pdf)

# Recorre la carpeta con archivos .docx
for archivo_docx in os.listdir(carpeta_docx):
    if archivo_docx.endswith(".docx"):
        # Ruta completa de entrada y salida para cada archivo
        ruta_docx = os.path.join(carpeta_docx, archivo_docx)
        nombre_pdf = os.path.splitext(archivo_docx)[0] + ".pdf"
        ruta_pdf = os.path.join(carpeta_pdf, nombre_pdf)

        # Convierte el archivo .docx a .pdf
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(ruta_docx)
        doc.SaveAs(ruta_pdf, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()

print("Conversión completada.")
