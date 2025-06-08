from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import os
import csv
import re

# Configuraciones
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
pdf_path = r'C:\Users\manue\OneDrive\manu\resumenes\resumen.pdf'
poppler_bin_path = r'C:\poppler\poppler-24.08.0\Library\bin'

# Convertir PDF a imágenes
pages = convert_from_path(pdf_path, dpi=300, poppler_path=poppler_bin_path)

# OCR y extracción de líneas
lines = []
for i, page in enumerate(pages):
    text = pytesseract.image_to_string(page)
    lines += text.split('\n')

# Limpiar líneas vacías
lines = [line.strip() for line in lines if line.strip()]

# Extraer movimientos en formato CSV
output_csv = "movimientos_extraidos.csv"
with open(output_csv, mode="w", newline="", encoding="utf-8") as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow(["Fecha", "Descripción", "Valor", "Saldo"])

    current_fecha = ""
    current_descripcion = ""

    for i in range(len(lines)):
        line = lines[i]

        # Detectar fecha
        if re.match(r"^\d{4}/\d{2}/\d{2}$", line):
            current_fecha = line
            current_descripcion = lines[i+1] if i+1 < len(lines) else ""
        
        # Detectar línea con valor y saldo
        elif re.search(r"\$\d", line) and current_fecha and current_descripcion:
            valores = re.findall(r"\$[\d,\.]+", line)
            if len(valores) == 2:
                writer.writerow([current_fecha, current_descripcion, valores[0], valores[1]])
                current_fecha = ""
                current_descripcion = ""

print(f"\n✅ Listo. CSV generado: {output_csv}")
