from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import os
import csv
import re


pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
pdf_path = r'C:\Users\manue\OneDrive\manu\resumenes\resumen.pdf'
poppler_bin_path = r'C:\poppler\poppler-24.08.0\Library\bin'


pages = convert_from_path(pdf_path, dpi=300, poppler_path=poppler_bin_path)


lines = []
for page in pages:
    text = pytesseract.image_to_string(page)
    lines += text.split('\n')


lines = [line.strip() for line in lines if line.strip()]


fechas = [line for line in lines if re.match(r"^\d{4}/\d{2}/\d{2}$", line)]
descripciones = []
valores_y_saldos = []

descripcion_flag = False
for line in lines:
    if "DESCRIPCION TRANSACCION" in line.upper():
        descripcion_flag = True
        continue
    if re.match(r"^\d{4}/\d{2}/\d{2}$", line):  # terminó la parte de descripciones
        descripcion_flag = False
    if descripcion_flag:
        descripciones.append(line)
    if re.search(r"\$\d", line) and len(re.findall(r"\$", line)) == 2:
        valores_y_saldos.append(line)


output_csv = "movimientos_extraidos.csv"
with open(output_csv, mode="w", newline="", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(["Fecha", "Descripción", "Valor", "Saldo"])

    for i in range(min(len(fechas), len(descripciones), len(valores_y_saldos))):
        fecha = fechas[i]
        descripcion = descripciones[i]
        valores = re.findall(r"\$[\d,\.]+", valores_y_saldos[i])
        if len(valores) == 2:
            writer.writerow([fecha, descripcion, valores[0], valores[1]])

print(f"\n✅ Listo. Se generó el archivo: {output_csv}")
