from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import os

# Ruta al ejecutable de Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Ruta al PDF
pdf_path = r'C:\Users\manue\OneDrive\manu\resumenes\resumen.pdf'

# Convertir PDF en imágenes
pages = convert_from_path(pdf_path, dpi=300)

# Crear carpeta temporal para imágenes
os.makedirs("paginas_temp", exist_ok=True)

# Crear archivo de salida
output_path = "texto_extraido.txt"
with open(output_path, "w", encoding="utf-8") as f_out:
    for i, page in enumerate(pages):
        img_path = f"paginas_temp/pagina_{i+1}.png"
        page.save(img_path, "PNG")

        text = pytesseract.image_to_string(Image.open(img_path), lang='spa')

        f_out.write(f"\n--- TEXTO EXTRAÍDO - Página {i+1} ---\n")
        f_out.write(text + "\n")

        print(f"Página {i+1} procesada.")

print(f"\n✅ Extracción finalizada. Texto guardado en: {output_path}")
