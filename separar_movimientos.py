import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.styles.numbers import FORMAT_CURRENCY_USD_SIMPLE

# Parámetro del mes
mes = 'mayo'

# Verificar que la carpeta base "Individuales" existe
base_dir = 'Individuales'
if not os.path.isdir(base_dir):
    raise FileNotFoundError(f'La carpeta base "{base_dir}" no existe. Por favor, créala manualmente antes de continuar.')

# Verificar si ya existe la carpeta del mes
output_dir = os.path.join(base_dir, mes)
if os.path.exists(output_dir):
    print(f'La carpeta "{output_dir}" ya existe. El trabajo de "{mes}" ya fue realizado.')
    exit()

# Crear carpeta del mes
os.makedirs(output_dir)

# Rutas
input_path = 'Crudos/dataset_movimientos_mayo2025_limpio.csv'

# Leer el archivo CSV
df = pd.read_csv(input_path)
nombres_unicos = df['Nombre'].unique()

# Exportar y aplicar formato
for nombre in nombres_unicos:
    df_persona = df[df['Nombre'] == nombre]
    nombre_archivo = f"{nombre.replace(' ', '_')}_{mes}.xlsx"
    ruta_archivo = os.path.join(output_dir, nombre_archivo)

    # Guardar como Excel
    df_persona.to_excel(ruta_archivo, index=False)

    # Abrir el archivo para aplicar formato
    wb = load_workbook(ruta_archivo)
    ws = wb.active

    # Formato encabezado
    header_fill = PatternFill(start_color="177086", end_color="177086", fill_type="solid")
    header_border = Border(
        left=Side(style="thin", color="177086"),
        right=Side(style="thin", color="177086"),
        top=Side(style="thin", color="177086"),
        bottom=Side(style="thin", color="177086")
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")
        cell.border = header_border

    # Formato moneda columna "Pesos" (columna E)
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
        for cell in row:
            cell.number_format = FORMAT_CURRENCY_USD_SIMPLE

    # Formato moneda columna "Dólares" (columna F)
    for row in ws.iter_rows(min_row=2, min_col=6, max_col=6):
        for cell in row:
            cell.number_format = '"USD" #,##0.00'

    # Agregar totales al final
    ultima_fila = ws.max_row + 1
    ws[f'E{ultima_fila}'] = f"=SUM(E2:E{ultima_fila - 1})"
    ws[f'F{ultima_fila}'] = f"=SUM(F2:F{ultima_fila - 1})"

    # Formato totales
    for col in ['E', 'F']:
        cell = ws[f'{col}{ultima_fila}']
        cell.font = Font(bold=True, color="444444")
        cell.fill = PatternFill(start_color="e3e8ea", end_color="e3e8ea", fill_type="solid")
        cell.number_format = FORMAT_CURRENCY_USD_SIMPLE
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(top=(Side(style="medium", color="222222")))

    # Guardar archivo final
    wb.save(ruta_archivo)

print("Archivos Excel individuales generados con formato, bordes y totales.")
