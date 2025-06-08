import pandas as pd
import xlsxwriter

# --- Configuración ---
valor_dolar = 1300  # Valor del dólar en pesos

# --- Cargar archivo ---
file_path = "movimientos_totales_2025.csv"
df = pd.read_csv(file_path)

# --- Parsear fechas ---
df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")
df["Mes"] = df["Fecha"].dt.to_period("M").astype(str)

# --- Gasto total por persona y por mes ---
gasto_persona_mes = df.groupby(["Nombre", "Mes"])[["Pesos", "Dolares"]].sum().reset_index()

# --- Gasto total por tipo por mes ---
gasto_tipo_mes = df.groupby(["Descripcion", "Mes"])[["Pesos", "Dolares"]].sum().reset_index()

# --- Ranking de mayores gastadores por mes (conversión de dólares a pesos) ---
ranking_por_mes = df.groupby(["Nombre", "Mes"])[["Pesos", "Dolares"]].sum().reset_index()
ranking_por_mes["Dolares_en_Pesos"] = ranking_por_mes["Dolares"] * valor_dolar
ranking_por_mes["Total_Convertido"] = ranking_por_mes["Pesos"] + ranking_por_mes["Dolares_en_Pesos"]
ranking_por_mes = ranking_por_mes.sort_values(by=["Mes", "Total_Convertido"], ascending=[True, False])

# --- Títulos de las tablas ---
titulo_persona_mes = "Gasto total por persona y mes"
titulo_tipo_mes = "Gasto total por tipo (Familia o Colaborador) por mes"
titulo_ranking = "Ranking de mayores gastadores por mes (con dólares convertidos a pesos)"

# --- Crear archivo Excel ---
output_path = "reporte_2025.xlsx"
workbook = xlsxwriter.Workbook(output_path)
worksheet = workbook.add_worksheet("Estadísticas")

# --- Función para escribir tabla con título ---
def escribir_tabla(titulo, dataframe, row_start):
    worksheet.write(row_start, 0, titulo)
    for col_num, value in enumerate(dataframe.columns):
        worksheet.write(row_start + 1, col_num, value)
    for i, row_data in enumerate(dataframe.values):
        for j, value in enumerate(row_data):
            worksheet.write(row_start + 2 + i, j, value)
    return row_start + 2 + len(dataframe)

# --- Escribir tablas una debajo de la otra ---
row = 0
row = escribir_tabla(titulo_persona_mes, gasto_persona_mes, row)
row += 2
row = escribir_tabla(titulo_tipo_mes, gasto_tipo_mes, row)
row += 2
row = escribir_tabla(titulo_ranking, ranking_por_mes, row)

# --- Cerrar Excel ---
workbook.close()

print("✅ Archivo Excel generado correctamente:", output_path)
