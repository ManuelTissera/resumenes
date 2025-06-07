import pandas as pd
import os

# Directorios
crudos_dir = 'Crudos'
salida_path = 'movimientos_totales_2025.csv'

# Obtener todos los CSV que empiecen con "movimientos_" y terminen en ".csv"
archivos = [f for f in os.listdir(crudos_dir) if f.startswith("movimientos_") and f.endswith(".csv")]

# Leer y acumular los archivos
dataframes = []
for archivo in archivos:
    ruta = os.path.join(crudos_dir, archivo)
    df = pd.read_csv(ruta)
    dataframes.append(df)

# Unir todos en uno solo
df_total = pd.concat(dataframes, ignore_index=True)

# Guardar en la raíz de /Resumenes
df_total.to_csv(salida_path, index=False)

print(f"Archivo combinado guardado como: {salida_path}")
