from converter import read_excel_file, clean_dataframe, convert_to_json

excel_path = "catalogos/catCFDI_V_4_20250618.xlsx"
# Leer el archivo Excel
df = read_excel_file(excel_path)

# Limpiar el DataFrame
df_clean = clean_dataframe(df)

# Convertir a JSON (como string)
json_str = convert_to_json(df_clean, output_file="catalogos/salida.json", pretty=True)

print("¡Conversión terminada!")