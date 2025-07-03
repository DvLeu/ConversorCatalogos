from converter import leer_excel, convertir_a_json, limpiar_dataframe

excel_path = "catalogos/catCFDI_V_4_20250618.xlsx"
df = leer_excel(excel_path)
df_clean = limpiar_dataframe(df)
json_str = convertir_a_json(df_clean, archivo_salida="catalogos/tablaPrueba.json", bonito=True)

print("¡Conversión terminada!")