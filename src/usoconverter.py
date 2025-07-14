from converter import leer_excel, convertir_a_json, limpiar_dataframe, seleccionar_columnas_ampersand
import os
import pandas as pd 

os.makedirs("catalogos/json_output", exist_ok=True)
excel_path = "catalogos/catCFDI_V_4_20250618.xlsx"
all_sheets = pd.ExcelFile(excel_path).sheet_names
hojas = [hoja for hoja in all_sheets if "&" in hoja]

print("PROCESAMIENTO MASIVO DE CATALOGOS SAT")
print()

resultados = {}

for hoja in hojas:
    try:
        print(f"Procesando: {hoja}")
        
        df = leer_excel(excel_path, hoja)
        df = seleccionar_columnas_ampersand(df)
        df_clean = limpiar_dataframe(df, limpiar_columnas=True, metodo_acentos='espanol')
        hoja_salida = hoja.replace("&", "")  
        salida = f"catalogos/json_output/{hoja_salida}_limpio.json" 
        convertir_a_json(df_clean, archivo_salida=salida, bonito=True)
        
        resultados[hoja_salida] = {  
            'status': 'exitoso',
            'registros': len(df_clean),
            'columnas': len(df_clean.columns),
            'archivo': salida
        }
        
        print(f"{hoja_salida}: {len(df_clean)} registros -> {salida}")
        print(f"Columnas: {list(df_clean.columns)[:3]}{'...' if len(df_clean.columns) > 3 else ''}")
        print()
        
    except Exception as e:
        print(f"Error procesando {hoja}: {e}")
        print()
        hoja_salida = hoja.replace("&", "")
        resultados[hoja_salida] = {
            'status': 'error',
            'error': str(e)
        }

print("REPORTE FINAL")
total_registros = 0
exitosos = 0

for hoja_salida, resultado in resultados.items():
    if resultado['status'] == 'exitoso':
        print(f"{hoja_salida}: {resultado['registros']} registros")
        total_registros += resultado['registros']
        exitosos += 1
    else:
        print(f"{hoja_salida}: {resultado['error']}")

print()
print(f"Catalogos procesados: {exitosos}/{len(hojas)}")
print(f"Total de registros: {total_registros}")
print(f"Archivos generados en: catalogos/json_output/")
print()
print("Procesamiento completado")