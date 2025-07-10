from converter import leer_excel, convertir_a_json, limpiar_dataframe
import os

os.makedirs("catalogos/json_output", exist_ok=True)

hojas = ["c_FormaPago", "c_Moneda", "c_TipoDeComprobante"]

print("PROCESAMIENTO MASIVO DE CATALOGOS SAT")
print()

resultados = {}

for hoja in hojas:
    try:
        print(f"Procesando: {hoja}")
        
        df = leer_excel("catalogos/catCFDI_V_4_20250618.xlsx", hoja)
        df_clean = limpiar_dataframe(df, limpiar_columnas=True, metodo_acentos='espanol')
        salida = f"catalogos/json_output/{hoja}_limpio.json"
        convertir_a_json(df_clean, archivo_salida=salida, bonito=True)
        
        resultados[hoja] = {
            'status': 'exitoso',
            'registros': len(df_clean),
            'columnas': len(df_clean.columns),
            'archivo': salida
        }
        
        print(f"{hoja}: {len(df_clean)} registros -> {salida}")
        print(f"Columnas: {list(df_clean.columns)[:3]}{'...' if len(df_clean.columns) > 3 else ''}")
        print()
        
    except Exception as e:
        print(f"Error procesando {hoja}: {e}")
        print()
        resultados[hoja] = {
            'status': 'error',
            'error': str(e)
        }

print("REPORTE FINAL")
total_registros = 0
exitosos = 0

for hoja, resultado in resultados.items():
    if resultado['status'] == 'exitoso':
        print(f"{hoja}: {resultado['registros']} registros")
        total_registros += resultado['registros']
        exitosos += 1
    else:
        print(f"{hoja}: {resultado['error']}")

print()
print(f"Catalogos procesados: {exitosos}/{len(hojas)}")
print(f"Total de registros: {total_registros}")
print(f"Archivos generados en: catalogos/json_output/")
print()
print("Procesamiento completado")