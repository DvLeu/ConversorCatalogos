"""
Convertidor de Excel a JSON

Script simple para convertir archivos Excel a JSON para pruebas locales.
Perfecto para probar la estructura de datos antes de insertar en MongoDB.
"""

import pandas as pd
import json
import argparse
import sys
from pathlib import Path

def leer_excel(ruta_archivo, hoja=None):
    """
    Lee un archivo Excel usando pandas.

    Args:
        ruta_archivo (str): Ruta al archivo Excel
        hoja (str, opcional): Nombre de la hoja a leer

    Returns:
        pandas.DataFrame: DataFrame con los datos del Excel
    """
    try:
        if hoja:
            df = pd.read_excel(ruta_archivo, sheet_name=hoja)
            print(f"✓ Hoja '{hoja}' leída desde {ruta_archivo}")
        else:
            df = pd.read_excel(ruta_archivo)
            print(f"✓ Archivo Excel leído: {ruta_archivo}")
        
        print(f"  - Filas: {df.shape[0]}, Columnas: {df.shape[1]}")
        print(f"  - Columnas: {list(df.columns)}")
        return df
    except FileNotFoundError:
        print(f"✗ Archivo no encontrado: {ruta_archivo}")
        sys.exit(1)
    except Exception as e:
        print(f"✗ Error al leer el archivo Excel: {e}")
        sys.exit(1)

def limpiar_dataframe(df):
    """
    Limpia el DataFrame para la conversión a JSON.

    Args:
        df (pandas.DataFrame): DataFrame de entrada

    Returns:
        pandas.DataFrame: DataFrame limpio
    """
    df = df.fillna("")  # Reemplaza NaN por cadenas vacías

    # Convierte columnas datetime a string ISO
    for col in df.select_dtypes(include=["datetime", "datetimetz"]):
        df[col] = df[col].astype(str)

    columnas_originales = list(df.columns)
    df.columns = df.columns.str.replace(' ', '_').str.replace('[^a-zA-Z0-9_]', '', regex=True)

    print(f"✓ DataFrame limpiado")
    if columnas_originales != list(df.columns):
        print(f"  - Columnas originales: {columnas_originales}")
        print(f"  - Columnas limpias: {list(df.columns)}")
    
    return df

def convertir_a_json(df, archivo_salida=None, bonito=True):
    """
    Convierte el DataFrame a JSON.

    Args:
        df (pandas.DataFrame): DataFrame a convertir
        archivo_salida (str, opcional): Ruta del archivo de salida
        bonito (bool): Si el JSON debe estar formateado

    Returns:
        str: Cadena JSON
    """
    registros = df.to_dict('records')
    if bonito:
        json_str = json.dumps(registros, indent=2, ensure_ascii=False)
    else:
        json_str = json.dumps(registros, ensure_ascii=False)
    
    if archivo_salida:
        with open(archivo_salida, 'w', encoding='utf-8') as f:
            f.write(json_str)
        print(f"✓ JSON guardado en: {archivo_salida}")
    
    return json_str

def main():
    parser = argparse.ArgumentParser(description="Convierte un archivo Excel a JSON")
    parser.add_argument("excel_file", help="Ruta al archivo Excel (.xls o .xlsx)")
    parser.add_argument("-s", "--sheet", help="Nombre de la hoja a leer (por defecto: primera hoja)")
    parser.add_argument("-o", "--output", help="Archivo JSON de salida (por defecto: imprime en consola)")
    parser.add_argument("--compact", action="store_true", help="Salida JSON compacta (sin formato bonito)")
    parser.add_argument("--preview", action="store_true", help="Muestra una vista previa de los datos antes de convertir")
    
    args = parser.parse_args()
    
    excel_path = Path(args.excel_file)
    if not excel_path.exists():
        print(f"✗ Archivo Excel no encontrado: {args.excel_file}")
        sys.exit(1)
    
    print(f"Convertidor de Excel a JSON")
    print(f"=" * 30)
    print(f"Archivo Excel: {args.excel_file}")
    print(f"Hoja: {args.sheet or 'Primera hoja'}")
    if args.output:
        print(f"Salida: {args.output}")
    print()
    
    df = leer_excel(args.excel_file, args.sheet)
    
    if args.preview:
        print("\n📋 Vista previa de los datos:")
        print(df.head())
        print(f"\nTipos de datos:")
        print(df.dtypes)
        print()
    
    df = limpiar_dataframe(df)
    
    bonito = not args.compact
    json_str = convertir_a_json(df, args.output, bonito)
    
    if not args.output:
        print("\n📄 Salida JSON:")
        print(json_str)
    
    print(f"\n✅ ¡Conversión de Excel a JSON exitosa!")
    print(f"   Registros: {len(df)}")
    print(f"   Campos: {len(df.columns)}")

if __name__ == "__main__":
    main()
