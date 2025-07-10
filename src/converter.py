"""
Convertidor de Excel a JSON

Script simple para convertir archivos Excel a JSON para pruebas locales.
Perfecto para probar la estructura de datos antes de insertar en MongoDB.
Incluye limpieza de acentos y ñ para encoding en español.
"""

import pandas as pd
import json
import argparse
import sys
import unicodedata
import re
from pathlib import Path

def limpiar_encoding_espanol(texto):
    if pd.isna(texto) or texto == "":
        return texto
    
    texto = str(texto)
    
    reemplazos = {
        'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u', 'ü': 'u',  
        'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U', 'Ü': 'U',
        'ñ': 'n', 'Ñ': 'N'
    }
    
    for original, reemplazo in reemplazos.items():
        texto = texto.replace(original, reemplazo)
    
    return texto

def limpiar_encoding_unicode(texto):
    """
    Método alternativo usando Unicode normalization.
    Elimina todos los acentos incluyendo ñ.
    """
    if pd.isna(texto) or texto == "":
        return texto
    
    texto = str(texto)
    
    # Normalizar Unicode y eliminar todos los acentos
    texto_nfd = unicodedata.normalize('NFD', texto)
    texto_sin_acentos = ''.join(c for c in texto_nfd if unicodedata.category(c) != 'Mn')
    
    return texto_sin_acentos

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

def limpiar_dataframe(df, limpiar_columnas=False, metodo_acentos='espanol'):
    """
    Limpia el DataFrame para la conversión a JSON.
    Reemplaza NaN por "", convierte fechas a string y limpia nombres de columnas.
    
    Args:
        df: DataFrame a limpiar
        limpiar_columnas (bool): Si debe limpiar acentos SOLO en nombres de columnas
        metodo_acentos (str): 'espanol' o 'unicode' para el método de limpieza
    """
    import numpy as np
    import datetime

    df = df.fillna("")

    # Convierte cualquier columna con al menos un valor datetime a string
    for col in df.columns:
        if df[col].apply(lambda x: isinstance(x, (datetime.datetime, datetime.date, np.datetime64))).any():
            df[col] = df[col].astype(str)

    # Limpiar nombres de columnas
    columnas_originales = list(df.columns)
    
    # Limpiar acentos en nombres de columnas si se solicita
    if limpiar_columnas:
        print(f"✓ Limpiando acentos en nombres de columnas usando método: {metodo_acentos}")
        
        if metodo_acentos == 'espanol':
            df.columns = [limpiar_encoding_espanol(col) for col in df.columns]
        else:
            df.columns = [limpiar_encoding_unicode(col) for col in df.columns]
            
        print(f"  - Ejemplo: 'Descripción' → 'Descripcion'")
    
    # Limpiar espacios y caracteres especiales en nombres de columnas
    df.columns = df.columns.str.replace(' ', '_').str.replace('[^a-zA-Z0-9_]', '', regex=True)

    print(f"✓ DataFrame limpiado")
    if columnas_originales != list(df.columns):
        print(f"  - Columnas originales: {columnas_originales[:3]}{'...' if len(columnas_originales) > 3 else ''}")
        print(f"  - Columnas limpias: {list(df.columns)[:3]}{'...' if len(df.columns) > 3 else ''}")
    
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
    parser.add_argument("--limpiar-columnas", action="store_true", help="Elimina acentos SOLO en nombres de columnas")
    parser.add_argument("--metodo-acentos", choices=['espanol', 'unicode'], default='espanol', 
                       help="Método para limpiar acentos en columnas: 'espanol' (mantiene estructura) o 'unicode' (más agresivo)")
    
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
    if args.limpiar_columnas:
        print(f"Limpieza de columnas: {args.metodo_acentos}")
    print()
    
    df = leer_excel(args.excel_file, args.sheet)
    
    if args.preview:
        print("\n📋 Vista previa de los datos:")
        print(df.head())
        print(f"\nTipos de datos:")
        print(df.dtypes)
        print()
    
    df = limpiar_dataframe(df, args.limpiar_columnas, args.metodo_acentos)
    
    bonito = not args.compact
    json_str = convertir_a_json(df, args.output, bonito)
    
    if not args.output:
        print("\n📄 Salida JSON:")
        # Solo mostrar los primeros 500 caracteres si es muy largo
        if len(json_str) > 500:
            print(json_str[:500] + "...")
            print(f"\n[JSON truncado - total: {len(json_str)} caracteres]")
        else:
            print(json_str)
    
    print(f"\n✅ ¡Conversión de Excel a JSON exitosa!")
    print(f"   Registros: {len(df)}")
    print(f"   Campos: {len(df.columns)}")
    if args.limpiar_columnas:
        print(f"   Limpieza de columnas: Aplicada ({args.metodo_acentos})")

if __name__ == "__main__":
    main()