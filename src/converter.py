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

def read_excel_file(file_path, sheet_name=None):
    """
    Lee un archivo Excel usando pandas.

    Args:
        file_path (str): Ruta al archivo Excel
        sheet_name (str, opcional): Nombre de la hoja a leer

    Returns:
        pandas.DataFrame: DataFrame con los datos del Excel
    """
    try:
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"✓ Hoja '{sheet_name}' leída desde {file_path}")
        else:
            df = pd.read_excel(file_path)
            print(f"✓ Archivo Excel leído: {file_path}")
        
        print(f"  - Forma: {df.shape[0]} filas, {df.shape[1]} columnas")
        print(f"  - Columnas: {list(df.columns)}")
        return df
    except FileNotFoundError:
        print(f"✗ Archivo no encontrado: {file_path}")
        sys.exit(1)
    except Exception as e:
        print(f"✗ Error al leer el archivo Excel: {e}")
        sys.exit(1)

def clean_dataframe(df):
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

    original_columns = list(df.columns)
    df.columns = df.columns.str.replace(' ', '_').str.replace('[^a-zA-Z0-9_]', '', regex=True)

    print(f"✓ DataFrame limpiado")
    if original_columns != list(df.columns):
        print(f"  - Columnas originales: {original_columns}")
        print(f"  - Columnas limpias: {list(df.columns)}")
    
    return df

def convert_to_json(df, output_file=None, pretty=True):
    """
    Convierte el DataFrame a JSON.

    Args:
        df (pandas.DataFrame): DataFrame a convertir
        output_file (str, opcional): Ruta del archivo de salida
        pretty (bool): Si el JSON debe estar formateado

    Returns:
        str: Cadena JSON
    """
    records = df.to_dict('records')
    if pretty:
        json_str = json.dumps(records, indent=2, ensure_ascii=False)
    else:
        json_str = json.dumps(records, ensure_ascii=False)
    
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(json_str)
        print(f"✓ JSON guardado en: {output_file}")
    
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
    
    df = read_excel_file(args.excel_file, args.sheet)
    
    if args.preview:
        print("\n📋 Vista previa de los datos:")
        print(df.head())
        print(f"\nTipos de datos:")
        print(df.dtypes)
        print()
    
    df = clean_dataframe(df)
    
    pretty = not args.compact
    json_str = convert_to_json(df, args.output, pretty)
    
    if not args.output:
        print("\n📄 Salida JSON:")
        print(json_str)
    
    print(f"\n✅ ¡Conversión de Excel a JSON exitosa!")
    print(f"   Registros: {len(df)}")
    print(f"   Campos: {len(df.columns)}")

if __name__ == "__main__": main()
"""
Excel to JSON Converter

Simple script to convert Excel files to JSON for local testing.
Perfect for testing data structure before MongoDB insertion.
"""

import pandas as pd
import json
import argparse
import sys
from pathlib import Path

def read_excel_file(file_path, sheet_name=None):
    """
    Read Excel file using pandas.
    
    Args:
        file_path (str): Path to the Excel file
        sheet_name (str, optional): Name of the sheet to read
    
    Returns:
        pandas.DataFrame: DataFrame containing the Excel data
    """
    try:
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"✓ Read sheet '{sheet_name}' from {file_path}")
        else:
            df = pd.read_excel(file_path)
            print(f"✓ Read Excel file: {file_path}")
        
        print(f"  - Shape: {df.shape[0]} rows, {df.shape[1]} columns")
        print(f"  - Columns: {list(df.columns)}")
        return df
    except FileNotFoundError:
        print(f"✗ File not found: {file_path}")
        sys.exit(1)
    except Exception as e:
        print(f"✗ Error reading Excel file: {e}")
        sys.exit(1)

def clean_dataframe(df):
    """
    Clean the DataFrame for JSON conversion.
    
    Args:
        df (pandas.DataFrame): Input DataFrame
    
    Returns:
        pandas.DataFrame: Cleaned DataFrame
    """
    # Handle NaN values
    df = df.fillna("")  # Replace NaN with empty strings
    
    # Clean column names (remove spaces, special chars)
    original_columns = list(df.columns)
    df.columns = df.columns.str.replace(' ', '_').str.replace('[^a-zA-Z0-9_]', '', regex=True)
    
    # Convert datetime columns to ISO string
    for col in df.select_dtypes(include=["datetime", "datetimetz"]):
        df[col] = df[col].astype(str)
    
    print(f"✓ Cleaned DataFrame")
    if original_columns != list(df.columns):
        print(f"  - Original columns: {original_columns}")
        print(f"  - Cleaned columns: {list(df.columns)}")
    
    return df

def convert_to_json(df, output_file=None, pretty=True):
    """
    Convert DataFrame to JSON.
    
    Args:
        df (pandas.DataFrame): DataFrame to convert
        output_file (str, optional): Output file path
        pretty (bool): Whether to format JSON nicely
    
    Returns:
        str: JSON string
    """
    # Convert to dictionary records (list of dicts)
    records = df.to_dict('records')
    
    # Convert to JSON
    if pretty:
        json_str = json.dumps(records, indent=2, ensure_ascii=False)
    else:
        json_str = json.dumps(records, ensure_ascii=False)
    
    # Save to file if specified
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(json_str)
        print(f"✓ Saved JSON to: {output_file}")
    
    return json_str

def main():
    parser = argparse.ArgumentParser(description="Convert Excel file to JSON")
    parser.add_argument("excel_file", help="Path to the Excel file (.xls or .xlsx)")
    parser.add_argument("-s", "--sheet", help="Sheet name to read (default: first sheet)")
    parser.add_argument("-o", "--output", help="Output JSON file (default: print to console)")
    parser.add_argument("--compact", action="store_true", help="Compact JSON output (no pretty formatting)")
    parser.add_argument("--preview", action="store_true", help="Show data preview before conversion")
    
    args = parser.parse_args()
    
    # Validate Excel file exists
    excel_path = Path(args.excel_file)
    if not excel_path.exists():
        print(f"✗ Excel file not found: {args.excel_file}")
        sys.exit(1)
    
    print(f"Excel to JSON Converter")
    print(f"=" * 30)
    print(f"Excel file: {args.excel_file}")
    print(f"Sheet: {args.sheet or 'First sheet'}")
    if args.output:
        print(f"Output: {args.output}")
    print()
    
    # Read Excel file
    df = read_excel_file(args.excel_file, args.sheet)
    
    # Show preview if requested
    if args.preview:
        print("\n📋 Data Preview:")
        print(df.head())
        print(f"\nData types:")
        print(df.dtypes)
        print()
    
    # Clean the data
    df = clean_dataframe(df)
    
    # Convert to JSON
    pretty = not args.compact
    json_str = convert_to_json(df, args.output, pretty)
    
    # Print to console if no output file specified
    if not args.output:
        print("\n📄 JSON Output:")
        print(json_str)
    
    print(f"\n✅ Successfully converted Excel to JSON!")
    print(f"   Records: {len(df)}")
    print(f"   Fields: {len(df.columns)}")

if __name__ == "__main__":
    main()
