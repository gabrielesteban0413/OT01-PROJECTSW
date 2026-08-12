import os
import sys
import pandas as pd
import psycopg2
from dotenv import load_dotenv

# Cargar variables de entorno
load_dotenv()
                       
DB_CONFIG = {
    'host': os.environ.get('PGHOST', 'localhost'),
    'port': os.environ.get('PGPORT', '5432'),
    'database': os.environ.get('PGDATABASE', 'quality'),
    'user': os.environ.get('PGUSER', 'postgres'),
    'password': os.environ.get('PGPASSWORD', '')
}

# Tabla a consultar: puedes cambiar a 'clean.asphia_clean' si lo prefieres
TABLA = 'staging.merge_corporativo_sync'
# Columna que contiene el identificador del cable (primer segmento de cable_acceso)
COLUMNA_CABLE = 'cable_principal'  # ajusta si usas otra columna (ej. n_acceso)

def consultar_cables(lista_cables):
    if not lista_cables:
        print("No se proporcionaron cables.")
        return pd.DataFrame()
    
    conn = psycopg2.connect(**DB_CONFIG)
    placeholders = ','.join(['%s'] * len(lista_cables))
    query = f"""
        SELECT *
        FROM {TABLA}
        WHERE {COLUMNA_CABLE} IN ({placeholders})
        ORDER BY cable_acceso;
    """
    df = pd.read_sql(query, conn, params=lista_cables)
    conn.close()
    return df

def main():
    if len(sys.argv) > 1:
        cables = sys.argv[1:]
    else:
        entrada = input("Ingrese los cables separados por espacio o coma: ")
        cables = [c.strip() for c in entrada.replace(',', ' ').split() if c.strip()]
        if not cables:
            print("No se ingresaron cables.")
            return
    
    print(f"Consultando cables: {cables}")
    df = consultar_cables(cables)
    if df.empty:
        print("No se encontraron registros para los cables especificados.")
    else:
        print(f"\nResultados ({len(df)} filas):")
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        print(df.to_string(index=False))

if __name__ == "__main__":
    main()