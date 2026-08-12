import os
import re
import unicodedata
import warnings
from datetime import datetime
from io import StringIO

import pandas as pd
import psycopg2
from dotenv import load_dotenv

load_dotenv()
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

DB_CONFIG = {
    'host': os.environ.get('PGHOST', 'localhost'),
    'port': os.environ.get('PGPORT', '5432'),
    'database': os.environ.get('PGDATABASE', 'quality'),
    'user': os.environ.get('PGUSER', 'postgres'),
    'password': os.environ.get('PGPASSWORD', '')
}

GS_BASE_PATH = os.environ.get('GS_BASE_PATH')
if GS_BASE_PATH is None:
    raise ValueError("La variable de entorno GS_BASE_PATH no está definida")

ARCHIVOS_CONFIG = [
    {
        'archivo': '03-linea_cables_responsables.xlsx',
        'hoja': 'TD ACCES0S',
        'tabla': 'corporativo_origin',
        'esquema': 'raw',
        'claves_unicas': None
    },
    {
        'archivo': '03-asph_report_origin.xlsx',
        'hoja': 'Hoja1',
        'tabla': 'asphia_origin',
        'esquema': 'raw',
        'claves_unicas': None
    },
    {
        'archivo': '03-base odf´s.xlsx',
        'hoja': 'Hoja1',
        'tabla': 'baseodf_origin',
        'esquema': 'raw',
        'claves_unicas': None
    }
]

ESQUEMA_CONTROL = 'raw'
TABLA_CONTROL = 'file_import_control'


def sanitizar_columna(nombre: str) -> str:
    if not nombre:
        return 'columna'
    nombre = unicodedata.normalize('NFKD', nombre).encode('ASCII', 'ignore').decode('ASCII')
    nombre = re.sub(r'[^\w]', '_', nombre)
    nombre = re.sub(r'_+', '_', nombre)
    if nombre and nombre[0].isdigit():
        nombre = '_' + nombre
    return nombre.lower()


def generar_columnas_unicas(columnas: list) -> list:
    resultado = []
    contador = {}
    for col in columnas:
        base = sanitizar_columna(col)
        if base in contador:
            contador[base] += 1
            base = f"{base}_{contador[base]}"
        else:
            contador[base] = 0
        resultado.append(base)
    return resultado


def obtener_fecha_modificacion(ruta: str) -> datetime:
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta}")
    return datetime.fromtimestamp(os.path.getmtime(ruta))


def crear_tabla_control(conn: psycopg2.extensions.connection) -> None:
    with conn.cursor() as cur:
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {ESQUEMA_CONTROL}")
        cur.execute(f"""
            CREATE TABLE IF NOT EXISTS {ESQUEMA_CONTROL}.{TABLA_CONTROL} (
                file_name TEXT PRIMARY KEY,
                last_modified TIMESTAMP NOT NULL,
                last_imported TIMESTAMP NOT NULL DEFAULT NOW()
            )
        """)
        conn.commit()


def archivo_requiere_carga(conn: psycopg2.extensions.connection, ruta_archivo: str) -> tuple:
    fecha_actual = obtener_fecha_modificacion(ruta_archivo)
    with conn.cursor() as cur:
        cur.execute(f"""
            SELECT last_modified FROM {ESQUEMA_CONTROL}.{TABLA_CONTROL}
            WHERE file_name = %s
        """, (os.path.basename(ruta_archivo),))
        row = cur.fetchone()
    if row is None:
        return True, fecha_actual
    return row[0] != fecha_actual, fecha_actual


def actualizar_control(conn: psycopg2.extensions.connection, ruta_archivo: str, fecha: datetime) -> None:
    with conn.cursor() as cur:
        cur.execute(f"""
            INSERT INTO {ESQUEMA_CONTROL}.{TABLA_CONTROL} (file_name, last_modified, last_imported)
            VALUES (%s, %s, NOW())
            ON CONFLICT (file_name) DO UPDATE
            SET last_modified = EXCLUDED.last_modified,
                last_imported = NOW()
        """, (os.path.basename(ruta_archivo), fecha))
        conn.commit()


def cargar_excel_incremental(
    conn: psycopg2.extensions.connection,
    ruta_excel: str,
    nombre_hoja: str,
    esquema_destino: str,
    tabla_destino: str,
    claves_unicas: list = None,
    verificar_cambios: bool = True
) -> tuple:
    if not os.path.exists(ruta_excel):
        raise FileNotFoundError(ruta_excel)

    if verificar_cambios:
        requiere, fecha = archivo_requiere_carga(conn, ruta_excel)
        if not requiere:
            return False, 0
    else:
        requiere = True
        fecha = obtener_fecha_modificacion(ruta_excel)

    df = pd.read_excel(ruta_excel, sheet_name=nombre_hoja, dtype=str, engine='openpyxl')
    df.columns = generar_columnas_unicas(df.columns)
    
    with conn.cursor() as cur:
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {esquema_destino}")
        
        if not claves_unicas:
            cur.execute(f"DROP TABLE IF EXISTS {esquema_destino}.{tabla_destino} CASCADE")
            cols_def = ', '.join([f'"{col}" TEXT' for col in df.columns])
            cur.execute(f"CREATE TABLE {esquema_destino}.{tabla_destino} ({cols_def})")
            
            buffer = StringIO()
            df.to_csv(buffer, sep='\t', index=False, header=False, na_rep='\\N')
            buffer.seek(0)
            sql_copy = f"COPY {esquema_destino}.{tabla_destino} FROM STDIN (FORMAT csv, DELIMITER '\t', NULL '\\N')"
            cur.copy_expert(sql_copy, buffer)
            
            resultado = len(df)
        
        else:
            claves_validas = [k for k in claves_unicas if k in df.columns]
            if not claves_validas:
                raise ValueError(f"Ninguna de las claves únicas {claves_unicas} encontrada en las columnas: {df.columns.tolist()}")
            
            cols_def = ', '.join([f'"{col}" TEXT' for col in df.columns])
            cur.execute(f"""
                CREATE TABLE IF NOT EXISTS {esquema_destino}.{tabla_destino} (
                    {cols_def}
                )
            """)
            
            idx_name = f"idx_{tabla_destino}_unique"
            col_idx = ', '.join([f'"{k}"' for k in claves_validas])
            cur.execute(f"""
                DO $$
                BEGIN
                    IF NOT EXISTS (SELECT 1 FROM pg_indexes WHERE indexname = '{idx_name}') THEN
                        CREATE UNIQUE INDEX {idx_name} 
                        ON {esquema_destino}.{tabla_destino} ({col_idx});
                    END IF;
                END $$;
            """)
            
            temp_table = f"{tabla_destino}_temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            cur.execute(f"CREATE TEMP TABLE {temp_table} (LIKE {esquema_destino}.{tabla_destino} INCLUDING DEFAULTS)")
            
            buffer = StringIO()
            df.to_csv(buffer, sep='\t', index=False, header=False, na_rep='\\N')
            buffer.seek(0)
            sql_copy = f"COPY {temp_table} FROM STDIN (FORMAT csv, DELIMITER '\t', NULL '\\N')"
            cur.copy_expert(sql_copy, buffer)
            
            conflict_cols = ', '.join([f'"{k}"' for k in claves_validas])
            update_cols = ', '.join([f'"{col}" = EXCLUDED."{col}"' for col in df.columns])
            insert_cols = ', '.join([f'"{col}"' for col in df.columns])
            
            cur.execute(f"""
                INSERT INTO {esquema_destino}.{tabla_destino} ({insert_cols})
                SELECT {insert_cols} FROM {temp_table}
                ON CONFLICT ({conflict_cols}) DO UPDATE SET
                    {update_cols}
            """)
            
            cur.execute(f"SELECT COUNT(*) FROM {temp_table}")
            resultado = cur.fetchone()[0]
            
            cur.execute(f"DROP TABLE {temp_table}")
        
        if verificar_cambios:
            actualizar_control(conn, ruta_excel, fecha)
        
        conn.commit()
        
        cur.execute(f"SELECT COUNT(*) FROM {esquema_destino}.{tabla_destino}")
        total = cur.fetchone()[0]
    
    return True, resultado


def main() -> None:
    conn = psycopg2.connect(**DB_CONFIG)
    conn.autocommit = False
    try:
        crear_tabla_control(conn)
        for cfg in ARCHIVOS_CONFIG:
            ruta = os.path.join(GS_BASE_PATH, cfg['archivo'])
            if not os.path.exists(ruta):
                print(f"Archivo no encontrado: {ruta}")
                continue
            
            cargado, filas = cargar_excel_incremental(
                conn,
                ruta,
                cfg['hoja'],
                cfg['esquema'],
                cfg['tabla'],
                claves_unicas=cfg.get('claves_unicas'),
                verificar_cambios=True
            )
            if cargado:
                modo = "incremental (UPSERT)" if cfg.get('claves_unicas') else "completa (DROP TABLE)"
                print(f"{cfg['archivo']} -> {cfg['esquema']}.{cfg['tabla']}: {filas} filas procesadas ({modo})")
            else:
                print(f"{cfg['archivo']} sin cambios")
    finally:
        conn.close()


if __name__ == "__main__":
    main()