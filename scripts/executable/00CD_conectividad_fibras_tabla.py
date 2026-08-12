import duckdb
import pandas as pd
import os
import re

RUTA_ENTRADA = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\smallworld\private_collections\00_find.txt"
RUTA_SALIDA_EXCEL = r"C:\A_GS1_PROYECTOS\reporte_fibra.xlsx"

con = duckdb.connect("auditoria_fibra.duckdb")
con.execute("SET memory_limit='8GB';")
con.execute("SET threads=4;")
con.execute("SET preserve_insertion_order=false;")

print("Iniciando procesamiento...")

con.execute(f"""
    CREATE OR REPLACE TABLE fibra_raw AS 
    SELECT * FROM read_csv(
        '{RUTA_ENTRADA}',
        delim='|',
        header=True,
        ignore_errors=True,
        nullstr=['', '<none>'],
        quote='',
        escape='',
        strict_mode=False,
        columns={{
            'SHEATH_ID': 'VARCHAR',
            'FUNDA': 'VARCHAR',
            'NOMBRE_ANTIGUO': 'VARCHAR',
            'HILO': 'VARCHAR',
            'ID_HILO': 'VARCHAR',
            'NOMBRE': 'VARCHAR',
            'ESTADO': 'VARCHAR',
            'TOPOLOGIA': 'VARCHAR',
            'TECNOLOGIA': 'VARCHAR',
            'STATUS': 'VARCHAR',
            'USO': 'VARCHAR',
            'CARACTERISTICA': 'VARCHAR',
            'RANGO': 'VARCHAR',
            'ANCHO_BANDA': 'VARCHAR'
        }}
    );
""")

con.execute("""
    CREATE OR REPLACE TABLE fibra AS
    SELECT 
        SHEATH_ID,
        FUNDA,
        REPLACE(REPLACE(REPLACE(NOMBRE_ANTIGUO, CHR(10), ''), CHR(13), ''), CHR(9), ' ') AS NOMBRE_ANTIGUO,
        HILO,
        ID_HILO,
        REPLACE(REPLACE(REPLACE(REPLACE(TRIM(NOMBRE), '\t', ' '), CHR(10), ''), CHR(13), ''), CHR(9), ' ') AS NOMBRE,
        ESTADO,
        TOPOLOGIA,
        TECNOLOGIA,
        STATUS,
        USO,
        RANGO,
        ANCHO_BANDA
    FROM fibra_raw;
""")

con.execute("DROP TABLE fibra_raw;")

total_filas = con.execute("SELECT COUNT(*) FROM fibra").fetchone()[0]
print(f"Total de registros: {total_filas:,}")

con.execute("""
    CREATE OR REPLACE TABLE fibra_con_estado AS
    SELECT 
        SHEATH_ID,
        FUNDA,
        NOMBRE_ANTIGUO,
        HILO,
        NOMBRE,
        STATUS,
        RANGO,
        CASE 
            WHEN STATUS = 'Ocupado' OR (STATUS = 'Libre' AND NOMBRE IS NOT NULL AND NOMBRE != '') 
            THEN 'OCUPADO' 
            ELSE 'LIBRE' 
        END AS estado_real,
        CASE 
            WHEN STATUS = 'Libre' AND NOMBRE IS NOT NULL AND NOMBRE != '' AND UPPER(NOMBRE) != 'RETIRO'
            THEN CAST(HILO AS VARCHAR) || ':' || NOMBRE
            ELSE NULL
        END AS error_libre_con_nombre
    FROM fibra;
""")

con.execute("""
    CREATE OR REPLACE TABLE sheath_resumen AS
    SELECT 
        FUNDA,
        SHEATH_ID,
        STRING_AGG(DISTINCT NOMBRE_ANTIGUO, ', ') AS NOMBRES_ANTIGUOS,
        COUNT(*) AS total_hilos,
        STRING_AGG(
            CASE WHEN estado_real = 'OCUPADO' 
            THEN CAST(HILO AS VARCHAR) 
            ELSE NULL END, 
            ', ' ORDER BY HILO
        ) AS hilos_ocupados_str,
        STRING_AGG(
            DISTINCT CASE WHEN estado_real = 'OCUPADO' THEN NOMBRE ELSE NULL END, 
            ', '
        ) AS clientes_sheath,
        STRING_AGG(
            DISTINCT error_libre_con_nombre,
            ', '
        ) AS errores_libre_con_nombre
    FROM fibra_con_estado
    GROUP BY FUNDA, SHEATH_ID;
""")

con.execute("""
    CREATE OR REPLACE TABLE sheath_validacion AS
    SELECT 
        FUNDA,
        SHEATH_ID,
        REPLACE(REPLACE(NOMBRES_ANTIGUOS, CHR(10), ''), CHR(13), '') AS NOMBRES_ANTIGUOS,
        total_hilos,
        REPLACE(REPLACE(hilos_ocupados_str, CHR(10), ''), CHR(13), '') AS hilos_ocupados_str,
        REPLACE(REPLACE(clientes_sheath, CHR(10), ''), CHR(13), '') AS clientes_sheath,
        REPLACE(REPLACE(errores_libre_con_nombre, CHR(10), ''), CHR(13), '') AS errores_libre_con_nombre,
        CASE 
            WHEN total_hilos != 24 THEN 'ERROR_HILOS'
            ELSE 'OK'
        END AS validacion_hilos
    FROM sheath_resumen;
""")

con.execute("DROP TABLE fibra_con_estado;")
con.execute("DROP TABLE sheath_resumen;")

con.execute("""
    CREATE OR REPLACE TABLE resumen_cable AS
    WITH 
    total_sheath_por_cable AS (
        SELECT 
            FUNDA,
            COUNT(*) AS TOTAL_SHEATHS
        FROM sheath_validacion
        GROUP BY FUNDA
    ),
    sheath_clientes_expandidos AS (
        SELECT 
            s.FUNDA,
            s.SHEATH_ID,
            UNNEST(STRING_TO_ARRAY(s.clientes_sheath, ', ')) AS cliente
        FROM sheath_validacion s
        WHERE s.clientes_sheath IS NOT NULL AND s.clientes_sheath != ''
    ),
    clientes_por_cable AS (
        SELECT 
            FUNDA,
            cliente,
            COUNT(DISTINCT SHEATH_ID) AS cantidad_sheath
        FROM sheath_clientes_expandidos
        GROUP BY FUNDA, cliente
    ),
    cable_con_clientes_comunes AS (
        SELECT 
            t.FUNDA,
            MAX(CASE 
                WHEN c.cantidad_sheath = t.TOTAL_SHEATHS THEN 1 
                ELSE 0 
            END) AS tiene_cliente_comun
        FROM total_sheath_por_cable t
        LEFT JOIN clientes_por_cable c ON t.FUNDA = c.FUNDA
        GROUP BY t.FUNDA
    ),
    cable_agrupado AS (
        SELECT 
            s.FUNDA,
            STRING_AGG(s.SHEATH_ID, ', ' ORDER BY s.SHEATH_ID) AS SHEATHS_IDS,
            STRING_AGG(DISTINCT s.NOMBRES_ANTIGUOS, ', ') AS NOMBRES_ANTIGUOS,
            LIST(s.total_hilos ORDER BY s.SHEATH_ID) AS lista_hilos,
            LIST(s.hilos_ocupados_str ORDER BY s.SHEATH_ID) AS lista_ocupados,
            STRING_AGG(DISTINCT s.clientes_sheath, ', ') AS CLIENTES_TODOS,
            STRING_AGG(s.validacion_hilos, ', ' ORDER BY s.SHEATH_ID) AS VALIDACIONES,
            cc.tiene_cliente_comun,
            STRING_AGG(
                DISTINCT s.errores_libre_con_nombre,
                ', '
            ) AS ERRORES_LIBRE_CON_NOMBRE
        FROM sheath_validacion s
        JOIN total_sheath_por_cable t ON s.FUNDA = t.FUNDA
        LEFT JOIN cable_con_clientes_comunes cc ON s.FUNDA = cc.FUNDA
        GROUP BY s.FUNDA, cc.tiene_cliente_comun
    ),
    hilos_formateados AS (
        SELECT 
            FUNDA,
            SHEATHS_IDS,
            NOMBRES_ANTIGUOS,
            lista_hilos,
            lista_ocupados,
            CLIENTES_TODOS,
            VALIDACIONES,
            tiene_cliente_comun,
            ERRORES_LIBRE_CON_NOMBRE,
            CASE 
                WHEN LEN(LIST_DISTINCT(lista_hilos)) = 1 
                THEN CAST(LIST_DISTINCT(lista_hilos)[1] AS VARCHAR)
                ELSE (
                    SELECT STRING_AGG(CAST(h AS VARCHAR), ', ')
                    FROM UNNEST(lista_hilos) AS t(h)
                )
            END AS HILOS_POR_SHEATH,
            CASE 
                WHEN LEN(LIST_DISTINCT(lista_ocupados)) = 1 
                THEN LIST_DISTINCT(lista_ocupados)[1]
                ELSE (
                    SELECT STRING_AGG(CAST(o AS VARCHAR), ', ')
                    FROM UNNEST(lista_ocupados) AS t(o)
                )
            END AS HILOS_OCUPADOS_INDIVIDUALES
        FROM cable_agrupado
    )
    SELECT 
        REPLACE(REPLACE(SHEATHS_IDS, CHR(10), ''), CHR(13), '') AS SHEATHS_IDS,
        FUNDA,
        REPLACE(REPLACE(NOMBRES_ANTIGUOS, CHR(10), ''), CHR(13), '') AS NOMBRES_ANTIGUOS,
        REPLACE(REPLACE(HILOS_POR_SHEATH, CHR(10), ''), CHR(13), '') AS HILOS_POR_SHEATH,
        REPLACE(REPLACE(HILOS_OCUPADOS_INDIVIDUALES, CHR(10), ''), CHR(13), '') AS HILOS_OCUPADOS_INDIVIDUALES,
        REPLACE(REPLACE(CLIENTES_TODOS, CHR(10), ''), CHR(13), '') AS CLIENTES,
        CASE 
            WHEN VALIDACIONES LIKE '%ERROR_HILOS%' 
            THEN 'RECHAZADO - HILOS INCONSISTENTES'
            WHEN ERRORES_LIBRE_CON_NOMBRE IS NOT NULL AND ERRORES_LIBRE_CON_NOMBRE != ''
            THEN 'RECHAZADO - LIBRE CON NOMBRE | ' || REPLACE(REPLACE(ERRORES_LIBRE_CON_NOMBRE, CHR(10), ''), CHR(13), '')
            WHEN LEN(CLIENTES_TODOS) > 0 AND CLIENTES_TODOS LIKE '%,%' AND CLIENTES_TODOS NOT LIKE '%SCJ%' 
            THEN 'RECHAZADO - CLIENTES DIFERENTES'
            WHEN CLIENTES_TODOS IS NULL OR CLIENTES_TODOS = '' 
            THEN 'SIN CLIENTES'
            ELSE 'APROBADO'
        END AS ESTADO_CALIDAD
    FROM hilos_formateados
    ORDER BY FUNDA;
""")

con.execute("DROP TABLE sheath_validacion;")

con.execute("COPY resumen_cable TO 'temp_resumen.csv' (HEADER, DELIMITER ',');")

df_resumen = pd.read_csv('temp_resumen.csv')

def limpiar_caracteres_ilegales(valor):
    if isinstance(valor, str):
        valor = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', valor)
        valor = re.sub(r'[^\x09\x0a\x0d\x20-\x7e]', '', valor)
        return valor
    return valor

df_resumen = df_resumen.applymap(limpiar_caracteres_ilegales)

with pd.ExcelWriter(RUTA_SALIDA_EXCEL, engine='openpyxl') as writer:
    df_resumen.to_excel(writer, sheet_name='Resumen', index=False)

print(f"Archivo Excel guardado en: {RUTA_SALIDA_EXCEL}")

stats = con.execute("""
    SELECT 
        COUNT(DISTINCT FUNDA) AS total_cables,
        COUNT(DISTINCT SHEATH_ID) AS total_segmentos
    FROM fibra;
""").fetchdf()

print(stats.to_string(index=False))

calidad_resumen = con.execute("""
    SELECT 
        ESTADO_CALIDAD,
        COUNT(*) AS cantidad
    FROM resumen_cable
    GROUP BY ESTADO_CALIDAD
    ORDER BY cantidad DESC;
""").fetchdf()
print(calidad_resumen.to_string(index=False))

os.remove('temp_resumen.csv')
con.close()