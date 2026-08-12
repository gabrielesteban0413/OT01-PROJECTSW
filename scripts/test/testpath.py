import os
import pandas as pd
import psycopg2
from dotenv import load_dotenv
import re

# ============================================================
# 1. Cargar variables de entorno
# ============================================================
load_dotenv(encoding='utf-8-sig')

DB_CONFIG = {
    'host': os.environ.get('PGHOST', 'localhost'),
    'port': os.environ.get('PGPORT', '5432'),
    'database': os.environ.get('PGDATABASE', 'quality'),
    'user': os.environ.get('PGUSER', 'postgres'),
    'password': os.environ.get('PGPASSWORD', '')
}

print("Conectando a PostgreSQL...")
conn = psycopg2.connect(**DB_CONFIG, client_encoding='UTF8')
print("✅ Conexión exitosa")

# ============================================================
# 2. Consulta SQL
# ============================================================
query = """
SELECT 
    cambio,
    actividad,
    n_acceso,
    anillo,
    fecha,
    historial_actividades,
    estado_ciclo,
    observacion_caso_atipico
FROM clean.asphia_clean
WHERE n_acceso NOT LIKE '%-%'
ORDER BY n_acceso, fecha;
"""

df = pd.read_sql(query, conn)
conn.close()

# Convertir fecha y limpiar nulos
df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
df = df.dropna(subset=['fecha'])

# ============================================================
# 3. Contar duplicados por (n_acceso, cambio) antes de filtrar
# ============================================================
df['conteo_duplicados'] = df.groupby(['n_acceso', 'cambio'])['cambio'].transform('count')

# ============================================================
# 4. Eliminar duplicados (quedarse con el más reciente)
# ============================================================
df_sorted = df.sort_values(['n_acceso', 'cambio', 'fecha'], ascending=[True, True, False])
df_final = df_sorted.drop_duplicates(subset=['n_acceso', 'cambio'], keep='first').copy()

# ============================================================
# 5. Función para detectar caracteres extraños en texto
# ============================================================
def tiene_caracteres_extraños(texto):
    if pd.isna(texto) or texto == '':
        return False
    # Detectar caracteres no imprimibles o reemplazos típicos de codificación
    patrones = [
        r'Ã‚Â', r'â†’', r'ÃƒÂ¡', r'Â', r'\ufffd',  # caracteres mal codificados
        r'[^\x00-\x7F]'  # cualquier carácter no ASCII
    ]
    for pat in patrones:
        if re.search(pat, str(texto)):
            return True
    return False

# ============================================================
# 6. Evaluar anomalías
# ============================================================
def evaluar_anomalia(row):
    motivos = []
    
    # 1. Duplicados (si conteo_duplicados > 1)
    if row['conteo_duplicados'] > 1:
        motivos.append(f"duplicado_{row['conteo_duplicados']}_veces")
    
    # 2. Observación atípica no permitida
    if row['observacion_caso_atipico'] not in ['OK', 'SIN_EVENTO_INSTALACION']:
        motivos.append(f"obs_atipica_{row['observacion_caso_atipico']}")
    
    # 3. Estado de ciclo no deseado
    if row['estado_ciclo'] in ['CANCELADO', 'ROLL_BACK', 'SUSPENDIDO']:
        motivos.append(f"estado_{row['estado_ciclo']}")
    
    # 4. Caracteres extraños en historial_actividades
    if tiene_caracteres_extraños(row['historial_actividades']):
        motivos.append("historial_con_caracteres_extraños")
    
    # 5. Campos clave nulos o vacíos
    if pd.isna(row['cambio']) or row['cambio'] == '':
        motivos.append("cambio_vacio")
    if pd.isna(row['actividad']) or row['actividad'] == '':
        motivos.append("actividad_vacia")
    if pd.isna(row['anillo']) or row['anillo'] == '':
        motivos.append("anillo_vacio")
    
    if motivos:
        return 'INCONSISTENCIA', ' | '.join(motivos)
    else:
        return 'OK', ''

# Aplicar evaluación
df_final[['estado_validacion', 'detalle_inconsistencia']] = df_final.apply(
    lambda row: pd.Series(evaluar_anomalia(row)), axis=1
)

# ============================================================
# 7. Reordenar columnas y guardar
# ============================================================
columnas_orden = [
    'n_acceso',
    'cambio',
    'actividad',
    'anillo',
    'fecha',
    'historial_actividades',
    'estado_ciclo',
    'observacion_caso_atipico',
    'estado_validacion',
    'detalle_inconsistencia',
    'conteo_duplicados'
]
df_final = df_final[columnas_orden].sort_values(['n_acceso', 'fecha']).reset_index(drop=True)

print(f"✅ Registros únicos por (n_acceso, cambio): {len(df_final)}")
print(f"📊 Distribución de validación:\n{df_final['estado_validacion'].value_counts()}")

# Guardar CSV
df_final.to_csv('resultado_validado.csv', index=False, encoding='utf-8-sig')
print("\n✅ Archivo guardado: resultado_validado.csv")