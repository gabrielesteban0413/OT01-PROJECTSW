"""
SISTEMA DE SINCRONIZACION DE DATOS
Version: 14.3
"""

import duckdb
import pandas as pd
import os
import sys
from datetime import datetime
import logging
import re

# ============================================================================
# CONFIGURACION
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format='%(message)s',
    handlers=[
        logging.FileHandler('sincronizacion.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# ============================================================================
# PARAMETROS EDITABLES
# ============================================================================

FECHA_INICIO = "01/08/2026"
FECHA_FIN = "10/08/2026"

RUTA_REPORTE_FIBRA = r"C:\A_GS1_PROYECTOS\reporte_fibra.xlsx"
RUTA_CLIENTES = r"C:\A_GS1_PROYECTOS\kineabaseetbClientes priorizados conectividad.xlsx"
RUTA_PORTS = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\01-Bulk export of ports.xlsx"
RUTA_SALIDA = r"C:\A_GS1_PROYECTOS\sincronizacion_resultado.xlsx"

MEMORY_LIMIT = '8GB'
THREADS = 4

# ============================================================================
# CLASE PRINCIPAL
# ============================================================================

class SincronizadorDatos:
    def __init__(self, fecha_inicio, fecha_fin):
        self.fecha_inicio = self._parse_fecha(fecha_inicio)
        self.fecha_fin = self._parse_fecha(fecha_fin)
        self.con = None
        self.df_resultado = None
        self.ports_values = set()
        self.ports_values_clean = set()
    
    def _parse_fecha(self, fecha_str):
        try:
            return datetime.strptime(fecha_str, "%d/%m/%Y")
        except ValueError:
            try:
                return datetime.strptime(fecha_str, "%Y-%m-%d")
            except:
                raise ValueError(f"Formato invalido: {fecha_str}")
    
    def _conectar_duckdb(self):
        self.con = duckdb.connect(":memory:")
        self.con.execute(f"SET memory_limit='{MEMORY_LIMIT}';")
        self.con.execute(f"SET threads={THREADS};")
        self.con.execute("SET preserve_insertion_order=false;")
    
    def _cargar_reporte_fibra(self):
        if not os.path.exists(RUTA_REPORTE_FIBRA):
            raise FileNotFoundError(f"Fibra no encontrado: {RUTA_REPORTE_FIBRA}")
        
        df_fibra = pd.read_excel(RUTA_REPORTE_FIBRA, sheet_name='Resumen', dtype=str)
        
        filas_expandidas = []
        for idx, row in df_fibra.iterrows():
            sheath_ids = str(row['SHEATHS_IDS'])
            lista_sheath = [s.strip() for s in re.split('[,;]', sheath_ids) if s.strip()]
            
            nombres = str(row['NOMBRES_ANTIGUOS'])
            lista_nombres = [n.strip() for n in re.split('[,;]', nombres) if n.strip()]
            
            for sheath in lista_sheath:
                for nombre in lista_nombres:
                    if sheath and nombre and sheath not in ['nan', 'None'] and nombre not in ['nan', 'None']:
                        nueva_fila = row.to_dict()
                        nueva_fila['SHEATH_ID_UNICO'] = sheath
                        nueva_fila['NOMBRE_ANTIGUO'] = nombre
                        filas_expandidas.append(nueva_fila)
        
        df_fibra = pd.DataFrame(filas_expandidas)
        df_fibra = df_fibra.drop_duplicates(subset=['SHEATH_ID_UNICO', 'NOMBRE_ANTIGUO'])
        
        self.con.register('fibra_df', df_fibra)
        self.con.execute("CREATE OR REPLACE TABLE fibra_raw AS SELECT * FROM fibra_df")
        
        return self.con.execute("SELECT COUNT(*) FROM fibra_raw").fetchone()[0]
    
    def _cargar_datos_clientes(self):
        if not os.path.exists(RUTA_CLIENTES):
            raise FileNotFoundError(f"Clientes no encontrado: {RUTA_CLIENTES}")
        
        df_clientes = pd.read_excel(RUTA_CLIENTES, dtype=str)
        df_clientes.columns = df_clientes.columns.str.strip().str.replace(' ', '_')
        
        col_fecha = None
        for col in df_clientes.columns:
            if col.strip() in ['Fecha', 'FECHA', 'fecha']:
                col_fecha = col
                break
        
        if col_fecha:
            def convertir_fecha(valor):
                if pd.isna(valor):
                    return None
                if isinstance(valor, (datetime, pd.Timestamp)):
                    return pd.to_datetime(valor)
                valor_str = str(valor).strip()
                for fmt in ['%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y']:
                    try:
                        return pd.to_datetime(valor_str, format=fmt)
                    except:
                        continue
                return None
            
            df_clientes['_fecha_temp'] = df_clientes[col_fecha].apply(convertir_fecha)
            mask = (df_clientes['_fecha_temp'] >= self.fecha_inicio) & (df_clientes['_fecha_temp'] <= self.fecha_fin)
            mask_nat = df_clientes['_fecha_temp'].isna()
            df_clientes = df_clientes[mask | mask_nat].drop('_fecha_temp', axis=1)
            df_clientes[col_fecha] = pd.to_datetime(df_clientes[col_fecha]).dt.strftime('%d/%m/%Y')
        
        col_cable = None
        for col in df_clientes.columns:
            if 'Cable_de_Acceso' in col or 'CABLE_DE_ACCESO' in col:
                col_cable = col
                break
            if 'cable' in col.lower() or 'acceso' in col.lower():
                col_cable = col
        
        if not col_cable:
            raise ValueError("Columna Cable_de_Acceso no encontrada")
        
        df_clientes['Cable_de_Acceso'] = df_clientes[col_cable].astype(str).str.strip()
        df_clientes['Cable_de_Acceso'] = df_clientes['Cable_de_Acceso'].replace(['nan', 'None', '', 'NULL', '-'], None)
        df_clientes = df_clientes[df_clientes['Cable_de_Acceso'].notna()]
        
        filas_expandidas = []
        for idx, row in df_clientes.iterrows():
            cables = str(row['Cable_de_Acceso'])
            lista_cables = [c.strip() for c in re.split('[,;/]', cables) if c.strip()]
            for cable in lista_cables:
                if cable and cable not in ['nan', 'None']:
                    nueva_fila = row.to_dict()
                    nueva_fila['CABLE_UNICO'] = cable
                    nueva_fila['ID_ORIGINAL'] = idx
                    filas_expandidas.append(nueva_fila)
        
        df_clientes = pd.DataFrame(filas_expandidas)
        if 'CABLE_UNICO' in df_clientes.columns:
            df_clientes.rename(columns={'CABLE_UNICO': 'Cable_de_Acceso'}, inplace=True)
        
        self.con.register('clientes_df', df_clientes)
        self.con.execute("CREATE OR REPLACE TABLE clientes_raw AS SELECT * FROM clientes_df")
        
        return self.con.execute("SELECT COUNT(*) FROM clientes_raw").fetchone()[0]
    
    def _cargar_datos_ports(self):
        if not os.path.exists(RUTA_PORTS):
            return 0
        
        df_ports = pd.read_excel(RUTA_PORTS, sheet_name='Port', dtype=str)
        
        all_values = set()
        all_values_clean = set()
        
        for col in df_ports.columns:
            valores = df_ports[col].astype(str).str.strip()
            valores = valores[valores.notna()]
            valores = valores[valores != 'nan']
            valores = valores[valores != 'None']
            valores = valores[valores != '']
            
            for v in valores:
                all_values.add(v)
                v_clean = re.sub(r'[^A-Za-z0-9]', '', v.upper())
                if v_clean:
                    all_values_clean.add(v_clean)
        
        self.ports_values = all_values
        self.ports_values_clean = all_values_clean
        
        logger.info(f"Ports: {len(all_values):,} valores | {len(all_values_clean):,} limpios")
        return len(self.ports_values)
    
    def _realizar_cruce(self):
        col_fecha = None
        df_sample = self.con.execute("SELECT * FROM clientes_raw LIMIT 1").fetchdf()
        for col in df_sample.columns:
            if 'Fecha' in col or 'FECHA' in col or 'fecha' in col:
                col_fecha = col
                break
        
        fecha_field = f'c."{col_fecha}" AS FECHA' if col_fecha else "'' AS FECHA"
        
        query = f"""
        CREATE TEMP TABLE fibra_norm AS
        SELECT 
            *,
            UPPER(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(
                COALESCE(NOMBRE_ANTIGUO, ''), '[-/_.]', '', 'g'), ' ', '', 'g'), '\\t', '', 'g')) 
            AS fibra_normalizado
        FROM fibra_raw;
        
        CREATE TEMP TABLE clientes_norm AS
        SELECT 
            *,
            UPPER(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(
                COALESCE(Cable_de_Acceso, ''), '[-/_.]', '', 'g'), ' ', '', 'g'), '\\t', '', 'g')) 
            AS cable_normalizado
        FROM clientes_raw;
        
        CREATE OR REPLACE TABLE resultado_cruce AS
        SELECT 
            c.ID_ORIGINAL,
            f.SHEATH_ID_UNICO AS SHEATHS_IDS,
            f.FUNDA,
            c.Cable_de_Acceso AS CABLE_ACCESO_CLIENTE,
            f.HILOS_OCUPADOS_INDIVIDUALES,
            c.Nombre AS NOMBRE_CLIENTE,
            c.Anillo,
            c.IDServicio,
            c.Responsable,
            c.Estatus,
            c.Observaciones,
            {fecha_field},
            CASE 
                WHEN f.SHEATH_ID_UNICO IS NULL THEN 'SYNC FIBER | SMALLWORLD DIFERENTE A BASE CORPORATIVO'
                ELSE 'OK'
            END AS ERROR
        FROM clientes_norm c
        LEFT JOIN fibra_norm f 
            ON c.cable_normalizado = f.fibra_normalizado
        """
        
        self.con.execute(query)
        self.df_resultado = self.con.execute("SELECT * FROM resultado_cruce").fetchdf()
        
        return self.df_resultado
    
    def _validar_ports(self):
        if len(self.ports_values) == 0:
            return self.df_resultado
        
        errores_actualizados = []
        for idx, row in self.df_resultado.iterrows():
            id_servicio = str(row['IDServicio']).strip()
            error_actual = row['ERROR']
            
            if id_servicio and id_servicio not in ['nan', 'None', '']:
                id_clean = re.sub(r'[^A-Za-z0-9]', '', id_servicio.upper())
                
                encontrado = False
                if id_servicio in self.ports_values:
                    encontrado = True
                elif id_clean and id_clean in self.ports_values_clean:
                    encontrado = True
                elif id_clean:
                    for pv in self.ports_values_clean:
                        if id_clean in pv or pv in id_clean:
                            encontrado = True
                            break
                
                if not encontrado:
                    if error_actual != 'OK':
                        errores_actualizados.append(error_actual)
                    else:
                        errores_actualizados.append('SYNC PORT | SMALLWORLD DIFERENTE A BASE CORPORATIVO')
                else:
                    errores_actualizados.append(error_actual)
            else:
                errores_actualizados.append(error_actual)
        
        self.df_resultado['ERROR'] = errores_actualizados
        return self.df_resultado
    
    def _agrupar_resultados(self):
        def agrupar_filas(df):
            grupos = []
            for idx, group in df.groupby('ID_ORIGINAL'):
                # Consolidar SHEATHS_IDS unicos
                sheath_ids = group['SHEATHS_IDS'].dropna().unique()
                sheath_ids = [s for s in sheath_ids if s and s not in ['nan', 'None']]
                sheath_ids_str = ', '.join(sheath_ids) if sheath_ids else ''
                
                # Consolidar HILOS_OCUPADOS_INDIVIDUALES - tomar todos los valores no nulos
                hilos_ocupados = group['HILOS_OCUPADOS_INDIVIDUALES'].dropna().unique()
                hilos_ocupados = [h for h in hilos_ocupados if h and h not in ['nan', 'None']]
                hilos_ocupados_str = ', '.join(hilos_ocupados) if hilos_ocupados else ''
                
                # Tomar primer valor de cada columna
                row_dict = {
                    'SHEATHS_IDS': sheath_ids_str,
                    'FUNDA': group['FUNDA'].iloc[0] if len(group['FUNDA'].dropna()) > 0 else '',
                    'CABLE_ACCESO_CLIENTE': group['CABLE_ACCESO_CLIENTE'].iloc[0],
                    'HILOS_OCUPADOS_INDIVIDUALES': hilos_ocupados_str,
                    'NOMBRE_CLIENTE': group['NOMBRE_CLIENTE'].iloc[0],
                    'Anillo': group['Anillo'].iloc[0],
                    'IDServicio': group['IDServicio'].iloc[0],
                    'Responsable': group['Responsable'].iloc[0],
                    'Estatus': group['Estatus'].iloc[0],
                    'Observaciones': group['Observaciones'].iloc[0],
                    'FECHA': group['FECHA'].iloc[0],
                }
                
                # Determinar error: prioridad
                # 1. SYNC FIBER (no existe fibra)
                # 2. HILO LIBRE (existe fibra pero sin hilos ocupados)
                # 3. SYNC PORT
                # 4. OK
                errores = group['ERROR'].unique()
                
                if any('SYNC FIBER' in str(e) for e in errores):
                    row_dict['ERROR'] = 'SYNC FIBER | SMALLWORLD DIFERENTE A BASE CORPORATIVO'
                elif not hilos_ocupados_str or hilos_ocupados_str == '':
                    row_dict['ERROR'] = 'SYNC FIBER | HILO LIBRE'
                elif any('SYNC PORT' in str(e) for e in errores):
                    row_dict['ERROR'] = 'SYNC PORT | SMALLWORLD DIFERENTE A BASE CORPORATIVO'
                else:
                    row_dict['ERROR'] = 'OK'
                
                grupos.append(row_dict)
            
            return pd.DataFrame(grupos)
        
        self.df_resultado = agrupar_filas(self.df_resultado)
        return self.df_resultado
    
    def _guardar_resultados(self):
        try:
            self.df_resultado.to_excel(RUTA_SALIDA, sheet_name='Cruce', index=False)
        except Exception as e:
            logger.error(f"Error guardando: {e}")
            raise
    
    def _limpiar_recursos(self):
        if self.con:
            self.con.close()
    
    def ejecutar(self):
        try:
            self._conectar_duckdb()
            self._cargar_reporte_fibra()
            self._cargar_datos_clientes()
            self._cargar_datos_ports()
            
            self.df_resultado = self._realizar_cruce()
            self.df_resultado = self._validar_ports()
            self.df_resultado = self._agrupar_resultados()
            
            total = len(self.df_resultado)
            ok = self.df_resultado[self.df_resultado['ERROR'] == 'OK'].shape[0]
            error_fiber = self.df_resultado[self.df_resultado['ERROR'].str.contains('SYNC FIBER', na=False)].shape[0]
            error_port = self.df_resultado[self.df_resultado['ERROR'].str.contains('SYNC PORT', na=False)].shape[0]
            error_hilo = self.df_resultado[self.df_resultado['ERROR'].str.contains('HILO LIBRE', na=False)].shape[0]
            
            print("=" * 60)
            print(f"Total: {total:,} | OK: {ok:,} ({ok/total*100:.1f}%)")
            print(f"SYNC FIBER: {error_fiber:,} | SYNC PORT: {error_port:,} | HILO LIBRE: {error_hilo:,}")
            print("=" * 60)
            
            self._guardar_resultados()
            print(f"Archivo: {RUTA_SALIDA}")
            return True
            
        except Exception as e:
            print(f"Error: {e}")
            import traceback
            print(traceback.format_exc())
            return False
        finally:
            self._limpiar_recursos()


# ============================================================================
# EJECUCION
# ============================================================================

if __name__ == "__main__":
    try:
        sincronizador = SincronizadorDatos(FECHA_INICIO, FECHA_FIN)
        exito = sincronizador.ejecutar()
        sys.exit(0 if exito else 1)
    except KeyboardInterrupt:
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)