#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Extrae conexiones EC-OP-OT desde archivos de traza.
VERSIÓN DEFINITIVA - Identifica el anillo troncal por la cantidad de empalmes ET.
"""

import re
from pathlib import Path
from typing import List, Dict, Set, Tuple, Optional
import pandas as pd
import multiprocessing as mp
from concurrent.futures import ProcessPoolExecutor, as_completed
from collections import defaultdict

# ----------------------------------------------------------------------
# CONFIGURACIÓN
# ----------------------------------------------------------------------
CARPETA_PADRE = Path(
    r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos"
    r"\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO"
    r"\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\11_Trace_Conectividad"
)

RUTA_EXCLUSIONES = Path(
    r"C:\A_GS1_PROYECTOS\0_Documents_gs\output\quality\03-TRAZE_MS_CONNECTION.xlsx"
)

RUTA_SALIDA = Path(r"C:\A_GS1_PROYECTOS") / "Tabla_Conexiones_EC_OP_OT_final.xlsx"

# Compilar patrones regex
RE_FIBRE = re.compile(r'Fibre\s*\(\s*(.*?)\s*\)')
RE_CABLE_FO = re.compile(r"'([^']*)'")
RE_PORT = re.compile(r'Port\s*\((.*?)\)')
RE_SHELF = re.compile(r'Shelf\s*\((.*?)\)')
RE_CARD = re.compile(r'Card\s*\((.*?)\)')
RE_SLOT = re.compile(r'Slot(\d+)')
RE_BANDEJA = re.compile(r'BANDEJA\s*([A-Z])')
RE_OT_TRONCAL = re.compile(r'OT\d+_(.+)')
RE_NUMEROS = re.compile(r'\D')
RE_ET = re.compile(r"'ET\d+_\d+'")

# ----------------------------------------------------------------------
# FUNCIONES DE EXTRACCIÓN
# ----------------------------------------------------------------------
def extraer_anillo_hilo_cable_de_linea(linea: str) -> Tuple[str, str, str]:
    """Extrae anillo, hilo y cable de una línea Fibre."""
    anillo = ''
    hilo = ''
    cable = ''
    
    if 'Fibre (' in linea:
        m = RE_FIBRE.search(linea)
        if m:
            contenido = m.group(1).strip()
            if ':' in contenido:
                parts = contenido.split(':', 1)
                anillo = parts[0].strip().upper()
                hilo_num = parts[1].strip()
                hilo = RE_NUMEROS.sub('', hilo_num)
    
    if 'Cable de FO (' in linea:
        m = RE_CABLE_FO.search(linea)
        if m:
            cable = m.group(1).strip()
    
    return anillo, hilo, cable

def identificar_hilo_troncal_por_ets(lineas: List[str]) -> Tuple[str, str, str]:
    """
    REGLA PRINCIPAL: El hilo troncal es el que tiene MÁS empalmes ET únicos.
    CORREGIDO: Ahora cuenta correctamente los ETs por cada fibra.
    """
    # Diccionario: {fibra: set de ETs}
    fibra_ets = defaultdict(set)
    # Diccionario: {fibra: cable}
    fibra_cable = {}
    
    # Recorrer línea por línea para mantener el contexto
    et_actual = None
    
    for i, linea in enumerate(lineas):
        # Buscar Empalme de Fibra (ET)
        if 'Empalme de Fibra (Tipo' in linea and "ET" in linea:
            m = RE_ET.search(linea)
            if m:
                et_actual = m.group(0).strip("'")
        
        # Buscar Fibre - SOLO si hay un ET actual
        if 'Fibre (' in linea and et_actual:
            anillo, hilo, cable = extraer_anillo_hilo_cable_de_linea(linea)
            if anillo and hilo:
                fibra_key = f"{anillo}:{hilo}"
                fibra_ets[fibra_key].add(et_actual)
                if cable and fibra_key not in fibra_cable:
                    fibra_cable[fibra_key] = cable
    
    if not fibra_ets:
        return '', '', ''
    
    # Encontrar el hilo con más ETs
    # Si hay empate, tomar el que tenga el anillo más largo (menos probable que sea cliente)
    hilo_troncal = max(fibra_ets.items(), 
                       key=lambda x: (len(x[1]), len(x[0].split(':')[0])))
    hilo_troncal = hilo_troncal[0]
    
    anillo_troncal, hilo_troncal_num = hilo_troncal.split(':', 1)
    cable_troncal = fibra_cable.get(hilo_troncal, '')
    
    return anillo_troncal, hilo_troncal_num, cable_troncal

def extraer_metadatos(linea: str) -> Tuple[str, str, str]:
    """Extrae port, shelf, card de una línea."""
    port = shelf = card = ''
    
    m = RE_PORT.search(linea)
    if m:
        port = m.group(1).strip()
    m = RE_SHELF.search(linea)
    if m:
        shelf = m.group(1).strip()
    m = RE_CARD.search(linea)
    if m:
        card = m.group(1).strip()
    
    return port, shelf, card

def clasificar_puerto(port: str, shelf: str) -> str:
    """Clasifica el puerto como EC, OT, OP u OTRO."""
    if '-' in port:
        return 'EC'
    if shelf.upper().startswith('OT'):
        return 'OT'
    if shelf.upper().startswith('OP'):
        return 'OP'
    return 'OTRO'

def extraer_slot(card: str) -> str:
    m = RE_SLOT.search(card)
    return m.group(1) if m else '0'

def extraer_numero_port(port: str) -> str:
    return RE_NUMEROS.sub('', port)

def extraer_odf_patcheo(shelf: str) -> str:
    if '_' in shelf:
        return shelf.split('_')[-1]
    return ''

def extraer_cassetera(card: str) -> str:
    m = RE_BANDEJA.search(card)
    return m.group(1) if m else ''

def extraer_odf_troncal(shelf: str) -> str:
    shelf_clean = shelf.split()[0] if shelf else ''
    m = RE_OT_TRONCAL.search(shelf_clean)
    return m.group(1) if m else shelf_clean

def es_tarjeta_xc(card: str) -> bool:
    return 'XC' in card.upper() or 'XG' in card.upper()

def determinar_medio(port: str, card: str) -> str:
    if 'XG' in port.upper():
        return '0'
    if '-' in port or 'G' in port.upper():
        return '1' if es_tarjeta_xc(card) else '0'
    return '0'

# ----------------------------------------------------------------------
# PROCESAMIENTO DE ARCHIVOS
# ----------------------------------------------------------------------
class Nodo:
    __slots__ = ('tipo', 'port', 'shelf', 'card', 'hijos', 'indent', 'padre', 'idx_linea')
    def __init__(self, tipo, port, shelf, card, indent, idx_linea, padre=None):
        self.tipo = tipo
        self.port = port
        self.shelf = shelf
        self.card = card
        self.hijos = []
        self.indent = indent
        self.padre = padre
        self.idx_linea = idx_linea

def construir_arbol(lineas_port: List[Tuple[str, int]]) -> Nodo:
    """Construye árbol jerárquico basado en indentación."""
    raiz = Nodo('RAIZ', '', '', '', -1, -1)
    pila = [raiz]
    
    for linea, idx in lineas_port:
        indent = len(linea) - len(linea.lstrip())
        port, shelf, card = extraer_metadatos(linea)
        tipo = clasificar_puerto(port, shelf)
        if tipo == 'OTRO':
            continue
        
        while len(pila) > 1 and pila[-1].indent >= indent:
            pila.pop()
        
        padre = pila[-1]
        nodo = Nodo(tipo, port, shelf, card, indent, idx, padre)
        padre.hijos.append(nodo)
        pila.append(nodo)
    
    return raiz

def encontrar_ot_descendiente(nodo: Nodo) -> Optional[Nodo]:
    """Busca el primer nodo OT en el subárbol."""
    if nodo.tipo == 'OT':
        return nodo
    for hijo in nodo.hijos:
        encontrado = encontrar_ot_descendiente(hijo)
        if encontrado:
            return encontrado
    return None

def recolectar_ops(nodo: Nodo) -> List[Nodo]:
    """Recolecta todos los nodos OP en el subárbol."""
    ops = []
    if nodo.tipo == 'OP':
        ops.append(nodo)
    for hijo in nodo.hijos:
        ops.extend(recolectar_ops(hijo))
    return ops

def procesar_bloque(lineas_bloque: List[str], lineas_originales: List[str], 
                   indices_inicio: int, nombre_archivo: str,
                   anillo_troncal: str, hilo_troncal: str, cable_troncal: str) -> List[Dict]:
    """Procesa un bloque de líneas usando los datos troncales globales."""
    port_lines = []
    for i, linea in enumerate(lineas_bloque):
        if 'Port (' in linea:
            port_lines.append((linea, indices_inicio + i))
    
    if not port_lines:
        return []

    raiz = construir_arbol(port_lines)
    resultados = []

    def recorrer(nodo: Nodo):
        if nodo.tipo == 'EC':
            # Buscar OT descendiente o relacionado
            ot_node = encontrar_ot_descendiente(nodo)
            if ot_node is None and nodo.padre:
                for hermano in nodo.padre.hijos:
                    if hermano.tipo == 'OT' and hermano != nodo:
                        ot_node = hermano
                        break
                if ot_node is None:
                    actual = nodo.padre
                    while actual is not None:
                        if actual.tipo == 'OT':
                            ot_node = actual
                            break
                        actual = actual.padre

            # Recolectar OP
            ops_actuales = recolectar_ops(nodo)
            if nodo.padre and nodo.padre.tipo == 'OP':
                ops_actuales.append(nodo.padre)

            if not ops_actuales or ot_node is None:
                for hijo in nodo.hijos:
                    recorrer(hijo)
                return

            # Datos del EC
            primer_op = ops_actuales[0]
            odf_patcheo = extraer_odf_patcheo(primer_op.shelf)
            cassetera = extraer_cassetera(primer_op.card)

            # port_pacheo
            nums_op = set()
            for op in ops_actuales:
                num = extraer_numero_port(op.port)
                if num:
                    nums_op.add(int(num))
            port_pacheo = '+'.join(str(n) for n in sorted(nums_op)) if nums_op else ''

            # port_ec
            slot = extraer_slot(nodo.card)
            puerto_num = extraer_numero_port(nodo.port)
            medio = determinar_medio(nodo.port, nodo.card)
            port_ec = f"{slot}/{medio}/{puerto_num}" if slot and puerto_num else nodo.port

            odf_troncal = extraer_odf_troncal(ot_node.shelf)

            fila = {
                'archivo': nombre_archivo,
                'ec': nodo.shelf,
                'port_ec': port_ec,
                'odf_patcheo': odf_patcheo,
                'cassetera': cassetera,
                'port_pacheo': port_pacheo,
                'fibra': cable_troncal,
                'odf_troncal': odf_troncal,
                'hilo': hilo_troncal,
                'anillo': anillo_troncal,
            }
            resultados.append(fila)

        for hijo in nodo.hijos:
            recorrer(hijo)

    recorrer(raiz)
    return resultados

def procesar_archivo(ruta_str: str) -> Tuple[str, List[Dict], str]:
    """Procesa un archivo completo."""
    try:
        with open(ruta_str, 'r', encoding='utf-8', errors='ignore') as f:
            lineas = f.read().splitlines()
        
        if not lineas:
            return Path(ruta_str).stem, [], "Sin conexiones"

        # ================================================================
        # REGLA PRINCIPAL: IDENTIFICAR EL HILO TRONCAL POR CANTIDAD DE ETs
        # ================================================================
        anillo_troncal, hilo_troncal, cable_troncal = identificar_hilo_troncal_por_ets(lineas)
        
        # Si no se pudo identificar por ETs, buscar el primer OT y su fibra asociada
        if not anillo_troncal or not hilo_troncal:
            for i, linea in enumerate(lineas):
                if 'Shelf (OT' in linea:
                    # Buscar fibra asociada al OT
                    for j in range(i - 1, max(0, i - 50) - 1, -1):
                        if 'Fibre (' in lineas[j]:
                            anillo_troncal, hilo_troncal, cable_troncal = extraer_anillo_hilo_cable_de_linea(lineas[j])
                            if anillo_troncal and hilo_troncal:
                                break
                    if anillo_troncal and hilo_troncal:
                        break
        
        if not anillo_troncal or not hilo_troncal:
            return Path(ruta_str).stem, [], "Sin hilo troncal identificado"

        # Dividir en bloques por Edificio
        bloques = []
        bloque_actual = []
        for linea in lineas:
            if 'Edificio' in linea:
                if bloque_actual:
                    bloques.append(bloque_actual)
                    bloque_actual = []
            bloque_actual.append(linea)
        if bloque_actual:
            bloques.append(bloque_actual)

        if not bloques:
            bloques = [lineas]

        resultados = []
        nombre_archivo = Path(ruta_str).stem
        indices_acum = 0
        
        for bloque in bloques:
            filas = procesar_bloque(bloque, lineas, indices_acum, nombre_archivo,
                                   anillo_troncal, hilo_troncal, cable_troncal)
            resultados.extend(filas)
            indices_acum += len(bloque)

        if not resultados:
            return nombre_archivo, [], "Sin conexiones"
        return nombre_archivo, resultados, "OK"
    except Exception as e:
        return Path(ruta_str).stem, [], f"Error: {str(e)}"

# ----------------------------------------------------------------------
# PROCESAMIENTO PRINCIPAL
# ----------------------------------------------------------------------
def cargar_exclusiones(ruta: Path) -> Set[str]:
    if not ruta.exists():
        return set()
    try:
        df = pd.read_excel(ruta, sheet_name="OM", engine="openpyxl", dtype=str, header=None)
        exclusions = set()
        for col in df.columns[:2]:
            for val in df[col].dropna():
                s = str(val).strip()
                if s:
                    exclusions.add(s)
        return exclusions
    except Exception:
        return set()

def generar_tabla(carpeta_padre: Path, exclusiones: Set[str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Genera tabla procesando archivos en paralelo."""
    
    print("Recolectando archivos...")
    archivos = []
    for ext in ['*.txt', '*.TXT']:
        archivos.extend([str(p) for p in carpeta_padre.rglob(ext)])
    
    archivos_filtrados = [a for a in archivos if Path(a).stem not in exclusiones]
    total = len(archivos_filtrados)
    
    if total == 0:
        print("No se encontraron archivos")
        return pd.DataFrame(), pd.DataFrame()
    
    print(f"Procesando {total} archivos con {mp.cpu_count()} CPUs...")
    
    resultados_totales = []
    resumen = []
    procesados = 0
    
    with ProcessPoolExecutor(max_workers=mp.cpu_count()) as executor:
        futures = {executor.submit(procesar_archivo, ruta): ruta for ruta in archivos_filtrados}
        
        for future in as_completed(futures):
            procesados += 1
            if procesados % 100 == 0:
                print(f"Progreso: {procesados}/{total}")
            
            nombre, filas, estado = future.result()
            resumen.append({'archivo': nombre, 'estado': estado})
            if filas:
                resultados_totales.extend(filas)
    
    print(f"✓ Procesados {total} archivos")
    
    if not resultados_totales:
        df_conexiones = pd.DataFrame(columns=['archivo', 'ec', 'port_ec', 'odf_patcheo', 
                                             'cassetera', 'port_pacheo', 'fibra', 
                                             'odf_troncal', 'hilo', 'anillo'])
    else:
        print("Creando DataFrame...")
        df_conexiones = pd.DataFrame(resultados_totales)
        df_conexiones = df_conexiones.drop_duplicates()
        df_conexiones = df_conexiones.sort_values(['anillo', 'fibra', 'hilo', 'ec']).reset_index(drop=True)
    
    df_resumen = pd.DataFrame(resumen)
    print(f"✓ Generadas {len(df_conexiones)} conexiones")
    return df_conexiones, df_resumen

def guardar_excel(df_conexiones: pd.DataFrame, df_resumen: pd.DataFrame, ruta_salida: Path):
    """Guarda en Excel."""
    with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
        if not df_conexiones.empty:
            df_conexiones.to_excel(writer, sheet_name='Conexiones', index=False)
        else:
            pd.DataFrame({'Mensaje': ['No se encontraron conexiones']}).to_excel(
                writer, sheet_name='Conexiones', index=False
            )
        
        if not df_resumen.empty:
            df_resumen.to_excel(writer, sheet_name='Resumen', index=False)
        else:
            pd.DataFrame({'Mensaje': ['No se procesaron archivos']}).to_excel(
                writer, sheet_name='Resumen', index=False
            )
    
    print(f"✓ Archivo guardado en {ruta_salida}")

def main():
    print("=== EXTRACCION DE CONEXIONES EC-OP-OT ===")
    print("REGLAS DE PRIORIDAD:")
    print("  1. Hilo con MÁS empalmes ET (REGLAS PRINCIPAL)")
    print("  2. Fallback: fibra asociada al OT")
    print(f"Usando {mp.cpu_count()} CPUs")
    
    print("\nCargando exclusiones...")
    exclusiones = cargar_exclusiones(RUTA_EXCLUSIONES)
    print(f"Exclusiones: {len(exclusiones)}")
    
    if not CARPETA_PADRE.exists():
        print(f"Error: La carpeta no existe: {CARPETA_PADRE}")
        return
    
    df_conexiones, df_resumen = generar_tabla(CARPETA_PADRE, exclusiones)
    guardar_excel(df_conexiones, df_resumen, RUTA_SALIDA)
    print("=== PROCESO COMPLETADO ===")

if __name__ == "__main__":
    main()