import gc
import os
import re
import time
import warnings
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
import tkinter as tk
from dotenv import load_dotenv
from openpyxl.styles import Border, Font, Side
from tkinter import simpledialog

load_dotenv()

BASE_PATH_STR = os.getenv("GS_BASE_PATH")
if not BASE_PATH_STR:
    raise ValueError("Variable GS_BASE_PATH no encontrada en .env")

BASE_PATH = Path(BASE_PATH_STR)
PROPIETARIO_VALIDO = os.getenv("PROPIETARIO_VALIDO", "DESCONOCIDO")

RUTA_MACRO          = BASE_PATH / "04-AuditoriaCalidadVisualAtributiva.xlsm"
RUTA_RESPONSABLES   = BASE_PATH / "03-linea_cables_responsables.xlsx"
RUTA_ODFS           = BASE_PATH / "04-Sync_SW_SP.xlsx"
RUTA_ASPHIA_CLIENTES= BASE_PATH / "03-asph_report_origin.xlsx"
RUTA_EDIFICIOS_BULK = BASE_PATH / "01-Bulk export of buildings.xlsx"
RUTA_FIBRAS_BULK    = BASE_PATH / "01-Bulk export of fibers.xlsx"
RUTA_EMPALMES_BULK  = BASE_PATH / "01-Bulk export of splices.xlsx"
RUTA_PORTS_BULK     = BASE_PATH / "01-Bulk export of ports.xlsx"

COLUMNAS_MASIVO = {
    "FIBRA": {
        "id": "id text",
        "nombre": "name",
        "colocacion": "colocacion type",
        "especificacion": "nombre especificación",
        "propietario": "propietario",
        "estado": "construction status",
        "nombre_ant": "nombre antiguo",
        "nodo_a": "nodo a",
        "nodo_b": "nodo b",
        "municipio_a": "municipio a id",
        "municipio_b": "municipio b id",
    },
    "EMPALME": {
        "id": "id texto",
        "nombre": "name",
        "especificacion": "nombre especificación",
        "propietario": "propietario",
        "estado": "construction status",
    },
    "EDIFICIO": {
        "id": "idtexto",
        "nombre": "name",
        "tipo": "type",
        "categoria": "categoria",
    },
    "PORT": {
        "id": "id",
        "description": "description",
        "tipo_senal": "tipo señal",
        "physical_status": "Physical Status",
        "type": "type",
        "connector_type": "connector type",
        "ancho_banda": "ancho de banda",
        "unidad_ancho_banda": "unidad ancho de banda",
        "service_level": "service level",
        "proyecto_red": "proyecto de red",
        "contratista": "contratista",
        "id_servicio": "id servicio",
        "equipo_central": "equipo central",
        "puerto_central": "puerto central",
        "tipo_servicio": "tipo servicio",
        "nombre_cliente": "nombre cliente",
        "direccion_cliente_id": "direccion cliente id",
        "comentarios": "comentarios",
        "connected": "connected?",
        "anillo": "anillo",
        "mit_card": "mit card",
        "direccion_puerto": "direccion puerto",
        "num_factibilidad": "numero de factibilidad-cable",
        "id_shelf": "id shelf",
        "cambio_service_manager": "cambio service manager",
    },
    "SHELF": {
        "id": "id shelf",
        "description": "description shelf",
        "estado": "construction status",
        "tipo": "tipo",
        "fabricante": "fabricante",
        "tecnologia": "tecnologia",
        "direccion_ip": "direccion ip",
        "nivel_bastidor": "nivel en bastidor",
        "contratista": "contratista",
        "comentarios": "comentarios",
        "nombre_anterior": "nombre anterior",
        "migracion_olt": "migracion olt",
        "id_bay": "bay.id",
    },
    "BAY": {
        "id": "bay.id",
        "number": "number",
        "description": "description bay",
        "estado": "construction status",
        "numero_salon": "numero de salon",
        "numero_fila": "numero de fila",
        "contratista": "contratista",
        "comentarios": "comentarios",
        "id_edificio": "building.id",
        "cambio_service_manager": "cambio service manager",
    },
}

COLUMNAS_ESTRUCTURA = {
    "nombre": "name",
    "floor_dimension_length": "floor dimension (length)",
    "floor_dimension_width": "floor dimension (width)",
    "idtexto": "idtexto",
    "number_of_rooms_length": "number of rooms (length)",
    "number_of_rooms_width": "number of rooms (width)",
    "number_of_floors": "number of floors",
}

def _agregar_errores_whitespace(jerarquia):
    for key in jerarquia:
        if "errores" in jerarquia[key]:
            if "ESPACIOS_DOBLES" not in jerarquia[key]["errores"]:
                jerarquia[key]["errores"].append("ESPACIOS_DOBLES")
            if "TABS" not in jerarquia[key]["errores"]:
                jerarquia[key]["errores"].append("TABS")
    return jerarquia

 # NO BORRAR (MERGE : SYNC PORT | SMALLWORLD DIFERENTE A BASE CORPORATIVO, SYNC PORT | SMALLWORLD TRAZE INEXISTENTE , )   

JERARQUIA_FIBRA = _agregar_errores_whitespace({
    1:   {"campo": "Id",                            "errores": ["DUPLICADO", "CONSECUTIVO"]},
    2:   {"campo": "Segmento",                      "errores": []},
    3:   {"campo": "Nombre",                        "errores": ["NOMBRE (Valor Inválido)", "NOMBRE (Debe estar en MAYÚSCULAS)", "NOMBRE (Corregir a", "NOMBRE (en", "COLOCACION (Troncal)", "COLOCACION (Ultima Milla)", "DE", "CARACTER INVALIDO", "el símbolo / solo se permite", "nombre no puede ser solo", "NOMBRE"]},
    4:   {"campo": "Nombre Antiguo",                "errores": ["NOMBRE_ANT", "NOMBRE_ANT (Falta guion cuando NOMBRE tiene guion)"]},
    5:   {"campo": "Municipio A",                   "errores": ["MUNICIPIO (Vacío)", "MUNICIPIO A INCONSISTENTE", "MUNICIPIO"]},
    6:   {"campo": "Municipio B",                   "errores": ["MUNICIPIO (Vacío)", "MUNICIPIO B INCONSISTENTE", "MUNICIPIO"]},
    7:   {"campo": "Nodo A",                        "errores": ["NODO INVÁLIDO", "NODO A Y B"]},
    8:   {"campo": "Nodo B",                        "errores": ["NODO INVÁLIDO", "NODO A Y B"]},
    9:   {"campo": "Especificación",                "errores": ["ESPECI (senc)", "ESPECI (Tipo 2)", "ESPECI ('revisar'Ultima Milla cantidad de Hilos)", "ESPECI (NO acorde a lo establecido)", "ESPECI EA (revisar Tipo)", "ESPECI INCONSISTENTE EN CABLE TRONCAL", "ESPECI INCONSISTENTE"]},
    10:  {"campo": "Estado de Construcción",        "errores": ["STATUS ADMITIDOS (En servicio/Plan Abandono/Futuro)", "C.STATUS", "STATUS"]},
    11:  {"campo": "Propietario",                   "errores": [f"PROPIETARIO ({PROPIETARIO_VALIDO})"]},
    12:  {"campo": "Tipo de Colocacion",            "errores": ["TIPO DE COLOCACION", "COLOCACION"]},
    13:  {"campo": "Longitud Calculada de Cable",   "errores": []},
    14:  {"campo": "Longitud Calculada de Fibra",   "errores": []},
    15:  {"campo": "Perdida Db Medida por Inserccion", "errores": []},
    16:  {"campo": "Fecha de Instalacion",          "errores": []},
    17:  {"campo": "Abcisa Inicial",                "errores": []},
    18:  {"campo": "Abcisa Final",                  "errores": []},
    19:  {"campo": "Reservas",                      "errores": []},
    20:  {"campo": "Contratista",                   "errores": []},
    21:  {"campo": "Comentarios",                   "errores": []},
    999: {"campo": "MERGE",                         "errores": ["TIPO DE CABLE INCONSISTENTE", "MERGE", "PROBLEMA EN FIBRA", "PROBLEMA EN SUB-ANILLO", "RECOMENDACIÓN"]},
})

JERARQUIA_EMPALME = _agregar_errores_whitespace({
    1:   {"campo": "Id",                            "errores": ["CONSECUTIVO", "DUPLICADO"]},
    2:   {"campo": "Segmento",                      "errores": []},
    3:   {"campo": "Nombre",                        "errores": ["NOMBRE(ERRONEO)", "NOMBRE(CONTIENE ESPACIOS)", "NOMBRE(LETRA CONSECUTIVA)", "NOMBRE(CARACTERES ERRONEOS)", "usar:"]},
    4:   {"campo": "Propietario",                   "errores": [f"PROPIETARIO ({PROPIETARIO_VALIDO})"]},
    5:   {"campo": "Municipio",                     "errores": []},
    6:   {"campo": "Estado de Construcción",        "errores": ["STATUS (En servicio/Plan Abandono/Futuro)"]},
    7:   {"campo": "Tipo de Caja",                  "errores": []},
    8:   {"campo": "Especificación",                "errores": ["ESPECIFICACION ET ERRONEA", "ESPECIFICACION (NO acorde a lo establecido)", "TIPO DE GRUPO", "esperado:", "actual:"]},
    9:   {"campo": "E2",                            "errores": []},
    10:  {"campo": "Fecha de Instalacion",          "errores": []},
    11:  {"campo": "Distancia a Splitter 1er nivel","errores": []},
    12:  {"campo": "Distancia a Splitter 2do nivel","errores": []},
    13:  {"campo": "Disponible Para",               "errores": []},
    14:  {"campo": "Tipo Restriccion",              "errores": []},
    15:  {"campo": "Fecha Inicio Restriccion",      "errores": []},
    16:  {"campo": "Responsable Restriccion",       "errores": []},
    17:  {"campo": "Sangria?",                      "errores": []},
    18:  {"campo": "Método",                        "errores": []},
    19:  {"campo": "fx Available Mounting Positions","errores": []},
    20:  {"campo": "Tiene Riser?",                  "errores": []},
    21:  {"campo": "Red Intema",                    "errores": []},
    22:  {"campo": "Puertos Derivacion Libres",     "errores": []},
    23:  {"campo": "Contratista",                   "errores": []},
    999: {"campo": "MERGE",                         "errores": ["TIPO DE CABLE INCONSISTENTE", "INCONSISTENCIA MERGE"]},
})

JERARQUIA_EDIFICIO = _agregar_errores_whitespace({
    1:  {"campo": "Id",                     "errores": []},
    2:  {"campo": "Nombre",                 "errores": ["NOMBRE (Corregir a", "NOMBRE", "NOMBRE (Valor Inválido)", "NOMBRE (Sufijo con espacios)"]},
    3:  {"campo": "Tipo",                   "errores": ["TIPO DE GRUPO"]},
    4:  {"campo": "Estado de Construcción", "errores": ["ESTADO"]},
    5:  {"campo": "Categoria",              "errores": ["CATEGORIA (AAA)"]},
    6:  {"campo": "Direccion Edificio",     "errores": []},
    7:  {"campo": "Numero de Factibilidad", "errores": []},
    8:  {"campo": "Comentarios",            "errores": []},
    9:  {"campo": "Limite",                 "errores": []},
    10: {"campo": "Direccion Cliente ID",   "errores": []},
    11: {"campo": "Posición",               "errores": []},
    12: {"campo": "Posición",               "errores": []},
})

JERARQUIA_ESTRUCTURA = _agregar_errores_whitespace({
    1: {"campo": "ESTRUCTURA", "errores": ["ESTRUCTURA (NUMBER_OF + FLOOR + ROOMS = 0)", "ESTRUCTURA_FLOOR_DIMENSION", "ESTRUCTURA_NUMBER_OF_ROOMS", "NUM_OF:", "FLOOR:", "ROOMS:"]},
})

JERARQUIA_PORT = _agregar_errores_whitespace({
    1:  {"campo": "Id",                         "errores": ["DUPLICADO", "CONSECUTIVO"]},
    2:  {"campo": "Tipo Señal",                  "errores": ["TIPO SEÑAL INVALIDO", "PUERTO RX/TX IDENTIFICADO"]},
    3:  {"campo": "Description",                 "errores": ["DESCRIPTION VACIO"]},
    4:  {"campo": "Physical Status",             "errores": ["PHYSICAL STATUS INVALIDO", "PHYSICAL STATUS VACIO"]},
    5:  {"campo": "Type",                        "errores": ["TYPE INVALIDO", "TYPE VACIO", "TYPE (solo Equipo Cliente)"]},
    6:  {"campo": "Connector Type",              "errores": ["CONNECTOR TYPE INVALIDO", "CONNECTOR TYPE VACIO", "CONNECTOR TYPE (solo Unknown o Biconic)"]},
    7:  {"campo": "Ancho De Banda",              "errores": ["ANCHO DE BANDA INVALIDO"]},
    8:  {"campo": "Unidad Ancho De Banda",       "errores": ["UNIDAD ANCHO DE BANDA INVALIDA", "UNIDAD ANCHO DE BANDA VACIA"]},
    9:  {"campo": "Service Level",               "errores": ["SERVICE LEVEL INVALIDO", "SERVICE LEVEL VACIO", "SERVICE LEVEL (solo Fiber)"]},
    10: {"campo": "Proyecto De Red",             "errores": ["PROYECTO DE RED VACIO"]},
    11: {"campo": "Contratista",                 "errores": ["CONTRATISTA VACIO"]},
    12: {"campo": "Id Servicio",                 "errores": ["ID SERVICIO FORMATO INVÁLIDO"]},
    13: {"campo": "Equipo Central",              "errores": ["EQUIPO CENTRAL INVALIDO", "EQUIPO CENTRAL VACIO", "PREFIJO EQUIPO CENTRAL (BO,VI,PA)", "MINIMO 4 NUMEROS", "PRIMEROS 6 CARACTERES DEBEN SER LETRAS"]},
    14: {"campo": "Puerto Central",              "errores": ["PUERTO CENTRAL INVALIDO", "PUERTO CENTRAL VACIO", "PUERTO CENTRAL INVALIDO", "PUERTO CENTRAL INVALIDO", "PUERTO CENTRAL INVALIDO"]},
    15: {"campo": "Tipo Servicio",               "errores": ["TIPO SERVICIO VACIO"]},
    16: {"campo": "Nombre Cliente",              "errores": ["NOMBRE CLIENTE (Valor Inválido)", "NOMBRE CLIENTE (Debe estar en MAYÚSCULAS)", "NOMBRE (Corregir a", "NOMBRE CLIENTE VACIO", "NOMBRE CLIENTE", "NOMBRE CLIENTE (Sufijo)"]},
    17: {"campo": "Direccion Cliente Id",        "errores": ["DIRECCION CLIENTE"]},
    18: {"campo": "Comentarios",                 "errores": []},
    19: {"campo": "Connected?",                  "errores": ["CONNECTED INVALIDO"]},
    20: {"campo": "Anillo",                      "errores": ["ANILLO VACIO"]},
    21: {"campo": "Mit Card",                    "errores": ["MIT CARD VACIO"]},
    22: {"campo": "Direccion Puerto",            "errores": ["DIRECCION PUERTO VACIO"]},
    23: {"campo": "Numero De Factibilidad-Cable","errores": ["NUMERO FACTIBILIDAD VACIO"]},
    24: {"campo": "Id shelf",                    "errores": ["ID SHELF VACIO", "ID SHELF INVALIDO"]},
    25: {"campo": "Cambio Service Manager",      "errores": ["SERVICE MANAGER"]},
    26: {"campo": "Id Servicio",                 "errores": ["ID SERVICIO"]},
})

JERARQUIA_SHELF = _agregar_errores_whitespace({
    1:  {"campo": "Id",                 "errores": ["DUPLICADO", "CONSECUTIVO"]},
    2:  {"campo": "Description",        "errores": ["DESCRIPTION VACIO", "DESCRIPTION INVALIDA", "DESCRIPTION CON ESPACIOS"]},
    3:  {"campo": "Construction Status","errores": ["CONSTRUCTION STATUS INVALIDO"]},
    4:  {"campo": "Tipo",               "errores": ["TIPO INVALIDO"]},
    5:  {"campo": "Fabricante",         "errores": []},
    6:  {"campo": "Tecnologia",         "errores": ["TECNOLOGIA VACIO"]},
    7:  {"campo": "Direccion Ip",       "errores": ["DIRECCION IP INVALIDA"]},
    8:  {"campo": "Nivel En Bastidor",  "errores": ["NIVEL BASTIDOR INVALIDO"]},
    9:  {"campo": "Contratista",        "errores": ["CONTRATISTA VACIO"]},
    10: {"campo": "Comentarios",        "errores": []},
    11: {"campo": "Nombre Anterior",    "errores": []},
    12: {"campo": "Migracion Olt",      "errores": ["MIGRACION OLT INVALIDA"]},
    13: {"campo": "Bay.id",             "errores": ["BAY ID VACIO", "BAY ID INVALIDO"]},
})

JERARQUIA_BAY = _agregar_errores_whitespace({
    1: {"campo": "Id",                  "errores": ["DUPLICADO", "CONSECUTIVO"]},
    2: {"campo": "Number",              "errores": ["NUMBER INVALIDO"]},
    3: {"campo": "Description",         "errores": ["DESCRIPTION VACIO", "DESCRIPTION CON ESPACIADO", "DESCRIPTION INVÁLIDA"]},
    4: {"campo": "Construction Status", "errores": ["CONSTRUCTION STATUS INVALIDO"]},
    5: {"campo": "Numero De Salon",     "errores": ["NUMERO SALON INVALIDO"]},
    6: {"campo": "Numero De Fila",      "errores": ["NUMERO FILA INVALIDO"]},
    7: {"campo": "Contratista",         "errores": ["CONTRATISTA VACIO"]},
    8: {"campo": "Comentarios",         "errores": []},
    9: {"campo": "Building.id",         "errores": ["BUILDING ID VACIO", "BUILDING ID INVALIDO"]},
})

RE_NOMBRE_FIBRA_TRONCAL        = re.compile(r"^[\dP/]+$")
RE_NOMBRE_FIBRA_TRONCAL_SIMPLE = re.compile(r"^[\d]+$")
RE_NOMBRE_FIBRA_UM_CON_P       = re.compile(r"^[\d\-P/]+$")
RE_NOMBRE_FIBRA_UM_SIN_P       = re.compile(r"^[\d\-]+$")
RE_NOMBRE_EMPALME_VALIDO       = re.compile(r"^(ET|NAP|EA|OB)[0-9_\-P]+$")
RE_NOMBRE_EMPALME_SUB          = re.compile(r"^(ET|NAP)\d+\-SUB_\d+$")
RE_NOMBRE_EMPALME_BASE         = re.compile(r"^(ET|NAP|EA|OB)[0-9\-_P]*")
RE_NOMBRE_FIBRA_BASE           = re.compile(r'^(\d+(?:-\d+)?)')
RE_NOMBRE_FIBRA_NUMERO         = re.compile(r'^(\d+)')
RE_ESPEC_HILOS                 = re.compile(r'(\d+)\s+hilos', re.IGNORECASE)
RE_NOMBRE_ANT_UM               = re.compile(r"^[\d\-P\#/]+$|^SUB$")
RE_NOMBRE_ANT_NO_UM            = re.compile(r"^[\d\-P\#]+$|^SUB$")
RE_NUMERO_SALON                = re.compile(r'^OF \d+$')
RE_DESCRIPCION_BASTIDOR        = re.compile(r'^BASTIDOR CLIENTE (\d+)$')
RE_DESCRIPCION_DEMARCADOR      = re.compile(r'^DEMARCADOR CLIENTE (\d+)$')
RE_IPV4                        = re.compile(r'^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})(/\d{1,2})?$')

PREFIJOS_EXCLUIR_VALIDACION = {"C.C", "ANT", "(SED)", "(REDP)", "COLEGIO", "TM", "EDIFICIO", "IED", "TORRE","NODO"}
PREFIJOS_ID_SERVICIO        = ("CAV", "ID", "IP", "PI", "SEG", "TI","CAV30-C")
PREFIJOS_SERVICE_MANAGER    = ("OC-", "C", "WO", "APR")
PREFIJOS_EQUIPO_CENTRAL     = ("BO", "VI", "PA","BOBSH930601","BOCHHU930601","CHCHHU930601","CUZPHU930601","SOPOHU930301","BOCADEAMEHUC1A","BOSIBATHUC1A")

UMBRAL_IGNORAR_NOMBRE_CLIENTE = 5

# ----------------------------------------------------------------------
# Funciones auxiliares (MODIFICADAS para mostrar valor original en errores whitespace)
# ----------------------------------------------------------------------
def detectar_errores_whitespace(valor, nombre_campo=""):
    """
    Detecta espacios dobles o tabs en un valor y retorna una lista de tuplas (tipo_error, valor_original)
    """
    if valor is None or not isinstance(valor, str):
        return []
    errores = []
    valor_str = str(valor)
    valor_str_original = valor_str  # guardamos para mostrar en el mensaje
    
    if '\t' in valor_str:
        errores.append(("TABS", valor_str_original))
    if '  ' in valor_str:
        errores.append(("ESPACIOS_DOBLES", valor_str_original))
    if valor_str.startswith(' ') or valor_str.endswith(' '):
        # Añadir solo si no es un doble espacio ya capturado
        if ("ESPACIOS_DOBLES", valor_str_original) not in errores:
            errores.append(("ESPACIOS_DOBLES", valor_str_original))
    # Eliminar duplicados exactos (mismo tipo y mismo valor)
    return list(dict.fromkeys(errores))

def limpiar_valor_reporte(valor):
    if valor is None or not isinstance(valor, str):
        return valor
    valor_limpio = valor.replace('\t', ' ')
    valor_limpio = re.sub(r'\s+', ' ', valor_limpio)
    valor_limpio = valor_limpio.strip()
    return valor_limpio

def validar_valor_con_whitespace(valor, nombre_campo, errores_lista):
    """
    Valida un valor y agrega errores de whitespace a la lista de errores,
    incluyendo el valor original en el mensaje.
    """
    if valor is None:
        return errores_lista
    errores_ws = detectar_errores_whitespace(valor, nombre_campo)
    for tipo_err, valor_orig in errores_ws:
        errores_lista.append(f"{nombre_campo.upper()}: {tipo_err} - valor: '{valor_orig}'")
    return errores_lista

def es_valor_cero(valor_raw, valor_str: str) -> bool:
    if valor_str == "0":
        return True
    if isinstance(valor_raw, (int, float)) and valor_raw == 0:
        return True
    try:
        return float(valor_str) == 0.0
    except (ValueError, TypeError):
        return False

def normalizar_id(valor):
    if valor is None or valor == "":
        return ""
    if isinstance(valor, (int, float)):
        try:
            return str(int(valor))
        except ValueError:
            return str(valor).strip()
    return str(valor).strip()

def traducir_estado_construccion(estado, modo: str):
    if not estado:
        return ""
    estado = str(estado).strip()
    if modo == "masivo":
        return {"is": "En servicio", "pa": "Plan Abandono", "pf": "Plan Uso Futuro"}.get(estado.lower(), estado)
    return estado

def extraer_indices_columnas(worksheet):
    result = {}
    for idx, cell in enumerate(worksheet[1], start=1):
        if cell.value:
            val = str(cell.value).strip().lower()
            result[val] = idx - 1
    return result

def validar_sufijos_nombre_empresa(nombre):
    if not nombre or pd.isna(nombre):
        return True, None, None
    nombre_str = str(nombre).strip()
    nombre_upper = nombre_str.upper()
    patrones_sufijos = [
        (r'\s+S\s+A\s+S?$', 'SA'),
        (r'\s+L\s+T\s+D\s+A$', 'LTDA'),
        (r'\s+L\s+T\s+D$', 'LTD'),
        (r'\s+I\s+N\s+C$', 'INC'),
        (r'\s+C\s+O\s+R\s+P$', 'CORP'),
        (r'\s+C\s+I\s+E$', 'CIE'),
        (r'\s+E\s+U$', 'EU'),
        (r'\s+N\s+I\s+T\s*\d*$', 'NIT'),
    ]
    for patron, _ in patrones_sufijos:
        if re.search(patron, nombre_upper):
            nombre_sin_espacios = re.sub(r'\s+', '', nombre_upper)
            msg = f"{JERARQUIA_EDIFICIO[2]['errores'][3]}'"
            return False, msg, nombre_sin_espacios
    return True, None, None

def validar_formato_nombre_cliente(nombre_original: str):
    nombre_upper = nombre_original.upper()
    errores = []
    patron_caracteres = r"^[A-Za-z0-9 ñÑ()&.,\-]+$"
    if not re.match(patron_caracteres, nombre_original):
        caracteres_invalidos = re.sub(r"[A-Za-z0-9 ñÑ()&.,\-]", "", nombre_original)
        unicos = "".join(sorted(set(caracteres_invalidos)))
        if unicos:
            errores.append(f"CARACTER INVALIDO: '{unicos}'")
    if nombre_original != nombre_upper:
        errores.append(f"{nombre_upper}")
    if errores:
        return False, f"NOMBRE CLIENTE: {' | '.join(errores)}", None
    return True, None, None

def validar_unidad_ancho_banda(valor_unidad):
    if pd.isna(valor_unidad) or str(valor_unidad).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_unidad).strip()
    if valor_str.lower() in ("megabyte", "kilobyte"):
        return "VÁLIDO", valor_str
    sugerencias = {
        "mb": "Megabyte", "mbps": "Megabyte", "mbyte": "Megabyte",
        "kb": "Kilobyte", "kbps": "Kilobyte", "kbyte": "Kilobyte",
    }
    for clave, sugerencia in sugerencias.items():
        if clave in valor_str.lower():
            return "UNIDAD_INCORRECTA", sugerencia
    return "INVÁLIDO", "Megabyte o Kilobyte"

def validar_physical_status(valor_status):
    if pd.isna(valor_status) or str(valor_status).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_status).strip()
    if valor_str.lower() in ("in service", "spare"):
        return "VÁLIDO", valor_str
    return "INVÁLIDO", "In Service o Spare"

def validar_service_level_fiber(valor_service_level):
    if pd.isna(valor_service_level) or str(valor_service_level).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_service_level).strip()
    if valor_str.lower() == "fiber":
        return "VÁLIDO", valor_str
    return "INVÁLIDO", "Fiber"

def validar_type_equipo_cliente(valor_type):
    if pd.isna(valor_type) or str(valor_type).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_type).strip()
    if valor_str.lower() == "equipo cliente":
        return "VÁLIDO", valor_str
    return "INVÁLIDO", "Equipo Cliente"

def validar_connector_type(valor_connector_type):
    if pd.isna(valor_connector_type) or str(valor_connector_type).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_connector_type).strip()
    if valor_str.lower() in ("unknown", "biconic"):
        return "VÁLIDO", valor_str
    return "INVÁLIDO", "Unknown o Biconic"

def validar_ipv4(direccion_ip):
    if pd.isna(direccion_ip) or str(direccion_ip).strip() == "":
        return "VACÍO", True
    direccion_str = str(direccion_ip).strip()
    match = RE_IPV4.match(direccion_str)
    if not match:
        return "INVÁLIDO", False
    for i in range(1, 5):
        if not (0 <= int(match.group(i)) <= 255):
            return "INVÁLIDO", False
    if match.group(5):
        mascara = int(match.group(5)[1:])
        if not (0 <= mascara <= 32):
            return "INVÁLIDO", False
    return "VÁLIDO", True

def validar_descripcion_bastidor(descripcion):
    if pd.isna(descripcion) or str(descripcion).strip() == "":
        return ["DESCRIPTION VACIO"]
    desc_str = str(descripcion).strip().upper()
    if desc_str == "BASTIDOR CLIENTE":
        return []
    match = RE_DESCRIPCION_BASTIDOR.match(desc_str)
    if match and int(match.group(1)) >= 2:
        return []
    return [f"DESCRIPTION INVALIDA. ACTUAL:'{str(descripcion).strip()}'"]

def validar_descripcion_demarcador(descripcion):
    if pd.isna(descripcion) or str(descripcion).strip() == "":
        return ["DESCRIPTION VACIO"]
    desc_str = str(descripcion).strip().upper()
    if desc_str == "DEMARCADOR CLIENTE":
        return []
    match = RE_DESCRIPCION_DEMARCADOR.match(desc_str)
    if match and int(match.group(1)) >= 2:
        return []
    return [f"DESCRIPTION INVALIDA. ACTUAL:'{str(descripcion).strip()}'"]

def validar_equipo_central(valor):
    if pd.isna(valor):
        return "VÁLIDO", ""
    valor_str = str(valor).strip()
    if valor_str == "" or valor_str == "0":
        return "VÁLIDO", ""
    if not re.match(r'^[A-Z0-9]+$', valor_str):
        return "CARACTERES_INVALIDOS", "Solo letras mayúsculas y números"
    if len(valor_str) < 6:
        return "LONGITUD_INSUFICIENTE", "Debe tener al menos 6 caracteres"
    primeros_6 = valor_str[:6]
    if not primeros_6.isalpha():
        return "PRIMEROS_6_NO_LETRAS", "Los primeros 6 caracteres deben ser letras"
    prefijo = primeros_6[:2]
    if prefijo not in PREFIJOS_EQUIPO_CENTRAL:
        return "PREFIJO_INVALIDO", f"Las primeras dos letras deben ser {', '.join(PREFIJOS_EQUIPO_CENTRAL)}"
    digitos = sum(c.isdigit() for c in valor_str)
    if digitos < 4:
        return "NUMEROS_INSUFICIENTES", "Debe contener al menos 4 números"
    return "VÁLIDO", valor_str

def validar_puerto_central(valor):
    if pd.isna(valor):
        return "VÁLIDO", ""
    valor_str = str(valor).strip()
    if valor_str == "" or valor_str == "0":
        return "VÁLIDO", ""
    if not re.match(r'^[0-9/-]+$', valor_str):
        return "CARACTERES_INVALIDOS", "Solo números, / y -"
    if re.search(r'[^0-9/-]', valor_str):
        return "CARACTERES_INVALIDOS", "Solo números, / y -"
    if '/' not in valor_str and '-' not in valor_str:
        return "SIN_SEPARADOR", "Debe contener al menos un / o -"
    separadores = valor_str.count('/') + valor_str.count('-')
    if separadores > 2:
        return "DEMASIADOS_SEPARADORES", "Máximo 2 separadores (/ o -)"
    if not valor_str[0].isdigit():
        return "INICIO_INVALIDO", "Debe empezar con dígito 0-9"
    segmentos = re.split(r'[/-]', valor_str)
    for seg in segmentos:
        if seg and not seg.isdigit():
            return "SEGMENTO_NO_NUMERICO", "Todos los segmentos deben ser números"
    return "VÁLIDO", valor_str

def extraer_cantidad_hilos(especificacion: str) -> int:
    match = RE_ESPEC_HILOS.search(str(especificacion))
    return int(match.group(1)) if match else 0

def validar_nombre_antiguo_fibra(nombre_ant, tipo_colocacion: str) -> bool:
    if pd.isna(nombre_ant) or str(nombre_ant).strip() == "":
        return True
    nombre_ant_str = str(nombre_ant).strip().upper()
    patron = RE_NOMBRE_ANT_UM if tipo_colocacion == "Ultima Milla" else RE_NOMBRE_ANT_NO_UM
    return bool(patron.match(nombre_ant_str))

def validar_nombre_fibra(nombre: str, tipo_colocacion: str) -> list:
    errores = []
    nombre_str = str(nombre).strip().upper()
    tipo_str = str(tipo_colocacion).strip()
    errores_whitespace = detectar_errores_whitespace(nombre, "NOMBRE")
    for tipo_err, valor_orig in errores_whitespace:
        errores.append(f"NOMBRE: {tipo_err} - valor: '{valor_orig}'")
    if tipo_str in ("Troncal-Acceso", "Troncal-Transporte"):
        if "-" in nombre_str:
            if "SUB" not in nombre_str:
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (en {tipo_str} el guion solo se permite con 'SUB')")
            elif not RE_NOMBRE_FIBRA_TRONCAL.match(nombre_str.replace("-SUB", "")):
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (formato inválido con SUB en {tipo_str})")
        elif "P" in nombre_str:
            if not RE_NOMBRE_FIBRA_TRONCAL.match(nombre_str):
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (en {tipo_str} debe contener solo números, P y / cuando hay P)")
            if "/" in nombre_str and "P" not in nombre_str:
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][8]} (en {tipo_str})")
            if nombre_str == "/" or nombre_str.replace("/", "") == "":
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][9]} (en {tipo_str})")
        else:
            if not RE_NOMBRE_FIBRA_TRONCAL_SIMPLE.match(nombre_str):
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (en {tipo_str} debe contener solo números cuando no hay P)")
    else:
        if "P" in nombre_str:
            if not RE_NOMBRE_FIBRA_UM_CON_P.match(nombre_str):
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][7]} (nombre con P)")
            if "/" in nombre_str and "P" not in nombre_str:
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][8]}")
            if nombre_str == "/" or nombre_str.replace("/", "").replace("-", "") == "":
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][9]}")
        else:
            if not RE_NOMBRE_FIBRA_UM_SIN_P.match(nombre_str):
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][7]} (nombre sin P)")
    return errores

def obtener_nombre_base_fibra(nombre_completo: str) -> str:
    if "-" not in nombre_completo:
        return nombre_completo
    partes = nombre_completo.split("-")
    if len(partes) == 2 and partes[1].isdigit():
        return nombre_completo
    if len(partes) == 3 and all(p.isdigit() for p in partes[1:]):
        return f"{partes[0]}-{partes[1]}"
    match = RE_NOMBRE_FIBRA_BASE.match(nombre_completo)
    return match.group(1) if match else nombre_completo

def validar_nombre_empalme(nombre: str) -> list:
    errores = []
    nombre_str = str(nombre).strip().upper()
    errores_whitespace = detectar_errores_whitespace(nombre, "NOMBRE")
    for tipo_err, valor_orig in errores_whitespace:
        errores.append(f"NOMBRE: {tipo_err} - valor: '{valor_orig}'")
    if " " in nombre_str:
        errores.append(f"{JERARQUIA_EMPALME[3]['errores'][1]}")
    if "-SUB" in nombre_str:
        if not RE_NOMBRE_EMPALME_SUB.match(nombre_str):
            errores.append(f"{JERARQUIA_EMPALME[3]['errores'][3]}: Formato inválido para empalme con SUB")
    else:
        if not RE_NOMBRE_EMPALME_VALIDO.match(nombre_str):
            caracteres_invalidos = re.sub(r"(ET|NAP|EA|OB)[0-9_\-P]", "", nombre_str)
            unicos = "".join(sorted(set(caracteres_invalidos)))
            if unicos:
                errores.append(f"{JERARQUIA_EMPALME[3]['errores'][3]}: '{unicos}'")
    if re.search(r'([A-Z])\1', nombre_str):
        letras_repetidas = {m.group(0) for m in re.finditer(r'([A-Z])\1', nombre_str)}
        errores.append(f"{JERARQUIA_EMPALME[3]['errores'][2]}: {', '.join(letras_repetidas)}")
    if nombre_str.startswith("ET") and "-" in nombre_str and "-SUB" not in nombre_str:
        errores.append(f"{JERARQUIA_EMPALME[3]['errores'][0]}")
    return errores

def validar_especificacion_empalme(nombre: str, especificacion: str) -> list:
    errores = []
    nombre_upper = nombre.upper()
    especificacion_str = str(especificacion).strip()
    errores_whitespace = detectar_errores_whitespace(especificacion, "ESPECIFICACION")
    for tipo_err, valor_orig in errores_whitespace:
        errores.append(f"ESPECIFICACION: {tipo_err} - valor: '{valor_orig}'")
    ESPEC_VALIDAS = {
        "Tipo 0", "Tipo 1", "Tipo 2", "Empalme Cliente",
        "Nap Sotano 12", "Nap Sotano 24", "Nap Sotano 36",
        "Nap Sotano 48", "Nap Sotano 96", "Nap Sotano 144",
    }
    if especificacion_str not in ESPEC_VALIDAS:
        if especificacion_str.startswith("Nap Sotano "):
            numero_str = especificacion_str.split("Nap Sotano ")[1]
            if not numero_str.isdigit():
                errores.append(f"{JERARQUIA_EMPALME[8]['errores'][1]} (actual: {especificacion_str})")
        else:
            errores.append(f"{JERARQUIA_EMPALME[8]['errores'][1]} (actual: {especificacion_str})")
    if nombre_upper.startswith("ET"):
        if especificacion_str not in {"Tipo 0", "Tipo 1", "Tipo 2", "Empalme Cliente"} and not especificacion_str.startswith("Nap Sotano"):
            errores.append(f"{JERARQUIA_EMPALME[8]['errores'][0]} (actual: {especificacion_str})")
    elif nombre_upper.startswith("NAP"):
        if not especificacion_str.startswith("Nap Sotano"):
            errores.append(f"ESPECIFICACION NAP ERRONEA (actual: {especificacion_str}, esperado: 'Nap Sotano X')")
    elif nombre_upper.startswith("EA"):
        if especificacion_str != "Tipo 2":
            errores.append(f"{JERARQUIA_EMPALME[8]['errores'][1]} (actual: {especificacion_str})")
    return errores

def determinar_tipo_empalme_esperado(especificacion_fibra: str) -> str:
    esp = especificacion_fibra.lower()
    if esp.startswith("ductado 144"):
        return "Tipo 0"
    if esp.startswith("ductado 96"):
        return "Tipo 1"
    if any(esp.startswith(f"ductado {n}") for n in ("12", "24", "36", "48")):
        return "Tipo 2"
    if "empalme cliente" in esp:
        return "Empalme Cliente"
    if "nap sotano" in esp:
        return "Nap Sotano"
    return "Desconocido"

def normalizar_nombre_empresa(nombre: str) -> str:
    if not nombre:
        return ""
    nombre_upper = nombre.upper().strip()
    sufijos = [
        r'\s+S\.?A\.?S\.?', r'\s+S\.?A\.?', r'\s+L\.?T\.?D\.?A\.?',
        r'\s+E\.?U\.?', r'\s+N\.?I\.?T\.?\s*\d+', r'\s+&?\s*CIE\.?',
        r'\s+LTDA\.?', r'\s+LTD\.?', r'\s+INC\.?', r'\s+CORP\.?',
        r'\s+S\.?C\.?S\.?', r'\s+E\.?I\.?R\.?L\.?',
    ]
    for sufijo in sufijos:
        nombre_upper = re.sub(sufijo, '', nombre_upper)
    nombre_upper = re.sub(r'[.,]', '', nombre_upper)
    nombre_upper = re.sub(r'\s+', ' ', nombre_upper).strip()
    return nombre_upper

def cargar_om_desde_macro(ruta_macro: Path, hoja: str = "OM"):
    ids_om, exc_ed = set(), set()
    wb = None
    try:
        wb = openpyxl.load_workbook(ruta_macro, data_only=True, read_only=True, keep_vba=True)
        if hoja not in wb.sheetnames:
            return ids_om, exc_ed
        ws = wb[hoja]
        idx_id = idx_elem = idx_tb = None
        for idx, cell in enumerate(ws[1], start=1):
            v = str(cell.value).strip().lower() if cell.value else ""
            if v == "id" and idx_id is None:
                idx_id = idx
            elif v in ("elemento", "element") and idx_elem is None:
                idx_elem = idx
            elif v == "tb" and idx_tb is None:
                idx_tb = idx
        for row in ws.iter_rows(min_row=2, values_only=True):
            if idx_id:
                val = row[idx_id - 1]
                if val:
                    ids_om.add(normalizar_id(val))
            if idx_elem and idx_tb:
                elem = row[idx_elem - 1]
                tb = row[idx_tb - 1]
                if elem and tb and str(tb).strip().upper() == "ED1":
                    exc_ed.add(str(elem).strip())
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"  [WARN] Error cargando OM: {e}")
    finally:
        if wb:
            try:
                wb.close()
            except:
                pass
    return ids_om, exc_ed

def cargar_responsables_por_cable(ruta: Path, hoja: str, columna_clave: str = "cable acceso"):
    try:
        df = pd.read_excel(ruta, sheet_name=hoja, engine="openpyxl", dtype=str)
    except (FileNotFoundError, ValueError):
        return {}
    df.columns = [str(c).strip().lower() for c in df.columns]
    clave_norm = columna_clave.lower()
    try:
        idx_clave = df.columns.get_loc(clave_norm)
        idx_resp = df.columns.get_loc("responsable red externa")
    except KeyError:
        return {}
    responsables = {}
    for _, row in df.iterrows():
        if clave_norm == "cable acceso":
            clave = str(row.iloc[idx_clave]).split("/")[0].strip() if pd.notna(row.iloc[idx_clave]) else ""
        elif clave_norm == "id edifico":
            clave = normalizar_id(row.iloc[idx_clave])
        else:
            clave = str(row.iloc[idx_clave]).strip() if pd.notna(row.iloc[idx_clave]) else ""
        if clave:
            responsables[clave] = str(row.iloc[idx_resp]).strip() if pd.notna(row.iloc[idx_resp]) else ""
    return responsables

def cargar_especificaciones_fibra(ruta_archivo: Path, ids_om: set, modo: str, incluir_colocacion: bool = False):
    especificaciones = {}
    colocacion_por_fibra = {} if incluir_colocacion else None
    wb = None
    try:
        wb = openpyxl.load_workbook(ruta_archivo, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        col = extraer_indices_columnas(ws)
        cols = COLUMNAS_MASIVO["FIBRA"]
        idx_id        = col.get(cols["id"], -1)
        idx_estado    = col.get(cols["estado"], -1)
        idx_nombre    = col.get(cols["nombre"], -1)
        idx_espec     = col.get(cols["especificacion"], -1)
        idx_colocacion= col.get(cols["colocacion"], -1) if incluir_colocacion else -1
        ids_om_set = set(ids_om)
        for fila in ws.iter_rows(min_row=2, values_only=True):
            if idx_id == -1 or idx_estado == -1:
                continue
            idv = normalizar_id(fila[idx_id])
            estado_raw = fila[idx_estado]
            if not idv or idv in ids_om_set or estado_raw in (None, ""):
                continue
            if idx_nombre == -1:
                continue
            nombre = str(fila[idx_nombre] or "").strip()
            base = nombre.split("/")[0].strip()
            esp_raw = str(fila[idx_espec] or "").strip() if idx_espec != -1 else ""
            if base:
                especificaciones[base] = esp_raw.lower()
                if incluir_colocacion and idx_colocacion != -1:
                    colocacion_por_fibra[base] = str(fila[idx_colocacion] or "").strip()
    except Exception as e:
        print(f"  [WARN] Error cargando especificaciones fibra: {e}")
    finally:
        if wb:
            try:
                wb.close()
            except:
                pass
    return (especificaciones, colocacion_por_fibra) if incluir_colocacion else especificaciones

def cargar_validaciones_odfs(ruta_odfs: Path):
    odfs_data = {}
    wb = None
    try:
        wb = openpyxl.load_workbook(ruta_odfs, data_only=True, read_only=True)
        if "ODFS" not in wb.sheetnames:
            return odfs_data
        ws = wb["ODFS"]
        headers = {str(cell.value).strip().lower(): idx for idx, cell in enumerate(ws[1], start=1) if cell.value}
        columnas_requeridas = ('validacion', 'troncal', 'ids_distribucion')
        if not all(col in headers for col in columnas_requeridas):
            return odfs_data
        idx_validacion     = headers['validacion']
        idx_troncal        = headers['troncal']
        idx_ids_distribucion = headers['ids_distribucion']
        for row in ws.iter_rows(min_row=2, values_only=True):
            validacion = str(row[idx_validacion]).strip().upper() if row[idx_validacion] else ""
            if validacion != "OK":
                continue
            troncal = str(row[idx_troncal]).strip() if row[idx_troncal] else ""
            ids_raw = str(row[idx_ids_distribucion]).strip() if row[idx_ids_distribucion] else ""
            if not troncal or not ids_raw:
                continue
            ids_list = [i.strip() for i in ids_raw.replace('"', '').split('|') if i.strip()]
            if ids_list:
                if troncal in odfs_data:
                    odfs_data[troncal].update(ids_list)
                else:
                    odfs_data[troncal] = set(ids_list)
    except Exception as e:
        print(f"  [WARN] Error cargando ODFS: {e}")
    finally:
        if wb:
            try:
                wb.close()
            except:
                pass
    return odfs_data

def cargar_nombres_edificios(ruta_ed: Path, modo: str):
    nombres_edificios = {}
    nombres_normalizados = defaultdict(set)
    wb = None
    try:
        print(f"  [DEBUG] Cargando edificios desde: {ruta_ed}")
        wb = openpyxl.load_workbook(ruta_ed, data_only=True, read_only=True)
        print(f"  [DEBUG] Hojas disponibles: {wb.sheetnames}")
        nombre_hoja = None
        for h in wb.sheetnames:
            h_lower = h.lower()
            if any(kw in h_lower for kw in ("edif", "building", "export", "output")):
                nombre_hoja = h
                break
        if not nombre_hoja:
            nombre_hoja = wb.sheetnames[0]
        print(f"  [DEBUG] Usando hoja: '{nombre_hoja}'")
        ws = wb[nombre_hoja]
        col = extraer_indices_columnas(ws)
        print(f"  [DEBUG] Columnas encontradas: {list(col.keys())}")
        cols_ed = COLUMNAS_MASIVO["EDIFICIO"]
        idx_id = -1
        idx_nombre = -1
        posibles_id = [cols_ed["id"].lower(), "id", "idtexto", "id texto", "building.id"]
        for posible in posibles_id:
            if posible in col:
                idx_id = col[posible]
                print(f"  [DEBUG] ID encontrado en columna: '{posible}' (idx={idx_id})")
                break
        posibles_nombre = [cols_ed["nombre"].lower(), "name", "nombre", "building name"]
        for posible in posibles_nombre:
            if posible in col:
                idx_nombre = col[posible]
                print(f"  [DEBUG] NOMBRE encontrado en columna: '{posible}' (idx={idx_nombre})")
                break
        if idx_id == -1 or idx_nombre == -1:
            print(f"  [WARN] No se encontraron columnas ID o NOMBRE. ID: {idx_id != -1}, NOMBRE: {idx_nombre != -1}")
            return nombres_edificios, nombres_normalizados
        count_filas = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if idx_id < len(row):
                idv = normalizar_id(row[idx_id])
            else:
                idv = ""
            if not idv:
                continue
            if idx_nombre < len(row):
                nombre = str(row[idx_nombre] or "").strip()
            else:
                nombre = ""
            if nombre:
                count_filas += 1
                nombre_upper = nombre.upper()
                nombres_edificios[idv] = nombre_upper
                nombres_normalizados[normalizar_nombre_empresa(nombre_upper)].add(idv)
        print(f"  [DEBUG] Edificios cargados: {count_filas}")
    except Exception as e:
        print(f"  [ERROR] cargar_nombres_edificios: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if wb:
            try:
                wb.close()
            except:
                pass
    return nombres_edificios, nombres_normalizados

def validar_nombre_cliente_contra_edificio(
    nombre_cliente, building_id, physical_status, comentarios,
    nombres_edificios: dict, nombres_normalizados: dict,
    contador_clientes_por_edificio: dict = None
):
    condicion_validar = False
    ps_upper = str(physical_status).strip().upper() if physical_status is not None else ""
    if ps_upper in ("IN SERVICE", "IS"):
        condicion_validar = True
    if comentarios and pd.notna(comentarios):
        if "ADICION" in str(comentarios).strip().upper():
            condicion_validar = True
    if not condicion_validar:
        return True, None, None, None
    if pd.isna(nombre_cliente) or str(nombre_cliente).strip() == "":
        return False, "", "NOMBRE CLIENTE VACIO", None
    nombre_original = str(nombre_cliente).strip()
    errores_whitespace = detectar_errores_whitespace(nombre_original, "NOMBRE CLIENTE")
    if errores_whitespace:
        msgs = [f"{tipo} - valor: '{valor}'" for tipo, valor in errores_whitespace]
        return False, nombre_original, f"NOMBRE CLIENTE: {' | '.join(msgs)}", None
    valido_sufijos, msg_sufijos, nombre_corregido = validar_sufijos_nombre_empresa(nombre_original)
    if not valido_sufijos:
        return False, nombre_original, msg_sufijos, nombre_corregido
    formato_valido, msg_formato, _ = validar_formato_nombre_cliente(nombre_original)
    if not formato_valido:
        return False, nombre_original, msg_formato, None
    skip_building_match = False
    building_name = None
    if building_id and building_id in nombres_edificios:
        building_name = nombres_edificios[building_id]
        building_norm = normalizar_nombre_empresa(building_name)
        if any(building_norm.startswith(p) or building_name.upper().startswith(p) for p in PREFIJOS_EXCLUIR_VALIDACION):
            skip_building_match = True
    if not skip_building_match and (not building_id or building_id not in nombres_edificios):
        cliente_norm = normalizar_nombre_empresa(nombre_original.upper())
        if cliente_norm in nombres_normalizados:
            edificios_encontrados = list(nombres_normalizados[cliente_norm])
            if edificios_encontrados:
                building_name = nombres_edificios.get(edificios_encontrados[0], "")
                if building_name:
                    building_norm = normalizar_nombre_empresa(building_name)
                    if any(building_norm.startswith(p) or building_name.upper().startswith(p) for p in PREFIJOS_EXCLUIR_VALIDACION):
                        skip_building_match = True
    if not skip_building_match and building_name:
        cliente_norm  = normalizar_nombre_empresa(nombre_original.upper())
        edificio_norm = normalizar_nombre_empresa(building_name.upper())
        if contador_clientes_por_edificio is not None:
            nombre_cliente_upper = nombre_original.upper()
            edificios_con_este_cliente = contador_clientes_por_edificio.get(nombre_cliente_upper, set())
            if len(edificios_con_este_cliente) > UMBRAL_IGNORAR_NOMBRE_CLIENTE:
                return True, None, None, None
        if cliente_norm in edificio_norm or edificio_norm in cliente_norm or cliente_norm == edificio_norm:
            return True, None, None, None
        error_msg = f"{JERARQUIA_PORT[16]['errores'][4]} | Posible esperado: '{building_name}'"
        return False, nombre_original, error_msg, building_name
    return True, None, None, None

class OrganizadorErrores:
    def __init__(self):
        self._errores_por_id: dict = defaultdict(list)
    def agregar_error(self, id_elemento: str, nombre_entidad: str, valores_error):
        prioridad, campo = self._resolver_prioridad_campo(nombre_entidad, valores_error)
        self._errores_por_id[id_elemento].append((prioridad, campo, nombre_entidad, valores_error))
    def _resolver_prioridad_campo(self, entidad: str, valores_error) -> tuple:
        mensaje = self._extraer_mensaje(valores_error)
        jerarquia = self._seleccionar_jerarquia(entidad, mensaje)
        if jerarquia is None:
            return (999, "Otros")
        mensaje_upper = mensaje.upper()
        for prioridad, config in jerarquia.items():
            for patron in config["errores"]:
                if patron.upper() in mensaje_upper:
                    return (prioridad, config["campo"])
        return (999, "Otros")
    @staticmethod
    def _seleccionar_jerarquia(entidad: str, mensaje: str):
        mapa = {
            "FIBRA":      JERARQUIA_FIBRA,
            "EMPALME":    JERARQUIA_EMPALME,
            "EDIFICIO":   JERARQUIA_EDIFICIO,
            "PORT":       JERARQUIA_PORT,
            "SHELF":      JERARQUIA_SHELF,
            "BAY":        JERARQUIA_BAY,
        }
        if entidad in mapa:
            return mapa[entidad]
        if entidad == "MERGE FI-EM":
            return JERARQUIA_FIBRA if "MERGE" in mensaje else JERARQUIA_EMPALME
        if "ESTRUCTURA" in entidad:
            return JERARQUIA_ESTRUCTURA
        return None
    @staticmethod
    def _extraer_mensaje(valores_error) -> str:
        if isinstance(valores_error, (list, tuple)):
            if len(valores_error) > 3:
                return str(valores_error[3])
            if len(valores_error) == 1:
                return str(valores_error[0])
            return str(valores_error)
        return str(valores_error)
    def obtener_errores_organizados(self) -> list:
        errores_finales = []
        for idv in sorted(self._errores_por_id.keys()):
            errores_id = self._errores_por_id[idv]
            errores_por_prioridad = defaultdict(list)
            for prio, campo, entidad, valores in errores_id:
                errores_por_prioridad[prio].append((campo, entidad, valores))
            mejor_prioridad = min(errores_por_prioridad.keys())
            mejores_errores = errores_por_prioridad[mejor_prioridad]
            agrupados = defaultdict(list)
            for campo, entidad, valores in mejores_errores:
                agrupados[(entidad, campo)].append(valores)
            for (entidad, _), lista_valores in agrupados.items():
                if len(lista_valores) == 1:
                    errores_finales.append((entidad, lista_valores[0]))
                else:
                    errores_finales.append((entidad, self._combinar_errores(lista_valores)))
        return errores_finales
    @staticmethod
    def _combinar_errores(lista_valores) -> tuple:
        id_val    = lista_valores[0][0]
        value_val = lista_valores[0][1]
        owner_val = lista_valores[0][2]
        mensajes = {vals[3] for vals in lista_valores if len(vals) > 3}
        return (id_val, value_val, owner_val, " | ".join(sorted(mensajes)))

# ----------------------------------------------------------------------
# Procesamiento de fibras (sin cambios)
# ----------------------------------------------------------------------
def procesar_fibras(ws_fi, col_fi: dict, cols_fi: dict, ids_om: set, responsables: dict, modo: str, odfs_data: dict, organizador: OrganizadorErrores):
    errores = []
    especificaciones = {}
    fibras_por_cable = defaultdict(list)
    fibras_por_cable_municipio = defaultdict(list)
    fibras_procesadas_nodos = set()
    idx_id          = col_fi.get(cols_fi["id"], -1)
    idx_nombre      = col_fi.get(cols_fi["nombre"], -1)
    idx_colocacion  = col_fi.get(cols_fi["colocacion"], -1)
    idx_especificacion = col_fi.get(cols_fi["especificacion"], -1)
    idx_propietario = col_fi.get(cols_fi["propietario"], -1)
    idx_estado      = col_fi.get(cols_fi["estado"], -1)
    idx_nombre_ant  = col_fi.get(cols_fi.get("nombre_ant", -1), -1)
    idx_nodo_a      = col_fi.get(cols_fi.get("nodo_a", -1), -1)
    idx_nodo_b      = col_fi.get(cols_fi.get("nodo_b", -1), -1)
    idx_municipio_a = col_fi.get(cols_fi.get("municipio_a", -1), -1)
    idx_municipio_b = col_fi.get(cols_fi.get("municipio_b", -1), -1)
    ids_om_set = set(ids_om)
    filas_procesadas = 0
    def registrar_error(idv, nombre, propietario, mensaje):
        err_tuple = (idv, nombre, propietario, mensaje)
        errores.append(err_tuple)
        organizador.agregar_error(idv, "FIBRA", err_tuple)
    for fila in ws_fi.iter_rows(min_row=2, values_only=True):
        filas_procesadas += 1
        if idx_id == -1 or idx_estado == -1:
            continue
        idv = normalizar_id(fila[idx_id])
        estado_raw = fila[idx_estado]
        if not idv or idv in ids_om_set or estado_raw in (None, ""):
            continue
        if idx_nombre == -1:
            continue
        nombre = str(fila[idx_nombre] or "").strip()
        base_raw = nombre.split("/")[0].strip()
        base = obtener_nombre_base_fibra(base_raw)
        tipo = str(fila[idx_colocacion] or "").strip() if idx_colocacion != -1 else ""
        esp_raw = str(fila[idx_especificacion] or "").strip() if idx_especificacion != -1 else ""
        esp = esp_raw.lower()
        prop = str(fila[idx_propietario] or "").strip() if idx_propietario != -1 else ""
        est = traducir_estado_construccion(estado_raw, modo)
        nombre_ant = fila[idx_nombre_ant] if idx_nombre_ant != -1 else ""
        if base:
            especificaciones[base] = esp
        errores_fila = []
        validar_valor_con_whitespace(nombre, "NOMBRE", errores_fila)
        validar_valor_con_whitespace(tipo, "COLOCACION", errores_fila)
        validar_valor_con_whitespace(esp_raw, "ESPECIFICACION", errores_fila)
        validar_valor_con_whitespace(prop, "PROPIETARIO", errores_fila)
        validar_valor_con_whitespace(est, "ESTADO", errores_fila)
        validar_valor_con_whitespace(nombre_ant, "NOMBRE_ANTIGUO", errores_fila)
        errores_fila.extend(validar_nombre_fibra(nombre, tipo))
        if "-" in nombre and tipo != "Ultima Milla" and not nombre.upper().endswith("-SUB"):
            errores_fila.append(f"{JERARQUIA_FIBRA[3]['errores'][6]} (actual: {tipo})")
        tipo_normalizado = tipo.lower().replace(" ", "-")
        if "-" not in nombre and tipo_normalizado not in ("troncal-acceso", "troncal-transporte"):
            if tipo_normalizado == "ultima-milla":
                pass
            else:
                errores_fila.append(f"{JERARQUIA_FIBRA[3]['errores'][5]} (actual: {tipo})")
        if "doble" in esp:
            errores_fila.append(f"{JERARQUIA_FIBRA[9]['errores'][0]} (actual: {esp})")
        if prop != PROPIETARIO_VALIDO:
            errores_fila.append(f"{JERARQUIA_FIBRA[11]['errores'][0]} (actual: {prop})")
        if est not in ("En servicio", "Plan Abandono", "Plan Uso Futuro"):
            errores_fila.append(f"{JERARQUIA_FIBRA[10]['errores'][0]} (actual: {est})")
        if tipo == "Ultima Milla":
            hilos = extraer_cantidad_hilos(esp_raw)
            if hilos > 48:
                errores_fila.append(f"{JERARQUIA_FIBRA[9]['errores'][2]} (actual: {hilos} Hilos)")
        if tipo == "Ultima Milla" and nombre_ant:
            if not validar_nombre_antiguo_fibra(nombre_ant, tipo):
                errores_fila.append(f"{JERARQUIA_FIBRA[4]['errores'][0]} (actual: {nombre_ant})")
            if "-" in nombre and "-" not in str(nombre_ant) and not nombre.upper().endswith("-SUB"):
                errores_fila.append(f"{JERARQUIA_FIBRA[4]['errores'][1]} (NOMBRE: {nombre}, NOMBRE_ANT: {nombre_ant})")
        owner = responsables.get(base, "-")
        for msg in errores_fila:
            registrar_error(idv, nombre, owner, msg)
        if "-" not in nombre and tipo in ("Troncal-Acceso", "Troncal-Transporte") and base:
            fibras_por_cable[base].append({
                'id': idv, 'nombre': nombre, 'especificacion': esp,
                'estado': est, 'tipo_colocacion': tipo, 'base_raw': base_raw,
            })
        if idx_municipio_a != -1 and idx_municipio_b != -1:
            muni_a = normalizar_id(fila[idx_municipio_a]) if idx_municipio_a != -1 else ""
            muni_b = normalizar_id(fila[idx_municipio_b]) if idx_municipio_b != -1 else ""
            match_num = RE_NOMBRE_FIBRA_NUMERO.match(nombre)
            cable_num = match_num.group(1) if match_num else None
            if cable_num:
                fibras_por_cable_municipio[cable_num].append({
                    'id': idv, 'nombre': nombre,
                    'municipio_a': muni_a, 'municipio_b': muni_b, 'tipo_colocacion': tipo,
                })
        if odfs_data and idv not in fibras_procesadas_nodos:
            fibras_procesadas_nodos.add(idv)
            match_base = RE_NOMBRE_FIBRA_BASE.match(nombre)
            if match_base:
                base_nombre = match_base.group(1)
                cable_odfs = ids_permitidos = None
                if base_nombre in odfs_data:
                    cable_odfs, ids_permitidos = base_nombre, odfs_data[base_nombre]
                if not cable_odfs and "-" in base_nombre:
                    partes = base_nombre.split("-")
                    if partes[0].isdigit() and partes[0] in odfs_data:
                        cable_odfs, ids_permitidos = partes[0], odfs_data[partes[0]]
                if not cable_odfs:
                    for troncal in odfs_data:
                        if base_nombre.startswith(troncal) or troncal.startswith(base_nombre.split("-")[0]):
                            cable_odfs, ids_permitidos = troncal, odfs_data[troncal]
                            break
                if cable_odfs and ids_permitidos:
                    nodo_a = normalizar_id(fila[idx_nodo_a]) if idx_nodo_a != -1 else ""
                    nodo_b = normalizar_id(fila[idx_nodo_b]) if idx_nodo_b != -1 else ""
                    if len(ids_permitidos) == 1:
                        id_req = next(iter(ids_permitidos))
                        if nodo_a != id_req or nodo_b != id_req:
                            msg = f"{JERARQUIA_FIBRA[7]['errores'][0]} | CORREGIR SEGÚN ODFS: A='{id_req}', B='{id_req}' | ACTUAL A='{nodo_a}' B='{nodo_b}'"
                            registrar_error(idv, nombre, prop or "-", msg)
                    else:
                        invalidos = []
                        if nodo_a and nodo_a not in ids_permitidos:
                            invalidos.append(f"A:'{nodo_a}'")
                        if nodo_b and nodo_b not in ids_permitidos:
                            invalidos.append(f"B:'{nodo_b}'")
                        if invalidos:
                            msg = f"{JERARQUIA_FIBRA[7]['errores'][1]} | Nodos no coinciden con ODFS"
                            registrar_error(idv, nombre, prop or "-", msg)
    for cable_base, fibras in fibras_por_cable.items():
        if len(fibras) <= 1:
            continue
        especs = [f['especificacion'] for f in fibras]
        cont_espec = Counter(especs)
        if len(cont_espec) > 1:
            esp_mayor = cont_espec.most_common(1)[0][0]
            for fib in fibras:
                if fib['especificacion'] != esp_mayor:
                    owner = responsables.get(cable_base, "-")
                    msg = (f"{JERARQUIA_FIBRA[999]['errores'][1]} | Cable '{cable_base}' tiene mayoría '{esp_mayor}' | "
                           f"MAYORÍA: '{esp_mayor}' , ACTUAL: '{fib['especificacion']}'")
                    registrar_error(fib['id'], fib['nombre'], owner, msg)
        estados = [f['estado'] for f in fibras]
        cont_est = Counter(estados)
        if len(cont_est) > 1:
            est_mayor = cont_est.most_common(1)[0][0]
            for fib in fibras:
                if fib['estado'] != est_mayor:
                    owner = responsables.get(cable_base, "-")
                    msg = (f"{JERARQUIA_FIBRA[10]['errores'][0]} | "
                           f"MAYORÍA: '{est_mayor}' , ACTUAL: '{fib['estado']}'")
                    registrar_error(fib['id'], fib['nombre'], owner, msg)
    municipios_predominantes = {}
    for cable_num, fibras in fibras_por_cable_municipio.items():
        if any(f['tipo_colocacion'] in ("Troncal-Acceso", "Troncal-Transporte") for f in fibras):
            continue
        todos_munis = [m for f in fibras for m in (f['municipio_a'], f['municipio_b']) if m]
        if todos_munis:
            cnt = Counter(todos_munis)
            mayor = cnt.most_common(1)[0][0]
            municipios_predominantes[cable_num] = (mayor, (cnt[mayor] / len(todos_munis)) * 100)
    for cable_num, fibras in fibras_por_cable_municipio.items():
        if cable_num not in municipios_predominantes:
            continue
        pred, pct = municipios_predominantes[cable_num]
        for fib in fibras:
            if fib['tipo_colocacion'] in ("Troncal-Acceso", "Troncal-Transporte"):
                continue
            owner = responsables.get(cable_num, "-")
            if fib['municipio_a'] == "":
                registrar_error(fib['id'], fib['nombre'], owner,
                    f"{JERARQUIA_FIBRA[5]['errores'][0]} | SUGERIDO: '{pred}' ({pct:.1f}%)")
            elif fib['municipio_a'] != pred:
                registrar_error(fib['id'], fib['nombre'], owner,
                    f"{JERARQUIA_FIBRA[5]['errores'][1]} | Municipio A '{fib['municipio_a']}' no coincide con predominante '{pred}' ({pct:.1f}%)")
            if fib['municipio_b'] == "":
                registrar_error(fib['id'], fib['nombre'], owner,
                    f"{JERARQUIA_FIBRA[6]['errores'][0]} | SUGERIDO: '{pred}' ({pct:.1f}%)")
            elif fib['municipio_b'] != pred:
                registrar_error(fib['id'], fib['nombre'], owner,
                    f"{JERARQUIA_FIBRA[6]['errores'][1]} | Municipio B '{fib['municipio_b']}' no coincide con predominante '{pred}' ({pct:.1f}%)")
    return errores, especificaciones, filas_procesadas

# ----------------------------------------------------------------------
# Procesamiento de empalmes (sin cambios)
# ----------------------------------------------------------------------
def procesar_empalmes(ws_em, col_em: dict, cols_em: dict, ids_om: set, responsables: dict, modo: str, organizador: OrganizadorErrores):
    errores = []
    datos_em = []
    grupos_especificacion = defaultdict(list)
    nombres_vistos = defaultdict(list)
    idx_id           = col_em.get(cols_em["id"], -1)
    idx_nombre       = col_em.get(cols_em["nombre"], -1)
    idx_especificacion = col_em.get(cols_em["especificacion"], -1)
    idx_propietario  = col_em.get(cols_em["propietario"], -1)
    idx_estado       = col_em.get(cols_em["estado"], -1)
    ids_om_set = set(ids_om)
    filas_procesadas = 0
    def registrar_error(idv, nombre, propietario, mensaje):
        err_tuple = (idv, nombre, propietario, mensaje)
        errores.append(err_tuple)
        organizador.agregar_error(idv, "EMPALME", err_tuple)
    for fila in ws_em.iter_rows(min_row=2, values_only=True):
        filas_procesadas += 1
        if idx_id == -1 or idx_estado == -1:
            continue
        idv = normalizar_id(fila[idx_id])
        estado_raw = fila[idx_estado]
        if not idv or idv in ids_om_set or estado_raw in (None, ""):
            continue
        if idx_nombre == -1:
            continue
        nombre = str(fila[idx_nombre] or "").strip()
        corto = nombre[3:] if nombre.upper().startswith("NAP") else (nombre[2:] if len(nombre) > 2 else nombre)
        esp   = str(fila[idx_especificacion] or "").strip() if idx_especificacion != -1 else ""
        prop  = str(fila[idx_propietario] or "").strip() if idx_propietario != -1 else ""
        est   = traducir_estado_construccion(estado_raw, modo)
        grupo = ""
        if (nombre.startswith("ET") or nombre.startswith("NAP")) and "_" in nombre:
            if "-SUB" in nombre.upper():
                match = re.match(r'^(ET|NAP)(\d+-SUB)', nombre.upper())
                grupo = f"{match.group(1)}{match.group(2)}" if match else ""
            else:
                grupo = nombre.split("_")[0]
        datos_em.append((idv, nombre, esp, prop, est, grupo, corto))
        nombres_vistos[nombre].append(idv)
        if grupo:
            agrupado = "Nap Sotano" if nombre.upper().startswith("NAP") and esp.startswith("Nap Sotano") else esp
            if agrupado in {"Tipo 0", "Tipo 1", "Tipo 2", "Empalme Cliente", "Nap Sotano"}:
                grupos_especificacion[grupo].append(agrupado)
        errores_fila = []
        validar_valor_con_whitespace(nombre, "NOMBRE", errores_fila)
        validar_valor_con_whitespace(esp, "ESPECIFICACION", errores_fila)
        validar_valor_con_whitespace(prop, "PROPIETARIO", errores_fila)
        validar_valor_con_whitespace(est, "ESTADO", errores_fila)
        errores_fila.extend(validar_nombre_empalme(nombre))
        errores_fila.extend(validar_especificacion_empalme(nombre, esp))
        if prop != PROPIETARIO_VALIDO:
            errores_fila.append(f"{JERARQUIA_EMPALME[4]['errores'][0]} (actual: {prop})")
        if est not in ("En servicio", "Plan Abandono", "Plan Uso Futuro"):
            errores_fila.append(f"{JERARQUIA_EMPALME[6]['errores'][0]} (actual: {est})")
        if "-SUB" not in nombre.upper():
            if not RE_NOMBRE_EMPALME_VALIDO.match(nombre.upper()):
                base_match = RE_NOMBRE_EMPALME_BASE.match(nombre.upper())
                nombre_corregido = base_match.group(0) if base_match else ""
                if nombre_corregido and nombre_corregido != nombre.upper():
                    errores_fila.append(f"{JERARQUIA_EMPALME[3]['errores'][4]} (usar: {nombre_corregido})")
        owner = responsables.get(corto, "-")
        for msg in errores_fila:
            registrar_error(idv, nombre, owner, msg)
    duplicados_reportados = set()
    for nombre, ids in nombres_vistos.items():
        if len(ids) > 1 and not ((nombre.startswith("ET") or nombre.startswith("NAP")) and "_" in nombre):
            ids_str = " , ".join(ids)
            corto_dup = nombre[3:] if nombre.upper().startswith("NAP") else (nombre[2:] if len(nombre) > 2 else nombre)
            owner = responsables.get(corto_dup, "-")
            err_tuple = (ids_str, nombre, owner, f"{JERARQUIA_EMPALME[1]['errores'][1]})")
            errores.append(err_tuple)
            organizador.agregar_error(ids_str, "EMPALME", err_tuple)
            duplicados_reportados.add(nombre)
    mayor_por_grupo = {g: Counter(v).most_common(1)[0][0] for g, v in grupos_especificacion.items() if v}
    for (idv, nombre, esp, prop, est, grupo, corto) in datos_em:
        if nombre in duplicados_reportados or not grupo or not mayor_por_grupo.get(grupo):
            continue
        esp_mayor = mayor_por_grupo[grupo]
        if grupo.startswith("NAP"):
            es_nap_actual = esp.startswith("Nap Sotano")
            es_nap_mayor  = esp_mayor == "Nap Sotano"
            if es_nap_actual != es_nap_mayor:
                owner = responsables.get(corto, "-")
                msg = f"{JERARQUIA_EMPALME[8]['errores'][2]} (para {grupo}, esperado: '{esp_mayor}', actual: '{esp}')"
                err_tuple = (idv, nombre, owner, msg)
                errores.append(err_tuple)
                organizador.agregar_error(idv, "EMPALME", err_tuple)
        elif esp in {"Tipo 0", "Tipo 1", "Tipo 2", "Empalme Cliente"} and esp != esp_mayor:
            owner = responsables.get(corto, "-")
            msg = f"{JERARQUIA_EMPALME[8]['errores'][2]} (esperado: '{esp_mayor}', actual: '{esp}')"
            err_tuple = (idv, nombre, owner, msg)
            errores.append(err_tuple)
            organizador.agregar_error(idv, "EMPALME", err_tuple)
    return errores, datos_em, filas_procesadas

def validar_consecutividad_empalmes(datos_em: list, ids_om: set, responsables: dict):
    errores = []
    grupos_sin_sub = defaultdict(list)
    grupos_con_sub = defaultdict(list)
    duplicados_reportados = set()
    ids_om_set = set(ids_om)
    for idv, nombre, esp, prop, est, grupo, corto in datos_em:
        if idv in ids_om_set or not idv:
            continue
        if not (nombre.startswith("ET") or nombre.startswith("NAP")) or "_" not in nombre:
            continue
        tiene_sub = "SUB" in nombre.upper()
        match = re.match(r'^(ET|NAP)(\d+)_(\d+)$', nombre.upper())
        if not match:
            continue
        prefijo_base  = match.group(1)
        numero_cable  = match.group(2)
        prefijo       = f"{prefijo_base}{numero_cable}"
        try:
            consecutivo = int(match.group(3))
        except ValueError:
            continue
        registro = {'id': idv, 'nombre': nombre, 'consecutivo': consecutivo, 'corto': corto}
        if tiene_sub:
            grupos_con_sub[prefijo].append(registro)
        else:
            grupos_sin_sub[prefijo].append(registro)
    def _analizar_grupo(grupos_dict):
        for prefijo, lista in grupos_dict.items():
            lista.sort(key=lambda x: x['consecutivo'])
            vistos = {}
            for em in lista:
                consec = em['consecutivo']
                clave_nombre = f"{prefijo}_{consec}"
                if consec in vistos:
                    if clave_nombre not in duplicados_reportados:
                        owner = responsables.get(em['corto'], "-")
                        ids_dup = f"{vistos[consec]['id']} , {em['id']}"
                        errores.append((ids_dup, vistos[consec]['nombre'], owner,
                                        f"{JERARQUIA_EMPALME[1]['errores'][1]})"))
                        duplicados_reportados.add(clave_nombre)
                else:
                    vistos[consec] = em
            if not duplicados_reportados:
                for i in range(len(lista) - 1):
                    actual   = lista[i]
                    siguiente = lista[i + 1]
                    diferencia = siguiente['consecutivo'] - actual['consecutivo']
                    owner = responsables.get(actual['corto'], "-")
                    if diferencia < 0:
                        errores.append((siguiente['id'], siguiente['nombre'], owner,
                                        f"{JERARQUIA_EMPALME[1]['errores'][0]} _{actual['consecutivo']} y un duplicado {prefijo}_{siguiente['consecutivo']})"))
                    elif diferencia == 2:
                        faltante = actual['consecutivo'] + 1
                        errores.append((actual['id'], actual['nombre'], owner,
                                        f"{JERARQUIA_EMPALME[1]['errores'][0]} _{faltante} después del _{actual['consecutivo']})"))
                    elif diferencia > 2:
                        inicio = actual['consecutivo'] + 1
                        fin    = siguiente['consecutivo'] - 1
                        if fin - inicio >= 1:
                            faltantes = list(range(inicio, fin + 1))
                            faltantes_str = ", ".join([f"_{n}" for n in faltantes]) if len(faltantes) <= 3 else f"_{inicio} al _{fin}"
                            errores.append((actual['id'], actual['nombre'], owner,
                                            f"{JERARQUIA_EMPALME[1]['errores'][0]} _{actual['consecutivo']} y el _{siguiente['consecutivo']}. Faltan: {faltantes_str})"))
    _analizar_grupo(grupos_sin_sub)
    _analizar_grupo(grupos_con_sub)
    return errores

def validar_consistencia_tipo_cable_empalmes(datos_em: list, ids_om: set, responsables: dict):
    errores = []
    cables_principales = defaultdict(list)
    sub_cables         = defaultdict(list)
    ESPEC_VALIDAS = {
        "Tipo 0", "Tipo 1", "Tipo 2", "Empalme Cliente",
        "Nap Sotano 12", "Nap Sotano 24", "Nap Sotano 36",
        "Nap Sotano 48", "Nap Sotano 96", "Nap Sotano 144",
    }
    ids_om_set = set(ids_om)
    for idv, nombre, esp, prop, est, grupo, corto in datos_em:
        if idv in ids_om_set or not idv:
            continue
        if not (nombre.startswith("ET") or nombre.startswith("NAP")):
            continue
        if esp not in ESPEC_VALIDAS and not esp.startswith("Nap Sotano "):
            continue
        match = re.match(r'^(ET|NAP)(\d+)', nombre.upper())
        if not match:
            continue
        prefijo_base = match.group(1)
        cable_base   = f"{prefijo_base}{match.group(2)}"
        registro     = {'id': idv, 'nombre': nombre, 'especificacion': esp, 'corto': corto}
        if "-SUB" in nombre.upper():
            match_sub = re.match(r'^(ET|NAP)(\d+)-SUB', nombre.upper())
            if match_sub:
                sub_cable = f"{match_sub.group(1)}{match_sub.group(2)}-SUB"
                sub_cables[sub_cable].append(registro)
        else:
            cables_principales[cable_base].append(registro)
    def _evaluar_inconsistencia(cable, empalmes):
        if len(empalmes) <= 1:
            return
        especs = [em['especificacion'] for em in empalmes]
        if "NAP" in cable:
            agrupado = ["Nap Sotano" if x.startswith("Nap Sotano") else x for x in especs]
            cont = Counter(agrupado)
        else:
            cont = Counter(especs)
        if len(cont) == 1:
            return
        esp_mayor = cont.most_common(1)[0][0]
        for em in empalmes:
            numero_match = re.search(r'(ET|NAP)(\d+)', cable)
            numero_cable = numero_match.group(2) if numero_match else cable[2:]
            owner = responsables.get(numero_cable, "-")
            if "NAP" in cable:
                es_nap_actual = em['especificacion'].startswith("Nap Sotano")
                es_nap_mayor  = esp_mayor == "Nap Sotano"
                if es_nap_actual != es_nap_mayor:
                    msg = (f"{JERARQUIA_EMPALME[999]['errores'][0]} | Cable {cable} tiene mayoría '{esp_mayor}' | "
                           f"MAYORÍA: '{esp_mayor}' , ACTUAL: '{em['especificacion']}'")
                    errores.append((em['id'], em['nombre'], owner, msg))
            elif em['especificacion'] != esp_mayor:
                msg = (f"{JERARQUIA_EMPALME[999]['errores'][0]} | Cable {cable} tiene mayoría '{esp_mayor}' | "
                       f"MAYORÍA: '{esp_mayor}' , ACTUAL: '{em['especificacion']}'")
                errores.append((em['id'], em['nombre'], owner, msg))
    for cable, empalmes in cables_principales.items():
        _evaluar_inconsistencia(cable, empalmes)
    for sub_cable, empalmes in sub_cables.items():
        _evaluar_inconsistencia(sub_cable, empalmes)
    return errores

# ----------------------------------------------------------------------
# PROCESAR PORTS (CORREGIDO Y CON EXCLUSIÓN DE CENTRALES EN SECUENCIA)
# ----------------------------------------------------------------------
def procesar_ports(ruta_ports: Path, ids_om: set, modo: str, nombres_edificios: dict, nombres_normalizados: dict):
    errores_port_raw = []
    errores_shelf = []
    errores_bay = []
    cliente_por_edificio = defaultdict(set)
    wb = None
    try:
        wb = openpyxl.load_workbook(ruta_ports, data_only=True, read_only=True)
        ws = wb["Port"] if "Port" in wb.sheetnames else wb.worksheets[0]
        col = extraer_indices_columnas(ws)
        cols_port = COLUMNAS_MASIVO["PORT"]
        cols_shelf = COLUMNAS_MASIVO["SHELF"]
        cols_bay = COLUMNAS_MASIVO["BAY"]
        def idx(cols, campo):
            return col.get(cols[campo].lower(), -1)
        idx_port_id = idx(cols_port, "id")
        idx_physical_status = idx(cols_port, "physical_status")
        idx_tipo_senal = idx(cols_port, "tipo_senal")
        idx_description = idx(cols_port, "description")
        idx_unidad_ancho_banda = idx(cols_port, "unidad_ancho_banda")
        idx_type = idx(cols_port, "type")
        idx_connector_type = idx(cols_port, "connector_type")
        idx_service_level = idx(cols_port, "service_level")
        idx_contratista = idx(cols_port, "contratista")
        idx_nombre_cliente = idx(cols_port, "nombre_cliente")
        idx_comentarios = idx(cols_port, "comentarios")
        idx_direccion_cliente_id = idx(cols_port, "direccion_cliente_id")
        idx_proyecto_red = idx(cols_port, "proyecto_red")
        idx_anillo = idx(cols_port, "anillo")
        idx_cambio_sm = idx(cols_port, "cambio_service_manager")
        idx_id_servicio = idx(cols_port, "id_servicio")
        idx_equipo_central = idx(cols_port, "equipo_central")
        idx_puerto_central = idx(cols_port, "puerto_central")
        idx_tipo_servicio = idx(cols_port, "tipo_servicio")
        idx_shelf_id = idx(cols_shelf, "id")
        idx_shelf_desc = idx(cols_shelf, "description")
        idx_shelf_estado = idx(cols_shelf, "estado")
        idx_shelf_tipo = idx(cols_shelf, "tipo")
        idx_shelf_fabricante = idx(cols_shelf, "fabricante")
        idx_shelf_tecnologia = idx(cols_shelf, "tecnologia")
        idx_shelf_ip = idx(cols_shelf, "direccion_ip")
        idx_shelf_nivel = idx(cols_shelf, "nivel_bastidor")
        idx_shelf_contratista = idx(cols_shelf, "contratista")
        idx_shelf_migracion_olt = idx(cols_shelf, "migracion_olt")
        idx_shelf_id_bay = idx(cols_shelf, "id_bay")
        idx_bay_id = idx(cols_bay, "id")
        idx_bay_number = idx(cols_bay, "number")
        idx_bay_description = idx(cols_bay, "description")
        idx_bay_estado = idx(cols_bay, "estado")
        idx_bay_numero_salon = idx(cols_bay, "numero_salon")
        idx_bay_numero_fila = idx(cols_bay, "numero_fila")
        idx_bay_contratista = idx(cols_bay, "contratista")
        idx_bay_id_edificio = idx(cols_bay, "id_edificio")
        idx_bay_cambio_sm = idx(cols_bay, "cambio_service_manager")
        col_name_building = next((k for k in col if "name_building" in k or "name building" in k), None)
        idx_name_building = col.get(col_name_building, -1) if col_name_building else -1
        col_building_id = next((k for k in col if "building.id" in k.lower() or "building id" in k.lower()), None)
        idx_building_id = col.get(col_building_id, -1) if col_building_id else -1
        ids_om_set = set(ids_om)
        filas_procesadas = 0
        
        # Primera pasada: recolectar datos sin agrupar aún
        datos_para_validacion = []
        for fila in ws.iter_rows(min_row=2, values_only=True):
            filas_procesadas += 1
            if idx_port_id != -1:
                idv = normalizar_id(fila[idx_port_id]) if idx_port_id < len(fila) else None
                if idv and idv not in ids_om_set:
                    building_id_val = normalizar_id(fila[idx_building_id]) if idx_building_id != -1 and idx_building_id < len(fila) else None
                    nombre_cliente_raw = fila[idx_nombre_cliente] if idx_nombre_cliente != -1 and idx_nombre_cliente < len(fila) else None
                    if nombre_cliente_raw is not None:
                        nombre_cliente_raw = limpiar_valor_reporte(str(nombre_cliente_raw))
                    if nombre_cliente_raw and building_id_val:
                        nombre_cliente_str = str(nombre_cliente_raw).strip().upper()
                        cliente_por_edificio[nombre_cliente_str].add(building_id_val)
                    datos_para_validacion.append((fila, idv))
        
        # Diccionario para almacenar relación puerto -> shelf y para validación de secuencia
        port_to_shelf = {}
        shelf_ports_info = defaultdict(list)  # para validación de secuencia (solo clientes)
        shelves_centrales = set()             # IDs de shelves que son centrales
        
        for fila, idv in datos_para_validacion:
            desc_shelf_raw = fila[idx_shelf_desc] if idx_shelf_desc != -1 and idx_shelf_desc < len(fila) else ""
            if desc_shelf_raw is None:
                desc_shelf_raw = ""
            else:
                desc_shelf_raw = str(desc_shelf_raw).strip()
            tipo_shelf_raw = fila[idx_shelf_tipo] if idx_shelf_tipo != -1 and idx_shelf_tipo < len(fila) else ""
            if tipo_shelf_raw is None:
                tipo_shelf_raw = ""
            else:
                tipo_shelf_raw = str(tipo_shelf_raw).strip()
            type_raw = fila[idx_type] if idx_type != -1 and idx_type < len(fila) else None
            es_central = False
            tipo_central = None
            if desc_shelf_raw:
                desc_upper = desc_shelf_raw.upper()
                if "OT" in desc_upper:
                    es_central = True
                    tipo_central = "OT"
                elif "OI" in desc_upper:
                    es_central = True
                    tipo_central = "OI"
                elif "PP" in desc_upper:
                    es_central = True
                    tipo_central = "PP"
            if es_central:
                # Guardar el shelf como central para omitir validación de secuencia
                shelf_idv = normalizar_id(fila[idx_shelf_id]) if idx_shelf_id != -1 and idx_shelf_id < len(fila) else None
                if shelf_idv:
                    shelves_centrales.add(shelf_idv)
                type_val = str(type_raw).strip() if type_raw is not None else ""
                if type_val.upper() != "ODF":
                    msg = f"TYPE INVALIDO. Debe ser 'ODF' para central (actual: '{type_val}')"
                    errores_port_raw.append((idv, str(fila[idx_description] if idx_description != -1 and idx_description < len(fila) else "") or "Sin descripción", "", msg))
                esperado = None
                if tipo_central == "OT":
                    esperado = "ODF Troncal"
                elif tipo_central == "OI":
                    esperado = "ODF Interconexión"
                elif tipo_central == "PP":
                    esperado = "ODF Pacheo"
                if esperado:
                    if tipo_shelf_raw.upper() != esperado.upper():
                        msg = f"TIPO INVALIDO. Esperado: '{esperado}', actual: '{tipo_shelf_raw}'"
                        shelf_id = normalizar_id(fila[idx_shelf_id]) if idx_shelf_id != -1 and idx_shelf_id < len(fila) else idv
                        errores_shelf.append((shelf_id, f"SHELF: {desc_shelf_raw or 'Sin descripción'}", "", msg))
                continue  # Los puertos de central no pasan por validaciones de cliente
            
            # --- Cliente: recolectar para validación de secuencia y errores normales ---
            if idx_shelf_id != -1:
                shelf_idv = normalizar_id(fila[idx_shelf_id]) if idx_shelf_id < len(fila) else None
                if shelf_idv:
                    port_to_shelf[idv] = shelf_idv
                    # Solo si tenemos el número de puerto desde Description
                    description_port = str(fila[idx_description] if idx_description != -1 and idx_description < len(fila) else "").strip()
                    if description_port and description_port.isdigit():
                        port_num = int(description_port)
                        ps_raw = fila[idx_physical_status] if idx_physical_status != -1 and idx_physical_status < len(fila) else None
                        if ps_raw is not None:
                            ps_upper = str(ps_raw).strip().upper()
                            ps = "IN SERVICE" if ps_upper == "IS" else ps_raw
                        else:
                            ps = None
                        status_str = str(ps).strip().upper() if ps else ""
                        shelf_ports_info[shelf_idv].append({
                            'port_num': port_num,
                            'status': status_str,
                            'idv': idv,
                            'description': description_port
                        })
            
            # Validaciones normales de cliente (generan errores en errores_port_raw)
            name_building = str(fila[idx_name_building] if idx_name_building != -1 and idx_name_building < len(fila) else "").strip()
            building_id = normalizar_id(fila[idx_building_id]) if idx_building_id != -1 and idx_building_id < len(fila) else ""
            tipo_senal = str(fila[idx_tipo_senal] if idx_tipo_senal != -1 and idx_tipo_senal < len(fila) else "").strip()
            description = str(fila[idx_description] if idx_description != -1 and idx_description < len(fila) else "").strip()
            ps_raw = fila[idx_physical_status] if idx_physical_status != -1 and idx_physical_status < len(fila) else None
            if ps_raw is not None:
                ps_upper = str(ps_raw).strip().upper()
                ps = "IN SERVICE" if ps_upper == "IS" else ps_raw
            else:
                ps = None
            nombre_cliente_raw = fila[idx_nombre_cliente] if idx_nombre_cliente != -1 and idx_nombre_cliente < len(fila) else None
            if nombre_cliente_raw is not None:
                nombre_cliente_raw = limpiar_valor_reporte(str(nombre_cliente_raw))
            comentarios_raw = fila[idx_comentarios] if idx_comentarios != -1 and idx_comentarios < len(fila) else None
            direccion_cliente_raw = fila[idx_direccion_cliente_id] if idx_direccion_cliente_id != -1 and idx_direccion_cliente_id < len(fila) else None
            proyecto_red_raw = fila[idx_proyecto_red] if idx_proyecto_red != -1 and idx_proyecto_red < len(fila) else None
            anillo_raw = fila[idx_anillo] if idx_anillo != -1 and idx_anillo < len(fila) else None
            cambio_sm_raw = fila[idx_cambio_sm] if idx_cambio_sm != -1 and idx_cambio_sm < len(fila) else None
            id_servicio_raw = fila[idx_id_servicio] if idx_id_servicio != -1 and idx_id_servicio < len(fila) else None
            unidad_ancho_banda_raw = fila[idx_unidad_ancho_banda] if idx_unidad_ancho_banda != -1 and idx_unidad_ancho_banda < len(fila) else None
            connector_type_raw = fila[idx_connector_type] if idx_connector_type != -1 and idx_connector_type < len(fila) else None
            service_level_raw = fila[idx_service_level] if idx_service_level != -1 and idx_service_level < len(fila) else None
            equipo_central_raw = fila[idx_equipo_central] if idx_equipo_central != -1 and idx_equipo_central < len(fila) else None
            puerto_central_raw = fila[idx_puerto_central] if idx_puerto_central != -1 and idx_puerto_central < len(fila) else None
            mensajes = []
            error_equipo_central = None
            error_puerto_central = None
            
            validar_valor_con_whitespace(tipo_senal, "TIPO SEÑAL", mensajes)
            validar_valor_con_whitespace(description, "DESCRIPTION", mensajes)
            validar_valor_con_whitespace(ps, "PHYSICAL STATUS", mensajes)
            validar_valor_con_whitespace(type_raw, "TYPE", mensajes)
            validar_valor_con_whitespace(connector_type_raw, "CONNECTOR TYPE", mensajes)
            validar_valor_con_whitespace(proyecto_red_raw, "PROYECTO DE RED", mensajes)
            validar_valor_con_whitespace(anillo_raw, "ANILLO", mensajes)
            validar_valor_con_whitespace(cambio_sm_raw, "CAMBIO SERVICE MANAGER", mensajes)
            validar_valor_con_whitespace(equipo_central_raw, "EQUIPO CENTRAL", mensajes)
            validar_valor_con_whitespace(puerto_central_raw, "PUERTO CENTRAL", mensajes)
            validar_valor_con_whitespace(name_building, "NAME BUILDING", mensajes)
            validar_valor_con_whitespace(building_id, "BUILDING ID", mensajes)
            validar_valor_con_whitespace(direccion_cliente_raw, "DIRECCION CLIENTE", mensajes)
            if nombre_cliente_raw is not None:
                errores_ws_nombre = detectar_errores_whitespace(nombre_cliente_raw, "NOMBRE CLIENTE")
                for tipo_err, valor_orig in errores_ws_nombre:
                    mensajes.append(f"NOMBRE CLIENTE: {tipo_err} - valor: '{valor_orig}'")
            
            if idx_equipo_central != -1:
                estado_ec, valor_ec = validar_equipo_central(equipo_central_raw)
                if estado_ec != "VÁLIDO":
                    valor_mostrar = str(equipo_central_raw).strip() if equipo_central_raw is not None else ""
                    error_equipo_central = f"EQUIPO CENTRAL ({valor_mostrar})"
            if idx_puerto_central != -1:
                estado_pc, valor_pc = validar_puerto_central(puerto_central_raw)
                if estado_pc != "VÁLIDO":
                    valor_mostrar = str(puerto_central_raw).strip() if puerto_central_raw is not None else ""
                    error_puerto_central = f"PUERTO CENTRAL ({valor_mostrar})"
            if idx_nombre_cliente != -1:
                es_valido, valor_actual, error_msg, _ = validar_nombre_cliente_contra_edificio(
                    nombre_cliente_raw, building_id, ps, comentarios_raw,
                    nombres_edificios, nombres_normalizados,
                    cliente_por_edificio
                )
                if not es_valido and error_msg is not None:
                    errores_port_raw.append((idv, valor_actual or "", "", error_msg))
                    continue
            if ps is not None and str(ps).strip().upper() == "IN SERVICE":
                dir_str = str(direccion_cliente_raw).strip() if direccion_cliente_raw is not None else ""
                if es_valor_cero(direccion_cliente_raw, dir_str):
                    mensajes.append(f"{JERARQUIA_PORT[17]['errores'][0]} = 0")
                proy_str = str(proyecto_red_raw).strip() if proyecto_red_raw is not None else ""
                if es_valor_cero(proyecto_red_raw, proy_str):
                    mensajes.append(f"{JERARQUIA_PORT[10]['errores'][0]} = 0")
                elif proyecto_red_raw is not None and proy_str:
                    proy_clean = re.sub(r'[\x00-\x1f\x7f]', '', proy_str).strip()
                    if not re.match(r'^[A-Z0-9]+$', proy_clean):
                        mensajes.append(f"{JERARQUIA_PORT[10]['errores'][0]}'{proy_str}'")
                    elif not any(proy_clean.startswith(p) for p in ("ME", "HU", "SD")):
                        mensajes.append(f"{JERARQUIA_PORT[10]['errores'][0]}: '{proy_str}'")
                if idx_anillo != -1 and not str(fila[idx_anillo] if idx_anillo < len(fila) else "").strip():
                    mensajes.append(f"{JERARQUIA_PORT[20]['errores'][0]}")
                if idx_cambio_sm != -1 and cambio_sm_raw is not None:
                    csm_str = str(cambio_sm_raw).strip()
                    csm_upper = csm_str.upper()
                    if csm_str:
                        if len(csm_upper) < 7:
                            mensajes.append(f"{JERARQUIA_PORT[25]['errores'][0]}: '{csm_str}'")
                        elif not any(csm_upper.startswith(p) for p in PREFIJOS_SERVICE_MANAGER):
                            mensajes.append(f"{JERARQUIA_PORT[25]['errores'][0]}: '{csm_str}'")
                        elif not csm_upper.replace("-", "").isalnum():
                            mensajes.append(f"{JERARQUIA_PORT[25]['errores'][0]}: '{csm_str}'")
                if idx_id_servicio != -1 and id_servicio_raw is not None:
                    id_serv_str = str(id_servicio_raw).strip()
                    id_serv_clean = re.sub(r'[\x00-\x1f\x7f]', '', id_serv_str).strip()
                    if id_serv_str and id_serv_str != "0":
                        if re.search(r'[\s\t]+', id_serv_clean):
                            mensajes.append(f"{JERARQUIA_PORT[26]['errores'][0]}: '{id_serv_str}'")
                        elif not re.match(r'^[A-Z0-9]+$', id_serv_clean):
                            mensajes.append(f"{JERARQUIA_PORT[26]['errores'][0]}: '{id_serv_str}'")
                        elif len(id_serv_clean) < 8:
                            mensajes.append(f"{JERARQUIA_PORT[26]['errores'][0]}: '{id_serv_str}'")
                        elif not id_serv_clean.startswith(PREFIJOS_ID_SERVICIO):
                            mensajes.append(f"{JERARQUIA_PORT[26]['errores'][0]}: '{id_serv_str}'")
            if error_equipo_central:
                mensajes.append(error_equipo_central)
            if error_puerto_central:
                mensajes.append(error_puerto_central)
            if tipo_senal != "Rx - Tx":
                mensajes.append(f"{JERARQUIA_PORT[2]['errores'][0]} ({'(VACÍO)' if not tipo_senal else f'(actual: {tipo_senal!r})'})")
            st, _ = validar_type_equipo_cliente(type_raw)
            if st != "VÁLIDO":
                if st == 'VACÍO':
                    mensajes.append(f"{JERARQUIA_PORT[5]['errores'][1]}")
                else:
                    mensajes.append(f"{JERARQUIA_PORT[5]['errores'][2]} (actual: '{type_raw}')")
            stc, _ = validar_connector_type(connector_type_raw)
            if stc != "VÁLIDO":
                if stc == 'VACÍO':
                    mensajes.append(f"{JERARQUIA_PORT[6]['errores'][1]}")
                else:
                    mensajes.append(f"{JERARQUIA_PORT[6]['errores'][2]} (actual: '{connector_type_raw}')")
            sts, _ = validar_service_level_fiber(service_level_raw)
            if sts != "VÁLIDO":
                if sts == 'VACÍO':
                    mensajes.append(f"{JERARQUIA_PORT[9]['errores'][1]}")
                else:
                    mensajes.append(f"{JERARQUIA_PORT[9]['errores'][2]} (actual: '{service_level_raw}')")
            if ps is not None:
                stp, _ = validar_physical_status(ps)
                if stp != "VÁLIDO":
                    if stp == 'VACÍO':
                        mensajes.append(f"{JERARQUIA_PORT[4]['errores'][1]}")
                    else:
                        mensajes.append(f"{JERARQUIA_PORT[4]['errores'][0]} (actual: '{ps}')")
            stu, _ = validar_unidad_ancho_banda(unidad_ancho_banda_raw)
            if stu != "VÁLIDO":
                if stu == 'VACÍO':
                    mensajes.append(f"{JERARQUIA_PORT[8]['errores'][1]}")
                else:
                    mensajes.append(f"{JERARQUIA_PORT[8]['errores'][0]} (actual: '{unidad_ancho_banda_raw}')")
            if mensajes:
                msg_final = " | ".join(mensajes)
                errores_port_raw.append((idv, description or "Sin descripción", "", msg_final))
            
            # Validaciones de shelf y bay (sin cambios)
            if idx_shelf_id != -1:
                shelf_idv = normalizar_id(fila[idx_shelf_id]) if idx_shelf_id < len(fila) else None
                if shelf_idv and shelf_idv not in ids_om_set:
                    desc_shelf = str(fila[idx_shelf_desc] if idx_shelf_desc != -1 and idx_shelf_desc < len(fila) else "").strip()
                    es_central_shelf = False
                    if desc_shelf and ("OT" in desc_shelf.upper() or "OI" in desc_shelf.upper() or "PP" in desc_shelf.upper()):
                        es_central_shelf = True
                    if es_central_shelf:
                        continue
                    description_shelf = str(fila[idx_shelf_desc] if idx_shelf_desc != -1 and idx_shelf_desc < len(fila) else "").strip()
                    estado_shelf = str(fila[idx_shelf_estado] if idx_shelf_estado != -1 and idx_shelf_estado < len(fila) else "").strip()
                    contratista_shelf = str(fila[idx_shelf_contratista] if idx_shelf_contratista != -1 and idx_shelf_contratista < len(fila) else "").strip()
                    tecnologia = str(fila[idx_shelf_tecnologia] if idx_shelf_tecnologia != -1 and idx_shelf_tecnologia < len(fila) else "").strip()
                    direccion_ip = str(fila[idx_shelf_ip] if idx_shelf_ip != -1 and idx_shelf_ip < len(fila) else "").strip()
                    nivel_bastidor = str(fila[idx_shelf_nivel] if idx_shelf_nivel != -1 and idx_shelf_nivel < len(fila) else "").strip()
                    migracion_olt = str(fila[idx_shelf_migracion_olt] if idx_shelf_migracion_olt != -1 and idx_shelf_migracion_olt < len(fila) else "").strip()
                    id_bay = normalizar_id(fila[idx_shelf_id_bay]) if idx_shelf_id_bay != -1 and idx_shelf_id_bay < len(fila) else ""
                    msgs_shelf = []
                    validar_valor_con_whitespace(description_shelf, "DESCRIPTION SHELF", msgs_shelf)
                    validar_valor_con_whitespace(estado_shelf, "ESTADO SHELF", msgs_shelf)
                    validar_valor_con_whitespace(tecnologia, "TECNOLOGIA", msgs_shelf)
                    validar_valor_con_whitespace(contratista_shelf, "CONTRATISTA SHELF", msgs_shelf)
                    validar_valor_con_whitespace(direccion_ip, "DIRECCION IP", msgs_shelf)
                    validar_valor_con_whitespace(nivel_bastidor, "NIVEL BASTIDOR", msgs_shelf)
                    validar_valor_con_whitespace(migracion_olt, "MIGRACION OLT", msgs_shelf)
                    validar_valor_con_whitespace(id_bay, "ID BAY", msgs_shelf)
                    if not description_shelf:
                        msgs_shelf.append(f"{JERARQUIA_SHELF[2]['errores'][0]}")
                    else:
                        if description_shelf != description_shelf.upper():
                            msgs_shelf.append("DESCRIPTION CON ESPACIOS")
                        if description_shelf.endswith(' '):
                            msgs_shelf.append("DESCRIPTION CON ESPACIOS")
                        msgs_shelf.extend(validar_descripcion_demarcador(description_shelf))
                    if not contratista_shelf:
                        msgs_shelf.append(f"{JERARQUIA_SHELF[9]['errores'][0]}")
                    if not estado_shelf:
                        msgs_shelf.append(f"{JERARQUIA_SHELF[3]['errores'][0]}")
                    elif estado_shelf.lower() not in ("is", "spare", "under construction", "retired"):
                        msgs_shelf.append(f"{JERARQUIA_SHELF[3]['errores'][0]} (actual: '{estado_shelf}')")
                    if not tecnologia:
                        msgs_shelf.append(f"{JERARQUIA_SHELF[6]['errores'][0]}")
                    if direccion_ip:
                        est_ip, _ = validar_ipv4(direccion_ip)
                        if est_ip != "VÁLIDO":
                            msgs_shelf.append(f"{JERARQUIA_SHELF[7]['errores'][0]} (actual: '{direccion_ip}')")
                    if nivel_bastidor:
                        try:
                            nivel = int(nivel_bastidor)
                            if not (0 <= nivel <= 50):
                                msgs_shelf.append(f"{JERARQUIA_SHELF[8]['errores'][0]} (actual: '{nivel_bastidor}')")
                        except ValueError:
                            msgs_shelf.append(f"{JERARQUIA_SHELF[8]['errores'][0]} (actual: '{nivel_bastidor}')")
                    if migracion_olt and migracion_olt.lower() not in ("yes", "no", "sí", "true", "false"):
                        msgs_shelf.append(f"{JERARQUIA_SHELF[12]['errores'][0]} (actual: '{migracion_olt}')")
                    if not id_bay:
                        msgs_shelf.append(f"{JERARQUIA_SHELF[13]['errores'][0]}")
                    if msgs_shelf:
                        unicos = list(dict.fromkeys(msgs_shelf))
                        error_msg = " | ".join(unicos)
                        errores_shelf.append((shelf_idv, f"SHELF: {description_shelf or 'Sin descripción'}", "", error_msg))
            
            if idx_bay_id != -1:
                bay_idv = normalizar_id(fila[idx_bay_id]) if idx_bay_id < len(fila) else None
                if bay_idv and bay_idv not in ids_om_set:
                    desc_shelf_bay = str(fila[idx_shelf_desc] if idx_shelf_desc != -1 and idx_shelf_desc < len(fila) else "").strip()
                    es_central_bay = False
                    if desc_shelf_bay and ("OT" in desc_shelf_bay.upper() or "OI" in desc_shelf_bay.upper() or "PP" in desc_shelf_bay.upper()):
                        es_central_bay = True
                    if es_central_bay:
                        continue
                    description_bay = str(fila[idx_bay_description] if idx_bay_description != -1 and idx_bay_description < len(fila) else "").strip()
                    number_bay = str(fila[idx_bay_number] if idx_bay_number != -1 and idx_bay_number < len(fila) else "").strip()
                    estado_bay = str(fila[idx_bay_estado] if idx_bay_estado != -1 and idx_bay_estado < len(fila) else "").strip()
                    numero_salon = str(fila[idx_bay_numero_salon] if idx_bay_numero_salon != -1 and idx_bay_numero_salon < len(fila) else "").strip()
                    numero_fila = str(fila[idx_bay_numero_fila] if idx_bay_numero_fila != -1 and idx_bay_numero_fila < len(fila) else "").strip()
                    contratista_bay = str(fila[idx_bay_contratista] if idx_bay_contratista != -1 and idx_bay_contratista < len(fila) else "").strip()
                    id_edificio = normalizar_id(fila[idx_bay_id_edificio]) if idx_bay_id_edificio != -1 and idx_bay_id_edificio < len(fila) else ""
                    msgs_bay = []
                    validar_valor_con_whitespace(description_bay, "DESCRIPTION BAY", msgs_bay)
                    validar_valor_con_whitespace(estado_bay, "ESTADO BAY", msgs_bay)
                    validar_valor_con_whitespace(contratista_bay, "CONTRATISTA BAY", msgs_bay)
                    validar_valor_con_whitespace(id_edificio, "ID EDIFICIO", msgs_bay)
                    if not description_bay:
                        msgs_bay.append(f"{JERARQUIA_BAY[3]['errores'][0]}")
                    msgs_bay.extend(validar_descripcion_bastidor(description_bay))
                    if not number_bay:
                        msgs_bay.append(f"{JERARQUIA_BAY[2]['errores'][0]}")
                    if not estado_bay:
                        msgs_bay.append(f"{JERARQUIA_BAY[4]['errores'][0]}")
                    if not numero_salon:
                        msgs_bay.append(f"{JERARQUIA_BAY[5]['errores'][0]}")
                    elif not RE_NUMERO_SALON.match(numero_salon):
                        msgs_bay.append(f"{JERARQUIA_BAY[5]['errores'][0]} (actual: '{numero_salon}')")
                    if not numero_fila:
                        msgs_bay.append(f"{JERARQUIA_BAY[6]['errores'][0]}")
                    elif numero_fila != "1":
                        msgs_bay.append(f"NUMERO FILA INVALIDO. ACTUAL: {numero_fila}")
                    if not contratista_bay:
                        msgs_bay.append(f"{JERARQUIA_BAY[7]['errores'][0]}")
                    if not id_edificio:
                        msgs_bay.append(f"{JERARQUIA_BAY[9]['errores'][0]}")
                    if msgs_bay:
                        unicos = list(dict.fromkeys(msgs_bay))
                        error_msg = " | ".join(unicos)
                        errores_bay.append((bay_idv, f"BAY: {description_bay or 'Sin descripción'}", "", error_msg))
        
        # --- Validación de secuencia (solo para shelves de cliente, NO para centrales) ---
        for shelf_id, ports in shelf_ports_info.items():
            if shelf_id in shelves_centrales:
                continue  # No aplicar regla de secuencia a shelves de central
            if not ports:
                continue
            ports_sorted = sorted(ports, key=lambda x: x['port_num'])
            last_in_service_idx = -1
            for idx, p in enumerate(ports_sorted):
                if p['status'] == "IN SERVICE":
                    last_in_service_idx = idx
            if last_in_service_idx == -1:
                p = ports_sorted[0]
                msg = "SECUENCIA: SIN_IN_SERVICE"
                errores_port_raw.append((p['idv'], p['description'], "", msg))
                continue
            for idx, p in enumerate(ports_sorted):
                if idx <= last_in_service_idx:
                    if p['status'] != "IN SERVICE":
                        last_port = ports_sorted[last_in_service_idx]['port_num']
                        msg = f"SECUENCIA: SPARE_ANTES_DE_IN_SERVICE (puerto {p['port_num']}, último {last_port})"
                        errores_port_raw.append((p['idv'], p['description'], "", msg))
                else:
                    if p['status'] == "IN SERVICE":
                        msg = f"SECUENCIA: IN_SERVICE_DESPUES_DE_SPARE (puerto {p['port_num']})"
                        errores_port_raw.append((p['idv'], p['description'], "", msg))
        
    except Exception as e:
        print(f"\n[ERROR en procesar_ports] {e}")
        import traceback
        traceback.print_exc()
    finally:
        if wb:
            try:
                wb.close()
            except:
                pass
    
    # --- AGRUPACIÓN DE ERRORES DE PUERTO POR SHELF Y MENSAJE ---
    grupos = defaultdict(lambda: {'ids': [], 'descriptions': [], 'owner': None, 'msg': None})
    for err in errores_port_raw:
        idv, desc, owner, msg = err
        shelf = port_to_shelf.get(idv, None)
        if shelf is None:
            key = (None, msg, idv)   # clave única para errores sin shelf
        else:
            key = (shelf, msg)
        grupos[key]['ids'].append(idv)
        grupos[key]['descriptions'].append(desc)
        if owner and not grupos[key]['owner']:
            grupos[key]['owner'] = owner
        grupos[key]['msg'] = msg
    
    errores_port_agrupados = []
    for key, data in grupos.items():
        ids_concat = " , ".join(data['ids'])
        desc_rep = data['descriptions'][0] if data['descriptions'] else ""
        errores_port_agrupados.append((ids_concat, desc_rep, data['owner'] or "", data['msg']))
    
    return (errores_port_agrupados, errores_shelf, errores_bay, filas_procesadas)

# ----------------------------------------------------------------------
# Procesamiento de edificios y estructuras (sin cambios)
# ----------------------------------------------------------------------
def procesar_edificios_y_estructuras(ruta_ed: Path, ids_om: set, exc_ed: set, responsables_edificio: dict, modo: str, organizador: OrganizadorErrores):
    errores_edificio = []
    errores_estructura = []
    filas_procesadas = 0
    wb = None
    try:
        print(f"  [DEBUG] Procesando edificios desde: {ruta_ed}")
        wb = openpyxl.load_workbook(ruta_ed, data_only=True, read_only=True)
        nombre_hoja = None
        for h in wb.sheetnames:
            h_lower = h.lower()
            if any(kw in h_lower for kw in ("edif", "building", "export", "output", "structure")):
                nombre_hoja = h
                break
        if not nombre_hoja:
            nombre_hoja = wb.sheetnames[0]
        print(f"  [DEBUG] Usando hoja para procesar: '{nombre_hoja}'")
        ws = wb[nombre_hoja]
        col = extraer_indices_columnas(ws)
        cols_ed = COLUMNAS_MASIVO["EDIFICIO"]
        cols_est = COLUMNAS_ESTRUCTURA
        idx_id = -1
        idx_nombre = -1
        idx_tipo = -1
        idx_categoria = -1
        posibles_id = [cols_ed["id"].lower(), "id", "idtexto", "id texto", "building.id"]
        for posible in posibles_id:
            if posible in col:
                idx_id = col[posible]
                print(f"  [DEBUG] ID encontrado en columna: '{posible}' (idx={idx_id})")
                break
        posibles_nombre = [cols_ed["nombre"].lower(), "name", "nombre", "building name"]
        for posible in posibles_nombre:
            if posible in col:
                idx_nombre = col[posible]
                print(f"  [DEBUG] NOMBRE encontrado en columna: '{posible}' (idx={idx_nombre})")
                break
        posibles_tipo = [cols_ed["tipo"].lower(), "type", "tipo", "building type"]
        for posible in posibles_tipo:
            if posible in col:
                idx_tipo = col[posible]
                break
        posibles_categoria = [cols_ed["categoria"].lower(), "categoria", "category"]
        for posible in posibles_categoria:
            if posible in col:
                idx_categoria = col[posible]
                break
        idx_id_est = -1
        idx_nombre_est = -1
        idx_floor_len = -1
        idx_floor_wid = -1
        idx_rooms_len = -1
        idx_rooms_wid = -1
        idx_floors = -1
        posibles_id_est = [cols_est["idtexto"].lower(), "idtexto", "id texto", "building.id", "id"]
        for posible in posibles_id_est:
            if posible in col:
                idx_id_est = col[posible]
                break
        posibles_nombre_est = [cols_est["nombre"].lower(), "name", "nombre", "building name"]
        for posible in posibles_nombre_est:
            if posible in col:
                idx_nombre_est = col[posible]
                break
        if "floor dimension (length)" in col:
            idx_floor_len = col["floor dimension (length)"]
        if "floor dimension (width)" in col:
            idx_floor_wid = col["floor dimension (width)"]
        if "number of rooms (length)" in col:
            idx_rooms_len = col["number of rooms (length)"]
        if "number of rooms (width)" in col:
            idx_rooms_wid = col["number of rooms (width)"]
        if "number of floors" in col:
            idx_floors = col["number of floors"]
        print(f"  [DEBUG] Índices EDIFICIO: id={idx_id}, nombre={idx_nombre}")
        ids_om_set = set(ids_om)
        datos_edificio = []
        grupos_tipo = defaultdict(list)
        count_edificios = 0
        count_estructuras = 0
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if idx_id != -1 and idx_id < len(row):
                idv = normalizar_id(row[idx_id])
                if idv and idv not in ids_om_set:
                    count_edificios += 1
                    if idx_nombre != -1 and idx_nombre < len(row):
                        nombre = str(row[idx_nombre] or "").strip()
                    else:
                        nombre = ""
                    base = nombre.split("/")[0].strip() if nombre else ""
                    errores_edificio_fila = []
                    validar_valor_con_whitespace(nombre, "NOMBRE", errores_edificio_fila)
                    if nombre:
                        valido_sufijos, msg_sufijos, _ = validar_sufijos_nombre_empresa(nombre)
                        if not valido_sufijos:
                            owner = responsables_edificio.get(idv, "RESPONSABLE NO ENCONTRADO")
                            err_tuple = (idv, nombre, owner, msg_sufijos)
                            errores_edificio.append(err_tuple)
                            organizador.agregar_error(idv, "EDIFICIO", err_tuple)
                    tipo = ""
                    if idx_tipo != -1 and idx_tipo < len(row):
                        tipo = str(row[idx_tipo] or "").strip()
                    cat = ""
                    if idx_categoria != -1 and idx_categoria < len(row):
                        cat = str(row[idx_categoria] or "").strip().upper()
                    datos_edificio.append((idv, nombre, tipo, cat, base))
                    if base:
                        grupos_tipo[base].append(tipo)
                    for err_msg in errores_edificio_fila:
                        err_tuple = (idv, nombre, "...", err_msg)
                        errores_edificio.append(err_tuple)
                        organizador.agregar_error(idv, "EDIFICIO", err_tuple)
            if idx_id_est != -1 and idx_id_est < len(row):
                idv_est = normalizar_id(row[idx_id_est])
                if idv_est and idv_est not in ids_om_set:
                    if idx_nombre_est != -1 and idx_nombre_est < len(row):
                        nombre_est = str(row[idx_nombre_est] or "").strip()
                    else:
                        nombre_est = ""
                    if nombre_est and nombre_est.upper().startswith("NODO"):
                        continue
                    count_estructuras += 1
                    def _val(idx_campo):
                        return normalizar_id(row[idx_campo]) if idx_campo != -1 and idx_campo < len(row) else ""
                    floor_len = _val(idx_floor_len)
                    floor_wid = _val(idx_floor_wid)
                    rooms_len = _val(idx_rooms_len)
                    rooms_wid = _val(idx_rooms_wid)
                    num_pisos = _val(idx_floors)
                    owner = responsables_edificio.get(idv_est, "RESPONSABLE NO ENCONTRADO")
                    errores_estructura_fila = []
                    validar_valor_con_whitespace(nombre_est, "NOMBRE", errores_estructura_fila)
                    validar_valor_con_whitespace(floor_len, "FLOOR LENGTH", errores_estructura_fila)
                    validar_valor_con_whitespace(floor_wid, "FLOOR WIDTH", errores_estructura_fila)
                    validar_valor_con_whitespace(rooms_len, "ROOMS LENGTH", errores_estructura_fila)
                    validar_valor_con_whitespace(rooms_wid, "ROOMS WIDTH", errores_estructura_fila)
                    validar_valor_con_whitespace(num_pisos, "NUMBER OF FLOORS", errores_estructura_fila)
                    for err_msg in errores_estructura_fila:
                        err_tuple = (idv_est, nombre_est, owner, err_msg)
                        errores_estructura.append(err_tuple)
                        organizador.agregar_error(idv_est, "ESTRUCTURA", err_tuple)
                    if num_pisos == "0" and floor_len == "0" and floor_wid == "0" and rooms_len == "0" and rooms_wid == "0":
                        detalles = f"NUM_OF:{num_pisos}|FLOOR:{floor_len},{floor_wid}|ROOMS:{rooms_len},{rooms_wid}"
                        err_msg = f"{JERARQUIA_ESTRUCTURA[1]['errores'][0]} ({detalles})"
                        err_tuple = (idv_est, nombre_est, owner, err_msg)
                        errores_estructura.append(err_tuple)
                        organizador.agregar_error(idv_est, "ESTRUCTURA", err_tuple)
                        continue
                    FLOOR_VALIDOS = {"700", "40000", "10000", "20000"}
                    ROOMS_VALIDOS = {"3", "0"}
                    regs = []
                    if floor_len and floor_len not in FLOOR_VALIDOS:
                        regs.append(f"{JERARQUIA_ESTRUCTURA[1]['errores'][1]} (actual: {floor_len})")
                    if floor_wid and floor_wid not in FLOOR_VALIDOS:
                        regs.append(f"{JERARQUIA_ESTRUCTURA[1]['errores'][1]} (actual: {floor_wid})")
                    if rooms_len and rooms_len not in ROOMS_VALIDOS:
                        regs.append(f"{JERARQUIA_ESTRUCTURA[1]['errores'][2]} (actual: {rooms_len})")
                    if rooms_wid and rooms_wid not in ROOMS_VALIDOS:
                        regs.append(f"{JERARQUIA_ESTRUCTURA[1]['errores'][2]} (actual: {rooms_wid})")
                    if regs:
                        err_tuple = (idv_est, nombre_est, owner, "; ".join(regs))
                        errores_estructura.append(err_tuple)
                        organizador.agregar_error(idv_est, "ESTRUCTURA", err_tuple)
        filas_procesadas = count_edificios
        print(f"  [DEBUG] Edificios procesados: {count_edificios}, Estructuras procesadas: {count_estructuras}")
        tipo_mayoría = {g: Counter(t).most_common(1)[0][0] for g, t in grupos_tipo.items() if t}
        for idv, nombre, tipo, cat, base in datos_edificio:
            msgs = []
            if cat and cat != "AAA":
                msgs.append(f"{JERARQUIA_EDIFICIO[5]['errores'][0]} (actual: {cat})")
            if base and tipo and base in tipo_mayoría and tipo != tipo_mayoría[base] and nombre not in exc_ed:
                msgs.append(f"{JERARQUIA_EDIFICIO[3]['errores'][0]} (esperado: '{tipo_mayoría[base]}', actual: '{tipo}')")
            if nombre:
                nombre_limpio = nombre.strip()
                if not (nombre_limpio.upper().startswith("C.C") or nombre_limpio.upper().startswith("NODO")):
                    if not re.fullmatch(r"[A-Za-z0-9 ñÑ()&.,\-]+", nombre_limpio):
                        invalidos = "".join(sorted(set(re.sub(r"[A-Za-z0-9 ñÑ()&.,\-]", "", nombre_limpio))))
                        if invalidos:
                            msgs.append(f"{JERARQUIA_EDIFICIO[2]['errores'][2]}: '{invalidos}'")
                    if nombre_limpio != nombre_limpio.upper():
                        msgs.append(f"{JERARQUIA_EDIFICIO[2]['errores'][1]} (usar: {nombre_limpio.upper()})")
            if msgs:
                err_tuple = (idv, nombre, "...", " | ".join(msgs))
                errores_edificio.append(err_tuple)
                organizador.agregar_error(idv, "EDIFICIO", err_tuple)
    except Exception as e:
        print(f"  [ERROR] procesar_edificios_y_estructuras: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if wb:
            try:
                wb.close()
            except:
                pass
    return errores_edificio, errores_estructura, filas_procesadas

def detectar_edificios_con_nombres_similares(ruta_archivo: Path):
    wb = None
    try:
        wb = openpyxl.load_workbook(ruta_archivo, data_only=True, read_only=True)
        ws = wb.active
    except FileNotFoundError:
        return []
    except Exception as e:
        print(f"  [WARN] Error detectando edificios similares: {e}")
        return []
    try:
        idx_nom = idx_id = None
        for idx, cell in enumerate(ws[1], start=1):
            v = str(cell.value).strip().lower() if cell.value else ""
            if v in ("nombre", "name"):
                idx_nom = idx
            if v in ("id", "idtexto", "id texto"):
                idx_id = idx
        if not idx_nom or not idx_id:
            return []
        conteo = defaultdict(int)
        ids_por_nombre = defaultdict(list)
        nombres_vistos = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            id_clean = normalizar_id(row[idx_id - 1])
            if not id_clean:
                continue
            nombre_raw = row[idx_nom - 1]
            if nombre_raw:
                n = str(nombre_raw).strip()
                conteo[n] += 1
                ids_por_nombre[n].append(id_clean)
                if n not in nombres_vistos:
                    nombres_vistos.append(n)
        def _preparar_para_comparacion(texto: str) -> str:
            t = texto.lower()
            t = re.sub(r'[.,;()]', '', t)
            t = re.sub(r'\s+', ' ', t).strip()
            palabras = t.split()
            resultado = []
            i = 0
            while i < len(palabras):
                if len(palabras[i]) == 1 and palabras[i].isalpha():
                    letras = []
                    while i < len(palabras) and len(palabras[i]) == 1 and palabras[i].isalpha():
                        letras.append(palabras[i])
                        i += 1
                    resultado.append(''.join(letras) if len(letras) > 1 else letras[0])
                else:
                    resultado.append(palabras[i])
                    i += 1
            return ' '.join(resultado)
        grupos_normalizados = defaultdict(list)
        for nombre in nombres_vistos:
            prep = _preparar_para_comparacion(nombre)
            if prep:
                grupos_normalizados[prep].append(nombre)
        def _excluir(nombre: str) -> bool:
            nl = nombre.lower()
            return "subred" in nl or "sede" in nl or nl.startswith("ant cm")
        def _puntuar(nombre: str) -> int:
            pts = 0
            if '.' not in nombre: pts += 1
            if '  ' not in nombre and nombre.count(' ') < 3: pts += 1
            if nombre.strip() == nombre: pts += 1
            if not any(nombre.upper().startswith(p) for p in ("C.C", "NODO")) and nombre == nombre.upper():
                pts += 1
            return pts
        salida = []
        for prep, grupo in grupos_normalizados.items():
            grupo_filtrado = [n for n in grupo if not _excluir(n)]
            if len(grupo_filtrado) < 2:
                continue
            max_count = max(conteo[n] for n in grupo_filtrado)
            mejores = [n for n in grupo_filtrado if conteo[n] == max_count]
            mejor = max(mejores, key=_puntuar)
            for nombre in grupo_filtrado:
                if nombre == mejor:
                    continue
                ids = ", ".join(ids_por_nombre[nombre])
                salida.append((ids, nombre, "...", f"{JERARQUIA_EDIFICIO[2]['errores'][0]} ({mejor})"))
        return salida
    finally:
        if wb:
            try:
                wb.close()
            except:
                pass

def buscar_id_fibra_por_nombre(ws_fi, col_fi: dict, cols_fi: dict, ids_om: set, nombre_fibra: str):
    if ws_fi is None or col_fi is None:
        return None
    idx_id = col_fi.get(cols_fi["id"], -1)
    idx_nombre = col_fi.get(cols_fi["nombre"], -1)
    idx_estado = col_fi.get(cols_fi["estado"], -1)
    ids_om_set = set(ids_om)
    for fila in ws_fi.iter_rows(min_row=2, values_only=True):
        if idx_id == -1 or idx_estado == -1:
            continue
        idv = normalizar_id(fila[idx_id])
        if not idv or idv in ids_om_set or fila[idx_estado] in (None, ""):
            continue
        if idx_nombre != -1:
            nombre_raw = str(fila[idx_nombre] or "").strip()
            base = obtener_nombre_base_fibra(nombre_raw.split("/")[0].strip())
            if base == nombre_fibra:
                return idv
    return None

def validar_merge_fibra_empalme(datos_em: list, espec_fi: dict, responsables: dict, ids_om: set, espec_fi_respaldo: dict = None, tipo_colocacion_por_fibra: dict = None, ws_fi=None, col_fi: dict = None, cols_fi: dict = None):
    errores = []
    espec_combinado = {}
    if espec_fi_respaldo:
        espec_combinado.update(espec_fi_respaldo)
    espec_combinado.update(espec_fi)
    fibras_sin_sub = {}
    fibras_con_sub = {}
    for nombre_fi, esp_fi in espec_combinado.items():
        mc = re.match(r'(\d+)-SUB$', nombre_fi.upper())
        ms = re.match(r'(\d+)$', nombre_fi)
        fuente = 'principal' if nombre_fi in espec_fi else 'respaldo'
        info = {'nombre_completo': nombre_fi, 'especificacion_fibra': esp_fi,
                'tipo_esperado': determinar_tipo_empalme_esperado(esp_fi), 'fuente': fuente}
        if mc:
            fibras_con_sub[mc.group(1)] = info
        elif ms:
            fibras_sin_sub[ms.group(1)] = info
    ids_om_set = set(ids_om)
    emp_sin_sub = defaultdict(list)
    emp_con_sub = defaultdict(list)
    for id_em, nom_em, esp_em, *_, corto in datos_em:
        if id_em in ids_om_set or not id_em:
            continue
        if not (nom_em.startswith("ET") or nom_em.startswith("NAP")) or "_" not in nom_em:
            continue
        if "-SUB" in nom_em.upper():
            m = re.match(r'^(ET|NAP)(\d+)-SUB', nom_em.upper())
            if m:
                clave = f"{m.group(1)}{m.group(2)}"
                emp_con_sub[clave].append({'id': id_em, 'nombre': nom_em, 'especificacion_empalme': esp_em, 'corto': corto, 'tipo': m.group(1), 'codigo_base': m.group(2)})
        else:
            m = re.match(r'^(ET|NAP)(\d+)_', nom_em.upper())
            if m:
                clave = f"{m.group(1)}{m.group(2)}"
                emp_sin_sub[clave].append({'id': id_em, 'nombre': nom_em, 'especificacion_empalme': esp_em, 'corto': corto, 'tipo': m.group(1), 'codigo_base': m.group(2)})
    def _evaluar_grupo(empalmes, fib_info, es_sub: bool):
        if not empalmes or not fib_info:
            return
        tipo_emp = empalmes[0]['tipo']
        cod_base = empalmes[0]['codigo_base']
        tipos_act = [e['especificacion_empalme'] for e in empalmes]
        if tipo_emp == "NAP":
            agrupado = ["Nap Sotano" if x.startswith("Nap Sotano") else x for x in tipos_act]
            cont = Counter(agrupado)
        else:
            cont = Counter(tipos_act)
        if not tipos_act:
            return
        tipo_mayor, cnt_mayor = cont.most_common(1)[0]
        pct = (cnt_mayor / len(empalmes)) * 100
        tipo_esp = fib_info['tipo_esperado']
        if pct == 100 and tipo_mayor != tipo_esp and tipo_esp != 'Desconocido':
            nota = f" [FI de archivo {fib_info['fuente']}]" if fib_info['fuente'] == 'respaldo' else ""
            etiqueta = ('SUB-ANILLO' if tipo_emp == 'ET' else 'NAP SUB-ANILLO') if es_sub else ('CABLE PRINCIPAL' if tipo_emp == 'ET' else 'NAP PRINCIPAL')
            msg = (f"{JERARQUIA_FIBRA[999]['errores'][1]} ({etiqueta}) | "
                   f"FI:'{fib_info['nombre_completo']}' ({fib_info['especificacion_fibra']} → actual: '{tipo_esp}' | "
                   f"Todos los empalmes ({len(empalmes)}) tienen: '{tipo_mayor}' | "
                   f"{'PROBLEMA EN SUB-ANILLO' if es_sub else 'PROBLEMA EN FIBRA PRINCIPAL'}: Todos los empalmes concuerdan entre sí pero no con la fibra | "
                   f"RECOMENDACIÓN: Corregir especificación de fibra a '{tipo_mayor}'{nota}")
            id_fib = buscar_id_fibra_por_nombre(ws_fi, col_fi, cols_fi, ids_om, fib_info['nombre_completo']) if ws_fi else None
            ref_id = id_fib or empalmes[0]['id']
            etq = 'sub-anillo' if es_sub else 'cable principal'
            errores.append((ref_id, f"FI:{fib_info['nombre_completo']} ({etq} '{cod_base}')", responsables.get(cod_base, "-"), msg))
    for clave, empalmes in emp_sin_sub.items():
        _evaluar_grupo(empalmes, fibras_sin_sub.get(empalmes[0]['codigo_base']), es_sub=False)
    for clave, empalmes in emp_con_sub.items():
        _evaluar_grupo(empalmes, fibras_con_sub.get(empalmes[0]['codigo_base']), es_sub=True)
    return errores

def aplicar_formato_hoja_resultados(ws, total_filas: int):
    borde_fino = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.border = borde_fino
    limite_filas = min(total_filas + 1, ws.max_row)
    for row_idx in range(2, limite_filas + 1):
        for col_idx in range(1, 7):
            ws.cell(row=row_idx, column=col_idx).border = borde_fino
        if row_idx % 10000 == 0:
            gc.collect()
    for col_idx, col in enumerate(ws.columns, 1):
        max_len = max((len(str(cell.value)) for cell in list(col)[:500] if cell.value), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

def ejecutar_validacion_calidad(modo: str) -> int:
    start_time = time.time()
    wb = None
    try:
        wb = openpyxl.load_workbook(RUTA_MACRO, keep_vba=True)
    except Exception as e:
        print(f"ERROR: No se pudo abrir el archivo macro: {e}")
        return 0
    try:
        if "CD" in wb.sheetnames:
            wb.remove(wb["CD"])
        ws_salida = wb.create_sheet("CD")
        ws_salida.append(["ELEMENT", "ID", "VALUE", "OWNER", "ERROR", "COMMENTS"])
        organizador = OrganizadorErrores()
        ids_om, exc_ed = cargar_om_desde_macro(RUTA_MACRO)
        responsables = cargar_responsables_por_cable(RUTA_RESPONSABLES, "TD ACCES0s", "cable acceso")
        responsables_edificio = cargar_responsables_por_cable(RUTA_RESPONSABLES, "TD ACCES0s", "id edifico")
        odfs_data = cargar_validaciones_odfs(RUTA_ODFS)
        nombres_edificios, nombres_normalizados = cargar_nombres_edificios(RUTA_EDIFICIOS_BULK, modo)
        cols_fi = COLUMNAS_MASIVO["FIBRA"]
        cols_em = COLUMNAS_MASIVO["EMPALME"]
        contadores = {'fibra': 0, 'empalme': 0, 'edificio': 0, 'estructura': 0, 'ports': 0}
        errores_por_tipo = {'FIBRA': 0, 'EMPALME': 0, 'EDIFICIO': 0, 'ESTRUCTURA': 0, 'PORTS': 0, 'MERGE': 0}
        datos_em = []
        espec_fi = {}
        ws_fi_ref = None
        col_fi_ref = None
        cols_fi_ref = None
        print("═" * 55)
        print()
        print(" DATOS PROCESADOS")
        print("├── FIBRA........: ", end="", flush=True)
        try:
            wb_fi = None
            try:
                wb_fi = openpyxl.load_workbook(RUTA_FIBRAS_BULK, data_only=True, read_only=True)
                ws_fi = wb_fi.worksheets[0]
                col_fi = extraer_indices_columnas(ws_fi)
                err_fi, espec_fi, filas_fi = procesar_fibras(ws_fi, col_fi, cols_fi, ids_om, responsables, modo, odfs_data, organizador)
                contadores['fibra'] = filas_fi
                errores_por_tipo['FIBRA'] = len(err_fi)
                ws_fi_ref = ws_fi
                col_fi_ref = col_fi
                cols_fi_ref = cols_fi
                print(f"{contadores['fibra']:,}")
            except Exception as e:
                print(f"0 [Error: {e}]")
            finally:
                if wb_fi:
                    try:
                        wb_fi.close()
                    except:
                        pass
        except Exception:
            print("0")
        print("├── EMPALME......: ", end="", flush=True)
        try:
            wb_em = None
            try:
                wb_em = openpyxl.load_workbook(RUTA_EMPALMES_BULK, data_only=True, read_only=True)
                ws_em = wb_em.worksheets[0]
                col_em = extraer_indices_columnas(ws_em)
                err_em, datos_em, filas_em = procesar_empalmes(ws_em, col_em, cols_em, ids_om, responsables, modo, organizador)
                contadores['empalme'] = filas_em
                errores_por_tipo['EMPALME'] = len(err_em)
                if datos_em:
                    err_consec = validar_consecutividad_empalmes(datos_em, ids_om, responsables)
                    err_tipo_cable = validar_consistencia_tipo_cable_empalmes(datos_em, ids_om, responsables)
                    errores_por_tipo['EMPALME'] += len(err_consec) + len(err_tipo_cable)
                    for error in err_consec:
                        organizador.agregar_error(error[0], "EMPALME", error)
                    for error in err_tipo_cable:
                        organizador.agregar_error(error[0], "EMPALME", error)
                print(f"{contadores['empalme']:,}")
            except Exception as e:
                print(f"0 [Error: {e}]")
            finally:
                if wb_em:
                    try:
                        wb_em.close()
                    except:
                        pass
        except Exception:
            print("0")
        print("├── EDIFICIO.....: ", end="", flush=True)
        try:
            err_ed, err_estructura, filas_ed = procesar_edificios_y_estructuras(
                RUTA_EDIFICIOS_BULK, ids_om, exc_ed, responsables_edificio, modo, organizador
            )
            contadores['edificio'] = filas_ed
            contadores['estructura'] = filas_ed
            errores_por_tipo['EDIFICIO'] = len(err_ed)
            errores_por_tipo['ESTRUCTURA'] = len(err_estructura)
            err_similares = detectar_edificios_con_nombres_similares(RUTA_EDIFICIOS_BULK)
            errores_por_tipo['EDIFICIO'] += len(err_similares)
            for error in err_similares:
                organizador.agregar_error(error[0], "EDIFICIO", error)
            print(f"{contadores['edificio']:,}")
        except Exception as e:
            print(f"0 [Error: {e}]")
        print("├── PORTS........: ", end="", flush=True)
        t_ports = time.time()
        try:
            if RUTA_PORTS_BULK.exists():
                err_ports, err_shelf, err_bay, filas_ports = procesar_ports(
                    RUTA_PORTS_BULK, ids_om, modo, nombres_edificios, nombres_normalizados
                )
                contadores['ports'] = filas_ports
                total_ports_errors = len(err_ports) + len(err_shelf) + len(err_bay)
                errores_por_tipo['PORTS'] = total_ports_errors
                for error in err_ports:
                    organizador.agregar_error(error[0], "PORT", error)
                for error in err_shelf:
                    organizador.agregar_error(error[0], "SHELF", error)
                for error in err_bay:
                    organizador.agregar_error(error[0], "BAY", error)
                print(f"{contadores['ports']:,} | {total_ports_errors} [{time.time() - t_ports:.1f}s]")
            else:
                print("0 (archivo no existe)")
        except Exception as e:
            print(f"0 [Error: {e}]")
        print("└── MERGE........: ", end="", flush=True)
        t_merge = time.time()
        try:
            espec_fi_respaldo = None
            colocacion_por_fibra = None
            if RUTA_FIBRAS_BULK.exists():
                espec_fi_respaldo, colocacion_por_fibra = cargar_especificaciones_fibra(
                    RUTA_FIBRAS_BULK, ids_om, "masivo", incluir_colocacion=True
                )
            if datos_em and (espec_fi or espec_fi_respaldo) and contadores['empalme'] > 0 and contadores['fibra'] > 0:
                todas_colocaciones = {}
                if colocacion_por_fibra:
                    todas_colocaciones.update(colocacion_por_fibra)
                err_merge = validar_merge_fibra_empalme(
                    datos_em, espec_fi, responsables, ids_om,
                    espec_fi_respaldo, todas_colocaciones,
                    ws_fi_ref, col_fi_ref, cols_fi_ref,
                )
                errores_por_tipo['MERGE'] = len(err_merge)
                for error in err_merge:
                    organizador.agregar_error(error[0], "MERGE FI-EM", error)
                print(f"{errores_por_tipo['MERGE']} [{time.time() - t_merge:.1f}s]")
            else:
                print("0 [0.0s]")
        except Exception as e:
            print(f"0 [0.0s] - {e}")
        print()
        print("═" * 55)
        print()
        print("PROCESANDO ESCRITURA... ", end="", flush=True)
        t_escritura = time.time()
        errores_finales = organizador.obtener_errores_organizados()
        for elemento, valores in errores_finales:
            if isinstance(valores, (list, tuple)):
                v = list(valores) + [""] * max(0, 4 - len(valores))
                v_limpios = [
                    v[0],
                    limpiar_valor_reporte(v[1]) if len(v) > 1 else "",
                    v[2] if len(v) > 2 else "",
                    v[3] if len(v) > 3 else "",
                ]
                ws_salida.append([elemento] + v_limpios + ["..."])
            else:
                valor_limpio = limpiar_valor_reporte(valores)
                ws_salida.append([elemento, valor_limpio, "", "", "Formato de error no reconocido", ""])
        print(f"[{time.time() - t_escritura:.1f}s]")
        print("├── FORMATO FINAL........: ", end="", flush=True)
        t_formato = time.time()
        aplicar_formato_hoja_resultados(ws_salida, len(errores_finales))
        print(f"[{time.time() - t_formato:.1f}s]")
        total_procesados = sum(contadores.values())
        total_errores_encontrados = sum(errores_por_tipo.values())
        print(f"├── TOTAL PROCESADOS........: {total_procesados:,}")
        print(f"├── ERRORES ENCONTRADOS.....: {total_errores_encontrados:,}")
        print(f"├── TIEMPO DE EJECUCIÓN.....: {time.time() - start_time:.2f} s")
        print("├── GS_EXPORT .........: ", end="", flush=True)
        t_guardado = time.time()
        try:
            wb.save(str(RUTA_MACRO))
            print(f"[{time.time() - t_guardado:.1f}s] (guardado correctamente)")
        except Exception as e:
            print(f"\n[ERROR] No se pudo guardar el archivo: {e}")
    finally:
        if wb:
            try:
                wb.close()
            except:
                pass
    return total_errores_encontrados

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    print("═" * 55)
    print()
    print(" RUTAS / COLUMNAS")
    archivos = [
        ("FIBRA",     RUTA_FIBRAS_BULK),
        ("EMPALME",   RUTA_EMPALMES_BULK),
        ("EDIFICIO",  RUTA_EDIFICIOS_BULK),
        ("ESTRUCTURA",RUTA_EDIFICIOS_BULK),
        ("PORTS",     RUTA_PORTS_BULK),
    ]
    hay_problemas = False
    for i, (nombre, ruta) in enumerate(archivos):
        status = "(OK)" if ruta.exists() else "(NO ENCONTRADO)"
        if not ruta.exists():
            hay_problemas = True
        prefijo = "└──" if i == len(archivos) - 1 else "├──"
        print(f"{prefijo} {nombre:12}: {status}")
    if not hay_problemas:
        total_errores = ejecutar_validacion_calidad("masivo")
        print(f"\n└──GS ERRORES......:[{total_errores}]")
    else:
        print("\nERROR: Se detectaron problemas en los archivos.")