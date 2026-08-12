#!/usr/bin/env python
# -*- coding: utf-8 -*-


import sys
import re
import logging
import multiprocessing
import time
import os
from pathlib import Path
from typing import Dict, List, Any, Set, Tuple
from collections import Counter
from datetime import datetime

import pandas as pd
import win32com.client as win32
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink


STATUS_OK = "OK"
ANOMALY_SEPARATOR = " | "


NO_ET_EMPALMES = "TRAZE ERRONEO"
NO_NODOS_NC = "TRAZE ERRONEO"
ED_WITHOUT_OB = "(FALTA OB {cable_id}, {hilo})"
OB_WITH_MULTIPLE_EDS = "MULTIPLES ELEMENTOS {num_ed} , ({cable_id}, {hilo})"
NO_RESPONSABLE = "SIN RESPONSABLE"
CLIENT_COUNT_MISMATCH = "NCLIENTES (TZ:{clientes_actuales}, RESP:{total_clientes_num})"




CARPETA_PADRE = Path(
    r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos"
    r"\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO"
    r"\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\11_Trace_Conectividad"
)

RUTA_MOTOR = Path(r"C:\A_GS1_PROYECTOS\00_GV_TRAZE.xlsm")
RUTA_RESPONSABLES = Path(r"C:\A_GS1_PROYECTOS\Reporte anillos-clientes.xlsx")
RUTA_SALIDA_BASE = Path(r"C:\A_GS1_PROYECTOS")
RUTA_EXCLUSIONES = Path(
    r"C:\A_GS1_PROYECTOS\0_Documents_gs\output\quality\03-TRAZE_MS_CONNECTION.xlsx"
)

MAX_WORKERS = 4
logging.basicConfig(level=logging.CRITICAL)
logger = logging.getLogger(__name__)


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


def normalizar_anillo(nombre: str) -> str:
    return nombre.strip().upper()


def parse_fecha(fecha_str: str) -> datetime:
    if not fecha_str:
        return datetime.min
    fecha_str = fecha_str.strip()
    try:
        return datetime.strptime(fecha_str, "%d-%b")
    except ValueError:
        try:
            return datetime.strptime(fecha_str, "%d-%b-%y")
        except ValueError:
            try:
                return datetime.strptime(fecha_str, "%d/%m/%Y")
            except ValueError:
                return datetime.min


class ResponsableLoader:
    def __init__(self, ruta: Path):
        self.ruta = ruta
        self.datos: Dict[str, Dict[str, str]] = {}

    def cargar(self) -> Dict[str, Dict[str, str]]:
        if not self.ruta.exists():
            return {}

        try:
            df = pd.read_excel(self.ruta, engine="openpyxl", dtype=str)
            df.columns = [str(c).strip().lower() for c in df.columns]

            col_anillo = None
            col_responsable = None
            col_total_clientes = None
            col_fecha = None

            for col in df.columns:
                if "anillo" in col:
                    col_anillo = col
                elif "responsable" in col:
                    col_responsable = col
                elif "total" in col and "cliente" in col:
                    col_total_clientes = col
                elif "clientes" in col:
                    col_total_clientes = col
                elif "fecha" in col:
                    col_fecha = col

            if col_anillo is None and len(df.columns) > 2:
                col_anillo = df.columns[2]
            if col_responsable is None and len(df.columns) > 1:
                col_responsable = df.columns[1]
            if col_total_clientes is None and len(df.columns) > 4:
                col_total_clientes = df.columns[4]
            if col_fecha is None and len(df.columns) > 3:
                col_fecha = df.columns[3]

            if col_anillo is None or col_responsable is None:
                return {}

            mejores = {}

            for _, row in df.iterrows():
                anillo = str(row[col_anillo]).strip() if pd.notna(row[col_anillo]) else ""
                if not anillo:
                    continue

                responsable = str(row[col_responsable]).strip() if pd.notna(row[col_responsable]) else ""
                total_clientes = ""
                if col_total_clientes and pd.notna(row[col_total_clientes]):
                    total_clientes = str(row[col_total_clientes]).strip()
                    if total_clientes and not total_clientes.replace('.', '', 1).isdigit():
                        total_clientes = ""

                fecha_str = ""
                if col_fecha and pd.notna(row[col_fecha]):
                    fecha_str = str(row[col_fecha]).strip()

                anillo_norm = normalizar_anillo(anillo)

                if anillo_norm in mejores:
                    fecha_existente = mejores[anillo_norm].get("fecha_parse", datetime.min)
                    fecha_nueva = parse_fecha(fecha_str)
                    if fecha_nueva > fecha_existente:
                        mejores[anillo_norm] = {
                            "responsable": responsable,
                            "total_clientes": total_clientes,
                            "fecha_parse": fecha_nueva
                        }
                else:
                    mejores[anillo_norm] = {
                        "responsable": responsable,
                        "total_clientes": total_clientes,
                        "fecha_parse": parse_fecha(fecha_str)
                    }

            datos = {}
            for anillo, info in mejores.items():
                datos[anillo] = {
                    "responsable": info["responsable"],
                    "total_clientes": info["total_clientes"],
                }
            return datos

        except Exception:
            return {}

    @staticmethod
    def extraer_anillo(nombre_hoja: str) -> str:
        partes = nombre_hoja.split("_")
        if len(partes) >= 3:
            return normalizar_anillo(partes[2])
        return normalizar_anillo(nombre_hoja)


class TableParser:
    def __init__(self):
        self.reset()

    def reset(self):
        self.fibraActual = ""
        self.cableActual = ""
        self.ultimoI = ""
        self.ultimoJ = ""
        self.ultimoK = ""
        self.ultimaFibraEscrita = ""
        self.forzarSiguienteEmpalme = False
        self.enEdificio = False
        self.edificioActual = ""
        self.fibraEdificio = ""
        self.puertosEdificio = ""
        self.initialPorts = ""
        self.hasSeenFibre = False
        self.lastEmpalmeNombre = ""
        self.lastEmpalmeCable = ""
        self.lastEmpalmeFibra = ""
        self.lastEmpalmeWasWritten = False
        self.esNodo = False
        self.resultados: List[Tuple[str, str, str]] = []

    def _buscar_parentesis_cierre(self, texto: str, posApertura: int) -> int:
        nivel = 1
        i = posApertura + 1
        while i < len(texto) and nivel > 0:
            if texto[i] == "(":
                nivel += 1
            elif texto[i] == ")":
                nivel -= 1
            i += 1
        return i - 1 if nivel == 0 else -1

    def _extraer_info_fibra_cable(self, linea: str) -> Tuple[str, str, str]:
        numFibra = ""
        hilosCable = ""
        idCable = ""

        posDosPuntos = linea.find(":")
        if posDosPuntos != -1:
            posParen = linea.find(")", posDosPuntos)
            if posParen != -1:
                numStr = linea[posDosPuntos + 1 : posParen].strip()
                m = re.match(r"^\d+", numStr)
                if m:
                    numFibra = m.group()

        posCable = linea.find("Cable de FO (")
        if posCable != -1:
            posParenCable = self._buscar_parentesis_cierre(
                linea, posCable + len("Cable de FO (") - 1
            )
            if posParenCable != -1:
                contenido = linea[posCable + len("Cable de FO (") : posParenCable]
                posDuctado = contenido.find("Ductado ")
                if posDuctado != -1:
                    posH = contenido.find("h", posDuctado + 8)
                    if posH != -1:
                        hilosCable = contenido[posDuctado + 8 : posH].strip()
                posComilla1 = contenido.find("'")
                if posComilla1 != -1:
                    posComilla2 = contenido.find("'", posComilla1 + 1)
                    if posComilla2 != -1:
                        idCable = contenido[posComilla1 + 1 : posComilla2]
        return numFibra, hilosCable, idCable

    def _extraer_nombre_edificio(self, linea: str) -> str:
        posEdificio = linea.find("Edificio")
        if posEdificio == -1:
            return ""
        posComilla1 = linea.find("'", posEdificio)
        if posComilla1 == -1:
            return ""
        posComilla2 = linea.find("'", posComilla1 + 1)
        if posComilla2 == -1:
            return ""
        nombre = linea[posComilla1 + 1 : posComilla2]
        if "NODO" in linea:
            return nombre
        else:
            return "(ED) " + nombre

    def _extraer_nombre_empalme(self, linea: str) -> str:
        posIni = linea.find("at Empalme de Fibra (")
        if posIni == -1:
            return ""
        posParen = self._buscar_parentesis_cierre(
            linea, posIni + len("at Empalme de Fibra (") - 1
        )
        if posParen == -1:
            return ""
        contenido = linea[posIni + len("at Empalme de Fibra (") : posParen]

        tipo = ""
        id_ = ""
        if "Empalme Cliente" in contenido:
            tipo = "EM CL"
            posComa = contenido.find(",")
            if posComa == -1:
                return ""
            posComilla1 = contenido.find("'", posComa)
            if posComilla1 == -1:
                return ""
            posComilla2 = contenido.find("'", posComilla1 + 1)
            if posComilla2 == -1:
                return ""
            id_ = contenido[posComilla1 + 1 : posComilla2]
        else:
            posTipo = contenido.find("Tipo ")
            if posTipo == -1:
                return ""
            numTipo = ""
            i = posTipo + 5
            while i < len(contenido) and contenido[i].isdigit():
                numTipo += contenido[i]
                i += 1
            tipo = "EM " + numTipo
            posComa = contenido.find(",", posTipo)
            if posComa == -1:
                return ""
            posComilla1 = contenido.find("'", posComa)
            if posComilla1 == -1:
                return ""
            posComilla2 = contenido.find("'", posComilla1 + 1)
            if posComilla2 == -1:
                return ""
            id_ = contenido[posComilla1 + 1 : posComilla2]
        return f"({tipo}) | {id_}"

    def _extraer_puerto(self, linea: str) -> str:
        puertoNum = ""
        cardName = ""
        shelfName = ""
        bayName = ""
        rackName = ""
        subrackName = ""

        pos = linea.find("Port (")
        if pos != -1:
            posParen = linea.find(")", pos + 6)
            if posParen != -1:
                puertoNum = linea[pos + 6 : posParen].strip()
        pos = linea.find("Card (")
        if pos != -1:
            posParen = linea.find(")", pos + 6)
            if posParen != -1:
                cardName = linea[pos + 6 : posParen].strip()
        pos = linea.find("Shelf (")
        if pos != -1:
            posParen = linea.find(")", pos + 7)
            if posParen != -1:
                shelfName = linea[pos + 7 : posParen].strip()
        pos = linea.find("Bay (")
        if pos != -1:
            posParen = linea.find(")", pos + 5)
            if posParen != -1:
                bayName = linea[pos + 5 : posParen].strip()
        pos = linea.find("Rack (")
        if pos != -1:
            posParen = linea.find(")", pos + 6)
            if posParen != -1:
                rackName = linea[pos + 6 : posParen].strip()
        pos = linea.find("Subrack (")
        if pos != -1:
            posParen = linea.find(")", pos + 9)
            if posParen != -1:
                subrackName = linea[pos + 9 : posParen].strip()

        resultado = ""
        if rackName:
            resultado = rackName
        if subrackName:
            resultado = (resultado + ", " if resultado else "") + subrackName
        if bayName:
            resultado = (resultado + ", " if resultado else "") + bayName
        if shelfName:
            resultado = (resultado + ", " if resultado else "") + shelfName
        if cardName:
            resultado = (resultado + ", " if resultado else "") + cardName
        if puertoNum:
            resultado = (resultado + ", " if resultado else "") + puertoNum
        return resultado

    def _agregar_resultado(
        self, elemento: str, cable: str, fibraHilo: str, force: bool = False
    ) -> bool:
        if not force:
            if (
                elemento == self.ultimoI
                and cable == self.ultimoJ
                and fibraHilo == self.ultimoK
            ):
                return False
        self.resultados.append((elemento, cable, fibraHilo))
        self.ultimoI = elemento
        self.ultimoJ = cable
        self.ultimoK = fibraHilo
        return True

    def parse(self, lineas: List[str]) -> List[Tuple[str, str, str]]:
        self.reset()

        for linea_raw in lineas:
            linea = linea_raw.strip()
            if not linea:
                continue

            if "Port (" in linea and not self.hasSeenFibre:
                if self.initialPorts == "":
                    self.initialPorts = linea_raw.rstrip("\n")
                else:
                    self.initialPorts += "\n" + linea_raw.rstrip("\n")

            if "Fibre (" in linea:
                numFibra, hilosCable, idCable = self._extraer_info_fibra_cable(linea)
                if numFibra:
                    self.fibraActual = numFibra
                if hilosCable and idCable:
                    self.cableActual = f"({hilosCable}h) | {idCable}"
                self.hasSeenFibre = True

            if "Edificio" in linea:
                if self.enEdificio:
                    if self.puertosEdificio:
                        elemento = f"{self.edificioActual} | {self.puertosEdificio}"
                        self._agregar_resultado(
                            elemento, self.cableActual, f"H{self.fibraEdificio}"
                        )
                    self.enEdificio = False
                    self.puertosEdificio = ""
                    self.esNodo = False

                nombre = self._extraer_nombre_edificio(linea)
                if nombre:
                    self.edificioActual = nombre
                    self.fibraEdificio = self.fibraActual
                    self.enEdificio = True
                    self.puertosEdificio = ""
                    self.esNodo = "NODO" in nombre
                    if self.initialPorts:
                        self.puertosEdificio = self.initialPorts
                        self.initialPorts = ""
                    self.forzarSiguienteEmpalme = True

            if "Empalme de Fibra" in linea:
                if self.enEdificio:
                    if self.puertosEdificio:
                        elemento = f"{self.edificioActual} | {self.puertosEdificio}"
                        self._agregar_resultado(
                            elemento, self.cableActual, f"H{self.fibraEdificio}"
                        )
                    self.enEdificio = False
                    self.puertosEdificio = ""
                    self.esNodo = False

                nombreEmpalme = self._extraer_nombre_empalme(linea)
                if nombreEmpalme:
                    self.lastEmpalmeNombre = nombreEmpalme
                    self.lastEmpalmeCable = self.cableActual
                    self.lastEmpalmeFibra = f"H{self.fibraActual}"
                    self.lastEmpalmeWasWritten = False

                    esET = "(EM " in nombreEmpalme and "(EM CL)" not in nombreEmpalme
                    if esET:
                        escribir = self.forzarSiguienteEmpalme or (
                            self.fibraActual != self.ultimaFibraEscrita
                        )
                    else:
                        escribir = True

                    if escribir:
                        if self._agregar_resultado(
                            nombreEmpalme, self.cableActual, f"H{self.fibraActual}"
                        ):
                            self.ultimaFibraEscrita = self.fibraActual
                            self.lastEmpalmeWasWritten = True

                    self.forzarSiguienteEmpalme = False

            if self.enEdificio and "Port (" in linea and self.hasSeenFibre:
                if self.esNodo:
                    lineaCompleta = linea_raw.rstrip("\n")
                    if self.puertosEdificio == "":
                        self.puertosEdificio = lineaCompleta
                    else:
                        self.puertosEdificio += "\n" + lineaCompleta
                else:
                    puertoInfo = self._extraer_puerto(linea)
                    if puertoInfo:
                        if self.puertosEdificio == "":
                            self.puertosEdificio = puertoInfo
                        else:
                            self.puertosEdificio += ", " + puertoInfo

        if not self.lastEmpalmeWasWritten and self.lastEmpalmeNombre:
            self._agregar_resultado(
                self.lastEmpalmeNombre,
                self.lastEmpalmeCable,
                self.lastEmpalmeFibra,
                force=True,
            )

        if self.enEdificio and self.puertosEdificio:
            elemento = f"{self.edificioActual} | {self.puertosEdificio}"
            self._agregar_resultado(elemento, self.cableActual, f"H{self.fibraEdificio}")

        if self.initialPorts:
            elemento = f"NODO INICIAL | {self.initialPorts}"
            self._agregar_resultado(elemento, "", "")

        return self.resultados


def analizar_tabla(filas: List[Tuple[str, str, str]]) -> str:
    anomalies = []

    et_count = sum(1 for elem, _, _ in filas if "ET" in elem and "(EM" in elem)
    if et_count == 0:
        anomalies.append(NO_ET_EMPALMES)

    grupos = {}
    for elem, cable, hilo in filas:
        key = (cable, hilo)
        tipo = None
        if "ED:" in elem or "(ED)" in elem:
            tipo = "ED"
        elif "(EM CL)" in elem:
            tipo = "OB"
        elif "(EM" in elem:
            if "EA" in elem:
                tipo = "EA"
            elif "ET" in elem:
                tipo = "ET"
        if tipo:
            grupos.setdefault(key, []).append(tipo)

    for (cable, hilo), tipos in grupos.items():
        if "ED" in tipos and "OB" not in tipos:
            if cable and "|" in cable:
                cable_id = cable.split("|", 1)[1].strip()
            else:
                cable_id = cable
            anomalies.append(ED_WITHOUT_OB.format(cable_id=cable_id, hilo=hilo))

    for (cable, hilo), tipos in grupos.items():
        if "OB" in tipos:
            num_ed = tipos.count("ED")
            if num_ed > 2:
                if cable and "|" in cable:
                    cable_id = cable.split("|", 1)[1].strip()
                else:
                    cable_id = cable
                anomalies.append(OB_WITH_MULTIPLE_EDS.format(num_ed=num_ed, cable_id=cable_id, hilo=hilo))

    if anomalies:
        return ANOMALY_SEPARATOR.join(anomalies)
    else:
        return STATUS_OK


def dibujar_grafica(ruta_archivo: Path, ruta_motor: Path, reintentos: int = 3) -> bool:
    excel = None
    for intento in range(reintentos):
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False

            wb_motor = excel.Workbooks.Open(str(ruta_motor), ReadOnly=True)
            wb_destino = excel.Workbooks.Open(str(ruta_archivo))

            for ws in wb_destino.Sheets:
                ws.Activate()
                try:
                    excel.Application.Run(f"'{ruta_motor.name}'!G01_02_DIAGRAMA_COMPONENTES")
                except Exception:
                    pass

            wb_destino.Save()
            wb_destino.Close()
            wb_motor.Close()
            return True
        except Exception:
            if intento < reintentos - 1:
                time.sleep(1)
                continue
            else:
                return False
        finally:
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
    return False


def procesar_carpeta(
    carpeta: Path, exclusiones: Set[str], responsables: Dict[str, Dict[str, str]]
) -> Tuple[List[Dict[str, str]], Set[str]]:
    cable_nombre = carpeta.name

    if cable_nombre in exclusiones:
        print(f"[{cable_nombre}] Excluida")
        return [], set()

    todos_txt = sorted(carpeta.glob("*.txt"))
    archivos_txt = []
    for archivo in todos_txt:
        if archivo.stem in exclusiones:
            print(f"[{cable_nombre}] Excluido: {archivo.name}")
        else:
            archivos_txt.append(archivo)

    if not archivos_txt:
        return [], set()

    anillos_presentes = set()
    info_hojas = []
    for archivo in archivos_txt:
        nombre_base = archivo.stem
        nombre_hoja = nombre_base[:31]
        anillo = ResponsableLoader.extraer_anillo(nombre_base)
        anillos_presentes.add(anillo)
        info_hojas.append(
            {
                "nombre_hoja": nombre_hoja,
                "anillo": anillo,
                "clientes_trace": "",
                "estado": "",
            }
        )

    responsables_anillos = {
        anillo: responsables.get(anillo, {}).get("responsable", "")
        for anillo in anillos_presentes
    }
    contador = Counter(responsables_anillos.values())
    responsable_mayoritario = contador.most_common(1)[0][0] if contador else ""
    if responsable_mayoritario:
        nombre_archivo = f"{cable_nombre}_{responsable_mayoritario}.xlsx"
    else:
        nombre_archivo = f"{cable_nombre}.xlsx"
    ruta_salida = RUTA_SALIDA_BASE / nombre_archivo

    wb = Workbook()
    wb.remove(wb.active)

    parser = TableParser()
    all_filas = []

    for idx, archivo in enumerate(archivos_txt):
        nombre_base = archivo.stem
        nombre_hoja = nombre_base[:31]
        with open(archivo, "r", encoding="utf-8", errors="ignore") as f:
            lineas = [l.rstrip("\n") for l in f]
        filas = parser.parse(lineas)
        all_filas.extend(filas)

        ws = wb.create_sheet(title=nombre_hoja)

        for row_idx, line in enumerate(lineas, start=4):
            ws.cell(row=row_idx, column=2, value=line)

        ws["I4"] = "ELEMENTO"
        ws["J4"] = "CABLE"
        ws["K4"] = "HILO"
        for i, (elem, cable, hilo) in enumerate(filas, start=5):
            ws.cell(row=i, column=9, value=elem)
            ws.cell(row=i, column=10, value=cable)
            ws.cell(row=i, column=11, value=hilo)

        clientes_actuales = sum(1 for elem, _, _ in filas if "ED:" in elem or "(ED)" in elem)
        info_hojas[idx]["clientes_trace"] = str(clientes_actuales)

        anillo = info_hojas[idx]["anillo"]
        responsable_anillo = responsables.get(anillo, {}).get("responsable", "")
        total_clientes_str = responsables.get(anillo, {}).get("total_clientes", "")
        total_clientes_num = 0
        if total_clientes_str and total_clientes_str.isdigit():
            total_clientes_num = int(total_clientes_str)

        estado_base = analizar_tabla(filas)
        anomalias_extra = []

        if not responsable_anillo:
            anomalias_extra.append(NO_RESPONSABLE)
        if total_clientes_str and total_clientes_str.isdigit() and clientes_actuales != total_clientes_num:
            anomalias_extra.append(CLIENT_COUNT_MISMATCH.format(clientes_actuales=clientes_actuales, total_clientes_num=total_clientes_num))

        if estado_base == STATUS_OK and not anomalias_extra:
            estado_final = STATUS_OK
        elif estado_base == STATUS_OK and anomalias_extra:
            estado_final = ANOMALY_SEPARATOR.join(anomalias_extra)
        elif estado_base != STATUS_OK and anomalias_extra:
            estado_final = f"{estado_base}{ANOMALY_SEPARATOR}{ANOMALY_SEPARATOR.join(anomalias_extra)}"
        else:
            estado_final = estado_base

        info_hojas[idx]["estado"] = estado_final

        for col in [2, 9, 10, 11]:
            max_len = 0
            for row in range(4, max(len(lineas), len(filas)) + 5):
                cell = ws.cell(row=row, column=col)
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[get_column_letter(col)].width = min(max_len + 2, 50)

    ws_idx = wb.create_sheet(title="ÍNDICE", index=0)
    headers = [
        "ARCHIVO",
        "HOJA",
        "CLIENTES (TRACE / RESPONSABLE)",
        "ESTADO",
    ]
    for col, titulo in enumerate(headers, start=1):
        ws_idx.cell(row=1, column=col, value=titulo)
        ws_idx.cell(row=1, column=col).font = Font(bold=True)

    for i, info in enumerate(info_hojas, start=2):
        nombre_hoja = info["nombre_hoja"]
        anillo = info["anillo"]
        clientes_trace = info["clientes_trace"]
        estado = info["estado"]

        cell = ws_idx.cell(row=i, column=1, value=nombre_hoja)
        cell.hyperlink = Hyperlink(ref=cell.coordinate, target=f"'{nombre_hoja}'!Q5")
        cell.font = Font(color="0563C1", underline="single")

        ws_idx.cell(row=i, column=2, value=nombre_hoja)

        info_anillo = responsables.get(anillo, {"responsable": "", "total_clientes": ""})
        total_clientes = info_anillo.get("total_clientes", "")
        clientes_display = f"{clientes_trace} / {total_clientes}" if total_clientes else clientes_trace
        ws_idx.cell(row=i, column=3, value=clientes_display)

        ws_idx.cell(row=i, column=4, value=estado)

    for col in range(1, 5):
        max_len = 0
        for row in range(1, len(info_hojas) + 2):
            cell = ws_idx.cell(row=row, column=col)
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws_idx.column_dimensions[get_column_letter(col)].width = min(max_len + 2, 30)

    try:
        wb.save(ruta_salida)
    except Exception as e:
        print(f"[{cable_nombre}] Error al guardar tabla: {e}")
        return [], set()

    time.sleep(0.5)

    hay_nc = any("NODO" in elem for elem, _, _ in all_filas)
    if hay_nc:
        dibujar_grafica(ruta_salida, RUTA_MOTOR)
    else:
        for info in info_hojas:
            if info["estado"] == STATUS_OK:
                info["estado"] = NO_NODOS_NC
            else:
                info["estado"] += f"{ANOMALY_SEPARATOR}{NO_NODOS_NC}"
        for i, info in enumerate(info_hojas, start=2):
            ws_idx.cell(row=i, column=4, value=info["estado"])

    tiene_anomalias = any(info["estado"] != STATUS_OK for info in info_hojas)
    if tiene_anomalias:
        print(f"ERROR | {nombre_archivo}")
    else:
        print(f"OK | {nombre_archivo}")

    resultados = []
    for info in info_hojas:
        if info["estado"] != STATUS_OK:
            resultados.append(
                {
                    "archivo": nombre_archivo,
                    "hoja": info["nombre_hoja"],
                    "estado": info["estado"],
                }
            )
    return resultados, anillos_presentes


def escribir_reporte(resultados: List[Dict[str, str]], ruta_reporte: Path):
    if not resultados:
        return

    df = pd.DataFrame(resultados)
    df.columns = ["ARCHIVO", "HOJA", "ANOMALÍAS"]

    if ruta_reporte.exists():
        wb = load_workbook(ruta_reporte)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    if "REPORT" in wb.sheetnames:
        del wb["REPORT"]

    ws = wb.create_sheet("REPORT")

    # Cabeceras: primera columna "RUTA"
    headers = ["RUTA", "HOJA", "ANOMALÍAS"]
    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=header)
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    # Directorio donde se guarda el reporte
    report_dir = ruta_reporte.parent

    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        archivo_nombre = row[0]
        hoja_nombre = row[1]
        anomalias = row[2]

        archivo_path = RUTA_SALIDA_BASE / archivo_nombre

        if archivo_path.exists():
            # Calcular ruta relativa desde la carpeta del reporte hasta el archivo
            try:
                rel_path = os.path.relpath(str(archivo_path), start=str(report_dir))
                # En Windows, relpath usa backslashes, perfecto para Excel
            except Exception:
                rel_path = str(archivo_path)  # fallback absoluto

            # Escapar comillas simples en el nombre de la hoja
            hoja_escapada = hoja_nombre.replace("'", "''")
            # Construir el destino: ruta_relativa#'hoja_escapada'!Q5
            target = f"{rel_path}#'{hoja_escapada}'!Q5"

            cell = ws.cell(row=row_idx, column=1, value=rel_path)
            link = Hyperlink(ref=cell.coordinate, target=target, tooltip=f"Ir a {hoja_nombre} Q5")
            cell.hyperlink = link
            cell.font = Font(color="0563C1", underline="single")
        else:
            # Si no existe, mostramos el nombre del archivo sin enlace
            ws.cell(row=row_idx, column=1, value=archivo_nombre)

        ws.cell(row=row_idx, column=2, value=hoja_nombre)
        ws.cell(row=row_idx, column=3, value=anomalias)

    # Ajustar ancho de columnas
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(ruta_reporte)
    print(f"Reporte actualizado en {ruta_reporte}")


def main():
    exclusiones = cargar_exclusiones(RUTA_EXCLUSIONES)
    loader = ResponsableLoader(RUTA_RESPONSABLES)
    responsables = loader.cargar()

    subcarpetas = [p for p in CARPETA_PADRE.iterdir() if p.is_dir()]
    if not subcarpetas:
        print("No se encontraron carpetas.")
        return

    print(f"Procesando {len(subcarpetas)} carpetas con {MAX_WORKERS} workers...")

    with multiprocessing.Pool(processes=MAX_WORKERS) as pool:
        args = [(carpeta, exclusiones, responsables) for carpeta in subcarpetas]
        resultados_parciales = pool.starmap(procesar_carpeta, args)

    todos_resultados = []
    anillos_presentes = set()
    for anomalias, anillos in resultados_parciales:
        todos_resultados.extend(anomalias)
        anillos_presentes.update(anillos)

    missing_anillos = set(responsables.keys()) - anillos_presentes
    for anillo in missing_anillos:
        todos_resultados.append(
            {
                "archivo": anillo,
                "hoja": "responsable",
                "estado": "FALTA ARCHIVO",
            }
        )

    if todos_resultados:
        escribir_reporte(todos_resultados, RUTA_EXCLUSIONES)
    else:
        print("No se encontraron anomalías, el reporte no se actualiza.")

    succeeded = len(subcarpetas) - sum(1 for r in resultados_parciales if not r[0] and r[1])
    failed = len(subcarpetas) - succeeded
    print(f"\nCompletado. Éxitos: {succeeded}, Fallos: {failed}")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()


    