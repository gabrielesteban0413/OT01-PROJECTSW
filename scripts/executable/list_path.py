"""
Extractor de datos desde archivos TXT (IP, anillo, ancho de banda, ID Servicio, coordenadas,
Tipo de Servicio, Equipo Central, Puerto Common, Puerto Owner, OC, Etiqueta)
Genera un archivo Excel con los resultados en el orden solicitado.
"""

import re
import pandas as pd
from pathlib import Path
from typing import Optional, Tuple, List, Dict, Any
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

RUTA_ORIGEN = Path(r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\02_ConversionCAD_Shape\Archivos_Service_PATH")
RUTA_DESTINO = Path(r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output")
RUTA_CENTRALES = Path(r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\03-props_centrales.xlsx")
NOMBRE_SALIDA = "01-Bulk_Path.xlsx"
EXTENSION = ".txt"

PATRONES = {
    'ip_demarcador': [
        r'IP\s+DEMARCADOR\s*:\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
        r'DEMARCADOR\s*:\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
        r'IP\s*DEMAR\s*:\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
        r'DEMARCADOR\s+(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
        r'IP\s+DEMARCADOR\s+(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
        r'IP\s*DEMAR\s+(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
    ],
    'ip_general': [
        r'Connected\s+to\s+(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
        r'IP\s*:\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})',
        r'\b(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b',
    ],
    'anillo': [
        r'ANILLO\s*O\s*NODO\s*:\s*([A-Z0-9]{8,15})',
        r'ANILLO:\s*([A-Z0-9]{8,15})',
        r'RING:\s*([A-Z0-9]{8,15})',
        r'SCJ\s*:\s*([A-Z0-9]{8,15})',
        r'ORIGEN\s*:\s*([A-Z0-9]{8,15})',
        r'\b(MEAD[A-Z0-9]{8,15})\b',
        r'\b([A-Z]{4,}[A-Z0-9]{4,})\b',
    ],
    'ancho_banda': [
        r'ANCHO DE BANDA FINAL\s*[:=]?\s*([^\n]+)',
        r'BW:\s*([^\n]+)',
        r'BW\s+([^\n]+)',
        r'ANCHO DE BANDA SOLICITADO\s*([^\n]+)',
    ],
    'id_servicio': [
        r'ID\s+DEL\s+SERVICIO\s+AFECTADO\s*:\s*([A-Z0-9]+)',
        r'ID\s+DEL\s+SERVICIO\s+AFECTADO\s*:\s*([^\n]+)',
        r'ID\s+SERVICIO\s*:\s*([A-Z0-9]+)',
        r'ID\s+SERVICIO\s*:\s*([^\n]+)',
        r'ID\s+DEL\s+SERVICIO\s+AFECTADO\s+([A-Z0-9]+)',
        r'ID\s+SERVICIO\s+([A-Z0-9]+)',
        r'\b(CAV\d{2}[A-Z]{2}\d{7,10})\b',
        r'\b([A-Z]{3}\d{2}[A-Z]{2}\d{7,10})\b',
    ],
    'coordenadas_multipoint': [
        r'COORDENADAS:\s*MultiPointZ\s*\(\s*\(\s*([-+]?\d+\.\d+)\s+([-+]?\d+\.\d+)',
        r'MultiPointZ\s*\(\s*\(\s*([-+]?\d+\.\d+)\s+([-+]?\d+\.\d+)',
    ],
    'lat_lon_decimal': [
        r'LATITUD:\s*([-+]?\s*\d+\.\d+)',
        r'LONGITUD:\s*([-+]?\s*\d+\.\d+)',
    ],
    'lat_lon_gms': [
        r'LATITUD:\s*([0-9°\'\"\.]+)',
        r'LONGITUD:\s*([0-9°\'\"\.]+)',
    ],
    'tipo_servicio': [
        r'TIPO DE SERVICIO:\s*([^\n]+)',
        r'TIPO DE SERVICIO\s+([^\n]+)',
    ],
    'puerto_common': [
        r'PUERTO COMMON:\s*([^\n]+)',
        r'PUERTO COMMON\s+([^\n]+)',
    ],
    'puerto_owner': [
        r'PUERTO OWNER:\s*([^\n]+)',
        r'PUERTO OWNER\s+([^\n]+)',
    ],
}


class CoordenadasFormatter:
    @staticmethod
    def a_grados_minutos_segundos(valor_decimal: float) -> str:
        es_negativo = valor_decimal < 0
        valor_abs = abs(valor_decimal)
        grados = int(valor_abs)
        minutos_decimal = (valor_abs - grados) * 60
        minutos = int(minutos_decimal)
        segundos = (minutos_decimal - minutos) * 60
        segundos_formateados = f"{segundos:.1f}"
        if es_negativo:
            return f"-{grados:02d}°{minutos:02d}'{segundos_formateados}\""
        return f"{grados:02d}°{minutos:02d}'{segundos_formateados}\""

    @staticmethod
    def formatear_coordenada_gms(coord_gms: str) -> str:
        if not coord_gms or coord_gms in ['NO', 'NO ENCONTRADA', 'ERROR']:
            return coord_gms
        coord = coord_gms.replace('"', '').strip()
        patron = r'(\d{1,3})[°]?(\d{2})\'(\d{1,3}(?:\.\d+)?)'
        match = re.search(patron, coord)
        if match:
            grados = match.group(1).zfill(2) if len(match.group(1)) < 3 else match.group(1).zfill(3)
            minutos = match.group(2)
            segundos = match.group(3)
            return f"{grados}°{minutos}'{segundos}\""
        return coord_gms


class DataExtractor:
    def __init__(self):
        self.formatter = CoordenadasFormatter()
        # Cargar equipos centrales válidos desde el Excel oficial
        self.equipos_validos = set()
        try:
            if RUTA_CENTRALES.exists():
                df_centrales = pd.read_excel(RUTA_CENTRALES, sheet_name='EC', engine='openpyxl')
                if 'cabecera' in df_centrales.columns:
                    self.equipos_validos = set(df_centrales['cabecera'].dropna().astype(str).str.strip())
        except Exception:
            pass

    def _extraer_token_de_linea(self, linea: str) -> Optional[str]:
        partes = linea.split(':', 1)
        if len(partes) < 2:
            return None
        resto = partes[1].strip()
        tokens = re.split(r'[\s\t]+', resto)
        for token in tokens:
            if token:
                return token
        return None

    def extraer_ip(self, texto: str) -> str:
        for patron in PATRONES['ip_demarcador']:
            match = re.search(patron, texto, re.IGNORECASE)
            if match:
                ip = match.group(1).strip()
                partes = ip.split('.')
                if len(partes) == 4 and all(p.isdigit() and 0 <= int(p) <= 255 for p in partes):
                    return ip
        
        for patron in PATRONES['ip_general']:
            match = re.search(patron, texto, re.IGNORECASE)
            if match:
                ip = match.group(1).strip()
                partes = ip.split('.')
                if len(partes) == 4 and all(p.isdigit() and 0 <= int(p) <= 255 for p in partes):
                    contexto_inicio = max(0, match.start() - 20)
                    contexto = texto[contexto_inicio:match.end() + 10]
                    if 'IP WAN' in contexto or '/30' in contexto:
                        continue
                    return ip
        return "SIN IP"

    def extraer_anillo(self, texto: str) -> str:
        for patron in PATRONES['anillo']:
            match = re.search(patron, texto, re.IGNORECASE)
            if match:
                anillo = match.group(1).strip().upper()
                anillo_limpio = re.sub(r'[^A-Z0-9]', '', anillo)
                if len(anillo_limpio) >= 8:
                    return anillo_limpio
        return "SIN ANILLO"

    def extraer_ancho_banda(self, texto: str) -> str:
        for patron in PATRONES['ancho_banda']:
            match = re.search(patron, texto, re.IGNORECASE)
            if match:
                valor = match.group(1).strip().strip(':,- ')
                numero_match = re.search(r'(\d+(?:\.\d+)?)', valor)
                if numero_match:
                    valor = numero_match.group(1)
                valor = re.sub(r'(?i)(mbps|mbit|megabite?)', '', valor).strip()
                if valor:
                    return valor
        return "N/A"

    def extraer_id_servicio(self, texto: str) -> str:
        for patron in PATRONES['id_servicio']:
            matches = re.finditer(patron, texto, re.IGNORECASE)
            for match in matches:
                id_serv = match.group(1).strip().strip(':,- ')
                if id_serv and len(id_serv) >= 10:
                    id_serv = re.sub(r'[^\w]', '', id_serv)
                    if id_serv:
                        return id_serv
        return "SIN ID SERVICIO"

    def extraer_tipo_servicio(self, texto: str) -> str:
        for patron in PATRONES['tipo_servicio']:
            match = re.search(patron, texto, re.IGNORECASE)
            if match:
                valor = match.group(1).strip().strip(':,- ')
                if valor:
                    return valor
        return "SIN TIPO SERVICIO"

    def extraer_equipo_central(self, texto: str) -> str:
        # 1. Sección RECURSOS METRO
        match = re.search(r'RECURSOS METRO.*?<([A-Z0-9]+)>', texto, re.IGNORECASE | re.DOTALL)
        if match:
            candidato = match.group(1)
            if candidato in self.equipos_validos:
                return candidato

        # 2. Líneas clave
        lineas_clave = ['METRO:', 'CABECERA:', 'PRINCIPAL:', 'ACCESO:']
        for linea in texto.splitlines():
            linea_upper = linea.upper()
            for clave in lineas_clave:
                if clave in linea_upper:
                    token = self._extraer_token_de_linea(linea)
                    if token and token in self.equipos_validos:
                        return token

        # 3. Búsqueda general
        posibles = re.findall(r'\b([A-Z]{2,}[A-Z0-9]{4,})\b', texto)
        for cand in posibles:
            if cand in self.equipos_validos:
                return cand

        return "SIN EQUIPO CENTRAL"

    def extraer_puerto_common(self, texto: str) -> str:
        # Primero buscar líneas explícitas "PUERTO COMMON:"
        for patron in PATRONES['puerto_common']:
            match = re.search(patron, texto, re.IGNORECASE)
            if match:
                valor = match.group(1).strip().strip(':,- ')
                if valor:
                    return self._limpiar_puerto(valor)

        # Buscar en sección RECURSOS METRO o en líneas que contengan "Common" y "GigabitEthernet/GE"
        patron_common = r'(?:COMMON|Common)[- ]*(?:GigabitEthernet|GE)[\s]*([0-9/]+)'
        match = re.search(patron_common, texto, re.IGNORECASE)
        if match:
            puerto = match.group(1).strip()
            return self._limpiar_puerto(puerto)

        # También buscar en líneas de "display erps ring" donde aparece "Common"
        patron_ring = r'(GE[0-9/]+)\s+Common'
        match = re.search(patron_ring, texto, re.IGNORECASE)
        if match:
            puerto = match.group(1).strip()
            return self._limpiar_puerto(puerto)

        return "SIN PUERTO COMMON"

    def extraer_puerto_owner(self, texto: str) -> str:
        # Primero buscar líneas explícitas "PUERTO OWNER:"
        for patron in PATRONES['puerto_owner']:
            match = re.search(patron, texto, re.IGNORECASE)
            if match:
                valor = match.group(1).strip().strip(':,- ')
                if valor:
                    return self._limpiar_puerto(valor)

        # Buscar en sección RECURSOS METRO o en líneas con "RPL OWNER"
        patron_owner = r'RPL\s+OWNER[- ]*(?:GigabitEthernet|GE)[\s]*([0-9/]+)'
        match = re.search(patron_owner, texto, re.IGNORECASE)
        if match:
            puerto = match.group(1).strip()
            return self._limpiar_puerto(puerto)

        # También buscar en líneas de "display erps ring" donde aparece "RPL Owner"
        patron_ring = r'(GE[0-9/]+)\s+RPL\s+Owner'
        match = re.search(patron_ring, texto, re.IGNORECASE)
        if match:
            puerto = match.group(1).strip()
            return self._limpiar_puerto(puerto)

        return "SIN PUERTO OWNER"

    def _limpiar_puerto(self, valor: str) -> str:
        if valor.startswith("SIN "):
            return valor
        valor_limpio = re.sub(r'\b(GigabitEthernet|GE)\s*', '', valor, flags=re.IGNORECASE).strip()
        return valor_limpio if valor_limpio else valor

    def extraer_coordenadas(self, texto: str) -> Tuple[Optional[str], Optional[str]]:
        for patron in PATRONES['coordenadas_multipoint']:
            match = re.search(patron, texto, re.IGNORECASE)
            if match:
                lon = float(match.group(1))
                lat = float(match.group(2))
                if lon > 0:
                    lon = -lon
                return (
                    self.formatter.a_grados_minutos_segundos(lat),
                    self.formatter.a_grados_minutos_segundos(lon)
                )

        match_lat = re.search(PATRONES['lat_lon_decimal'][0], texto, re.IGNORECASE)
        match_lon = re.search(PATRONES['lat_lon_decimal'][1], texto, re.IGNORECASE)
        if match_lat and match_lon:
            try:
                lat = float(match_lat.group(1).replace(' ', ''))
                lon = float(match_lon.group(1).replace(' ', ''))
                if lon > 0:
                    lon = -lon
                return (
                    self.formatter.a_grados_minutos_segundos(lat),
                    self.formatter.a_grados_minutos_segundos(lon)
                )
            except ValueError:
                pass

        match_lat_gms = re.search(PATRONES['lat_lon_gms'][0], texto, re.IGNORECASE)
        match_lon_gms = re.search(PATRONES['lat_lon_gms'][1], texto, re.IGNORECASE)
        if match_lat_gms and match_lon_gms:
            lat = self.formatter.formatear_coordenada_gms(match_lat_gms.group(1))
            lon = self.formatter.formatear_coordenada_gms(match_lon_gms.group(1))
            if lon and not lon.startswith('-'):
                lon = f"-{lon}"
            return lat, lon

        return None, None

    def extraer_oc(self, texto: str) -> str:
        """
        Extrae el número de OC incluyendo el prefijo 'OC-'.
        Ejemplo: 'OC-1785338' en lugar de solo '1785338'.
        """
        match = re.search(r'OC-\d+', texto, re.IGNORECASE)
        if match:
            return match.group(0)
        return "SIN OC"

    def extraer_etiqueta(self, texto: str) -> str:
        """
        Extrae la etiqueta que comienza con TO_ hasta el final de la línea.
        Prioriza líneas con 'ETIQUETA METRO:' o 'DESCRIPCION:'.
        """
        # 1. Buscar en línea ETIQUETA METRO:
        for linea in texto.splitlines():
            if 'ETIQUETA METRO:' in linea.upper():
                match = re.search(r'TO_.*$', linea)
                if match:
                    return match.group(0).strip()

        # 2. Buscar en líneas que contengan 'DESCRIPCION:' o 'description'
        for linea in texto.splitlines():
            if 'DESCRIPCION:' in linea.upper() or 'description' in linea.lower():
                match = re.search(r'TO_.*$', linea)
                if match:
                    return match.group(0).strip()

        # 3. Buscar en líneas con 'TO_DEM' o 'TO_' (cualquier línea)
        for linea in texto.splitlines():
            if 'TO_DEM' in linea or 'TO_' in linea:
                match = re.search(r'TO_.*$', linea)
                if match:
                    return match.group(0).strip()

        # 4. Búsqueda general en todo el texto (primera ocurrencia)
        match = re.search(r'TO_[^\s]+', texto)
        if match:
            return match.group(0)

        return "SIN ETIQUETA"

    def extraer_todos(self, texto: str) -> Dict[str, Any]:
        ip = self.extraer_ip(texto)
        anillo = self.extraer_anillo(texto)
        ancho_banda = self.extraer_ancho_banda(texto)
        id_servicio = self.extraer_id_servicio(texto)
        lat, lon = self.extraer_coordenadas(texto)
        tipo_servicio = self.extraer_tipo_servicio(texto)
        equipo_central = self.extraer_equipo_central(texto)
        puerto_common = self.extraer_puerto_common(texto)
        puerto_owner = self.extraer_puerto_owner(texto)
        oc = self.extraer_oc(texto)
        etiqueta = self.extraer_etiqueta(texto)

        if lat and lon:
            coordenadas = f"{lat} {lon}"
        else:
            coordenadas = ""

        return {
            'ip': ip,
            'anillo': anillo,
            'ancho_banda': ancho_banda,
            'id_servicio': id_servicio,
            'coordenadas': coordenadas,
            'latitud': lat,
            'longitud': lon,
            'tipo_servicio': tipo_servicio,
            'equipo_central': equipo_central,
            'puerto_common': puerto_common,
            'puerto_owner': puerto_owner,
            'oc': oc,
            'etiqueta': etiqueta,
        }


class FileProcessor:
    def __init__(self, ruta_origen: Path):
        self.ruta_origen = ruta_origen
        self.extractor = DataExtractor()
        self.estadisticas = {
            'total': 0,
            'con_coordenadas': 0,
            'con_ip': 0,
            'con_anillo': 0,
            'con_ancho_banda': 0,
            'con_id_servicio': 0,
            'errores': 0,
        }

    def obtener_archivos(self) -> List[Path]:
        if not self.ruta_origen.exists():
            return []
        return [f for f in self.ruta_origen.iterdir() if f.suffix.lower() == EXTENSION]

    def procesar_archivo(self, archivo: Path) -> Optional[List[Any]]:
        try:
            with open(archivo, 'r', encoding='utf-8', errors='ignore') as f:
                texto = f.read()
        except Exception:
            self.estadisticas['errores'] += 1
            return None
        
        datos = self.extractor.extraer_todos(texto)
        nombre_archivo = archivo.stem

        self.estadisticas['total'] += 1
        if datos['coordenadas']:
            self.estadisticas['con_coordenadas'] += 1
        if datos['ip'] != "SIN IP":
            self.estadisticas['con_ip'] += 1
        if datos['anillo'] != "SIN ANILLO":
            self.estadisticas['con_anillo'] += 1
        if datos['ancho_banda'] != "N/A":
            self.estadisticas['con_ancho_banda'] += 1
        if datos['id_servicio'] != "SIN ID SERVICIO":
            self.estadisticas['con_id_servicio'] += 1

        # Devolvemos en el orden que luego reordenaremos en procesar_todos
        return {
            'Archivo': nombre_archivo,
            'ID_Servicio': datos['id_servicio'],
            'Anillo': datos['anillo'],
            'IP': datos['ip'],
            'ancho_banda': datos['ancho_banda'],
            'Coordenadas': datos['coordenadas'],
            'Tipo_Servicio': datos['tipo_servicio'],
            'Equipo_Central': datos['equipo_central'],
            'Puerto_Common': datos['puerto_common'],
            'Puerto_Owner': datos['puerto_owner'],
            'OC': datos['oc'],
            'Etiqueta': datos['etiqueta'],
        }

    def procesar_todos(self) -> pd.DataFrame:
        archivos = self.obtener_archivos()
        if not archivos:
            return pd.DataFrame()

        filas = []
        for archivo in archivos:
            datos = self.procesar_archivo(archivo)
            if datos:
                filas.append(datos)

        df = pd.DataFrame(filas)

        # Reordenar columnas según lo solicitado:
        nuevo_orden = [
            'Equipo_Central',
            'Anillo',
            'Etiqueta',
            'Puerto_Common',
            'Puerto_Owner',
            'ID_Servicio',
            'IP',
            'ancho_banda',
            'Coordenadas',
            'Tipo_Servicio',
            'OC',
            'Archivo'
        ]
        # Asegurar que todas las columnas existan
        columnas_existentes = [col for col in nuevo_orden if col in df.columns]
        # Agregar cualquier columna que falte en el orden al final
        resto = [col for col in df.columns if col not in columnas_existentes]
        df = df[columnas_existentes + resto]

        return df


class ExcelWriter:
    def __init__(self, ruta_destino: Path):
        self.ruta_destino = ruta_destino
        self.ruta_destino.mkdir(parents=True, exist_ok=True)

    def guardar(self, df: pd.DataFrame) -> Optional[Path]:
        if df.empty:
            return None

        ruta_completa = self.ruta_destino / NOMBRE_SALIDA

        try:
            with pd.ExcelWriter(ruta_completa, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='DATOS')
                hoja = writer.sheets['DATOS']
                for columna in hoja.columns:
                    max_len = 0
                    col_letra = columna[0].column_letter
                    for celda in columna:
                        try:
                            max_len = max(max_len, len(str(celda.value)))
                        except:
                            pass
                    hoja.column_dimensions[col_letra].width = min(max_len + 2, 50)
            return ruta_completa
        except Exception:
            ruta_csv = self.ruta_destino / f"{NOMBRE_SALIDA}.csv"
            df.to_csv(ruta_csv, index=False, encoding='utf-8-sig')
            return ruta_csv


def main():
    if not RUTA_ORIGEN.exists():
        print("Ruta de origen no existe.")
        return

    processor = FileProcessor(RUTA_ORIGEN)
    df = processor.procesar_todos()

    if df.empty:
        print("No se encontraron datos.")
        return

    writer = ExcelWriter(RUTA_DESTINO)
    archivo_guardado = writer.guardar(df)
    if archivo_guardado:
        print(f"Archivo guardado en: {archivo_guardado}")
    else:
        print("Error al guardar el archivo.")


if __name__ == "__main__":
    main()