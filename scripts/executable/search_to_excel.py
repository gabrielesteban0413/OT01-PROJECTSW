import os
import pandas as pd

directory_path = r"\\10.112.71.143\Smallworld"
palabras_clave = ["ds_date","date_created", "sys!change_info","creation_time"] 
extensiones_validas = ('.magik', '.keymap', '.xml','.dat', 'txt','dmp')

if not os.path.exists(directory_path):
    print(f"La ruta {directory_path} no existe o no es accesible.")
else:
    search_results = []

    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith(extensiones_validas):
                file_path = os.path.join(root, file)
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as archivo:
                        lineas = archivo.readlines()
                        for num_linea, linea in enumerate(lineas, 1):
                            linea_limpia = linea.strip().lower()
                            if any(palabra in linea_limpia for palabra in palabras_clave):
                                search_results.append({
                                    'File Path': file_path,
                                    'Line Number': num_linea,
                                    'Line Content': linea.strip()
                                })
                except Exception as e:
                    print(f"No se pudo leer el archivo {file_path}: {e}")


    if not search_results:
        print("No se encontraron ocurrencias relacionadas.")
    else:
        df = pd.DataFrame(search_results, columns=["File Path", "Line Number", "Line Content"])
        output_path = r"C:\A_GS1_PROYECTOS\reporte2.xlsx"
        df.to_excel(output_path, index=False)
        print(f"Archivo Excel guardado en: {output_path}")







# para buscar en todos los archivos :  

import os
import re
import mmap
import pandas as pd
from multiprocessing import Pool, cpu_count
from functools import partial

# Configuración
directory_path = r"\\10.112.71.143\Smallworld"
palabras_clave = ["Conn", "Connection", "traze", "connected"]

# Compilar expresión regular en bytes (insensible a mayúsculas/minúsculas)
pattern_bytes = re.compile(
    b'|'.join(re.escape(palabra.encode('utf-8', errors='ignore')) for palabra in palabras_clave),
    re.IGNORECASE
)

def buscar_en_archivo(file_path, pattern):
    """Busca el patrón en un archivo usando mmap. Retorna lista de coincidencias."""
    resultados = []
    try:
        with open(file_path, 'rb') as f:
            with mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ) as mm:
                for match in re.finditer(pattern, mm):
                    start = mm.rfind(b'\n', 0, match.start()) + 1
                    end = mm.find(b'\n', match.end())
                    if end == -1:
                        end = len(mm)
                    linea = mm[start:end].decode('utf-8', errors='ignore').strip()
                    num_linea = mm[:start].count(b'\n') + 1
                    resultados.append({
                        'File Path': file_path,
                        'Line Number': num_linea,
                        'Line Content': linea
                    })
    except Exception:
        # Ignorar archivos que no se pueden leer (permisos, bloqueos, etc.)
        pass
    return resultados

def procesar_archivo(file_path, pattern):
    return buscar_en_archivo(file_path, pattern)

def obtener_lista_archivos(ruta):
    """Recorre recursivamente todos los archivos usando os.scandir (más rápido)."""
    archivos = []
    try:
        with os.scandir(ruta) as it:
            for entry in it:
                if entry.is_file():
                    archivos.append(entry.path)
                elif entry.is_dir():
                    archivos.extend(obtener_lista_archivos(entry.path))
    except PermissionError:
        # Ignorar directorios sin permisos
        pass
    return archivos

if __name__ == '__main__':
    if not os.path.exists(directory_path):
        print(f"La ruta {directory_path} no existe o no es accesible.")
    else:
        print("Recorriendo directorios...")
        archivos = obtener_lista_archivos(directory_path)
        print(f"Se encontraron {len(archivos)} archivos. Procesando en paralelo...")

        num_procesos = max(1, cpu_count() - 1)
        with Pool(processes=num_procesos) as pool:
            func = partial(procesar_archivo, pattern=pattern_bytes)
            resultados_por_archivo = pool.map(func, archivos)

        # Aplanar resultados
        search_results = []
        for res in resultados_por_archivo:
            search_results.extend(res)

        if not search_results:
            print("No se encontraron ocurrencias relacionadas.")
        else:
            df = pd.DataFrame(search_results, columns=["File Path", "Line Number", "Line Content"])
            output_path = r"C:\A_GS1_PROYECTOS\reporte2.xlsx"
            df.to_excel(output_path, index=False)
            print(f"Archivo Excel guardado en: {output_path} ({len(search_results)} coincidencias)")

            