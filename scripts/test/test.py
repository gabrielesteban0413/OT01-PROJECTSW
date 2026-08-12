#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os

def transformar_linea(linea):
    """
    Transforma una línea:
    - Si contiene '°' en la parte anterior al primer '|', la divide en varios IDs
      y genera una línea por cada ID con el resto de los campos.
    - Si no tiene '°', devuelve la línea original.
    - Si la línea empieza con '|' (sin ID), la devuelve tal cual.
    """
    linea = linea.rstrip('\n')
    if not linea:
        return []

    # Buscar el primer pipe
    pos_pipe = linea.find('|')
    if pos_pipe == -1:
        # Línea sin pipe: se devuelve tal cual
        return [linea]

    ids_part = linea[:pos_pipe]          # parte antes del primer '|'
    resto = linea[pos_pipe:]             # desde el primer '|' inclusive

    if '°' not in ids_part:
        # Un solo ID (o ninguno si ids_part está vacío)
        return [linea]

    # Dividir los IDs por '°'
    ids = [id.strip() for id in ids_part.split('°') if id.strip()]
    if not ids:
        return [linea]

    # Generar una línea por cada ID
    lineas_salida = []
    for id_actual in ids:
        lineas_salida.append(id_actual + resto)

    return lineas_salida


def main():
    # Ruta del archivo de entrada
    archivo_entrada = r"C:\A_GS1_PROYECTOS\a1.txt"
    # Ruta del archivo de salida (puedes cambiarla)
    archivo_salida = r"C:\A_GS1_PROYECTOS\a1_procesado.txt"

    # Verificar que el archivo de entrada existe
    if not os.path.isfile(archivo_entrada):
        print(f"Error: No se encontró el archivo '{archivo_entrada}'")
        sys.exit(1)

    # Contador de líneas procesadas
    contador = 0

    try:
        with open(archivo_entrada, 'r', encoding='utf-8') as f_in, \
             open(archivo_salida, 'w', encoding='utf-8') as f_out:

            for linea in f_in:
                lineas_transformadas = transformar_linea(linea)
                for nueva_linea in lineas_transformadas:
                    f_out.write(nueva_linea + '\n')
                    contador += 1

        print(f"✅ Procesamiento completado.")
        print(f"📁 Archivo de entrada: {archivo_entrada}")
        print(f"📁 Archivo de salida: {archivo_salida}")
        print(f"📊 Total de líneas generadas: {contador}")

    except Exception as e:
        print(f"❌ Ocurrió un error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()