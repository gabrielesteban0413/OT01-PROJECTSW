"sw - users"

import pyautogui
import pyperclip
import win32com.client as win32
import win32gui
import win32con
import time
import keyboard
import re

NOMBRE_HOJA = "CD"
COLUMNA_ID = "ID"
COLUMNA_OWNER = "OWNER"
COLUMNA_ELEMENT = "ELEMENT"

COORD_BORRAR = (1125, 99)
COORD_BUSCAR = (1197, 97)
COORD_PEGAR_ID = (1095, 212)
COORD_INICIO_LISTA = (766, 203)
COORD_CERRAR_POPUP = (666, 377)  

pyautogui.PAUSE = 0.0

pausado = False
detener = False

POPUP_KEYWORDS = ["información", "mensaje", "error", "advertencia", "confirmar", "warning", "message", "info"]

def extraer_primer_id(id_texto):
    if ',' in str(id_texto):
        primer_id = str(id_texto).split(',')[0].strip()
        return primer_id
    return str(id_texto).strip()

def verificar_y_cerrar_popup():
    def enum_callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and any(keyword in title.lower() for keyword in POPUP_KEYWORDS):
                extra.append(hwnd)
        return True
    ventanas = []
    win32gui.EnumWindows(enum_callback, ventanas)
    if ventanas:
        hwnd = ventanas[0]
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.05)
        pyautogui.press('enter')
        time.sleep(0.1)
        print("[POPUP] Cerrado con Enter")
        return True
    return False

def cerrar_popup_no_resultados():
    def enum_windows_callback(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and ("no port records" in title.lower()):
                results.append(hwnd)
                return True
            def enum_child(hwnd_child, child_results):
                child_text = win32gui.GetWindowText(hwnd_child)
                if "no port records" in child_text.lower():
                    child_results.append(hwnd_child)
            child_windows = []
            win32gui.EnumChildWindows(hwnd, enum_child, child_windows)
            if child_windows:
                results.append(hwnd)
        return True
    ventanas = []
    win32gui.EnumWindows(enum_windows_callback, ventanas)
    if ventanas:
        hwnd = ventanas[0]
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.1)
        pyautogui.click(COORD_CERRAR_POPUP[0], COORD_CERRAR_POPUP[1])
        time.sleep(0.2)
        print("[POPUP] Ventana 'no encontrado' cerrada con clic")
        return True
    return False

def pausar_o_reanudar():
    global pausado
    pausado = not pausado
    print(f"\n[CONTROL] {'PAUSADO' if pausado else 'REANUDADO'}")

def detener_proceso():
    global detener
    detener = True
    print("\n[CONTROL] DETENIENDO")

keyboard.add_hotkey('0', pausar_o_reanudar)
keyboard.add_hotkey('1', detener_proceso)

def esperar_si_pausado():
    global detener, pausado
    while pausado and not detener:
        time.sleep(0.05)
        verificar_y_cerrar_popup()
        if keyboard.is_pressed('q'):
            detener = True
            break

def esperar_con_control(segundos):
    inicio = time.time()
    while time.time() - inicio < segundos:
        if detener:
            return False
        esperar_si_pausado()
        verificar_y_cerrar_popup()
        time.sleep(0.05)
    return not detener

def obtener_excel_abierto():
    try:
        excel = win32.GetActiveObject("Excel.Application")
        wb = excel.ActiveWorkbook
        print(f"Excel: {wb.Name}")
        # Suprimir alertas y eventos para evitar diálogos
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        return excel, wb
    except:
        print("No se pudo conectar a Excel. Asegúrate de tenerlo abierto.")
        input("Presiona ENTER cuando Excel esté listo...")
        try:
            excel = win32.GetActiveObject("Excel.Application")
            wb = excel.ActiveWorkbook
            print(f"Excel conectado: {wb.Name}")
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            return excel, wb
        except:
            return None, None

def fila_visible(hoja, fila):
    try:
        return not hoja.Rows(fila).Hidden
    except:
        return True

def tiene_formato_fecha(texto):
    patron = r'^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}'
    return re.match(patron, texto) is not None

def obtener_ultimo_valor_lista():
    if detener:
        return ""
    esperar_si_pausado()
    pyautogui.click(COORD_INICIO_LISTA[0], COORD_INICIO_LISTA[1])
    if not esperar_con_control(0.1): return ""
    pyautogui.hotkey('ctrl', 'end')
    if not esperar_con_control(0.2): return ""
    pyautogui.hotkey('ctrl', 'c')
    if not esperar_con_control(0.1): return ""
    valor = pyperclip.paste().strip()
    if tiene_formato_fecha(valor):
        if '/' in valor:
            return valor
        else:
            return "|"
    pyautogui.click(COORD_INICIO_LISTA[0], COORD_INICIO_LISTA[1])
    if not esperar_con_control(0.1): return ""
    ultimo_valor = ""
    for _ in range(20):
        if detener: return ""
        pyautogui.press('end')
        if not esperar_con_control(0.1): return ""
        pyautogui.hotkey('ctrl', 'c')
        if not esperar_con_control(0.1): return ""
        valor_actual = pyperclip.paste().strip()
        if valor_actual == ultimo_valor: break
        ultimo_valor = valor_actual
        if tiene_formato_fecha(valor_actual):
            if '/' in valor_actual:
                return valor_actual
            else:
                return "|"
    pyautogui.click(COORD_INICIO_LISTA[0], COORD_INICIO_LISTA[1])
    if not esperar_con_control(0.1): return ""
    ultimo_valor = ""
    for _ in range(20):
        if detener: return ""
        pyautogui.press('pagedown')
        if not esperar_con_control(0.2): return ""
        pyautogui.hotkey('ctrl', 'c')
        if not esperar_con_control(0.1): return ""
        valor_actual = pyperclip.paste().strip()
        if valor_actual == ultimo_valor: break
        ultimo_valor = valor_actual
        if tiene_formato_fecha(valor_actual):
            if '/' in valor_actual:
                return valor_actual
            else:
                return "|"
    return ultimo_valor if ultimo_valor and '/' in ultimo_valor else "|"

def limpiar_campo():
    if detener: return
    esperar_si_pausado()
    pyautogui.click(COORD_BORRAR[0], COORD_BORRAR[1])
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')

def cerrar_dialogos_excel(excel):
    """Cierra cualquier cuadro de diálogo que Excel pueda mostrar"""
    try:
        # Si hay un cuadro de diálogo modal, SendKeys Esc suele cerrarlo
        excel.SendKeys("{ESC}")
        time.sleep(0.2)
    except:
        pass

def escribir_celda_seguro(hoja, fila, columna, valor, excel, wb, reintentos=3):
    """Intenta escribir en la celda con reintentos y reconexión si falla"""
    global detener
    for intento in range(reintentos):
        if detener:
            return False
        try:
            hoja.Cells(fila, columna).Value = valor
            return True
        except Exception as e:
            print(f"  Error al escribir (intento {intento+1}/{reintentos}): {e}")
            if intento == reintentos - 1:
                raise
            # Cerrar posibles diálogos de Excel
            cerrar_dialogos_excel(excel)
            time.sleep(1)
            # Reintentar sin reconectar primero
            continue
    return False

def procesar_fila(hoja, fila, id_texto, idx_owner, total, indice, es_prueba=False, excel=None, wb=None):
    global detener
    if detener:
        return False

    id_texto = extraer_primer_id(id_texto)
    print(f"[{indice}/{total}] ID {id_texto} (fila {fila})", end=" ")
    pyperclip.copy(id_texto)

    esperar_si_pausado()
    if detener: return False

    pyautogui.hotkey('alt', 'tab')
    if not esperar_con_control(0.15): return False

    limpiar_campo()
    verificar_y_cerrar_popup()
    if detener: return False

    pyautogui.click(COORD_PEGAR_ID[0], COORD_PEGAR_ID[1])
    pyautogui.hotkey('ctrl', 'v')
    if not esperar_con_control(0.5): return False

    pyautogui.click(COORD_BUSCAR[0], COORD_BUSCAR[1])
    time.sleep(0.5)

    if cerrar_popup_no_resultados():
        print("-> ID NO EXISTE")
        ultimo_valor = "|"
    else:
        ultimo_valor = obtener_ultimo_valor_lista()
        verificar_y_cerrar_popup()

    if not ultimo_valor or '/' not in ultimo_valor:
        ultimo_valor = "|"

    if es_prueba:
        print(f"\n   Capturado: '{ultimo_valor}'")
        opcion = input("   s=guardar, n=omitir, m=escribir manual: ").lower()
        if opcion == 'n':
            print("   omitida")
            return False
        elif opcion == 'm':
            ultimo_valor = input("   Escribe valor: ").strip()
    else:
        print("-> OK", end=" ")

    pyautogui.hotkey('alt', 'tab')
    if not esperar_con_control(0.1): return False

    # Escribir con reintentos y manejo de errores
    try:
        escribir_celda_seguro(hoja, fila, idx_owner, ultimo_valor, excel, wb)
    except Exception as e:
        print(f"\n  ERROR CRÍTICO al escribir en Excel: {e}")
        print("  Se recomienda pausar (0) y revisar si Excel tiene algún diálogo abierto.")
        return False

    if not es_prueba:
        print(f"('{ultimo_valor}')")
    else:
        print(f"   OWNER actualizado: '{ultimo_valor}'")
    return True

def procesar_excel():
    global detener, pausado
    detener = False
    pausado = False

    excel, wb = obtener_excel_abierto()
    if not excel or not wb:
        return

    try:
        hoja = wb.Sheets(NOMBRE_HOJA)
    except:
        print(f"Hoja '{NOMBRE_HOJA}' no existe.")
        return

    encabezados = [hoja.Cells(1, col).Value for col in range(1, 100) if hoja.Cells(1, col).Value]
    if COLUMNA_ID not in encabezados or COLUMNA_OWNER not in encabezados or COLUMNA_ELEMENT not in encabezados:
        print("No se encontraron las columnas requeridas")
        return

    idx_id = encabezados.index(COLUMNA_ID) + 1
    idx_owner = encabezados.index(COLUMNA_OWNER) + 1
    idx_element = encabezados.index(COLUMNA_ELEMENT) + 1

    print("Buscando primera fila visible con OWNER vacío...")
    primera_fila = None
    fila = 2
    while True:
        id_valor = hoja.Cells(fila, idx_id).Value
        if id_valor is None:
            break
        if fila_visible(hoja, fila) and (hoja.Cells(fila, idx_owner).Value is None or str(hoja.Cells(fila, idx_owner).Value).strip() == ""):
            primera_fila = fila
            break
        fila += 1

    if primera_fila is None:
        print("No se encontró ninguna fila visible con OWNER vacío.")
        return

    valor_element_ref = str(hoja.Cells(primera_fila, idx_element).Value).strip()
    print(f"Primera fila visible: {primera_fila}, ELEMENT = '{valor_element_ref}'")

    print("Recopilando filas del bloque...")
    filas_a_procesar = []
    fila = primera_fila
    while True:
        id_valor = hoja.Cells(fila, idx_id).Value
        if id_valor is None:
            break
        if fila_visible(hoja, fila):
            element_valor = str(hoja.Cells(fila, idx_element).Value).strip()
            if element_valor != valor_element_ref:
                break
            if hoja.Cells(fila, idx_owner).Value is None or str(hoja.Cells(fila, idx_owner).Value).strip() == "":
                filas_a_procesar.append((fila, str(id_valor)))
        fila += 1

    if not filas_a_procesar:
        print("No hay filas visibles para procesar en este bloque.")
        return

    total = len(filas_a_procesar)
    print(f"\nTotal de filas a procesar: {total}")
    print("[CONTROL] 0=pausa, 1=detener\n")



    for i, (f, txt) in enumerate(filas_a_procesar, start=1):
        if detener: break
        esperar_si_pausado()
        if detener: break
        procesar_fila(hoja, f, txt, idx_owner, total, i, es_prueba=False, excel=excel, wb=wb)

    # Restaurar configuraciones de Excel
    try:
        excel.DisplayAlerts = True
        excel.EnableEvents = True
    except:
        pass
    print("\nProceso finalizado.")

if __name__ == "__main__":
    procesar_excel()