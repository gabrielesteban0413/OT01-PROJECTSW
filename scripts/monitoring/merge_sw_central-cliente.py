import pandas as pd
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# Rutas
RUTA_PORTS = Path(r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\01-Bulk export of ports.xlsx")
RUTA_SALIDA = RUTA_PORTS.parent / "01-Bulk Merge_Smallworld.xlsx"

# Prefijos centrales
PREFIJOS_CENTRALES = ("OT", "OI", "O", "PP")

def es_cliente(row, col_desc_shelf, col_type):
    if col_type and pd.notna(row[col_type]):
        if str(row[col_type]).strip().lower() == "equipo cliente":
            return True
    if col_desc_shelf and pd.notna(row[col_desc_shelf]):
        desc = str(row[col_desc_shelf]).strip().upper()
        if not desc.startswith(PREFIJOS_CENTRALES):
            return True
    return False

def main():
    if not RUTA_PORTS.exists():
        print(f"Error: Archivo no encontrado: {RUTA_PORTS}")
        return

    try:
        df_portcard = pd.read_excel(RUTA_PORTS, sheet_name="PortCard", dtype=str)
        df_port = pd.read_excel(RUTA_PORTS, sheet_name="Port", dtype=str)
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        return

    # ========== PortCard ==========
    col_id_pc = "Id"
    col_type_pc = "Type"
    col_proyecto_pc = "Proyecto De Red"
    col_equipo_pc = "Equipo Central"
    col_puerto_pc = "Puerto Central"
    col_desc_pc = "Description"

    for col in [col_id_pc, col_type_pc, col_proyecto_pc, col_equipo_pc, col_puerto_pc, col_desc_pc]:
        if col not in df_portcard.columns:
            print(f"Error: Columna '{col}' no encontrada en PortCard")
            return

    df_portcard = df_portcard[df_portcard[col_proyecto_pc].notna()]
    df_portcard = df_portcard[df_portcard[col_proyecto_pc].astype(str).str.strip() != '']
    df_portcard = df_portcard[df_portcard[col_proyecto_pc].astype(str).str.strip() != '0']

    central_por_proyecto = {}

    for _, row in df_portcard.iterrows():
        proyecto = str(row[col_proyecto_pc]).strip()
        tipo = str(row[col_type_pc]).strip().lower()
        id_val = str(row[col_id_pc]).strip()
        puerto = str(row[col_puerto_pc]).strip() if pd.notna(row[col_puerto_pc]) else ""
        desc = str(row[col_desc_pc]).strip() if pd.notna(row[col_desc_pc]) else ""
        equipo_nombre = str(row[col_equipo_pc]).strip() if pd.notna(row[col_equipo_pc]) else ""

        if proyecto not in central_por_proyecto:
            central_por_proyecto[proyecto] = {
                "equipos_centrales": [],
                "odf_pacheos": []
            }

        if tipo == "equipo central":
            central_por_proyecto[proyecto]["equipos_centrales"].append({
                "id": id_val,
                "puerto": puerto,
                "desc": desc,
                "equipo": equipo_nombre
            })
        elif tipo == "odf pacheo":
            central_por_proyecto[proyecto]["odf_pacheos"].append({
                "id": id_val,
                "desc": desc
            })

    # ========== Port ==========
    col_id_p = "Id"
    col_type_p = "Type"
    col_proyecto_p = "Proyecto De Red"
    col_puerto_p = "Puerto Central"
    col_nombre_cliente = "Nombre Cliente"
    col_desc_shelf = "Description Shelf"
    col_equipo_cliente = "Equipo Central"   # esta columna existe en Port

    for col in [col_id_p, col_type_p, col_proyecto_p, col_puerto_p, col_nombre_cliente, col_desc_shelf, col_equipo_cliente]:
        if col not in df_port.columns:
            print(f"Error: Columna '{col}' no encontrada en Port")
            return

    df_port = df_port[df_port[col_proyecto_p].notna()]
    df_port = df_port[df_port[col_proyecto_p].astype(str).str.strip() != '']
    df_port = df_port[df_port[col_proyecto_p].astype(str).str.strip() != '0']

    odf_troncal_por_proyecto = {}

    for _, row in df_port.iterrows():
        tipo = str(row[col_type_p]).strip().lower()
        if tipo == "odf":
            desc_shelf = str(row[col_desc_shelf]).strip() if pd.notna(row[col_desc_shelf]) else ""
            if desc_shelf.upper().startswith("OT"):
                proyecto = str(row[col_proyecto_p]).strip()
                id_val = str(row[col_id_p]).strip()
                if proyecto not in odf_troncal_por_proyecto:
                    odf_troncal_por_proyecto[proyecto] = []
                odf_troncal_por_proyecto[proyecto].append({
                    "id": id_val,
                    "desc": desc_shelf
                })

    # Extraer clientes (incluyendo su Equipo Central)
    clientes = []
    for _, row in df_port.iterrows():
        if es_cliente(row, col_desc_shelf, col_type_p):
            proyecto = str(row[col_proyecto_p]).strip()
            if not proyecto or proyecto == "0":
                continue
            id_cliente = str(row[col_id_p]).strip() if pd.notna(row[col_id_p]) else ""
            puerto_cliente = str(row[col_puerto_p]).strip() if pd.notna(row[col_puerto_p]) else ""
            nombre_cliente = str(row[col_nombre_cliente]).strip() if pd.notna(row[col_nombre_cliente]) else ""
            equipo_cliente = str(row[col_equipo_cliente]).strip() if pd.notna(row[col_equipo_cliente]) else ""
            clientes.append({
                "proyecto": proyecto,
                "id_cliente": id_cliente,
                "puerto_cliente": puerto_cliente,
                "nombre_cliente": nombre_cliente,
                "equipo_cliente": equipo_cliente
            })

    if not clientes:
        print("No se encontraron clientes con proyectos válidos.")
        return

    # ========== Construir salida ==========
    registros = []

    for cliente in clientes:
        proyecto = cliente["proyecto"]
        central = central_por_proyecto.get(proyecto, {})
        equipos = central.get("equipos_centrales", [])
        pacheos = central.get("odf_pacheos", [])
        troncales = odf_troncal_por_proyecto.get(proyecto, [])

        ids_ec = "|".join([e["id"] for e in equipos]) if equipos else ""
        puertos_ec = "|".join([e["puerto"] for e in equipos]) if equipos else ""
        nombre_ec = equipos[0]["equipo"] if equipos else ""

        ids_pacheo = "|".join([p["id"] for p in pacheos]) if pacheos else ""
        desc_pacheo = "|".join([p["desc"] for p in pacheos]) if pacheos else ""

        ids_troncal = "|".join([t["id"] for t in troncales]) if troncales else ""
        desc_troncal = "|".join([t["desc"] for t in troncales]) if troncales else ""

        # Anomalía: si no hay equipos centrales, usar el equipo del cliente
        if not equipos:
            equipo_cliente = cliente["equipo_cliente"]
            if equipo_cliente:
                anomalia = f"no existe equipo central (posible {equipo_cliente})"
            else:
                anomalia = "no existe equipo central (posible desconocido)"
        else:
            anomalia = ""

        registros.append({
            "Proyecto De Red": proyecto,
            "ID_Cliente": cliente["id_cliente"],
            "Puerto_Cliente": cliente["puerto_cliente"],
            "Nombre_Cliente": cliente["nombre_cliente"],
            "Equipo_Central": nombre_ec,
            "Puerto_Central_EC": puertos_ec,
            "ID_Equipo_Central": ids_ec,
            "ID_ODF_Pacheo": ids_pacheo,
            "Desc_ODF_Pacheo": desc_pacheo,
            "ID_ODF_Troncal": ids_troncal,
            "Desc_ODF_Troncal": desc_troncal,
            "Anomalias": anomalia
        })

    df_salida = pd.DataFrame(registros)

    columnas_orden = [
        "Proyecto De Red",
        "ID_Cliente", "Puerto_Cliente", "Nombre_Cliente",
        "Equipo_Central", "Puerto_Central_EC", "ID_Equipo_Central",
        "ID_ODF_Pacheo", "Desc_ODF_Pacheo",
        "ID_ODF_Troncal", "Desc_ODF_Troncal",
        "Anomalias"
    ]
    for col in columnas_orden:
        if col not in df_salida.columns:
            df_salida[col] = ""
    df_salida = df_salida[columnas_orden]

    try:
        with pd.ExcelWriter(RUTA_SALIDA, engine='openpyxl') as writer:
            df_salida.to_excel(writer, sheet_name="Smallworld", index=False)
        print(f"Archivo generado: {RUTA_SALIDA}")
    except Exception as e:
        print(f"Error al guardar: {e}")

if __name__ == "__main__":
    main()