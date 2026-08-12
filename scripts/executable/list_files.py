import os

# Ruta de la carpeta
ruta = r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\12_Cables_troncales_ conectados_en central"

try:

    if os.path.exists(ruta):
        archivos = [f for f in os.listdir(ruta) if os.path.isfile(os.path.join(ruta, f))]  
        print("Archivos encontrados:")
        for archivo in archivos:
            print(archivo)
        
        print(f"\nTotal de archivos: {len(archivos)}")
    else:
        print(f"La ruta no existe: {ruta}")
        
except PermissionError:
    print("Error: No tienes permisos para acceder a esta carpeta")
except FileNotFoundError:
    print("Error: La ruta no se encuentra disponible")
except Exception as e:
    print(f"Error inesperado: {e}")