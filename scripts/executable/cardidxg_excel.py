import os
import pandas as pd
import subprocess
import tempfile
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

GS_BASE_CARDI = os.getenv("GS_BASE_CARDI")
GS_BASE_PATH = os.getenv("GS_BASE_PATH")

if not GS_BASE_CARDI:
    raise ValueError("Variable GS_BASE_CARDI no encontrada en .env")
if not GS_BASE_PATH:
    raise ValueError("Variable GS_BASE_PATH no encontrada en .env")

ruta_base = Path(GS_BASE_CARDI)
ruta_salida = Path(GS_BASE_PATH)
ruta_salida.mkdir(parents=True, exist_ok=True)

def obtener_archivos_dwg_con_metadatos():
    temp_csv = tempfile.NamedTemporaryFile(delete=False, suffix='.csv', mode='w', encoding='utf-8')
    temp_csv.close()
    
    ps_command = f"""
    $results = Get-ChildItem -Path "{ruta_base}" -Recurse -File -Filter "*.dwg" -ErrorAction SilentlyContinue | ForEach-Object {{
        [PSCustomObject]@{{
            ARCHIVO = $_.Name
            RUTA_COMPLETA = $_.FullName
            RUTA_CARPETA = $_.DirectoryName
            TAMAÑO_MB = [math]::Round($_.Length / 1MB, 2)
            FECHA_MODIFICACION = $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            FECHA_CREACION = $_.CreationTime.ToString('yyyy-MM-dd HH:mm:ss')
        }}
    }}
    $results | Export-Csv -Path "{temp_csv.name}" -NoTypeInformation -Encoding UTF8
    """
    try:
        subprocess.run(["powershell", "-Command", ps_command], check=True, capture_output=True, text=True)
        df = pd.read_csv(temp_csv.name, encoding='utf-8')
        os.unlink(temp_csv.name)
        return df
    except subprocess.CalledProcessError as e:
        print(f"Error en PowerShell: {e.stderr}")
        return pd.DataFrame()

def crear_buscador_batch(df):
    buscador_path = ruta_salida / "BUSCADOR_DWG.bat"
    with open(buscador_path, 'w', encoding='utf-8') as f:
        f.write("@echo off\n")
        f.write("title BUSCADOR DE ARCHIVOS DWG\n")
        f.write("echo Escribe parte del nombre del archivo DWG\n")
        f.write("set /p busqueda=Buscar: \n")
        f.write("set \"tempfile=%TEMP%\\dwg_resultados.txt\"\n")
        f.write(f'> \"%%tempfile%%\" echo Resultados para: %%busqueda%%\n')
        f.write(f'>> \"%%tempfile%%\" echo.\n')
        
        contador = 1
        for idx, row in df.iterrows():
            archivo = row['ARCHIVO']
            ruta_completa = row['RUTA_COMPLETA']
            ruta_carpeta = row['RUTA_CARPETA']
            nombre_sin_extension = archivo.replace('.dwg', '').replace('.DWG', '')
            
            f.write(f'echo "%nombre_sin_extension%" | findstr /i "%%busqueda%%" > nul\n')
            f.write(f'if not errorlevel 1 (\n')
            f.write(f'  echo [%%contador%%] {archivo} >> "%%tempfile%%"\n')
            f.write(f'  echo     {ruta_carpeta} >> "%%tempfile%%"\n')
            f.write(f'  echo. >> "%%tempfile%%"\n')
            f.write(f'  set "archivo_%%contador%%={ruta_completa}"\n')
            f.write(f'  set /a contador+=1\n')
            f.write(f')\n')
        
        f.write('type "%tempfile%" && del "%tempfile%" 2>nul\n')
        f.write('echo.\n')
        f.write('if %contador%==1 (echo No se encontraron archivos & pause & exit)\n')
        f.write('set /p opcion=Selecciona el numero: \n')
        
        for idx, row in df.iterrows():
            ruta_completa = row['RUTA_COMPLETA']
            f.write(f'if "%%opcion%%"=="{idx+1}" start explorer /select, "{ruta_completa}" & exit\n')
        
        f.write('echo Opcion no valida & pause\n')
    return buscador_path

def main():
    print("Obteniendo archivos DWG y sus metadatos usando PowerShell (puede tomar varios minutos)...")
    df = obtener_archivos_dwg_con_metadatos()
    
    if df.empty:
        print("No se encontraron archivos DWG o hubo un error.")
        return
    
    df = df.drop_duplicates(subset=['RUTA_COMPLETA'])
    df = df.sort_values('ARCHIVO')
    
    ruta_excel = ruta_salida / "03_cardidwg_excel.xlsx"
    with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Archivos_DWG', index=False)
    
    buscador_path = crear_buscador_batch(df)
    escritorio = Path(os.environ['USERPROFILE']) / "Desktop"
    acceso_directo = escritorio / "BUSCADOR_DWG.bat"
    with open(acceso_directo, 'w', encoding='utf-8') as f:
        f.write(f'@start "" "{buscador_path}"')
    
    print(f"Proceso completado: {len(df)} archivos DWG encontrados")

if __name__ == "__main__":
    main()