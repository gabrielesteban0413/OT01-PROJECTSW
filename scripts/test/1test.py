import pandas as pd
import numpy as np
import sys

path_corp = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\03-asph_report_origin.xlsx"
path_port = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\01-Bulk export of ports.xlsx"
path_output = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\01-Bulk Merge_Clientes.xlsx"

try:
    import polars as pl
    USE_POLARS = True
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "polars"])
    import polars as pl
    USE_POLARS = True

def read_excel_fast(filepath, columns=None, dtype=str):
    if USE_POLARS:
        if columns is None:
            df = pl.read_excel(filepath, infer_schema_length=0)
        else:
            df = pl.read_excel(filepath, columns=columns, infer_schema_length=0)
        return df.to_pandas().astype(dtype)
    else:
        return pd.read_excel(filepath, dtype=dtype, usecols=columns, engine='openpyxl')

df_corp = read_excel_fast(path_corp, columns=None)

mask = df_corp['Estatus'].str.upper() == 'CONECTADO'
df_corp_filtrado = df_corp.loc[mask].copy()
conectados_count = len(df_corp_filtrado)

if conectados_count == 0:
    pd.DataFrame().to_excel(path_output, index=False)
    print(f"CONECTADO en dbcorporativo: 0")
    print(f"  - IDs NO encontrados en dbport: 0")
    sys.exit(0)

df_port = read_excel_fast(
    path_port,
    columns=['Id Servicio', 'Proyecto De Red']
)

df_port_clean = df_port.dropna(subset=['Id Servicio', 'Proyecto De Red'])
id_set = set(df_port_clean['Id Servicio'].unique())
pair_multiindex = pd.MultiIndex.from_frame(
    df_port_clean[['Id Servicio', 'Proyecto De Red']].drop_duplicates()
)

id_exists = df_corp_filtrado['IDServicio'].isin(id_set)
conectados_mi = pd.MultiIndex.from_arrays(
    [df_corp_filtrado['IDServicio'], df_corp_filtrado['Anillo']]
)
pair_exists = conectados_mi.isin(pair_multiindex)

df_corp_filtrado['validación idservicio'] = np.where(id_exists, 'OK', 'NO ENCONTRADO')
anillo_val = np.select(
    [~id_exists, id_exists & pair_exists, id_exists & ~pair_exists],
    ['NO APLICA', 'OK', 'ERROR (no coincide)'],
    default=''
)
df_corp_filtrado['validación anillo'] = anillo_val

if 'Estatus' in df_corp_filtrado.columns:
    df_corp_filtrado.drop(columns=['Estatus'], inplace=True)

try:
    import xlsxwriter
    engine_write = 'xlsxwriter'
except ImportError:
    engine_write = 'openpyxl'

with pd.ExcelWriter(path_output, engine=engine_write) as writer:
    df_corp_filtrado.to_excel(writer, index=False, sheet_name='Sheet1')

no_encontrados = df_corp_filtrado.loc[~id_exists, 'IDServicio'].unique()

print(f"CONECTADO en dbcorporativo: {conectados_count}")
print(f"  - IDs NO encontrados en dbport: {len(no_encontrados)}")