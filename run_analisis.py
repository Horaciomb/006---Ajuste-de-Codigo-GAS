import pandas as pd
import numpy as np
import json, sys, io, warnings
warnings.filterwarnings("ignore")

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import os
os.chdir(r"C:\Users\hmolina\Documents\PERSONAL BEX\006 - Ajustre de Codigo GAS")

PATH_FUENTE    = "PERSONAL YAPE.csv"
PATH_CONTRATOS = "Copia de CONTROL DE CONTRATOS POR PROYECTOS.xlsx"
HOJA_CONTRATOS = "CONTRATOS YAPE"

df_yape = pd.read_csv(PATH_FUENTE, dtype=str, encoding="utf-8")
df_yape = df_yape.fillna("")
df_yape.columns = df_yape.columns.str.strip()

df_contratos = pd.read_excel(PATH_CONTRATOS, sheet_name=HOJA_CONTRATOS, dtype=str)
df_contratos = df_contratos.fillna("")
df_contratos.columns = df_contratos.columns.str.strip()

print(f"Fuente YAPE:        {len(df_yape):>5} filas | {df_yape.shape[1]} columnas")
print(f"Contratos YAPE:     {len(df_contratos):>5} filas | {df_contratos.shape[1]} columnas")

# === SECCIÓN 1: Validación de llaves ===
print("\n=== LLAVE: solo CODIGO ===")
print(f"Duplicados en YAPE fuente:    {df_yape['CODIGO'].str.strip().duplicated().sum()}")
print(f"Duplicados en CONTRATOS:      {df_contratos['CODIGO'].str.strip().duplicated().sum()}")
print(f"Únicos en YAPE:               {df_yape['CODIGO'].str.strip().nunique()}")
print(f"Únicos en CONTRATOS:          {df_contratos['CODIGO'].str.strip().nunique()}")

df_yape["KEY"]      = df_yape["CODIGO"].str.strip() + "|" + df_yape["C.I."].str.strip()
df_contratos["KEY"] = df_contratos["CODIGO"].str.strip() + "|" + df_contratos["C.I."].str.strip()

print("\n=== LLAVE: CODIGO + C.I. ===")
print(f"Duplicados en YAPE fuente:    {df_yape['KEY'].duplicated().sum()}")
print(f"Duplicados en CONTRATOS:      {df_contratos['KEY'].duplicated().sum()}")
print(f"Únicos en YAPE:               {df_yape['KEY'].nunique()}")
print(f"Únicos en CONTRATOS:          {df_contratos['KEY'].nunique()}")

# === SECCIÓN 2: Faltantes ===
keys_yape      = set(df_yape["KEY"])
keys_contratos = set(df_contratos["KEY"])
solo_en_yape      = keys_yape - keys_contratos
solo_en_contratos = keys_contratos - keys_yape
en_ambos          = keys_yape & keys_contratos

print(f"\nMatch (en ambos):                   {len(en_ambos)}")
print(f"Solo en fuente YAPE (faltantes):    {len(solo_en_yape)}")
print(f"Solo en CONTRATOS (sin match):      {len(solo_en_contratos)}")

faltantes = df_yape[df_yape["KEY"].isin(solo_en_yape)].copy()
print(f"\n=== FALTANTES EN CONTRATOS ===")
print(f"Total: {len(faltantes)}")
print("\nPor estado:")
print(faltantes["ESTADO (ACTIVO/INACTIVO)"].value_counts().to_string())

activos_falt = faltantes[faltantes["ESTADO (ACTIVO/INACTIVO)"].str.upper().str.strip() == "ACTIVO"]
print(f"\nACTIVOS faltantes (deben insertarse): {len(activos_falt)}")

# Filtrar ABR/MAY 2026
def parse_fecha(f):
    try: return pd.to_datetime(f, format="%m/%d/%Y")
    except:
        try: return pd.to_datetime(f)
        except: return pd.NaT

faltantes["FECHA_PARSED"] = faltantes["FECHA DE INGRESO"].apply(parse_fecha)
recientes = faltantes[faltantes["FECHA_PARSED"] >= "2026-04-01"]
print(f"\nFaltantes con ingreso ABR/MAY 2026: {len(recientes)}")
print(recientes[["CODIGO","C.I.","NOMBRE","ESTADO (ACTIVO/INACTIVO)","FECHA DE INGRESO","CIUDAD"]].to_string())

# === SECCIÓN 3: Discrepancias ===
print("\n=== DISCREPANCIAS EN REGISTROS EXISTENTES ===")
MAPEO = {
    "CELULAR":                  "CELULAR",
    "ESTADO (ACTIVO/INACTIVO)": "ESTADO",
    "CARGO":                    "CARGO",
    "SUPERVISOR":               "SUPERVISOR",
    "CIUDAD":                   "CIUDAD",
}
df_yape_idx      = df_yape.set_index("KEY")
df_contratos_idx = df_contratos.drop_duplicates(subset="KEY", keep="first").set_index("KEY")

resumen_disc = {}
for col_fuente, col_contrato in MAPEO.items():
    if col_fuente not in df_yape_idx.columns: continue
    if col_contrato not in df_contratos_idx.columns: continue
    comun = df_yape_idx.index.intersection(df_contratos_idx.index)
    a = df_yape_idx.loc[comun, col_fuente].str.strip().str.upper()
    b = df_contratos_idx.loc[comun, col_contrato].str.strip().str.upper()
    diffs = (a != b).sum()
    resumen_disc[col_fuente] = diffs
    if diffs > 0:
        print(f"\nDISCREPANCIA {col_fuente} -> {col_contrato}: {diffs} diferencias")
        muestra = comun[a != b][:10]
        for k in muestra:
            print(f"  {k}: fuente={df_yape_idx.loc[k,col_fuente]!r}  contrato={df_contratos_idx.loc[k,col_contrato]!r}")

print("\n=== RESUMEN DISCREPANCIAS ===")
for col, n in resumen_disc.items():
    marca = "OK" if n == 0 else f"DIFERENCIAS: {n}"
    print(f"  {col:<35} {marca}")

# === SECCIÓN 4: Duplicados en CONTRATOS ===
dups = df_contratos[df_contratos["KEY"].duplicated(keep=False)]
print(f"\n=== DUPLICADOS EN CONTRATOS ===")
print(f"Filas duplicadas (mismo CODIGO+CI): {len(dups)}")

# === SECCIÓN 5: Columnas manuales ===
print("\n=== COLUMNAS MANUALES: datos no vacíos ===")
cols_manuales = ["FECHA DE CONTRATO","FIRMA","ESTADO DE CONTRATO","OBS","OBS SUPERVISOR","OBS SUBGERENTE"]
for col in cols_manuales:
    if col in df_contratos.columns:
        n = (df_contratos[col].str.strip() != "").sum()
        print(f"  {col:<30}: {n:>5} registros con dato")
    else:
        print(f"  {col:<30}: columna no encontrada")

# === SECCIÓN 6: Activos a insertar ===
activos_insertar = faltantes[faltantes["ESTADO (ACTIVO/INACTIVO)"].str.upper().str.strip() == "ACTIVO"]
print(f"\nTotal ACTIVOS a insertar: {len(activos_insertar)}")
if len(activos_insertar) > 0:
    print(activos_insertar[["CODIGO","C.I.","NOMBRE","ESTADO (ACTIVO/INACTIVO)","FECHA DE INGRESO","CIUDAD"]].to_string())
