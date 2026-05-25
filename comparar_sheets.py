import pandas as pd
import numpy as np
import os, sys, io, warnings
warnings.filterwarnings("ignore")

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

BASE = r"C:\Users\hmolina\Documents\PERSONAL BEX\006 - Ajustre de Codigo GAS\comparacion"

ORIGINAL = os.path.join(BASE, "CONTROL DE CONTRATOS POR PROYECTOS (2).xlsx")
COPIA    = os.path.join(BASE, "Copia de CONTROL DE CONTRATOS POR PROYECTOS (2).xlsx")

HOJAS = ["CONTRATOS YAPE", "CONTRATOS BNB", "CONTRATOS ZAS"]

def leer_hoja(path, hoja):
    try:
        df = pd.read_excel(path, sheet_name=hoja, dtype=str)
        df = df.fillna("").apply(lambda c: c.str.strip() if c.dtype == object else c)
        return df
    except Exception as e:
        print(f"  ERROR leyendo {hoja}: {e}")
        return None

for hoja in HOJAS:
    print("=" * 70)
    print(f"HOJA: {hoja}")
    print("=" * 70)

    orig = leer_hoja(ORIGINAL, hoja)
    copia = leer_hoja(COPIA, hoja)

    if orig is None or copia is None:
        continue

    print(f"  Original : {len(orig):>5} filas | {orig.shape[1]} columnas")
    print(f"  Con cambios: {len(copia):>5} filas | {copia.shape[1]} columnas")
    print()

    # Llave CODIGO + CI
    col_codigo = "CODIGO"
    col_ci     = "C.I."
    if col_codigo not in orig.columns or col_ci not in orig.columns:
        print(f"  No se encontraron columnas CODIGO / C.I. — columnas disponibles:")
        print(f"  {list(orig.columns)}")
        continue

    orig["_KEY"]  = orig[col_codigo].str.strip()  + "|" + orig[col_ci].str.strip()
    copia["_KEY"] = copia[col_codigo].str.strip() + "|" + copia[col_ci].str.strip()

    keys_orig  = set(orig["_KEY"].unique())
    keys_copia = set(copia["_KEY"].unique())

    solo_orig  = keys_orig  - keys_copia
    solo_copia = keys_copia - keys_orig
    en_ambos   = keys_orig  & keys_copia

    print(f"  Llaves únicas original   : {len(keys_orig)}")
    print(f"  Llaves únicas con cambios: {len(keys_copia)}")
    print(f"  En ambos                 : {len(en_ambos)}")
    print()

    # Filas solo en original (eliminadas o sin match)
    if solo_orig:
        print(f"  SOLO EN ORIGINAL (no están en la copia): {len(solo_orig)}")
        muestra = orig[orig["_KEY"].isin(solo_orig)][["CODIGO","C.I.","NOMBRE COMPLETO","ESTADO"]].head(10)
        print(muestra.to_string(index=False))
        print()

    # Filas nuevas (solo en copia)
    if solo_copia:
        print(f"  NUEVAS EN COPIA (no estaban en original): {len(solo_copia)}")
        cols_vista = [c for c in ["CODIGO","C.I.","NOMBRE COMPLETO","ESTADO","FECHA DE INGRESO","CIUDAD"] if c in copia.columns]
        muestra = copia[copia["_KEY"].isin(solo_copia)][cols_vista].head(20)
        print(muestra.to_string(index=False))
        print()
    else:
        print(f"  Sin filas nuevas.")
        print()

    # Diferencias en filas que existen en ambos
    # Usar primera aparición de cada llave para comparar
    orig_idx  = orig.drop_duplicates("_KEY").set_index("_KEY")
    copia_idx = copia.drop_duplicates("_KEY").set_index("_KEY")

    cols_comunes = [c for c in orig_idx.columns if c in copia_idx.columns and c != "_KEY"]

    diffs_por_col = {}
    for col in cols_comunes:
        comun = orig_idx.index.intersection(copia_idx.index)
        a = orig_idx.loc[comun, col]
        b = copia_idx.loc[comun, col]
        mask = a != b
        if mask.any():
            diffs_por_col[col] = mask.sum()

    if diffs_por_col:
        print(f"  COLUMNAS CON DIFERENCIAS (en registros existentes en ambos):")
        for col, n in sorted(diffs_por_col.items(), key=lambda x: -x[1]):
            print(f"    {col:<30}: {n:>4} filas cambiaron")
        print()

        # Mostrar muestra de las columnas más cambiadas (top 3)
        top3 = sorted(diffs_por_col.items(), key=lambda x: -x[1])[:3]
        for col, n in top3:
            comun = orig_idx.index.intersection(copia_idx.index)
            mask  = orig_idx.loc[comun, col] != copia_idx.loc[comun, col]
            filas = comun[mask][:5]
            print(f"  Muestra de cambios en [{col}]:")
            for k in filas:
                v_orig  = orig_idx.loc[k, col]
                v_copia = copia_idx.loc[k, col]
                nombre  = orig_idx.loc[k, "NOMBRE COMPLETO"] if "NOMBRE COMPLETO" in orig_idx.columns else k
                print(f"    {nombre[:35]:<35} | antes: {str(v_orig)[:25]:<25} | ahora: {str(v_copia)[:25]}")
            print()
    else:
        print(f"  Sin diferencias en columnas de registros existentes.")
        print()

    # Duplicados
    dups_orig  = orig[orig["_KEY"].duplicated(keep=False)]
    dups_copia = copia[copia["_KEY"].duplicated(keep=False)]
    print(f"  Filas duplicadas (CODIGO+CI): original={len(dups_orig)}, con cambios={len(dups_copia)}")
    print()

print("=" * 70)
print("FIN DEL ANÁLISIS")
print("=" * 70)
