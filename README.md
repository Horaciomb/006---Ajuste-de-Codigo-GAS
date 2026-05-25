# Control de Contratos por Proyectos — Sincronización GAS + Análisis Python

Herramienta para sincronizar el sheet de control de contratos de RRHH con las planillas fuente de YAPE, BNB y ZAS. Combina un script de Google Apps Script para actualizar el Google Sheet en producción, y un notebook de Jupyter para analizar discrepancias antes y después de cada sincronización.

---

## Descripción del proyecto

El sheet **CONTROL DE CONTRATOS POR PROYECTOS** centraliza el seguimiento de contratos del personal activo/inactivo de tres empresas. Los datos del personal viven en planillas separadas (publicadas como CSV desde Google Sheets) y deben sincronizarse periódicamente con el sheet de contratos, respetando columnas que el equipo completa manualmente.

### Problema que resuelve

Las versiones anteriores del script usaban el celular o el CODIGO como llave de identificación de cada persona. Ambos campos cambian o se reutilizan, lo que causaba que:

- Ingresos nuevos de abril/mayo no aparecieran en contratos
- Actualizaciones de campos (supervisor, cargo, estado) se perdieran silenciosamente
- Se generaran filas duplicadas en el sheet

La solución implementada usa `CODIGO + C.I.` como llave compuesta, que es única por persona en las tres fuentes.

---

## Estructura del repositorio

```
├── original.gs              # Script de Google Apps Script (producción)
│                            # Contiene actualizarYAPE(), actualizarBNB(), actualizarZAS()
│
├── actualizarYape.gs        # Versión draft del fix de YAPE (referencia)
├── actualizarYAPE_v2.3.gs   # Versión standalone de actualizarYAPE() (referencia)
│
├── analisis_yape.ipynb      # Notebook de análisis y validación
│
├── requirements.txt         # Dependencias Python
└── .gitignore
```

> Los archivos de datos (`*.csv`, `*.xlsx`) están excluidos del repositorio. Deben descargarse manualmente antes de ejecutar el notebook.

---

## Configuración del entorno Python

El notebook requiere Python 3.x y las librerías listadas en `requirements.txt`.

### 1. Crear el entorno virtual

```bash
python -m venv venv
```

### 2. Activar el entorno

**Windows (PowerShell):**
```powershell
.\venv\Scripts\Activate.ps1
```

**Windows (CMD):**
```cmd
venv\Scripts\activate.bat
```

**macOS / Linux:**
```bash
source venv/bin/activate
```

### 3. Instalar las dependencias

```bash
pip install -r requirements.txt
```

### 4. Abrir el notebook

```bash
jupyter notebook analisis_yape.ipynb
```

---

## Archivos de datos necesarios (no incluidos en el repo)

Antes de ejecutar el notebook, colocar en la raíz del proyecto:

| Archivo | Descripción | Cómo obtenerlo |
|---------|-------------|----------------|
| `PERSONAL YAPE.csv` | Planilla fuente de personal YAPE | Descargar desde Google Sheets → Archivo → Descargar → CSV |
| `Copia de CONTROL DE CONTRATOS POR PROYECTOS.xlsx` | Copia del sheet de contratos | Descargar desde Google Sheets → Archivo → Descargar → XLSX |

---

## Qué hace el notebook (`analisis_yape.ipynb`)

Compara la planilla fuente YAPE contra el sheet de contratos e identifica:

1. **Validación de llave** — Confirma que `CODIGO + C.I.` es única en fuente (0 duplicados) y detecta cuántas hay en contratos.
2. **Registros faltantes** — Lista personas en fuente que no están en contratos, separadas por estado (ACTIVO / INACTIVO). Los ACTIVOS faltantes son los que deben insertarse.
3. **Ingresos recientes** — Filtra los faltantes con fecha de ingreso en abril/mayo del año en curso.
4. **Discrepancias** — Para registros que sí existen en ambos, compara columnas automáticas (CELULAR, ESTADO, CARGO, SUPERVISOR, CIUDAD) e informa diferencias.
5. **Duplicados en contratos** — Detecta filas con el mismo `CODIGO + C.I.` repetido.
6. **Columnas manuales** — Valida que los campos `FECHA DE CONTRATO`, `FIRMA`, `ESTADO DE CONTRATO`, `OBS`, `OBS SUPERVISOR` y `OBS SUBGERENTE` no están vacíos antes y después de la sincronización.
7. **Reporte exportable** — Genera `reporte_faltantes_yape.csv` con los ACTIVOS a insertar, listo para revisión.

---

## Qué hace el Apps Script (`original.gs`)

Contiene tres funciones principales que se ejecutan desde el menú **Procesar** del Google Sheet:

### `actualizarYAPE()` / `actualizarBNB()` / `actualizarZAS()`

Para cada empresa, el script:

1. **Descarga** el CSV de la planilla fuente desde su URL pública.
2. **Lee** el sheet de contratos completo en memoria.
3. **Construye un índice** de registros existentes usando la llave `CODIGO + C.I.`
4. **Para cada persona en la fuente:**
   - Si ya existe en contratos → actualiza todas las columnas automáticas (A–D, E–L, N, O) y preserva las columnas manuales (M, P, Q, R, S, T).
   - Si no existe y está ACTIVA → inserta una fila nueva con los campos manuales en blanco.
   - Si no existe y está INACTIVA → la omite (salió antes de ser registrada).
5. **Escribe** todos los cambios en un solo bloque (`setValues` único), minimizando el número de llamadas a la API.

### Columnas del sheet de contratos

| Cols | Tipo | Descripción |
|------|------|-------------|
| A–D | Automático | Empresa, CODIGO, C.I., CELULAR |
| E–L | Automático | Nombres, CARGO, CIUDAD, DPTO, SUPERVISOR |
| M | **Manual** | FECHA DE CONTRATO — nunca se sobreescribe |
| N | Automático | FECHA DE INGRESO |
| O | Automático | ESTADO (ACTIVO / INACTIVO) |
| P–T | **Manual** | FIRMA, ESTADO DE CONTRATO, OBS, OBS SUPERVISOR, OBS SUBGERENTE |

### Cómo instalar el script en Google Sheets

1. Abrir el Google Sheet → **Extensiones → Apps Script**
2. Seleccionar todo el contenido del editor (`Ctrl+A`) y reemplazarlo con el contenido de `original.gs`
3. Guardar (`Ctrl+S`) y recargar el Google Sheet
4. El menú **Procesar** aparecerá en la barra del sheet

---

## Flujo de trabajo recomendado

```
1. Descargar CSV y XLSX actualizados de Google Sheets
2. Ejecutar analisis_yape.ipynb → revisar faltantes y discrepancias
3. Ejecutar "Procesar → Actualizar YAPE/BNB/ZAS" en el Google Sheet
4. Volver a descargar el XLSX actualizado
5. Re-ejecutar el notebook para confirmar que el resultado es el esperado
```

---

## Dependencias principales

| Librería | Uso |
|----------|-----|
| `pandas` | Lectura de CSV/XLSX y comparación de datos |
| `openpyxl` | Lectura de archivos `.xlsx` |
| `numpy` | Soporte numérico para pandas |
| `ipykernel` | Ejecución del notebook en Jupyter |
