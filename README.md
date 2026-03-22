# Excel Analyzer

Library for detecting, extracting, and analyzing tables from messy Excel files with merged cells, empty rows/columns, and non-contiguous data blocks.

## Installation

```bash
pip install -e .
```

**Dependencies:** `pandas`, `numpy`, `openpyxl`, Python >= 3.9

---

## Modules

### Detection — `_core.py` / `_region_detector.py`

Automatically detects tables in an Excel sheet using a header-first strategy. Handles merged cells, multi-level headers, titles, side-by-side tables, and stacked tables.

```python
from excel_analyzer import detectar_todas_las_tablas, detectar_tabla, crear_tabla, diagnosticar

tablas = detectar_todas_las_tablas("archivo.xlsx", "Hoja1")
tabla  = detectar_tabla("archivo.xlsx", "Hoja1")          # first/best table
df     = crear_tabla(tabla)                                # export as DataFrame
diagnosticar("archivo.xlsx", "Hoja1")                     # print summary
```

**Auto-detected table types:**

| Tipo | Descripción |
|------|-------------|
| `manager_metrica` | MultiIndex (Manager, Métrica) × columnas de fecha |
| `solo_metrica` | Métrica index × columnas de fecha |
| `tabla_generica` | Primera columna = etiquetas, resto = datos |
| `tabla_cruzada` | Formato largo [Métrica, Header, Valor] |
| `clave_valor` | Pares clave–valor simples |

Other helpers:
```python
extraer_fila(tabla, "AUM")          # row by metric name
extraer_columna(tabla, "2024-01")   # column by date
subdividir_por_anio(tabla)          # split by year
reemplazar_valores(tabla, mapping)  # value replacement
```

---

### Query Engine — `query_engine.py`

Semantic search with fuzzy matching and synonym expansion across one or multiple sheets.

```python
from excel_analyzer import excel_query, buscar_tabla, describir_tablas

df = excel_query(
    "archivo.xlsx",
    metric="assets under management",   # fuzzy, case-insensitive
    sheet="Hoja1",                       # optional
    table=0,                             # optional table index
    where={"Manager": "Fondo A"},        # optional filter
    mode="best",                         # "best" | "merge"
    min_score=0.6
)
# df.attrs contains: sheet, table_id, matched_metric, score

describir_tablas(tablas)
tabla = buscar_tabla(tablas, "ROA")
```

Built-in synonym groups: `aum/assets under management`, `roa/return on assets`, `roi/return on investment`, and more.

---

### Table Builder — `table_builder.py`

Builds analytical tables with automatic pivoting based on detected table type.

```python
from excel_analyzer import excel_build_table

df = excel_build_table(
    "archivo.xlsx",
    sheet="Hoja1",
    table=0,
    metrics=["AUM", "ROA"],     # subset of metrics (fuzzy matched)
    columns=["ene-24", "feb-24"] # subset of columns
)
# df.attrs: sheet, table_id, table_tipo, matched_metrics
```

Pivot strategies per table type: `manager_metrica`, `solo_metrica`, `tabla_generica`, `tabla_cruzada`, `clave_valor`.

---

### Formula Navigator — `formula_navigator.py`

Move and update Excel formulas using semantic references (metric names and column headers), not raw cell coordinates.

```python
from excel_analyzer import (
    construir_mapa, inspeccionar_formulas,
    mover_a_ultima_columna, mover_a_ultima_fila,
    recorrer_columnas, recorrer_filas,
    apuntar_a_ultimo, actualizar_a_ultimo,
    agregar_columna_formulas, reapuntar_a_ultima_columna,
    actualizar_ref_absoluta, recorrer_columnas_rango,
    mapear_por_periodo,
)

mapa = construir_mapa("archivo.xlsx", "Hoja1", tabla_idx=0)
inspeccionar_formulas("archivo.xlsx", "Hoja1")

# Slide all temporal column references to the last column
mover_a_ultima_columna("archivo.xlsx", "Hoja1", tabla_idx=0)

# Shift column references by N positions
recorrer_columnas("archivo.xlsx", "Hoja1", n=1)

# Point snapshot formulas to the latest period
apuntar_a_ultimo("archivo.xlsx", "Hoja1")
actualizar_a_ultimo("archivo.xlsx", "Hoja1")
```

`TablaMap` provides bidirectional mappings: Excel column number ↔ header name, Excel row ↔ metric name.

---

### Chart Updater — `chart_updater.py`

Diagnose and automatically update chart ranges when new periods are added to a workbook.

```python
from excel_analyzer import diagnosticar_graficas, actualizar_graficas

infos = diagnosticar_graficas("archivo.xlsx")
# Each ChartInfo: hoja, indice, tipo_chart, clasificacion, necesita_actualizacion

resultado = actualizar_graficas(
    "archivo.xlsx",
    archivo_datos="datos.xlsx",      # source of new data
    archivo_salida="salida.xlsx",
    borrar_rotas=True,               # delete #REF! charts
    actualizar_titulos=True,         # update month names in titles
    ventana_fija=12,                 # keep rolling 12-month window
    verbose=True
)
```

**Chart classifications:**

| Clasificación | Descripción |
|---------------|-------------|
| `ventana_temporal` | Rango horizontal deslizante (meses) |
| `ultimo_valor` | Rango vertical o celda única (snapshot) |
| `estatica` | Rango fijo sin dependencia temporal |
| `rota` | Contiene errores `#REF!` |
| `mixta` | Combinación de horizontal y vertical |

---

### Chart Creator — `chart_creator.py`

Create charts from a configuration object.

```python
from excel_analyzer import (
    GraficaConfig, SerieConfig,
    crear_graficas_desde_config,
    ventana_temporal_series, snapshot_series,
)

series = ventana_temporal_series(
    hoja_datos="Datos", fila_header=2, filas_datos=[3,4,5],
    fila_labels_col=1, col_fin=14, ventana=12
)

config = GraficaConfig(
    hoja_destino="Dashboard",
    estilo="bar_stacked",          # bar_stacked | bar_clustered | bar_stacked_100 | pie | bar_cambio
    posicion="B2",
    series=series,
    titulo="AUM por Fondo",
    width=600, height=300,
)

crear_graficas_desde_config("archivo.xlsx", [config], archivo_salida="salida.xlsx")
```

---

### Formula Annexes — `Anexos_formulas.py`

Export a sheet with all its formula dependencies resolved, annexing every referenced sheet.

```python
from excel_analyzer import analizar_y_exportar, exportar_todas_con_formulas

# Export one sheet + all sheets it references
resultado = analizar_y_exportar("archivo.xlsx", "Hoja1", "salida.xlsx")
# resultado: {"hoja", "anexos", "n_formulas"}

# Export ALL sheets that contain formulas, without duplication
resultado = exportar_todas_con_formulas("archivo.xlsx", "salida.xlsx")
# resultado: {"hojas_con_formulas", "hojas_copiadas", "total_formulas"}
```

---

## Running Tests

```bash
python excel_analyzer/test_basico.py
```

Tests cover: datetime dates, string dates (`ene-24`), merged cell headers, stacked tables, and noisy sheets (empty rows/columns + title).

## Build

```bash
python -m build
# outputs: dist/excel_analyzer-0.1.0-py3-none-any.whl
```
