"""
excel_analyzer
==============
Librería para detectar y consultar tablas en archivos Excel desordenados
con merged cells, filas/columnas vacías y bloques de datos no contiguos.

MÓDULOS
=======
    _core           Motor de detección: SheetScanner, TableAnalyzer, API base.
    query_engine    Motor de consulta: excel_query(), sinónimos, fuzzy matching.
    table_builder   Motor de construcción: excel_build_table(), pivots analíticos.
    formula_navigator  Navegación y movimiento de fórmulas por mapeo semántico.
    chart_updater   Detección y actualización automática de gráficas existentes.
    chart_creator   Creación de gráficas desde cero con estilos predefinidos.

API RÁPIDA
==========

    # Detección básica
    from excel_analyzer import detectar_tabla, detectar_todas_las_tablas
    tabla  = detectar_tabla("archivo.xlsx", "Hoja1")
    tablas = detectar_todas_las_tablas("archivo.xlsx", "Hoja1")

    # Extracción
    from excel_analyzer import extraer_fila, extraer_columna, crear_tabla
    extraer_fila(tabla, "AUM")
    extraer_columna(tabla, "ene-26")

    # Query engine — busca en todas las hojas automáticamente
    from excel_analyzer import excel_query
    df = excel_query("archivo.xlsx", metric="AUM")
    df = excel_query("archivo.xlsx", metric="assets")           # sinónimo
    df = excel_query("archivo.xlsx", metric="AUM",
                     where="Manager == 'First Trust'")          # filtro
    df = excel_query("archivo.xlsx", metric="AUM",
                     sheet=2, table=1)                          # hoja específica
    df = excel_query("archivo.xlsx", metric="AUM",
                     mode="merge")                              # todas las hojas

    # Table builder — tablas analíticas con pivot
    from excel_analyzer import excel_build_table
    df = excel_build_table(
        "archivo.xlsx", sheet=2, table=1,
        metrics=["AUM", "ROA"],
        columns=["ene-24", "feb-24", "mar-24"]
    )
    #             AUM                    ROA
    #          ene-24 feb-24 mar-24   ene-24 feb-24 mar-24
    # Manager
    # First Trust  …    …    …          …      …      …

    # Exploración
    from excel_analyzer import describir_tablas, buscar_tabla, tabla_mas_grande
    describir_tablas(tablas)
    tabla = buscar_tabla(tablas, "AUM")
    tabla = tabla_mas_grande(tablas)

    # Gráficas — diagnóstico, actualización y creación
    from excel_analyzer import diagnosticar_graficas, actualizar_graficas
    diagnosticar_graficas("archivo.xlsx")
    actualizar_graficas("archivo.xlsx", borrar_rotas=True, ventana_fija=13)

    from excel_analyzer import crear_graficas_desde_config, GraficaConfig
    crear_graficas_desde_config("archivo.xlsx", [config1, config2])
"""

# ── Motor de detección (núcleo) ───────────────────────────────────────────────
from ._core import (
    # Clases internas
    WorkbookLoader,
    SheetScanner,
    TableParser,
    TableAnalyzer,
    TableRegion,
    # API de detección
    detectar_tabla,
    detectar_todas_las_tablas,
    # API de extracción
    extraer_fila,
    extraer_columna,
    crear_tabla,
    replicar_tabla,
    # Diagnóstico
    diagnosticar,
    subdividir_por_anio,
    reemplazar_valores,
)

# ── Motor de consulta (query engine) ─────────────────────────────────────────
from .query_engine import (
    # Función principal
    excel_query,
    # Utilidades de exploración
    buscar_tabla,
    tabla_mas_grande,
    describir_tablas,
    # Motor semántico (accesible para extensión y tests)
    _find_best_metric,
    _get_synonym_candidates,
    _similarity,
    _normalize_text,
)

# ── Motor de construcción de tablas ──────────────────────────────────────────
from .table_builder import (
    excel_build_table,
)

# ── Motor de fórmulas y anexos ──────────────────────────────────────────
from .Anexos_formulas import (
    analizar_y_exportar,
    exportar_todas_con_formulas,
)

from .formula_navigator import (
    inspeccionar_formulas,
    mover_a_ultima_columna,
    mover_a_ultima_fila,
    recorrer_columnas,
    recorrer_filas,
    construir_mapa,
    mapear_por_periodo,
    apuntar_a_ultimo,
    actualizar_a_ultimo,
    agregar_columna_formulas,
    recorrer_columnas_rango,
    reapuntar_a_ultima_columna,
    actualizar_ref_absoluta,
)

# ── Motor de gráficas — actualización ────────────────────────────────────────
from .chart_updater import (
    diagnosticar_graficas,
    actualizar_graficas,
)

# ── Motor de gráficas — creación ─────────────────────────────────────────────
from .chart_creator import (
    GraficaConfig,
    SerieConfig,
    crear_grafica,
    crear_graficas_desde_config,
    ventana_temporal_series,
    snapshot_series,
)

__all__ = [
    # Clases
    "WorkbookLoader",
    "SheetScanner",
    "TableParser",
    "TableAnalyzer",
    "TableRegion",
    # Detección
    "detectar_tabla",
    "detectar_todas_las_tablas",
    # Extracción
    "extraer_fila",
    "extraer_columna",
    "crear_tabla",
    "replicar_tabla",
    # Diagnóstico / transformación
    "diagnosticar",
    "subdividir_por_anio",
    "reemplazar_valores",
    # Query engine
    "excel_query",
    "buscar_tabla",
    "tabla_mas_grande",
    "describir_tablas",
    # Semántico (para extensión / tests)
    "_find_best_metric",
    "_get_synonym_candidates",
    "_similarity",
    "_normalize_text",
    # Table builder
    "excel_build_table",
    # Fórmulas y anexos
    "analizar_y_exportar",
    "exportar_todas_con_formulas",
    # Navigator
    "inspeccionar_formulas",
    "mover_a_ultima_columna",
    "mover_a_ultima_fila",
    "recorrer_columnas",
    "recorrer_filas",
    "construir_mapa",
    "mapear_por_periodo",
    "apuntar_a_ultimo",
    "actualizar_a_ultimo",
    "agregar_columna_formulas",
    "recorrer_columnas_rango",
    "reapuntar_a_ultima_columna",
    "actualizar_ref_absoluta",
    # Chart updater
    "diagnosticar_graficas",
    "actualizar_graficas",
    # Chart creator
    "GraficaConfig",
    "SerieConfig",
    "crear_grafica",
    "crear_graficas_desde_config",
    "ventana_temporal_series",
    "snapshot_series",
]