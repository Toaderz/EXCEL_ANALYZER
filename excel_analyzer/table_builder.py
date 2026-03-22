"""
excel_analyzer/table_builder.py
================================
Motor de construcción de tablas analíticas para Excel Analyzer.

Construido SOBRE la infraestructura existente (detectar_todas_las_tablas,
_resolve_sheets, _select_table) sin modificar el motor de detección.

API pública:
    excel_build_table(archivo, sheet, table, metrics, columns)

DIAGRAMA DE FLUJO
=================

    excel_build_table(archivo, sheet, table, metrics, columns)
              │
              ▼
    _resolve_sheets(sheet)      ← 1-based int → nombre de hoja  (query_engine)
              │
              ▼
    detectar_todas_las_tablas() ← de _core
              │
              ▼
    _select_table(table)        ← filtra por id 1-based         (query_engine)
              │
              ▼
    _match_metrics()            ← exact → case-insensitive → substring → fuzzy
              │
              ▼
    _validate_columns()         ← verifica que existan en el DataFrame
              │
              ▼
    _pivot_table()              ← unstack + swaplevel + reindex (orden garantizado)
              │
              ▼
    pd.DataFrame con MultiIndex de columnas:
        Métrica (outer)  ×  Fecha/Columna (inner)
        Manager (filas)   ← o Fecha (filas) para solo_metrica

PIVOT PARA manager_metrica
===========================

    Entrada:  MultiIndex (Manager, Métrica) × date columns
    Objetivo:
              AUM              ROA
           ene-24 feb-24    ene-24 feb-24
    First Trust  …    …        …    …
    Columbia     …    …        …    …

    Pasos:
      1. df[mask][cols]                   → filtra métricas y columnas
      2. .unstack('Métrica')              → (date, Métrica) como columnas
      3. .swaplevel(axis=1)               → (Métrica, date) como columnas
      4. .reindex(desired_multiindex)     → fuerza orden exacto del usuario

PIVOT PARA solo_metrica (sin Manager)
======================================

    Entrada:  Index (Métrica,) × date columns
    Objetivo (Métrica → columnas, fechas → filas):

              Salarios    Impuestos
    ene-26    1_289_845    462_260
    feb-26    1_510_000    200_000

    Pasos:
      1. df[mask][cols]   → filtra
      2. .T               → transpone: fechas como filas, métricas como columnas
      3. reindex(metrics) → fuerza orden de métricas

MATCHING DE MÉTRICAS
====================
Para cada nombre en 'metrics' se intenta en orden:
  1. Exacto          "AUM"  → "AUM"
  2. Case-insensitive "aum" → "AUM"
  3. Substring        "Ingreso" → "Ingreso Generado"
  4. Fuzzy (_similarity)  score >= 0.4

Si no hay match, se lanza ValueError con todos los errores acumulados.
"""

from __future__ import annotations

from typing import Any

import pandas as pd

from ._core import detectar_todas_las_tablas
from .query_engine import _resolve_sheets, _select_table, _similarity, _get_synonym_candidates


# ══════════════════════════════════════════════════════════════════════════════
# MATCHING DE MÉTRICAS
# ══════════════════════════════════════════════════════════════════════════════

_FUZZY_MIN_SCORE = 0.4


def _match_metric(query: str, available: list[str]) -> str:
    """
    Resuelve un nombre de métrica solicitado contra la lista de métricas disponibles.

    Estrategia en orden de prioridad:
      1. Coincidencia exacta.
      2. Case-insensitive (ej. "aum" → "AUM").
      3. Substring (ej. "Ingreso" → "Ingreso Generado").
      4. Fuzzy matching con _similarity (score >= _FUZZY_MIN_SCORE).

    Args:
        query    : Nombre tal como lo pasó el usuario.
        available: Lista de nombres de métricas en el DataFrame.

    Returns:
        El nombre exacto de la métrica en el DataFrame.

    Raises:
        ValueError si no se encontró ningún match con score suficiente.
    """
    # 1. Exacto
    if query in available:
        return query

    ql = query.lower()

    # 2. Case-insensitive
    for m in available:
        if m.lower() == ql:
            return m

    # 3. Substring (query contenido en métrica o viceversa)
    for m in available:
        if ql in m.lower() or m.lower() in ql:
            return m

    # 4. Fuzzy con expansión de sinónimos: el mismo grupo que usa excel_query.
    #    "return on assets" → candidatos ["return on assets", "roa"]
    #    → max_score("roa", "ROA") = 1.0  ✓
    candidates = _get_synonym_candidates(query)
    best_m, best_s = None, 0.0
    for cand in candidates:
        for m in available:
            s = _similarity(cand, m)
            if s > best_s:
                best_s, best_m = s, m

    if best_m is not None and best_s >= _FUZZY_MIN_SCORE:
        return best_m

    raise ValueError(
        f"Métrica '{query}' no encontrada.\n"
        f"Métricas disponibles: {available}"
    )


def _match_metrics(
    requested: list[str],
    available: list[str],
) -> list[str]:
    """
    Resuelve una lista de métricas solicitadas, acumulando todos los errores.

    Returns:
        Lista de nombres reales en el mismo orden que 'requested'.

    Raises:
        ValueError con todos los nombres no encontrados si hay alguno.
    """
    results: list[str] = []
    errors:  list[str] = []

    for q in requested:
        try:
            results.append(_match_metric(q, available))
        except ValueError:
            errors.append(q)

    if errors:
        raise ValueError(
            f"Las siguientes métricas no se encontraron: {errors}\n"
            f"Métricas disponibles: {available}"
        )

    return results


# ══════════════════════════════════════════════════════════════════════════════
# VALIDACIÓN DE COLUMNAS
# ══════════════════════════════════════════════════════════════════════════════

def _validate_columns(
    requested: list[str],
    available: list[str],
) -> list[str]:
    """
    Verifica que todas las columnas solicitadas existan en el DataFrame.

    Returns:
        La misma lista 'requested' si todas son válidas.

    Raises:
        ValueError con las columnas no encontradas y las disponibles.
    """
    invalid = [c for c in requested if c not in available]
    if invalid:
        raise ValueError(
            f"Las siguientes columnas no se encontraron: {invalid}\n"
            f"Columnas disponibles: {available}"
        )
    return requested

# ══════════════════════════════════════════════════════════════════════════════
# PIVOT ENGINE — una función por tipo de tabla
# ══════════════════════════════════════════════════════════════════════════════

def _pivot_manager_metrica(
    df: pd.DataFrame,
    matched_metrics: list[str],
    columns: list[str],
) -> pd.DataFrame:
    """
    Construye la tabla pivotada para tablas de tipo manager_metrica.

    Entrada:  MultiIndex (Manager, Métrica) × date columns
    Salida:   Manager × (Métrica, Fecha) con el orden exacto solicitado.

        Métrica         AUM              ROA
        Fecha        ene-24 feb-24    ene-24 feb-24
        Manager
        First Trust    …      …         …      …
        Columbia       …      …         …      …

    El orden de métricas y columnas fecha sigue exactamente las listas recibidas.
    """
    mask     = df.index.get_level_values("Métrica").isin(matched_metrics)
    filtered = df.loc[mask, columns]
    unstacked = filtered.unstack(level="Métrica")
    swapped   = unstacked.swaplevel(axis=1)
    desired   = pd.MultiIndex.from_tuples(
        [(m, c) for m in matched_metrics for c in columns],
        names=["Métrica", swapped.columns.names[1]],
    )
    return swapped.reindex(columns=desired)


def _pivot_solo_metrica(
    df: pd.DataFrame,
    matched_metrics: list[str],
    columns: list[str],
) -> pd.DataFrame:
    """
    Construye la tabla pivotada para tablas de tipo solo_metrica.

    Sin Manager — transpone para que las fechas sean filas y las métricas columnas.

        Métrica    Salarios y Beneficios  Impuestos
        Fecha
        ene-26             1_289_845      462_260
        feb-26             1_510_000      200_000
    """
    mask     = df.index.get_level_values("Métrica").isin(matched_metrics)
    filtered = df.loc[mask, columns]
    result   = filtered.T
    result.columns.name = "Métrica"
    result.index.name   = "Fecha"
    return result[matched_metrics]


def _pivot_generic(
    df: pd.DataFrame,
    matched_metrics: list[str],
    columns: list[str],
) -> pd.DataFrame:
    """
    Construye la tabla para tablas genéricas o rotadas.

    Estas tablas tienen estructura plana:
        primera columna = etiquetas de fila (equivalente a métricas)
        columnas restantes = columnas de datos

    Ejemplo:
        Producto  Ventas  Costos
        A         100     50
        B         120     70

        excel_build_table(..., metrics=["A","B"], columns=["Ventas","Costos"])

        →   Ventas  Costos
        A   100     50
        B   120     70

    La primera columna se usa como índice del resultado.
    """
    label_col = df.columns[0]
    mask      = df[label_col].isin(matched_metrics)
    cols_sel  = [label_col] + columns if columns else [label_col]
    filtered  = df.loc[mask, cols_sel]
    result    = filtered.set_index(label_col)
    # Preserve metric order as requested
    met_order = [m for m in matched_metrics if m in result.index]
    return result.loc[met_order]


def _pivot_cruzada(
    df: pd.DataFrame,
    matched_metrics: list[str],
    columns: list[str],
) -> pd.DataFrame:
    """
    Re-pivota una tabla cruzada normalizada (formato largo) a formato ancho.

    Después de que _parse_generic_table aplica melt, la tabla cruzada tiene:
        Métrica  Header  Valor

    Este pivot la devuelve al formato analítico wide con orden garantizado:
        [col headers como columnas, metrics como filas]

    Ejemplo:
        Métrica  Header  Valor    →    2023  2024
        AUM      2023    10            AUM   10    20
        AUM      2024    20            ROA   1.2   1.4
        ROA      2023    1.2
        ROA      2024    1.4
    """
    filtered = df[df["Métrica"].isin(matched_metrics)].copy()
    if columns:
        filtered = filtered[filtered["Header"].isin(columns)]

    result = filtered.pivot(index="Métrica", columns="Header", values="Valor")
    result.columns.name = None

    # Reorder rows and cols to match user request
    met_order = [m for m in matched_metrics if m in result.index]
    col_order = [c for c in (columns or list(result.columns)) if c in result.columns]
    return result.loc[met_order, col_order]


def _pivot_key_value(
    df: pd.DataFrame,
    matched_metrics: list[str],
    columns: list[str],
) -> pd.DataFrame:
    """
    Devuelve una tabla clave-valor filtrada por las filas y columnas solicitadas.

    Las tablas clave_valor ya tienen 'Nombre' como índice y 'Valor' como única columna.
    metrics → filtra el índice (nombres de fila)
    columns → selecciona columnas (generalmente ['Valor'])
    """
    valid_metrics = [m for m in matched_metrics if m in df.index]
    result = df.loc[valid_metrics] if valid_metrics else df.copy()
    if columns:
        valid_cols = [c for c in columns if c in result.columns]
        result = result[valid_cols] if valid_cols else result
    return result


def _get_available_metrics(df: pd.DataFrame, tipo: str) -> list[str]:
    """
    Devuelve la lista de métricas disponibles según el tipo de tabla.

    La estrategia varía porque cada tipo almacena las etiquetas de diferente manera:
      - Tablas financieras: en el índice 'Métrica'.
      - Tabla cruzada (post-melt): en la columna 'Métrica'.
      - Tabla genérica / rotada: en la primera columna del DataFrame.
      - Tabla clave-valor: en el índice 'Nombre'.
    """
    if tipo in ("manager_metrica", "solo_metrica"):
        return df.index.get_level_values("Métrica").unique().tolist()
    if tipo == "tabla_cruzada":
        return df["Métrica"].unique().tolist()
    if tipo == "clave_valor":
        return df.index.unique().tolist()
    # tabla_generica, tabla_rotada
    return df.iloc[:, 0].dropna().astype(str).unique().tolist()


def _get_available_columns(df: pd.DataFrame, tipo: str) -> list[str]:
    """
    Devuelve la lista de columnas válidas para el parámetro 'columns'.

    Para tabla_cruzada los valores usables son los valores únicos de la columna
    'Header' (no las columnas del DataFrame, que son Métrica/Header/Valor).
    Para tabla_generica/rotada se excluye la primera columna (etiquetas de fila).
    """
    if tipo in ("manager_metrica", "solo_metrica", "clave_valor"):
        return list(df.columns)
    if tipo == "tabla_cruzada":
        return df["Header"].unique().tolist()
    # tabla_generica, tabla_rotada
    return list(df.columns[1:])


# ══════════════════════════════════════════════════════════════════════════════
# API PÚBLICA — excel_build_table
# ══════════════════════════════════════════════════════════════════════════════

def excel_build_table(
    archivo: str,
    sheet: int,
    table: int,
    metrics: list[str],
    columns: list[str],
) -> pd.DataFrame:
    """
    Construye una tabla analítica a partir de cualquier tabla detectada en Excel.

    Funciona con todos los tipos de tabla que Excel Analyzer puede detectar:

        manager_metrica  → MultiIndex (Manager, Métrica) × fechas
                           → resultado con (Métrica, Fecha) como columnas agrupadas
        solo_metrica     → Métrica × fechas, sin Manager
                           → resultado transpuesto (fechas como filas)
        tabla_generica   → DataFrame plano, primera col = etiquetas, resto = datos
        tabla_rotada     → ídem (ya normalizada por el parser)
        tabla_cruzada    → post-melt [Métrica, Header, Valor]
                           → re-pivotada a formato ancho
        clave_valor      → índice Nombre × col Valor
                           → devuelta filtrada

    Args:
        archivo : Ruta al archivo .xlsx.
        sheet   : Número de hoja (1-based, como en Excel UI).
        table   : ID de tabla dentro de la hoja (1-based).
        metrics : Métricas o etiquetas de fila a incluir (orden preservado).
                  Para tipos financieros: acepta exact, case-insensitive,
                  substring o fuzzy match.
                  Para tipos genéricos: coincidencia exacta contra la primera columna.
        columns : Columnas o headers a incluir (orden preservado).
                  Para tipos financieros: fechas exactas como aparecen.
                  Para tabla_cruzada: valores de 'Header' (ej. "2023", "Q1").
                  Para tabla_generica: nombres de columna exactos (ej. "Ventas").
                  Lista vacía [] → incluir todas las columnas disponibles.

    Returns:
        pd.DataFrame con el resultado pivotado/filtrado.

        Metadatos en df.attrs:
            df.attrs["sheet"]           → nombre de hoja (str)
            df.attrs["table_id"]        → id de la tabla (int)
            df.attrs["table_tipo"]      → tipo de tabla
            df.attrs["matched_metrics"] → nombres reales de métricas resueltos

    Raises:
        ValueError si sheet o table están fuera de rango.
        ValueError si alguna métrica no se encontró.
        ValueError si alguna columna no existe.
        ValueError si el tipo de tabla no está soportado.

    Ejemplos:
        # Tabla financiera con Manager y Métrica
        df = excel_build_table(
            "presupuesto.xlsx", sheet=2, table=1,
            metrics=["AUM", "ROA"],
            columns=["ene-24", "feb-24", "mar-24"],
        )

        # Tabla genérica (Producto | Ventas | Costos)
        df = excel_build_table(
            "reporte.xlsx", sheet=1, table=1,
            metrics=["A", "B"],
            columns=["Ventas", "Costos"],
        )

        # Tabla cruzada (ya normalizada por el parser a [Métrica, Header, Valor])
        df = excel_build_table(
            "pivot.xlsx", sheet=1, table=2,
            metrics=["AUM", "ROA"],
            columns=["2023", "2024"],
        )
    """
    # ── 1. Cargar tabla ──────────────────────────────────────────────────
    hojas  = _resolve_sheets(archivo, sheet)
    hoja   = hojas[0]
    tablas = detectar_todas_las_tablas(archivo, hoja)
    tabla  = _select_table(tablas, table, hoja)[0]
    df     = tabla["data"]
    tipo   = tabla.get("tipo", "tabla_generica")

    # ── 2. Determinar métricas y columnas disponibles según tipo ─────────
    available_metrics = _get_available_metrics(df, tipo)
    available_columns = _get_available_columns(df, tipo)

    # ── 3. Resolver métricas ─────────────────────────────────────────────
    # Tipos financieros: fuzzy matching completo (exact → case-i → substring → fuzzy).
    # Tipos genéricos: solo exacto, porque las etiquetas son datos del usuario,
    # no nombres de métricas con convenciones (no tiene sentido fuzzy "A" → "B").
    if tipo in ("manager_metrica", "solo_metrica"):
        matched_metrics = _match_metrics(metrics, available_metrics)
    else:
        # Coincidencia exacta — acumula errores igual que _match_metrics
        errors = [m for m in metrics if m not in available_metrics]
        if errors:
            raise ValueError(
                f"Las siguientes métricas no se encontraron: {errors}\n"
                f"Disponibles: {available_metrics}"
            )
        matched_metrics = list(metrics)

    # ── 4. Validar columnas ──────────────────────────────────────────────
    if columns:
        _validate_columns(columns, available_columns)

    # ── 5. Dispatch por tipo ─────────────────────────────────────────────
    if tipo == "manager_metrica":
        result = _pivot_manager_metrica(df, matched_metrics, columns)

    elif tipo == "solo_metrica":
        result = _pivot_solo_metrica(df, matched_metrics, columns)

    elif tipo in ("tabla_generica", "tabla_rotada"):
        result = _pivot_generic(df, matched_metrics, columns)

    elif tipo == "tabla_cruzada":
        result = _pivot_cruzada(df, matched_metrics, columns)

    elif tipo == "clave_valor":
        result = _pivot_key_value(df, matched_metrics, columns)

    else:
        raise ValueError(
            f"Tipo de tabla no soportado: '{tipo}'. "
            f"Tipos válidos: manager_metrica, solo_metrica, tabla_generica, "
            f"tabla_rotada, tabla_cruzada, clave_valor."
        )

    # ── 6. Metadatos ─────────────────────────────────────────────────────
    result.attrs["sheet"]           = hoja
    result.attrs["table_id"]        = table
    result.attrs["table_tipo"]      = tipo
    result.attrs["matched_metrics"] = matched_metrics

    return result