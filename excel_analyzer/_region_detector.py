"""
_region_detector.py
===================
Detector unificado de regiones de tabla.

Implementa el método de referencia (header-first) sobre las masks numpy
del SheetScanner.  Reemplaza tanto el antiguo RegionDetector (densidad de
filas) como el ProjectionRegionDetector (proyecciones) con un solo
algoritmo más robusto y predecible.

Filosofía (método de referencia):
  1. Mapa binario de densidad  →  ya precalculado en SheetScanner._mask_meaningful
  2. Firma de encabezado       →  fila de strings seguida de fila con números
  3. Título flotante           →  celda solitaria arriba de un encabezado
  4. Bloque de datos           →  empieza debajo del header, termina en fila vacía
  5. Tablas lado a lado        →  grupos de columnas separados por columnas vacías
  6. Tablas incrustadas        →  fila de strings en medio de datos numéricos

Todo el trabajo se hace sobre _mask_meaningful, _mask_numeric y _mask_datetime
del scanner — sin llamadas a openpyxl, sin loops celda-por-celda sobre ws.

Complejidad: O(filas × columnas) una sola pasada por señal.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np


@dataclass
class DetectedTable:
    """
    Región de tabla detectada con metadatos de estructura.
    Todos los índices son base-0.
    """
    fila_titulo:      int | None
    fila_header:      int
    fila_header_fin:  int
    fila_data_inicio: int
    fila_data_fin:    int
    grupos_columnas:  list[list[int]]
    titulo:           str | None = None


class HeaderFirstDetector:
    """
    Detecta tablas en una hoja usando la estrategia header-first:
    primero encuentra encabezados, luego delimita los bloques de datos
    alrededor de ellos.

    Señales de encabezado (de más fuerte a más débil):
      1. Fila donde ≥70% de celdas no-vacías son strings Y la fila
         siguiente tiene al menos un valor numérico.
      2. Fila con ≥ 2 fechas detectadas (header temporal).
      3. Fila con mezcla texto + temporal (header mixto financiero).

    Señales de título:
      - Fila con exactamente una celda no-vacía de tipo string, ubicada
        justo arriba de un encabezado detectado.

    Señales de fin de bloque:
      - Fila completamente vacía (densidad 0 en el rango de columnas).
      - Fin de la hoja.

    Señales de tablas incrustadas:
      - Dentro de un bloque de datos, fila donde ≥80% de celdas no-vacías
        son strings y la fila anterior tenía números.

    Señales de tablas lado a lado:
      - En las filas de encabezado, grupos de columnas separados por ≥ 1
        columna completamente vacía.
    """

    def __init__(self, sc: Any) -> None:
        self.sc = sc

    # ── API pública ────────────────────────────────────────────────────────

    def detectar_regiones(self) -> list[tuple[int, int, int | None, int | None]]:
        """
        Detecta todas las tablas de la hoja.

        Returns:
            Lista de (r0, r1, c0, c1) base-0.  c0/c1 pueden ser None si la
            tabla abarca todas las columnas con datos del rango de filas.
        """
        tablas = self._detectar_tablas()
        regiones: list[tuple[int, int, int | None, int | None]] = []

        for tabla in tablas:
            r0 = tabla.fila_titulo if tabla.fila_titulo is not None else tabla.fila_header
            r1 = tabla.fila_data_fin

            if len(tabla.grupos_columnas) <= 1:
                regiones.append((r0, r1, None, None))
            else:
                for grupo in tabla.grupos_columnas:
                    regiones.append((r0, r1, min(grupo), max(grupo)))

        return regiones

    # ── Detección interna ──────────────────────────────────────────────────

    def _detectar_tablas(self) -> list[DetectedTable]:
        """Escanea la hoja completa y retorna todas las tablas detectadas."""
        nrows = self.sc.nrows
        tablas: list[DetectedTable] = []
        filas_procesadas: set[int] = set()

        r = 0
        while r < nrows:
            if r in filas_procesadas:
                r += 1
                continue

            # ¿Esta fila es un encabezado clásico o temporal?
            if self._es_fila_encabezado(r) or self._es_fila_header_temporal(r):
                tabla = self._construir_tabla_desde_header(r, filas_procesadas)
                if tabla is not None:
                    tablas.append(tabla)
                    inicio = tabla.fila_titulo if tabla.fila_titulo is not None else tabla.fila_header
                    for f in range(inicio, tabla.fila_data_fin + 1):
                        filas_procesadas.add(f)
                    r = tabla.fila_data_fin + 1
                    continue

            r += 1

        # Fallback: regiones con datos que no quedaron cubiertas
        for r0, r1 in self._filas_no_cubiertas(filas_procesadas):
            densidad = int(self.sc._mask_meaningful[r0: r1 + 1, :].sum())
            if densidad >= 4 and (r1 - r0) >= 1:
                tabla = self._construir_tabla_fallback(r0, r1)
                if tabla is not None:
                    tablas.append(tabla)

        return tablas

    def _construir_tabla_desde_header(
        self, fila_header: int, filas_procesadas: set[int]
    ) -> DetectedTable | None:
        """Dado un encabezado, construye la tabla: título arriba, datos abajo, columnas."""
        nrows = self.sc.nrows

        # Si esta "fila header" es en realidad un título (merged: todos valores iguales),
        # buscar el header real en la fila siguiente.
        actual_header = fila_header
        titulo_fila = None
        titulo_texto = None

        if self._es_titulo(fila_header):
            titulo_fila = fila_header
            titulo_texto = self._extraer_titulo(fila_header)
            # Buscar el header real abajo
            for r_next in range(fila_header + 1, min(fila_header + 4, nrows)):
                if r_next in filas_procesadas:
                    break
                if self._es_fila_encabezado(r_next) or self._es_fila_header_temporal(r_next):
                    actual_header = r_next
                    break
            else:
                # No se encontró header debajo del título — tratar como tabla sin header
                return None
            if actual_header == fila_header:
                return None

        # Header multinivel: buscar filas complementarias arriba y abajo
        header_inicio = actual_header
        header_fin = actual_header

        for r_look in range(actual_header - 1, max(actual_header - 4, -1), -1):
            if r_look < 0 or r_look in filas_procesadas:
                break
            if titulo_fila is not None and r_look <= titulo_fila:
                break  # no pasar por encima del título
            if self._es_fila_header_complementaria(r_look):
                header_inicio = r_look
            else:
                break

        for r_look in range(actual_header + 1, min(actual_header + 4, nrows)):
            if r_look in filas_procesadas:
                break
            if (self._es_fila_encabezado(r_look)
                    or self._es_fila_header_temporal(r_look)
                    or self._es_fila_header_complementaria(r_look)):
                header_fin = r_look
            else:
                break

        # Si no encontramos título arriba del header de manera especial,
        # buscar con la lógica estándar (1 celda solitaria o merged)
        if titulo_fila is None:
            r_titulo = header_inicio - 1
            if r_titulo >= 0 and r_titulo not in filas_procesadas:
                if self._es_titulo(r_titulo):
                    titulo_fila = r_titulo
                    titulo_texto = self._extraer_titulo(r_titulo)

        # Rango de columnas del header
        c0_h, c1_h = self._bbox_cols_rango(header_inicio, header_fin)
        if c0_h is None:
            return None

        # Bloque de datos debajo del header
        fila_data_inicio = header_fin + 1
        fila_data_fin = self._encontrar_fin_bloque(
            fila_data_inicio, c0_h, c1_h, filas_procesadas
        )

        if fila_data_fin < fila_data_inicio:
            return None

        # Tabla incrustada: segundo header dentro del bloque de datos
        fila_incrustada = self._detectar_header_incrustado(
            fila_data_inicio, fila_data_fin, c0_h, c1_h
        )
        if fila_incrustada is not None:
            fila_data_fin = fila_incrustada - 1
            if fila_data_fin < fila_data_inicio:
                return None

        # Tablas lado a lado: grupos de columnas
        grupos = self._encontrar_grupos_columnas(header_inicio, header_fin)

        return DetectedTable(
            fila_titulo=titulo_fila,
            fila_header=header_inicio,
            fila_header_fin=header_fin,
            fila_data_inicio=fila_data_inicio,
            fila_data_fin=fila_data_fin,
            grupos_columnas=grupos,
            titulo=titulo_texto,
        )

    def _construir_tabla_fallback(self, r0: int, r1: int) -> DetectedTable | None:
        """Tabla sin header estándar: primera fila como header si mayoría strings."""
        first_row = None
        for r in range(r0, r1 + 1):
            if self.sc._mask_meaningful[r, :].any():
                first_row = r
                break
        if first_row is None:
            return None

        mask_m = self.sc._mask_meaningful[first_row, :]
        if mask_m.sum() == 0:
            return None

        vals = self.sc._mat[first_row, mask_m]
        n_str = sum(1 for v in vals if isinstance(v, str))
        is_header = n_str >= len(vals) * 0.5 and first_row < r1

        if is_header:
            return DetectedTable(
                fila_titulo=None, fila_header=first_row, fila_header_fin=first_row,
                fila_data_inicio=first_row + 1, fila_data_fin=r1,
                grupos_columnas=[],
            )
        return DetectedTable(
            fila_titulo=None, fila_header=first_row, fila_header_fin=first_row,
            fila_data_inicio=first_row, fila_data_fin=r1,
            grupos_columnas=[],
        )

    # ── Señales de encabezado ──────────────────────────────────────────────

    def _es_fila_encabezado(self, r: int) -> bool:
        """
        Fila de strings con ≥2 celdas, seguida (directa o indirectamente)
        de una fila con números.

        Mejora sobre el método de referencia: permite cadenas de headers
        consecutivos (título merged → header → datos), buscando números
        en las próximas 3 filas en vez de solo la inmediata.
        """
        if r >= self.sc.nrows - 1:
            return False
        mask_m = self.sc._mask_meaningful[r, :]
        n_meaningful = int(mask_m.sum())
        if n_meaningful < 2:
            return False
        # ≥70% strings
        vals = self.sc._mat[r, mask_m]
        n_str = sum(1 for v in vals if isinstance(v, str))
        if n_str < n_meaningful * 0.7:
            return False
        # Buscar números en las próximas 3 filas (permite headers consecutivos)
        for lookahead in range(1, min(4, self.sc.nrows - r)):
            if int(self.sc._mask_numeric[r + lookahead, :].sum()) >= 1:
                return True
        return False

    def _es_fila_header_temporal(self, r: int) -> bool:
        """Fila con ≥ 2 fechas detectadas."""
        if r >= self.sc.nrows:
            return False
        return int(self.sc._mask_datetime[r, :].sum()) >= 2

    def _es_fila_header_complementaria(self, r: int) -> bool:
        """Fila de años o fechas que complementa un header principal."""
        mask_m = self.sc._mask_meaningful[r, :]
        n_meaningful = int(mask_m.sum())
        if n_meaningful < 2:
            return False
        # ¿Tiene ≥2 fechas?
        if int(self.sc._mask_datetime[r, :].sum()) >= 2:
            return True
        # ¿Todos los números parecen años?
        vals = self.sc._mat[r, mask_m]
        nums = [v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
        if len(nums) >= 3:
            todos_anio = all(
                float(v) == int(v) and ((20 <= int(v) <= 99) or (2000 <= int(v) <= 2100))
                for v in nums
            )
            if todos_anio:
                return True
        return False

    def _es_titulo(self, r: int) -> bool:
        """
        Título = fila con una sola celda de texto, O fila donde todas las
        celdas tienen el mismo valor (merged cell expandido).

        Maneja el caso común de títulos en merged cells que el scanner
        expande a todas las columnas con el mismo string.
        """
        mask_m = self.sc._mask_meaningful[r, :]
        n_meaningful = int(mask_m.sum())
        if n_meaningful == 0:
            return False

        # Caso clásico: exactamente 1 celda
        if n_meaningful == 1:
            idx = np.where(mask_m)[0][0]
            return isinstance(self.sc._mat[r, idx], str)

        # Caso merged: todas las celdas tienen el mismo valor string
        vals = self.sc._mat[r, mask_m]
        if not all(isinstance(v, str) for v in vals):
            return False
        unique_vals = set(str(v).strip() for v in vals)
        return len(unique_vals) == 1

    def _extraer_titulo(self, r: int) -> str | None:
        """Texto del título de una fila."""
        mask_m = self.sc._mask_meaningful[r, :]
        if not mask_m.any():
            return None
        idx = np.where(mask_m)[0][0]
        v = self.sc._mat[r, idx]
        return str(v).strip() if v is not None else None

    # ── Límites de bloque ──────────────────────────────────────────────────

    def _encontrar_fin_bloque(
        self, fila_inicio: int, c0: int, c1: int, filas_procesadas: set[int]
    ) -> int:
        """Fin del bloque = primera fila vacía (tolerancia: 1 fila vacía interna)."""
        fila_fin = fila_inicio - 1
        vacias = 0
        for r in range(fila_inicio, self.sc.nrows):
            if r in filas_procesadas:
                break
            if int(self.sc._mask_meaningful[r, c0: c1 + 1].sum()) == 0:
                vacias += 1
                if vacias > 1:
                    break
            else:
                vacias = 0
                fila_fin = r
        return fila_fin

    def _detectar_header_incrustado(
        self, r_inicio: int, r_fin: int, c0: int, c1: int
    ) -> int | None:
        """Fila de strings en medio de datos numéricos = segundo header."""
        if r_fin - r_inicio < 3:
            return None
        for r in range(r_inicio + 2, r_fin + 1):
            mask_m = self.sc._mask_meaningful[r, c0: c1 + 1]
            n_m = int(mask_m.sum())
            if n_m < 2:
                continue
            vals = self.sc._mat[r, c0: c1 + 1][mask_m]
            n_str = sum(1 for v in vals if isinstance(v, str))
            if n_str < n_m * 0.8:
                continue
            # Fila anterior tenía números
            if r - 1 >= r_inicio and int(self.sc._mask_numeric[r - 1, c0: c1 + 1].sum()) >= 1:
                return r
        return None

    # ── Tablas lado a lado ─────────────────────────────────────────────────

    def _encontrar_grupos_columnas(
        self, r_header_inicio: int, r_header_fin: int
    ) -> list[list[int]]:
        """Grupos de columnas en el header, separados por columnas vacías."""
        mask_combined = np.zeros(self.sc.ncols, dtype=bool)
        for r in range(r_header_inicio, r_header_fin + 1):
            mask_combined |= self.sc._mask_meaningful[r, :]
        if not mask_combined.any():
            return []
        grupos: list[list[int]] = []
        grupo: list[int] = []
        for c in range(self.sc.ncols):
            if mask_combined[c]:
                grupo.append(c)
            else:
                if grupo:
                    grupos.append(grupo)
                    grupo = []
        if grupo:
            grupos.append(grupo)
        return grupos

    # ── Fallback: filas no cubiertas ───────────────────────────────────────

    def _filas_no_cubiertas(self, filas_procesadas: set[int]) -> list[tuple[int, int]]:
        """Rangos de filas con datos no cubiertas por ninguna tabla."""
        filas = [
            r for r in range(self.sc.nrows)
            if r not in filas_procesadas and self.sc._mask_meaningful[r, :].any()
        ]
        if not filas:
            return []
        grupos: list[tuple[int, int]] = []
        start = prev = filas[0]
        for f in filas[1:]:
            if f - prev > 2:
                grupos.append((start, prev))
                start = f
            prev = f
        grupos.append((start, prev))
        return grupos

    def _bbox_cols_rango(self, r0: int, r1: int) -> tuple[int | None, int | None]:
        """Columna mínima y máxima con datos en [r0, r1]."""
        sub = self.sc._mask_meaningful[r0: r1 + 1, :]
        cols = np.where(sub.any(axis=0))[0]
        if len(cols) == 0:
            return None, None
        return int(cols[0]), int(cols[-1])


# ── Alias para imports existentes en _core.py ──────────────────────────────
ProjectionRegionDetector = HeaderFirstDetector