"""Lectores robustos de Excel para el Facturador Henkel.

Principios:
- Emparejar columnas por nombre normalizado (sin acentos, mayúsculas, alfanum)
  para sobrevivir a re-exportaciones.
- Elegir dinámicamente la hoja con datos (no hardcodear "Sheet1").
- Excluir archivos de bloqueo de Excel (~$).
- Devolver DataFrames con nombres de columna estables y limpios.
"""
from __future__ import annotations

import calendar
import re
import unicodedata
import warnings
from pathlib import Path
from typing import Optional

import pandas as pd

import config


# ---------------------------------------------------------------------------
# Utilidades de normalización y búsqueda de columnas
# ---------------------------------------------------------------------------

def normalize(value) -> str:
    """Normaliza un texto estilo fnTexto del Power Query.

    NFKD quita acentos -> ASCII -> MAYÚSCULAS -> solo alfanuméricos.
    Se usa para emparejar nombres de columna y también valores de texto libre
    (p. ej. cliente <-> CEDI) donde conviene ignorar acentos, puntuación y mayúsculas.
    """
    if value is None:
        return ""
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^0-9A-Z]", "", text.upper())


def entrega_key(value) -> Optional[str]:
    """Normaliza un id de entrega/delivery a texto comparable.

    Los Delivery/ENTREGA vienen como enteros; los normalizamos a texto sin '.0'
    para que 934709063, 934709063.0 y '934709063' emparejen. Devuelve None si es nulo.
    """
    if value is None or pd.isna(value):
        return None
    try:
        f = float(value)
        if f.is_integer():
            return str(int(f))
    except (TypeError, ValueError):
        pass
    return str(value).strip().upper()


def find_column(columns, wanted: str, occurrence: int = 0) -> str:
    """Devuelve el nombre real de columna cuyo normalize() coincide con `wanted`.

    `occurrence=0` devuelve la 1ª coincidencia (resuelve el "Material" duplicado).
    Si no hay coincidencia exacta, cae a contención de substring normalizado.
    Lanza KeyError con un mensaje claro si no halla nada.
    """
    wanted_n = normalize(wanted)
    matches = [c for c in columns if normalize(c) == wanted_n]
    if not matches:
        # Fallback: contención (queremos "MATERIAL" dentro del nombre).
        matches = [c for c in columns if wanted_n and wanted_n in normalize(c)]
    if not matches:
        raise KeyError(
            f"No se encontró la columna '{wanted}' (normalizada '{wanted_n}'). "
            f"Columnas disponibles: {list(columns)}"
        )
    idx = min(occurrence, len(matches) - 1)
    return matches[idx]


def find_column_index(columns, wanted: str, occurrence: int = 0) -> int:
    """Igual que find_column pero devuelve la POSICIÓN (índice 0-based).

    Útil para leer con `usecols=` por posición: así se evita el problema de
    columnas con nombre duplicado (p.ej. dos "Material") y se leen solo las
    columnas necesarias, mucho más rápido.
    """
    wanted_n = normalize(wanted)
    positions = [i for i, c in enumerate(columns) if normalize(c) == wanted_n]
    if not positions:
        positions = [i for i, c in enumerate(columns) if wanted_n and wanted_n in normalize(c)]
    if not positions:
        raise KeyError(
            f"No se encontró la columna '{wanted}' (normalizada '{wanted_n}'). "
            f"Columnas disponibles: {list(columns)}"
        )
    idx = min(occurrence, len(positions) - 1)
    return positions[idx]


# ---------------------------------------------------------------------------
# Auditoría de tipos de datos (fechas/números) sobre la columna CRUDA
# ---------------------------------------------------------------------------

# Patrón "d/m/a..." para clasificar fechas en texto (p. ej. "02/21/2026").
_DATE_DMY_RE = re.compile(r"^\s*(\d{1,2})/(\d{1,2})/(\d{2,4})")


def _classify_date_str(value) -> str | None:
    """Clasifica un string de fecha como 'invertida' (mm/dd), 'invalida' o None (OK).

    Determinista (sin depender de la heurística de dayfirst de pandas, que es
    inestable en arreglos mixtos):
    - Si calza d/m/a: válida si día=a y mes=b son una fecha real (calendar.monthrange);
      si no, es 'invertida' cuando mes=a/día=b sí calza (a<=12, b>12); y 'invalida' si no.
    - Si no calza d/m/a (ISO "2026-06-05", texto...): válida si pandas la parsea.
    """
    s = str(value).strip()
    m = _DATE_DMY_RE.match(s)
    if m:
        a, b, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if yr < 100:  # año de 2 dígitos -> normalizar.
            yr += 2000 if yr < 70 else 1900
        if 1 <= b <= 12 and 1 <= a <= calendar.monthrange(yr, b)[1]:
            return None  # fecha dd/mm/aaaa válida
        if a <= 12 and 1 <= b <= calendar.monthrange(yr, a)[1]:
            return "invertida"  # no es dd/mm, pero calza como mm/dd (mes/día)
        return "invalida"
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            pd.to_datetime(s)
        return None
    except (ValueError, TypeError):
        return "invalida"


def audit_value_column(raw, col_label: str, kind: str, file: str, extra_valid=()) -> list[dict]:
    """Audita una columna CRUDA (antes del `errors="coerce"`) y devuelve issues de
    error (`{severity:"error", kind:"audit", file, msg}`) para celdas no-vacías con
    datos inválidos. Los vacíos legítimos (NaN/None/""/"nan") NUNCA se flaggean.

    - kind="number": celdas no-vacías que no son número (texto basura como "s").
      `extra_valid` admite literales extra válidos (p. ej. "-" en Ocupación = 0).
    - kind="date": sobre valores str no-vacíos detecta (a) inválidas (NaT al
      parsear) y (b) formato invertido mm/dd (string "d/m/..." con mes≤12 y día>12,
      p. ej. "02/21/2026"). Los datetime nativos de Excel (dtype datetime64) se
      aceptan sin auditar (Excel ya los validó como fecha real). Los **números**
      sueltos (p. ej. 45951 = número de serie de Excel sin formato de fecha) se
      marcan como error: `pd.to_datetime` los leería como ns desde 1970 y el paso
      se omitiría silenciosamente como "0 filas en el rango".
      Limitación: si Excel convirtió el valor a fecha real (locale en-US), el
      invertido ya no es detectable a posteriori.
    - kind="money": celdas no-vacías que `_parse_dinero` no convierte (para
      tarifa/costo de Otros, que admiten texto moneda "$ 13.933.333").

    Devuelve [] si todo está OK. Vectorizado; si la columna fecha es datetime64 no
    audita nada (rápido, caso normal en línea base).
    """
    if raw is None or len(raw) == 0:
        return []

    # Máscara de vacíos legítimos: NaN/None o strings en blanco ("", "nan", "NaT"…).
    stripped = raw.astype(str).str.strip()
    blank = raw.isna() | stripped.isin(["", "nan", "none", "<na>", "None", "NaT"])
    nonblank = ~blank

    def _issue(n: int, what: str, examples: list[str]) -> dict:
        ex = ", ".join(f"'{e}'" for e in examples[:2])
        return {
            "severity": "error",
            "kind": "audit",
            "file": file,
            "msg": f"{n} celda(s) {what} en '{col_label}'. Ej: {ex}",
        }

    if kind == "number":
        parsed = pd.to_numeric(raw, errors="coerce")
        bad = nonblank & parsed.isna()
        if extra_valid:
            ev = {str(x).strip() for x in extra_valid}
            bad = bad & ~stripped.isin(ev)
        n = int(bad.sum())
        if n:
            return [_issue(n, "con texto en columna numérica", stripped[bad].tolist())]

    elif kind == "money":
        bad = nonblank & raw.map(lambda v: _parse_dinero(v) is None)
        n = int(bad.sum())
        if n:
            return [_issue(n, "con valor no numérico/moneda inválido", stripped[bad].tolist())]

    elif kind == "date":
        # datetime64 puro -> Excel ya validó las fechas; nada que auditar.
        if pd.api.types.is_datetime64_any_dtype(raw):
            return []
        # Números sueltos en columna de fecha: Excel dejó la fecha como número de
        # serie (p. ej. 45951) sin formato de fecha. `pd.to_datetime` los leería
        # como nanosegundos desde 1970 -> fechas absurdas y el filtro por rango
        # daría 0 filas (el paso se omitiría silenciosamente como "0 filas en el
        # rango"). Se marca como error para que se corrija el formato en la fuente.
        num_mask = nonblank & raw.map(
            lambda v: pd.api.types.is_number(v) and not isinstance(v, bool)
        )
        n_num = int(num_mask.sum())
        if n_num:
            ex = ", ".join(f"'{e}'" for e in raw[num_mask].head(2).tolist())
            return [{
                "severity": "error",
                "kind": "audit",
                "file": file,
                "msg": (
                    f"{n_num} celda(s) con número en columna de fecha '{col_label}' "
                    f"(¿formato de fecha perdido en Excel? Ej: {ex}). Corrija el "
                    f"formato a fecha real y vuelva a generar."
                ),
            }]
        # Solo auditan los strings (fechas que Excel dejó como texto, p. ej.
        # "02/21/2026" que un locale dd/mm no reconoce y por tanto no convirtió).
        str_mask = nonblank & raw.map(lambda v: isinstance(v, str))
        if not str_mask.any():
            return []
        vals = raw[str_mask]
        classes = vals.map(_classify_date_str)
        inv_mask = classes == "invertida"
        bad_mask = classes == "invalida"
        n_inv = int(inv_mask.sum())
        n_bad = int(bad_mask.sum())
        if n_inv or n_bad:
            parts: list[str] = []
            examples: list[str] = []
            if n_inv:
                parts.append(f"{n_inv} con formato invertido sospechoso (mes/día)")
                examples += [str(v) for v in vals[inv_mask].head(2).tolist()]
            if n_bad:
                parts.append(f"{n_bad} inválida(s)/no parseable(s)")
                examples += [str(v) for v in vals[bad_mask].head(2).tolist()]
            ex = ", ".join(f"'{e}'" for e in examples[:2])
            return [{
                "severity": "error",
                "kind": "audit",
                "file": file,
                "msg": f"Columna de fecha '{col_label}': {' y '.join(parts)}. Ej: {ex}",
            }]

    return []


def excel_engine() -> str:
    """Motor de lectura preferido: 'calamine' (rápido, ~6x) si está disponible,
    si no 'openpyxl'. calamine reduce la lectura de los salidas grandes de ~95s a ~15s."""
    try:
        import python_calamine  # noqa: F401

        return "calamine"
    except ImportError:
        return "openpyxl"


def _exclude_locks(paths):
    """Filtra archivos de bloqueo de Excel (~$...) de una lista de rutas."""
    return [p for p in paths if not p.name.startswith(config.LOCK_PREFIX)]


# ---------------------------------------------------------------------------
# Lectura de salidas (salidas_cons / salidas_prof)
# ---------------------------------------------------------------------------

def pick_salidas_sheet(path: Path) -> str:
    """Devuelve el nombre de la hoja que contiene la cabecera MATERIAL.

    Evita hardcodear "Sheet1" y salta hojas vacías como "Hoja1".
    Solo lee la fila de cabecera de cada hoja (rápido).
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    for sheet in xls.sheet_names:
        header = pd.read_excel(path, sheet_name=sheet, nrows=0, engine=engine)
        if any(normalize(c) == "MATERIAL" for c in header.columns):
            return sheet
    raise KeyError(f"El archivo {path.name} no tiene ninguna hoja con columna 'Material'.")


# Nombres esperados (normalizados) de las columnas en el camino rápido.
_FAST_NORMALIZED = [normalize(config.SALIDAS_COLS[k]) for k in ("delivery", "cliente", "material", "cantidad", "fecha")]


def read_salidas(path: Path, area: str, audit: list | None = None) -> pd.DataFrame:
    """Lee un archivo de salidas y devuelve [delivery, cliente, material, cantidad, fecha, area].

    Estrategia de rendimiento:
    - Camino rápido: lee SOLO las 5 columnas en posiciones conocidas en un único
      parseo (calamine honra usecols). Verifica que los nombres cuadren.
    - Fallback robusto: si el layout cambió, lee la hoja completa y selecciona por
      nombre (maneja el "Material" duplicado tomando la 1ª ocurrencia).
    Cada archivo se abre una sola vez (calamine parsea el archivo entero en cada
    llamada, así que se minimizan los parseos).
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        # Camino rápido: posiciones conocidas.
        try:
            df = xls.parse(sheet_name=sheet, usecols=config.SALIDAS_FAST_POSITIONS)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        norm_cols = [normalize(c) for c in df.columns]
        if norm_cols == _FAST_NORMALIZED:
            return _select_salidas(df, area, audit, path.name)
        # El layout cambió: fallback robusto con lectura completa de esta hoja.
        full = xls.parse(sheet_name=sheet)
        fcols = list(full.columns)
        raw_cantidad = full[find_column(fcols, config.SALIDAS_COLS["cantidad"])]
        raw_fecha = full[find_column(fcols, config.SALIDAS_COLS["fecha"])]
        if audit is not None:
            audit.extend(audit_value_column(raw_cantidad, config.SALIDAS_COLS["cantidad"], "number", path.name))
            audit.extend(audit_value_column(raw_fecha, config.SALIDAS_COLS["fecha"], "date", path.name))
        out = pd.DataFrame()
        out["delivery"] = full[find_column(fcols, config.SALIDAS_COLS["delivery"])]
        out["cliente"] = full[find_column(fcols, config.SALIDAS_COLS["cliente"], occurrence=0)]
        out["material"] = full[find_column(fcols, config.SALIDAS_COLS["material"], occurrence=0)]
        out["cantidad"] = pd.to_numeric(raw_cantidad, errors="coerce")
        out["fecha"] = pd.to_datetime(raw_fecha, errors="coerce")
        out["area"] = area
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con las columnas esperadas. {last_err or ''}"
    )


def _select_salidas(df: pd.DataFrame, area: str, audit: list | None = None, file: str | None = None) -> pd.DataFrame:
    """Selecciona/limpia las columnas ya leídas por posición (camino rápido)."""
    out = pd.DataFrame()
    out["delivery"] = df.iloc[:, 0]
    out["cliente"] = df.iloc[:, 1]
    out["material"] = df.iloc[:, 2]
    out["cantidad"] = pd.to_numeric(df.iloc[:, 3], errors="coerce")
    out["fecha"] = pd.to_datetime(df.iloc[:, 4], errors="coerce")
    out["area"] = area
    if audit is not None and file is not None:
        # Auditar las columnas crudas (posición 3=cantidad, 4=fecha) antes del coerce.
        audit.extend(audit_value_column(df.iloc[:, 3], config.SALIDAS_COLS["cantidad"], "number", file))
        audit.extend(audit_value_column(df.iloc[:, 4], config.SALIDAS_COLS["fecha"], "date", file))
    return out


def find_salidas_files(base_dir: Optional[Path] = None):
    """Localiza los archivos salidas_cons* y salidas_prof* (excluye ~$).

    Devuelve lista de tuplas (ruta, area).
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    results = []
    for area, pattern in config.SALIDAS_GLOBS.items():
        folder = config.DIRS["consumer"] if area == "cons" else config.DIRS["profesional"]
        if not folder.exists():
            continue
        files = _exclude_locks(sorted(folder.glob(pattern)))
        for f in files:
            results.append((f, area))
    return results


# ---------------------------------------------------------------------------
# Lectura de ingresos (ingresos_cons / ingresos_prof): Paso 3
# ---------------------------------------------------------------------------

def read_ingresos(path: Path, area: str, audit: list | None = None) -> pd.DataFrame:
    """Lee un archivo de ingresos y devuelve
    [posting_date, cantidad, material, documento, referencia, area].

    Lee por nombre (find_column): los ingresos son archivos pequeños (~1.800 filas), así que
    no hace falta el camino rápido por posición de las salidas. Toma la primera hoja con las
    columnas esperadas. El filtro por rango (sobre `posting_date`) y el de `documento_cruce`
    se aplican en la pipeline, no aquí.
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        cols = list(df.columns)
        try:
            out = pd.DataFrame()
            raw_posting = df[find_column(cols, config.INGRESOS_COLS["posting_date"])]
            raw_cantidad = df[find_column(cols, config.INGRESOS_COLS["cantidad"])]
            if audit is not None:
                audit.extend(audit_value_column(raw_posting, config.INGRESOS_COLS["posting_date"], "date", path.name))
                audit.extend(audit_value_column(raw_cantidad, config.INGRESOS_COLS["cantidad"], "number", path.name))
            out["posting_date"] = pd.to_datetime(raw_posting, errors="coerce")
            out["cantidad"] = pd.to_numeric(raw_cantidad, errors="coerce")
            out["material"] = df[find_column(cols, config.INGRESOS_COLS["material"])]
            out["documento"] = df[find_column(cols, config.INGRESOS_COLS["documento"])]
            out["referencia"] = df[find_column(cols, config.INGRESOS_COLS["referencia"])]
        except KeyError:
            # Esta hoja no tiene las columnas de ingresos; probar la siguiente.
            continue
        out["material"] = out["material"].astype(str).str.strip().str.upper()
        out["area"] = area
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con las columnas esperadas "
        f"de ingresos. {last_err or ''}"
    )


def find_ingresos_files(base_dir: Optional[Path] = None):
    """Localiza los archivos ingresos_cons* y ingresos_prof* (excluye ~$).

    Devuelve lista de tuplas (ruta, area). Mismo molde que find_salidas_files: el área
    (cons/prof) se deriva del prefijo del nombre, no del contenido.
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    results = []
    for area, pattern in config.INGRESOS_GLOBS.items():
        folder = config.DIRS["consumer"] if area == "cons" else config.DIRS["profesional"]
        if not folder.exists():
            continue
        files = _exclude_locks(sorted(folder.glob(pattern)))
        for f in files:
            results.append((f, area))
    return results


# ---------------------------------------------------------------------------
# Lectura de ocupación (ocupacion_cons / ocupacion_prof): Paso 4
# ---------------------------------------------------------------------------

def read_ocupacion(path: Path, area: str, audit: list | None = None) -> pd.DataFrame:
    """Lee un archivo de ocupación y devuelve [fecha, almacen, tipo, ocupacion, area].

    Lee por nombre (find_column): los archivos son pequeños (~256-384 filas), así que no
    hace falta el camino rápido por posición de las salidas. Toma la primera hoja con las
    columnas esperadas. "Ocupación" y "% Ocupación" normalizan ambas a OCUPACION, así que se
    toma la 1ª ocurrencia (la de valor) con occurrence=0 — igual que el "Material" duplicado
    de salidas. `Tipo` y `Almacen` se dejan en MAYÚSCULAS/trim (fnTexto de logica.txt); el
    parseo numérico de `ocupacion` ("-" / "" -> 0) lo hace la pipeline (fnNumero).
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        cols = list(df.columns)
        try:
            out = pd.DataFrame()
            raw_fecha = df[find_column(cols, config.OCUPACION_COLS["fecha"])]
            raw_ocup = df[find_column(cols, config.OCUPACION_COLS["ocupacion"], occurrence=0)]
            if audit is not None:
                audit.extend(audit_value_column(raw_fecha, config.OCUPACION_COLS["fecha"], "date", path.name))
                # Ocupación admite "-" (= 0 en fnNumero): se trata como válido.
                audit.extend(audit_value_column(raw_ocup, config.OCUPACION_COLS["ocupacion"], "number", path.name, extra_valid=("-",)))
            out["fecha"] = raw_fecha
            out["almacen"] = df[find_column(cols, config.OCUPACION_COLS["almacen"])]
            out["tipo"] = df[find_column(cols, config.OCUPACION_COLS["tipo"])]
            out["ocupacion"] = raw_ocup
        except KeyError:
            # Esta hoja no tiene las columnas de ocupación; probar la siguiente.
            continue
        out["almacen"] = out["almacen"].astype(str).str.strip().str.upper()
        out["tipo"] = out["tipo"].astype(str).str.strip().str.upper()
        out["area"] = area
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con las columnas esperadas "
        f"de ocupación. {last_err or ''}"
    )


def find_ocupacion_files(base_dir: Optional[Path] = None):
    """Localiza los archivos ocupacion_cons* y ocupacion_prof* (excluye ~$).

    Devuelve lista de tuplas (ruta, area). Mismo molde que find_salidas_files /
    find_ingresos_files: el área (cons/prof) se deriva del prefijo del nombre.
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    results = []
    for area, pattern in config.OCUPACION_GLOBS.items():
        folder = config.DIRS["consumer"] if area == "cons" else config.DIRS["profesional"]
        if not folder.exists():
            continue
        files = _exclude_locks(sorted(folder.glob(pattern)))
        for f in files:
            results.append((f, area))
    return results


# ---------------------------------------------------------------------------
# Lectura de traslados (traslados_cons / traslados_prof): Paso 5
# ---------------------------------------------------------------------------

def read_traslados(path: Path, area: str, audit: list | None = None) -> pd.DataFrame:
    """Lee un archivo de traslados y devuelve [delivery, shu, con, area].

    Lee por nombre (find_column): los archivos son pequeños (~5-12 filas), así que no hace
    falta el camino rápido por posición de las salidas. Toma la primera hoja con las columnas
    esperadas. Solo se leen `delivery` (para el filtro "<> null" de logica.txt), `shu` y `con`
    (los 2 servicios); Ship-To/Nombre/No.Of Lines no se usan en la agregación. El filtro de
    Delivery nulo y el sumatorio por negocio los aplica la pipeline, no aquí.
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        cols = list(df.columns)
        try:
            out = pd.DataFrame()
            raw_shu = df[find_column(cols, config.TRASLADOS_COLS["shu"])]
            raw_con = df[find_column(cols, config.TRASLADOS_COLS["con"])]
            if audit is not None:
                audit.extend(audit_value_column(raw_shu, config.TRASLADOS_COLS["shu"], "number", path.name))
                audit.extend(audit_value_column(raw_con, config.TRASLADOS_COLS["con"], "number", path.name))
            out["delivery"] = df[find_column(cols, config.TRASLADOS_COLS["delivery"])]
            out["shu"] = raw_shu
            out["con"] = raw_con
        except KeyError:
            # Esta hoja no tiene las columnas de traslados; probar la siguiente.
            continue
        out["area"] = area
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con las columnas esperadas "
        f"de traslados. {last_err or ''}"
    )


def find_traslados_files(base_dir: Optional[Path] = None):
    """Localiza los archivos traslados_cons* y traslados_prof* (excluye ~$).

    Devuelve lista de tuplas (ruta, area). Mismo molde que find_ocupacion_files: el área
    (cons/prof) se deriva del prefijo del nombre, no del contenido.
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    results = []
    for area, pattern in config.TRASLADOS_GLOBS.items():
        folder = config.DIRS["consumer"] if area == "cons" else config.DIRS["profesional"]
        if not folder.exists():
            continue
        files = _exclude_locks(sorted(folder.glob(pattern)))
        for f in files:
            results.append((f, area))
    return results


# ---------------------------------------------------------------------------
# Lectura de maquila (maquila_cons / maquila_prof): Paso 6
# ---------------------------------------------------------------------------

def read_maquila(path: Path, area: str, audit: list | None = None) -> pd.DataFrame:
    """Lee un archivo de maquila y devuelve [posting_date, cantidad, material, area].

    Molde `read_ingresos` pero con solo 3 columnas (los maquila no usan documento ni
    referencia). Lee por nombre (find_column): los archivos son pequeños (~370-770
    filas), así que no hace falta el camino rápido por posición de las salidas. Toma la
    primera hoja con las columnas esperadas. El filtro por rango (sobre `posting_date`)
    y los cruces (huellas/idh) los aplica la pipeline, no aquí.
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        cols = list(df.columns)
        try:
            out = pd.DataFrame()
            raw_posting = df[find_column(cols, config.MAQUILA_COLS["posting_date"])]
            raw_cantidad = df[find_column(cols, config.MAQUILA_COLS["cantidad"])]
            if audit is not None:
                audit.extend(audit_value_column(raw_posting, config.MAQUILA_COLS["posting_date"], "date", path.name))
                audit.extend(audit_value_column(raw_cantidad, config.MAQUILA_COLS["cantidad"], "number", path.name))
            out["posting_date"] = pd.to_datetime(raw_posting, errors="coerce")
            out["cantidad"] = pd.to_numeric(raw_cantidad, errors="coerce")
            out["material"] = df[find_column(cols, config.MAQUILA_COLS["material"])]
        except KeyError:
            # Esta hoja no tiene las columnas de maquila; probar la siguiente.
            continue
        out["material"] = out["material"].astype(str).str.strip().str.upper()
        out["area"] = area
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con las columnas esperadas "
        f"de maquila. {last_err or ''}"
    )


def find_maquila_files(base_dir: Optional[Path] = None):
    """Localiza los archivos maquila_cons* y maquila_prof* (excluye ~$).

    Devuelve lista de tuplas (ruta, area). **Diferencia clave** con find_salidas_files
    y compañía: los `maquila_cons*` y `maquila_prof*` viven **ambos en la carpeta
    MAQUILA/** (no en CONSUMER/ y PROFESIONAL/), así que se hace glob de los dos
    patrones sobre la misma carpeta. El área (cons/prof) se deriva del prefijo.
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    folder = config.DIRS["maquila"]
    if not folder.exists():
        return []
    results = []
    for area, pattern in config.MAQUILA_GLOBS.items():
        files = _exclude_locks(sorted(folder.glob(pattern)))
        for f in files:
            results.append((f, area))
    return results


# ---------------------------------------------------------------------------
# Lectura de exportaciones (exportacion_cons / exportacion_prof): Paso 7
# ---------------------------------------------------------------------------

def read_exportacion(path: Path, area: str, audit: list | None = None) -> pd.DataFrame:
    """Lee un archivo de exportación y devuelve
    [delivery, fecha, material, cantidad, canal, area].

    Molde `read_ingresos`. Lee por nombre (find_column): los archivos no son enormes
    (~360-2.200 filas), así que no hace falta el camino rápido por posición de las
    salidas. Toma la primera hoja con las columnas esperadas. El filtro `Delivery <> null`
    y el de rango (sobre `fecha` = Fecha factura) los aplica la pipeline, no aquí.
    `canal` se conserva tal cual (MAYÚSCULAS/trim): forma parte del nombre del servicio
    (`servicio = base & " " & canal` en logica.txt). `Texto breve de material` no se lee.
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        cols = list(df.columns)
        try:
            out = pd.DataFrame()
            raw_fecha = df[find_column(cols, config.EXPORTACION_COLS["fecha"])]
            raw_cantidad = df[find_column(cols, config.EXPORTACION_COLS["cantidad"])]
            if audit is not None:
                audit.extend(audit_value_column(raw_fecha, config.EXPORTACION_COLS["fecha"], "date", path.name))
                audit.extend(audit_value_column(raw_cantidad, config.EXPORTACION_COLS["cantidad"], "number", path.name))
            out["delivery"] = df[find_column(cols, config.EXPORTACION_COLS["delivery"])]
            out["fecha"] = pd.to_datetime(raw_fecha, errors="coerce")
            out["material"] = df[find_column(cols, config.EXPORTACION_COLS["material"])]
            out["cantidad"] = pd.to_numeric(raw_cantidad, errors="coerce")
            out["canal"] = df[find_column(cols, config.EXPORTACION_COLS["canal"])]
        except KeyError:
            # Esta hoja no tiene las columnas de exportación; probar la siguiente.
            continue
        out["material"] = out["material"].astype(str).str.strip().str.upper()
        out["canal"] = out["canal"].astype(str).str.strip().str.upper()
        out["area"] = area
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con las columnas esperadas "
        f"de exportación. {last_err or ''}"
    )


def find_exportacion_files(base_dir: Optional[Path] = None):
    """Localiza los archivos exportacion_cons* y exportacion_prof* (excluye ~$).

    Devuelve lista de tuplas (ruta, area). Molde `find_maquila_files`: los
    `exportacion_cons*` y `exportacion_prof*` viven **ambos en la carpeta
    EXPORTACIONES/** (no en CONSUMER/ y PROFESIONAL/), así que se hace glob de los dos
    patrones sobre la misma carpeta. El área (cons/prof) se deriva del prefijo del
    nombre. `paletizado*`/`trincaje*` (que también viven en EXPORTACIONES/) NO se
    procesan: logica.txt sólo admite ficheros que empiecen por "exportacion".
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    folder = config.DIRS["exportaciones"]
    if not folder.exists():
        return []
    results = []
    for area, pattern in config.EXPORTACION_GLOBS.items():
        files = _exclude_locks(sorted(folder.glob(pattern)))
        for f in files:
            results.append((f, area))
    return results


# ---------------------------------------------------------------------------
# Lectura de etiquetas (etiquetas_*): Paso 8
# ---------------------------------------------------------------------------

def read_etiquetas(path: Path, audit: list | None = None) -> pd.DataFrame:
    """Lee un archivo de etiquetas y devuelve [cajas, tipo].

    Molde `read_traslados`: archivo pequeño (~50 filas), lectura por nombre
    (find_column). Sólo se leen `cajas` (la medida que se suma -> valor) y `tipo`
    (que mapea al nombre del servicio). `fecha` y `entrega` existen en el fuente pero
    el PQ las descarta (la `fecha` final viene del nombre del archivo), así que no se
    leen. negocio es siempre "CONSUMER" (constante del PQ), así que no hace falta área.
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        cols = list(df.columns)
        try:
            out = pd.DataFrame()
            raw_cajas = df[find_column(cols, config.ETIQUETAS_COLS["cajas"])]
            if audit is not None:
                audit.extend(audit_value_column(raw_cajas, config.ETIQUETAS_COLS["cajas"], "number", path.name))
            out["cajas"] = raw_cajas
            out["tipo"] = df[find_column(cols, config.ETIQUETAS_COLS["tipo"])]
        except KeyError:
            # Esta hoja no tiene las columnas de etiquetas; probar la siguiente.
            continue
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con las columnas esperadas "
        f"de etiquetas. {last_err or ''}"
    )


def find_etiquetas_files(base_dir: Optional[Path] = None) -> list[Path]:
    """Localiza los archivos etiquetas* (excluye ~$).

    El fichero `etiquetas_<fecha>.xlsx` vive en la carpeta MAQUILA/ (sin separación
    cons/prof), así que se hace glob de `etiquetas*` sobre esa carpeta. negocio es
    siempre "CONSUMER" (constante del PQ), así que no se devuelve área (a diferencia
    del resto de lectores, que devuelven tuplas (ruta, area)).
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    folder = config.DIRS["maquila"]
    if not folder.exists():
        return []
    return _exclude_locks(sorted(folder.glob(config.ETIQUETAS_GLOB)))


# ---------------------------------------------------------------------------
# Lectura de paletizado (paletizado_*): Paso 9
# ---------------------------------------------------------------------------

def read_paletizado(path: Path, audit: list | None = None) -> pd.DataFrame:
    """Lee un archivo de paletizado y devuelve [area, despacho, total, canal].

    Molde `read_traslados`: archivo pequeño (~90 filas), lectura por nombre
    (find_column). Sólo se leen `area` (-> negocio vía PALETIZADO_AREA_MAP), `despacho`
    (filtro "<> null" de logica.txt), `total` (la medida que se suma -> valor) y `canal`
    (forma parte del nombre del servicio). Las demás columnas no se usan en la
    agregación. El mapeo AREA -> negocio y el sumatorio por (negocio, canal) los aplica
    la pipeline, no aquí.
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        cols = list(df.columns)
        try:
            out = pd.DataFrame()
            raw_total = df[find_column(cols, config.PALETIZADO_COLS["total"])]
            if audit is not None:
                audit.extend(audit_value_column(raw_total, config.PALETIZADO_COLS["total"], "number", path.name))
            out["area"] = df[find_column(cols, config.PALETIZADO_COLS["area"])]
            out["despacho"] = df[find_column(cols, config.PALETIZADO_COLS["despacho"])]
            out["total"] = pd.to_numeric(raw_total, errors="coerce")
            out["canal"] = df[find_column(cols, config.PALETIZADO_COLS["canal"])]
        except KeyError:
            # Esta hoja no tiene las columnas de paletizado; probar la siguiente.
            continue
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con las columnas esperadas "
        f"de paletizado. {last_err or ''}"
    )


def find_paletizado_files(base_dir: Optional[Path] = None) -> list[Path]:
    """Localiza los archivos paletizado* (excluye ~$).

    El fichero `paletizado_<fecha>.xlsx` vive en la carpeta EXPORTACIONES/ (junto a
    `exportacion_*`/`trincaje_*`), así que se hace glob de `paletizado*` sobre esa
    carpeta. El `negocio` se deriva de la columna AREA (no del nombre), así que no se
    devuelve área. `trincaje*` no se procesa aquí (logica.txt lo admite en su propia
    query).
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    folder = config.DIRS["exportaciones"]
    if not folder.exists():
        return []
    return _exclude_locks(sorted(folder.glob(config.PALETIZADO_GLOB)))


# ---------------------------------------------------------------------------
# Lectura de trincaje (trincaje_*): Paso 10
# ---------------------------------------------------------------------------

def read_trincaje(path: Path) -> pd.DataFrame:
    """Lee un archivo de trincaje y devuelve [despacho].

    Molde `read_etiquetas`/`read_paletizado`: archivo pequeño (~30 filas), lectura por
    nombre (find_column). El PQ lee despacho/exp/destino/contenedor/n-cont, pero tras el
    filtro `despacho <> null` sólo conserva el NOMBRE del archivo (la `fecha` final viene
    del nombre), así que aquí sólo se lee `despacho` (para el filtro "<> null"). El conteo
    `× 2` -> valor lo aplica la pipeline, no aquí.
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        cols = list(df.columns)
        try:
            out = pd.DataFrame()
            out["despacho"] = df[find_column(cols, config.TRINCAJE_COLS["despacho"])]
        except KeyError:
            # Esta hoja no tiene la columna despacho; probar la siguiente.
            continue
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con la columna esperada "
        f"de trincaje. {last_err or ''}"
    )


def find_trincaje_files(base_dir: Optional[Path] = None) -> list[Path]:
    """Localiza los archivos trincaje* (excluye ~$).

    El fichero `trincaje_<fecha>.xlsx` vive en la carpeta EXPORTACIONES/ (junto a
    `exportacion_*`/`paletizado_*`), así que se hace glob de `trincaje*` sobre esa
    carpeta. negocio es siempre "CONSUMER" (constante del PQ), así que no se devuelve
    área (igual que etiquetas).
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    folder = config.DIRS["exportaciones"]
    if not folder.exists():
        return []
    return _exclude_locks(sorted(folder.glob(config.TRINCAJE_GLOB)))


# ---------------------------------------------------------------------------
# Lectura de planta (planta_*): Paso 11
# ---------------------------------------------------------------------------

def read_planta(path: Path, audit: list | None = None) -> pd.DataFrame:
    """Lee un archivo de planta y devuelve [consumer, profesional].

    Molde `read_etiquetas`: archivo pequeño (~30 filas), lectura por nombre
    (find_column). El PQ elimina las cols `fecha` y `semana` del archivo (la `fecha`
    final viene del NOMBRE del archivo) y unpivotea las dos cols de estibas, así que
    aquí sólo se leen `estibas_consumer` (-> negocio CONSUMER) y `estibas_profesional`
    (-> negocio PROFESIONAL). El sumatorio por negocio (que forma `valor`) lo aplica la
    pipeline, no aquí. `semana`/`fecha` no se leen.
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        cols = list(df.columns)
        try:
            out = pd.DataFrame()
            raw_consumer = df[find_column(cols, config.PLANTA_COLS["consumer"])]
            raw_profesional = df[find_column(cols, config.PLANTA_COLS["profesional"])]
            if audit is not None:
                audit.extend(audit_value_column(raw_consumer, config.PLANTA_COLS["consumer"], "number", path.name))
                audit.extend(audit_value_column(raw_profesional, config.PLANTA_COLS["profesional"], "number", path.name))
            out["consumer"] = pd.to_numeric(raw_consumer, errors="coerce")
            out["profesional"] = pd.to_numeric(raw_profesional, errors="coerce")
        except KeyError:
            # Esta hoja no tiene las columnas de planta; probar la siguiente.
            continue
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con las columnas esperadas "
        f"de planta (estibas_consumer/estibas_profesional). {last_err or ''}"
    )


def find_planta_files(base_dir: Optional[Path] = None) -> list[Path]:
    """Localiza los archivos planta* (excluye ~$).

    El fichero `planta_<fecha>.xlsx` cubre AMBOS negocios en un solo archivo (el
    `negocio` viene del nombre de la columna, no de la carpeta), pero en la práctica
    vive en CONSUMER/ o PROFESIONAL/. Para ser fiel a `logica.txt` (que busca cualquier
    `planta*` en FACTURACION, sin restringir subcarpeta) se hace glob de `planta*` sobre
    AMBAS carpetas de negocio y se devuelven todas las rutas (dedupe por ruta). No se
    devuelve área (el negocio lo da la columna, no el nombre/ubicación).
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    found: list[Path] = []
    for key in ("consumer", "profesional"):
        folder = config.DIRS.get(key)
        if folder and folder.exists():
            found.extend(folder.glob(config.PLANTA_GLOB))
    # Dedupe conservando el orden (mismos nombres en ambas carpetas sería raro).
    seen, unique = set(), []
    for p in found:
        rp = p.resolve()
        if rp not in seen:
            seen.add(rp)
            unique.append(p)
    return _exclude_locks(sorted(unique))


# ---------------------------------------------------------------------------
# Lectura de material (ocupacionMaterial_*): Paso 12
# ---------------------------------------------------------------------------

def read_material(path: Path, audit: list | None = None) -> pd.DataFrame:
    """Lee un archivo de ocupacionMaterial y devuelve [fecha, valor].

    Molde `read_ocupacion` (simplificado): archivo pequeño (~32 filas), lectura por nombre
    (find_column). El PQ sólo usa `fecha` (día de almacenamiento) y `cant_bodega_8` (la
    única medida). El rango del usuario sobre `fecha` y el `ceil(promedio)` -> valor los
    aplica la pipeline, no aquí. A diferencia de Ocupación no hay Almacen/Tipo/Ocupación:
    una sola medida y servicio/negocio fijos.
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        cols = list(df.columns)
        try:
            out = pd.DataFrame()
            raw_fecha = df[find_column(cols, config.MATERIAL_COLS["fecha"])]
            raw_valor = df[find_column(cols, config.MATERIAL_COLS["valor"])]
            if audit is not None:
                audit.extend(audit_value_column(raw_fecha, config.MATERIAL_COLS["fecha"], "date", path.name))
                audit.extend(audit_value_column(raw_valor, config.MATERIAL_COLS["valor"], "number", path.name))
            out["fecha"] = pd.to_datetime(raw_fecha, errors="coerce")
            out["valor"] = pd.to_numeric(raw_valor, errors="coerce")
        except KeyError:
            # Esta hoja no tiene las columnas de material; probar la siguiente.
            continue
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con las columnas esperadas "
        f"de ocupacionMaterial (fecha/cant_bodega_8). {last_err or ''}"
    )


def find_material_files(base_dir: Optional[Path] = None) -> list[Path]:
    """Localiza los archivos ocupacionMaterial* (excluye ~$).

    El fichero `ocupacionMaterial_<fecha>.xlsx` vive en la carpeta OTROS/ (junto a
    `otros_*`), así que se hace glob de `ocupacionMaterial*` sobre esa carpeta. Negocio/
    facturador/servicio son constantes (no se derivan del nombre), así que no se devuelve
    área.
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    folder = config.DIRS["otros"]
    if not folder.exists():
        return []
    return _exclude_locks(sorted(folder.glob(config.MATERIAL_GLOB)))


# ---------------------------------------------------------------------------
# Lectura de falabella (falabella_cons / falabella_prof): Paso 13
# ---------------------------------------------------------------------------

def read_falabella(path: Path, area: str) -> pd.DataFrame:
    """Lee un archivo de falabella y devuelve [entrega, area].

    Molde `read_traslados`: archivo pequeño (~114 filas), lectura por nombre
    (find_column). El PQ trae 8 columnas (DIGITO VERIFICADOR, NIT HENKEL,
    CONSECUTIVO…, NUMERACION, LARGO 18 DIGITOS, TIENDA, Entrega, MES), pero tras el
    filtro `Entrega <> null` sólo conserva el NOMBRE del archivo (la `fecha` final viene
    del nombre), así que aquí sólo se lee `entrega` (para el filtro "<> null"). El conteo
    de filas por negocio -> valor lo aplica la pipeline, no aquí. `MES` no se usa.
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        cols = list(df.columns)
        try:
            out = pd.DataFrame()
            out["entrega"] = df[find_column(cols, config.FALABELLA_COLS["entrega"])]
        except KeyError:
            # Esta hoja no tiene la columna Entrega; probar la siguiente.
            continue
        out["area"] = area
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con la columna esperada "
        f"de falabella (Entrega). {last_err or ''}"
    )


def find_falabella_files(base_dir: Optional[Path] = None):
    """Localiza los archivos falabella_cons* y falabella_prof* (excluye ~$).

    Devuelve lista de tuplas (ruta, area). Mismo molde que find_traslados_files: el
    área (cons/prof) se deriva del prefijo del nombre (falabella_cons* -> CONSUMER,
    falabella_prof* -> PROFESIONAL), no del contenido.
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    results = []
    for area, pattern in config.FALABELLA_GLOBS.items():
        folder = config.DIRS["consumer"] if area == "cons" else config.DIRS["profesional"]
        if not folder.exists():
            continue
        files = _exclude_locks(sorted(folder.glob(pattern)))
        for f in files:
            results.append((f, area))
    return results


# ---------------------------------------------------------------------------
# Lectura de otros (otros_<fecha>): Paso 14
# ---------------------------------------------------------------------------

def _parse_dinero(value):
    """Convierte un valor de `tarifa`/`costo` (que puede venir como texto moneda) a float.

    El archivo otros_* trae estos campos como `object` (mezcla): números puros
    (`62895.4147`, `93759.7371`) y texto moneda colombiano (`"$ 13.933.333"`,
    `"$   6.348.884"`). Power Query los convierte con `type number` (línea 33); aquí se
    replica la limpieza:

    - int/float -> float(value) (NaN/None -> None).
    - str -> quitar `$`, espacios; si hay "," y "." asumir formato europeo (punto=miles,
      coma=decimal); si hay >1 "." son miles; 1 "." -> decimal. No parseable -> None.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if s in ("", "-", "nan", "None"):
        return None
    s = s.replace("$", "").replace(" ", "").replace("\xa0", "")
    if "," in s and "." in s:
        # p. ej. "13.933.333,50" -> punto=miles, coma=decimal.
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        # sólo coma -> decimal.
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        # varios puntos -> miles (p. ej. "13.933.333").
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def read_otros(path: Path, audit: list | None = None) -> pd.DataFrame:
    """Lee un archivo `otros_*` y devuelve las 9 columnas que el bot anexa al Excel.

    A diferencia del resto de pasos, este archivo es una tabla de servicios **pre-armados**
    (cánones, horas extras, cargues expo, descargues…): ya trae `valor`/`tarifa`/`costo`
    calculados. No se calcula nada aquí; sólo se leen y se parsean a número. Las columnas
    `fecha` y `tabla` del contenido **no se leen**: el PQ descarta `fecha` (línea 31; la
    final viene del NOMBRE del archivo, que el bot reemplaza por `periodo`), y `tabla` se
    fija a "OTROS" para todos los registros (decisión del usuario). Molde `read_etiquetas`
    (lectura por nombre con find_column; sin área, pues el `negocio` viene del propio archivo).

    Lanza KeyError si faltan columnas obligatorias (OTROS_REQUIRED) -> la pipeline lo
    convierte en BlockingError.
    """
    engine = excel_engine()
    xls = pd.ExcelFile(path, engine=engine)
    last_err = None
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet)
        except Exception as exc:
            last_err = exc
            continue
        if df is None or df.empty:
            continue
        cols = list(df.columns)
        try:
            out = pd.DataFrame()
            raw_valor = df[find_column(cols, config.OTROS_COLS["valor"])]
            raw_tarifa = df[find_column(cols, config.OTROS_COLS["tarifa"])]
            raw_costo = df[find_column(cols, config.OTROS_COLS["costo"])]
            if audit is not None:
                audit.extend(audit_value_column(raw_valor, config.OTROS_COLS["valor"], "number", path.name))
                # tarifa/costo admiten texto moneda ("$ 13.933.333"): money = no parseable por _parse_dinero.
                audit.extend(audit_value_column(raw_tarifa, config.OTROS_COLS["tarifa"], "money", path.name))
                audit.extend(audit_value_column(raw_costo, config.OTROS_COLS["costo"], "money", path.name))
            out["negocio"] = df[find_column(cols, config.OTROS_COLS["negocio"])]
            out["negocio_facturador"] = df[find_column(cols, config.OTROS_COLS["negocio_facturador"])]
            out["servicio"] = df[find_column(cols, config.OTROS_COLS["servicio"])]
            out["valor"] = pd.to_numeric(raw_valor, errors="coerce")
            out["proceso_extendido"] = df[find_column(cols, config.OTROS_COLS["proceso_extendido"])]
            out["macro_proceso"] = df[find_column(cols, config.OTROS_COLS["macro_proceso"])]
            out["proceso_abreviado"] = df[find_column(cols, config.OTROS_COLS["proceso_abreviado"])]
            out["tarifa"] = raw_tarifa.map(_parse_dinero)
            out["costo"] = raw_costo.map(_parse_dinero)
        except KeyError:
            # Esta hoja no tiene las columnas esperadas; probar la siguiente.
            continue
        return out
    raise KeyError(
        f"El archivo {path.name} no tiene una hoja legible con las columnas esperadas "
        f"de otros (negocio/negocio_facturador/servicio/valor/proceso_*/tarifa/costo). "
        f"{last_err or ''}"
    )


def find_otros_files(base_dir: Optional[Path] = None) -> list[Path]:
    """Localiza los archivos `otros*` (excluye ~$).

    El fichero `otros_<fecha>.xlsx` vive en la carpeta OTROS/ (junto a
    `ocupacionMaterial_*`), así que se hace glob de `otros*` sobre esa carpeta. Como el
    `negocio` viene del propio archivo (no del nombre ni de la carpeta), no se devuelve área.
    """
    base = Path(base_dir) if base_dir else config.BASE_DIR
    folder = config.DIRS["otros"]
    if not folder.exists():
        return []
    return _exclude_locks(sorted(folder.glob(config.OTROS_GLOB)))


# ---------------------------------------------------------------------------
# Lectura de tablas de lookup
# ---------------------------------------------------------------------------

def read_idh(base_dir: Optional[Path] = None) -> tuple[pd.DataFrame, int]:
    """Lee idh_especiales -> ([material, negocio], n_duplicados).

    Si hubiera materiales duplicados, conserva el primero y reporta el conteo.
    """
    path = config.FILES["idh_especiales"]
    df = pd.read_excel(path, engine=excel_engine())
    cols = list(df.columns)
    out = pd.DataFrame()
    out["material"] = df[find_column(cols, config.IDH_COLS["material"])]
    out["negocio"] = df[find_column(cols, config.IDH_COLS["negocio"])]
    out["material"] = out["material"].astype(str).str.strip().str.upper()
    out["negocio"] = out["negocio"].astype(str).str.strip().str.upper()
    dup_count = int(out["material"].duplicated(keep=False).sum())
    return out.drop_duplicates(subset=["material"], keep="first"), dup_count


def read_equivalencias() -> tuple[pd.DataFrame, int]:
    """Lee equivalencias_almacenamiento.xlsx -> ([archivo(Tipo), conversion], n_duplicados).

    Molde read_idh. `archivo` (p. ej. "PALLET (MEDIO)") es la llave de cruce con el `Tipo`
    de ocupación; `conversion` (p. ej. "MEDIO PALLET") forma el servicio y se conserva tal
    cual (MAYÚSCULAS/trim, SIN normalizar agresivamente: sus espacios son significativos).
    Si hubiera Tipos duplicados, conserva el primero y reporta el conteo.
    """
    path = config.FILES["equivalencias"]
    df = pd.read_excel(path, engine=excel_engine())
    cols = list(df.columns)
    out = pd.DataFrame()
    out["archivo"] = df[find_column(cols, config.EQUIVALENCIAS_COLS["archivo"])]
    out["conversion"] = df[find_column(cols, config.EQUIVALENCIAS_COLS["conversion"])]
    out["archivo"] = out["archivo"].astype(str).str.strip().str.upper()
    out["conversion"] = out["conversion"].astype(str).str.strip().str.upper()
    out = out[out["archivo"].notna() & (out["archivo"] != "")]
    dup_count = int(out["archivo"].duplicated(keep=False).sum())
    return out.drop_duplicates(subset=["archivo"], keep="first").reset_index(drop=True), dup_count


def read_huellas() -> tuple[pd.DataFrame, int]:
    """Lee huellas.xlsx (cabecera en fila 2) -> ([producto, pallet, caja], n_duplicados).

    Desempata los Producto duplicados quedándose con el PRIMER valor único que tenga
    dato correcto (pallet y caja presentes, numéricos y > 0); los registros que vienen
    detrás se descartan. Esto afecta a TODOS los pasos que cruzan huellas (SALIDAS,
    INGRESOS, MAQUILA, EXPORTACIONES).

    El archivo se relee en cada ejecución porque se actualiza todos los meses y muy
    probablemente llegue con duplicados. La regla "primer valor correcto" evita que un
    batch erróneo appended al final (p. ej. el 2025-10-11 con valores absurdos) pise la
    huella buena: la primera aparición de cada producto en el orden del archivo es la
    válida.
    """
    import numpy as np

    path = config.FILES["huellas"]
    df = pd.read_excel(path, header=config.HUELLAS_HEADER_ROW, engine=excel_engine())
    cols = list(df.columns)
    out = pd.DataFrame()
    out["producto"] = df[find_column(cols, config.HUELLAS_COLS["producto"])]
    out["pallet"] = pd.to_numeric(df[find_column(cols, config.HUELLAS_COLS["pallet"])], errors="coerce")
    out["caja"] = pd.to_numeric(df[find_column(cols, config.HUELLAS_COLS["caja"])], errors="coerce")
    # Normaliza la llave ANTES del desempate (mismo criterio que el cruce downstream).
    out["producto"] = out["producto"].astype(str).str.strip().str.upper()

    # Dato correcto: pallet y caja presentes y > 0.
    dup_count = int(out["producto"].duplicated(keep=False).sum())
    out["_valid"] = (
        out["pallet"].notna() & (out["pallet"] > 0)
        & out["caja"].notna() & (out["caja"] > 0)
    )
    out["_ord"] = np.arange(len(out))
    # Filas válidas primero y, dentro de cada grupo, por orden del archivo: drop_duplicates
    # keep="first" deja el primer valor único correcto. Si un producto no tuviera ninguna
    # fila válida, se queda con su primera aparición (pallet/caja NaN, como antes).
    out = (
        out.sort_values(["_valid", "_ord"], ascending=[False, True], kind="stable")
        .drop(columns=["_valid", "_ord"])
        .drop_duplicates(subset=["producto"], keep="first")
    )

    return out.reset_index(drop=True), dup_count


def read_tarifas() -> pd.DataFrame:
    """Lee tarifas -> [servicio, valor, um, fecha] como tabla de referencia."""
    path = config.FILES["tarifas"]
    df = pd.read_excel(path, engine=excel_engine())
    cols = list(df.columns)
    out = pd.DataFrame()
    out["servicio"] = df[find_column(cols, config.TARIFAS_COLS["servicio"])]
    out["valor"] = pd.to_numeric(df[find_column(cols, config.TARIFAS_COLS["valor"])], errors="coerce")
    out["um"] = df[find_column(cols, config.TARIFAS_COLS["um"])]
    try:
        out["fecha"] = pd.to_datetime(df[find_column(cols, config.TARIFAS_COLS["fecha"])], errors="coerce")
    except KeyError:
        out["fecha"] = pd.NaT
    return out


def find_adicionales_files(base_dir: Optional[Path] = None) -> list[Path]:
    """Localiza los archivos adicionales* en CONSUMER/ y PROFESIONAL/ (excluye ~$)."""
    base = Path(base_dir) if base_dir else config.BASE_DIR
    results: list[Path] = []
    for key in ("consumer", "profesional"):
        folder = config.DIRS[key]
        if not folder.exists():
            continue
        results.extend(_exclude_locks(sorted(folder.glob(config.ADICIONALES_GLOB))))
    return results


def read_adicionales(base_dir: Optional[Path] = None) -> tuple[pd.DataFrame, int, list[str]]:
    """Lee todos los adicionales* -> ([entrega(key), tipo], n_duplicados, archivos_leidos).

    Concatena adicionales_cons* y adicionales_prof*. La columna `entrega` se normaliza
    con entrega_key() para emparejar con el Delivery de salidas. Lanza FileNotFoundError
    si no hay archivos legibles.
    """
    paths = find_adicionales_files(base_dir)
    if not paths:
        raise FileNotFoundError("No se encontraron archivos adicionales* en CONSUMER/ y PROFESIONAL/.")
    frames = []
    for path in paths:
        df = pd.read_excel(path, engine=excel_engine())
        cols = list(df.columns)
        out = pd.DataFrame()
        out["entrega"] = df[find_column(cols, config.ADICIONALES_COLS["entrega"])]
        out["tipo"] = df[find_column(cols, config.ADICIONALES_COLS["tipo"])]
        frames.append(out)
    all_df = pd.concat(frames, ignore_index=True)
    all_df["entrega"] = all_df["entrega"].map(entrega_key)
    all_df["tipo"] = all_df["tipo"].astype(str).str.strip().str.upper()
    all_df = all_df[all_df["entrega"].notna()]
    dup_count = int(all_df["entrega"].duplicated(keep=False).sum())
    out = all_df.drop_duplicates(subset=["entrega"], keep="first").reset_index(drop=True)
    return out, dup_count, [p.name for p in paths]


def read_tipo_despacho() -> tuple[pd.DataFrame, int]:
    """Lee tipo_despacho.xlsx -> ([cedi(key), tipo], n_duplicados).

    El CEDI se normaliza con normalize() para emparejar con el cliente de salidas
    (Nombre completo con tratamiento y título). Si hubiera CEDI duplicados, conserva
    el primero y reporta el conteo.
    """
    path = config.FILES["tipo_despacho"]
    df = pd.read_excel(path, engine=excel_engine())
    cols = list(df.columns)
    out = pd.DataFrame()
    out["cedi"] = df[find_column(cols, config.TIPO_DESPACHO_COLS["cedi"])]
    out["tipo"] = df[find_column(cols, config.TIPO_DESPACHO_COLS["tipo"])]
    out["cedi"] = out["cedi"].map(lambda v: None if pd.isna(v) else normalize(v))
    out["tipo"] = out["tipo"].astype(str).str.strip().str.upper()
    out = out[out["cedi"].notna()]
    dup_count = int(out["cedi"].duplicated(keep=False).sum())
    return out.drop_duplicates(subset=["cedi"], keep="first").reset_index(drop=True), dup_count


# ---------------------------------------------------------------------------
# Lectura de destrucción (destruccion_*): Paso 2
# ---------------------------------------------------------------------------

def find_destruccion_files(base_dir: Optional[Path] = None) -> list[Path]:
    """Localiza los archivos destruccion* en CONSUMER/ y PROFESIONAL/ (excluye ~$)."""
    base = Path(base_dir) if base_dir else config.BASE_DIR
    results: list[Path] = []
    for key in ("consumer", "profesional"):
        folder = config.DIRS[key]
        if not folder.exists():
            continue
        results.extend(_exclude_locks(sorted(folder.glob(config.DESTRUCCION_GLOB))))
    return results


def read_destruccion(base_dir: Optional[Path] = None):
    """Lee todos los destruccion* -> (df[negocio, valor], archivos_leidos, notas).

    Para cada archivo: `negocio` = 'Almacen' en MAYÚSCULAS (no nulo); `valor` = conteo de
    filas por negocio (suma entre archivos). Los archivos **vacíos** (0 filas) no son error:
    se anotan y se continúa. Si un archivo no se puede leer o no tiene la columna 'Almacen'
    se lanza ValueError (la pipeline lo convierte en error grave que detiene el proceso).
    Lanza FileNotFoundError si no hay ningún archivo destruccion*.
    """
    paths = find_destruccion_files(base_dir)
    if not paths:
        raise FileNotFoundError("No se encontraron archivos destruccion* en CONSUMER/ y PROFESIONAL/.")

    frames = []
    files_read = []
    notes = []
    for path in paths:
        files_read.append(path.name)
        try:
            df = pd.read_excel(path, engine=excel_engine())
        except Exception as exc:  # corrupto / formato raro
            raise ValueError(f"{path.name}: no se pudo leer el archivo ({exc}).") from exc
        if df is None or df.empty:
            notes.append(f"{path.name}: archivo vacío (0 filas), se omite.")
            continue
        try:
            col = find_column(list(df.columns), config.DESTRUCCION_COLS["almacen"])
        except KeyError as exc:
            raise ValueError(f"{path.name}: falta la columna 'Almacen' ({exc}).") from exc
        almacen = df[col]
        almacen = almacen[almacen.notna()].astype(str).str.strip().str.upper()
        almacen = almacen[almacen != ""]
        if not almacen.empty:
            frames.append(almacen.to_frame(name="negocio"))

    if frames:
        all_neg = pd.concat(frames, ignore_index=True)
        grouped = all_neg.groupby("negocio").size().reset_index(name="valor")
        grouped["valor"] = grouped["valor"].astype(int)
    else:
        grouped = pd.DataFrame(columns=["negocio", "valor"])
    return grouped, files_read, notes
