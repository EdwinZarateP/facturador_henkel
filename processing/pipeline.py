"""Núcleo del Facturador Henkel (Paso 1 SALIDAS + Paso 2 DESTRUCCIÓN + Paso 3 INGRESOS +
Paso 4 OCUPACIÓN + Paso 5 TRASLADOS + Paso 6 MAQUILA + Paso 7 EXPORTACIONES +
Paso 8 ETIQUETAS + Paso 9 PALETIZADO + Paso 10 TRINCAJE + Paso 11 PLANTA + Paso 12 MATERIAL).

Lógica pura (sin FastAPI): lee las salidas, filtra por rango de fechas, clasifica
cada material en una familia (Negocio) y calcula cajas y estibas; luego procesa
los archivos `destruccion*` (Paso 2), `ingresos*` (Paso 3), `ocupacion*` (Paso 4),
`traslados*` (Paso 5), `maquila*` (Paso 6), `exportacion*` (Paso 7), `etiquetas*`
(Paso 8), `paletizado*` (Paso 9), `trincaje*` (Paso 10), `planta*` (Paso 11) y
`ocupacionMaterial*` (Paso 12) y combina todo en un único listado de servicios.
Devuelve un Step1Result listo para serializar a JSON o exportar a Excel.

Robustez frente a fuentes con errores:
- **Errores graves de archivo** (corrupto, mal formato, falta archivo obligatorio)
  lanzan `BlockingError` y DETIENEN el proceso para que el usuario corrija.
- El resto (duplicados, lookups opcionales ausentes, archivos vacíos) se reporta
  como `issues` (warning) y el proceso continúa.
- Hay un estado de ejecución global (`RunState`) que la API sondea para mostrar
  el avance (% y etapa) y los problemas en vivo.
- El resultado se cachea por rango (start|end) para que "Exportar" sea instantáneo
  tras "Generar" (app local, un solo usuario).
"""
from __future__ import annotations

import math
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

import config
from processing import io_utils


# Caché en memoria: (start|end) -> Step1Result. App local de un usuario.
_RESULT_CACHE: dict[str, "Step1Result"] = {}


# ---------------------------------------------------------------------------
# Errores graves y estado de ejecución (avance + bloqueo)
# ---------------------------------------------------------------------------

class BlockingError(Exception):
    """Error grave de archivo que debe DETENER el proceso (el usuario debe corregir)."""


@dataclass
class RunState:
    """Estado de la ejecución en curso, sondeado por GET /api/progress."""

    run_key: str = ""        # "start|end"
    stage: str = ""          # etiqueta de la etapa actual
    percent: int = 0         # 0..100
    done: bool = False
    blocked: bool = False    # True si se detuvo por un error grave
    error: str = ""          # mensaje del error grave (si blocked)
    issues: list = field(default_factory=list)
    has_result: bool = False  # True si hay un resultado exportable
    started_at: float = 0.0   # epoch del inicio (para el cronómetro)
    elapsed_seconds: float = 0.0  # tiempo total al terminar


_RUN_STATE = RunState()
_RUN_LOCK = threading.Lock()


def start_run(start: str, end: str) -> None:
    """Reinicia el estado de ejecución para un nuevo rango."""
    with _RUN_LOCK:
        _RUN_STATE.run_key = f"{start.strip()}|{end.strip()}"
        _RUN_STATE.stage = "Iniciando…"
        _RUN_STATE.percent = 0
        _RUN_STATE.done = False
        _RUN_STATE.blocked = False
        _RUN_STATE.error = ""
        _RUN_STATE.issues = []
        _RUN_STATE.has_result = False
        _RUN_STATE.started_at = time.time()
        _RUN_STATE.elapsed_seconds = 0.0


def set_progress(stage: str, percent: int) -> None:
    with _RUN_LOCK:
        _RUN_STATE.stage = stage
        _RUN_STATE.percent = int(percent)


def add_issue(issue: dict) -> None:
    with _RUN_LOCK:
        _RUN_STATE.issues.append(issue)


def finish_run(*, blocked: bool = False, error: str = "", has_result: bool = False) -> None:
    with _RUN_LOCK:
        if _RUN_STATE.started_at:
            _RUN_STATE.elapsed_seconds = time.time() - _RUN_STATE.started_at
        _RUN_STATE.done = True
        _RUN_STATE.blocked = blocked
        _RUN_STATE.error = error
        _RUN_STATE.has_result = has_result
        if not blocked:
            _RUN_STATE.percent = 100
            _RUN_STATE.stage = "Listo"


def get_progress() -> dict:
    """Snapshot del estado actual para la API (sin filas de detalle)."""
    with _RUN_LOCK:
        if _RUN_STATE.done:
            elapsed = _RUN_STATE.elapsed_seconds
        elif _RUN_STATE.started_at:
            elapsed = time.time() - _RUN_STATE.started_at
        else:
            elapsed = 0.0
        return {
            "run_key": _RUN_STATE.run_key,
            "stage": _RUN_STATE.stage,
            "percent": _RUN_STATE.percent,
            "done": _RUN_STATE.done,
            "blocked": _RUN_STATE.blocked,
            "error": _RUN_STATE.error,
            "issues": list(_RUN_STATE.issues),
            "has_result": _RUN_STATE.has_result,
            "elapsed_seconds": round(elapsed, 1),
        }


def is_running() -> bool:
    with _RUN_LOCK:
        return _RUN_STATE.run_key != "" and not _RUN_STATE.done


def invalidate_cache(start: str, end: str) -> None:
    """Elimina el resultado cacheado de un rango (p. ej. tras un run bloqueado)."""
    _RESULT_CACHE.pop(f"{start.strip()}|{end.strip()}", None)


@dataclass
class Step1Result:
    """Resultado del proceso (SALIDAS + DESTRUCCIÓN)."""

    totals: dict                   # servicios (SALIDAS + DESTRUCCION) + totales
    diagnostics: dict              # matched/unmatched, archivos leídos, advertencias
    issues: list[dict]             # [{severity, msg, ...}] errores/advertencias detectadas
    generated_at: str              # timestamp ISO
    excel_path: Optional[str] = None  # ruta del Excel pre-construido (None si aún no se ha generado)

    def to_summary_dict(self) -> dict:
        """Versión que devuelve la API. No incluye filas de detalle."""
        return {
            "totals": self.totals,
            "diagnostics": self.diagnostics,
            "issues": self.issues,
            "generated_at": self.generated_at,
        }


# ---------------------------------------------------------------------------
# Fechas
# ---------------------------------------------------------------------------

def parse_date_range(start: str, end: str) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Convierte dd/mm/yyyy -> (Timestamp, Timestamp) inclusivo. Valida start <= end."""
    try:
        d_start = datetime.strptime(start.strip(), config.USER_DATE_FORMAT).date()
        d_end = datetime.strptime(end.strip(), config.USER_DATE_FORMAT).date()
    except ValueError as exc:
        raise ValueError(
            f"Formato de fecha inválido. Use {config.USER_DATE_FORMAT} (dd/mm/yyyy). "
            f"Recibido: start='{start}', end='{end}'."
        ) from exc
    if d_start > d_end:
        raise ValueError(
            f"La fecha inicial ({start}) no puede ser mayor que la final ({end})."
        )
    return pd.Timestamp(d_start), pd.Timestamp(d_end)


# ---------------------------------------------------------------------------
# Cálculo de cajas y estibas
# ---------------------------------------------------------------------------

def _ceil_div(cantidad, divisor) -> Optional[int]:
    """ceil(cantidad / divisor); None si divisor es nulo o 0. (uso escalar)"""
    if pd.isna(cantidad) or pd.isna(divisor) or divisor == 0:
        return None
    return int(math.ceil(cantidad / divisor))


def _ceil_series(cantidad, divisor) -> pd.Series:
    """ceil(cantidad / divisor) vectorizado; None donde cantidad o divisor falten o sean 0."""
    import numpy as np

    c = pd.to_numeric(cantidad, errors="coerce")
    d = pd.to_numeric(divisor, errors="coerce")
    valid = c.notna() & d.notna() & (d != 0)
    result = pd.Series([pd.NA] * len(c), index=c.index, dtype=object)
    if valid.any():
        result.loc[valid] = np.ceil(c.loc[valid] / d.loc[valid]).astype("int64")
    return result


def _floor_series(cantidad, divisor) -> pd.Series:
    """floor(cantidad / divisor) vectorizado; None donde cantidad o divisor falten o sean 0.

    Es la división ENTERA (pallets LLENOS) que usa INGRESOS: a diferencia de las estibas
    (ceil), aquí interesa la cantidad de pallets completos — el residuo va aparte
    (`_mod_series`) y se usa para las medidas de MATERIAL DE EMPAQUE.
    """
    import numpy as np

    c = pd.to_numeric(cantidad, errors="coerce")
    d = pd.to_numeric(divisor, errors="coerce")
    valid = c.notna() & d.notna() & (d != 0)
    result = pd.Series([pd.NA] * len(c), index=c.index, dtype=object)
    if valid.any():
        result.loc[valid] = np.floor(c.loc[valid] / d.loc[valid]).astype("int64")
    return result


def _mod_series(cantidad, divisor) -> pd.Series:
    """cantidad mod divisor vectorizado; None donde cantidad o divisor falten o sean 0."""
    import numpy as np

    c = pd.to_numeric(cantidad, errors="coerce")
    d = pd.to_numeric(divisor, errors="coerce")
    valid = c.notna() & d.notna() & (d != 0)
    result = pd.Series([pd.NA] * len(c), index=c.index, dtype=object)
    if valid.any():
        result.loc[valid] = (c.loc[valid] % d.loc[valid]).astype("int64")
    return result


def _fn_numero_series(values) -> pd.Series:
    """Réplica de `fnNumero` de logica.txt para la columna de ocupación.

    - null / no parseable -> NaN (null).
    - "-" o "" (tras trim) -> 0.
    - resto -> número.

    Vectorizado: parsea todo a numérico (coerce -> NaN) y luego sobreescribe con 0 las
    posiciones cuyo texto es "-" o vacío.
    """
    s = pd.Series(values)
    num = pd.to_numeric(s, errors="coerce")
    txt = s.astype(str).str.strip()
    is_zero = txt.isin(["-", ""])
    return num.mask(is_zero, 0)


# ---------------------------------------------------------------------------
# Carga de datos
# ---------------------------------------------------------------------------

def _load_all_salidas(strict: bool = True):
    """Concatena todos los salidas_cons* y salidas_prof*.

    Devuelve (df, archivos_leidos_ok, issues_por_archivo).
    - strict=True (por defecto, para el proceso): ante cualquier archivo que no se
      pueda leer, o si no hay archivos, lanza BlockingError (detiene el proceso).
    - strict=False (calendario por defecto): tolerante, reporta y continúa.
    """
    files = io_utils.find_salidas_files()
    if not files:
        msg = "No se encontraron archivos salidas_cons*/salidas_prof* en CONSUMER/ y PROFESIONAL/."
        if strict:
            raise BlockingError(msg)
        return pd.DataFrame(), [], [{"severity": "error", "msg": msg}]
    frames = []
    ok_names = []
    issues = []
    for path, area in files:
        try:
            frames.append(io_utils.read_salidas(path, area))
            ok_names.append(path.name)
        except Exception as exc:  # archivo corrupto, sin columnas, hoja rara, etc.
            if strict:
                raise BlockingError(f"{path.name}: no se pudo leer el archivo ({exc}).") from exc
            issues.append(
                {
                    "severity": "error",
                    "file": path.name,
                    "msg": f"No se pudo leer el archivo: {exc}",
                }
            )
    df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return df, ok_names, issues


# ---------------------------------------------------------------------------
# Pipeline SALIDAS (Paso 1) — cuerpo puro, sin cachear ni finalizar
# ---------------------------------------------------------------------------

def _run_salidas_pipeline(
    start: str,
    end: str,
    ts_start: pd.Timestamp,
    ts_end: pd.Timestamp,
    emit,                       # callable(issue) para reportar issues en vivo
    progress=None,              # callable(stage, percent)
) -> tuple[dict, dict]:
    """Ejecuta el Paso 1 (SALIDAS) y devuelve (totals, diagnostics).

    No cachea ni construye Step1Result: eso lo hace run_all(). Los errores graves
    lanzan BlockingError. 0 filas en rango → warning (no detiene; run_all sigue
    con la destrucción). `emit` reporta cada issue (la API los muestra en vivo).
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    # 1) Cargar y concatenar salidas (estricto: error grave si algo falla).
    p("Leyendo archivos de salidas…", 5)
    salidas, files_read, load_issues = _load_all_salidas(strict=True)
    for iss in load_issues:
        emit(iss)
    rows_total = int(len(salidas))

    diagnostics: dict = {
        "files_read": files_read,
        "rows_total": rows_total,
        "rows_in_range": 0,
        "range_start": start,
        "range_end": end,
    }

    if salidas.empty:
        emit(
            {
                "severity": "warning",
                "msg": "No hay filas en los archivos de salidas (los archivos están vacíos).",
            }
        )
        return {}, diagnostics

    # Normalizar llaves y cantidad.
    salidas["material"] = salidas["material"].astype(str).str.strip().str.upper()
    salidas["cantidad"] = pd.to_numeric(salidas["cantidad"], errors="coerce")

    # 2) Filtrar por Fecha factura en el rango (inclusivo, ignora la hora).
    p("Filtrando por fecha…", 20)
    fecha = pd.to_datetime(salidas["fecha"], errors="coerce").dt.normalize()
    in_range = (fecha >= ts_start) & (fecha <= ts_end)
    salidas = salidas.loc[in_range].copy()
    salidas["fecha"] = pd.to_datetime(salidas["fecha"], errors="coerce")
    rows_in_range = int(len(salidas))
    diagnostics["rows_in_range"] = rows_in_range

    if rows_in_range == 0:
        emit(
            {
                "severity": "warning",
                "msg": f"0 filas con 'Fecha factura' dentro del rango {start} – {end} (solo destrucción, si la hay).",
            }
        )
        return {}, diagnostics

    # 3) material -> negocio via idh (default por área).
    p("Cruzando catálogos (huellas, idh, adicionales, despacho)…", 35)
    idh_matched, idh_defaulted, idh_dup = 0, 0, 0
    try:
        idh, idh_dup = io_utils.read_idh()
        before = salidas.shape[0]
        salidas = salidas.merge(idh, on="material", how="left")
        idh_matched = int(salidas["negocio"].notna().sum())
        idh_defaulted = before - idh_matched
    except FileNotFoundError:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontró idh_especiales.xlsx. Todo queda con Negocio por defecto.",
            }
        )
        salidas["negocio"] = pd.NA
    except Exception as exc:
        emit({"severity": "warning", "msg": f"Error leyendo idh_especiales: {exc}"})
        salidas["negocio"] = pd.NA

    # Rellenar negocio por defecto según el área de origen (vectorizado).
    default_map = salidas["area"].map(config.AREA_DEFAULT).fillna("CONSUMER")
    salidas["negocio"] = salidas["negocio"].fillna(default_map)

    # 4) material -> pallet/caja via huellas.
    huellas_matched, huellas_missing, huellas_dup = 0, 0, 0
    try:
        huellas, huellas_dup = io_utils.read_huellas()
        salidas = salidas.merge(
            huellas.rename(columns={"producto": "material"}), on="material", how="left"
        )
        huellas_matched = int(salidas["pallet"].notna().sum())
        huellas_missing = rows_in_range - huellas_matched
    except FileNotFoundError:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontró huellas.xlsx. Estibas y cajas quedarán vacías.",
            }
        )
        salidas["pallet"] = pd.NA
        salidas["caja"] = pd.NA
    except Exception as exc:
        emit({"severity": "warning", "msg": f"Error leyendo huellas: {exc}"})
        salidas["pallet"] = pd.NA
        salidas["caja"] = pd.NA

    # 5) Calcular estibas y cajas (vectorizado con numpy).
    salidas["estibas"] = _ceil_series(salidas["cantidad"], salidas["pallet"])
    salidas["cajas"] = _ceil_series(salidas["cantidad"], salidas["caja"])

    null_estibas = int(salidas["estibas"].isna().sum())
    null_cajas = int(salidas["cajas"].isna().sum())

    # 6) tipo_trabajo: Delivery -> ENTREGA/TIPO en adicionales (default NORMAL).
    tt_matched, tt_dup, tt_files = 0, 0, []
    salidas["tipo_trabajo"] = config.TIPO_TRABAJO_DEFAULT
    try:
        adic, tt_dup, tt_files = io_utils.read_adicionales()
        if not adic.empty:
            tt_map = dict(zip(adic["entrega"], adic["tipo"]))
            dkey = salidas["delivery"].map(io_utils.entrega_key)
            matched = dkey.notna() & dkey.isin(set(tt_map))
            salidas.loc[matched, "tipo_trabajo"] = dkey.loc[matched].map(tt_map)
            tt_matched = int(matched.sum())
    except FileNotFoundError:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos adicionales*. tipo_trabajo queda como NORMAL.",
            }
        )
    except Exception as exc:
        emit({"severity": "warning", "msg": f"Error leyendo adicionales: {exc}"})

    # 7) tipo_despacho: cliente -> CEDI/TIPO en tipo_despacho.xlsx (default ESTANDAR).
    td_matched, td_dup = 0, 0
    salidas["tipo_despacho"] = config.TIPO_DESPACHO_DEFAULT
    try:
        td, td_dup = io_utils.read_tipo_despacho()
        if not td.empty:
            td_map = dict(zip(td["cedi"], td["tipo"]))
            ckey = salidas["cliente"].map(lambda v: None if pd.isna(v) else io_utils.normalize(v))
            matched = ckey.notna() & ckey.isin(set(td_map))
            salidas.loc[matched, "tipo_despacho"] = ckey.loc[matched].map(td_map)
            td_matched = int(matched.sum())
    except FileNotFoundError:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontró tipo_despacho.xlsx. tipo_despacho queda como ESTANDAR.",
            }
        )
    except Exception as exc:
        emit({"severity": "warning", "msg": f"Error leyendo tipo_despacho: {exc}"})

    diagnostics.update(
        {
            "idh_matched": idh_matched,
            "idh_defaulted": idh_defaulted,
            "idh_duplicates": idh_dup,
            "huellas_matched": huellas_matched,
            "huellas_missing": huellas_missing,
            "huellas_duplicates_resolved": huellas_dup,
            "null_estibas": null_estibas,
            "null_cajas": null_cajas,
            "tipo_trabajo_matched": tt_matched,
            "adicionales_files": tt_files,
            "adicionales_duplicates": tt_dup,
            "tipo_despacho_matched": td_matched,
            "tipo_despacho_duplicates": td_dup,
        }
    )

    # 8) Issues derivados de la calidad de los datos (no detienen).
    if huellas_missing > 0:
        emit(
            {
                "severity": "warning",
                "msg": (
                    f"{huellas_missing:,} registros sin huella (material no encontrado en huellas.xlsx): "
                    "sus estibas/cajas quedaron vacías."
                ).replace(",", "."),
            }
        )
    if idh_dup > 0:
        emit(
            {
                "severity": "warning",
                "msg": f"idh_especiales tenía {idh_dup} materiales duplicados; se conservó el primero.",
            }
        )
    if tt_dup > 0:
        emit(
            {
                "severity": "warning",
                "msg": f"adicionales tenía {tt_dup} entregas duplicadas; se conservó la primera.",
            }
        )
    if td_dup > 0:
        emit(
            {
                "severity": "warning",
                "msg": f"tipo_despacho tenía {td_dup} CEDI duplicados; se conservó el primero.",
            }
        )

    # 9) Totales SALIDAS: agrupar + AddServicio + FilterServicio.
    p("Clasificando servicios de salidas…", 55)
    totals = _build_totals(salidas)
    return totals, diagnostics


# ---------------------------------------------------------------------------
# Pipeline DESTRUCCIÓN (Paso 2)
# ---------------------------------------------------------------------------

_DESTRUCCION_SERVICIO = "ALISTAMIENTO Y DESPACHO PALLETS DESTRUCCION"


def run_step2_destruccion(emit, progress=None) -> list[dict]:
    """Procesa los archivos destruccion* -> lista de líneas de servicio.

    Lee io_utils.read_destruccion(): negocio = 'Almacen' (MAYÚSCULAS), valor = conteo
    de filas por negocio. Sin archivos -> warning (no detiene). Error de lectura o
    falta columna 'Almacen' -> BlockingError (detiene). Archivos vacíos -> se omiten.
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    p("Procesando destrucción…", 65)
    try:
        grouped, files_read, notes = io_utils.read_destruccion()
    except FileNotFoundError:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos destruccion* (Paso 2 omitido).",
            }
        )
        return []
    except ValueError as exc:
        raise BlockingError(str(exc)) from exc
    except Exception as exc:
        raise BlockingError(f"Error leyendo destrucción: {exc}") from exc

    for note in notes:
        emit({"severity": "warning", "msg": note})

    servicios = []
    for _, r in grouped.iterrows():
        neg = str(r["negocio"]).strip().upper()
        valor = int(r["valor"])
        servicios.append(
            {
                "negocio": neg,
                "negocio_facturador": neg,
                "servicio": _DESTRUCCION_SERVICIO,
                "valor": valor,
                "unidades": 0,
                "proceso_extendido": f"PICKING {neg}",
                "macro_proceso": "OUT BOUND",
                "proceso_abreviado": "OUB",
                "tabla": "DESTRUCCION",
            }
        )
    servicios.sort(key=lambda s: s["negocio"])
    return servicios


# ---------------------------------------------------------------------------
# Pipeline INGRESOS (Paso 3) — recibo de mercancía
# ---------------------------------------------------------------------------

def _ing_servicio(negocio, negocio_facturador, servicio, valor, unidades) -> dict:
    """Construye una línea de servicio de INGRESOS con sus constantes (logica.txt)."""
    return {
        "negocio": negocio,
        "negocio_facturador": negocio_facturador,
        "servicio": servicio,
        "valor": valor,
        "unidades": unidades,
        "proceso_extendido": "RECIBO",
        "macro_proceso": "IN BOUND",
        "proceso_abreviado": "INB",
        "tabla": "INGRESOS",
    }


def _run_ingresos_pipeline(
    start: str,
    end: str,
    ts_start: pd.Timestamp,
    ts_end: pd.Timestamp,
    emit,
    progress=None,
) -> list[dict]:
    """Paso 3 (INGRESOS): procesa ingresos_cons*/ingresos_prof* -> lista de servicios.

    El rango de fechas lo define el usuario: se filtra la columna `Posting Date` con
    [start, end] (NO se usa RANGOS_FECHAS ni la fecha del nombre). El área/negocio viene
    del prefijo del archivo (ingresos_cons* -> CONSUMER, ingresos_prof* -> PROFESIONAL).

    1. Lee y concatena los ingresos* (estricto: error grave si un archivo obligatorio falla).
    2. documento_cruce = referencia si no nula, si no documento; descarta filas sin ninguno.
    3. Filtra por Posting Date en el rango.
    4. Cruza huellas (pallet/caja) e idh (negocio, incl. MATERIAL DE EMPAQUE).
    5. Calcula las 6 medidas de recibo (vectorizado).
    6. Agrupa por (negocio, negocio_facturador) y emite los servicios según ME:
       - no-ME -> RECIBO CAJAS (valor=recibo_cajas, unidades=cantidad).
       - ME    -> RECIBO CAJAS ME (valor=cajas_material, unidades=unidades_residuo) y
                  RECIBO PALLETS ME (valor=recibo_pallets, unidades=unidades_pallet_me).
    Sin archivos -> warning y []. Error de lectura -> BlockingError.
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    p("Leyendo archivos de ingresos…", 72)
    files = io_utils.find_ingresos_files()
    if not files:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos ingresos* (Paso 3 omitido).",
            }
        )
        return []

    frames, ok_names = [], []
    for path, area in files:
        try:
            frames.append(io_utils.read_ingresos(path, area))
            ok_names.append(path.name)
        except Exception as exc:  # corrupto, sin columnas, hoja rara
            raise BlockingError(
                f"{path.name}: no se pudo leer el archivo de ingresos ({exc})."
            ) from exc

    ing = pd.concat(frames, ignore_index=True)
    if ing.empty:
        emit(
            {
                "severity": "warning",
                "msg": "Los archivos de ingresos están vacíos (Paso 3 omitido).",
            }
        )
        return []

    # 2) documento_cruce = referencia si no nula, si no documento. Descarta filas sin ninguno.
    documento_cruce = ing["referencia"].where(ing["referencia"].notna(), ing["documento"])
    ing = ing.loc[documento_cruce.notna()].copy()

    # 3) Filtro por Posting Date en el rango del usuario (inclusivo, ignora la hora).
    p("Filtrando ingresos por fecha…", 76)
    posting = pd.to_datetime(ing["posting_date"], errors="coerce").dt.normalize()
    ing = ing.loc[(posting >= ts_start) & (posting <= ts_end)].copy()
    if ing.empty:
        emit(
            {
                "severity": "warning",
                "msg": f"0 filas de ingresos con 'Posting Date' en el rango {start} – {end}.",
            }
        )
        return []

    ing["cantidad"] = pd.to_numeric(ing["cantidad"], errors="coerce")

    # 4) Cruces huellas (pallet/caja) e idh (negocio, incluye MATERIAL DE EMPAQUE).
    p("Cruzando catálogos de ingresos (huellas, idh)…", 80)
    try:
        huellas, _ = io_utils.read_huellas()
        ing = ing.merge(huellas.rename(columns={"producto": "material"}), on="material", how="left")
    except FileNotFoundError:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontró huellas.xlsx. Los servicios de ingresos quedarán vacíos.",
            }
        )
        ing["pallet"] = pd.NA
        ing["caja"] = pd.NA
    except Exception as exc:
        emit({"severity": "warning", "msg": f"Error leyendo huellas para ingresos: {exc}"})
        ing["pallet"] = pd.NA
        ing["caja"] = pd.NA

    try:
        idh, _ = io_utils.read_idh()
        ing = ing.merge(idh, on="material", how="left")
    except FileNotFoundError:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontró idh_especiales.xlsx. negocio_facturador de ingresos queda por área.",
            }
        )
        ing["negocio"] = pd.NA
    except Exception as exc:
        emit({"severity": "warning", "msg": f"Error leyendo idh para ingresos: {exc}"})
        ing["negocio"] = pd.NA

    # negocio_facturador = idh si match, si no área. negocio = MATERIAL DE EMPAQUE si nf es ME.
    area_label = ing["area"].map(config.AREA_DEFAULT).fillna("CONSUMER")
    nf = ing["negocio"].fillna(area_label)
    ing["negocio_facturador"] = nf
    me_mask = nf.eq("MATERIAL DE EMPAQUE")
    ing["negocio"] = area_label.mask(me_mask, "MATERIAL DE EMPAQUE")

    # 5) Medidas de recibo (vectorizado). recibo_pallets es FLOOR (pallets llenos), no ceil.
    cantidad = ing["cantidad"]
    pallet = pd.to_numeric(ing["pallet"], errors="coerce")
    caja = pd.to_numeric(ing["caja"], errors="coerce")

    recibo_cajas = pd.to_numeric(_ceil_series(cantidad, caja), errors="coerce")
    recibo_pallets = pd.to_numeric(_floor_series(cantidad, pallet), errors="coerce")
    unidades_residuo = pd.to_numeric(_mod_series(cantidad, pallet), errors="coerce")
    # unidades_pallet_me = pallet * recibo_pallets (null si pallet o recibo_pallets null).
    up_me = pd.Series([pd.NA] * len(ing), index=ing.index, dtype=object)
    up_valid = pallet.notna() & recibo_pallets.notna()
    up_me.loc[up_valid] = (pallet.loc[up_valid] * recibo_pallets.loc[up_valid]).astype("int64")
    cajas_material = pd.to_numeric(_ceil_series(unidades_residuo, caja), errors="coerce")

    ing["_recibo_cajas"] = recibo_cajas
    ing["_recibo_pallets"] = recibo_pallets
    ing["_cajas_material"] = cajas_material
    ing["_unidades_residuo"] = unidades_residuo
    ing["_unidades_pallet_me"] = pd.to_numeric(up_me, errors="coerce")

    # 6) Agrupar por (negocio, negocio_facturador) y emitir servicios.
    grouped = (
        ing.groupby(["negocio", "negocio_facturador"], dropna=False)
        .agg(
            recibo_cajas=("_recibo_cajas", "sum"),
            recibo_pallets=("_recibo_pallets", "sum"),
            cajas_material=("_cajas_material", "sum"),
            unidades_residuo=("_unidades_residuo", "sum"),
            unidades_pallet_me=("_unidades_pallet_me", "sum"),
            recibo_unidades=("cantidad", "sum"),
        )
        .reset_index()
    )

    servicios = []
    for _, r in grouped.iterrows():
        nf = str(r["negocio_facturador"]) if pd.notna(r["negocio_facturador"]) else ""
        neg = str(r["negocio"]) if pd.notna(r["negocio"]) else ""
        if nf == "MATERIAL DE EMPAQUE":
            v_cajas = _safe_num(r["cajas_material"])
            if v_cajas:
                servicios.append(
                    _ing_servicio(neg, nf, "RECIBO CAJAS ME", v_cajas, _safe_num(r["unidades_residuo"]))
                )
            v_pme = _safe_num(r["recibo_pallets"])
            if v_pme:
                servicios.append(
                    _ing_servicio(neg, nf, "RECIBO PALLETS ME", v_pme, _safe_num(r["unidades_pallet_me"]))
                )
        else:
            v = _safe_num(r["recibo_cajas"])
            if v:
                servicios.append(_ing_servicio(neg, nf, "RECIBO CAJAS", v, _safe_num(r["recibo_unidades"])))

    servicios.sort(key=lambda s: (s["negocio"], s["negocio_facturador"], s["servicio"]))
    return servicios


# ---------------------------------------------------------------------------
# Pipeline OCUPACIÓN (Paso 4) — almacenamiento
# ---------------------------------------------------------------------------

# proceso_extendido fijo por servicio (logica.txt, AgregarFinal). default -> BIN.
_OCUP_PROCESO_EXTENDIDO = {
    "ALMACENAMIENTO PALLET BODEGA 7": "PALLETS",
    "ALMACENAMIENTO PALLET BODEGA 8": "PALLETS",
    "ALMACENAMIENTO PALLET BODEGA 8 ME": "PALLETS",
    "ALMACENAMIENTO MEDIO PALLET": "MEDIO PALLET",
    "ALMACENAMIENTO FLOW RACK": "FLOW RACK",
    "ALMACENAMIENTO BIN": "BIN",
    "ALMACENAMIENTO MODULA": "MODULA",
}
_OCUP_PROCESO_DEFAULT = "BIN"


def _run_ocupacion_pipeline(
    start: str,
    end: str,
    ts_start: pd.Timestamp,
    ts_end: pd.Timestamp,
    emit,
    progress=None,
) -> list[dict]:
    """Paso 4 (OCUPACIÓN): procesa ocupacion_cons*/ocupacion_prof* -> lista de servicios.

    Facturación de ALMACENAMIENTO. A diferencia de SALIDAS/INGRESOS, **no usa huellas ni
    idh**: usa la tabla EQUIVALENCIAS (`Tipo` -> `Conversion liquidacion`) y el valor es un
    **promedio diario redondeado hacia arriba** (ceil), con el caso especial
    `ALMACENAMIENTO MODULA` = ceil(ocup * 350.3). El rango de fechas lo define el usuario
    sobre la columna `Fecha` (no RANGOS_FECHAS ni la fecha del nombre).

    1. Lee y concatena ocupacion* (estricto: error grave si un archivo obligatorio falla).
    2. `ocupacion` -> numérico (fnNumero: '-' o '' -> 0); filtra Fecha no nula.
    3. Filtra por Fecha en el rango del usuario (inclusivo).
    4. `negocio` = área; `negocio_facturador` = LAUNDRY si `Almacen` contiene "LDRY", si no
       `negocio`.
    5. Join EQUIVALENCIAS por `Tipo` -> `conversion`; descarta filas sin conversión
       (Tipos no facturables / sin servicio convertible).
    6. Suma por día `(fecha, negocio, nf, conversion)`; descarta días con suma 0/nula
       (`FiltrarOcupacionValida`).
    7. Promedia por `(negocio, nf, conversion)` -> average de los días -> ceil.
    8. Servicios: `servicio` = "ALMACENAMIENTO "+conversion; `valor` = ocup, salvo MODULA =
       ceil(ocup * 350.3) (sobre el promedio ya redondeado); `nf` final = NATTURA si
       (PROFESIONAL + ALMACENAMIENTO BIN). unidades=0, macro_proceso=ALMACENAMIENTO,
       proceso_abreviado=WHS, tabla=OCUPACION. **Excepción:** si conversion = "PALLET
       BODEGA 8 ME" y `nf` es CONSUMER/PROFESIONAL, el servicio pasa a
       "ALMACENAMIENTO PALLET BODEGA 8" (se agrupa/suma con el BODEGA 8 existente);
       LAUNDRY conserva "… 8 ME".
    Sin archivos -> warning y []. Sin equivalencias -> warning y []. Error de lectura ->
    BlockingError.
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    p("Leyendo archivos de ocupación…", 84)
    files = io_utils.find_ocupacion_files()
    if not files:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos ocupacion* (Paso 4 omitido).",
            }
        )
        return []

    frames, ok_names = [], []
    for path, area in files:
        try:
            frames.append(io_utils.read_ocupacion(path, area))
            ok_names.append(path.name)
        except Exception as exc:  # corrupto, sin columnas, hoja rara
            raise BlockingError(
                f"{path.name}: no se pudo leer el archivo de ocupación ({exc})."
            ) from exc

    ocu = pd.concat(frames, ignore_index=True)
    if ocu.empty:
        emit(
            {
                "severity": "warning",
                "msg": "Los archivos de ocupación están vacíos (Paso 4 omitido).",
            }
        )
        return []

    # 2) ocupacion -> numérico (fnNumero); fecha de almacenamiento; filtra Fecha no nula.
    ocu["ocupacion"] = _fn_numero_series(ocu["ocupacion"])
    ocu["_fecha"] = pd.to_datetime(ocu["fecha"], errors="coerce").dt.normalize()
    ocu = ocu.loc[ocu["_fecha"].notna()].copy()

    # 3) Filtro por Fecha en el rango del usuario (inclusivo).
    p("Filtrando ocupación por fecha…", 87)
    ocu = ocu.loc[(ocu["_fecha"] >= ts_start) & (ocu["_fecha"] <= ts_end)].copy()
    if ocu.empty:
        emit(
            {
                "severity": "warning",
                "msg": f"0 filas de ocupación con 'Fecha' en el rango {start} – {end}.",
            }
        )
        return []

    # 4) negocio = área; nf = LAUNDRY si Almacen contiene "LDRY".
    area_label = ocu["area"].map(config.AREA_DEFAULT).fillna("CONSUMER")
    ocu["negocio"] = area_label
    ocu["negocio_facturador"] = area_label.mask(
        ocu["almacen"].astype(str).str.contains("LDRY", na=False), "LAUNDRY"
    )

    # 5) Join EQUIVALENCIAS por Tipo -> conversion; descarta filas sin conversión.
    p("Cruzando equivalencias de ocupación…", 90)
    try:
        equiv, equiv_dup = io_utils.read_equivalencias()
    except FileNotFoundError:
        emit(
            {
                "severity": "warning",
                "msg": (
                    "No se encontró equivalencias_almacenamiento.xlsx. Ocupación sin "
                    "servicios convertibles (Paso 4 omitido)."
                ),
            }
        )
        return []
    except Exception as exc:
        emit({"severity": "warning", "msg": f"Error leyendo equivalencias: {exc}"})
        return []
    if equiv_dup > 0:
        emit(
            {
                "severity": "warning",
                "msg": f"equivalencias_almacenamiento tenía {equiv_dup} Tipos duplicados; se conservó el primero.",
            }
        )

    ocu = ocu.merge(
        equiv.rename(columns={"archivo": "tipo"}), on="tipo", how="left"
    )
    ocu = ocu.loc[ocu["conversion"].notna()].copy()
    if ocu.empty:
        return []

    # 6) Suma por día (List.Sum(List.RemoveNulls)) -> descarta días con suma 0/nula.
    ocu["ocupacion"] = pd.to_numeric(ocu["ocupacion"], errors="coerce")
    daily = (
        ocu.groupby(["_fecha", "negocio", "negocio_facturador", "conversion"], dropna=False)
        .agg(ocupacion=("ocupacion", "sum"))
        .reset_index()
    )
    daily = daily.loc[daily["ocupacion"].notna() & (daily["ocupacion"] != 0)]
    if daily.empty:
        return []

    # 7) Promedia por (negocio, nf, conversion) -> average -> ceil.
    period = (
        daily.groupby(["negocio", "negocio_facturador", "conversion"], dropna=False)
        .agg(ocupacion=("ocupacion", "mean"))
        .reset_index()
    )
    period["ocupacion"] = period["ocupacion"].apply(
        lambda v: math.ceil(v) if pd.notna(v) else pd.NA
    )
    period = period.loc[period["ocupacion"].notna()]

    # 8) Servicios finales.
    servicios = []
    for _, r in period.iterrows():
        neg = str(r["negocio"])
        nf = str(r["negocio_facturador"])
        conversion = str(r["conversion"])
        servicio = f"ALMACENAMIENTO {conversion}"
        # PALLET BODEGA 8 ME de CONSUMER/PROFESIONAL se factura como PALLET BODEGA 8:
        # al compartir nombre (y proceso_extendido "PALLETS") se agrupa/suma con el
        # BODEGA 8 existente en _aggregate_servicios. LAUNDRY conserva el "… ME".
        if conversion == "PALLET BODEGA 8 ME" and nf in ("CONSUMER", "PROFESIONAL"):
            servicio = "ALMACENAMIENTO PALLET BODEGA 8"
        ocup = int(r["ocupacion"])
        # MODULA se valora sobre el promedio YA redondeado: ceil(ocup * 350.3).
        valor = math.ceil(ocup * 350.3) if servicio == "ALMACENAMIENTO MODULA" else ocup
        if nf == "PROFESIONAL" and servicio == "ALMACENAMIENTO BIN":
            nf_final = "NATTURA"
        else:
            nf_final = nf
        servicios.append(
            {
                "negocio": neg,
                "negocio_facturador": nf_final,
                "servicio": servicio,
                "valor": valor,
                "unidades": 0,
                "proceso_extendido": _OCUP_PROCESO_EXTENDIDO.get(servicio, _OCUP_PROCESO_DEFAULT),
                "macro_proceso": "ALMACENAMIENTO",
                "proceso_abreviado": "WHS",
                "tabla": "OCUPACION",
            }
        )

    servicios.sort(key=lambda s: (s["negocio"], s["negocio_facturador"], s["servicio"]))
    return servicios


# ---------------------------------------------------------------------------
# Pipeline TRASLADOS (Paso 5) — centro de traslados
# ---------------------------------------------------------------------------

_TRASLADO_SERVICIO_CAJAS = "ALISTAMIENTO Y DESPACHO CAJAS CENTRO DE TRASLADOS"
_TRASLADO_SERVICIO_UNIDADES = "ALISTAMIENTO Y DESPACHO UNIDADES CENTRO DE TRASLADOS"


def _run_traslados_pipeline(emit, progress=None) -> list[dict]:
    """Paso 5 (TRASLADOS): procesa traslados_cons*/traslados_prof* -> lista de servicios.

    Es la pipeline más simple: **sin huellas, sin idh, sin equivalencias, sin filtro de fecha**
    (decisión del usuario, molde DESTRUCCIÓN: se procesan TODOS los traslados* encontrados).
    El área/negocio viene del prefijo del archivo (traslados_cons* -> CONSUMER, ..._prof* ->
    PROFESIONAL). Traducción de logica.txt:

    1. Lee y concatena los traslados* (estricto: error grave si un archivo obligatorio falla).
    2. Filtra filas con `Delivery` <> null (`#"Filas filtradas"` de logica.txt).
    3. `negocio` = área; agrupa por `negocio` sumando `shu` y `con`.
    4. Emite 2 servicios por negocio (unpivot de logica.txt):
       - CAJAS CENTRO DE TRASLADOS:   valor = sum(SHU), unidades = 0.
       - UNIDADES CENTRO DE TRASLADOS: valor = sum(CON), unidades = sum(CON).
       Constantes: negocio_facturador = negocio, proceso_extendido = "PICKING "+negocio,
       macro_proceso = "OUT BOUND", proceso_abreviado = "OUB", tabla = "TRASLADOS".
       Convención del bot: se omiten los servicios con valor 0 (igual que el resto de pasos).
    Sin archivos -> warning y []. Error de lectura -> BlockingError.
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    p("Procesando traslados…", 91)
    files = io_utils.find_traslados_files()
    if not files:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos traslados* (Paso 5 omitido).",
            }
        )
        return []

    frames, ok_names = [], []
    for path, area in files:
        try:
            frames.append(io_utils.read_traslados(path, area))
            ok_names.append(path.name)
        except Exception as exc:  # corrupto, sin columnas, hoja rara
            raise BlockingError(
                f"{path.name}: no se pudo leer el archivo de traslados ({exc})."
            ) from exc

    tras = pd.concat(frames, ignore_index=True)
    if tras.empty:
        emit(
            {
                "severity": "warning",
                "msg": "Los archivos de traslados están vacíos (Paso 5 omitido).",
            }
        )
        return []

    # 2) Filas con Delivery <> null (logica.txt: #"Filas filtradas").
    tras = tras.loc[tras["delivery"].notna()].copy()
    tras["shu"] = pd.to_numeric(tras["shu"], errors="coerce").fillna(0)
    tras["con"] = pd.to_numeric(tras["con"], errors="coerce").fillna(0)

    if tras.empty:
        return []

    # 3) negocio = área; agrupar sumando shu y con.
    tras["negocio"] = tras["area"].map(config.AREA_DEFAULT).fillna("CONSUMER")
    grouped = (
        tras.groupby("negocio", dropna=False)
        .agg(shu=("shu", "sum"), con=("con", "sum"))
        .reset_index()
    )

    # 4) Servicios finales (omitiendo valor 0, convención del bot).
    servicios = []
    for _, r in grouped.iterrows():
        neg = str(r["negocio"])
        shu = int(r["shu"])
        con = int(r["con"])
        if shu:
            servicios.append(
                {
                    "negocio": neg,
                    "negocio_facturador": neg,
                    "servicio": _TRASLADO_SERVICIO_CAJAS,
                    "valor": shu,
                    "unidades": 0,
                    "proceso_extendido": f"PICKING {neg}",
                    "macro_proceso": "OUT BOUND",
                    "proceso_abreviado": "OUB",
                    "tabla": "TRASLADOS",
                }
            )
        if con:
            servicios.append(
                {
                    "negocio": neg,
                    "negocio_facturador": neg,
                    "servicio": _TRASLADO_SERVICIO_UNIDADES,
                    "valor": con,
                    "unidades": con,
                    "proceso_extendido": f"PICKING {neg}",
                    "macro_proceso": "OUT BOUND",
                    "proceso_abreviado": "OUB",
                    "tabla": "TRASLADOS",
                }
            )

    servicios.sort(key=lambda s: (s["negocio"], s["servicio"]))
    return servicios


# ---------------------------------------------------------------------------
# Pipeline MAQUILA (Paso 6) — subcontratación / maquila
# ---------------------------------------------------------------------------

def _mq_servicio(negocio, negocio_facturador, servicio, valor, unidades) -> dict:
    """Construye una línea de servicio de MAQUILA con sus constantes (logica.txt)."""
    return {
        "negocio": negocio,
        "negocio_facturador": negocio_facturador,
        "servicio": servicio,
        "valor": valor,
        "unidades": unidades,
        "proceso_extendido": "MAQUILA",
        "macro_proceso": "OTROS",
        "proceso_abreviado": "MAQ",
        "tabla": "MAQUILA",
    }


def _run_maquila_pipeline(
    start: str,
    end: str,
    ts_start: pd.Timestamp,
    ts_end: pd.Timestamp,
    emit,
    progress=None,
) -> list[dict]:
    """Paso 6 (MAQUILA): procesa maquila_cons*/maquila_prof* -> lista de servicios.

    Traducción de la query `FilesMaquila` de logica.txt. Es casi idéntica a INGRESOS
    (Paso 3): mismas 6 medidas de recibo (`floor`/`mod`/`ceil`), mismos cruces
    (huellas + idh, con MATERIAL DE EMPAQUE vía idh_especiales.Negocio). La diferencia
    es el set de servicios y las constantes finales.

    El rango de fechas lo define el usuario: se filtra la columna `Posting Date` con
    [start, end] (NO se usa RANGOS_FECHAS ni la fecha del nombre). El área/negocio viene
    del prefijo del archivo (maquila_cons* -> CONSUMER, maquila_prof* -> PROFESIONAL);
    ambos ficheros viven en la carpeta MAQUILA/.

    1. Lee y concatena los maquila* (estricto: error grave si un archivo obligatorio falla).
       Filtra Material no nulo (logica.txt: #"Filas filtradas2").
    2. Filtra por Posting Date en el rango del usuario (inclusivo).
    3. Cruza huellas (pallet/caja) e idh (negocio, incl. MATERIAL DE EMPAQUE).
       negocio_facturador = idh si match, si no área; negocio = MATERIAL DE EMPAQUE si
       nf es ME (mismo molde que ingresos).
    4. Calcula las 6 medidas (vectorizado, mismos helpers que ingresos).
    5. Agrupa por (negocio, negocio_facturador) y emite los servicios según ME:
       - no-ME -> ALISTAMIENTO DE MAQUILA CAJAS (valor=cajas_generales, unidades=cantidad).
       - ME    -> PICKING PALLETS MAQUILA ME (valor=pallets_maquila, unidades=unidades_pallet_maquila)
                  + PICKING CAJAS MAQUILA ME (valor=cajas_maquila, unidades=unidades_cajas_maquila).

    Mapeo de `unidades` (decisión del usuario, "coherente" — NO el literal del PQ): el
    `AddUnidadesServicio` de logica.txt compara contra nombres "RECIBO…" que no existen
    en MAQUILA, así que en el PQ original todo cae a `unidades_pallet_maquila` (bug
    heredado de plantilla). Aquí se asigna la unidad coherente con cada servicio.

    Sin archivos -> warning y []. Error de lectura -> BlockingError.
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    p("Leyendo archivos de maquila…", 94)
    files = io_utils.find_maquila_files()
    if not files:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos maquila* (Paso 6 omitido).",
            }
        )
        return []

    frames, ok_names = [], []
    for path, area in files:
        try:
            frames.append(io_utils.read_maquila(path, area))
            ok_names.append(path.name)
        except Exception as exc:  # corrupto, sin columnas, hoja rara
            raise BlockingError(
                f"{path.name}: no se pudo leer el archivo de maquila ({exc})."
            ) from exc

    mq = pd.concat(frames, ignore_index=True)
    if mq.empty:
        emit(
            {
                "severity": "warning",
                "msg": "Los archivos de maquila están vacíos (Paso 6 omitido).",
            }
        )
        return []

    # 1) Filas con Material no nulo (logica.txt: #"Filas filtradas2").
    mq = mq.loc[mq["material"].notna()].copy()

    # 2) Filtro por Posting Date en el rango del usuario (inclusivo, ignora la hora).
    posting = pd.to_datetime(mq["posting_date"], errors="coerce").dt.normalize()
    mq = mq.loc[(posting >= ts_start) & (posting <= ts_end)].copy()
    if mq.empty:
        emit(
            {
                "severity": "warning",
                "msg": f"0 filas de maquila con 'Posting Date' en el rango {start} – {end}.",
            }
        )
        return []

    mq["cantidad"] = pd.to_numeric(mq["cantidad"], errors="coerce")

    # 3) Cruces huellas (pallet/caja) e idh (negocio, incluye MATERIAL DE EMPAQUE).
    p("Cruzando catálogos de maquila (huellas, idh)…", 96)
    try:
        huellas, _ = io_utils.read_huellas()
        mq = mq.merge(huellas.rename(columns={"producto": "material"}), on="material", how="left")
    except FileNotFoundError:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontró huellas.xlsx. Los servicios de maquila quedarán vacíos.",
            }
        )
        mq["pallet"] = pd.NA
        mq["caja"] = pd.NA
    except Exception as exc:
        emit({"severity": "warning", "msg": f"Error leyendo huellas para maquila: {exc}"})
        mq["pallet"] = pd.NA
        mq["caja"] = pd.NA

    try:
        idh, _ = io_utils.read_idh()
        mq = mq.merge(idh, on="material", how="left")
    except FileNotFoundError:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontró idh_especiales.xlsx. negocio_facturador de maquila queda por área.",
            }
        )
        mq["negocio"] = pd.NA
    except Exception as exc:
        emit({"severity": "warning", "msg": f"Error leyendo idh para maquila: {exc}"})
        mq["negocio"] = pd.NA

    # negocio_facturador = idh si match, si no área. negocio = MATERIAL DE EMPAQUE si nf es ME.
    area_label = mq["area"].map(config.AREA_DEFAULT).fillna("CONSUMER")
    nf = mq["negocio"].fillna(area_label)
    mq["negocio_facturador"] = nf
    me_mask = nf.eq("MATERIAL DE EMPAQUE")
    mq["negocio"] = area_label.mask(me_mask, "MATERIAL DE EMPAQUE")

    # 4) Medidas de maquila (mismas que ingresos, vectorizado).
    cantidad = mq["cantidad"]
    pallet = pd.to_numeric(mq["pallet"], errors="coerce")
    caja = pd.to_numeric(mq["caja"], errors="coerce")

    pallets_maquila = pd.to_numeric(_floor_series(cantidad, pallet), errors="coerce")
    unidades_cajas_maquila = pd.to_numeric(_mod_series(cantidad, pallet), errors="coerce")
    cajas_maquila = pd.to_numeric(_ceil_series(unidades_cajas_maquila, caja), errors="coerce")
    cajas_generales = pd.to_numeric(_ceil_series(cantidad, caja), errors="coerce")
    # unidades_pallet_maquila = pallet * pallets_maquila (null si alguno null).
    up_me = pd.Series([pd.NA] * len(mq), index=mq.index, dtype=object)
    up_valid = pallet.notna() & pallets_maquila.notna()
    up_me.loc[up_valid] = (pallet.loc[up_valid] * pallets_maquila.loc[up_valid]).astype("int64")
    unidades_pallet_maquila = pd.to_numeric(up_me, errors="coerce")

    mq["_cajas_generales"] = cajas_generales
    mq["_pallets_maquila"] = pallets_maquila
    mq["_cajas_maquila"] = cajas_maquila
    mq["_unidades_cajas_maquila"] = unidades_cajas_maquila
    mq["_unidades_pallet_maquila"] = unidades_pallet_maquila
    mq["_unidades_generales"] = cantidad

    # 5) Agrupar por (negocio, negocio_facturador) y emitir servicios (mapeo coherente).
    grouped = (
        mq.groupby(["negocio", "negocio_facturador"], dropna=False)
        .agg(
            cajas_generales=("_cajas_generales", "sum"),
            pallets_maquila=("_pallets_maquila", "sum"),
            cajas_maquila=("_cajas_maquila", "sum"),
            unidades_cajas_maquila=("_unidades_cajas_maquila", "sum"),
            unidades_pallet_maquila=("_unidades_pallet_maquila", "sum"),
            unidades_generales=("_unidades_generales", "sum"),
        )
        .reset_index()
    )

    servicios = []
    for _, r in grouped.iterrows():
        nf = str(r["negocio_facturador"]) if pd.notna(r["negocio_facturador"]) else ""
        neg = str(r["negocio"]) if pd.notna(r["negocio"]) else ""
        if nf == "MATERIAL DE EMPAQUE":
            v_pme = _safe_num(r["pallets_maquila"])
            if v_pme:
                servicios.append(
                    _mq_servicio(
                        neg, nf, "PICKING PALLETS MAQUILA ME",
                        v_pme, _safe_num(r["unidades_pallet_maquila"]),
                    )
                )
            v_cme = _safe_num(r["cajas_maquila"])
            if v_cme:
                servicios.append(
                    _mq_servicio(
                        neg, nf, "PICKING CAJAS MAQUILA ME",
                        v_cme, _safe_num(r["unidades_cajas_maquila"]),
                    )
                )
        else:
            v = _safe_num(r["cajas_generales"])
            if v:
                servicios.append(
                    _mq_servicio(
                        neg, nf, "ALISTAMIENTO DE MAQUILA CAJAS",
                        v, _safe_num(r["unidades_generales"]),
                    )
                )

    servicios.sort(key=lambda s: (s["negocio"], s["negocio_facturador"], s["servicio"]))
    return servicios


# ---------------------------------------------------------------------------
# Pipeline EXPORTACIONES (Paso 7) — exportaciones / export
# ---------------------------------------------------------------------------

# Bases de los 3 servicios (logica.txt: Grouped + UnpivotServicios). El nombre final
# lleva el `canal` pegado: `servicio = base & " " & canal` (p. ej. "... EXPO IC").
_EXPO_BASE_PALLETS = "ALISTAMIENTO Y DESPACHO PALLETS EXPO"
_EXPO_BASE_CAJAS = "ALISTAMIENTO Y DESPACHO CAJAS EXPO"
_EXPO_BASE_UND = "ALISTAMIENTO Y DESPACHO UND EXPO"


def _expo_servicio(negocio: str, servicio: str, valor, unidades) -> dict:
    """Construye una línea de servicio de EXPORTACIONES con sus constantes (logica.txt).

    `negocio_facturador` = `negocio` (área): el `JoinIDHEspeciales` de logica.txt es
    vestigial (hace el NestedJoin pero NUNCA lo expande) y luego duplica
    `negocio -> negocio_facturador`, así que nf = negocio = área (sin idh).
    """
    return {
        "negocio": negocio,
        "negocio_facturador": negocio,
        "servicio": servicio,
        "valor": valor,
        "unidades": unidades,
        "proceso_extendido": f"PICKING {negocio}",
        "macro_proceso": "OUT BOUND",
        "proceso_abreviado": "OUB",
        "tabla": "EXPORTACIONES",
    }


def _run_exportacion_pipeline(
    start: str,
    end: str,
    ts_start: pd.Timestamp,
    ts_end: pd.Timestamp,
    emit,
    progress=None,
) -> list[dict]:
    """Paso 7 (EXPORTACIONES): procesa exportacion_cons*/exportacion_prof* -> servicios.

    Traducción de la query `FilesExportacion` de logica.txt. Lee los `exportacion*` de
    la carpeta EXPORTACIONES/ (ambos prefijos en la misma carpeta, como maquila). El
    área/negocio viene del prefijo del nombre (exportacion_cons* -> CONSUMER, ..._prof* ->
    PROFESIONAL). `paletizado*`/`trincaje*` no se procesan (logica.txt sólo admite
    "exportacion*").

    El rango de fechas lo define el usuario: se filtra la columna `Fecha factura` con
    [start, end] (NO se usa RANGOS_FECHAS ni la fecha del nombre — la ventana del PQ
    se reemplaza por el rango libre del calendario, igual que en los pasos previos).

    1. Lee y concatena los exportacion* (estricto: error grave si un archivo falla).
       Filtra `Delivery` <> null (logica.txt: #"Filas filtradas").
    2. Filtra por Fecha factura en el rango del usuario (inclusivo).
    3. `negocio` = área; `negocio_facturador` = `negocio` (el `JoinIDHEspeciales` del PQ
       es vestigial — nunca se expande — y luego duplica negocio -> nf; sin idh).
    4. Cruza huellas (pallet/caja); NO hay cruce idh.
    5. Medidas (vectorizado): `pallets = ceil(cant/pallet)`, `cajas = ceil(cant/caja)`,
       `unidades = caja * cajas` (redondea hacia ARRIBA a múltiplo de caja; **no** es
       `cantidad` — es la unidades "facturadas" del grupo).
    6. Agrupa por (negocio, canal) y emite 3 servicios (unpivot de logica.txt):
       - PALLETS EXPO <canal>: valor = sum(pallets),  unidades = sum(unidades).
       - CAJAS EXPO <canal>:   valor = sum(cajas),    unidades = sum(unidades).
       - UND EXPO <canal>:     valor = sum(unidades), unidades = sum(unidades).
       `servicio = base & " " & canal`. Convención del bot: se omiten valor 0/nulo.

    Sin archivos -> warning y []. Error de lectura -> BlockingError. Sin huellas ->
    medidas vacías (advertencia).
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    p("Leyendo archivos de exportación…", 96)
    files = io_utils.find_exportacion_files()
    if not files:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos exportacion* (Paso 7 omitido).",
            }
        )
        return []

    frames, ok_names = [], []
    for path, area in files:
        try:
            frames.append(io_utils.read_exportacion(path, area))
            ok_names.append(path.name)
        except Exception as exc:  # corrupto, sin columnas, hoja rara
            raise BlockingError(
                f"{path.name}: no se pudo leer el archivo de exportación ({exc})."
            ) from exc

    expo = pd.concat(frames, ignore_index=True)
    if expo.empty:
        emit(
            {
                "severity": "warning",
                "msg": "Los archivos de exportación están vacíos (Paso 7 omitido).",
            }
        )
        return []

    # 1) Filas con Delivery <> null (logica.txt: #"Filas filtradas").
    expo = expo.loc[expo["delivery"].notna()].copy()

    # 2) Filtro por Fecha factura en el rango del usuario (inclusivo, ignora la hora).
    p("Filtrando exportación por fecha…", 97)
    fecha = pd.to_datetime(expo["fecha"], errors="coerce").dt.normalize()
    expo = expo.loc[(fecha >= ts_start) & (fecha <= ts_end)].copy()
    if expo.empty:
        emit(
            {
                "severity": "warning",
                "msg": f"0 filas de exportación con 'Fecha factura' en el rango {start} – {end}.",
            }
        )
        return []

    expo["cantidad"] = pd.to_numeric(expo["cantidad"], errors="coerce")

    # 3) negocio = área; negocio_facturador = negocio (el join idh del PQ es vestigial).
    expo["negocio"] = expo["area"].map(config.AREA_DEFAULT).fillna("CONSUMER")

    # 4) Cruce huellas (pallet/caja). No hay cruce idh (el del PQ no se usa).
    p("Cruzando catálogos de exportación (huellas)…", 97)
    try:
        huellas, _ = io_utils.read_huellas()
        expo = expo.merge(huellas.rename(columns={"producto": "material"}), on="material", how="left")
    except FileNotFoundError:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontró huellas.xlsx. Los servicios de exportación quedarán vacíos.",
            }
        )
        expo["pallet"] = pd.NA
        expo["caja"] = pd.NA
    except Exception as exc:
        emit({"severity": "warning", "msg": f"Error leyendo huellas para exportación: {exc}"})
        expo["pallet"] = pd.NA
        expo["caja"] = pd.NA

    # 5) Medidas: pallets=ceil(cant/pallet), cajas=ceil(cant/caja),
    #    unidades = caja * cajas (null si caja o cajas null). NO es `cantidad`.
    cantidad = expo["cantidad"]
    pallet = pd.to_numeric(expo["pallet"], errors="coerce")
    caja = pd.to_numeric(expo["caja"], errors="coerce")
    pallets = pd.to_numeric(_ceil_series(cantidad, pallet), errors="coerce")
    cajas = pd.to_numeric(_ceil_series(cantidad, caja), errors="coerce")
    und = pd.Series([pd.NA] * len(expo), index=expo.index, dtype=object)
    und_valid = caja.notna() & cajas.notna()
    und.loc[und_valid] = (caja.loc[und_valid] * cajas.loc[und_valid]).astype("int64")
    unidades = pd.to_numeric(und, errors="coerce")

    expo["_pallets"] = pallets
    expo["_cajas"] = cajas
    expo["_unidades"] = unidades

    # 6) Agrupar por (negocio, canal) sumando las medidas (nf = negocio, así no hace
    #    falta incluirlo en la clave). `fecha` del nombre se descarta (usa `periodo`).
    grouped = (
        expo.groupby(["negocio", "canal"], dropna=False)
        .agg(
            pallets=("_pallets", "sum"),
            cajas=("_cajas", "sum"),
            unidades=("_unidades", "sum"),
        )
        .reset_index()
    )

    # 7) Servicios finales (unpivot de logica.txt): servicio = base & " " & canal.
    servicios = []
    for _, r in grouped.iterrows():
        neg = str(r["negocio"]) if pd.notna(r["negocio"]) else ""
        canal = str(r["canal"]).strip() if pd.notna(r["canal"]) else ""
        canal_suf = f" {canal}" if canal else ""
        und = _safe_num(r["unidades"])
        v_p = _safe_num(r["pallets"])
        if v_p:
            servicios.append(_expo_servicio(neg, f"{_EXPO_BASE_PALLETS}{canal_suf}", v_p, und))
        v_c = _safe_num(r["cajas"])
        if v_c:
            servicios.append(_expo_servicio(neg, f"{_EXPO_BASE_CAJAS}{canal_suf}", v_c, und))
        if und:
            servicios.append(_expo_servicio(neg, f"{_EXPO_BASE_UND}{canal_suf}", und, und))

    servicios.sort(key=lambda s: (s["negocio"], s["servicio"]))
    return servicios


# ---------------------------------------------------------------------------
# Pipeline ETIQUETAS (Paso 8) — etiquetado / reempaque
# ---------------------------------------------------------------------------

def _etiq_servicio(servicio: str, valor, unidades: int = 0) -> dict:
    """Construye una línea de servicio de ETIQUETAS con sus constantes (logica.txt).

    negocio = "CONSUMER" (constante del PQ: `#"Personalizada agregada"`). El PQ duplica
    `servicio -> proceso_extendido` (`#"Columna duplicada1"`), así que aquí también
    proceso_extendido = servicio.
    """
    return {
        "negocio": config.ETIQUETAS_NEGOCIO,
        "negocio_facturador": config.ETIQUETAS_NEGOCIO,
        "servicio": servicio,
        "valor": valor,
        "unidades": unidades,
        "proceso_extendido": servicio,
        "macro_proceso": "OTROS",
        "proceso_abreviado": "OTR",
        "tabla": "ETIQUETAS",
    }


def _run_etiquetas_pipeline(emit, progress=None) -> list[dict]:
    """Paso 8 (ETIQUETAS): procesa etiquetas* -> lista de servicios.

    Traducción de la query `etiquetas` de logica.txt. Es tan simple como TRASLADOS:
    **sin huellas, sin idh, sin cruce y sin filtro de fecha** (el PQ no aplica
    RANGOS_FECHAS: solo agrupa por la fecha del NOMBRE del archivo, que el bot descarta
    y reemplaza por `periodo`). El fichero `etiquetas_<fecha>.xlsx` vive en MAQUILA/
    (sin separación cons/prof); negocio es siempre "CONSUMER" (constante del PQ).

    1. Lee y concatena los etiquetas* (estricto: error grave si un archivo falla).
       Sólo `cajas` (la medida que se suma -> valor) y `tipo` (que mapea al servicio).
    2. `tipo` -> servicio (ReplaceValue de logica.txt: REGULARES -> "ETIQUETAS
       REGULARES", REEMPAQUE -> "ETIQUETAS REEMPAQUE"; otros valores pasan tal cual).
    3. Agrupa por `servicio` sumando `cajas` -> `valor`.
    4. Emite una línea por servicio (negocio = CONSUMER, unidades = 0, proceso_extendido
       = servicio, macro_proceso = "OTROS", proceso_abreviado = "OTR", tabla = "ETIQUETAS").
       Convención del bot: se omiten los servicios con valor 0.

    Sin archivos -> warning y []. Error de lectura -> BlockingError.
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    p("Procesando etiquetas…", 97)
    files = io_utils.find_etiquetas_files()
    if not files:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos etiquetas* (Paso 8 omitido).",
            }
        )
        return []

    frames, ok_names = [], []
    for path in files:
        try:
            frames.append(io_utils.read_etiquetas(path))
            ok_names.append(path.name)
        except Exception as exc:  # corrupto, sin columnas, hoja rara
            raise BlockingError(
                f"{path.name}: no se pudo leer el archivo de etiquetas ({exc})."
            ) from exc

    etiq = pd.concat(frames, ignore_index=True)
    if etiq.empty:
        emit(
            {
                "severity": "warning",
                "msg": "Los archivos de etiquetas están vacíos (Paso 8 omitido).",
            }
        )
        return []

    # 2) tipo -> servicio (logica.txt: ReplaceValue). Otros valores pasan tal cual.
    etiq["tipo"] = etiq["tipo"].astype(str).str.strip().str.upper()
    etiq["cajas"] = pd.to_numeric(etiq["cajas"], errors="coerce").fillna(0)
    etiq["servicio"] = etiq["tipo"].map(config.ETIQUETAS_SERVICIO_MAP).fillna(etiq["tipo"])

    # 3) Agrupar por servicio sumando cajas (la fecha del nombre se descarta -> periodo).
    grouped = (
        etiq.groupby("servicio", dropna=False)
        .agg(cajas=("cajas", "sum"))
        .reset_index()
    )

    # 4) Servicios finales (omitiendo valor 0, convención del bot).
    servicios = []
    for _, r in grouped.iterrows():
        servicio = str(r["servicio"]).strip()
        valor = int(r["cajas"])
        if valor:
            servicios.append(_etiq_servicio(servicio, valor))

    servicios.sort(key=lambda s: s["servicio"])
    return servicios


# ---------------------------------------------------------------------------
# Pipeline PALETIZADO (Paso 9) — paletizado de exportación
# ---------------------------------------------------------------------------

def _pal_servicio(negocio: str, servicio: str, valor, unidades: int = 0) -> dict:
    """Construye una línea de servicio de PALETIZADO con sus constantes (logica.txt).

    `negocio_facturador` = `negocio` (el PQ duplica `negocio -> negocio_facturador`,
    `#"Columna duplicada"`). `proceso_extendido` es la constante "EXPORTACION
    (PALETIZADO)" (no depende de negocio ni servicio).
    """
    return {
        "negocio": negocio,
        "negocio_facturador": negocio,
        "servicio": servicio,
        "valor": valor,
        "unidades": unidades,
        "proceso_extendido": "EXPORTACION (PALETIZADO)",
        "macro_proceso": "OTROS",
        "proceso_abreviado": "PAL",
        "tabla": "PALETIZADO",
    }


def _run_paletizado_pipeline(emit, progress=None) -> list[dict]:
    """Paso 9 (PALETIZADO): procesa paletizado* -> lista de servicios.

    Traducción de la query `paletizado` de logica.txt. Tan simple como ETIQUETAS/
    TRASLADOS: **sin huellas, sin idh, sin cruce y sin filtro de fecha** (el PQ no aplica
    RANGOS_FECHAS: solo agrupa por la fecha del NOMBRE del archivo, que el bot descarta
    y reemplaza por `periodo`). El fichero `paletizado_<fecha>.xlsx` vive en EXPORTACIONES/.

    A diferencia del resto de pasos, el `negocio` se deriva de la columna **AREA**
    (HENKEL.PF -> PROFESIONAL, HENKEL.RT -> CONSUMER), no del prefijo del nombre ni del
    idh. El `canal` forma parte del nombre del servicio (como en EXPORTACIONES).

    1. Lee y concatena los paletizado* (estricto: error grave si un archivo falla).
       Sólo `area` (-> negocio), `despacho` (filtro), `total` (-> valor) y `canal` (->
       sufijo del servicio).
    2. Filtra filas con `DESPACHO` <> null (logica.txt: #"Filas filtradas1").
    3. `negocio` = AREA mapeada (HENKEL.PF/HENKEL.RT); otros valores pasan tal cual.
       Descarta filas con AREA nula (filtro final negocio <> null).
    4. Agrupa por (negocio, canal) sumando `total` -> `valor` (la fecha del nombre se
       descarta -> periodo).
    5. Servicios: `servicio = "PALETIZADO EXPO " + canal`; negocio_facturador = negocio;
       unidades = 0; proceso_extendido = "EXPORTACION (PALETIZADO)", macro_proceso =
       "OTROS", proceso_abreviado = "PAL", tabla = "PALETIZADO". Convención del bot: se
       omiten los servicios con valor 0.

    Sin archivos -> warning y []. Error de lectura -> BlockingError.
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    p("Procesando paletizado…", 97)
    files = io_utils.find_paletizado_files()
    if not files:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos paletizado* (Paso 9 omitido).",
            }
        )
        return []

    frames, ok_names = [], []
    for path in files:
        try:
            frames.append(io_utils.read_paletizado(path))
            ok_names.append(path.name)
        except Exception as exc:  # corrupto, sin columnas, hoja rara
            raise BlockingError(
                f"{path.name}: no se pudo leer el archivo de paletizado ({exc})."
            ) from exc

    pal = pd.concat(frames, ignore_index=True)
    if pal.empty:
        emit(
            {
                "severity": "warning",
                "msg": "Los archivos de paletizado están vacíos (Paso 9 omitido).",
            }
        )
        return []

    # 2) Filas con DESPACHO <> null (logica.txt: #"Filas filtradas1").
    pal = pal.loc[pal["despacho"].notna()].copy()
    pal["canal"] = pal["canal"].astype(str).str.strip().str.upper()
    pal["total"] = pd.to_numeric(pal["total"], errors="coerce").fillna(0)

    # 3) negocio = AREA mapeada; descarta AREA nula (filtro final negocio <> null).
    #    Otros valores de AREA pasan tal cual (el PQ sólo reemplaza HENKEL.PF/HENKEL.RT).
    pal = pal.loc[pal["area"].notna()].copy()
    area_upper = pal["area"].astype(str).str.strip().str.upper()
    pal["negocio"] = area_upper.map(config.PALETIZADO_AREA_MAP).fillna(area_upper)
    if pal.empty:
        return []

    # 4) Agrupar por (negocio, canal) sumando total (la fecha del nombre se descarta).
    grouped = (
        pal.groupby(["negocio", "canal"], dropna=False)
        .agg(total=("total", "sum"))
        .reset_index()
    )

    # 5) Servicios finales (servicio = "PALETIZADO EXPO " + canal; omitir valor 0).
    servicios = []
    for _, r in grouped.iterrows():
        neg = str(r["negocio"]).strip()
        canal = str(r["canal"]).strip() if pd.notna(r["canal"]) else ""
        canal_suf = f" {canal}" if canal else ""
        valor = int(r["total"])
        if valor:
            servicios.append(_pal_servicio(neg, f"PALETIZADO EXPO{canal_suf}", valor))

    servicios.sort(key=lambda s: (s["negocio"], s["servicio"]))
    return servicios


# ---------------------------------------------------------------------------
# Pipeline TRINCAJE (Paso 10) — trincaje de exportación
# ---------------------------------------------------------------------------

def _tri_servicio(negocio: str, servicio: str, valor, unidades: int = 0) -> dict:
    """Construye la (única) línea de servicio de TRINCAJE con sus constantes (logica.txt).

    `negocio_facturador` = `negocio` (el PQ duplica `negocio -> negocio_facturador`,
    `#"Columna duplicada"`). `proceso_extendido` = `servicio` = "DOBLE TRINCAJE" (el PQ
    duplica `proceso_extendido -> servicio`, `#"Columna duplicada1"`).
    """
    return {
        "negocio": negocio,
        "negocio_facturador": negocio,
        "servicio": servicio,
        "valor": valor,
        "unidades": unidades,
        "proceso_extendido": servicio,
        "macro_proceso": "OTROS",
        "proceso_abreviado": "OTR",
        "tabla": "TRINCAJE",
    }


def _run_trincaje_pipeline(emit, progress=None) -> list[dict]:
    """Paso 10 (TRINCAJE): procesa trincaje* -> lista de servicios.

    Traducción de la query `trincaje` de logica.txt. Es la pipeline **más simple de
    todas** (junto a TRASLADOS/ETIQUETAS): **sin huellas, sin idh, sin cruce y sin filtro
    de fecha** (el PQ no aplica RANGOS_FECHAS: solo agrupa por la fecha del NOMBRE del
    archivo, que el bot descarta y reemplaza por `periodo`). El fichero
    `trincaje_<fecha>.xlsx` vive en EXPORTACIONES/.

    El `valor` es el **conteo de filas × 2** (logica.txt: `valor = Table.RowCount(_) * 2`
    por cada fecha del nombre). `negocio` es siempre "CONSUMER" (constante del PQ, no se
    deriva del nombre ni del contenido). Genera **una sola línea**: servicio =
    "DOBLE TRINCAJE".

    1. Lee y concatena los trincaje* (estricto: error grave si un archivo falla). Sólo
       `despacho` (filtro "<> null").
    2. Filtra filas con `despacho` <> null (logica.txt: #"Filas filtradas").
    3. `valor = (nº de filas) × TRINCAJE_FACTOR` (la fecha del nombre se descarta ->
       periodo); negocio = negocio_facturador = "CONSUMER"; servicio = proceso_extendido
       = "DOBLE TRINCAJE", macro_proceso = "OTROS", proceso_abreviado = "OTR",
       tabla = "TRINCAJE", unidades = 0. Convención del bot: si valor = 0 se omite.

    Sin archivos -> warning y []. Error de lectura -> BlockingError.
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    p("Procesando trincaje…", 97)
    files = io_utils.find_trincaje_files()
    if not files:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos trincaje* (Paso 10 omitido).",
            }
        )
        return []

    frames = []
    for path in files:
        try:
            frames.append(io_utils.read_trincaje(path))
        except Exception as exc:  # corrupto, sin columna despacho, hoja rara
            raise BlockingError(
                f"{path.name}: no se pudo leer el archivo de trincaje ({exc})."
            ) from exc

    tri = pd.concat(frames, ignore_index=True)
    if tri.empty:
        emit(
            {
                "severity": "warning",
                "msg": "Los archivos de trincaje están vacíos (Paso 10 omitido).",
            }
        )
        return []

    # 2) Filas con despacho <> null (logica.txt: #"Filas filtradas").
    tri = tri.loc[tri["despacho"].notna()].copy()

    # 3) valor = conteo de filas × TRINCAJE_FACTOR (la fecha del nombre se descarta).
    valor = int(len(tri) * config.TRINCAJE_FACTOR)
    if not valor:
        return []

    servicios = [_tri_servicio(config.TRINCAJE_NEGOCIO, config.TRINCAJE_SERVICIO, valor)]
    return servicios


# ---------------------------------------------------------------------------
# Pipeline PLANTA (Paso 11) — traslado de pallets planta → CEDI
# ---------------------------------------------------------------------------

def _pla_servicio(negocio: str, servicio: str, valor, unidades: int = 0) -> dict:
    """Construye una línea de servicio de PLANTA con sus constantes (logica.txt).

    `negocio_facturador` = `negocio` (el PQ duplica `negocio -> negocio_facturador`,
    `#"Columna duplicada"`). `proceso_extendido` es la constante "TRASLADO" y el
    `servicio` la constante "TRASLADO PALLETS PLANTA - CEDI" (ninguno depende del
    negocio).
    """
    return {
        "negocio": negocio,
        "negocio_facturador": negocio,
        "servicio": servicio,
        "valor": valor,
        "unidades": unidades,
        "proceso_extendido": "TRASLADO",
        "macro_proceso": "OTROS",
        "proceso_abreviado": "TRA",
        "tabla": "PLANTA",
    }


def _run_planta_pipeline(emit, progress=None) -> list[dict]:
    """Paso 11 (PLANTA): procesa planta* -> lista de servicios.

    Traducción de la query `planta` de logica.txt. Tan simple como TRINCAJE/PALETIZADO:
    **sin huellas, sin idh, sin equivalencias y sin filtro de fecha** (el PQ elimina las
    cols `fecha`/`semana` del archivo y no aplica RANGOS_FECHAS: agrupa por la fecha del
    NOMBRE del archivo, que el bot descarta y reemplaza por `periodo`). Un SOLO fichero
    `planta_<fecha>.xlsx` cubre ambos negocios.

    El PQ unpivotea `estibas_consumer`/`estibas_profesional` en una columna `negocio` y
    agrupa por `(fecha del nombre, negocio)` sumando -> valor. El bot descarta la fecha
    del nombre y suma por `negocio` (CONSUMER/PROFESIONAL) sobre todos los `planta*`.

    1. Lee y concatena los planta* (estricto: error grave si un archivo falla). Sólo
       `consumer` (= estibas_consumer) y `profesional` (= estibas_profesional).
    2. `valor_CONSUMER = sum(estibas_consumer)`, `valor_PROFESIONAL = sum(estibas_profesional)`.
    3. Servicios (uno por negocio con valor > 0): `servicio = "TRASLADO PALLETS PLANTA -
       CEDI"`; negocio_facturador = negocio; unidades = 0; proceso_extendido = "TRASLADO",
       macro_proceso = "OTROS", proceso_abreviado = "TRA", tabla = "PLANTA".

    Sin archivos -> warning y []. Error de lectura -> BlockingError.
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    p("Procesando planta…", 97)
    files = io_utils.find_planta_files()
    if not files:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos planta* (Paso 11 omitido).",
            }
        )
        return []

    frames = []
    for path in files:
        try:
            frames.append(io_utils.read_planta(path))
        except Exception as exc:  # corrupto, sin estibas_consumer/profesional, hoja rara
            raise BlockingError(
                f"{path.name}: no se pudo leer el archivo de planta ({exc})."
            ) from exc

    pla = pd.concat(frames, ignore_index=True)
    if pla.empty:
        emit(
            {
                "severity": "warning",
                "msg": "Los archivos de planta están vacíos (Paso 11 omitido).",
            }
        )
        return []

    # 2) Sumar estibas por negocio (la fecha del nombre se descarta -> periodo).
    sumas = {
        "consumer": int(pd.to_numeric(pla["consumer"], errors="coerce").fillna(0).sum()),
        "profesional": int(
            pd.to_numeric(pla["profesional"], errors="coerce").fillna(0).sum()
        ),
    }

    # 3) Servicios finales (uno por negocio con valor > 0).
    servicios = []
    for key, valor in sumas.items():
        if valor:
            servicios.append(
                _pla_servicio(config.PLANTA_NEGOCIOS[key], config.PLANTA_SERVICIO, valor)
            )
    servicios.sort(key=lambda s: s["negocio"])
    return servicios


# ---------------------------------------------------------------------------
# Pipeline MATERIAL (Paso 12) — almacenamiento material de empaque (Bodega 8 ME)
# ---------------------------------------------------------------------------

def _mat_servicio(negocio: str, nf: str, servicio: str, valor, unidades: int = 0) -> dict:
    """Construye la (única) línea de servicio de MATERIAL con sus constantes (logica.txt).

    Todas las constantes vienen fijas en la query `FilesOcupacionMaterial`. OJO:
    `negocio_facturador` = "MATERIAL EMPAQUE" (SIN "DE"), distinto a `negocio`.
    """
    return {
        "negocio": negocio,
        "negocio_facturador": nf,
        "servicio": servicio,
        "valor": valor,
        "unidades": unidades,
        "proceso_extendido": "PALLETS",
        "macro_proceso": "ALMACENAMIENTO",
        "proceso_abreviado": "WHS",
        "tabla": "MATERIAL",
    }


def _run_material_pipeline(
    start: str,
    end: str,
    ts_start: pd.Timestamp,
    ts_end: pd.Timestamp,
    emit,
    progress=None,
) -> list[dict]:
    """Paso 12 (MATERIAL): procesa ocupacionMaterial* -> lista de servicios.

    Traducción de la query `FilesOcupacionMaterial` de logica.txt. Gemelo SIMPLIFICADO de
    Ocupación (Paso 4): una sola medida (`cant_bodega_8`), servicio/negocio/nf fijos y sin
    join EQUIVALENCIAS. El valor es el **promedio diario redondeado hacia arriba (ceil)** de
    `cant_bodega_8` en el rango.

    El PQ filtra los días por la ventana de RANGOS_FECHAS (joined por la fecha del NOMBRE
    del archivo, sobre la `fecha` del contenido). El bot la reemplaza por el **rango libre
    del calendario** del usuario sobre la columna `fecha` (igual que Ocupación).

    1. Lee y concatena los ocupacionMaterial* (estricto: error grave si un archivo falla).
       Sólo `fecha` y `cant_bodega_8` (-> valor).
    2. Filtra `fecha` por el rango del usuario (inclusivo).
    3. `valor = ceil(promedio(cant_bodega_8))` sobre los días en rango
       (`Number.RoundUp(List.Average([valor]))` del PQ).
    4. Servicio final (**una línea**): `negocio = "MATERIAL DE EMPAQUE"`,
       `negocio_facturador = "MATERIAL EMPAQUE"` (sin "DE"), `servicio = "ALMACENAMIENTO
       PALLET BODEGA 8 ME GENERAL"`, `unidades = 0`, `proceso_extendido = "PALLETS"`,
       `macro_proceso = "ALMACENAMIENTO"`, `proceso_abreviado = "WHS"`, `tabla = "MATERIAL"`.
       Convención del bot: si `valor = 0` se omite.

    Sin archivos -> warning y []. 0 filas en rango -> warning y []. Error de lectura ->
    BlockingError.
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    p("Procesando material…", 97)
    files = io_utils.find_material_files()
    if not files:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos ocupacionMaterial* (Paso 12 omitido).",
            }
        )
        return []

    frames = []
    for path in files:
        try:
            frames.append(io_utils.read_material(path))
        except Exception as exc:  # corrupto, sin fecha/cant_bodega_8, hoja rara
            raise BlockingError(
                f"{path.name}: no se pudo leer el archivo de material ({exc})."
            ) from exc

    mat = pd.concat(frames, ignore_index=True)
    if mat.empty:
        emit(
            {
                "severity": "warning",
                "msg": "Los archivos de material están vacíos (Paso 12 omitido).",
            }
        )
        return []

    # 2) fecha -> datetime; filtra fecha no nula y por el rango del usuario (inclusivo).
    mat["_fecha"] = pd.to_datetime(mat["fecha"], errors="coerce").dt.normalize()
    mat = mat.loc[mat["_fecha"].notna()].copy()
    mat = mat.loc[(mat["_fecha"] >= ts_start) & (mat["_fecha"] <= ts_end)].copy()
    if mat.empty:
        emit(
            {
                "severity": "warning",
                "msg": f"0 filas de material con 'fecha' en el rango {start} – {end}.",
            }
        )
        return []

    # 3) valor = ceil(promedio(cant_bodega_8)) sobre los días en rango.
    valores = pd.to_numeric(mat["valor"], errors="coerce").dropna()
    if valores.empty:
        return []
    valor = int(math.ceil(float(valores.mean())))
    if not valor:
        return []

    # 4) Servicio final (una línea).
    return [
        _mat_servicio(
            config.MATERIAL_NEGOCIO,
            config.MATERIAL_NEGOCIO_FACTURADOR,
            config.MATERIAL_SERVICIO,
            valor,
        )
    ]


# ---------------------------------------------------------------------------
# Pipeline FALABELLA (Paso 13) — etiquetado Falabella (conteo de entregas)
# ---------------------------------------------------------------------------

def _fal_servicio(negocio: str, servicio: str, valor, unidades: int = 0) -> dict:
    """Construye una línea de servicio de FALABELLA con sus constantes (logica.txt).

    `negocio_facturador` = `negocio` (el PQ duplica `negocio -> negocio_facturador`,
    `#"Columna duplicada"`). `proceso_extendido` = `servicio` = "ETIQUETADO (FALABELLA)"
    (el PQ duplica `proceso_extendido -> servicio`, `#"Personalizada agregada"` y
    `#"Personalizada agregada1"` con el mismo valor).
    """
    return {
        "negocio": negocio,
        "negocio_facturador": negocio,
        "servicio": servicio,
        "valor": valor,
        "unidades": unidades,
        "proceso_extendido": servicio,
        "macro_proceso": "OTROS",
        "proceso_abreviado": "OTR",
        "tabla": "FALABELLA",
    }


def _run_falabella_pipeline(emit, progress=None) -> list[dict]:
    """Paso 13 (FALABELLA): procesa falabella* -> lista de servicios.

    Traducción de la query `falabella` de logica.txt. Tan simple como
    TRASLADOS/TRINCAJE/PLANTA: **sin huellas, sin idh, sin cruce y sin filtro de fecha**
    (el PQ no aplica RANGOS_FECHAS: agrupa por la fecha del NOMBRE del archivo, que el
    bot descarta y reemplaza por `periodo`). El fichero `falabella_<negocio>_<fecha>.xlsx`
    vive en CONSUMER/ (falabella_cons*) o PROFESIONAL/ (falabella_prof*).

    El `valor` es el **conteo de filas con `Entrega <> null`** (logica.txt:
    `valor = Table.RowCount(_)` por cada `(fecha, negocio)`); el bot descarta la fecha
    del nombre y suma por `negocio` (CONSUMER/PROFESIONAL vía AREA_DEFAULT) sobre todos
    los `falabella*`.

    1. Lee y concatena los falabella* (estricto: error grave si un archivo falla). Sólo
       `entrega` (filtro "<> null"). `MES` no se usa (el PQ la descarta).
    2. Filtra filas con `Entrega` <> null (logica.txt: #"Filas filtradas").
    3. Cuenta por `area` -> valor por negocio; servicio = proceso_extendido =
       "ETIQUETADO (FALABELLA)", negocio_facturador = negocio, macro_proceso = "OTROS",
       proceso_abreviado = "OTR", tabla = "FALABELLA", unidades = 0. Convención del bot:
       si valor = 0 se omite.

    Sin archivos -> warning y []. Error de lectura -> BlockingError.
    """
    def p(stage, pct):
        if progress:
            progress(stage, pct)

    p("Procesando falabella…", 97)
    files = io_utils.find_falabella_files()
    if not files:
        emit(
            {
                "severity": "warning",
                "msg": "No se encontraron archivos falabella* (Paso 13 omitido).",
            }
        )
        return []

    frames = []
    for path, area in files:
        try:
            frames.append(io_utils.read_falabella(path, area))
        except Exception as exc:  # corrupto, sin columna Entrega, hoja rara
            raise BlockingError(
                f"{path.name}: no se pudo leer el archivo de falabella ({exc})."
            ) from exc

    fal = pd.concat(frames, ignore_index=True)
    if fal.empty:
        emit(
            {
                "severity": "warning",
                "msg": "Los archivos de falabella están vacíos (Paso 13 omitido).",
            }
        )
        return []

    # 2) Filas con Entrega <> null (logica.txt: #"Filas filtradas"); conteo por area.
    fal = fal.loc[fal["entrega"].notna()].copy()

    # 3) Servicios finales (uno por negocio con valor > 0). La fecha del nombre se descarta.
    servicios = []
    for area, grupo in fal.groupby("area"):
        valor = int(len(grupo))
        if valor:
            servicios.append(
                _fal_servicio(config.AREA_DEFAULT[area], config.FALABELLA_SERVICIO, valor)
            )
    servicios.sort(key=lambda s: s["negocio"])
    return servicios


# ---------------------------------------------------------------------------
# Orquestador: SALIDAS + DESTRUCCIÓN + INGRESOS + OCUPACIÓN
# ---------------------------------------------------------------------------

def _periodo_for_range(end: str) -> str:
    """Primer día del ÚLTIMO mes del rango (el mes de la fecha final), en dd/mm/yyyy.

    Es la marca de periodo que se coloca igual en todos los registros del Excel.
    P. ej. rango 21/05–21/06 -> 01/06 (del año de la fecha final).
    """
    d = datetime.strptime(end.strip(), config.USER_DATE_FORMAT).date()
    return d.replace(day=1).strftime(config.USER_DATE_FORMAT)


def _aggregate_servicios(servicios: list[dict], periodo: str) -> list[dict]:
    """Agrega las líneas por (periodo, negocio, negocio_facturador, servicio, tabla):
    suma `valor` y `unidades`; las columnas descriptivas se conservan (first). Así
    `servicio` es un nivel de agrupación (las ramas que mapeaban al mismo servicio
    con distinto tipo_trabajo/tipo_despacho quedan sumadas en una sola línea).
    """
    if not servicios:
        return []
    df = pd.DataFrame(servicios)
    df["periodo"] = periodo
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0)
    df["unidades"] = pd.to_numeric(df["unidades"], errors="coerce").fillna(0)
    group_keys = ["periodo", "negocio", "negocio_facturador", "servicio", "tabla"]
    agg = (
        df.groupby(group_keys, dropna=False)
        .agg(
            valor=("valor", "sum"),
            unidades=("unidades", "sum"),
            proceso_extendido=("proceso_extendido", "first"),
            macro_proceso=("macro_proceso", "first"),
            proceso_abreviado=("proceso_abreviado", "first"),
        )
        .reset_index()
    )
    order = [
        "periodo", "negocio", "negocio_facturador", "servicio", "valor", "unidades",
        "proceso_extendido", "macro_proceso", "proceso_abreviado", "tabla",
    ]
    agg = agg[order]
    agg["valor"] = agg["valor"].astype(int)
    agg["unidades"] = agg["unidades"].astype(int)
    agg = agg.sort_values(["tabla", "negocio", "negocio_facturador", "servicio"]).reset_index(drop=True)
    return agg.to_dict("records")


def _apply_tarifas(servicios: list[dict], emit) -> list[dict]:
    """Cruza cada línea agregada con `tarifas.xlsx` (llave `servicio`) y calcula el costo.

    - Lee `tarifas` (hoja activa, una tarifa por `servicio`); la **fecha no importa**
      (decisión del usuario): se arma un lookup plano `servicio normalizado -> {um, tarifa}`.
      Si hubiera varias filas para el mismo `servicio` con `valor` distinto, gana la última
      y se avisa (en la hoja activa actual no ocurre: 1 fila por servicio).
    - Para cada línea con match: añade `um`, `tarifa` (= tarifas.valor) y `costo_total` =
      `valor × tarifa` redondeado a 2 decimales.
    - Las líneas **sin tarifa** se **eliminan del Excel** (decisión del usuario: "las que no
      no las debes tener en cuenta en el excel final"). **No se reportan**: no impiden facturar
      (son servicios que simplemente no tienen tarifa configurada); el panel de avisos es solo
      para errores que impidan facturar.
    - También se eliminan las líneas con `valor == 0` ("las cosas que tengan como valor cero
      no deben aparecer").
    - Si `tarifas.xlsx` no existe/no se puede leer: aviso y **degradación elegante** — se
      conservan todas las líneas sin columnas de costo (no se filtra por tarifa). Así un
      auxiliar ausente no vacía el Excel.

    `emit(issue)` reporta los avisos (tarifas ausente, duplicados) — **no** los servicios sin tarifa.
    """
    # 1) Construir el lookup plano servicio -> {um, tarifa}.
    try:
        df_tar = io_utils.read_tarifas()
    except (FileNotFoundError, KeyError, ValueError) as exc:
        emit(
            {
                "severity": "warning",
                "msg": f"tarifas.xlsx no disponible ({exc}); no se calcula costo ni se filtra por tarifa.",
            }
        )
        for s in servicios:
            s.setdefault("um", "")
            s.setdefault("tarifa", None)
            s.setdefault("costo_total", None)
        return servicios

    lookup: dict[str, dict] = {}
    conflicts = []
    for rec in df_tar.to_dict(orient="records"):
        serv = rec.get("servicio")
        if serv is None or (isinstance(serv, float) and pd.isna(serv)):
            continue
        key = io_utils.normalize(serv)
        tarifa = rec.get("valor")
        tarifa = None if pd.isna(tarifa) else float(tarifa)
        um = rec.get("um")
        um = "" if pd.isna(um) else str(um).strip()
        prev = lookup.get(key)
        if prev is not None and prev["tarifa"] is not None and tarifa is not None and prev["tarifa"] != tarifa:
            conflicts.append(str(serv))
        lookup[key] = {"um": um, "tarifa": tarifa}
    if conflicts:
        emit(
            {
                "severity": "warning",
                "msg": f"tarifas.xlsx tiene tarifas distintas para el mismo servicio (se usa la última): {sorted(set(conflicts))}.",
            }
        )

    # 2) Enriquecer/filtrar líneas. Las líneas sin tarifa o con valor 0 simplemente se
    #    omiten del Excel (no se reportan: no impiden facturar — son servicios sin tarifa
    #    configurada; el panel de avisos es solo para errores que impidan facturar).
    enriquecidas = []
    for s in servicios:
        valor = s.get("valor", 0)
        try:
            valor_num = float(valor)
        except (TypeError, ValueError):
            valor_num = 0.0
        # Descartar valor cero (no debe aparecer en el Excel).
        if valor_num == 0:
            continue
        key = io_utils.normalize(s.get("servicio", ""))
        tar = lookup.get(key)
        if tar is None or tar["tarifa"] is None:
            continue
        costo = round(valor_num * tar["tarifa"], 2)
        s2 = dict(s)
        s2["um"] = tar["um"]
        s2["tarifa"] = tar["tarifa"]
        s2["costo_total"] = costo
        enriquecidas.append(s2)

    return enriquecidas


def run_all(start: str, end: str, *, progress=None, on_issue=None) -> Step1Result:
    """Ejecuta Paso 1 (SALIDAS) + Paso 2 (DESTRUCCIÓN) + Paso 3 (INGRESOS) +
    Paso 4 (OCUPACIÓN) + Paso 5 (TRASLADOS) + Paso 6 (MAQUILA) + Paso 7 (EXPORTACIONES)
    + Paso 8 (ETIQUETAS) + Paso 9 (PALETIZADO) + Paso 10 (TRINCAJE) + Paso 11 (PLANTA)
    + Paso 12 (MATERIAL) + Paso 13 (FALABELLA) y combina los servicios.

    Lanza ValueError ante rango inválido y BlockingError ante errores graves de
    archivo. `progress(stage, percent)` y `on_issue(issue)` permiten a la API
    mostrar avance y problemas en vivo. Cachea el resultado por rango.
    """
    ts_start, ts_end = parse_date_range(start, end)
    generated_at = datetime.now().isoformat(timespec="seconds")
    issues: list[dict] = []

    def emit(issue):
        issues.append(issue)
        if on_issue:
            on_issue(issue)

    def p(stage, pct):
        if progress:
            progress(stage, pct)

    # Paso 1: SALIDAS.
    totals, diagnostics = _run_salidas_pipeline(start, end, ts_start, ts_end, emit, p)

    # Paso 2: DESTRUCCIÓN.
    p("Leyendo archivos de destrucción…", 60)
    destr_services = run_step2_destruccion(emit, p)

    # Paso 3: INGRESOS.
    ingr_services = _run_ingresos_pipeline(start, end, ts_start, ts_end, emit, p)

    # Paso 4: OCUPACIÓN.
    ocup_services = _run_ocupacion_pipeline(start, end, ts_start, ts_end, emit, p)

    # Paso 5: TRASLADOS.
    trasl_services = _run_traslados_pipeline(emit, p)

    # Paso 6: MAQUILA.
    maquila_services = _run_maquila_pipeline(start, end, ts_start, ts_end, emit, p)

    # Paso 7: EXPORTACIONES.
    expo_services = _run_exportacion_pipeline(start, end, ts_start, ts_end, emit, p)

    # Paso 8: ETIQUETAS.
    etiq_services = _run_etiquetas_pipeline(emit, p)

    # Paso 9: PALETIZADO.
    pal_services = _run_paletizado_pipeline(emit, p)

    # Paso 10: TRINCAJE.
    tri_services = _run_trincaje_pipeline(emit, p)

    # Paso 11: PLANTA.
    pla_services = _run_planta_pipeline(emit, p)

    # Paso 12: MATERIAL.
    mat_services = _run_material_pipeline(start, end, ts_start, ts_end, emit, p)

    # Paso 13: FALABELLA.
    fal_services = _run_falabella_pipeline(emit, p)

    # Combinar servicios (SALIDAS + DESTRUCCIÓN + INGRESOS + OCUPACIÓN + TRASLADOS +
    # MAQUILA + EXPORTACIONES + ETIQUETAS + PALETIZADO + TRINCAJE + PLANTA + MATERIAL +
    # FALABELLA), periodo y agregar.
    servicios = (
        list(totals.get("servicios", []))
        + destr_services
        + ingr_services
        + ocup_services
        + trasl_services
        + maquila_services
        + expo_services
        + etiq_services
        + pal_services
        + tri_services
        + pla_services
        + mat_services
        + fal_services
    )
    periodo = _periodo_for_range(end)
    servicios = _aggregate_servicios(servicios, periodo)
    # Tarifa → costo: cruza por `servicio` con tarifas.xlsx, calcula costo_total y
    # elimina las líneas sin tarifa (y las de valor 0).
    servicios = _apply_tarifas(servicios, emit)
    combined = dict(totals)
    combined["servicios"] = servicios

    if not servicios:
        issues.append(
            {
                "severity": "error",
                "msg": "No hay servicios para facturar en el rango seleccionado (ni salidas ni destrucción).",
            }
        )

    p("Construyendo Excel…", 98)
    return _finalize(combined, diagnostics, issues, generated_at)


def _finalize(totals, diagnostics, issues, generated_at) -> Step1Result:
    """Construye el Step1Result y lo cachea por rango."""
    result = Step1Result(
        totals=totals,
        diagnostics=diagnostics,
        issues=issues,
        generated_at=generated_at,
    )
    key = f"{diagnostics.get('range_start')}|{diagnostics.get('range_end')}"
    _RESULT_CACHE[key] = result
    return result


def run_all_cached(start: str, end: str) -> Step1Result:
    """Igual que run_all pero aprovecha la caché por rango si existe."""
    key = f"{start.strip()}|{end.strip()}"
    cached = _RESULT_CACHE.get(key)
    if cached is not None:
        return cached
    return run_all(start, end)


# Alias legacy: el "step1" ahora incluye también la destrucción (run_all).
def run_step1(start: str, end: str, *, base_dir: Optional[Path] = None) -> Step1Result:
    return run_all(start, end)


def run_step1_cached(start: str, end: str) -> Step1Result:
    return run_all_cached(start, end)


# ---------------------------------------------------------------------------
# AddServicio + FilterServicio (traducción de logica.txt, ~15 ramas)
# ---------------------------------------------------------------------------

# Medida (atributo en logica.txt) -> campo agregado de cada grupo.
_PICKING_MEDIDAS = (
    ("PICKING PALLETS", "estibas"),
    ("PICKING CAJAS", "cajas"),
    ("PICKING UNIDADES", "unidades"),
)


def _servicio(
    negocio_facturador: str, tipo_trabajo: str, tipo_despacho: str, atributo: str
) -> Optional[str]:
    """Clasifica el servicio según (negocio_facturador, tipo_trabajo, tipo_despacho, atributo).

    Réplica del `AddServicio` de logica.txt: cadena if/elseif, gana la 1ª rama que
    coincide. Devuelve None si ningún servicio aplica (esas líneas se descartan luego).
    """
    nf, tt, td, a = negocio_facturador, tipo_trabajo, tipo_despacho, atributo
    if nf == "LAUNDRY" and tt == "NORMAL" and td == "ESTANDAR" and a == "PICKING UNIDADES":
        return "ALISTAMIENTO Y DESPACHO UNIDADES NORMAL LAUNDRY"
    if nf == "LAUNDRY" and tt == "EXTRA E.S" and td == "CROSS DOCKING" and a == "PICKING UNIDADES":
        return "DESTELLE UNIDADES CROSS DOCKING EXTRA E.S LAUNDRY"
    if nf == "LAUNDRY" and tt == "NORMAL" and td == "CROSS DOCKING" and a == "PICKING UNIDADES":
        return "DESTELLE UNIDADES CROSS DOCKING NORMAL LAUNDRY"
    if nf == "LAUNDRY" and tt == "EXTRA E.S" and td == "CROSS DOCKING" and a == "PICKING CAJAS":
        return "DESTELLE CAJAS CROSS DOCKING EXTRA E.S LAUNDRY"
    if nf == "LAUNDRY" and td == "ESTANDAR" and a == "PICKING CAJAS":
        return "ALISTAMIENTO Y DESPACHO CAJAS ESTANDAR LAUNDRY"
    if tt == "NORMAL" and td == "ESTANDAR" and a == "PICKING CAJAS":
        return "ALISTAMIENTO Y DESPACHO CAJAS ESTANDAR"
    if tt == "NORMAL" and td == "CROSS DOCKING" and a == "PICKING UNIDADES":
        return "DESTELLE UNIDADES CROSS DOCKING"
    if tt == "NORMAL" and td == "CROSS DOCKING" and a == "PICKING CAJAS":
        return "DESTELLE CAJAS CROSS DOCKING"
    if tt == "EXTRA E.S" and td == "CROSS DOCKING" and a == "PICKING CAJAS":
        return "DESTELLE CAJAS CROSS DOCKING JORNADA EXTENDIDA"
    if tt == "EXTRA E.S" and td == "ESTANDAR" and a == "PICKING CAJAS":
        return "ALISTAMIENTO Y DESPACHO CAJAS ESTANDAR EXTRA E.S"
    if tt == "NORMAL" and td == "ESTANDAR" and a == "PICKING UNIDADES":
        return "ALISTAMIENTO Y DESPACHO UNIDADES ESTANDAR"
    if tt == "EXTRA E.S" and td == "ESTANDAR" and a == "PICKING UNIDADES":
        return "ALISTAMIENTO Y DESPACHO UNIDADES ESTANDAR EXTRA E.S"
    if tt == "EXTRA E.S" and a == "PICKING PALLETS":
        return "ALISTAMIENTO Y DESPACHO PALLETS ESTANDAR EXTRA E.S"
    if a == "PICKING PALLETS":
        return "ALISTAMIENTO Y DESPACHO PALLETS ESTANDAR"
    return None


def _keep_servicio(negocio_facturador: str, servicio: Optional[str]) -> bool:
    """Réplica del `FilterServicio` de logica.txt. False si la línea no se factura."""
    if servicio is None:
        return False
    if servicio == "DESTELLE CAJAS CROSS DOCKING" and negocio_facturador == "LAUNDRY":
        return False
    if (
        negocio_facturador != "LAUNDRY"
        and servicio == "ALISTAMIENTO Y DESPACHO UNIDADES ESTANDAR EXTRA E.S"
    ):
        return False
    return True


def _build_totals(salidas: pd.DataFrame) -> dict:
    """Construye las líneas de servicio facturables (AddServicio + FilterServicio de
    logica.txt) + totales globales para el resumen.

    1. Agrega por (negocio[área], negocio_facturador, tipo_trabajo, tipo_despacho).
    2. Desdinamiza las 3 medidas (PICKING PALLETS/CAJAS/UNIDADES) en líneas.
    3. Asigna `servicio` y aplica el filtro. `valor` = cantidad de esa línea.
    Devuelve {servicios: [...], total_unidades, total_estibas, total_cajas}.
    """
    grouped = (
        salidas.groupby(
            ["area", "negocio", "tipo_trabajo", "tipo_despacho"], dropna=False
        )
        .agg(
            unidades=("cantidad", "sum"),
            estibas=("estibas", "sum"),
            cajas=("cajas", "sum"),
        )
        .reset_index()
    )

    servicios = []
    for _, r in grouped.iterrows():
        area_label = config.AREA_DEFAULT.get(
            str(r["area"]), str(r["area"]) if pd.notna(r["area"]) else ""
        )
        nf = str(r["negocio"]) if pd.notna(r["negocio"]) else ""
        tt = str(r["tipo_trabajo"]) if pd.notna(r["tipo_trabajo"]) else ""
        td = str(r["tipo_despacho"]) if pd.notna(r["tipo_despacho"]) else ""
        unidades = _safe_num(r["unidades"])

        for atributo, campo in _PICKING_MEDIDAS:
            valor = _safe_num(r[campo])
            if not valor:  # None o 0: nada que facturar en esa medida
                continue
            servicio = _servicio(nf, tt, td, atributo)
            if not _keep_servicio(nf, servicio):
                continue
            servicios.append(
                {
                    "negocio": area_label,
                    "negocio_facturador": nf,
                    "servicio": servicio,
                    "valor": valor,
                    "unidades": unidades,
                    "proceso_extendido": f"PICKING {area_label}",
                    "macro_proceso": "OUT BOUND",
                    "proceso_abreviado": "OUB",
                    "tabla": "SALIDAS",
                }
            )

    servicios.sort(
        key=lambda s: (
            s["negocio"],
            s["negocio_facturador"],
            s["servicio"],
        )
    )
    return {
        "servicios": servicios,
        "total_unidades": _safe_num(salidas["cantidad"].sum()),
        "total_estibas": _safe_num(salidas["estibas"].sum()),
        "total_cajas": _safe_num(salidas["cajas"].sum()),
    }


def _safe_num(value) -> Optional[int]:
    if pd.isna(value):
        return None
    return int(value)


# ---------------------------------------------------------------------------
# Tarifas (panel de referencia, para pasos siguientes) y rango por defecto
# ---------------------------------------------------------------------------

def load_tarifas_reference() -> list[dict]:
    """Devuelve tarifas como lista de {servicio, valor, um, fecha}. Para uso futuro."""
    try:
        df = io_utils.read_tarifas()
    except FileNotFoundError:
        return []
    records = []
    for rec in df.to_dict(orient="records"):
        rec["valor"] = None if pd.isna(rec.get("valor")) else float(rec["valor"])
        fecha = rec.get("fecha")
        rec["fecha"] = fecha.strftime("%d/%m/%Y") if pd.notna(fecha) else None
        records.append(rec)
    return records


def default_date_range() -> dict:
    """Min/Max de Fecha factura en todos los archivos, para precargar el calendario.
    Tolerante (strict=False): un archivo roto no rompe la precarga del calendario."""
    salidas, _, _ = _load_all_salidas(strict=False)
    if salidas.empty:
        return {"start": None, "end": None}
    fechas = pd.to_datetime(salidas["fecha"], errors="coerce").dropna()
    if fechas.empty:
        return {"start": None, "end": None}
    return {
        "start": fechas.min().strftime(config.USER_DATE_FORMAT),
        "end": fechas.max().strftime(config.USER_DATE_FORMAT),
    }


# ---------------------------------------------------------------------------
# Punto de entrada para prueba rápida desde consola
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import json

    try:
        res = run_all("20/05/2026", "19/06/2026")
    except BlockingError as exc:
        print("PROCESO DETENIDO (error grave de archivo):", exc)
        raise SystemExit(1)

    print("Issues:")
    print(json.dumps(res.issues, indent=2, ensure_ascii=False))

    salidas_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "SALIDAS")
    destr_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "DESTRUCCION")
    ingr_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "INGRESOS")
    ocup_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "OCUPACION")
    trasl_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "TRASLADOS")
    maq_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "MAQUILA")
    expo_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "EXPORTACIONES")
    etiq_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "ETIQUETAS")
    pal_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "PALETIZADO")
    tri_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "TRINCAJE")
    pla_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "PLANTA")
    mat_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "MATERIAL")
    fal_n = sum(1 for s in res.totals["servicios"] if s["tabla"] == "FALABELLA")
    print(
        f"\nFilas en rango: {res.diagnostics.get('rows_in_range')} | "
        f"Estibas: {res.totals.get('total_estibas')} | Cajas: {res.totals.get('total_cajas')}"
    )
    print(
        f"SALIDAS: {salidas_n} líneas | DESTRUCCION: {destr_n} líneas | "
        f"INGRESOS: {ingr_n} líneas | OCUPACION: {ocup_n} líneas | "
        f"TRASLADOS: {trasl_n} líneas | MAQUILA: {maq_n} líneas | "
        f"EXPORTACIONES: {expo_n} líneas | ETIQUETAS: {etiq_n} líneas | "
        f"PALETIZADO: {pal_n} líneas | TRINCAJE: {tri_n} líneas | "
        f"PLANTA: {pla_n} líneas | MATERIAL: {mat_n} líneas | "
        f"FALABELLA: {fal_n} líneas | "
        f"total {len(res.totals['servicios'])}"
    )
    print("\nLíneas de servicio:")
    total_costo = 0.0
    for s in res.totals["servicios"]:
        costo = s.get("costo_total") or 0
        total_costo += costo
        print(
            f"  [{s['tabla']:<13}] {s['negocio']:<12} "
            f"{s['servicio']:<50} valor={s['valor']:>8} "
            f"tarifa={s.get('tarifa')} costo={costo:>14.2f}"
        )
    print(f"\nCOSTO TOTAL: {total_costo:,.2f}")
