"""Rutas JSON del Facturador Henkel (FastAPI).

Contrato uniforme:
  - Éxito:  {"ok": true,  "data": ...}
  - Error:  HTTP 4xx/5xx con {"ok": false, "error": "...", "detail": "..."}

Endpoints:
  GET  /api/tarifas              -> tarifas (referencia, para pasos siguientes)
  GET  /api/daterange/default    -> min/max de Fecha factura para precargar el calendario
  POST /api/run                  -> arranca el proceso en segundo plano (SALIDAS + DESTRUCCION +
                                    INGRESOS + OCUPACION + TRASLADOS + MAQUILA + EXPORTACIONES)
  GET  /api/progress             -> avance (% y etapa) + problemas detectados (en vivo)
  GET  /api/export               -> descarga el Excel (hoja Servicios: los 7 pasos combinados)

El proceso corre en un hilo daemon; el cliente sondea /api/progress. Al terminar
sin errores graves, el Excel se pre-construye en segundo plano y /api/export lo
sirve al instante (o lo construye en el momento si aún no existía).
"""
from __future__ import annotations

import threading
from pathlib import Path

from fastapi import APIRouter
from fastapi.responses import FileResponse, JSONResponse

import config
from processing import excel_export, pipeline

router = APIRouter()


def _ok(data):
    return {"ok": True, "data": data}


def _err(message: str, detail: str = "", status: int = 400):
    return JSONResponse(
        status_code=status,
        content={"ok": False, "error": message, "detail": detail},
    )


def _excel_out_path(start: str, end: str) -> Path:
    return config.OUTPUT_DIR / excel_export.build_export_filename(start, end)


def _build_excel(result, start: str, end: str) -> Path:
    """Construye el Excel del resultado y marca result.excel_path. Devuelve la ruta."""
    out_path = _excel_out_path(start, end)
    excel_export.export_services(result.totals, out_path)
    result.excel_path = str(out_path)
    return out_path


# ---------------------------------------------------------------------------
# Tarifas (panel de referencia, para pasos siguientes)
# ---------------------------------------------------------------------------

@router.get("/tarifas")
def get_tarifas():
    return _ok(pipeline.load_tarifas_reference())


# ---------------------------------------------------------------------------
# Rango por defecto del calendario
# ---------------------------------------------------------------------------

@router.get("/daterange/default")
def get_default_daterange():
    return _ok(pipeline.default_date_range())


# ---------------------------------------------------------------------------
# Arranque del proceso (segundo plano) + avance
# ---------------------------------------------------------------------------

def _run_job(start: str, end: str):
    """Hilo daemon: ejecuta los 7 pasos (SALIDAS + DESTRUCCION + INGRESOS + OCUPACION +
    TRASLADOS + MAQUILA + EXPORTACIONES) con avance en vivo."""
    pipeline.start_run(start, end)
    try:
        result = pipeline.run_all(
            start,
            end,
            progress=pipeline.set_progress,
            on_issue=pipeline.add_issue,
        )
    except pipeline.BlockingError as exc:
        pipeline.invalidate_cache(start, end)
        pipeline.finish_run(blocked=True, error=str(exc))
        return
    except Exception as exc:  # salvaguarda: cualquier otro error se reporta claro.
        pipeline.invalidate_cache(start, end)
        pipeline.finish_run(blocked=True, error=f"Error inesperado: {exc}")
        return

    # Éxito: pre-construir el Excel en segundo plano (Descargar será instantáneo).
    try:
        _build_excel(result, start, end)
    except Exception:
        # Si falla, se dejará excel_path=None y se construirá síncrono al descargar.
        pass

    pipeline.finish_run(has_result=bool(result.totals.get("servicios")))


@router.post("/run")
def post_run(payload: dict):
    start = (payload or {}).get("start", "").strip()
    end = (payload or {}).get("end", "").strip()
    if not start or not end:
        return _err("Faltan las fechas.", "Se requieren 'start' y 'end' en formato dd/mm/yyyy.")

    # Validar el rango antes de arrancar (feedback inmediato).
    try:
        pipeline.parse_date_range(start, end)
    except ValueError as exc:
        return _err("Rango de fechas inválido.", str(exc), status=400)

    if pipeline.is_running():
        return _err(
            "Ya hay un proceso en curso.",
            "Espera a que termine el proceso actual antes de generar de nuevo.",
            status=409,
        )

    threading.Thread(target=_run_job, args=(start, end), daemon=True).start()
    return _ok({"started": True, "run_key": f"{start}|{end}"})


@router.get("/progress")
def get_progress():
    """Avance del proceso: {stage, percent, done, blocked, error, issues, has_result}."""
    return _ok(pipeline.get_progress())


# ---------------------------------------------------------------------------
# Exportación a Excel
# ---------------------------------------------------------------------------

@router.get("/export")
def get_export(start: str, end: str):
    start = (start or "").strip()
    end = (end or "").strip()
    if not start or not end:
        return _err("Faltan las fechas para exportar.", "Use ?start=dd/mm/yyyy&end=dd/mm/yyyy.")

    try:
        result = pipeline.run_all_cached(start, end)
    except pipeline.BlockingError as exc:
        return _err(
            "No se pudo exportar: hay un error grave de archivo pendiente.",
            str(exc),
            status=400,
        )
    except ValueError as exc:
        return _err("No se pudo exportar.", str(exc), status=400)
    except Exception as exc:
        return _err("Error inesperado al exportar.", str(exc), status=500)

    if not result.totals.get("servicios"):
        return _err(
            "No hay datos en el rango seleccionado para exportar.",
            "Genere primero un rango con resultados.",
            status=400,
        )

    # Servir el Excel pre-construido si existe; si no, construirlo ahora.
    if result.excel_path and Path(result.excel_path).exists():
        out_path = Path(result.excel_path)
    else:
        try:
            out_path = _build_excel(result, start, end)
        except Exception as exc:
            return _err("No se pudo escribir el Excel.", str(exc), status=500)

    return FileResponse(
        path=str(out_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=out_path.name,
    )
