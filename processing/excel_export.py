"""Exportación del resultado (SALIDAS + DESTRUCCIÓN) a Excel.

Una sola hoja 'Servicios' con una fila por línea de servicio facturable:
- SALIDAS: salida de AddServicio + FilterServicio de logica.txt.
- DESTRUCCION: conteo de filas por negocio (Paso 2).

Cada grupo SALIDAS se desdinamiza en hasta 3 servicios (pallets, cajas, unidades);
`valor` es la cantidad de esa línea; `costo_total` = `valor × tarifa` (cruce con
tarifas.xlsx por `servicio`). Las líneas sin tarifa o con valor 0 no llegan a esta hoja.
La columna `tabla` distingue SALIDAS de DESTRUCCION.

Motor: openpyxl (ya instalado). Sin dependencias web.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

# Orden estable de columnas para la hoja Servicios. `periodo` va primero (primer día
# del último mes del rango, igual en todos los registros). Sin tipo_trabajo/tipo_despacho
# (no interesan en el Excel); `servicio` es un nivel de agrupación.
# `um`/`tarifa`/`costo_total` vienen del cruce con tarifas.xlsx (Paso tarifa→costo):
# costo_total = valor × tarifa. Las líneas sin tarifa (ni las de valor 0) llegan hasta aquí.
SERVICIO_COLUMNS = [
    "periodo",
    "negocio",
    "negocio_facturador",
    "servicio",
    "valor",
    "unidades",
    "um",
    "tarifa",
    "costo_total",
    "proceso_extendido",
    "macro_proceso",
    "proceso_abreviado",
    "tabla",
]


def export_services(totals: dict, out_path: Path) -> Path:
    """Escribe el Excel de salida (hoja Servicios) y devuelve la ruta usada.

    `totals["servicios"]` ya viene combinado (SALIDAS + DESTRUCCION). Las filas
    DESTRUCCION llevan tipo_trabajo/tipo_despacho vacíos.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)

    servicios = totals.get("servicios", [])
    hoja = pd.DataFrame(servicios, columns=SERVICIO_COLUMNS)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        hoja.to_excel(writer, sheet_name="Servicios", index=False)
        # Ajustar ancho de columnas para que sea legible.
        for sheet in writer.book.worksheets:
            for column_cells in sheet.columns:
                length = max(
                    (len(str(c.value)) for c in column_cells if c.value is not None),
                    default=10,
                )
                sheet.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 55)

    return out_path


def build_export_filename(start: str, end: str) -> str:
    """Nombre estable por periodo: facturacion_<ddmmyyyy>_<ddmmyyyy>.xlsx.

    Sin timestamp: así el Excel pre-construido se reutiliza entre Generar y Descargar
    y se sobrescribe si se regenera el mismo periodo.
    """
    s = start.replace("/", "")
    e = end.replace("/", "")
    return f"facturacion_{s}_{e}.xlsx"
