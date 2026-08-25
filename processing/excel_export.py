"""Exportación del resultado de facturación a Excel.

La hoja 'Servicios' contiene una fila por línea facturable y la hoja 'Resumen'
presenta consolidaciones por macroproceso, negocio facturador y negocio.
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
from openpyxl.styles import Alignment, Font, PatternFill

from processing import io_utils

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
    "minima",
]

_DISPLAY_NEGOCIO = {
    "CONSUMER": "CONSUMER",
    "PROFESIONAL": "PROFESSIONAL",
}


def _macro_resumen(value) -> str:
    macro = "" if pd.isna(value) else str(value).strip().upper()
    return "ALMACENAMIENTO" if macro == "ALMCENAMIENTO" else macro


def _nombre_negocio(value) -> str:
    nombre = "" if pd.isna(value) else str(value).strip().upper()
    return _DISPLAY_NEGOCIO.get(nombre, nombre)


def _config_minimos() -> tuple[dict[str, float], dict[str, str]]:
    """Devuelve mínimos por macro y el macro aplicable a cada servicio elegible."""
    try:
        tarifas = io_utils.read_tarifas()
    except (FileNotFoundError, KeyError, ValueError):
        return {}, {}
    aplica = tarifas["minima"].map(io_utils.normalize).eq(
        io_utils.normalize("Aplica minima")
    )
    tarifas = tarifas.loc[aplica].copy()
    tarifas["_macro"] = tarifas["macro_proceso"].map(_macro_resumen)
    tarifas = tarifas.loc[tarifas["_macro"].ne("") & tarifas["_macro"].ne("OTROS")]
    minimos = (
        tarifas.groupby("_macro")["minima_valor_subproceso"].sum().astype(float).to_dict()
    )
    tarifas["_servicio_key"] = tarifas["servicio"].map(io_utils.normalize)
    servicio_macro = (
        tarifas.drop_duplicates("_servicio_key", keep="last")
        .set_index("_servicio_key")["_macro"]
        .to_dict()
    )
    return minimos, servicio_macro


def _build_summary_tables(servicios: list[dict]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Construye los tres resúmenes solicitados a partir del resultado final."""
    df = pd.DataFrame(servicios)
    if df.empty:
        df = pd.DataFrame(columns=SERVICIO_COLUMNS)
    for col in ("valor", "costo_total"):
        df[col] = pd.to_numeric(df.get(col), errors="coerce").fillna(0.0)
    df["_macro"] = df.get("macro_proceso", pd.Series(index=df.index, dtype=object)).map(_macro_resumen)
    df["_tabla"] = df.get("tabla", pd.Series(index=df.index, dtype=object)).fillna("").astype(str).str.upper()
    df["_servicio_key"] = df.get("servicio", pd.Series(index=df.index, dtype=object)).map(io_utils.normalize)
    minimos, servicio_macro = _config_minimos()
    df["_macro_minima"] = df["_servicio_key"].map(servicio_macro)
    df["_aplica_minima"] = df["_macro_minima"].notna()

    # 1) Usa exactamente la misma base del cálculo de mínimos: sólo servicios marcados
    #    Aplica minima. Así TOTAL FACTURADO + AJUSTE se puede comparar con TARIFA MINIMA.
    normales = df.loc[df["_aplica_minima"] & df["_tabla"].ne("AJUSTE")].copy()
    normales["_macro"] = normales["_macro_minima"]
    ajustes = df.loc[df["_tabla"].eq("AJUSTE")]
    macros = sorted(set(normales["_macro"].dropna()) | set(minimos) | set(ajustes["_macro"].dropna()))
    macros = [m for m in macros if m and m != "OTROS"]
    macro_rows = []
    for macro in macros:
        base = normales.loc[normales["_macro"].eq(macro)]
        macro_rows.append(
            {
                "CONCEPTO": macro,
                "CANTIDAD FACTURADA": float(base["valor"].sum()),
                "TOTAL FACTURADO": float(base["costo_total"].sum()),
                "TARIFA MINIMA": float(minimos.get(macro, 0)),
                "AJUSTE": float(ajustes.loc[ajustes["_macro"].eq(macro), "costo_total"].sum()),
            }
        )
    macro_rows.append(
        {
            "CONCEPTO": "Total general",
            "CANTIDAD FACTURADA": sum(r["CANTIDAD FACTURADA"] for r in macro_rows),
            "TOTAL FACTURADO": sum(r["TOTAL FACTURADO"] for r in macro_rows),
            "TARIFA MINIMA": sum(r["TARIFA MINIMA"] for r in macro_rows),
            "AJUSTE": sum(r["AJUSTE"] for r in macro_rows),
        }
    )
    por_macro = pd.DataFrame(macro_rows)

    # 2) Total final por negocio_facturador; los ajustes forman una categoría propia.
    nf_totals: dict[str, float] = {}
    for row in df.to_dict("records"):
        concepto = "AJUSTE" if row.get("_tabla") == "AJUSTE" else _nombre_negocio(row.get("negocio_facturador"))
        # En este cuadro LAUNDRY forma parte consolidada de CONSUMER.
        if concepto == "LAUNDRY":
            concepto = "CONSUMER"
        if concepto:
            nf_totals[concepto] = nf_totals.get(concepto, 0.0) + float(row.get("costo_total", 0))
    preferred_nf = ["MATERIAL DE EMPAQUE", "NATTURA", "PROFESSIONAL", "CONSUMER"]
    nf_names = list(preferred_nf)
    nf_names += sorted(n for n in nf_totals if n not in preferred_nf and n != "AJUSTE")
    if "AJUSTE" in nf_totals:
        nf_names.append("AJUSTE")
    fechas_periodo = pd.to_datetime(df.get("periodo"), errors="coerce", dayfirst=True)
    anio = int(fechas_periodo.max().year) if fechas_periodo.notna().any() else pd.Timestamp.now().year
    total_col = f"TOTAL TARIFA {anio}"
    nf_rows = [{"PROCESO": n, total_col: nf_totals.get(n, 0.0)} for n in nf_names]
    nf_rows.append({"PROCESO": "TOTAL", total_col: sum(nf_totals.values())})
    por_nf = pd.DataFrame(nf_rows)

    # 3) Por negocio: operación normal, ajuste y conceptos OTROS.
    negocio_totals: dict[str, dict[str, float]] = {}
    for row in df.to_dict("records"):
        negocio = _nombre_negocio(row.get("negocio"))
        if not negocio:
            continue
        bucket = negocio_totals.setdefault(
            negocio, {"ALMACENAMIENTO": 0.0, "AJUSTE A MINIMA": 0.0, "OTROS CONCEPTOS": 0.0}
        )
        costo = float(row.get("costo_total", 0))
        if row.get("_tabla") == "AJUSTE":
            bucket["AJUSTE A MINIMA"] += costo
        elif row.get("_aplica_minima"):
            bucket["ALMACENAMIENTO"] += costo
        else:
            bucket["OTROS CONCEPTOS"] += costo
    preferred_neg = ["CONSUMER", "PROFESSIONAL", "MATERIAL DE EMPAQUE"]
    neg_names = list(preferred_neg)
    neg_names += sorted(n for n in negocio_totals if n not in preferred_neg)
    neg_rows = []
    for nombre in neg_names:
        vals = negocio_totals.get(
            nombre, {"ALMACENAMIENTO": 0.0, "AJUSTE A MINIMA": 0.0, "OTROS CONCEPTOS": 0.0}
        )
        neg_rows.append({"NEGOCIO": nombre, **vals, "TOTAL": sum(vals.values())})
    total_vals = {
        col: sum(r[col] for r in neg_rows)
        for col in ("ALMACENAMIENTO", "AJUSTE A MINIMA", "OTROS CONCEPTOS", "TOTAL")
    }
    neg_rows.append({"NEGOCIO": "TOTALES", **total_vals})
    por_negocio = pd.DataFrame(neg_rows)
    return por_macro, por_nf, por_negocio


def _format_summary_sheet(sheet, table_ranges: list[tuple[int, int, int]]) -> None:
    """Aplica formato legible a títulos, encabezados, totales y valores monetarios."""
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    total_fill = PatternFill("solid", fgColor="D9EAF7")
    currency_fmt = '$ #,##0;[Red]-$ #,##0;$ -'
    quantity_fmt = '#,##0.####'
    for title_row, header_row, end_row in table_ranges:
        sheet.cell(title_row, 1).font = Font(bold=True, size=13, color="1F4E78")
        for cell in sheet[header_row]:
            if cell.value is not None:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
        for cell in sheet[end_row]:
            if cell.value is not None:
                cell.fill = total_fill
                cell.font = Font(bold=True)
    for row in sheet.iter_rows():
        for cell in row:
            if cell.column >= 2 and isinstance(cell.value, (int, float)):
                cell.number_format = quantity_fmt if cell.column == 2 and row[0].value in {
                    "ALMACENAMIENTO", "IN BOUND", "OUT BOUND", "Total general"
                } else currency_fmt
    sheet.freeze_panes = "A3"


def export_services(totals: dict, out_path: Path) -> Path:
    """Escribe las hojas Servicios y Resumen, y devuelve la ruta usada.

    `totals["servicios"]` ya viene combinado (SALIDAS + DESTRUCCION). Las filas
    DESTRUCCION llevan tipo_trabajo/tipo_despacho vacíos.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)

    servicios = totals.get("servicios", [])
    hoja = pd.DataFrame(servicios, columns=SERVICIO_COLUMNS)
    por_macro, por_nf, por_negocio = _build_summary_tables(servicios)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        hoja.to_excel(writer, sheet_name="Servicios", index=False)
        resumen = "Resumen"
        negocio_title_row = 1
        negocio_header_row = 2
        por_negocio.to_excel(writer, sheet_name=resumen, index=False, startrow=negocio_header_row - 1)
        ws = writer.book[resumen]
        ws.cell(negocio_title_row, 1, "RESUMEN POR NEGOCIO")

        macro_title_row = negocio_header_row + len(por_negocio) + 3
        macro_header_row = macro_title_row + 1
        por_macro.to_excel(writer, sheet_name=resumen, index=False, startrow=macro_header_row - 1)
        ws.cell(macro_title_row, 1, "RESUMEN POR MACROPROCESO")

        nf_title_row = macro_header_row + len(por_macro) + 3
        nf_header_row = nf_title_row + 1
        por_nf.to_excel(writer, sheet_name=resumen, index=False, startrow=nf_header_row - 1)
        ws.cell(nf_title_row, 1, "RESUMEN POR NEGOCIO FACTURADOR")
        _format_summary_sheet(
            ws,
            [
                (negocio_title_row, negocio_header_row, negocio_header_row + len(por_negocio)),
                (macro_title_row, macro_header_row, macro_header_row + len(por_macro)),
                (nf_title_row, nf_header_row, nf_header_row + len(por_nf)),
            ],
        )
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
