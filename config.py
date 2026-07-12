"""Configuración central del Facturador Henkel.

Rutas a las carpetas de datos, mapa de nombres de columnas y constantes.
No depende de nada de la capa web (FastAPI/uvicorn): cualquier módulo puede
importar de aquí. Si las fuentes renombran columnas, se editan aquí nomás.
"""
from __future__ import annotations

import os
from pathlib import Path

# Raíz del proyecto: la carpeta que contiene AUXILIARES/, CONSUMER/, etc.
# Se puede sobreescribir con la variable de entorno FACTURADOR_BASE_DIR.
BASE_DIR = Path(os.environ.get("FACTURADOR_BASE_DIR", Path(__file__).resolve().parent)).resolve()

# Carpetas de datos (viven en la raíz del proyecto).
DIRS = {
    "auxiliares": BASE_DIR / "AUXILIARES",
    "consumer": BASE_DIR / "CONSUMER",
    "profesional": BASE_DIR / "PROFESIONAL",
    "huellas": BASE_DIR / "HUELLAS",
    "maquila": BASE_DIR / "MAQUILA",
    "exportaciones": BASE_DIR / "EXPORTACIONES",
    "otros": BASE_DIR / "OTROS",
}

# Salida de los Excel generados (separada de los datos fuente).
OUTPUT_DIR = BASE_DIR / "FACTURAS_GENERADAS"

# Archivos de lookup fijos.
FILES = {
    "tarifas": DIRS["auxiliares"] / "tarifas.xlsx",
    "idh_especiales": DIRS["auxiliares"] / "idh_especiales.xlsx",
    "huellas": DIRS["huellas"] / "huellas.xlsx",
    "tipo_despacho": DIRS["auxiliares"] / "tipo_despacho.xlsx",
    # Equivalencias de almacenamiento (Paso 4 Ocupación): Tipo -> Conversion liquidacion.
    # Es el único cruce de Ocupación (no usa huellas ni idh).
    "equivalencias": DIRS["auxiliares"] / "equivalencias_almacenamiento.xlsx",
}

# Patrones para hallar los archivos de salidas por área.
# Se busca en mayúsculas/minúsculas porque a veces llegan como .XLSX y otras .xlsx.
SALIDAS_GLOBS = {
    "cons": "salidas_cons*.[xX][lL][sS][xX]",
    "prof": "salidas_prof*.[xX][lL][sS][xX]",
}

# Negocio por defecto según el área de origen del archivo.
AREA_DEFAULT = {
    "cons": "CONSUMER",
    "prof": "PROFESIONAL",
}

# Adicionales: tablas que empiezan por "adicionales" (viven en CONSUMER/ y PROFESIONAL/).
# Tipifican el trabajo de una entrega (p. ej. "EXTRA E.S") cruzando por Delivery/ENTREGA.
ADICIONALES_GLOB = "adicionales*.[xX][lL][sS][xX]"

# Destrucción (Paso 2): archivos que empiezan por "destruccion" en CONSUMER/ y PROFESIONAL/.
# No traen fecha en su contenido: se procesan TODOS los encontrados (sin filtrar por rango).
# negocio = columna "Almacen" en MAYÚSCULAS; valor = conteo de filas por negocio.
DESTRUCCION_GLOB = "destruccion*.[xX][lL][sS][xX]"

# Ingresos (Paso 3): archivos que empiezan por "ingresos" en CONSUMER/ y PROFESIONAL/.
# El área/negocio se deriva del prefijo del nombre (ingresos_cons* -> CONSUMER,
# ingresos_prof* -> PROFESIONAL), igual que con las salidas. El filtro de fecha se aplica
# sobre la columna "Posting Date" usando el rango que eligió el usuario en el calendario
# (no se usa RANGOS_FECHAS ni la fecha del nombre del archivo).
INGRESOS_GLOBS = {
    "cons": "ingresos_cons*.[xX][lL][sS][xX]",
    "prof": "ingresos_prof*.[xX][lL][sS][xX]",
}

# Ocupación (Paso 4): archivos que empiezan por "ocupacion" en CONSUMER/ y PROFESIONAL/.
# El área/negocio se deriva del prefijo del nombre (ocupacion_cons* -> CONSUMER,
# ocupacion_prof* -> PROFESIONAL), igual que con salidas/ingresos. El filtro de fecha se
# aplica sobre la columna "Fecha" (fecha de almacenamiento) con el rango del usuario
# (no se usa RANGOS_FECHAS ni la fecha del nombre del archivo).
OCUPACION_GLOBS = {
    "cons": "ocupacion_cons*.[xX][lL][sS][xX]",
    "prof": "ocupacion_prof*.[xX][lL][sS][xX]",
}

# Traslados (Paso 5): archivos que empiezan por "traslados" en CONSUMER/ y PROFESIONAL/.
# El área/negocio se deriva del prefijo del nombre (traslados_cons* -> CONSUMER,
# traslados_prof* -> PROFESIONAL), igual que con salidas/ingresos/ocupación. **Sin filtro de
# fecha** (decisión del usuario, molde DESTRUCCIÓN): el contenido no trae columna de periodo
# y se procesan TODOS los encontrados; el valor es el sumatorio de SHU/CON por negocio.
TRASLADOS_GLOBS = {
    "cons": "traslados_cons*.[xX][lL][sS][xX]",
    "prof": "traslados_prof*.[xX][lL][sS][xX]",
}

# Maquila (Paso 6): archivos que empiezan por "maquila" en la carpeta MAQUILA/. A
# diferencia de los pasos previos, los ficheros `maquila_cons*` y `maquila_prof*` viven
# **ambos en MAQUILA/** (no en CONSUMER/ y PROFESIONAL/), así que find_maquila_files
# hace glob de los dos patrones sobre la misma carpeta. El área/negocio se deriva del
# prefijo del nombre (maquila_cons* -> CONSUMER, maquila_prof* -> PROFESIONAL), igual
# que con ingresos. El filtro de fecha se aplica sobre "Posting Date" con el rango del
# usuario (no se usa RANGOS_FECHAS ni la fecha del nombre).
MAQUILA_GLOBS = {
    "cons": "maquila_cons*.[xX][lL][sS][xX]",
    "prof": "maquila_prof*.[xX][lL][sS][xX]",
}

# Exportaciones (Paso 7): archivos que empiezan por "exportacion" en la carpeta
# EXPORTACIONES/. Igual que con maquila, los `exportacion_cons*` y `exportacion_prof*`
# viven **ambos en EXPORTACIONES/** (no en CONSUMER/ y PROFESIONAL/), así que
# find_exportacion_files hace glob de los dos patrones sobre la misma carpeta. El
# área/negocio se deriva del prefijo (exportacion_cons* -> CONSUMER, ..._prof* ->
# PROFESIONAL). El filtro de fecha se aplica sobre "Fecha factura" con el rango del
# usuario (no se usa RANGOS_FECHAS ni la fecha del nombre). `paletizado*`/`trincaje*`
# (que también viven en EXPORTACIONES/) NO se procesan: logica.txt sólo admite ficheros
# que empiecen por "exportacion".
EXPORTACION_GLOBS = {
    "cons": "exportacion_cons*.[xX][lL][sS][xX]",
    "prof": "exportacion_prof*.[xX][lL][sS][xX]",
}

# Etiquetas (Paso 8): ficheros que empiezan por "etiquetas" en la carpeta MAQUILA/ (ahí
# vive `etiquetas_<fecha>.xlsx`, SIN separación cons/prof, a diferencia de maquila/expo).
# El área/negocio no se deriva del nombre: el PQ fija negocio = "CONSUMER" para todas las
# filas (ver ETIQUETAS_NEGOCIO). Sin filtro de fecha: el PQ no aplica RANGOS_FECHAS, solo
# agrupa por la fecha del NOMBRE del archivo (que el bot descarta y reemplaza por `periodo`).
ETIQUETAS_GLOB = "etiquetas*.[xX][lL][sS][xX]"

# Paletizado (Paso 9): ficheros que empiezan por "paletizado" en la carpeta EXPORTACIONES/
# (ahí vive `paletizado_<fecha>.xlsx`, junto a `exportacion_*`/`trincaje_*`; la query de
# exportación del Paso 7 NO lo procesa, logica.txt lo admite en su propia query). A
# diferencia del resto de pasos, el `negocio` se deriva de la columna AREA (ver
# PALETIZADO_AREA_MAP), no del prefijo del nombre ni del idh. Sin filtro de fecha: el PQ
# no aplica RANGOS_FECHAS, solo agrupa por la fecha del NOMBRE del archivo (que el bot
# descarta y reemplaza por `periodo`).
PALETIZADO_GLOB = "paletizado*.[xX][lL][sS][xX]"

# Trincaje (Paso 10): ficheros que empiezan por "trincaje" en la carpeta EXPORTACIONES/
# (ahí vive `trincaje_<fecha>.xlsx`, junto a `exportacion_*`/`paletizado_*`; la query de
# exportación del Paso 7 NO lo procesa, logica.txt lo admite en su propia query). Igual de
# simple que ETIQUETAS/TRASLADOS: sin huellas, sin idh, sin cruce y sin filtro de fecha. El
# PQ sólo filtra `despacho <> null`, agrupa por la fecha del NOMBRE y cuenta filas `× 2`
# -> valor (negocio = "CONSUMER" fijo; servicio = proceso_extendido = "DOBLE TRINCAJE").
TRINCAJE_GLOB = "trincaje*.[xX][lL][sS][xX]"

# Planta (Paso 11): ficheros que empiezan por "planta" en las carpetas CONSUMER/ y/o
# PROFESIONAL/ (ahí vive `planta_<fecha>.xlsx`; en la data actual está en PROFESIONAL/,
# pero un SOLO fichero cubre ambos negocios: el `negocio` viene del NOMBRE de la columna,
# no de la carpeta — ver PLANTA_COLS). Tan simple como TRINCAJE/PALETIZADO: sin huellas,
# sin idh, sin equivalencias y sin filtro de fecha (el PQ no aplica RANGOS_FECHAS: elimina
# las cols `fecha`/`semana` del archivo y agrupa por la fecha del NOMBRE, que el bot
# descarta y reemplaza por `periodo`). Se unpivotean las dos cols de estibas y se suma.
PLANTA_GLOB = "planta*.[xX][lL][sS][xX]"

# Material (Paso 12): ficheros que empiezan por "ocupacionMaterial" en la carpeta OTROS/
# (ahí vive `ocupacionMaterial_<fecha>.xlsx`). Gemelo SIMPLIFICADO de Ocupación (Paso 4):
# misma idea (promedio diario redondeado hacia arriba, ceil) pero con UNA sola medida
# (`cant_bodega_8`), servicio/negocio/nf fijos y sin join EQUIVALENCIAS. El PQ filtra por
# la ventana de RANGOS_FECHAS (joined por la fecha del NOMBRE); el bot la reemplaza por el
# rango libre del calendario del usuario sobre la columna `fecha` (igual que Ocupación).
MATERIAL_GLOB = "ocupacionMaterial*.[xX][lL][sS][xX]"

# Defaults cuando un registro NO hace match en el cruce correspondiente.
TIPO_TRABAJO_DEFAULT = "NORMAL"     # si el Delivery no está en adicionales
TIPO_DESPACHO_DEFAULT = "ESTANDAR"  # si el cliente no está en tipo_despacho

# Prefijo de los archivos de bloqueo de Excel (abiertos). Siempre se excluyen.
LOCK_PREFIX = "~$"

# ---- Mapa de columnas de las fuentes (nombres exactos tal como vienen) ----

# salidas_cons / salidas_prof (hoja con la cabecera MATERIAL, fila 1).
# "Material" aparece dos veces en el archivo; se usa la 1ª ocurrencia (occurrence=0).
SALIDAS_COLS = {
    "delivery": "Delivery",
    "cliente": "Nombre completo con tratamiento y título",
    "material": "Material",
    "cantidad": "Ctd Ent.(UMV)",
    "fecha": "Fecha factura",
}

# Posiciones conocidas (índices 0-based) de las columnas anteriores en la export
# de SAP. Camino rápido: se leen solo estas 5 columnas en un único parseo y se
# verifican por nombre. Si el layout cambiara, se cae al fallback robusto (lectura
# completa + emparejamiento por nombre). Orden: delivery, cliente, material, cantidad, fecha.
SALIDAS_FAST_POSITIONS = [0, 7, 16, 19, 63]

# huellas.xlsx: los encabezados empiezan en la fila 2 (header=1 en pandas).
HUELLAS_HEADER_ROW = 1  # índice 0-based -> fila física 2
HUELLAS_COLS = {
    "producto": "Producto",
    "pallet": "Unidades por pallet",
    "caja": "Unidades por caja",
    "fupd": "Fecha actualizacion",
}

# idh_especiales.xlsx
IDH_COLS = {
    "material": "Material",
    "negocio": "Negocio",
}

# adicionales_cons / adicionales_prof (hoja con cabecera ENTREGA, TIPO).
ADICIONALES_COLS = {
    "entrega": "ENTREGA",
    "tipo": "TIPO",
}

# tipo_despacho.xlsx (hoja con cabecera CEDI, TIPO).
TIPO_DESPACHO_COLS = {
    "cedi": "CEDI",
    "tipo": "TIPO",
}

# destruccion_* (hoja con cabecera Almacen). Solo se usa Almacen para derivar `negocio`.
DESTRUCCION_COLS = {
    "almacen": "Almacen",
}

# ingresos_cons / ingresos_prof (hoja Sheet1, export de SAP). El área/negocio viene del
# prefijo del nombre del archivo (no de una columna). El filtro de fecha va sobre Posting Date.
# "documento_cruce" = referencia si no es nula, si no documento (Texto cab.documento); las
# filas sin ninguno de los dos se descartan (fidelidad a logica.txt).
INGRESOS_COLS = {
    "posting_date": "Posting Date",
    "cantidad": "Cantidad",
    "material": "Material",
    "documento": "Texto cab.documento",
    "referencia": "Referencia",
}

# ocupacion_cons / ocupacion_prof (fecha de almacenamiento por fila). El área viene del
# prefijo del nombre del archivo. "Ocupación" y "% Ocupación" normalizan ambas a OCUPACION:
# se usa la 1ª ocurrencia (la de valor) con find_column(..., occurrence=0), igual que el
# "Material" duplicado de salidas. El parseo de "-" / "" -> 0 (fnNumero) lo hace la pipeline.
OCUPACION_COLS = {
    "fecha": "Fecha",
    "almacen": "Almacen",
    "tipo": "Tipo",
    "ocupacion": "Ocupación",
}

# traslados_cons / traslados_prof (export de SAP de entregas por Ship-To). El área/negocio
# viene del prefijo del nombre del archivo. Solo se usan estas 3 columnas: `delivery` para el
# filtro "Delivery <> null" de logica.txt, y `shu`/`con` para los 2 servicios (cajas/unidades).
TRASLADOS_COLS = {
    "delivery": "Delivery",
    "shu": "SHU",
    "con": "CON",
}

# maquila_cons / maquila_prof (export de SAP de movimientos de subcontratación). El
# área/negocio viene del prefijo del nombre del archivo. Solo se usan estas 3 columnas:
# `posting_date` (filtro por rango), `material` (cruces huellas + idh) y `cantidad`
# (base de las 6 medidas, mismas que ingresos). `documento`/`referencia` no aplican.
MAQUILA_COLS = {
    "posting_date": "Posting Date",
    "cantidad": "Cantidad",
    "material": "Material",
}

# exportacion_cons / exportacion_prof (export de SAP de entregas de exportación). El
# área/negocio viene del prefijo del nombre. `delivery` sólo para el filtro "<> null" de
# logica.txt; `fecha` (Fecha factura) para el filtro por rango del usuario; `material`
# para el cruce huellas (pallet/caja); `cantidad` base de las 3 medidas; `canal`
# (Canal distribución, p. ej. "EX"/"IC") forma parte del nombre del servicio
# (`servicio = base & " " & canal`). `Texto breve de material` no se usa en la agregación.
EXPORTACION_COLS = {
    "delivery": "Delivery",
    "fecha": "Fecha factura",
    "material": "Material",
    "cantidad": "Ctd Ent.(UMV)",
    "canal": "Canal distribución",
}

# etiquetas_<fecha> (hoja con cabecera fecha/entrega/cajas/tipo). Sólo se usan `cajas`
# (la medida que se suma -> valor) y `tipo` (que mapea al nombre del servicio). `fecha`
# y `entrega` existen en el fuente pero el PQ las descarta (la `fecha` final viene del
# nombre del archivo), así que no se leen.
ETIQUETAS_COLS = {
    "cajas": "cajas",
    "tipo": "tipo",
}

# Mapeo tipo -> servicio (logica.txt: ReplaceValue REGULARES/REEMPAQUE -> ETIQUETAS ...).
# Otros valores de `tipo` pasan tal cual (el PQ sólo reemplaza esos dos).
ETIQUETAS_SERVICIO_MAP = {
    "REGULARES": "ETIQUETAS REGULARES",
    "REEMPAQUE": "ETIQUETAS REEMPAQUE",
}

# Negocio fijo de etiquetas (el PQ hardcodea "CONSUMER" para todas las filas).
ETIQUETAS_NEGOCIO = "CONSUMER"

# paletizado_<fecha> (hoja con AREA/DESPACHO/TOTAL/CANAL...). Sólo se usan `area` (->
# negocio vía PALETIZADO_AREA_MAP), `despacho` (filtro "<> null" de logica.txt), `total`
# (la medida que se suma -> valor) y `canal` (forma parte del nombre del servicio). Las
# demás cols (N-DE EXPO, ORDEN SAP, CONTE/OR, N-CONTEOR, TIPO, ESTATUS) no se usan.
PALETIZADO_COLS = {
    "area": "AREA",
    "despacho": "DESPACHO",
    "total": "TOTAL",
    "canal": "CANAL",
}

# Mapeo AREA -> negocio (logica.txt: ReplaceValue HENKEL.PF -> PROFESIONAL,
# HENKEL.RT -> CONSUMER). Otros valores de AREA pasan tal cual (el PQ sólo reemplaza
# esos dos); las filas con AREA nula se descartan (filtro final negocio <> null).
PALETIZADO_AREA_MAP = {
    "HENKEL.PF": "PROFESIONAL",
    "HENKEL.RT": "CONSUMER",
}

# trincaje_<fecha> (hoja con despacho/exp/destino/contenedor/n-cont). El PQ lee las 5
# cols pero tras el filtro sólo conserva el NOMBRE del archivo (la `fecha` final viene
# del nombre): sólo se usa `despacho` (filtro "<> null" de logica.txt). Las demás cols
# (exp/destino/contenedor/n-cont) no se usan en la agregación.
TRINCAJE_COLS = {
    "despacho": "despacho",
}

# Negocio fijo y servicio de trincaje (el PQ hardcodea negocio = "CONSUMER" y
# proceso_extendido/servicio = "DOBLE TRINCAJE" para todas las filas).
TRINCAJE_NEGOCIO = "CONSUMER"
TRINCAJE_SERVICIO = "DOBLE TRINCAJE"
# Factor del conteo: logica.txt hace `valor = Table.RowCount(_) * 2` por cada fecha.
TRINCAJE_FACTOR = 2

# planta_<fecha> (hoja con semana/fecha/estibas_consumer/estibas_profesional). El PQ elimina
# `fecha` y `semana` (la `fecha` final viene del NOMBRE del archivo) y unpivotea las dos cols
# de estibas: `estibas_consumer -> negocio CONSUMER`, `estibas_profesional -> negocio
# PROFESIONAL`. Las llaves de este dict son internas (consumer/profesional) y mapean al
# `negocio` final vía PLANTA_NEGOCIOS.
PLANTA_COLS = {
    "consumer": "estibas_consumer",
    "profesional": "estibas_profesional",
}
PLANTA_NEGOCIOS = {
    "consumer": "CONSUMER",
    "profesional": "PROFESIONAL",
}
# Negocio fijo del servicio (el PQ hardcodea "TRASLADO PALLETS PLANTA - CEDI" para todas las
# filas); `negocio_facturador = negocio` (el PQ duplica `negocio -> negocio_facturador`).
PLANTA_SERVICIO = "TRASLADO PALLETS PLANTA - CEDI"

# ocupacionMaterial_<fecha> (hoja con fecha/cant_bodega_8). Sólo se leen `fecha` (día de
# almacenamiento, sobre la que el bot aplica el rango del usuario) y `cant_bodega_8` (la
# única medida -> valor = ceil del promedio en el rango).
MATERIAL_COLS = {
    "fecha": "fecha",
    "valor": "cant_bodega_8",
}
# Negocio/negocio_facturador/servicio FIJOS del PQ. OJO: `negocio_facturador` es
# "MATERIAL EMPAQUE" (SIN "DE"), distinto a `negocio` "MATERIAL DE EMPAQUE" (AddNegocio vs
# AddNegocioFacturador de logica.txt).
MATERIAL_NEGOCIO = "MATERIAL DE EMPAQUE"
MATERIAL_NEGOCIO_FACTURADOR = "MATERIAL EMPAQUE"
MATERIAL_SERVICIO = "ALMACENAMIENTO PALLET BODEGA 8 ME GENERAL"

# Falabella (Paso 13): ficheros falabella_cons* en CONSUMER/ y falabella_prof* en
# PROFESIONAL/ (molde TRASLADOS/OCUPACION: el área/negocio viene del prefijo del nombre
# + carpeta). **Sin filtro de fecha** (el PQ no aplica RANGOS_FECHAS, igual que
# TRASLADOS/ETIQUETAS/TRINCAJE/PLANTA: agrupa por la fecha del NOMBRE, que el bot
# descarta y reemplaza por `periodo`). Sólo se usa `Entrega` (filtro "<> null" de
# logica.txt); el conteo de filas -> valor. `MES` existe en el fuente pero el PQ la
# descarta (no llega al output).
FALABELLA_GLOBS = {
    "cons": "falabella_cons*.[xX][lL][sS][xX]",
    "prof": "falabella_prof*.[xX][lL][sS][xX]",
}
FALABELLA_COLS = {"entrega": "Entrega"}
FALABELLA_SERVICIO = "ETIQUETADO (FALABELLA)"

# equivalencias_almacenamiento.xlsx (join de Ocupación). `archivo` (= Tipo de ubicación,
# p. ej. "PALLET (MEDIO)") es la llave; `conversion` es el valor que forma el servicio
# ("MEDIO PALLET") y se conserva tal cual (no se normaliza agresivamente: lleva espacios).
EQUIVALENCIAS_COLS = {
    "archivo": "Archivo Alarmasblu",
    "conversion": "Conversion liquidacion",
}

# Columnas de tarifas (panel de referencia).
TARIFAS_COLS = {
    "servicio": "servicio",
    "valor": "valor",
    "um": "UM",
    "fecha": "fecha",
}

# Formato de fecha que usa el usuario en el calendario (dd/mm/yyyy).
USER_DATE_FORMAT = "%d/%m/%Y"
