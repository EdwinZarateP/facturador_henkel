# Facturador Henkel

Herramienta **local** (se ejecuta en el propio equipo, los datos no salen de la máquina)
que genera la facturación de servicios de almacén a partir de los Excel que llegan en
las carpetas `CONSUMER/`, `PROFESIONAL/`, `AUXILIARES/` y `HUELLAS/`. Tiene una
interfaz web sencilla: se elige un **rango de fechas con calendario**, se pulsa
**Generar** (se ve una **barra de avance con cronómetro**; si una fuente tiene un error
grave el proceso se **detiene** y lo indica para que lo corrijas) y luego
**Descargar Excel**.

> **Estado.** Construcción incremental frente a `logica.txt` (Power Query de referencia).
> - **Hecho:** lectura de salidas + filtro por fecha + cruces (`idh`, `huellas`,
>   `adicionales`, `tipo_despacho`) + `AddServicio`/`FilterServicio` → Excel de **Servicios**;
>   **Paso 2 (Destrucción)** que suma líneas `DESTRUCCION` a la misma hoja;
>   **Paso 3 (Ingresos)** que suma líneas `INGRESOS` (RECIBO) a la misma hoja, con filtro
>   de fecha del usuario sobre `Posting Date` y rama `MATERIAL DE EMPAQUE` vía `idh_especiales`;
>   **Paso 4 (Ocupación)** que suma líneas `ALMACENAMIENTO …` a la misma hoja, con join
>   EQUIVALENCIAS, promedio diario redondeado hacia arriba (ceil) y casos especiales
>   MODULA × 350.3 / PROFESIONAL+BIN → NATTURA;
>   **Paso 5 (Traslados)** que suma líneas `ALISTAMIENTO Y DESPACHO … CENTRO DE TRASLADOS`
>   a la misma hoja (suma de `SHU`/`CON` por negocio, sin filtro de fecha);
>   **Paso 6 (Maquila)** que suma líneas `ALISTAMIENTO DE MAQUILA CAJAS` y
>   `PICKING … MAQUILA ME` a la misma hoja (cajas/pallets sobre `Cantidad` + huellas,
>   filtro de fecha del usuario sobre `Posting Date`, `MATERIAL DE EMPAQUE` vía `idh`);
>   **Paso 7 (Exportaciones)** que suma líneas `ALISTAMIENTO Y DESPACHO PALLETS/CAJAS/UND
>   EXPO <canal>` a la misma hoja (pallets/cajas sobre `Ctd Ent.(UMV)` + huellas, filtro
>   de fecha del usuario sobre `Fecha factura`, `servicio = base & " " & canal`);
>   **Paso 8 (Etiquetas)** que suma líneas `ETIQUETAS REGULARES`/`ETIQUETAS REEMPAQUE`
>   a la misma hoja (suma de `cajas` por `tipo`, `negocio = CONSUMER` fijo, sin filtro
>   de fecha);
>   **Paso 9 (Paletizado)** que suma líneas `PALETIZADO EXPO <canal>` a la misma hoja
>   (suma de `TOTAL` por `(negocio, canal)` con `DESPACHO <> null`, `negocio` desde
>   `AREA` — HENKEL.PF/HENKEL.RT — y sin filtro de fecha);
>   **Paso 10 (Trincaje)** que suma una línea `DOBLE TRINCAJE` a la misma hoja
>   (conteo de filas con `despacho <> null` × 2, `negocio = CONSUMER` fijo, sin filtro
>   de fecha);
>   **Paso 11 (Planta)** que suma líneas `TRASLADO PALLETS PLANTA - CEDI` a la misma hoja
>   (suma de `estibas_consumer`/`estibas_profesional` por `negocio`, sin filtro de fecha);
>   **Paso 12 (Material)** que suma una línea `ALMACENAMIENTO PALLET BODEGA 8 ME GENERAL`
>   a la misma hoja (`ceil` del promedio diario de `cant_bodega_8` en el rango,
>   `negocio = MATERIAL DE EMPAQUE`);
>   **Paso 13 (Falabella)** que suma líneas `ETIQUETADO (FALABELLA)` a la misma hoja
>   (conteo de filas con `Entrega <> null` por `negocio`, sin filtro de fecha);
>   **Paso 14 (Otros)** que **anexa** los servicios pre-armados de `OTROS/otros_*.xlsx`
>   (cánones de arrendamiento, horas extras, cargues expo, descargues, sobre-costos…) a
>   la misma hoja — el archivo ya viene con su `valor`/`tarifa`/`costo` calculados, así que
>   no se calcula nada: se respetan tal cual (con tabla `OTROS` fija), sin filtro de
>   fecha (se anexan todas las filas); se inyecta **tras** `_apply_tarifas`;
>   **proceso en segundo plano con barra de avance** y **detención ante errores graves**
>   de archivo (corrupto / mal formato / faltante).
> - **Pendiente:** `RANGOS_FECHAS`. (Tarifa → costo **HECHO**: cada línea se cruza por
>   `servicio` con `tarifas.xlsx` y se calcula `costo_total = valor × tarifa`; las líneas
>   sin tarifa o con `valor = 0` no llegan al Excel.)
>   (Nota: `MATERIAL DE EMPAQUE` **no** requiere archivo aparte — se identifica por la columna
>   `Negocio` de `idh_especiales.xlsx`, que ya vale `MATERIAL DE EMPAQUE` para esos materiales.)
>
> Pensado para que otra instancia de Claude (o un desarrollador) retome el contexto.

---

## Cómo ejecutarlo

1. Doble-clic en **`iniciar.bat`** (levanta el servidor y abre **una sola pestaña** en
   el navegador en `http://127.0.0.1:8000`). Para detenerlo, cierra la ventana negra o `Ctrl+C`.
2. Alternativa manual: `python app.py` y abrir `http://127.0.0.1:8000`.

Requisitos (ya instalados en este entorno): Python 3.13, `fastapi`, `uvicorn`,
`pandas`, `openpyxl`, `python-calamine`. Ver `requirements.txt`.

---

## Qué hace el Paso 1

1. Lee `salidas_cons*.xlsx` (CONSUMER) y `salidas_prof*.xlsx` (PROFESIONAL).
2. Filtra las filas cuya **`Fecha factura`** cae dentro del rango elegido (inclusivo).
3. Asigna **Negocio** a cada material vía `AUXILIARES/idh_especiales.xlsx`
   (llave `Material` → `Negocio`). Si no hay coincidencia, el default depende de la
   carpeta de origen: `salidas_cons*` → **CONSUMER**, `salidas_prof*` → **PROFESIONAL**.
4. Trae **Unidades por pallet** y **Unidades por caja** de `HUELLAS/huellas.xlsx`
   (llave `Producto`).
5. Calcula, a partir de `Ctd Ent.(UMV)`:
   - `estibas = techo(cantidad / unidades_por_pallet)`
   - `cajas   = techo(cantidad / unidades_por_caja)`
   - (nulo si falta el divisor o es 0)
6. **`tipo_trabajo`**: cruza `Delivery` con `ENTREGA` de los archivos `adicionales*`
   (en `CONSUMER/` y `PROFESIONAL/`). Si hay match → el `TIPO` del archivo
   (p. ej. **EXTRA E.S**); si no → **NORMAL**.
7. **`tipo_despacho`**: cruza el cliente (`Nombre completo con tratamiento y título`)
   con `CEDI` de `AUXILIARES/tipo_despacho.xlsx`. Si hay match → el `TIPO`
   (**CROSS DOCKING**); si no → **ESTANDAR**.
8. **`AddServicio`** (de `logica.txt`, ~15 ramas): desdinamiza cada grupo en hasta
   3 medidas (**PICKING PALLETS / CAJAS / UNIDADES**) y asigna un **`servicio`**
   según `(negocio_facturador, tipo_trabajo, tipo_despacho, atributo)`. Luego
   **`FilterServicio`** descarta las que no facturan (nulas y dos exclusiones).
   `valor` = cantidad de esa línea (no es dinero todavía).
9. Exporta un Excel con una sola hoja **Servicios** en `FACTURAS_GENERADAS/`:
   una fila por servicio (`periodo, negocio, negocio_facturador, servicio, valor,
   unidades, proceso_extendido, macro_proceso, proceso_abreviado, tabla`). `periodo`
   = primer día del **último mes del rango** (igual en todas las filas; p. ej.
   21/05–21/06 → `01/06/2026`). `tipo_trabajo`/`tipo_despacho` **no aparecen** en el
   Excel (sólo se usan internamente para derivar `servicio`); las líneas se **agrupan
   por `servicio`** (sumando `valor`/`unidades`). No se vuelcan filas de detalle del salidas.

La lógica original de referencia está en **`logica.txt`** (Power Query / M).

---

## Qué hace el Paso 2 (Destrucción)

1. Lee todos los archivos `destruccion*` de `CONSUMER/` y `PROFESIONAL/` (ignora `~$`).
   **No filtra por fecha** (el contenido no la trae): se procesan **todos** los encontrados.
2. `negocio` = columna `Almacen` en MAYÚSCULAS (p. ej. `Consumer` → **CONSUMER**).
3. Cuenta las filas por `negocio` → ese conteo es `valor`.
4. Genera **una línea por negocio**: `servicio = ALISTAMIENTO Y DESPACHO PALLETS DESTRUCCION`,
   `unidades = 0`, `tabla = DESTRUCCION`, `proceso_extendido = "PICKING "+negocio`,
   `macro_proceso = OUT BOUND`, `proceso_abreviado = OUB`.
5. Se **suman** a la misma hoja **Servicios** del Paso 1 (la columna `tabla` distingue).

Errores de archivo (no se puede leer, falta la columna `Almacen`, archivo corrupto)
**detienen el proceso**. Un archivo `destruccion*` **vacío** (0 filas) **no es error**:
se omite y se continúa.

---

## Qué hace el Paso 3 (Ingresos)

1. Lee `ingresos_cons*.xlsx` (CONSUMER) y `ingresos_prof*.xlsx` (PROFESIONAL). El **área**
   viene del **prefijo del nombre** (como en salidas); el **filtro de fecha** se aplica sobre
   la columna **`Posting Date`** con el rango que eligió el usuario — **no** se usa
   `RANGOS_FECHAS` ni la fecha del nombre.
2. `documento_cruce = referencia` si no es nula, si no **`Texto cab.documento`**; las filas
   sin ninguno de los dos se descartan (fidelidad a `logica.txt`).
3. Trae **Unidades por pallet** y **Unidades por caja** de `huellas.xlsx` (llave `Producto`).
4. Asigna **`negocio_facturador`** por material vía `idh_especiales.xlsx` (llave `Material` →
   `Negocio`: puede ser `MATERIAL DE EMPAQUE`, `NATTURA`, `LAUNDRY`, …); si no hay match,
   default = el área (`ingresos_cons*` → CONSUMER, `ingresos_prof*` → PROFESIONAL).
   - Si `negocio_facturador = MATERIAL DE EMPAQUE`, entonces **`negocio = MATERIAL DE EMPAQUE`**
     (no el área). En caso contrario `negocio = área`.
5. Calcula, a partir de `Cantidad`:
   - `recibo_cajas    = techo(cantidad / unidades_por_caja)`
   - `recibo_pallets  = suelo(cantidad / unidades_por_pallet)` — **división entera** (pallets
     LLENOS, distinto a las estibas de salidas que usan techo)
   - `unidades_residuo = cantidad mod unidades_por_pallet` (el residuo fuera de pallets llenos)
   - `unidades_pallet_me = unidades_por_pallet × recibo_pallets`
   - `cajas_material  = techo(unidades_residuo / unidades_por_caja)`
   - (nulo donde falte el divisor o sea 0)
6. **Agrupa** por `(negocio, negocio_facturador)` sumando las medidas. Luego genera los
   servicios (unpivot + `FixAtributoME` + `FilterServicio` de `logica.txt`):
   - grupo **no-ME** → **`RECIBO CAJAS`** (`valor` = sum recibo_cajas, `unidades` = sum cantidad).
   - grupo **ME** → **`RECIBO CAJAS ME`** (`valor` = sum cajas_material, `unidades` = sum
     unidades_residuo) y **`RECIBO PALLETS ME`** (`valor` = `unidades` = sum unidades_pallet_me).
   - Se omiten las líneas con `valor` 0/nulo.
7. Se **suman** a la misma hoja **Servicios** con constantes `proceso_extendido = RECIBO`,
   `macro_proceso = IN BOUND`, `proceso_abreviado = INB`, `tabla = INGRESOS`. La columna
   `tabla` distingue de SALIDAS/DESTRUCCION. **No incluye `fecha`** como dimensión (usa
   `periodo` como el resto); el agrupado por `servicio` (suma `valor`/`unidades`) lo hace
   `_aggregate_servicios`.

Errores de archivo (no se puede leer un `ingresos*`, falta columna obligatoria) **detienen**
el proceso. **Sin archivos `ingresos*`** → advertencia y el Paso 3 se omite (no detiene);
SALIDAS y DESTRUCCIÓN siguen saliendo. Sin `huellas.xlsx`, los servicios de ingresos quedan
vacíos (advertencia).

---

## Paso 4 (Ocupacion) — ✅ HECHO

> Especificación (ya implementada; se conserva como documentación de la lógica). La
> traducción Power BI de referencia está en `logica.txt` (la query que empieza en
> `ArchivosOcupacion0`). Todo lo siguiente está verificado contra los archivos reales.

Facturación de **almacenamiento**. A diferencia de SALIDAS/INGRESOS, **no usa huellas ni idh**:
usa la tabla **EQUIVALENCIAS** (`AUXILIARES/equivalencias_almacenamiento.xlsx`, `Tipo` →
`Conversion liquidacion`) y el valor es un **promedio diario redondeado hacia arriba** (con
casos especiales MODULA × 350.3 y NATTURA para PROFESIONAL + BIN).

**Adaptaciones (mismo patrón que los pasos previos):**
1. Rango de fechas del **usuario** sobre la columna `Fecha` (no `RANGOS_FECHAS` ni la fecha del
   nombre). El **promedio** se calcula sobre los días de almacenamiento que caen **dentro del
   rango** elegido.
2. Área desde el prefijo del nombre: `ocupacion_cons*` → CONSUMER, `ocupacion_prof*` → PROFESIONAL.

**Datos verificados** (`ocupacion_cons_01-06-2026.xlsx` 384 filas, `ocupacion_prof_01-06-2026.xlsx`
256 filas):
- Columnas: `Fecha, Almacen, Tipo, Cantidad Instalada, Ocupación, % Ocupación`.
- **`Ocupación` y `% Ocupación` normalizan ambas a `OCUPACION`** → usar la **1ª ocurrencia**
  (la de valor, índice 4) con `find_column(..., occurrence=0)` (como el "Material" duplicado de salidas).
- 32 fechas de almacenamiento por archivo (21/05–21/06) → confirma el promedio diario.
- `Ocupación` trae `"-"` (→ `0` en `fnNumero`); el resto son enteros (p. ej. 3399).
- `Almacen`: `BN01 - HENKEL.RT` y `BN01 - HENKEL.RT - LDRY` → el texto **"LDRY"** identifica LAUNDRY.

**Mapeo EQUIVALENCIAS** (`Tipo` → `Conversion liquidacion`, 8 filas, sin duplicados):
`PALLET→PALLET BODEGA 7`, `PALLET (MEDIO)→MEDIO PALLET`, `PALLET (PICKING)→PALLET BODEGA 7`,
`PALLET (TIPO 2)→PALLET BODEGA 8`, `BIN→BIN`, `FLOW RACK→FLOW RACK`,
`PALLET (TIPO 3)→PALLET BODEGA 8 ME`, `MODULA→MODULA`.

**Lógica a traducir (paso a paso):**
1. Lee `ocupacion_cons*`/`ocupacion_prof*`: columnas `Fecha, Almacen, Tipo, Ocupación`.
2. `ocupacion` → numérico (`fnNumero`: `"-"` o vacío → `0`, resto a número, no parseable → null).
   `Tipo`/`Almacen` → upper/trim. Filtra `Fecha` no nula.
3. **Filtra `Fecha` por el rango del usuario** (inclusivo).
4. `negocio` = área; `negocio_facturador` = **"LAUNDRY" si `Almacen` contiene "LDRY"**, si no `negocio`.
5. **Join EQUIVALENCIAS** por `Tipo` (normalizado) → `conversion_liquidacion`.
6. **Suma por día**: `groupby (fecha_almacenamiento, negocio, nf, conversion)` sum `ocupacion`;
   descarta los días con suma `0` o nula (`FiltrarOcupacionValida`).
7. **Promedia por periodo**: `groupby (negocio, nf, conversion)` → **average** de los días, luego **`ceil`**.
8. **Servicios finales** por grupo:
   - `servicio = "ALMACENAMIENTO " + conversion`.
   - **`PALLET BODEGA 8 ME` de CONSUMER/PROFESIONAL → `ALMACENAMIENTO PALLET BODEGA 8`** (se agrupa/suma
     con el `BODEGA 8` existente en `_aggregate_servicios`, porque comparten `proceso_extendido="PALLETS"`).
     LAUNDRY conserva `… 8 ME`.
   - `proceso_extendido` (mapa fijo): `PALLET BODEGA 7/8/8 ME`→`PALLETS`, `MEDIO PALLET`→`MEDIO PALLET`,
     `FLOW RACK`→`FLOW RACK`, `BIN`→`BIN`, `MODULA`→`MODULA`, default→`BIN`.
   - `valor` = `ocupacion`, **excepto** `ALMACENAMIENTO MODULA` = `ceil(ocupacion * 350.3)`.
   - `negocio_facturador` final: si `nf == "PROFESIONAL"` y `servicio == "ALMACENAMIENTO BIN"` → **"NATTURA"**,
     si no `nf`.
   - `unidades = 0`, `macro_proceso = "ALMACENAMIENTO"`, `proceso_abreviado = "WHS"`, `tabla = "OCUPACION"`.

**Dónde tocar al implementar** (molde: Paso 3 / `_run_ingresos_pipeline`):
- `config.py`: `OCUPACION_GLOBS = {"cons": "ocupacion_cons*...", "prof": "ocupacion_prof*..."}`,
  `FILES["equivalencias"] = DIRS["auxiliares"] / "equivalencias_almacenamiento.xlsx"`,
  `OCUPACION_COLS = {fecha:"Fecha", almacen:"Almacen", tipo:"Tipo", ocupacion:"Ocupación"}`,
  `EQUIVALENCIAS_COLS = {archivo:"Archivo Alarmasblu", conversion:"Conversion liquidacion"}`.
- `processing/io_utils.py`: `find_ocupacion_files(base_dir)`, `read_ocupacion(path, area)` (con
  `ocupacion` parseado y `Ocupación` por 1ª ocurrencia), `read_equivalencias()` (molde `read_idh`).
- `processing/pipeline.py`: `_run_ocupacion_pipeline(start, end, ts_start, ts_end, emit, progress)`
  → `list[dict]`; wirear en `run_all` **tras INGRESOS** y sumar sus servicios antes de
  `_aggregate_servicios`; rebalancear progreso y ampliar `__main__`.
- `processing/excel_export.py`: **sin cambios** (las columnas ya cubren OCUPACION).

**Salida esperada** (rango 20/05–19/06/2026, ya calculada): **10 servicios OCUPACION**
(el `PALLET BODEGA 8 ME` de CONSUMER/PROFESIONAL se factura como `PALLET BODEGA 8` y se suma al
existente; LAUNDRY conservaría `… 8 ME`):
- CONSUMER / CONSUMER: MEDIO PALLET 110, PALLET BODEGA 7 3.748, PALLET BODEGA 8 187 (= 130 + 57 ME).
- CONSUMER / LAUNDRY: PALLET BODEGA 8 758.
- PROFESIONAL / NATTURA: BIN 798.
- PROFESIONAL / PROFESIONAL: FLOW RACK 141, MEDIO PALLET 336, MODULA 1.051, PALLET BODEGA 7 1.272,
  PALLET BODEGA 8 31 (= 6 + 25 ME).

> Validación (verificada): `python -m processing.pipeline` da **37 líneas** en la hoja
> Servicios (las 27 de SALIDAS/DESTRUCCION/INGRESOS + 10 de OCUPACION) **sin alterar** los pasos
> previos, y los 10 valores coinciden con la "Salida esperada" de arriba. Sin `ocupacion*` →
> advertencia y se omite (no detiene). Sin `equivalencias` → OCUPACION sin servicios convertibles.
> Archivo obligatorio roto → `BlockingError`.

---

## Paso 5 (Traslados) — ✅ HECHO

> Especificación (ya implementada). Traducción de la query `traslados*` de `logica.txt`
> (Power Query). Es la pipeline más simple de todas.

Facturación del **centro de traslados**. A diferencia de los pasos previos, **no usa huellas,
idh ni equivalencias, y no filtra por fecha**: el valor es el **sumatorio** directo de `SHU` y
`CON` por `negocio`, y se procesan **todos** los `traslados*` encontrados (decisión del usuario,
molde DESTRUCCIÓN — `logica.txt` trae la fecha del nombre, pero el contenido no tiene columna
de periodo).

**Adaptaciones (mismo patrón que los pasos previos):**
1. **Sin filtro de fecha**: se procesan TODOS los `traslados*` (no se usa el rango del calendario
   ni `RANGOS_FECHAS` ni la fecha del nombre).
2. Área desde el prefijo del nombre: `traslados_cons*` → CONSUMER, `traslados_prof*` → PROFESIONAL.

**Datos verificados** (`traslados_cons_01-06-2026.xlsx` 5 filas, `traslados_prof_01-06-2026.xlsx`
12 filas): export de SAP de entregas con `Delivery, SHU, CON` (entre otras). Solo se leen esas
tres columnas (`Ship-To`/`Nombre`/`No.Of Lines` no se usan en la agregación).

**Lógica a traducir (paso a paso):**
1. Lee `traslados_cons*`/`traslados_prof*`.
2. **Filtra filas con `Delivery <> null`** (`#"Filas filtradas"` de `logica.txt`).
3. `negocio` = área; agrupa por `negocio` sumando `SHU` y `CON`.
4. Servicios finales por grupo (unpivot de `logica.txt`):
   - **`ALISTAMIENTO Y DESPACHO CAJAS CENTRO DE TRASLADOS`**: `valor = sum(SHU)`, `unidades = 0`.
   - **`ALISTAMIENTO Y DESPACHO UNIDADES CENTRO DE TRASLADOS`**: `valor = sum(CON)`, `unidades = sum(CON)`.
   - `negocio_facturador = negocio`, `proceso_extendido = "PICKING "+negocio`,
     `macro_proceso = "OUT BOUND"`, `proceso_abreviado = "OUB"`, `tabla = "TRASLADOS"`.
   - Se omiten los servicios con `valor = 0` (convención del bot; en `logica.txt` se conservan).

**Dónde tocar al implementar** (molde: Paso 2 / `run_step2_destruccion`, sin `start/end`):
- `config.py`: `TRASLADOS_GLOBS = {"cons": "traslados_cons*...", "prof": "traslados_prof*..."}`,
  `TRASLADOS_COLS = {delivery:"Delivery", shu:"SHU", con:"CON"}`.
- `processing/io_utils.py`: `find_traslados_files(base_dir)`, `read_traslados(path, area)`
  (molde `read_ocupacion`).
- `processing/pipeline.py`: `_run_traslados_pipeline(emit, progress)` → `list[dict]`; wirear en
  `run_all` **tras OCUPACIÓN** y sumar sus servicios antes de `_aggregate_servicios`.
- `processing/excel_export.py`: **sin cambios** (las columnas ya cubren TRASLADOS).

**Salida esperada** (rango 20/05–19/06/2026, ya calculada): **4 servicios TRASLADOS**
- CONSUMER / CONSUMER: `CAJAS CENTRO DE TRASLADOS` 71, `UNIDADES CENTRO DE TRASLADOS` 489.
- PROFESIONAL / PROFESIONAL: `CAJAS CENTRO DE TRASLADOS` 125, `UNIDADES CENTRO DE TRASLADOS` 29.949.

> Validación (verificada): `python -m processing.pipeline` da **41 líneas** en la hoja Servicios
> (las 37 de SALIDAS/DESTRUCCION/INGRESOS/OCUPACION + 4 de TRASLADOS) **sin alterar** los pasos
> previos, y los 4 valores coinciden con la "Salida esperada". Sin `traslados*` → advertencia y
> se omite (no detiene). Archivo obligatorio roto (sin columna `Delivery`/`SHU`/`CON`) → `BlockingError`.

---

## Paso 6 (Maquila) — ✅ HECHO

> Especificación (ya implementada). Traducción de la query `FilesMaquila` de `logica.txt`
> (Power Query, líneas 1–336). Es la pipeline **gemela de INGRESOS**: mismas 6 medidas de
> recibo (`floor`/`mod`/`ceil`), mismos cruces (huellas + idh, `MATERIAL DE EMPAQUE` vía
> `idh_especiales.Negocio`); sólo cambian los servicios y las constantes finales.

Facturación de **subcontratación (maquila)**. Lee `maquila_cons*`/`maquila_prof*` de la
carpeta **`MAQUILA/`** (ojo: ambos ficheros viven en esa carpeta, no en `CONSUMER/` y
`PROFESIONAL/` como el resto). El área viene del prefijo del nombre.

**Adaptaciones (mismo patrón que INGRESOS):**
1. Rango de fechas del **usuario** sobre la columna `Posting Date` (no `RANGOS_FECHAS`
   ni la fecha del nombre).
2. Área desde el prefijo del nombre: `maquila_cons*` → CONSUMER, `maquila_prof*` → PROFESIONAL.
3. `MATERIAL DE EMPAQUE` se identifica por la columna `Negocio` de `idh_especiales.xlsx`
   (no hay tabla aparte); si `negocio_facturador` es `MATERIAL DE EMPAQUE`, `negocio`
   también lo es (y se generan los servicios `… ME`), igual que en ingresos.

**Medidas** (sobre `Cantidad` + `unidades_pallet`/`unidades_caja` de `huellas.xlsx`):
`pallets_maquila = floor(Cantidad / pallet)` (pallets llenos), `unidades_cajas_maquila = Cantidad mod pallet`
(residuo), `cajas_maquila = ceil(residuo / caja)`, `cajas_generales = ceil(Cantidad / caja)`,
`unidades_generales = Cantidad`, `unidades_pallet_maquila = pallet × pallets_maquila`. Son
**las mismas fórmulas de recibo** de Ingresos (reutiliza `_floor_series`/`_mod_series`/`_ceil_series`).

**Lógica a traducir (paso a paso):**
1. Lee `maquila_cons*`/`maquila_prof*`: columnas `Posting Date, Cantidad, Material`.
2. Filtra filas con `Material` no nulo (`#"Filas filtradas2"` de `logica.txt`).
3. **Filtra `Posting Date` por el rango del usuario** (inclusivo).
4. `negocio` = área; `negocio_facturador` = idh si match (incl. `MATERIAL DE EMPAQUE`), si no área.
   Si `nf = MATERIAL DE EMPAQUE` → `negocio = MATERIAL DE EMPAQUE`.
5. Cruces huellas (pallet/caja) e idh (negocio).
6. Calcula las 6 medidas (vectorizado).
7. **Agrupa** por `(negocio, negocio_facturador)` sumando las medidas; genera los servicios
   (`KeepFlag` + `MapServicio` de `logica.txt`):
   - grupo **no-ME** → **`ALISTAMIENTO DE MAQUILA CAJAS`** (`valor` = sum cajas_generales,
     `unidades` = sum Cantidad).
   - grupo **ME** → **`PICKING PALLETS MAQUILA ME`** (`valor` = sum pallets_maquila,
     `unidades` = sum unidades_pallet_maquila) y **`PICKING CAJAS MAQUILA ME`** (`valor` = sum
     cajas_maquila, `unidades` = sum unidades_cajas_maquila).
   - Se omiten las líneas con `valor` 0/nulo.
8. Se **suman** a la misma hoja **Servicios** con constantes `proceso_extendido = MAQUILA`,
   `macro_proceso = OTROS`, `proceso_abreviado = MAQ`, `tabla = MAQUILA`.

> **Nota sobre `unidades`** (decisión del usuario): el `AddUnidadesServicio` de `logica.txt`
> compara `[atributo]` contra los nombres `"RECIBO CAJAS ME"`/`"RECIBO CAJAS"` (que son de
> **Ingresos**, no de Maquila), así que en el PQ original **todo** cae al `else` y `unidades`
> = `unidades_pallet_maquila` para los 5 servicios (bug heredado de plantilla). Aquí se aplica
> un **mapeo coherente**: cajas → `unidades_generales`, cajas ME → `unidades_cajas_maquila`,
> pallets ME → `unidades_pallet_maquila`. Los `valor` sí son fieles al PQ.

**Dónde tocar al implementar** (molde: Paso 3 / `_run_ingresos_pipeline`):
- `config.py`: `DIRS["maquila"] = BASE_DIR / "MAQUILA"`,
  `MAQUILA_GLOBS = {"cons": "maquila_cons*...", "prof": "maquila_prof*..."}`,
  `MAQUILA_COLS = {posting_date:"Posting Date", cantidad:"Cantidad", material:"Material"}`.
- `processing/io_utils.py`: `find_maquila_files(base_dir)` (los DOS globs sobre la misma
  carpeta `MAQUILA/`) y `read_maquila(path, area)` (molde `read_ingresos`, 3 cols + `area`).
- `processing/pipeline.py`: `_run_maquila_pipeline(start, end, ts_start, ts_end, emit, progress)`
  → `list[dict]`; wirear en `run_all` **tras TRASLADOS** y sumar sus servicios antes de
  `_aggregate_servicios`; rebalancear progreso (MAQUILA 94/96, Excel 98) y ampliar `__main__`.
- `processing/excel_export.py`: **sin cambios** (las columnas ya cubren MAQUILA).

**Salida esperada** (rango 20/05–19/06/2026, ya calculada): **5 servicios MAQUILA**
- CONSUMER / CONSUMER: `ALISTAMIENTO DE MAQUILA CAJAS` valor 40.754 (unid. 572.928).
- PROFESIONAL / NATTURA: `ALISTAMIENTO DE MAQUILA CAJAS` valor 1.223 (unid. 53.833).
- PROFESIONAL / PROFESIONAL: `ALISTAMIENTO DE MAQUILA CAJAS` valor 7.410 (unid. 91.392).
- MATERIAL DE EMPAQUE / MATERIAL DE EMPAQUE: `PICKING CAJAS MAQUILA ME` valor 895
  (unid. 222.231) y `PICKING PALLETS MAQUILA ME` valor 14 (unid. 92.160).

> Validación (verificada): `python -m processing.pipeline` da **46 líneas** en la hoja Servicios
> (las 41 de SALIDAS/DESTRUCCION/INGRESOS/OCUPACION/TRASLADOS + 5 de MAQUILA) **sin alterar**
> los pasos previos, y los 5 valores coinciden con la "Salida esperada". Sin `maquila*` →
> advertencia y se omite (no detiene). Sin `huellas`/`idh` → MAQUILA con valores vacíos
> (advertencia). Archivo obligatorio roto (sin `Posting Date`/`Material`/`Cantidad`) → `BlockingError`.

---

## Paso 7 (Exportaciones) — ✅ HECHO

> Especificación (ya implementada). Traducción de la query `FilesExportacion` de
> `logica.txt` (Power Query, líneas 1–213). Sigue el molde de INGRESOS/MAQUILA
> (huellas + filtro de fecha del usuario), con dos particularidades: el **`canal`** forma
> parte del nombre del servicio, y `negocio_facturador` = `negocio` (el `JoinIDHEspeciales`
> del PQ es vestigial — ver abajo).

Facturación de **exportaciones**. Lee `exportacion_cons*`/`exportacion_prof*` de la
carpeta **`EXPORTACIONES/`** (ambos ficheros viven ahí, como `MAQUILA/`; no en
`CONSUMER/` y `PROFESIONAL/`). El área viene del prefijo del nombre. Los ficheros
`paletizado*`/`trincaje*` (que también viven en `EXPORTACIONES/`) **no se procesan**:
`logica.txt` sólo admite ficheros que empiecen por `exportacion`.

**Adaptaciones (mismo patrón que INGRESOS/MAQUILA):**
1. Rango de fechas del **usuario** sobre la columna `Fecha factura` (no `RANGOS_FECHAS`
   ni la fecha del nombre — la ventana del PQ se reemplaza por el rango libre del
   calendario, igual que en los pasos previos).
2. Área desde el prefijo del nombre: `exportacion_cons*` → CONSUMER,
   `exportacion_prof*` → PROFESIONAL.

**Datos verificados** (`exportacion_cons_01-06-2026.xlsx` 359 filas,
`exportacion_prof_01-06-2026.xlsx` 2.178 filas, 85 columnas):
- Columnas leídas: `Delivery, Fecha factura, Material, Ctd Ent.(UMV), Canal distribución`
  (`Texto breve de material` no se usa en la agregación).
- `Fecha factura` ya viene como datetime (sin nulos); rango cons 22/05–19/06, prof 21/05–20/06.
- **`Canal distribución`**: valores **`EX`** e **`IC`** (cons tiene ambos: EX 40 / IC 319;
  prof sólo IC). Se conserva tal cual (MAYÚSCULAS/trim) y se pega al nombre del servicio.
- `Material` es numérico (p. ej. 3083475); se normaliza a texto para el cruce huellas,
  igual que en salidas. `Material` (col 16) y `Material.1` (col 66) son nombres distintos,
  así que no hay ambigüedad de ocurrencia (no como el `Material` duplicado de salidas).

**Lógica a traducir (paso a paso):**
1. Lee `exportacion_cons*`/`exportacion_prof*`: columnas `Delivery, Fecha factura, Material,
   Ctd Ent.(UMV), Canal distribución`.
2. **Filtra filas con `Delivery <> null`** (`#"Filas filtradas"` de `logica.txt`).
3. **Filtra `Fecha factura` por el rango del usuario** (inclusivo).
4. `negocio` = área; **`negocio_facturador` = `negocio`** (ver nota sobre el idh vestigial).
5. Cruce **huellas** (pallet/caja) por `material` (foto única por producto, como salidas).
   **No hay cruce idh** (el del PQ no se usa).
6. Medidas (sobre `Ctd Ent.(UMV)` + `unidades_pallet`/`unidades_caja`):
   - `pallets = ceil(cantidad / unidades_pallet)` (RoundUp, como estibas).
   - `cajas   = ceil(cantidad / unidades_caja)` (RoundUp).
   - `unidades = unidades_caja × cajas` — **redondea hacia ARRIBA a múltiplo de caja**
     (`caja × ceil(cant/caja)`); **NO es `cantidad`** (son las unidades "facturadas").
   - (nulo donde falte el divisor o sea 0).
7. **Agrupa** por `(negocio, canal)` sumando `pallets`, `cajas` y `unidades` (la `fecha`
   del nombre del PQ se descarta: el bot usa `periodo` como el resto). Luego emite los
   servicios (unpivot + `servicio = base & " " & canal` de `logica.txt`):
   - **`ALISTAMIENTO Y DESPACHO PALLETS EXPO <canal>`**: `valor` = sum(pallets),
     `unidades` = sum(unidades).
   - **`ALISTAMIENTO Y DESPACHO CAJAS EXPO <canal>`**: `valor` = sum(cajas),
     `unidades` = sum(unidades).
   - **`ALISTAMIENTO Y DESPACHO UND EXPO <canal>`**: `valor` = `unidades` = sum(unidades).
   - Se omiten las líneas con `valor` 0/nulo (convención del bot; en `logica.txt` no hay
     `FilterServicio` para exportaciones).
8. Se **suman** a la misma hoja **Servicios** con constantes `proceso_extendido =
   "PICKING "+negocio`, `macro_proceso = "OUT BOUND"`, `proceso_abreviado = "OUB"`,
   `tabla = "EXPORTACIONES"`.

> **Nota sobre `negocio_facturador` (idh vestigial):** la query `FilesExportacion` hace
> `JoinIDHEspeciales` (NestedJoin `material`→`idh`) pero **nunca expande** esa columna
> —es código muerto— y luego `#"Columna duplicada"` copia `negocio → negocio_facturador`.
> Por tanto, en el PQ `negocio_facturador = negocio = área` (sin LAUNDRY/NATTURA/MATERIAL
> DE EMPAQUE). El bot es fiel a eso: `nf = negocio`. (Si se quisiera `nf` por idh, habría
> que decidirlo aparte — no es lo que hace `logica.txt`.)

**Dónde tocar al implementar** (molde: Paso 3 / `_run_ingresos_pipeline`):
- `config.py`: `DIRS["exportaciones"] = BASE_DIR / "EXPORTACIONES"`,
  `EXPORTACION_GLOBS = {"cons": "exportacion_cons*...", "prof": "exportacion_prof*..."}`,
  `EXPORTACION_COLS = {delivery:"Delivery", fecha:"Fecha factura", material:"Material",
  cantidad:"Ctd Ent.(UMV)", canal:"Canal distribución"}`.
- `processing/io_utils.py`: `find_exportacion_files(base_dir)` (los DOS globs sobre la
  misma carpeta `EXPORTACIONES/`, molde `find_maquila_files`) y `read_exportacion(path, area)`
  (molde `read_ingresos`, 5 cols + `area`).
- `processing/pipeline.py`: `_run_exportacion_pipeline(start, end, ts_start, ts_end, emit, progress)`
  → `list[dict]`; wirear en `run_all` **tras MAQUILA** y sumar sus servicios antes de
  `_aggregate_servicios`; rebalancear progreso (EXPO 96/97, Excel 98) y ampliar `__main__`.
- `processing/excel_export.py`: **sin cambios** (las columnas ya cubren EXPORTACIONES).

**Salida esperada** (rango 20/05–19/06/2026, ya calculada): **9 servicios EXPORTACIONES**
(3 grupos `(negocio, canal)` × 3 medidas; `paletizado*`/`trincaje*` no se procesan):
- CONSUMER / EX: `PALLETS EXPO EX` 67, `CAJAS EXPO EX` 4.709, `UND EXPO EX` 79.188.
- CONSUMER / IC: `PALLETS EXPO IC` 921, `CAJAS EXPO IC` 85.048, `UND EXPO IC` 1.071.900.
- PROFESIONAL / IC: `PALLETS EXPO IC` 2.467, `CAJAS EXPO IC` 82.642, `UND EXPO IC` 4.602.297.

> Validación (verificada): `python -m processing.pipeline` da **55 líneas** en la hoja
> Servicios (las 46 de SALIDAS/DESTRUCCION/INGRESOS/OCUPACION/TRASLADOS/MAQUILA + 9 de
> EXPORTACIONES) **sin alterar** los pasos previos, y los 9 valores coinciden con la
> "Salida esperada". Coherencia OK: `unidades` es compartida por los 3 servicios de cada
> grupo y, para `UND EXPO`, `valor == unidades` (fiel al `DuplicateUND` del PQ). Sin
> `exportacion*` → advertencia y se omite (no detiene). Sin `huellas` → medidas vacías
> (advertencia). Archivo obligatorio roto (sin `Delivery`/`Fecha factura`/`Material`/
> `Ctd Ent.(UMV)`/`Canal distribución`) → `BlockingError`.

---

## Paso 8 (Etiquetas) — ✅ HECHO

> Especificación (ya implementada). Traducción de la query `etiquetas` de `logica.txt`
> (Power Query). Es, junto a TRASLADOS, **la pipeline más simple de todas**: sin huellas,
> sin idh, sin equivalencias y **sin filtro de fecha** (el PQ no aplica `RANGOS_FECHAS`:
> solo agrupa por la fecha del **nombre** del archivo, que el bot descarta y reemplaza
> por `periodo`). Lee el fichero `etiquetas_<fecha>.xlsx` que vive en la carpeta
> **`MAQUILA/`** (sin separación cons/prof, a diferencia de maquila/expo).

Facturación de **etiquetado / reempaque**. La columna `tipo` (`REGULARES` / `REEMPAQUE`)
mapea al nombre del servicio; el `valor` es el **sumatorio** de `cajas` por servicio.
`negocio` es siempre **`CONSUMER`** (constante del PQ, no se deriva del nombre ni del
contenido).

**Datos verificados** (`etiquetas_01-06-2026.xlsx` 51 filas):
- Columnas en el fuente: `fecha, entrega, cajas, tipo`. El PQ **descarta `fecha` y
  `entrega`** (la `fecha` final viene del **nombre** del archivo), así que el bot sólo
  lee `cajas` y `tipo`.
- `tipo`: valores **`REGULARES`** (6 filas, 3.433 cajas) y **`REEMPAQUE`** (45 filas,
  247 cajas). `cajas` es entero; total 3.680.

**Lógica a traducir (paso a paso):**
1. Lee `etiquetas*` (de `MAQUILA/`): columnas `cajas, tipo`.
2. `tipo` → servicio (`ReplaceValue` de `logica.txt`): **`REGULARES` → `ETIQUETAS
   REGULARES`**, **`REEMPAQUE` → `ETIQUETAS REEMPAQUE`**. Otros valores de `tipo`
   pasan tal cual (el PQ sólo reemplaza esos dos).
3. **Agrupa** por `servicio` sumando `cajas` → `valor` (el PQ agrupa por `(fecha del
   nombre, servicio)`; el bot descarta la fecha del nombre y usa `periodo`, así que
   agrega solo por `servicio`).
4. Servicios finales (una línea por `servicio`): `negocio = negocio_facturador =
   "CONSUMER"` (constante), `valor = sum(cajas)`, `unidades = 0`,
   `proceso_extendido = servicio` (`DuplicateColumn` del PQ), `macro_proceso = "OTROS"`,
   `proceso_abreviado = "OTR"`, `tabla = "ETIQUETAS"`. Se omiten los servicios con
   `valor = 0` (convención del bot; el PQ no los filtra).

**Dónde tocar al implementar** (molde: Paso 5 / `_run_traslados_pipeline`, sin `start/end`):
- `config.py`: `ETIQUETAS_GLOB = "etiquetas*.[xX][lL][sS][xX]"` (sobre `DIRS["maquila"]`),
  `ETIQUETAS_COLS = {cajas:"cajas", tipo:"tipo"}`,
  `ETIQUETAS_SERVICIO_MAP = {REGULARES:"ETIQUETAS REGULARES", REEMPAQUE:"ETIQUETAS REEMPAQUE"}`,
  `ETIQUETAS_NEGOCIO = "CONSUMER"`.
- `processing/io_utils.py`: `find_etiquetas_files(base_dir)` (glob `etiquetas*` sobre
  `DIRS["maquila"]`, devuelve solo rutas — sin área, pues `negocio` es constante) y
  `read_etiquetas(path)` (molde `read_traslados`, 2 cols).
- `processing/pipeline.py`: `_run_etiquetas_pipeline(emit, progress)` → `list[dict]`;
  wirear en `run_all` **tras EXPORTACIONES** y sumar sus servicios antes de
  `_aggregate_servicios` (progreso 97; Excel sigue en 98); ampliar `__main__`.
- `processing/excel_export.py`: **sin cambios** (las columnas ya cubren ETIQUETAS).

**Salida esperada** (rango 20/05–19/06/2026, ya calculada): **2 servicios ETIQUETAS**
- CONSUMER / CONSUMER: `ETIQUETAS REGULARES` 3.433, `ETIQUETAS REEMPAQUE` 247.

> Validación (verificada): `python -m processing.pipeline` da **57 líneas** en la hoja
> Servicios (las 55 de SALIDAS/DESTRUCCION/INGRESOS/OCUPACION/TRASLADOS/MAQUILA/EXPORTACIONES
> + 2 de ETIQUETAS) **sin alterar** los pasos previos y **sin issues**, y los 2 valores
> coinciden con la "Salida esperada". Sin `etiquetas*` → advertencia y se omite (no
> detiene). Archivo obligatorio roto (sin `cajas`/`tipo`) → `BlockingError`.

---

## Paso 9 (Paletizado) — ✅ HECHO

> Especificación (ya implementada). Traducción de la query `paletizado` de `logica.txt`
> (Power Query). Tan simple como ETIQUETAS/TRASLADOS: **sin huellas, sin idh, sin cruce
> y sin filtro de fecha** (el PQ no aplica `RANGOS_FECHAS`: solo agrupa por la fecha del
> **nombre** del archivo, que el bot descarta y reemplaza por `periodo`). Lee el fichero
> `paletizado_<fecha>.xlsx` que vive en **`EXPORTACIONES/`** (junto a `exportacion_*`/
> `trincaje_*`; la query de exportación del Paso 7 **no** lo procesa, `logica.txt` lo
> admite en su propia query).

Facturación del **paletizado de exportación**. El `canal` forma parte del nombre del
servicio (`PALETIZADO EXPO <canal>`, como en EXPORTACIONES); el `valor` es el **sumatorio**
de `TOTAL` por `(negocio, canal)`. La particularidad frente al resto de pasos es que el
`negocio` se deriva de la columna **`AREA`** (`HENKEL.PF` → PROFESIONAL, `HENKEL.RT` →
CONSUMER), no del prefijo del nombre ni del idh.

**Datos verificados** (`paletizado_01-06-2026.xlsx` 91 filas, 10 columnas):
- Columnas en el fuente: `AREA, N-DE EXPO, DESPACHO, ORDEN SAP, CONTE/OR, N-CONTEOR,
  TIPO, ESTATUS, TOTAL, CANAL`. Sólo se usan `AREA` (→ negocio), `DESPACHO` (filtro),
  `TOTAL` (→ valor) y `CANAL` (→ sufijo del servicio); las demás no se leen.
- `AREA`: `HENKEL.PF` (→ PROFESIONAL) y `HENKEL.RT` (→ CONSUMER).
- `CANAL`: `IC`. `DESPACHO`: sin nulos (91/91). `TOTAL`: entero, suma 1.257.

**Lógica a traducir (paso a paso):**
1. Lee `paletizado*` (de `EXPORTACIONES/`): columnas `area, despacho, total, canal`.
2. **Filtra filas con `DESPACHO <> null`** (`#"Filas filtradas1"` de `logica.txt`).
3. `negocio` = `AREA` mapeada (`ReplaceValue`): `HENKEL.PF` → **PROFESIONAL**,
   `HENKEL.RT` → **CONSUMER**; otros valores de `AREA` pasan tal cual. Descarta filas
   con `AREA` nula (filtro final `negocio <> null`).
4. **Agrupa** por `(negocio, canal)` sumando `TOTAL` → `valor` (el PQ agrupa por
   `(fecha del nombre, negocio, canal)`; el bot descarta la fecha del nombre y usa
   `periodo`, así que agrega por `(negocio, canal)`).
5. Servicios finales (una línea por `(negocio, canal)`): `servicio = "PALETIZADO EXPO "
   + canal`, `negocio_facturador = negocio` (`DuplicateColumn` del PQ), `valor =
   sum(TOTAL)`, `unidades = 0`, `proceso_extendido = "EXPORTACION (PALETIZADO)"`,
   `macro_proceso = "OTROS"`, `proceso_abreviado = "PAL"`, `tabla = "PALETIZADO"`. Se
   omiten los servicios con `valor = 0` (convención del bot; el PQ no los filtra).

**Dónde tocar al implementar** (molde: Paso 8 / `_run_etiquetas_pipeline`, sin `start/end`):
- `config.py`: `PALETIZADO_GLOB = "paletizado*.[xX][lL][sS][xX]"` (sobre
  `DIRS["exportaciones"]`), `PALETIZADO_COLS = {area:"AREA", despacho:"DESPACHO",
  total:"TOTAL", canal:"CANAL"}`, `PALETIZADO_AREA_MAP = {HENKEL.PF:"PROFESIONAL",
  HENKEL.RT:"CONSUMER"}`.
- `processing/io_utils.py`: `find_paletizado_files(base_dir)` (glob `paletizado*` sobre
  `DIRS["exportaciones"]`, devuelve solo rutas — sin área, pues `negocio` viene de la
  columna `AREA`) y `read_paletizado(path)` (molde `read_traslados`, 4 cols).
- `processing/pipeline.py`: `_run_paletizado_pipeline(emit, progress)` → `list[dict]`;
  wirear en `run_all` **tras ETIQUETAS** y sumar sus servicios antes de
  `_aggregate_servicios` (progreso 97; Excel sigue en 98); ampliar `__main__`.
- `processing/excel_export.py`: **sin cambios** (las columnas ya cubren PALETIZADO).

**Salida esperada** (rango 20/05–19/06/2026, ya calculada): **2 servicios PALETIZADO**
- CONSUMER / CONSUMER: `PALETIZADO EXPO IC` 58.
- PROFESIONAL / PROFESIONAL: `PALETIZADO EXPO IC` 1.199.

> Validación (verificada): `python -m processing.pipeline` da **59 líneas** en la hoja
> Servicios (las 57 de los pasos 1–8 + 2 de PALETIZADO) **sin alterar** los pasos previos
> y **sin issues**, y los 2 valores coinciden con la "Salida esperada". Sin `paletizado*`
> → advertencia y se omite (no detiene). Archivo obligatorio roto (sin `AREA`/`DESPACHO`/
> `TOTAL`/`CANAL`) → `BlockingError`.

---

## Paso 10 (Trincaje) — ✅ HECHO

> Especificación (ya implementada). Traducción de la query `trincaje` de `logica.txt`
> (Power Query). Es la pipeline **más simple de todas** (junto a TRASLADOS/ETIQUETAS):
> sin huellas, sin idh, sin equivalencias y **sin filtro de fecha** (el PQ no aplica
> `RANGOS_FECHAS`: solo agrupa por la fecha del **nombre** del archivo, que el bot
> descarta y reemplaza por `periodo`). Lee el fichero `trincaje_<fecha>.xlsx` que vive en
> **`EXPORTACIONES/`** (junto a `exportacion_*`/`paletizado_*`; la query de exportación
> del Paso 7 **no** lo procesa, `logica.txt` lo admite en su propia query).

Facturación del **trincaje (doble trincaje)** de exportación. El `valor` es el
**conteo de filas con `despacho <> null` × 2** (`logica.txt`: `Table.RowCount(_) * 2`).
`negocio` es siempre **`CONSUMER`** (constante del PQ, no se deriva del nombre ni del
contenido), y se genera **una sola línea**: `servicio = proceso_extendido =
"DOBLE TRINCAJE"`.

**Datos verificados** (`trincaje_01-06-2026.xlsx` 27 filas, 5 columnas):
- Columnas en el fuente: `despacho, exp, destino, contenedor, n-cont`. El PQ las lee
  todas pero, tras el filtro, **sólo conserva el NOMBRE del archivo** (la `fecha` final
  viene del nombre); así que el bot **solo lee `despacho`** (para el filtro `<> null`).
  `exp`/`destino`/`contenedor`/`n-cont` no se usan en la agregación.
- `despacho`: 27/27 no nulas → `valor = 27 × 2 = 54`.

**Lógica a traducir (paso a paso):**
1. Lee `trincaje*` (de `EXPORTACIONES/`): columna `despacho`.
2. **Filtra filas con `despacho <> null`** (`#"Filas filtradas"` de `logica.txt`).
3. `valor = (nº de filas) × 2` (`#"Filas agrupadas"` del PQ: agrupa por la fecha del
   nombre y hace `Table.RowCount(_) * 2`; el bot descarta la fecha del nombre y suma
   todos los archivos → `(total filas) × 2`, idéntico con un solo archivo).
4. Servicio final (**una línea**): `negocio = negocio_facturador = "CONSUMER"` (constante),
   `servicio = proceso_extendido = "DOBLE TRINCAJE"` (el PQ duplica `proceso_extendido
   → servicio`), `valor = conteo × 2`, `unidades = 0`, `macro_proceso = "OTROS"`,
   `proceso_abreviado = "OTR"`, `tabla = "TRINCAJE"`. Convención del bot: si `valor = 0`
   se omite (el PQ no lo filtra).

**Dónde tocar al implementar** (molde: Paso 9 / `_run_paletizado_pipeline`, sin `start/end`):
- `config.py`: `TRINCAJE_GLOB = "trincaje*.[xX][lL][sS][xX]"` (sobre `DIRS["exportaciones"]`),
  `TRINCAJE_COLS = {despacho:"despacho"}`, `TRINCAJE_NEGOCIO = "CONSUMER"`,
  `TRINCAJE_SERVICIO = "DOBLE TRINCAJE"`, `TRINCAJE_FACTOR = 2`.
- `processing/io_utils.py`: `find_trincaje_files(base_dir)` (glob `trincaje*` sobre
  `DIRS["exportaciones"]`, devuelve solo rutas — sin área, pues `negocio` es constante)
  y `read_trincaje(path)` (molde `read_etiquetas`, 1 col `despacho`).
- `processing/pipeline.py`: `_run_trincaje_pipeline(emit, progress)` → `list[dict]`;
  wirear en `run_all` **tras PALETIZADO** y sumar sus servicios antes de
  `_aggregate_servicios` (progreso 97; Excel sigue en 98); ampliar `__main__`.
- `processing/excel_export.py`: **sin cambios** (las columnas ya cubren TRINCAJE).

**Salida esperada** (rango 20/05–19/06/2026, ya calculada): **1 servicio TRINCAJE**
- CONSUMER / CONSUMER: `DOBLE TRINCAJE` 54 (= 27 filas × 2).

> Validación (verificada): `python -m processing.pipeline` da **60 líneas** en la hoja
> Servicios (las 59 de los pasos 1–9 + 1 de TRINCAJE) **sin alterar** los pasos previos
> y **sin issues**, y el valor (54) coincide con la "Salida esperada". Sin `trincaje*`
> → advertencia y se omite (no detiene). Archivo obligatorio roto (sin `despacho`)
> → `BlockingError`.

---

## Paso 11 (Planta) — ✅ HECHO

> Especificación (ya implementada). Traducción de la query `planta` de `logica.txt`
> (Power Query). Tan simple como TRINCAJE/PALETIZADO: **sin huellas, sin idh, sin
> equivalencias y sin filtro de fecha** (el PQ elimina las columnas `fecha`/`semana` del
> archivo y no aplica `RANGOS_FECHAS`: agrupa por la fecha del **nombre** del archivo,
> que el bot descarta y reemplaza por `periodo`). Un SOLO fichero `planta_<fecha>.xlsx`
> cubre ambos negocios.

Facturación del **traslado de pallets de planta → CEDI**. El archivo trae, por día, las
estibas movidas por negocio en dos columnas (`estibas_consumer`, `estibas_profesional`);
el `valor` es el **sumatorio** de cada columna. El `negocio` viene del **nombre de la
columna** (`estibas_consumer` → CONSUMER, `estibas_profesional` → PROFESIONAL), **no** de
la carpeta ni del idh.

**Datos verificados** (`planta_01-06-2026.xlsx` 27 filas, 4 columnas, en `PROFESIONAL/`):
- Columnas en el fuente: `semana, fecha, estibas_consumer, estibas_profesional`. El PQ
  **elimina `fecha` y `semana`** (`#"Columnas quitadas"`: la `fecha` final viene del
  NOMBRE del archivo), así que el bot sólo lee las dos de estibas.
- 27 días (21/05–20/06), sin nulos. Sumas: `estibas_consumer` = 2.067,
  `estibas_profesional` = 1.293.

**Lógica a traducir (paso a paso):**
1. Lee `planta*` (de `CONSUMER/` y/o `PROFESIONAL/`): columnas `estibas_consumer`,
   `estibas_profesional` (`semana`/`fecha` se descartan).
2. El PQ unpivotea ambas en una columna `negocio` (`#"Columna de anulación de
   dinamización"`) y agrupa por `(fecha del nombre, negocio)` sumando → `valor`. El bot
   descarta la fecha del nombre (usa `periodo`) y suma por `negocio` sobre todos los
   `planta*`.
3. Servicios finales (uno por negocio con valor > 0): `negocio = negocio_facturador` =
   CONSUMER o PROFESIONAL; `servicio = "TRASLADO PALLETS PLANTA - CEDI"`; `valor` =
   sum(estibas_consumer) para CONSUMER, sum(estibas_profesional) para PROFESIONAL;
   `unidades = 0`, `proceso_extendido = "TRASLADO"`, `macro_proceso = "OTROS"`,
   `proceso_abreviado = "TRA"`, `tabla = "PLANTA"`. Convención del bot: se omiten los
   servicios con `valor = 0` (el PQ no los filtra).

**Dónde tocar al implementar** (molde: Paso 10 / `_run_trincaje_pipeline`, sin `start/end`):
- `config.py`: `PLANTA_GLOB = "planta*.[xX][lL][sS][xX]"` (sobre `DIRS["consumer"]` y
  `DIRS["profesional"]`), `PLANTA_COLS = {consumer:"estibas_consumer",
  profesional:"estibas_profesional"}`, `PLANTA_NEGOCIOS = {consumer:"CONSUMER",
  profesional:"PROFESIONAL"}`, `PLANTA_SERVICIO = "TRASLADO PALLETS PLANTA - CEDI"`.
- `processing/io_utils.py`: `find_planta_files(base_dir)` (glob `planta*` sobre AMBAS
  carpetas de negocio, dedup; devuelve solo rutas — sin área, pues `negocio` viene de la
  columna) y `read_planta(path)` (molde `read_etiquetas`, 2 cols numéricas).
- `processing/pipeline.py`: `_run_planta_pipeline(emit, progress)` → `list[dict]`;
  wirear en `run_all` **tras TRINCAJE** y sumar sus servicios antes de
  `_aggregate_servicios` (progreso 97; Excel sigue en 98); ampliar `__main__`.
- `processing/excel_export.py`: **sin cambios** (las columnas ya cubren PLANTA).

**Salida esperada** (rango 20/05–19/06/2026, ya calculada): **2 servicios PLANTA**
- CONSUMER / CONSUMER: `TRASLADO PALLETS PLANTA - CEDI` 2.067.
- PROFESIONAL / PROFESIONAL: `TRASLADO PALLETS PLANTA - CEDI` 1.293.

> Validación (verificada): `python -m processing.pipeline` da **62 líneas** en la hoja
> Servicios (las 60 de los pasos 1–10 + 2 de PLANTA) **sin alterar** los pasos previos
> y **sin issues**, y los valores (2.067 / 1.293) coinciden con la "Salida esperada".
> Sin `planta*` → advertencia y se omite (no detiene). Archivo obligatorio roto (sin
> `estibas_consumer`/`estibas_profesional`) → `BlockingError`.

---

## Paso 12 (Material) — ✅ HECHO

> Especificación (ya implementada). Traducción de la query `FilesOcupacionMaterial` de
> `logica.txt` (Power Query). **Gemelo SIMPLIFICADO de Ocupación (Paso 4)**: misma idea
> (promedio diario redondeado hacia arriba, `ceil`) pero con **una sola medida**
> (`cant_bodega_8`), servicio/negocio/nf **fijos** y **sin join EQUIVALENCIAS**. Lee
> `ocupacionMaterial_<fecha>.xlsx` de la carpeta **`OTROS/`** (junto a `otros_*`).

Facturación del **almacenamiento de material de empaque en Bodega 8 ME**. El valor es el
**promedio diario redondeado hacia arriba** (`ceil`) de `cant_bodega_8` en el rango.

**Adaptaciones (mismo patrón que Ocupación):**
1. Rango de fechas del **usuario** sobre la columna `fecha` (no `RANGOS_FECHAS` ni la fecha
   del nombre). El `ceil(promedio)` se calcula sobre los días que caen **dentro del rango**.
2. `negocio`/`negocio_facturador`/`servicio` son **constantes** del PQ (no se derivan del
   nombre ni del contenido).

**Datos verificados** (`ocupacionMaterial_01-06-2026.xlsx` 32 filas, 2 columnas, en `OTROS/`):
- Columnas: `fecha, cant_bodega_8`. El PQ sólo usa ésas (la `fecha` del NOMBRE se usa sólo
  para el join con `RANGOS_FECHAS`, que el bot reemplaza por el rango del calendario).
- 32 días (21/05–21/06), sin nulos. `cant_bodega_8` entero positivo.

**Lógica a traducir (paso a paso):**
1. Lee `ocupacionMaterial*` (de `OTROS/`): columnas `fecha, cant_bodega_8`.
2. **Filtra `fecha` por el rango del usuario** (inclusivo).
3. `valor = ceil(promedio(cant_bodega_8))` sobre los días en rango
   (`Number.RoundUp(List.Average([valor]))` del PQ).
4. Servicio final (**una línea**): `negocio = "MATERIAL DE EMPAQUE"`,
   `negocio_facturador = "MATERIAL DE EMPAQUE"` (en el PQ original venía sin "DE",
   distinto a `negocio`; aquí se unifica con "DE" por decisión del usuario),
   `servicio = "ALMACENAMIENTO PALLET BODEGA 8 ME GENERAL"`,
   `unidades = 0`, `proceso_extendido = "PALLETS"`, `macro_proceso = "ALMACENAMIENTO"`,
   `proceso_abreviado = "WHS"`, `tabla = "MATERIAL"`. Convención del bot: si `valor = 0` se
   omite.

**Dónde tocar al implementar** (molde: Paso 4 / `_run_ocupacion_pipeline`, simplificado):
- `config.py`: `DIRS["otros"] = BASE_DIR / "OTROS"`,
  `MATERIAL_GLOB = "ocupacionMaterial*.[xX][lL][sS][xX]"` (sobre `DIRS["otros"]`),
  `MATERIAL_COLS = {fecha:"fecha", valor:"cant_bodega_8"}`, `MATERIAL_NEGOCIO = "MATERIAL DE
  EMPAQUE"`, `MATERIAL_NEGOCIO_FACTURADOR = "MATERIAL DE EMPAQUE"`, `MATERIAL_SERVICIO =
  "ALMACENAMIENTO PALLET BODEGA 8 ME GENERAL"`.
- `processing/io_utils.py`: `find_material_files(base_dir)` (glob `ocupacionMaterial*` sobre
  `DIRS["otros"]`; devuelve solo rutas) y `read_material(path)` (molde `read_ocupacion`,
  2 cols `fecha`/`cant_bodega_8`).
- `processing/pipeline.py`: `_run_material_pipeline(start, end, ts_start, ts_end, emit,
  progress)` → `list[dict]`; wirear en `run_all` **tras PLANTA** y sumar sus servicios antes
  de `_aggregate_servicios`; ampliar `__main__`.
- `processing/excel_export.py`: **sin cambios** (las columnas ya cubren MATERIAL).

**Salida esperada** (rango 20/05–19/06/2026, ya calculada): **1 servicio MATERIAL**
- MATERIAL DE EMPAQUE / MATERIAL DE EMPAQUE: `ALMACENAMIENTO PALLET BODEGA 8 ME GENERAL` 553
  (= `ceil` del promedio de 30 días en rango, avg 552.8).

> Validación (verificada): `python -m processing.pipeline` da **63 líneas** en la hoja
> Servicios (las 62 de los pasos 1–11 + 1 de MATERIAL) **sin alterar** los pasos previos
> y **sin issues**, y el valor (553) coincide con la "Salida esperada". Sin
> `ocupacionMaterial*` → advertencia y se omite (no detiene). 0 filas con `fecha` en el
> rango → advertencia y se omite. Archivo obligatorio roto (sin `fecha`/`cant_bodega_8`)
> → `BlockingError`. **Paridad Power BI:** el PQ filtra por la ventana de `RANGOS_FECHAS`
> (21/05–21/06, 32 días → `ceil` 551); el bot usa rango libre, así que con el baseline
> 20/05–19/06 (30 días) da 553 — para paridad exacta elegir 21/05–21/06 (mismo caveat que
> Ocupación/Exportaciones).

---

## Paso 13 (Falabella) — ✅ HECHO

> Especificación (ya implementada). Traducción de la query `falabella` de `logica.txt`
> (Power Query). Es, junto a TRASLADOS/TRINCAJE/ETIQUETAS/PLANTA, **una de las pipelines
> más simples de todas**: sin huellas, sin idh, sin equivalencias, sin cruce y **sin
> filtro de fecha** (el PQ no aplica `RANGOS_FECHAS`: solo agrupa por la fecha del
> **nombre** del archivo, que el bot descarta y reemplaza por `periodo`).

Facturación del **etiquetado para Falabella**. El `valor` es el **conteo de filas con
`Entrega <> null`** por `negocio` (`Table.RowCount(_)` de `logica.txt`). Se genera una
línea `ETIQUETADO (FALABELLA)` por negocio y se suma a la misma hoja **Servicios**.

**Adaptaciones (mismo patrón que TRASLADOS/TRINCAJE):**
1. **Sin filtro de fecha**: se procesan TODOS los `falabella*` (no se usa el rango del
   calendario ni `RANGOS_FECHAS` ni la fecha del nombre — molde DESTRUCCIÓN/TRASLADOS).
2. Área desde el prefijo del nombre: `falabella_cons*` → CONSUMER, `falabella_prof*` →
   PROFESIONAL (vía `AREA_DEFAULT`, **no** el literal `cons`/`profesional` del PQ — ver nota).

**Datos verificados** (`falabella_prof_01-06-2026.xlsx` 114 filas, 8 columnas, en `PROFESIONAL/`):
- Columnas: `DIGITO VERIFICADOR, NIT HENKEL, CONSECUTIVO ASIGNADO POR EL PROVEEDOR BLU LOGISTICS,
  NUMERACION, LARGO 18 DIGITOS, TIENDA, Entrega, MES`. El PQ **sólo conserva `Entrega`** (para
  el filtro `<> null`) y descarta el resto (la `fecha` final viene del **nombre** del archivo;
  `MES` no llega al output), así que el bot sólo lee `Entrega`.
- `Entrega`: 114/114 no nulas → `valor = 114`.

**Lógica a traducir (paso a paso):**
1. Lee `falabella_cons*`/`falabella_prof*`: columna `Entrega`.
2. **Filtra filas con `Entrega <> null`** (`#"Filas filtradas"` de `logica.txt`).
3. Cuenta las filas por `negocio` → ese conteo es `valor` (`#"Filas agrupadas"` del PQ:
   `Table.RowCount(_)` por `(fecha del nombre, negocio)`; el bot descarta la fecha del nombre
   y suma todos los archivos por `negocio`).
4. Servicios finales (uno por negocio con valor > 0): `servicio = proceso_extendido =
   "ETIQUETADO (FALABELLA"` (constante), `negocio_facturador = negocio` (`DuplicateColumn`),
   `unidades = 0`, `macro_proceso = "OTROS"`, `proceso_abreviado = "OTR"`, `tabla = "FALABELLA"`.
   Convención del bot: se omiten los servicios con `valor = 0` (el PQ no los filtra).

> **Nota sobre `negocio`/`negocio_facturador`:** la query divide el nombre por `_`
> (`falabella_<negocio>_<fecha>`) y sólo hace `ReplaceValue("prof" → "profesional")`,
> dejando `cons` literal (inconsistencia del PQ). El bot normaliza al estándar de toda la
> hoja Servicios vía `AREA_DEFAULT`: `cons` → **CONSUMER**, `prof` → **PROFESIONAL**
> (`negocio_facturador = negocio`). Si se quisiera el literal `cons`/`profesional` habría
> que decidirlo aparte — no es lo que conviene para el Excel.

**Dónde tocar al implementar** (molde: Paso 10 / `_run_trincaje_pipeline`, sin `start/end`):
- `config.py`: `FALABELLA_GLOBS = {"cons": "falabella_cons*...", "prof": "falabella_prof*..."}`,
  `FALABELLA_COLS = {entrega:"Entrega"}`, `FALABELLA_SERVICIO = "ETIQUETADO (FALABELLA)"`.
- `processing/io_utils.py`: `find_falabella_files(base_dir)` (globs `falabella_cons*`/`falabella_prof*`
  sobre `CONSUMER/`/`PROFESIONAL/`, molde `find_traslados_files`, devuelve `(ruta, área)`) y
  `read_falabella(path, area)` (molde `read_traslados`, 1 col `Entrega`).
- `processing/pipeline.py`: `_run_falabella_pipeline(emit, progress)` → `list[dict]`;
  wirear en `run_all` **tras MATERIAL** y sumar sus servicios antes de `_aggregate_servicios`
  (progreso 97; Excel sigue en 98); ampliar `__main__`.
- `processing/excel_export.py`: **sin cambios** (las columnas ya cubren FALABELLA).

**Salida esperada** (rango 20/05–19/06/2026, ya calculada): **1 servicio FALABELLA**
- PROFESIONAL / PROFESIONAL: `ETIQUETADO (FALABELLA)` 114 (= 114 filas con `Entrega` no nula).

> Validación (verificada): `python -m processing.pipeline` da **64 líneas** en la hoja
> Servicios (las 63 de los pasos 1–12 + 1 de FALABELLA) **sin alterar** los pasos previos
> y **sin issues**, y el valor (114) coincide con la "Salida esperada". Sin `falabella*`
> → advertencia y se omite (no detiene). Archivo obligatorio roto (sin `Entrega`)
> → `BlockingError`.

---

## Paso 14 (Otros) — ✅ HECHO

> Especificación (ya implementada). Traducción de la query `otros` de `logica.txt`
> (Power Query, líneas 1–41). **Distinto a todos los pasos anteriores**: el archivo
> `otros_*.xlsx` no trae datos crudos que haya que calcular, sino una **tabla de servicios
> pre-armados a mano** (cánones de arrendamiento, horas extras, cargues expo puntuales,
> descargues, sobre-costos, cotizaciones…). Ya viene con sus `valor`, `tarifa` y `costo`
> calculados por el usuario, así que el bot **no calcula nada**: los **anexa** a la hoja
> Servicios respetando lo que trae. El fichero `otros_<fecha>.xlsx` vive en **`OTROS/`**
> (junto a `ocupacionMaterial_*` del Paso 12).

**Adaptaciones (mismo patrón que TRASLADOS/ETIQUETAS/FALABELLA):**
1. **Sin filtro de fecha**: el PQ descarta la columna `fecha` del contenido (línea 31) y
   toma la fecha del **nombre** del archivo, que el bot descarta y reemplaza por `periodo`.
   Se anexan **todas** las filas del archivo.
2. **`tabla` = "OTROS" fija** para todos los registros (decisión del usuario): el archivo
   trae OTROS/ALM/OUB en su columna `tabla`, pero se **ignora** — en el Excel todas las filas
   de este paso llevan tabla `OTROS`.

**Datos verificados** (`otros_01-06-2026.xlsx`, 33 filas, `Sheet1`, en `OTROS/`):
- Columnas en el fuente: `negocio, negocio_facturador, servicio, valor, proceso_extendido,
  macro_proceso, proceso_abreviado, tabla, tarifa, costo, fecha`. El PQ **descarta `fecha`**
  (la final viene del nombre), así que el bot lee las otras 10.
- `tabla` en el fuente: OTROS (28), ALM (4), OUB (1) — **pero se ignora**: en el Excel todas
  llevan tabla `OTROS` (decisión del usuario).
- `negocio`: PROFESIONAL 20, CONSUMER 13. `negocio_facturador`: PROFESIONAL 20, CONSUMER 9,
  LAUNDRY 4.
- `valor`: 0 nulos, 0 ceros; **algunos decimales** (cánones prorrateados: 68.16, 5.73…),
  a diferencia del resto de pasos que son enteros.
- `tarifa`/`costo`: dtype `object` — mezcla de número (`62895.4147`) y **texto moneda**
  (`"$ 13.933.333"`, `"$   6.348.884"`); `io_utils._parse_dinero` los limpia a `float`
  (replicando el `type number` del PQ, línea 33).

**Lógica a traducir (paso a paso):**
1. Lee `otros*` de `OTROS/`: las 10 columnas (sin `fecha`).
2. Filtra filas con `negocio <> null` (líneas 30 y 39 del PQ).
3. `valor`/`tarifa`/`costo` → número (`_parse_dinero` para la moneda). Añade `unidades = 0`
   (línea 35 del PQ).
4. **No agrupa** (no hay `#"Filas agrupadas"`): **1 fila del archivo → 1 fila del Excel**.
5. Cada línea se mapea al esquema del Excel: `costo → costo_total`, `periodo` (del rango,
   `_periodo_for_range(end)`), `um = ""` (no aplica). **No se recalcula el costo**: en
   algunas filas `costo ≠ valor × tarifa` (p. ej. traslado de pallets con `valor=40` y
   `costo=13.933.333`) y se respeta el `costo` que viene. Se omiten las filas con `valor`
   nulo.

> **Inyección TRAS `_apply_tarifas` (clave):** los Pasos 2–13 se suman a `servicios` *antes*
> de `_aggregate_servicos` y pasan por `_apply_tarifas`. Los de **Otros NO pueden pasar por
> ahí**: (a) no existen en `tarifas.xlsx` → se eliminarían; (b) `_aggregate_servicos` fuerza
> `valor` a `int` → perdería los decimales de los cánones. Por eso `_run_otros_pipeline` se
> llama **después** de `_apply_tarifas` y se concatena directamente a `servicios`.

**Dónde tocar al implementar** (molde: Paso 13 / `_run_falabella_pipeline`, sin `start/end`):
- `config.py`: `OTROS_GLOB = "otros*.[xX]..."` (sobre `DIRS["otros"]`, ya existente),
  `OTROS_COLS = {negocio, negocio_facturador, servicio, valor, proceso_extendido,
  macro_proceso, proceso_abreviado, tarifa, costo}` (sin `tabla`), `OTROS_TABLA = "OTROS"`,
  `OTROS_REQUIRED` (cols obligatorias → `BlockingError` si faltan).
- `processing/io_utils.py`: `_parse_dinero(v)` (limpia `$`/puntos de miles/coma decimal),
  `read_otros(path)` (molde `read_etiquetas`; 9 cols — sin `fecha` ni `tabla`, parsea
  `valor`/`tarifa`/`costo`), `find_otros_files(base_dir)` (glob `otros*` sobre `DIRS["otros"]`).
- `processing/pipeline.py`: `_run_otros_pipeline(periodo, emit, progress)` → `list[dict]`;
  `_otros_servicio` fija `tabla = config.OTROS_TABLA`; wirear en `run_all` **TRAS
  `_apply_tarifas`** (`servicios = servicios + otros_services`); ampliar `__main__`
  (conteo `tabla == "OTROS"`).
- `processing/excel_export.py`: **sin cambios** (las columnas ya cubren OTROS; `valor`
  decimal es válido).

**Salida esperada** (rango 20/05–19/06/2026, ya calculada): **33 servicios** con tabla
`OTROS` (el archivo traía 28 OTROS + 4 ALM + 1 OUB, pero todos se vuelcan como OTROS).
Algunos: `CANON DE ARRENDAMIENTO DE OFICINAS` (CONSUMER 68.16,
PROFESIONAL 32.82), `ALMACENAMIENTO BODEGA 3` (ALM, 777/153/287), `Auxiliares Extra
Operación Junio` (OUB, 4), `Horas extras …` (4 líneas), `Re: [EXT] RE: 📦 Cotización BCP
Junio` (1). Costo total de los 33 ≈ **\$247.779.585**.

> Validación (verificada): `python -m processing.pipeline` da **80 líneas** en la hoja
> Servicios (las 47 de los pasos 1–13 con tarifa + 33 de OTROS) **sin alterar** los pasos
> previos y **sin issues**, el parsing de moneda es correcto (`"$ 13.933.333"` → 13933333),
> `valor` conserva decimales, `tabla` = **OTROS** para todas y el costo total da
> **\$1.675.172.793**.
> Sin `otros*` → advertencia y se omite (no detiene); los 47 servicios previos intactos.
> Archivo obligatorio roto (sin `negocio`/`servicio`/`valor`/`tabla`) → `BlockingError`.

---

## Tarifa → costo (dinero) — ✅ HECHO

Después de `_aggregate_servicios`, `_apply_tarifas` (en `processing/pipeline.py`) cruza cada
línea agregada con `tarifas.xlsx` y calcula el costo:

- **Cruce por `servicio`** (llave normalizada con `io_utils.normalize`). La **fecha no importa**
  (decisión del usuario): se arma un lookup **plano** `servicio → {um, tarifa}` desde la hoja
  activa de `tarifas.xlsx` (hoja `Sheet1`, una tarifa por servicio; la hoja `todos` es histórico
  y no se usa). Si el mismo `servicio` tuviera tarifas distintas, gana la última + aviso.
- **Columnas añadidas** a cada línea: `um`, `tarifa` (= `tarifas.valor`) y
  `costo_total = valor × tarifa` (redondeado a 4 decimales). `valor` sigue siendo la **cantidad**
  (no se renombra).
- **No aparecen en el Excel final** (decisión del usuario):
  - las líneas **sin tarifa** (servicio no encontrado en `tarifas.xlsx`) — se omiten
    **silenciosamente** (no se reportan: no impiden facturar; el panel de avisos es solo
    para errores que impidan facturar);
  - las líneas con **`valor = 0`**.
- **Sin mínimo facturable**: sólo `valor × tarifa` (las columnas `minima`/`minima_valor_subproceso`
  de `tarifas.xlsx` se ignoran por ahora).
- **Degradación elegante**: si falta `tarifas.xlsx`, se emite un aviso y se conservan todas las
  líneas (sin columnas de costo, sin filtrar) — un auxiliar ausente no vacía el Excel.

**Columnas del Excel** (`excel_export.SERVICIO_COLUMNS`):
`periodo, negocio, negocio_facturador, servicio, valor, unidades, um, tarifa, costo_total,
proceso_extendido, macro_proceso, proceso_abreviado, tabla`.

> **Brecha de cobertura de tarifas (data, no código):** en el baseline, de los 64 servicios
> calculados **47 encuentran tarifa** y **17 (9 servicios únicos) se eliminan** por no existir en
> `tarifas.xlsx` — son las variantes `PALLETS`/`UNIDADES`/`UND` donde tarifas sólo trae la versión
> `CAJAS`, más `CAJAS ESTANDAR LAUNDRY` y `DESTELLE UNIDADES CROSS DOCKING`. Para que facturen,
> hay que añadir esas tarifas a `tarifas.xlsx` (hoja `Sheet1`); al hacerlo se rellenan solas, sin
> tocar código. Los pasos 1–13 dejan **47 líneas** con tarifa; sumando las **33 de OTROS** (Paso 14,
> que se inyecta tras `_apply_tarifas` y no se filtra), el Excel del baseline queda con **80 líneas**,
> costo total ≈ **\$1.675.172.793**.

---

## Reglas de negocio y decisiones importantes

- **`negocio` vs `negocio_facturador`** (nombres alineados a `logica.txt`):
  - `negocio` = **área de origen** del archivo (`cons`→CONSUMER, `prof`→PROFESIONAL).
    Equivale al `negocio` del Power Query (derivado del área).
  - `negocio_facturador` = negocio por **material** vía `idh_especiales`
    (CONSUMER / LAUNDRY / NATTURA / PROFESIONAL / **MATERIAL DE EMPAQUE**), con default = `negocio`.
    En `logica.txt` la prioridad es `MATERIAL_EMPAQUE > idh > área` con una tabla aparte;
    aquí la marca `MATERIAL DE EMPAQUE` **vive en la columna `Negocio` del propio
    `idh_especiales`** (no hay archivo aparte), así que queda `idh > área`. En **Ingresos**,
    si `negocio_facturador = MATERIAL DE EMPAQUE`, entonces `negocio` también pasa a
    `MATERIAL DE EMPAQUE` (y se generan los servicios `… ME`).
- **`tipo_trabajo`** desde `adicionales*`: `Delivery`↔`ENTREGA`. Match → el `TIPO`
  del archivo (p. ej. **EXTRA E.S**); si no → **NORMAL** (`config.TIPO_TRABAJO_DEFAULT`).
- **`tipo_despacho`** desde `tipo_despacho.xlsx`: cliente↔`CEDI`. Match → **CROSS DOCKING**;
  si no → **ESTANDAR** (`config.TIPO_DESPACHO_DEFAULT`).
- **`AddServicio`** (~15 ramas) clasifica cada `(negocio_facturador, tipo_trabajo,
  tipo_despacho, atributo)` en un `servicio`. Cada grupo se desdinamiza en hasta 3
  medidas (**PICKING PALLETS / CAJAS / UNIDADES**); `valor` = la cantidad de esa línea
  (**no es dinero** — el cálculo con tarifas es un paso posterior).
- **`FilterServicio`** descarta: servicios `null` y dos exclusiones —
  `DESTELLE CAJAS CROSS DOCKING` para LAUNDRY, y
  `ALISTAMIENTO Y DESPACHO UNIDADES ESTANDAR EXTRA E.S` para no-LAUNDRY.
- **Llaves de cruce**: entregas vía `io_utils.entrega_key` (int/float → texto sin `.0`);
  CEDI/cliente vía `io_utils.normalize` (mayúsculas, sin acentos ni puntuación).
- **`Material` aparece duplicado** en los salidas (col 16 y col 66). Se usa la
  **1ª ocurrencia** (col 16).
- **Emparejamiento de columnas por nombre normalizado** (sin acentos, mayúsculas,
  solo alfanum) para sobrevivir a re-exportaciones. La lectura usa un **camino rápido**
  por posición (`SALIDAS_FAST_POSITIONS`) con verificación por nombre y **fallback
  robusto** (lectura completa + nombre) si el layout cambia.
- **`huellas.xlsx`**: encabezados en la **fila 2** (`header=1`). Tiene ~26 `Producto`
  duplicados; se desempata quedándose con el **primer valor único que tenga dato correcto**
  (pallet y caja presentes y > 0), en el orden del archivo — `io_utils.read_huellas` ordena
  las filas válidas primero y hace `drop_duplicates(keep="first")`. La primera aparición de
  cada producto es la huella buena; un batch erróneo appended al final (p. ej. el 2025-10-11
  con valores absurdos como 900000/477000000) **no** la pisa. Afecta a TODOS los pasos que
  cruzan huellas (SALIDAS, INGRESOS, MAQUILA, EXPORTACIONES). El archivo se relee cada
  ejecución (se actualiza todos los meses). (El archivo real es una foto por producto; la
  lógica de `logica.txt` asumía capacidad por `(idh, fecha)` — ver "Diferencias con Power BI".)
- Se **ignoran** los archivos `~$*.xlsx` (bloqueo de Excel abierto).

---

## Arquitectura / estructura de archivos

```
app.py                 # FastAPI + uvicorn (delgado): monta /api y sirve static/
iniciar.bat            # lanzador de doble clic
requirements.txt
config.py              # rutas, mapa de columnas, constantes (sin deps web)
processing/
  io_utils.py          # lectura robusta de Excel (calamine), normalizador + lectores de lookup
  pipeline.py          # núcleo: SALIDAS + Destrucción + Ingresos + Ocupación + Traslados +
                       #        Maquila + Exportaciones + Etiquetas + Paletizado + Trincaje +
                       #        Planta + Material + Falabella + Otros, AddServicio/FilterServicio,
                       #        agrupación por servicio, BlockingError + RunState
                       #        (avance/cronómetro). Puro, sin FastAPI.
  excel_export.py      # escritura del Excel de salida (hoja Servicios)
api/
  routes.py            # router FastAPI (JSON + export)
static/
  index.html, app.css, app.js   # UI mínima (calendario, Generar, Descargar, errores)
FACTURAS_GENERADAS/    # salida de los Excel generados
logica.txt             # lógica original en Power Query (referencia / roadmap)
```

**Principio:** `processing/` no importa nada de FastAPI. La capa web (`api/`, `app.py`)
solo transporta entradas/salidas. Los pasos siguientes extienden `pipeline.py`.

### Endpoints (API)

| Método | Path | Devuelve |
|---|---|---|
| GET | `/api/tarifas` | tarifas (referencia; para pasos siguientes) |
| GET | `/api/daterange/default` | min/max de `Fecha factura` para precargar el calendario |
| GET | `/api/validate` | checklist de archivos presentes/faltantes en todas las carpetas (antes de facturar) |
| POST | `/api/run` | arranca el proceso en segundo plano (los 14 pasos) |
| GET | `/api/progress` | `{stage, percent, done, blocked, error, issues, has_result, elapsed_seconds}` (sondeo) |
| GET | `/api/export?start=&end=` | `.xlsx` (hoja **Servicios**: los 14 pasos combinados) |

Contrato uniforme: `{ok:true,data}` / `{ok:false,error,detail}`.

---

## Detección de errores en las fuentes

**Validación previa (checklist):** `GET /api/validate` (`pipeline.validate_sources()`)
escanea todas las carpetas y devuelve, por fuente, si está presente (✅), si falta y es
opcional (⚠️ — ese paso se omite con aviso) o si falta y es requerido (⛔ — detiene la
facturación; sólo Salidas). Sólo verifica **presencia** (no lee contenidos): rápido y
tolerante. La UI lo muestra en la tarjeta **Fuentes de datos** y lo auto-ejecuta al cargar
la página (botón *Validar fuentes* para repetirlo).

El procesamiento distingue dos tipos de problemas:
- **Errores graves** (archivo corrupto, mal formato, falta un archivo obligatorio, una
  columna requerida, o **datos inválidos en celdas clave** — ver *Auditoría de tipos* más
  abajo) **detienen el proceso** y avisan al usuario para que corrija y vuelva a generar.
  La barra de avance se congela en la etapa donde ocurrió.
- **Advertencias** (lookups opcionales ausentes como `huellas`/`idh`/`adicionales`/
  `tipo_despacho`, duplicados, registros sin huella, archivos `destruccion*` vacíos)
  **se reportan pero no detienen**.

`/api/progress` incluye `issues` con `severity` (`error` / `warning`):

- No hay archivos `salidas_cons*`/`salidas_prof*`.
- Un archivo no se pudo leer (corrupto, sin la columna `Material`, hoja rara).
- 0 filas con `Fecha factura` en el rango.
- No hay archivos `ingresos_cons*`/`ingresos_prof*` (Paso 3 omitido) o 0 filas con
  `Posting Date` en el rango.
- Falta `idh_especiales.xlsx` (todo queda por defecto).
- Falta `huellas.xlsx` (estibas/cajas vacías).
- Materiales sin huella (sus estibas/cajas quedan vacías).
- Materiales duplicados en `idh_especiales`.
- Falta `adicionales*` (todo queda `tipo_trabajo = NORMAL`).
- Falta `tipo_despacho.xlsx` (todo queda `tipo_despacho = ESTANDAR`).
- Entregas duplicadas en `adicionales` o CEDI duplicados en `tipo_despacho`.
- **Auditoría de tipos** (fecha invertida como `02/21/2026`, texto en columna numérica
  como `s`, o fechas inválidas/no parseables): **detiene el proceso** con detalle
  `archivo · columna · ejemplo` (ver abajo).

### Auditoría de tipos de datos (fechas y números)

El bot **no se traga silenciosamente** valores inválidos en las columnas que usa para
filtrar por fecha o para calcular. Al leer cada archivo audita las columnas clave
**antes** del `errors="coerce"` (que es donde un valor raro se volvería `NaN`/`NaT` sin
avisar). Si encuentra alguno, **detiene la facturación** (error grave) y lista **todos**
los problemas de **todas** las fuentes a la vez, con `archivo · columna · ejemplo`, para
que corrijas y vuelvas a generar.

Detecta (helper `io_utils.audit_value_column`, determinista vía `calendar.monthrange` —
no depende de la heurística inestable de `dayfirst` de pandas):

- **Fechas** (`Fecha factura`, `Posting Date`, `Fecha`, `fecha` de Material):
  **formato invertido mes/día** (`02/21/2026` → avisa; `13/02/2026` dd/mm válido **no**
  avisa), **inválidas/no parseables** (`31/02/2026`, `abc`) y basura. Los `datetime`
  nativos de Excel (celda fecha real) se aceptan sin auditar. `"05/06/2026"` (ambiguo)
  **no** se flaggea.
- **Números** (`Ctd Ent.(UMV)`, `Cantidad`, `Ocupación`, `SHU`, `CON`, `TOTAL`, `cajas`,
  `cant_bodega_8`, `estibas_consumer/profesional`, `valor`): texto que no es número
  (`s`). Los vacíos legítimos se ignoran.
- **Dinero** (`tarifa`/`costo` de Otros): admite texto moneda `$ 13.933.333`; sólo pita
  lo que `_parse_dinero` no puede convertir.

Excepciones para evitar falsos positivos: `"-"` en `Ocupación` (= 0) es válido;
`"$ 13.933.333"` en Otros es válido.

**Mecánica:** cada `_run_*_pipeline` audita al leer y, si hay error, salta el cálculo
(early-return) para no procesar basura; `run_all` sigue al siguiente paso para acumular
**todos** los problemas y, antes de generar el Excel, lanza `BlockingError` con el conteo
→ la UI muestra el panel rojo ⛔ con cada problema. Desde consola,
`python -m processing.pipeline` imprime `PROCESO DETENIDO` + los detalles
(archivo/columna/ejemplo).

---

## Rendimiento

- **`python-calamine`** lee Excel ~6x más rápido que openpyxl. Si no está instalado,
  la app cae a openpyxl (más lento pero funcional).
- **Generar** corre en segundo plano (SALIDAS + Destrucción, ~30–40 s; leer los salidas
  grandes es lo pesado). La UI **sondea `/api/progress`** y muestra solo el **avance**, un
  **cronómetro** (tiempo total al acabar) y los **problemas** (sin cifras). Si hay un error
  grave, se **detiene** y lo indica.
- Al **Generar** se pre-construye el Excel en **segundo plano**; por eso **Descargar**
  es prácticamente instantáneo. El resultado se cachea por rango `start|end`.
- Los cálculos (estibas/cajas) están **vectorizados** con numpy.

### Línea base validada (regresión) — rango 20/05–19/06/2026
101.491 filas · estibas 101.852 · cajas 237.822 · **64 líneas de servicio** de los pasos 1–13
(línea base interna; **+33 de OTROS** = 97 calculadas → **80 en el Excel** tras el filtro de
tarifa, costo total ≈ **\$1.675.172.793**):
- **SALIDAS**: 19 líneas agrupadas por `servicio`.
- **DESTRUCCION**: 2 líneas (CONSUMER 24, PROFESIONAL 3).
- **INGRESOS**: 6 líneas — `RECIBO CAJAS` (CONSUMER/CONSUMER 304.672, CONSUMER/LAUNDRY 3.997,
  PROFESIONAL/PROFESIONAL 8.384, PROFESIONAL/NATTURA 111.018) y, para `MATERIAL DE EMPAQUE`,
  `RECIBO CAJAS ME` 183 + `RECIBO PALLETS ME` 19 (unidades 65.700).
- **OCUPACION**: 10 líneas — `ALMACENAMIENTO …` (CONSUMER/CONSUMER: MEDIO PALLET 110, PALLET
  BODEGA 7 3.748, PALLET BODEGA 8 187 [=130+57 ME]; CONSUMER/LAUNDRY: PALLET BODEGA 8 758;
  PROFESIONAL/NATTURA: BIN 798; PROFESIONAL/PROFESIONAL: FLOW RACK 141, MEDIO PALLET 336,
  MODULA 1.051, PALLET BODEGA 7 1.272, PALLET BODEGA 8 31 [=6+25 ME]).
- **TRASLADOS**: 4 líneas — `ALISTAMIENTO Y DESPACHO … CENTRO DE TRASLADOS` (CONSUMER/CONSUMER:
  CAJAS 71, UNIDADES 489; PROFESIONAL/PROFESIONAL: CAJAS 125, UNIDADES 29.949).
- **MAQUILA**: 5 líneas — `ALISTAMIENTO DE MAQUILA CAJAS` (CONSUMER/CONSUMER 40.754,
  PROFESIONAL/NATTURA 1.223, PROFESIONAL/PROFESIONAL 7.410) y, para `MATERIAL DE EMPAQUE`,
  `PICKING CAJAS MAQUILA ME` 895 + `PICKING PALLETS MAQUILA ME` 14.
- **EXPORTACIONES**: 9 líneas — `ALISTAMIENTO Y DESPACHO … EXPO <canal>` (CONSUMER/EX:
  PALLETS 67, CAJAS 4.709, UND 79.188; CONSUMER/IC: PALLETS 921, CAJAS 85.048, UND 1.071.900;
  PROFESIONAL/IC: PALLETS 2.467, CAJAS 82.642, UND 4.602.297).
- **ETIQUETAS**: 2 líneas — `ETIQUETAS …` (CONSUMER/CONSUMER: REGULARES 3.433, REEMPAQUE 247).
- **PALETIZADO**: 2 líneas — `PALETIZADO EXPO IC` (CONSUMER/CONSUMER 58, PROFESIONAL/PROFESIONAL 1.199).
- **TRINCAJE**: 1 línea — `DOBLE TRINCAJE` (CONSUMER/CONSUMER 54 = 27 filas × 2).
- **PLANTA**: 2 líneas — `TRASLADO PALLETS PLANTA - CEDI` (CONSUMER/CONSUMER 2.067, PROFESIONAL/PROFESIONAL 1.293).
- **MATERIAL**: 1 línea — `ALMACENAMIENTO PALLET BODEGA 8 ME GENERAL` (MATERIAL DE EMPAQUE/MATERIAL DE EMPAQUE 553 = `ceil` del promedio de 30 días).
- **FALABELLA**: 1 línea — `ETIQUETADO (FALABELLA)` (PROFESIONAL/PROFESIONAL 114 = 114 filas con `Entrega` no nula; sin `falabella_cons*` no hay línea CONSUMER).
- **OTROS**: 33 líneas **anexadas pre-armadas** de `OTROS/otros_*.xlsx` (todas con tabla
  `OTROS`; el archivo traía 28 OTROS + 4 ALM + 1 OUB pero se ignora)
  — no pasan por el filtro de tarifa (se inyectan tras `_apply_tarifas`); p. ej. `CANON DE
  ARRENDAMIENTO DE OFICINAS`, `ALMACENAMIENTO BODEGA 3`, `Auxiliares Extra Operación Junio`,
  `Horas extras …`, con `valor` decimal y `tabla` respetada. Costo ≈ 247,8M.

Solo CONSUMER: 18.943 filas · estibas 19.303 · cajas 148.377 (CONSUMER 18.286 + LAUNDRY 657).
PROFESIONAL: 61.036 + NATTURA 21.512.

> Estos totales son una **línea base interna de regresión** (reproducible entre
> ejecuciones del bot), **no** una validación cruzada contra la salida de Power BI.

---

## Roadmap (frente a `logica.txt`)

El usuario construye **paso a paso**:

1. ~~**Tarifa → valor (dinero)**: cruzar `servicio` con `tarifas.xlsx` y multiplicar por `valor`.~~
   **✓ HECHO.** `_apply_tarifas` cruza cada línea agregada por `servicio` con la hoja activa de
   `tarifas.xlsx` (la **fecha no importa**: lookup plano, una tarifa por servicio), añade `um`,
   `tarifa` y `costo_total = valor × tarifa` (redondeo a 4 decimales), y **elimina** del Excel
   las líneas sin tarifa y las de `valor = 0`. Sin mínimo facturable (sólo valor×tarifa). Ver
   sección "Tarifa → costo (dinero)".
2. ~~**`tipo_trabajo`** desde `ADICIONALES` (entrega → NORMAL / EXTRA E.S).~~ **✓ HECHO.**
3. ~~**`tipo_despacho`** desde `TIPO_DESPACHO` (CEDI → ESTANDAR / CROSS DOCKING).~~ **✓ HECHO.**
4. **`negocio_facturador`** con prioridad compuesta: `MATERIAL_EMPAQUE` > `idh` > `area`.
   **✓ RESUELTO VÍA `idh_especiales`:** la marca `MATERIAL DE EMPAQUE` está en la columna
   `Negocio` del propio `idh_especiales.xlsx` (no hace falta archivo aparte); prioridad
   efectiva `idh > área`.
5. ~~**`AddServicio`** (~15 ramas) + **`FilterServicio`**.~~ **✓ HECHO.**
6. ~~**Unpivot + agrupación** + columnas finales (`proceso_extendido`, `macro_proceso`,
   `proceso_abreviado`, `tabla`).~~ **✓ HECHO.** Nota: el output **agrupa por `servicio`**
   (suma `valor`/`unidades`), **no incluye `fecha`** como dimensión y añade `periodo` (primer
   día del último mes del rango); `tipo_trabajo`/`tipo_despacho` **no van al Excel** (sólo se
   usan internamente para derivar `servicio`) — ver "Diferencias con Power BI".
7. ~~**Destrucción** (Paso 2): `destruccion*` → línea `ALISTAMIENTO Y DESPACHO PALLETS
   DESTRUCCION` por `negocio` (conteo de filas), sumada a la hoja Servicios. Sin filtro
   de fecha (se procesan todos los archivos).~~ **✓ HECHO.**
8. ~~**Ingresos** (Paso 3): `ingresos*` → servicios `RECIBO CAJAS` / `RECIBO CAJAS ME` /
   `RECIBO PALLETS ME` por `(negocio, negocio_facturador)`, con `MATERIAL DE EMPAQUE` vía
   `idh_especiales`. Filtro de fecha del usuario sobre `Posting Date`. Sumado a la hoja
   Servicios.~~ **✓ HECHO.**
9. ~~**Ocupación (Paso 4)**: `ocupacion*` → servicios `ALMACENAMIENTO …` por `(negocio,
   negocio_facturador)`, valor = promedio diario redondeado hacia arriba, con join
   EQUIVALENCIAS y casos especiales MODULA × 350.3 / PROFESIONAL+BIN → NATTURA.~~ **✓ HECHO.**
   Detalle y salida esperada (10 servicios) en la sección "Paso 4 (Ocupacion) — ✅ HECHO".
10. ~~**Traslados (Paso 5)**: `traslados*` → servicios `ALISTAMIENTO Y DESPACHO … CENTRO DE
    TRASLADOS` por `negocio` (CAJAS = sum(SHU), UNIDADES = sum(CON)), sin filtro de fecha
    (se procesan todos). Sumado a la hoja Servicios.~~ **✓ HECHO.**
11. ~~**Maquila (Paso 6)**: `maquila*` → servicios `ALISTAMIENTO DE MAQUILA CAJAS` y
    `PICKING … MAQUILA ME` por `(negocio, negocio_facturador)`, con las mismas 6 medidas de
    recibo de Ingresos (`floor`/`mod`/`ceil`), filtro de fecha del usuario sobre `Posting Date`
    y `MATERIAL DE EMPAQUE` vía `idh_especiales`. Sumado a la hoja Servicios.~~ **✓ HECHO.**
    Detalle y salida esperada (5 servicios) en la sección "Paso 6 (Maquila) — ✅ HECHO".
12. ~~**Exportaciones (Paso 7)**: `exportacion*` → servicios `ALISTAMIENTO Y DESPACHO
    PALLETS/CAJAS/UND EXPO <canal>` por `(negocio, canal)`, con `pallets`/`cajas`/`unidades`
    (`unidades = unidades_caja × cajas`, redondeo hacia arriba a múltiplo de caja), filtro de
    fecha del usuario sobre `Fecha factura` y `negocio_facturador = negocio` (el `JoinIDHEspeciales`
    del PQ es vestigial). Sumado a la hoja Servicios.~~ **✓ HECHO.** Detalle y salida esperada
    (9 servicios) en la sección "Paso 7 (Exportaciones) — ✅ HECHO".
13. ~~**Etiquetas (Paso 8)**: `etiquetas*` (de `MAQUILA/`) → servicios `ETIQUETAS REGULARES`/
    `ETIQUETAS REEMPAQUE` por `servicio` (suma de `cajas`), `negocio = "CONSUMER"` fijo, sin
    filtro de fecha (el PQ agrupa por la fecha del nombre, que el bot descarta → `periodo`).
    Sumado a la hoja Servicios.~~ **✓ HECHO.** Detalle y salida esperada (2 servicios) en la
    sección "Paso 8 (Etiquetas) — ✅ HECHO".
14. ~~**Paletizado (Paso 9)**: `paletizado*` (de `EXPORTACIONES/`) → servicios `PALETIZADO EXPO
    <canal>` por `(negocio, canal)` (suma de `TOTAL` con `DESPACHO <> null`), `negocio` desde
    `AREA` (HENKEL.PF → PROFESIONAL, HENKEL.RT → CONSUMER), sin filtro de fecha. Sumado a la
    hoja Servicios.~~ **✓ HECHO.** Detalle y salida esperada (2 servicios) en la sección
    "Paso 9 (Paletizado) — ✅ HECHO".
15. ~~**Trincaje (Paso 10)**: `trincaje*` (de `EXPORTACIONES/`) → servicio `DOBLE TRINCAJE`
    (conteo de filas con `despacho <> null` × 2), `negocio = "CONSUMER"` fijo, sin filtro de
    fecha (el PQ agrupa por la fecha del nombre, que el bot descarta → `periodo`). Sumado a
    la hoja Servicios.~~ **✓ HECHO.** Detalle y salida esperada (1 servicio) en la sección
    "Paso 10 (Trincaje) — ✅ HECHO".
16. ~~**Planta (Paso 11)**: `planta*` (de `CONSUMER/`/`PROFESIONAL/`) → servicios `TRASLADO
    PALLETS PLANTA - CEDI` por `negocio` (suma de `estibas_consumer`/`estibas_profesional`),
    `negocio` desde el nombre de columna, sin filtro de fecha (el PQ agrupa por la fecha del
    nombre, que el bot descarta → `periodo`). Sumado a la hoja Servicios.~~ **✓ HECHO.**
    Detalle y salida esperada (2 servicios) en la sección "Paso 11 (Planta) — ✅ HECHO".
17. ~~**Material (Paso 12)**: `ocupacionMaterial*` (de `OTROS/`) → servicio `ALMACENAMIENTO
    PALLET BODEGA 8 ME GENERAL` (`ceil` del promedio diario de `cant_bodega_8` en el rango),
    `negocio = MATERIAL DE EMPAQUE` / `nf = MATERIAL DE EMPAQUE`, rango de fechas del usuario
    sobre `fecha`. Sumado a la hoja Servicios.~~ **✓ HECHO.** Detalle y salida esperada
    (1 servicio) en la sección "Paso 12 (Material) — ✅ HECHO".
18. ~~**Falabella (Paso 13)**: `falabella*` (de `CONSUMER/`/`PROFESIONAL/`) → servicio
    `ETIQUETADO (FALABELLA)` por `negocio` (conteo de filas con `Entrega <> null`),
    `negocio` desde el prefijo del nombre vía `AREA_DEFAULT` (CONSUMER/PROFESIONAL), sin
    filtro de fecha (el PQ agrupa por la fecha del nombre, que el bot descarta → `periodo`).
    Sumado a la hoja Servicios.~~ **✓ HECHO.** Detalle y salida esperada (1 servicio) en la
    sección "Paso 13 (Falabella) — ✅ HECHO".
19. ~~**Otros (Paso 14)**: anexar los servicios pre-armados de `OTROS/otros_*.xlsx` (cánones,
    horas extras, cargues expo, descargues, sobre-costos…) a la hoja Servicios. El archivo ya
    trae `valor`/`tarifa`/`costo`, así que **no se calcula nada**: se respetan tal cual (con
    `tabla` OTROS/ALM/OUB), sin filtro de fecha (se anexan todas las filas) y se inyectan
    **tras `_apply_tarifas`**.~~ **✓ HECHO.** Detalle y salida esperada (33 servicios) en la
    sección "Paso 14 (Otros) — ✅ HECHO".

### Diferencias con Power BI (a revisar si se busca paridad exacta)

- **Filtro de fechas**: Power BI usa `RANGOS_FECHAS` (cada archivo → ventana
  `fecha_inicial/fecha_final`, filtra `Fecha factura` ahí). El bot usa un **rango libre**
  del calendario. (`rangos_fechas_facturacion.xlsx` existe pero no se usa.)
- **Huellas**: `logica.txt` une por `(material, fecha)` sobre una huella con capacidad por
  fecha; el `huellas.xlsx` real es una **foto por producto** con duplicados, así que el bot
  se queda con el **primer valor correcto** de cada producto (orden del archivo). Evita que
  un batch corrupto appended al final (2025-10-11) pise la huella buena — esto era lo que
  hacía que `RECIBO PALLETS ME` diera 55 en vez de 63 (materiales 2184249 y 2472922).
- **`fecha` como dimensión**: Power BI agrupa también por la fecha del archivo; el bot no.
- **Forma del output**: Power BI **desdinamiza** las medidas en filas (`atributo`/`valor`);
  el bot lo hace para clasificar el servicio y luego **agrupa por `servicio`** (suma
  `valor`/`unidades`). `tipo_trabajo`/`tipo_despacho` **no van al Excel** (sólo internos) y se
  añade `periodo` (primer día del último mes del rango). Mismos números.
- **`valor`**: en Power BI es la cantidad de la línea; aquí igual. El **dinero** (tarifa ×
  valor) no está en `logica.txt` visible ni en el bot.
- **Destrucción**: Power BI agrupa por `(fecha del nombre, negocio)` y filtra por
  `RANGOS_FECHAS`; el bot **agrega por `negocio` todos los archivos** (sin fecha) — los
  archivos no traen fecha en su contenido, así que se procesan todos los encontrados.
- **Ingresos**: Power BI filtra por ventana de `RANGOS_FECHAS` (vía `fecha` del nombre) y
  agrupa por `(fecha, negocio, negocio_facturador)`; el bot **filtra `Posting Date` por el
  rango del usuario** y **agrupa por `(negocio, negocio_facturador)`** (sin `fecha`; usa
  `periodo`). La marca `MATERIAL DE EMPAQUE` viene de `idh_especiales.Negocio` (Power BI la
  tomaba de una tabla `MATERIAL_EMPAQUE` aparte). Mismos números por grupo.
- **Ocupación** (verificado 2026-07-11): Power BI filtra los días de almacenamiento por la
  ventana de `RANGOS_FECHAS` (joined por la **fecha del nombre** del archivo; `logica.txt`
  322–349). Para junio esa ventana es **21/05–21/06** (32 días). El bot filtra por el **rango
  libre del calendario**, y el **default del calendario** se calcula del min/máx de fechas
  de **TODAS** las fuentes (`default_date_range()` → `_collect_range_dates`: salidas, ingresos,
  ocupación, maquila, exportaciones, material), así que el *Hasta* por defecto **alcanza** el
  fin de la ventana de ocupación/material/exportaciones (21/06). Si el *Hasta* no llega al fin
  de la ventana, los servicios cuyo promedio roza un entero difieren de Power BI (junio con la
  ventana corta: `PALLET BODEGA 7` CONSUMER **3754** vs 3760, PROFESIONAL **1272** vs 1271;
  `PALLET BODEGA 8` LAUNDRY **762** vs 766); los demás coinciden porque su `ceil` no se mueve.
  **Paridad exacta:** con el *Hasta* por defecto (21/05–21/06 para junio) el bot da
  **3760 / 1271 / 766**, idéntico a Power BI. (Decisión del usuario: mantener rango libre;
  **no** aplicar `RANGOS_FECHAS` automáticamente.)
- **Traslados**: Power BI agrupa por `(fecha del nombre, negocio)` y filtra por la ventana de
  `RANGOS_FECHAS`. El bot **no filtra por fecha** (decisión del usuario, molde DESTRUCCIÓN):
  procesa **todos** los `traslados*` y agrega solo por `negocio`. Además, el bot **omite** los
  servicios con `valor = 0` (Power BI los conserva). Mismos números por negocio mientras llegue
  un solo archivo por periodo.
- **Maquila**: Power BI filtra por la ventana de `RANGOS_FECHAS` (joined por la fecha del
  **nombre** del archivo) sobre `Posting Date`, y agrupa por `(fecha, negocio, negocio_facturador)`.
  El bot filtra `Posting Date` por el **rango libre del calendario** y agrupa por
  `(negocio, negocio_facturador)` (sin `fecha`; usa `periodo`). La marca `MATERIAL DE EMPAQUE`
  viene de `idh_especiales.Negocio` (Power BI la tomaba de una tabla `MATERIAL_EMPAQUE` aparte).
  **`unidades`**: el `AddUnidadesServicio` del PQ tiene un bug de plantilla (compara contra
  nombres "RECIBO…" inexistentes → todo cae a `unidades_pallet_maquila`); el bot aplica un mapeo
  **coherente** (cajas→`unidades_generales`, cajas ME→`unidades_cajas_maquila`, pallets ME→
  `unidades_pallet_maquila`). Los `valor` son idénticos al PQ. `etiquetas*` no se procesa
  (`FilesMaquila` sólo admite ficheros que empiecen por `maquila`).
- **Exportaciones**: Power BI filtra por la ventana de `RANGOS_FECHAS` (joined por la fecha del
  **nombre** del archivo) sobre `Fecha factura`, y agrupa por `(fecha del nombre, negocio,
  negocio_facturador, canal)`. El bot filtra `Fecha factura` por el **rango libre del
  calendario** y agrupa por `(negocio, canal)` (sin `fecha`; usa `periodo`; `nf = negocio`).
  **`negocio_facturador`**: el `JoinIDHEspeciales` del PQ es vestigial (NestedJoin que nunca se
  expande) + `DuplicateColumn(negocio → nf)`, así que `nf = negocio = área` en ambos (sin
  LAUNDRY/NATTURA/MATERIAL DE EMPAQUE). **`unidades`** = `unidades_caja × cajas` (ceil a
  múltiplo de caja), idéntico al PQ. **Sin `FilterServicio`**: el PQ no filtra servicios; el
  bot omite `valor = 0` (convención). `paletizado*`/`trincaje*` no se procesan (`FilesExportacion`
  sólo admite ficheros que empiecen por `exportacion`). Mismos números por `(negocio, canal)`
  con un solo archivo por área; al igual que en ocupación, el baseline 20/05–19/06 queda un día
  corto frente a la ventana 21/05–21/06 de PBI (las filas de prof con `Fecha factura` 20/06
  quedan fuera) — para paridad exacta elegir 21/05–21/06.
- **Etiquetas**: el PQ agrupa por `(fecha del nombre, negocio=CONSUMER, servicio)` (sin
  `RANGOS_FECHAS`). El bot **descarta la fecha del nombre** (usa `periodo`) y agrega solo por
  `servicio` (suma `cajas`); `negocio = CONSUMER` se mantiene fijo. Además, el bot **omite** los
  servicios con `valor = 0` (Power BI los conservaría). Con un solo `etiquetas*` por periodo no
  hay diferencia; si llegaran varios, el bot los sumaría todos en una línea por servicio.
- **Paletizado**: el PQ agrupa por `(fecha del nombre, negocio, canal)` (sin `RANGOS_FECHAS`),
  con `negocio` desde `AREA` (HENKEL.PF/HENKEL.RT). El bot **descarta la fecha del nombre** (usa
  `periodo`) y agrega por `(negocio, canal)` (suma `TOTAL` con `DESPACHO <> null`). Además, el
  bot **omite** los servicios con `valor = 0` (Power BI los conservaría). Con un solo
  `paletizado*` por periodo no hay diferencia; si llegaran varios, el bot los sumaría todos en
  una línea por `(negocio, canal)`.
- **Trincaje**: el PQ agrupa por `(fecha del nombre)` (sin `RANGOS_FECHAS`) y hace
  `valor = Table.RowCount(_) * 2`, con `negocio = "CONSUMER"` fijo. El bot **descarta la fecha
  del nombre** (usa `periodo`) y suma todas las filas con `despacho <> null` de todos los
  `trincaje*`: `valor = (total filas) × 2` (idéntico con un solo archivo). Además, el bot
  **omite** el servicio si `valor = 0` (Power BI lo conservaría). `exp`/`destino`/`contenedor`/
  `n-cont` se leen en el PQ pero se descartan tras el filtro (sólo `despacho` importa).
- **Planta**: el PQ elimina las cols `fecha`/`semana` del archivo y agrupa por
  `(fecha del nombre, negocio)` (sin `RANGOS_FECHAS`), donde `negocio` viene de unpivotear
  `estibas_consumer`/`estibas_profesional` (no de la carpeta). El bot **descarta la fecha del
  nombre** (usa `periodo`) y suma por `negocio` sobre todos los `planta*`. Además, el bot
  **omite** los servicios con `valor = 0` (Power BI los conservaría). Con un solo `planta*` por
  periodo no hay diferencia; si llegaran varios, el bot los sumaría todos en una línea por
  negocio.
- **Material**: el PQ filtra los días por la ventana de `RANGOS_FECHAS` (joined por la fecha
  del NOMBRE del archivo, sobre la `fecha` del contenido) y hace `valor = ceil(promedio(
  cant_bodega_8))`, con `negocio = MATERIAL DE EMPAQUE` y `nf = MATERIAL EMPAQUE` (en el PQ
  venía sin "DE"; el bot lo unifica con "DE" por decisión del usuario) constantes. El bot
  filtra `fecha` por el **rango libre del calendario** (mismo modelo que
  Ocupación) y aplica el mismo `ceil(promedio)`. Al igual que en Ocupación/Exportaciones, el
  baseline 20/05–19/06 queda **un día corto** frente a la ventana 21/05–21/06 de PBI (30 vs 32
  días → 553 vs 551); para paridad exacta elegir 21/05–21/06. Además, el bot **omite** el
  servicio si `valor = 0` (Power BI lo conservaría).

- **Falabella**: el PQ agrupa por `(fecha del nombre, negocio)` (sin `RANGOS_FECHAS`) y hace
  `valor = Table.RowCount(_)` (conteo de filas con `Entrega <> null`), con `negocio` derivado
  del nombre split por `_` y un `ReplaceValue` parcial (`prof`→`profesional`, `cons` se queda
  literal). El bot **descarta la fecha del nombre** (usa `periodo`) y cuenta por `negocio` sobre
  todos los `falabella*`, **normalizando** `cons`/`prof` → CONSUMER/PROFESIONAL vía `AREA_DEFAULT`
  (en vez del literal `cons`/`profesional` del PQ). Además, el bot **omite** el servicio si
  `valor = 0` (Power BI lo conservaría). Con un solo `falabella*` por periodo y negocio no hay
  diferencia numérica; si llegaran varios, el bot los sumaría todos en una línea por negocio.

### Dónde tocar para añadir un paso
- Lógica: amplía `run_all(...)` (o añade `run_stepN(...)`) en `processing/pipeline.py`,
  reutilizando `io_utils` (añade ahí nuevos lectores, p. ej. `read_destruccion`).
- Columnas/posiciones de las fuentes: en `config.py`. Columnas del Excel de salida:
  `excel_export.SERVICIO_COLUMNS`.
- API: en `api/routes.py`. El proceso corre en un **hilo daemon**; sigue el patrón de
  `POST /api/run` + sondeo de `GET /api/progress` (+ `GET /api/export`).
- UI: `static/index.html` + `app.js` (+ `app.css` para la barra de avance).

---

## Notas para otro Claude

- Empieza leyendo `logica.txt` (Power Query) y este README. La lógica pura está en
  `processing/pipeline.py`; los tests manuales se corren con
  `python -m processing.pipeline` (imprime SALIDAS/DESTRUCCION y los `issues`).
- **Proceso asíncrono + bloqueo**: `run_all(...)` corre en un hilo daemon desde
  `POST /api/run`; el avance se sigue con `RunState` (`get_progress()`, lleva
  `elapsed_seconds`/cronómetro). Los **errores graves** lanzan `BlockingError` y **detienen**
  (archivo corrupto / mal formato / falta columna o archivo obligatorio); el resto son
  `issues` warning que **no detienen**. `default_date_range()` (precarga del calendario, vía
  `_collect_range_dates`) junta las fechas de **todas** las fuentes para que el *Hasta* por
  defecto cubra el fin de la ventana de ocupación/material/exportaciones (no sólo salidas).
- **Output**: una fila por `servicio` (agrupada por `_aggregate_servicios`); `periodo` =
  primer día del último mes del rango (`_periodo_for_range`). `tipo_trabajo`/`tipo_despacho`
  no van al Excel (sólo internos). Columnas: `excel_export.SERVICIO_COLUMNS`.
- **No rompas el camino rápido** de `read_salidas` (posiciones + verificación por nombre):
  calamine parsea el archivo entero en cada llamada, así que hay que minimizar parseos.
- Mantén la separación processing (sin web) ↔ api/static.
- Valida cualquier cambio contra la **línea base** de arriba.
