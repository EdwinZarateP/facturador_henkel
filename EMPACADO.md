# Empaquetado y entrega (PyInstaller)

Cómo generar el ejecutable y armar el paquete que se le entrega al cliente.
El cliente **no instala Python**: doble-clic y se abre en el navegador.

---

## 1. Generar el ejecutable

```bash
pyinstaller FacturadorHenkel.spec --noconfirm
```

- Resultado: `dist/FacturadorHenkel/` (modo **carpeta / onedir**, ~113 MB).
  Contiene `FacturadorHenkel.exe` + `_internal/` (librerías + `static/`).
- `static/` (HTML/CSS/JS/logo) va **dentro** del bundle (se sirve desde `_MEIPASS`).
- Los **datos** (Excels del cliente) van **junto al exe** (carpeta donde está el `.exe`).
- `python-calamine` **sí se empaqueta** → lectura rápida de Excel (~6× vs openpyxl).

> Si cambias `app.py` / `config.py` / `static/`, **reconstruye** con el mismo comando.

---

## 2. Armar la carpeta entregable

Copia `dist/FacturadorHenkel/` a una carpeta de entrega (p. ej. `Facturador Henkel/`)
y añade **al lado del `.exe`**:

```
Facturador Henkel/
├─ FacturadorHenkel.exe        ← viene de dist/
├─ _internal/                  ← viene de dist/ (librerías; NO tocar)
├─ iniciar.bat                 ← lanzador de doble clic (ver §3)
├─ tarifas.xlsx                ← AUXILIARES/tarifas.xlsx (decide si lo entregas)
├─ CONSUMER/                   ← vacías: el cliente deja aquí sus Excels
├─ PROFESIONAL/
├─ AUXILIARES/                 ← (idh, tipo_despacho, equivalencias… si aplican)
├─ HUELLAS/                    ← huellas.xlsx
├─ MAQUILA/
├─ EXPORTACIONES/
├─ OTROS/
└─ MANUAL DE USUARIO.pdf       ← (pendiente de redactar)
```

Notas:
- Las carpetas de datos pueden ir **vacías** (con un `LEEME.txt` que explique qué archivo
  va en cada una). El cliente arrastra ahí los Excels que le lleguen.
- `FACTURAS_GENERADAS/` **se crea sola** al generar la primera factura (junto al exe).
- Sin `tarifas.xlsx` no salen los costos (sí las cantidades).

---

## 3. `iniciar.bat` para el ejecutable

Pon este `iniciar.bat` junto al `.exe` (doble-clic lo levanta y abre el navegador):

```bat
@echo off
cd /d "%~dp0"
title Facturador Henkel
color 0F
echo.
echo   ================================================================
echo                      FACTURADOR HENKEL
echo   ================================================================
echo.
echo   Iniciando... el navegador se abrira solo en http://127.0.0.1:8000
echo   Deja esta ventana abierta mientras lo usas.
echo   Para detener: cierra esta ventana o presiona Ctrl+C.
echo.
echo   ----------------------------------------------------------------
FacturadorHenkel.exe
echo.
echo   El Facturador se detuvo. Puedes cerrar esta ventana.
echo   ----------------------------------------------------------------
pause
```

> Si el puerto 8000 estuviera ocupado en la máquina del cliente, se puede cambiar
> con la variable de entorno `FACTURADOR_PORT` (p. ej. `set FACTURADOR_PORT=8011`
> antes de lanzar el exe). El host, con `FACTURADOR_HOST`.

---

## 4. Verificación rápida del ejecutable

```bash
# Levantarlo en un puerto libre, apuntando a los datos del proyecto:
FACTURADOR_PORT=8011 FACTURADOR_BASE_DIR="<raiz del proyecto>" ./dist/FacturadorHenkel/FacturadorHenkel.exe
# En otra terminal:
curl http://127.0.0.1:8011/api/validate      # -> checklist de fuentes
curl http://127.0.0.1:8011/                  # -> la pagina (logo + UI)
```

Verificado (2026-07-11): `/`, `/logo.webp`, `/app.js`, `/api/validate` y
`/api/daterange/default` responden 200 con datos correctos.

---

## 5. Pendientes / Known issues

- **Carga del calendario lenta (~16–28s):** `GET /api/daterange/default` lee **todos** los
  archivos para calcular el rango (los salidas tienen ~100k filas). Mejora propuesta:
  cachear el resultado por *mtime* de las fuentes (instantáneo si los archivos no
  cambiaron). Es una optimización pendiente, no un error del empaquetado.
- **Manual de usuario** (español plano): pendiente de redactar.
- **Licencia / contrato de uso**: definir qué se entrega (uso vs. fuente), soporte y
  prohibición de reventa.
- **Calamine en el exe**: confirmado empaquetado (`_python_calamine.cp313-*.pyd`).
