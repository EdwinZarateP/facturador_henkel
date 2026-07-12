@echo off
REM Facturador Henkel - lanzador de doble clic.
REM Levanta el servidor local. El navegador lo abre app.py (1 sola pestaña,
REM cuando el servidor ya esta listo), asi que aqui NO se abre el navegador.
cd /d "%~dp0"
title Facturador Henkel
color 0F
echo.
echo  ================================================================
echo                      FACTURADOR HENKEL
echo  ================================================================
echo.
echo   Iniciando servidor local...
echo   El navegador se abrira solo en http://127.0.0.1:8000
echo.
echo   ----------------------------------------------------------------
echo    IMPORTANTE: deja esta ventana abierta mientras lo usas.
echo    Para detener el Facturador: cierra esta ventana o presiona Ctrl+C.
echo   ----------------------------------------------------------------
echo.

REM Verificar que Python este disponible.
where python >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] No se encontro 'python'.
    echo          Instala Python 3.x o agrega su carpeta al PATH y vuelve a intentarlo.
    echo.
    pause
    exit /b 1
)

REM Si el puerto 8000 ya esta en uso, uvicorn lo indicara en esta ventana.
python app.py

echo.
echo  ================================================================
echo   El Facturador se detuvo. Puedes cerrar esta ventana.
echo  ================================================================
pause
