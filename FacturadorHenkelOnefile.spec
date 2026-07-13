# -*- mode: python ; coding: utf-8 -*-
# Spec de PyInstaller para el Facturador Henkel (modo ONEFILE: un solo .exe).
#
# Construcción:  pyinstaller FacturadorHenkelOnefile.spec --noconfirm
# Resultado:     dist/FacturadorHenkel.exe  (un ÚNICO archivo autoextraíble)
#
# A diferencia del spec ONEDIR (FacturadorHenkel.spec), aquí NO hay carpeta
# _internal ni COLLECT: todo (binarios + datos estáticos) va EMPAQUETADO dentro
# del propio .exe. Al ejecutarse, PyInstaller lo extrae a un temporal (_MEIxxxx
# bajo %TEMP%) y arranca la app desde ahí.
#
# Ventaja para distribución: el cliente descarga UN solo archivo; no hay _internal
# que se pierda al comprimir/subir a una compartida ni DLLs sueltas que el
# antivirus pueda borrar una a una. Trade-off: arranque +3-5s (extracción a temp
# en cada ejecución). Los Excels del cliente van JUNTO al .exe
# (config.BASE_DIR = carpeta de sys.executable).
#
# El spec ONEDIR (FacturadorHenkel.spec) se conserva por si se quiere volver a
# esa modalidad (arranque más rápido, carpeta con _internal).
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

# uvicorn importa dinámicamente loops/protocols/lifespan -> recolectar todos sus
# submódulos para que no falten en tiempo de ejecución.
hidden = collect_submodules("uvicorn") + collect_submodules("starlette")
hidden += [
    "h11",             # backend HTTP de uvicorn
    "openpyxl",        # lectura/escritura Excel (fallback si calamine fallara)
    "python_calamine", # lectura Excel ~6x más rápida (calamine SÍ está en el entorno)
    "numpy",
    "pandas",
    "webbrowser",
]

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("static", "static"),   # index.html, app.css, app.js, logo.webp
    ],
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # No se usan -> reducen tamaño del bundle.
        "tkinter", "matplotlib", "PyQt5", "PyQt6", "PySide2", "PySide6",
        "scipy", "pytest", "IPython", "notebook", "jupyter",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ONEFILE: el EXE incluye TODO (scripts + binaries + zipfiles + datas). No hay
# COLLECT ni exclude_binaries=True.
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="FacturadorHenkel",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,           # ventana visible: el cliente la cierra para detener el servidor
    disable_windowed_traceback=False,
    icon=None,
)
