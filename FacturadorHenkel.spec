# -*- mode: python ; coding: utf-8 -*-
# Spec de PyInstaller para el Facturador Henkel (modo ONEDIR: una carpeta con el .exe).
#
# Construcción:  pyinstaller FacturadorHenkel.spec --noconfirm
# Resultado:     dist/FacturadorHenkel/FacturadorHenkel.exe
#
# El cliente recibe la carpeta dist/FacturadorHenkel/ y deja sus Excels en las
# subcarpetas (CONSUMER/, PROFESIONAL/, ...) AL LADO del .exe (config.BASE_DIR =
# carpeta del exe cuando está empaquetado). Los archivos estáticos van DENTRO del
# bundle (servidos desde _MEIPASS).
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

# uvicorn importa dinámicamente loops/protocols/lifespan -> recolectar todos sus
# submódulos para que no falten en tiempo de ejecución.
hidden = collect_submodules("uvicorn") + collect_submodules("starlette")
hidden += [
    "h11",            # backend HTTP de uvicorn
    "openpyxl",       # lectura/escritura Excel (calamine no está instalado -> openpyxl)
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

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="FacturadorHenkel",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,           # ventana visible: el cliente la cierra para detener el servidor
    disable_windowed_traceback=False,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="FacturadorHenkel",
)
