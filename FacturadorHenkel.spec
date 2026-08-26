# -*- mode: python ; coding: utf-8 -*-
# Spec de PyInstaller para el Facturador Henkel (modo ONEFILE: un único .exe).
#
# Construcción:  pyinstaller FacturadorHenkel.spec --noconfirm
# Resultado:     dist/FacturadorHenkel.exe  (un solo archivo, ~50 MB)
#
# El cliente recibe SOLO el .exe y deja sus Excels en las subcarpetas
# (CONSUMER/, PROFESIONAL/, ...) AL LADO del .exe (config.BASE_DIR = carpeta del
# exe cuando está empaquetado). Los archivos estáticos van DENTRO del bundle
# (servidos desde _MEIPASS).
#
# Nota: ONEFILE descomprime el bundle a temp en cada arranque -> la primera
# apertura tarda 1-2 minutos (pandas/numpy pesan); es normal.
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
    a.binaries,
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
