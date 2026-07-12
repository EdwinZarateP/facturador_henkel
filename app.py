"""Servidor del Facturador Henkel (FastAPI + uvicorn).

Arranque:  python app.py   (o doble-clic en iniciar.bat)
Navegador: http://127.0.0.1:8000

Sirve el frontend (static/) en / y la API en /api.
"""
from __future__ import annotations

import os
import sys
import threading
import webbrowser
from pathlib import Path

import uvicorn
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

from api.routes import router

# Si la app corre empaquetada (PyInstaller), los archivos estáticos (HTML/CSS/JS/logo)
# están dentro del bundle (carpetas extraídas desde _MEIPASS). En desarrollo, junto a app.py.
if getattr(sys, "frozen", False):
    STATIC_DIR = Path(sys._MEIPASS) / "static"
else:
    STATIC_DIR = Path(__file__).resolve().parent / "static"
HOST = os.environ.get("FACTURADOR_HOST", "127.0.0.1")
PORT = int(os.environ.get("FACTURADOR_PORT", "8000"))


def create_app() -> FastAPI:
    app = FastAPI(title="Facturador Henkel")

    # No cachear los archivos estáticos (index.html/app.js/app.css): garantiza que el
    # navegador siempre cargue la última versión tras un cambio (local, sin versionado).
    @app.middleware("http")
    async def _no_cache_static(request, call_next):
        response = await call_next(request)
        path = request.url.path.lower()
        if path == "/" or path.endswith((".html", ".js", ".css")):
            response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
            response.headers["Pragma"] = "no-cache"
            response.headers["Expires"] = "0"
        return response

    app.include_router(router, prefix="/api")
    app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
    return app


app = create_app()


def _open_browser():
    webbrowser.open(f"http://{HOST}:{PORT}")


if __name__ == "__main__":
    # Abrir el navegador 1.5s después de arrancar el servidor.
    threading.Timer(1.5, _open_browser).start()
    uvicorn.run(app, host=HOST, port=PORT, log_level="info")
