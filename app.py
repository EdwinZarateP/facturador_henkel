"""Servidor del Facturador Henkel (FastAPI + uvicorn).

Arranque:  python app.py   (o doble-clic en iniciar.bat)
Navegador: http://127.0.0.1:8000

Sirve el frontend (static/) en / y la API en /api.
"""
from __future__ import annotations

import threading
import webbrowser
from pathlib import Path

import uvicorn
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

from api.routes import router

STATIC_DIR = Path(__file__).resolve().parent / "static"
HOST = "127.0.0.1"
PORT = 8000


def create_app() -> FastAPI:
    app = FastAPI(title="Facturador Henkel")
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
