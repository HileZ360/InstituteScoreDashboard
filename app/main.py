from pathlib import Path

from fastapi import FastAPI
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from .data_loader import DataStore

BASE_DIR = Path(__file__).resolve().parents[1]
WEB_DIR = BASE_DIR / "web"
ASSET_DIR = WEB_DIR / "assets"

app = FastAPI(title="Institute Score Dashboard")
store = DataStore(BASE_DIR)


@app.get("/api/data")
def api_data(force: bool = False) -> JSONResponse:
    return JSONResponse(store.load(force=force))


@app.post("/api/refresh")
def api_refresh() -> JSONResponse:
    return JSONResponse(store.load(force=True))


@app.get("/api/health")
def api_health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/")
def index() -> FileResponse:
    return FileResponse(WEB_DIR / "index.html")


if ASSET_DIR.exists():
    app.mount("/assets", StaticFiles(directory=ASSET_DIR), name="assets")
