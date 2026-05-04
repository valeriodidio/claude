"""
Entry point FastAPI.
Avvio in produzione (Windows Service via NSSM):
    python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
"""
import logging
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles

from app.config import get_settings
from app.routers import turnover, pl_prodotti

settings = get_settings()

logging.basicConfig(
    level=settings.log_level,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
log = logging.getLogger("reports")

app = FastAPI(
    title="Yeppon Reports API",
    version="0.1.0",
    docs_url="/api/reports/docs",
    openapi_url="/api/reports/openapi.json",
)

if settings.cors_origins:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.cors_origins,
        allow_credentials=True,
        allow_methods=["GET"],
        allow_headers=["*"],
    )


@app.middleware("http")
async def internal_token_guard(request: Request, call_next):
    """
    Barriera minima: l'autenticazione utente vera e' sull'ASP.
    Se INTERNAL_TOKEN e' vuoto, il controllo e' disabilitato.
    """
    path = request.url.path
    if path in ("/api/reports/health", "/api/reports/docs", "/api/reports/openapi.json"):
        return await call_next(request)
    if settings.enable_static_test_ui and path.startswith("/static/"):
        return await call_next(request)
    if settings.internal_token:
        token = request.headers.get("X-Internal-Token") or request.query_params.get("token", "")
        if token != settings.internal_token:
            return JSONResponse(status_code=401, content={"detail": "Invalid internal token"})
    return await call_next(request)


@app.get("/api/reports/health")
def health():
    return {"status": "ok"}


# ── Router ────────────────────────────────────────────────────────────────────
app.include_router(turnover.router,    prefix="/api/reports/turnover",    tags=["turnover"])
app.include_router(pl_prodotti.router, prefix="/api/reports/pl_prodotti", tags=["pl_prodotti"])


# ── Pagina di test statica (solo dev) ─────────────────────────────────────────
if settings.enable_static_test_ui:
    _STATIC_DIR = Path(__file__).resolve().parent / "static"
    if _STATIC_DIR.is_dir():
        app.mount("/static", StaticFiles(directory=str(_STATIC_DIR)), name="static")
        log.warning(
            "Pagina di test statica ABILITATA su /static/. "
            "Non attivare ENABLE_STATIC_TEST_UI in produzione."
        )


@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    log.exception("Unhandled error on %s", request.url)
    return JSONResponse(status_code=500, content={"detail": "Internal server error"})
