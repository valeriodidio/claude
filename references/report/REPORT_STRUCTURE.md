# Struttura Generica Report — Python Backend + ASP

Questo documento descrive l'architettura standard da replicare per ogni nuovo report
dell'admin Yeppon. Segui questa struttura per mantenere coerenza tra i progetti.

---

## Struttura cartelle

```
python_backend/
├── .env                        ← credenziali + token (NON in git)
├── requirements.txt
├── venv/                       ← creato localmente, NON in git
├── app/
│   ├── __init__.py
│   ├── main.py                 ← FastAPI entry point + middleware token
│   ├── config.py               ← Settings (pydantic-settings, legge .env)
│   ├── db.py                   ← Engine SQLAlchemy singleton
│   ├── routers/
│   │   ├── __init__.py
│   │   └── <nome_report>.py   ← endpoints HTTP del report
│   ├── services/
│   │   ├── __init__.py
│   │   ├── <nome>_query.py    ← logica SQL / costruzione filtri
│   │   └── excel_export.py    ← generazione XLSX (se serve)
│   └── static/
│       └── test_<nome>.html   ← pagina HTML di test (solo dev, disabilitata in prod)
└── dev/
    ├── docker-compose.yml      ← MySQL locale
    ├── init.sql                ← schema + utente reports
    ├── start-local.bat         ← avvio ambiente locale
    ├── stop-local.bat          ← stop container
    ├── reset-db.bat            ← reset completo volume MySQL
    ├── seed_from_prod.py       ← importa dati reali da produzione
    ├── scripts/
    │   └── seed_test_data.py   ← genera dati fittizi per test
    └── .env.prod.example       ← template credenziali prod (copia in .env.prod)

asp_nuovo/
└── <nome_report>/
    ├── <nome_report>.asp           ← pagina principale (HTML + JS + chiamate AJAX)
    └── download-<nome>-xlsx.asp    ← proxy per download binario Excel (se serve)

deploy/
├── deploy.ps1      ← script deploy staging → produzione
├── watcher.ps1     ← watcher auto-deploy (opzionale)
└── README.md       ← istruzioni deploy
```

---

## config.py — pattern standard

```python
from functools import lru_cache
from pathlib import Path
from typing import List
from pydantic_settings import BaseSettings, SettingsConfigDict

_ENV_FILE = Path(__file__).resolve().parent.parent / ".env"

class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=str(_ENV_FILE),
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    # DB
    db_host: str = "localhost"
    db_port: int = 3306
    db_user: str = "reports_reader"
    db_password: str = ""
    db_name: str = "yeppon_stats"

    # App
    app_host: str = "127.0.0.1"
    app_port: int = 8000
    app_reload: bool = False

    # Security
    internal_token: str = ""
    allowed_origins: str = ""

    # Dev-only
    enable_static_test_ui: bool = False

    log_level: str = "INFO"

    @property
    def mysql_url(self) -> str:
        return (
            f"mysql+pymysql://{self.db_user}:{self.db_password}"
            f"@{self.db_host}:{self.db_port}/{self.db_name}?charset=utf8mb4"
        )

    @property
    def cors_origins(self) -> List[str]:
        if not self.allowed_origins.strip():
            return []
        return [o.strip() for o in self.allowed_origins.split(",") if o.strip()]

@lru_cache
def get_settings() -> Settings:
    return Settings()
```

---

## db.py — engine singleton

```python
from sqlalchemy import create_engine
from sqlalchemy.engine import Engine
from app.config import get_settings

_engine: Engine | None = None

def get_engine() -> Engine:
    global _engine
    if _engine is None:
        settings = get_settings()
        _engine = create_engine(
            settings.mysql_url,
            pool_size=5,
            max_overflow=10,
            pool_pre_ping=True,
            pool_recycle=1800,
            future=True,
        )
    return _engine
```

---

## main.py — entry point e middleware

```python
import logging
from pathlib import Path
from fastapi import FastAPI, Request, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from app.config import get_settings
from app.routers import nome_report   # ← importa il tuo router

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

app.include_router(nome_report.router, prefix="/api/reports/nome_report", tags=["nome_report"])

if settings.enable_static_test_ui:
    _STATIC_DIR = Path(__file__).resolve().parent / "static"
    if _STATIC_DIR.is_dir():
        app.mount("/static", StaticFiles(directory=str(_STATIC_DIR)), name="static")

@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    log.exception("Unhandled error on %s", request.url)
    return JSONResponse(status_code=500, content={"detail": "Internal server error"})
```

---

## Router pattern (routers/<nome>.py)

```python
from fastapi import APIRouter, Depends, Query
from app.db import get_engine
from app.services import nome_query as q

router = APIRouter()

def _filters(
    dal: date = Query(...),
    al: date = Query(...),
    # ... altri filtri
) -> q.Filters:
    return q.Filters(dal=dal, al=al, ...)

@router.get("/summary")
def get_summary(filters: q.Filters = Depends(_filters)):
    engine = get_engine()
    return q.query_summary(engine, filters)

@router.get("/export.xlsx", response_class=Response)
def export_xlsx(filters: q.Filters = Depends(_filters)):
    engine = get_engine()
    data = q.query_for_export(engine, filters)
    xlsx_bytes = build_xlsx(data)
    filename = f"report_{filters.dal}_{filters.al}.xlsx"
    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
```

---

## .env template

```ini
# Connessione DB (locale: punta al Docker MySQL)
DB_HOST=127.0.0.1
DB_PORT=3307          # porta mappata da docker-compose
DB_USER=reports_reader
DB_PASSWORD=dev_password_change_me
DB_NAME=yeppon_stats

# Security — in produzione metti una stringa lunga e casuale
# In locale può restare vuoto (controllo disabilitato)
INTERNAL_TOKEN=

# Dev: espone /static/test_<nome>.html tramite FastAPI
# NON attivare in produzione
ENABLE_STATIC_TEST_UI=true

LOG_LEVEL=DEBUG
```

---

## init.sql template

```sql
-- Eseguito automaticamente al primo avvio del container MySQL (docker-compose)

CREATE DATABASE IF NOT EXISTS yeppon_stats
    DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

USE yeppon_stats;

CREATE TABLE IF NOT EXISTS nome_tabella (
    id          BIGINT AUTO_INCREMENT PRIMARY KEY,
    campo1      VARCHAR(50),
    campo2      DOUBLE,
    dataordine  TIMESTAMP,
    INDEX idx_dataordine (dataordine)
);

-- Se serve accedere a un DB esterno (es. smart2):
CREATE DATABASE IF NOT EXISTS smart2
    DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

USE smart2;
CREATE TABLE IF NOT EXISTS ordini_cliente (
    id       INT PRIMARY KEY,
    idUtente INT NOT NULL,
    INDEX idx_id_utente (idUtente)
);

USE yeppon_stats;

-- Utente di sola lettura usato dall'app
CREATE USER IF NOT EXISTS 'reports_reader'@'%' IDENTIFIED BY 'dev_password_change_me';
GRANT SELECT ON yeppon_stats.nome_tabella   TO 'reports_reader'@'%';
GRANT SELECT ON smart2.ordini_cliente       TO 'reports_reader'@'%';
FLUSH PRIVILEGES;
```

> **Attenzione**: `init.sql` viene eseguito **solo al primo avvio** del volume Docker.
> Se aggiungi tabelle o GRANT dopo aver già creato il volume, devi applicarli
> manualmente con `docker exec` oppure eseguire `reset-db.bat` per ripartire da zero.

---

## Checklist nuovo report

- [ ] Crea cartella `python_backend/` con struttura sopra
- [ ] Crea `dev/init.sql` con le tabelle necessarie
- [ ] Configura `dev/docker-compose.yml` (porta MySQL unica per evitare conflitti)
- [ ] Scrivi `app/services/<nome>_query.py` con i filtri e le query
- [ ] Scrivi `app/routers/<nome>.py` con gli endpoint
- [ ] Registra il router in `app/main.py`
- [ ] Crea `asp_nuovo/<nome>/<nome>.asp` seguendo `REPORT_ASP_STRUCTURE.md`
- [ ] Testa localmente con `dev/start-local.bat`
- [ ] Deploy con `deploy/deploy.ps1` seguendo `REPORT_DEPLOY.md`

---

## Produzione — avvio NSSM

```powershell
# Installa il servizio Windows (una tantum)
nssm install YepponReportsAPI "D:\admin_yeppon_python\apps\yeppon_reports\python_backend\venv\Scripts\python.exe"
nssm set YepponReportsAPI AppParameters "-m uvicorn app.main:app --host 127.0.0.1 --port 8000"
nssm set YepponReportsAPI AppDirectory "D:\admin_yeppon_python\apps\yeppon_reports\python_backend"
nssm set YepponReportsAPI AppStdout "D:\admin_yeppon_python\logs\uvicorn_stdout.log"
nssm set YepponReportsAPI AppStderr "D:\admin_yeppon_python\logs\uvicorn_stderr.log"
nssm set YepponReportsAPI Start SERVICE_AUTO_START
nssm start YepponReportsAPI
```
