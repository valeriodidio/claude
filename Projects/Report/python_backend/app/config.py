"""
Config caricata da .env (con python-dotenv + pydantic-settings).
"""
from functools import lru_cache
from pathlib import Path
from typing import List

from pydantic_settings import BaseSettings, SettingsConfigDict

# Percorso assoluto al .env: funziona indipendentemente da dove viene
# lanciato uvicorn (dalla root del progetto, da python_backend, ecc.)
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

    # Dev-only: espone /static/test_turnover.html tramite FastAPI.
    # In produzione DEVE restare False: l'UI e' servita dall'admin ASP.
    # Per debug locale o rapido in staging, si puo' abilitare mettendo
    # ENABLE_STATIC_TEST_UI=true nel .env.
    enable_static_test_ui: bool = False

    # Logging
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
