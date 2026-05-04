"""Settings caricati da .env via pydantic-settings."""
from functools import lru_cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict

_ENV_FILE = Path(__file__).resolve().parent.parent / ".env"


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=str(_ENV_FILE), extra="ignore")

    # Database
    db_host: str = "localhost"
    db_port: int = 3306
    db_user: str = "reports_reader"
    db_password: str = ""
    db_name: str = "smart2"

    # App
    app_host: str = "127.0.0.1"
    app_port: int = 8000
    app_reload: bool = False

    # Sicurezza
    internal_token: str = ""           # token condiviso con l'ASP
    allowed_origins: str = ""          # CORS, domini separati da virgola
    enable_static_test_ui: bool = False
    log_level: str = "INFO"

    # Default report
    default_periodo_giorni: int = 180

    @property
    def mysql_url(self) -> str:
        return (
            f"mysql+pymysql://{self.db_user}:{self.db_password}"
            f"@{self.db_host}:{self.db_port}/{self.db_name}?charset=utf8mb4"
        )

    @property
    def cors_origins(self) -> list[str]:
        return [o.strip() for o in self.allowed_origins.split(",") if o.strip()]


@lru_cache
def get_settings() -> Settings:
    return Settings()
