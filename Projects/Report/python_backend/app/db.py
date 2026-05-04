"""
Engine SQLAlchemy singleton con pool di connessioni MySQL.
"""
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
            pool_pre_ping=True,   # riconnette se il DB ha chiuso la connessione
            pool_recycle=1800,    # ricicla ogni 30 min per evitare wait_timeout
            future=True,
        )
    return _engine
