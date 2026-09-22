import os
from urllib.parse import quote_plus

from dotenv import load_dotenv
from sqlalchemy import create_engine

load_dotenv()

_engine = None


def get_engine():
    """Lazily create a SQLAlchemy engine for the checkpoint DB from env vars
    (DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_NAME). See .env.example."""
    global _engine
    if _engine is None:
        user = os.environ["DB_USER"]
        password = quote_plus(os.environ["DB_PASSWORD"])
        host = os.environ["DB_HOST"]
        port = os.environ.get("DB_PORT", "3306")
        name = os.environ["DB_NAME"]
        _engine = create_engine(
            f"mysql+mysqlconnector://{user}:{password}@{host}:{port}/{name}"
        )
    return _engine
