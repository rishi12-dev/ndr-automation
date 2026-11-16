from pathlib import Path
from sqlalchemy import create_engine, text
from sqlalchemy.orm import DeclarativeBase, sessionmaker
from .config import settings

if settings.database_url.startswith("sqlite"):
    Path("data").mkdir(exist_ok=True)
    engine = create_engine(settings.database_url, connect_args={"check_same_thread": False})
else:
    engine = create_engine(settings.database_url, pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, expire_on_commit=False)

class Base(DeclarativeBase):
    pass

def init_db():
    from . import models
    Base.metadata.create_all(engine)
    if settings.database_url.startswith("sqlite"):
        with engine.begin() as connection:
            columns = {row[1] for row in connection.execute(text("PRAGMA table_info(orders)"))}
            for name in ("dx_mail", "cx_mail"):
                if name not in columns: connection.execute(text(f"ALTER TABLE orders ADD COLUMN {name} VARCHAR(255) DEFAULT ''"))
            rto_columns = {row[1] for row in connection.execute(text("PRAGMA table_info(rto_cases)"))}
            if "store_received_date" not in rto_columns: connection.execute(text("ALTER TABLE rto_cases ADD COLUMN store_received_date DATE"))
