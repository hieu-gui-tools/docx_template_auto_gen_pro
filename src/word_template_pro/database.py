"""Database models và engine cho Word Template Pro."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from sqlalchemy import (
    Boolean, Column, DateTime, Integer, String, Text, create_engine, event
)
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker

# ── DB path ──────────────────────────────────────────────────────────────────
APP_DIR = Path(__file__).parent.parent.parent          # project root
DB_PATH = APP_DIR / "word_template_pro.db"
ENGINE = create_engine(f"sqlite:///{DB_PATH}", echo=False)


class Base(DeclarativeBase):
    pass


# ── Models ───────────────────────────────────────────────────────────────────

class Template(Base):
    __tablename__ = "danh_sach_template"

    id            = Column(Integer, primary_key=True, autoincrement=True)
    name          = Column(String(200), nullable=False)
    file_path     = Column(String(500), nullable=False, unique=True)
    output_folder = Column(String(500), default="")
    filename_pattern = Column(String(300), default="{template_name}_{date}")
    created_at    = Column(DateTime, default=datetime.now)
    updated_at    = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __repr__(self):
        return f"<Template id={self.id} name={self.name!r}>"


class Settings(Base):
    __tablename__ = "settings"

    key   = Column(String(100), primary_key=True)
    value = Column(Text, default="")


# ── Helpers ──────────────────────────────────────────────────────────────────

def get_session() -> Session:
    SessionLocal = sessionmaker(bind=ENGINE)
    return SessionLocal()


def init_db() -> None:
    Base.metadata.create_all(ENGINE)
    # seed default settings
    with get_session() as s:
        for k, v in {
            "default_output_folder": str(APP_DIR / "output"),
            "theme": "dark",
        }.items():
            if not s.get(Settings, k):
                s.add(Settings(key=k, value=v))
        s.commit()
    # Chuẩn hoá path cũ (từ các phiên bản trước)
    migrate_paths()


def get_setting(key: str, default: str = "") -> str:
    with get_session() as s:
        row = s.get(Settings, key)
        return row.value if row else default


def set_setting(key: str, value: str) -> None:
    with get_session() as s:
        row = s.get(Settings, key)
        if row:
            row.value = value
        else:
            s.add(Settings(key=key, value=value))
        s.commit()


def migrate_paths() -> None:
    """Chuẩn hoá tất cả file_path trong DB về dạng absolute backslash (Windows)."""
    with get_session() as s:
        rows = s.query(Template).all()
        changed = 0
        for row in rows:
            try:
                norm = str(Path(row.file_path).resolve())
                if norm != row.file_path:
                    row.file_path = norm
                    changed += 1
            except Exception:
                pass
        if changed:
            s.commit()
