"""Settings read from environment variables (optionally from a root .env file)."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parents[2]
load_dotenv(ROOT / ".env")


def _int(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, default))
    except ValueError:
        return default


@dataclass
class Settings:
    max_upload_mb: int = 25
    max_uncompressed_mb: int = 300
    max_rows: int = 200_000
    max_columns: int = 500
    max_cells: int = 2_000_000
    max_sheets: int = 50
    max_processing_seconds: int = 60
    job_ttl_minutes: int = 30
    max_active_jobs: int = 50
    cors_origins: list[str] = field(
        default_factory=lambda: ["http://localhost:5173", "http://127.0.0.1:5173"]
    )

    @classmethod
    def from_env(cls) -> "Settings":
        origins = os.environ.get("CORS_ORIGINS", "http://localhost:5173,http://127.0.0.1:5173")
        return cls(
            max_upload_mb=_int("MAX_UPLOAD_MB", 25),
            max_uncompressed_mb=_int("MAX_UNCOMPRESSED_MB", 300),
            max_rows=_int("MAX_ROWS", 200_000),
            max_columns=_int("MAX_COLUMNS", 500),
            max_cells=_int("MAX_CELLS", 2_000_000),
            max_sheets=_int("MAX_SHEETS", 50),
            max_processing_seconds=_int("MAX_PROCESSING_SECONDS", 60),
            job_ttl_minutes=_int("JOB_TTL_MINUTES", 30),
            max_active_jobs=_int("MAX_ACTIVE_JOBS", 50),
            cors_origins=[o.strip() for o in origins.split(",") if o.strip()],
        )
