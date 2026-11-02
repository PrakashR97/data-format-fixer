"""Per-upload temporary workspaces. Each job gets its own private folder and an unguessable id."""
from __future__ import annotations

import re
import secrets
import shutil
import tempfile
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path

from .errors import UserError


def safe_stem(filename: str) -> str:
    stem = Path(filename.replace("\\", "/")).stem
    stem = re.sub(r"[^\w\-. ()]", "_", stem, flags=re.UNICODE).strip(" .")[:100]
    return stem or "file"


@dataclass
class Job:
    id: str
    dir: Path
    filename: str
    kind: str  # xlsx | csv
    created: float = field(default_factory=time.time)
    touched: float = field(default_factory=time.time)
    lock: threading.Lock = field(default_factory=threading.Lock)

    @property
    def source(self) -> Path:
        return self.dir / f"original.{self.kind}"

    @property
    def output(self) -> Path:
        return self.dir / "corrected.xlsx"

    @property
    def download_name(self) -> str:
        return f"{safe_stem(self.filename)}_corrected.xlsx"


class JobStore:
    def __init__(self, ttl_minutes: int, max_active: int):
        self.root = Path(tempfile.mkdtemp(prefix="data_format_fixer_"))
        self.ttl = ttl_minutes * 60
        self.max_active = max_active
        self.jobs: dict[str, Job] = {}
        self._lock = threading.Lock()

    def create(self, filename: str, kind: str) -> Job:
        with self._lock:
            if len(self.jobs) >= self.max_active:
                raise UserError("The server is busy. Please try again in a few minutes.", 503)
            job_id = secrets.token_urlsafe(24)
            d = self.root / job_id
            d.mkdir(mode=0o700)
            job = Job(job_id, d, filename, kind)
            self.jobs[job_id] = job
            return job

    def get(self, job_id: str) -> Job:
        with self._lock:
            job = self.jobs.get(job_id)
            if job:
                job.touched = time.time()
        if not job:
            raise UserError("This file is no longer available. Please upload it again.", 404)
        return job

    def delete(self, job_id: str) -> None:
        with self._lock:
            job = self.jobs.pop(job_id, None)
        if job:
            shutil.rmtree(job.dir, ignore_errors=True)

    def purge_expired(self) -> None:
        cutoff = time.time() - self.ttl
        with self._lock:
            old = [j for j in self.jobs.values() if j.touched < cutoff]
        for j in old:
            self.delete(j.id)

    def close(self) -> None:
        shutil.rmtree(self.root, ignore_errors=True)
