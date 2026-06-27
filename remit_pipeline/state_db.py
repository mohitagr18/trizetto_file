"""
State Database Module - SQLite wrapper to track processed remittance files.
"""

import logging
import sqlite3
from pathlib import Path
from typing import Set, Tuple

logger = logging.getLogger(__name__)


class StateDB:
    """Manages SQLite database state for read-only remittance tracking."""

    def __init__(self, db_path: Path):
        self.db_path = db_path
        self._conn = None
        self._init_db()

    def _connect(self) -> sqlite3.Connection:
        """Establish connection to SQLite database."""
        if not self._conn:
            self.db_path.parent.mkdir(parents=True, exist_ok=True)
            self._conn = sqlite3.connect(str(self.db_path))
            self._conn.row_factory = sqlite3.Row
        return self._conn

    def close(self) -> None:
        """Close connection."""
        if self._conn:
            self._conn.close()
            self._conn = None

    def _init_db(self) -> None:
        """Initialize schema if it doesn't exist."""
        conn = self._connect()
        cursor = conn.cursor()
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS processed_files (
                filename TEXT PRIMARY KEY,
                file_size INTEGER,
                downloaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                processed_at TIMESTAMP,
                status TEXT
            );
            """
        )
        conn.commit()

    def get_processed_files(self) -> Set[Tuple[str, int]]:
        """
        Get set of (filename, file_size) tuples of successfully processed files.
        """
        conn = self._connect()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT filename, file_size FROM processed_files WHERE status = 'SUCCESS'"
        )
        return {(row["filename"], row["file_size"]) for row in cursor.fetchall()}

    def mark_file(self, filename: str, file_size: int, status: str) -> None:
        """
        Record file download/process result in database.
        Updates on conflict to allow retry of failed files.
        """
        conn = self._connect()
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO processed_files (filename, file_size, status, processed_at)
            VALUES (?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(filename) DO UPDATE SET
                file_size = excluded.file_size,
                status = excluded.status,
                processed_at = CURRENT_TIMESTAMP
            """,
            (filename, file_size, status),
        )
        conn.commit()
        logger.info("Marked file '%s' (%d bytes) as %s", filename, file_size, status)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
