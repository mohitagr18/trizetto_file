"""
Configuration module — loads settings from .env file.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env from project root
load_dotenv(Path(__file__).resolve().parent.parent / ".env")


class Config:
    """Central configuration loaded from environment variables."""

    # SFTP
    SFTP_HOST: str = os.getenv("SFTP_HOST", "sftp.gatewayedi.com")
    SFTP_PORT: int = int(os.getenv("SFTP_PORT", "22"))
    SFTP_USERNAME: str = os.getenv("SFTP_USERNAME", "")
    SFTP_PASSWORD: str = os.getenv("SFTP_PASSWORD", "")
    SFTP_REMOTE_DIR: str = os.getenv("SFTP_REMOTE_DIR", "remit/")

    # Local paths
    RAW_DATA_DIR: Path = Path(os.getenv("RAW_DATA_DIR", "./data/raw/"))
    PROCESSED_DATA_DIR: Path = Path(os.getenv("PROCESSED_DATA_DIR", "./data/processed/"))
    OUTPUT_DIR: Path = Path(os.getenv("OUTPUT_DIR", "./output/"))

    # Behavior
    FILE_EXISTS_BEHAVIOR: str = os.getenv("FILE_EXISTS_BEHAVIOR", "skip")  # "skip" or "overwrite"

    @classmethod
    def ensure_directories(cls) -> None:
        """Create local directories if they don't exist."""
        cls.RAW_DATA_DIR.mkdir(parents=True, exist_ok=True)
        cls.PROCESSED_DATA_DIR.mkdir(parents=True, exist_ok=True)
        cls.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    @classmethod
    def validate(cls) -> None:
        """Validate that required config values are present."""
        if not cls.SFTP_USERNAME:
            raise ValueError("SFTP_USERNAME is not set. Check your .env file.")
        if not cls.SFTP_PASSWORD:
            raise ValueError("SFTP_PASSWORD is not set. Check your .env file.")
