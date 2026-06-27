"""
SFTP client module — connect, list, and download .rmt files.
"""

import logging
import os
from pathlib import Path
from typing import List, Optional, Tuple

import paramiko

from remit_pipeline.config import Config

logger = logging.getLogger(__name__)


class SFTPClient:
    """Manages SFTP connection and file operations for EDI 835 remittance files."""

    def __init__(self):
        self._transport: Optional[paramiko.Transport] = None
        self._sftp: Optional[paramiko.SFTPClient] = None

    def connect(self) -> None:
        """Establish SFTP connection using credentials from Config."""
        Config.validate()
        logger.info(
            "Connecting to SFTP: %s:%d as %s",
            Config.SFTP_HOST,
            Config.SFTP_PORT,
            Config.SFTP_USERNAME,
        )
        self._transport = paramiko.Transport((Config.SFTP_HOST, Config.SFTP_PORT))
        self._transport.connect(
            username=Config.SFTP_USERNAME,
            password=Config.SFTP_PASSWORD,
        )
        self._sftp = paramiko.SFTPClient.from_transport(self._transport)
        logger.info("SFTP connection established successfully.")

    def disconnect(self) -> None:
        """Close the SFTP connection."""
        if self._sftp:
            self._sftp.close()
        if self._transport:
            self._transport.close()
        logger.info("SFTP connection closed.")


    def list_rmt_files(self, remote_dir: Optional[str] = None) -> List[str]:
        """
        List all .rmt files in the remote directory.

        Args:
            remote_dir: Remote directory path. Defaults to Config.SFTP_REMOTE_DIR.

        Returns:
            List of .rmt filenames found in the directory.
        """
        if self._sftp is None:
            raise RuntimeError("Not connected. Call connect() first.")

        remote_dir = remote_dir or Config.SFTP_REMOTE_DIR
        logger.info("Listing files in remote directory: %s", remote_dir)

        all_files = self._sftp.listdir(remote_dir)
        rmt_files = sorted([f for f in all_files if f.lower().endswith(".rmt")])

        logger.info("Found %d .rmt files in %s", len(rmt_files), remote_dir)
        return rmt_files

    def list_rmt_files_with_attrs(self, remote_dir: Optional[str] = None) -> List[Tuple[str, int]]:
        """
        List all .rmt files in the remote directory along with their file sizes.

        Args:
            remote_dir: Remote directory path. Defaults to Config.SFTP_REMOTE_DIR.

        Returns:
            List of (filename, file_size) tuples found in the directory.
        """
        if self._sftp is None:
            raise RuntimeError("Not connected. Call connect() first.")

        remote_dir = remote_dir or Config.SFTP_REMOTE_DIR
        logger.info("Listing files with attributes in remote directory: %s", remote_dir)

        attrs = self._sftp.listdir_attr(remote_dir)
        rmt_files = [
            (a.filename, a.st_size)
            for a in attrs
            if a.filename and a.filename.lower().endswith(".rmt")
        ]
        rmt_files = sorted(rmt_files, key=lambda x: x[0])

        logger.info("Found %d .rmt files in %s", len(rmt_files), remote_dir)
        return rmt_files


    def download_files(
        self,
        rmt_files: List[str],
        remote_dir: Optional[str] = None,
        local_dir: Optional[Path] = None,
        overwrite: Optional[bool] = None,
    ) -> List[Path]:
        """
        Download .rmt files from the remote directory to a local directory.

        Args:
            rmt_files: List of filenames to download.
            remote_dir: Remote directory path. Defaults to Config.SFTP_REMOTE_DIR.
            local_dir: Local directory to save files. Defaults to Config.RAW_DATA_DIR.
            overwrite: Whether to overwrite existing files. Defaults to Config setting.

        Returns:
            List of local file paths that were downloaded.
        """
        if self._sftp is None:
            raise RuntimeError("Not connected. Call connect() first.")

        remote_dir = remote_dir or Config.SFTP_REMOTE_DIR
        local_dir = local_dir or Config.RAW_DATA_DIR
        local_dir.mkdir(parents=True, exist_ok=True)

        if overwrite is None:
            overwrite = Config.FILE_EXISTS_BEHAVIOR.lower() == "overwrite"

        downloaded: List[Path] = []

        for filename in rmt_files:
            remote_path = f"{remote_dir}/{filename}" if not remote_dir.endswith("/") else f"{remote_dir}{filename}"
            local_path = local_dir / filename

            if local_path.exists() and not overwrite:
                logger.info("SKIP (exists): %s", filename)
                continue

            try:
                logger.info("Downloading: %s → %s", remote_path, local_path)
                self._sftp.get(remote_path, str(local_path))
                file_size = local_path.stat().st_size
                logger.info("  ✓ Downloaded %s (%s bytes)", filename, f"{file_size:,}")
                downloaded.append(local_path)
            except Exception as e:
                logger.error("  ✗ Failed to download %s: %s", filename, e)

        logger.info(
            "Download complete: %d files downloaded, %d skipped",
            len(downloaded),
            len(rmt_files) - len(downloaded),
        )
        return downloaded

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()
        return False
