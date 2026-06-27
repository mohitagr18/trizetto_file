"""
EDI 835 Remittance Processing Pipeline — Entry Point

Usage:
    uv run python main.py                    # Full pipeline (download + parse + report)
    uv run python main.py --download-only    # Only download from SFTP
    uv run python main.py --parse-only       # Parse already-downloaded files
    uv run python main.py --output-dir DIR   # Override output directory
"""

import argparse
import logging
import sys
from pathlib import Path

from remit_pipeline.config import Config


def setup_logging() -> None:
    """Configure logging with timestamped, leveled output."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def run_download() -> int:
    """Steps 2–3: Connect to SFTP, list and download .rmt files."""
    from remit_pipeline.sftp_client import SFTPClient

    logger = logging.getLogger("pipeline.download")

    try:
        with SFTPClient() as sftp:
            rmt_files = sftp.list_rmt_files()
            if not rmt_files:
                logger.warning("No .rmt files found on server.")
                return 0

            logger.info("Found %d .rmt files on server.", len(rmt_files))
            downloaded = sftp.download_files(rmt_files)
            logger.info("Downloaded %d new files.", len(downloaded))
            return len(rmt_files)

    except Exception as e:
        logger.error("SFTP download failed: %s", e, exc_info=True)
        return -1


def run_parse_and_report(output_dir: Path = None) -> str:
    """Steps 4–5: Parse downloaded .rmt files and build the master Excel report."""
    from remit_pipeline.edi_parser import parse_all_rmt_files
    from remit_pipeline.report_builder import build_and_write_report

    logger = logging.getLogger("pipeline.parse")

    # Parse all .rmt files
    parsed_df = parse_all_rmt_files()

    if parsed_df.empty:
        logger.warning("No data parsed from .rmt files. Check data/raw/ directory.")
        return ""

    # Build and write the master Excel report
    output_path = build_and_write_report(parsed_df, output_dir=output_dir)
    return str(output_path)


def main():
    parser = argparse.ArgumentParser(
        description="EDI 835 Remittance Processing Pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  uv run python main.py                  Full pipeline (download + parse + report)
  uv run python main.py --download-only  Only download .rmt files from SFTP
  uv run python main.py --parse-only     Parse local files and generate Excel report
  uv run python main.py --output-dir ./my_output/
        """,
    )
    parser.add_argument(
        "--download-only",
        action="store_true",
        help="Only connect to SFTP and download .rmt files (Steps 2–3)",
    )
    parser.add_argument(
        "--parse-only",
        action="store_true",
        help="Skip SFTP download; parse already-downloaded files (Steps 4–5)",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None,
        help="Override the default output directory",
    )
    args = parser.parse_args()

    # Validate mutually exclusive flags
    if args.download_only and args.parse_only:
        print("ERROR: --download-only and --parse-only are mutually exclusive.")
        sys.exit(1)

    # Setup
    setup_logging()
    logger = logging.getLogger("pipeline")

    if args.output_dir:
        Config.OUTPUT_DIR = Path(args.output_dir)

    Config.ensure_directories()

    print()
    print("=" * 64)
    print("  EDI 835 Remittance Processing Pipeline")
    print("=" * 64)
    print(f"  Raw data dir:       {Config.RAW_DATA_DIR}")
    print(f"  Processed data dir: {Config.PROCESSED_DATA_DIR}")
    print(f"  Output dir:         {Config.OUTPUT_DIR}")
    print(f"  Mode:               {'Download Only' if args.download_only else 'Parse Only' if args.parse_only else 'Full Pipeline'}")
    print("=" * 64)
    print()

    # ---- Execute pipeline ----

    if args.parse_only:
        # Skip download, go straight to parse + report
        logger.info("Mode: --parse-only (skipping SFTP download)")
        output_path = run_parse_and_report(
            output_dir=Config.OUTPUT_DIR,
        )
        if output_path:
            _print_summary(output_path)
        else:
            logger.warning("No output generated.")
            sys.exit(1)

    elif args.download_only:
        # Only download
        logger.info("Mode: --download-only")
        count = run_download()
        if count < 0:
            sys.exit(1)
        print(f"\n✓ Download complete. {count} .rmt files available in {Config.RAW_DATA_DIR}")

    else:
        # Full pipeline: download → parse → report
        logger.info("Mode: Full pipeline")

        # Step 1: Download
        count = run_download()
        if count < 0:
            logger.error("Download failed. Aborting pipeline.")
            sys.exit(1)
        if count == 0:
            logger.warning("No files to process.")
            sys.exit(0)

        # Step 2: Parse + Report
        output_path = run_parse_and_report(
            output_dir=Config.OUTPUT_DIR,
        )
        if output_path:
            _print_summary(output_path)
        else:
            logger.warning("No output generated.")
            sys.exit(1)


def _print_summary(output_path: str) -> None:
    """Print a final summary of the generated report."""
    import os

    from openpyxl import load_workbook

    path = Path(output_path)
    size = path.stat().st_size

    wb = load_workbook(str(path), read_only=True)
    sheet_info = []
    for name in wb.sheetnames:
        ws = wb[name]
        rows = ws.max_row - 4  # Subtract header rows
        sheet_info.append((name, rows))
    wb.close()

    print()
    print("=" * 64)
    print("  ✓ Pipeline Complete")
    print("=" * 64)
    print(f"  Output file: {output_path}")
    print(f"  File size:   {size:,} bytes ({size / 1024:.1f} KB)")
    print()
    for name, rows in sheet_info:
        print(f"  Sheet: {name:30s} → {rows:,} rows")
    print("=" * 64)


if __name__ == "__main__":
    main()
