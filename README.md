# EDI 835 Remittance Processing Pipeline

Automated pipeline to download, parse, and report on X12 EDI 835 remittance files from an SFTP server for a home healthcare provider.

## Setup

### Prerequisites
- Python 3.9+
- [`uv`](https://docs.astral.sh/uv/) package manager

### Install Dependencies
```bash
uv sync
```

This will create a virtual environment (`.venv/`) and install all dependencies from `pyproject.toml`.

### Configure Credentials
1. Copy the environment template:
   ```bash
   cp .env.example .env
   ```
2. Edit `.env` with your SFTP credentials.

## Usage

### Full Pipeline (Download + Parse + Report)
```bash
uv run python main.py
```

### Download Only (SFTP → local)
```bash
uv run python main.py --download-only
```

### Parse Only (local files → Excel report)
```bash
uv run python main.py --parse-only
```

### Custom Output Directory
```bash
uv run python main.py --output-dir ./my_output/
```

## Project Structure
```
trizetto_file/
├── main.py                      # CLI entry point
├── pyproject.toml               # Dependencies & project metadata
├── .python-version              # Python version pin
├── .env                         # Credentials (git-ignored)
├── .env.example                 # Credential template
├── remit_pipeline/
│   ├── __init__.py
│   ├── config.py                # Environment config loader
│   ├── sftp_client.py           # SFTP connection & download
│   ├── edi_parser.py            # X12 EDI 835 segment parser
│   └── report_builder.py        # Master Excel report generator
├── data/
│   ├── raw/                     # Downloaded .rmt files
│   ├── processed/               # Parsed CSVs
│   └── V6.15...xlsx             # Master report template
└── output/                      # Generated Excel reports
```

## EDI 835 Parser Design

Uses **custom segment-level parsing** (not a third-party 835 library) for full control over:
- Hierarchical context tracking (ISA → ST → BPR/TRN → CLP → SVC)
- Specific field extraction (REF*6R, REF*LU, AMT*B6, MOA, LQ, etc.)
- Multiple CAS segment handling per service line
- Error resilience — failed files are logged and skipped