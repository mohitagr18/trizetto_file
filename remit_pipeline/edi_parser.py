"""
EDI 835 (X12) segment-level parser.

Parses .rmt files into structured records — one row per SVC (service line) segment.
Maintains hierarchical context: ISA → GS → ST → BPR/TRN → CLP → SVC.

Segment terminator: ~
Element delimiter: *
Component separator: : (from ISA16)
"""

import csv
import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from remit_pipeline.config import Config

logger = logging.getLogger(__name__)

# Maximum number of CAS (adjustment) groups to expand into columns
MAX_CAS_GROUPS = 3


def _safe_get(elements: List[str], index: int, default: str = "") -> str:
    """Safely get an element by index, returning default if out of range."""
    if index < len(elements):
        return elements[index].strip()
    return default


def _parse_date(date_str: str) -> Optional[str]:
    """Convert CCYYMMDD date string to M/D/YYYY format."""
    if not date_str or len(date_str) != 8:
        return date_str or ""
    try:
        dt = datetime.strptime(date_str, "%Y%m%d")
        return dt.strftime("%-m/%-d/%Y")
    except (ValueError, TypeError):
        return date_str


def _parse_procedure_code(svc01: str) -> str:
    """
    Extract procedure code from SVC01.
    SVC01 is a composite element like 'HC:T1019' or 'HC:T1005:HQ'.
    """
    if not svc01:
        return ""
    # Return the full composite — e.g. 'HC:T1019' or 'HC:T1005:HQ'
    return svc01


def parse_edi_835(filepath: str, filename: str = "") -> List[Dict[str, Any]]:
    """
    Parse a single EDI 835 file into a list of records (one per SVC line).

    Args:
        filepath: Path to the .rmt file.
        filename: Original filename for the Source File column.

    Returns:
        List of dicts, each representing one service line.
    """
    if not filename:
        filename = os.path.basename(filepath)

    with open(filepath, "r", encoding="latin-1") as f:
        content = f.read()

    # Split into segments by ~ terminator
    raw_segments = content.split("~")
    segments = [s.strip() for s in raw_segments if s.strip()]

    records: List[Dict[str, Any]] = []

    # --- Transaction-level context (persists across all CLPs in this ST) ---
    trn_check_number = ""
    bpr_payment_date = ""
    bpr_total_amount = ""
    payer_name = ""
    payee_name = ""
    payee_npi = ""
    tax_id = ""

    # --- Claim-level context (resets at each CLP) ---
    clp_data: Dict[str, str] = {}
    patient_last = ""
    patient_first = ""
    member_id = ""
    rendering_npi = ""
    program_waiver = ""
    claim_first_dos = ""
    claim_last_dos = ""
    claim_allowed = ""

    # --- Service-line level accumulators ---
    svc_data: Optional[Dict[str, Any]] = None

    def _flush_svc():
        """Flush the current SVC record into results."""
        nonlocal svc_data
        if svc_data is None:
            return

        # Merge claim-level and transaction-level context
        svc_data["Check/EFT Number"] = trn_check_number
        svc_data["Payment Date"] = _parse_date(bpr_payment_date)
        svc_data["Payer Name"] = payer_name
        svc_data["Payee Name"] = payee_name
        svc_data["Payee NPI"] = payee_npi
        svc_data["Tax ID"] = tax_id
        svc_data["Patient Last Name"] = patient_last
        svc_data["Patient First Name"] = patient_first
        svc_data["Member ID"] = member_id
        svc_data["Rendering Provider NPI"] = rendering_npi
        svc_data["Program / Waiver"] = program_waiver
        svc_data["Claim Control Number (TCN)"] = clp_data.get("clp07", "")
        svc_data["Claim Status Code"] = clp_data.get("clp02", "")
        svc_data["Charge Amount"] = clp_data.get("clp03", "")
        svc_data["Payment Amount"] = clp_data.get("clp04", "")
        svc_data["Allowed Amount"] = claim_allowed
        svc_data["First DOS"] = _parse_date(claim_first_dos)
        svc_data["Last DOS"] = _parse_date(claim_last_dos)
        svc_data["Source File"] = filename

        records.append(svc_data)
        svc_data = None

    for seg_str in segments:
        elements = seg_str.split("*")
        seg_id = elements[0].upper()

        # ---- Envelope / header segments ----
        if seg_id == "BPR":
            bpr_total_amount = _safe_get(elements, 2)
            bpr_payment_date = _safe_get(elements, 16)

        elif seg_id == "TRN":
            trn_check_number = _safe_get(elements, 2)

        elif seg_id == "DTM":
            qualifier = _safe_get(elements, 1)
            date_val = _safe_get(elements, 2)
            if qualifier == "405":
                # Production date (header level) — use as payment date if BPR16 empty
                if not bpr_payment_date:
                    bpr_payment_date = date_val
            elif qualifier == "232":
                claim_first_dos = date_val
            elif qualifier == "233":
                claim_last_dos = date_val
            elif qualifier == "472":
                # Service date — attach to current SVC
                if svc_data is not None:
                    svc_data["Service Date"] = _parse_date(date_val)

        elif seg_id == "N1":
            qualifier = _safe_get(elements, 1)
            if qualifier == "PR":
                payer_name = _safe_get(elements, 2)
            elif qualifier == "PE":
                payee_name = _safe_get(elements, 2)
                payee_npi = _safe_get(elements, 4)

        elif seg_id == "REF":
            qualifier = _safe_get(elements, 1)
            value = _safe_get(elements, 2)
            if qualifier == "TJ":
                tax_id = value
            elif qualifier == "CE":
                program_waiver = value
            elif qualifier == "6R":
                if svc_data is not None:
                    svc_data["Service Line Reference"] = value
            elif qualifier == "LU":
                if svc_data is not None:
                    svc_data["Location/Modifier Code"] = value

        elif seg_id == "NM1":
            entity_id = _safe_get(elements, 1)
            if entity_id == "QC":
                # Patient
                patient_last = _safe_get(elements, 3)
                patient_first = _safe_get(elements, 4)
                member_id = _safe_get(elements, 9)
            elif entity_id == "74":
                # Rendering / corrected provider
                rendering_npi = _safe_get(elements, 9)

        # ---- Claim-level segment ----
        elif seg_id == "CLP":
            # Flush any pending SVC from the previous claim
            _flush_svc()

            # Reset claim-level context
            clp_data = {
                "clp01": _safe_get(elements, 1),  # Claim submitter ID
                "clp02": _safe_get(elements, 2),  # Status code
                "clp03": _safe_get(elements, 3),  # Charge amount
                "clp04": _safe_get(elements, 4),  # Payment amount
                "clp07": _safe_get(elements, 7),  # Payer claim control number (TCN)
            }
            patient_last = ""
            patient_first = ""
            member_id = ""
            rendering_npi = ""
            program_waiver = ""
            claim_first_dos = ""
            claim_last_dos = ""
            claim_allowed = ""

        elif seg_id == "AMT":
            qualifier = _safe_get(elements, 1)
            amount = _safe_get(elements, 2)
            if qualifier == "AU":
                claim_allowed = amount
            elif qualifier == "B6":
                # Line-level paid amount
                if svc_data is not None:
                    svc_data["Paid Amount per Line"] = amount

        # ---- Service-line segment ----
        elif seg_id == "SVC":
            # Flush the previous SVC if any
            _flush_svc()

            svc_data = {
                "Procedure Code": _parse_procedure_code(_safe_get(elements, 1)),
                "Billed Amount (SVC)": _safe_get(elements, 2),
                "Paid Amount (SVC)": _safe_get(elements, 3),
                "Units": _safe_get(elements, 5),
                "Service Date": "",
                "Service Line Reference": "",
                "Location/Modifier Code": "",
                "Paid Amount per Line": "",
                "Remark Codes": "",
                "_cas_index": 0,
            }
            # Initialize CAS columns
            for i in range(1, MAX_CAS_GROUPS + 1):
                svc_data[f"CAS{i}_Group"] = ""
                svc_data[f"CAS{i}_Reason"] = ""
                svc_data[f"CAS{i}_Amount"] = ""

        elif seg_id == "CAS":
            if svc_data is not None:
                cas_group = _safe_get(elements, 1)  # CO, PR, OA, etc.
                # CAS can have multiple reason/amount triplets within the same segment
                # Elements: CAS*group*reason1*amount1*qty1*reason2*amount2*qty2*...
                triplet_start = 2
                while triplet_start < len(elements):
                    reason = _safe_get(elements, triplet_start)
                    amount = _safe_get(elements, triplet_start + 1)
                    if not reason:
                        break
                    idx = svc_data["_cas_index"] + 1
                    if idx <= MAX_CAS_GROUPS:
                        svc_data[f"CAS{idx}_Group"] = cas_group
                        svc_data[f"CAS{idx}_Reason"] = reason
                        svc_data[f"CAS{idx}_Amount"] = amount
                    svc_data["_cas_index"] = idx
                    triplet_start += 3  # Skip reason, amount, quantity

        elif seg_id in ("MOA", "LQ"):
            # Remark codes
            if svc_data is not None:
                remark = _safe_get(elements, 2) if seg_id == "LQ" else _safe_get(elements, 1)
                if remark:
                    existing = svc_data.get("Remark Codes", "")
                    svc_data["Remark Codes"] = f"{existing}; {remark}".lstrip("; ") if existing else remark

        elif seg_id in ("SE", "GE", "IEA"):
            # End of transaction / group / interchange — flush
            _flush_svc()

    # Final flush
    _flush_svc()

    # Clean up internal fields
    for rec in records:
        rec.pop("_cas_index", None)

    logger.info("Parsed %s: %d service lines from file", filename, len(records))
    return records


def parse_all_rmt_files(
    raw_dir: Optional[Path] = None,
    processed_dir: Optional[Path] = None,
    files_to_parse: Optional[List[str]] = None,
) -> pd.DataFrame:
    """
    Parse .rmt files in the raw directory (or a specific list of files).

    Saves per-file CSVs to processed_dir and returns a combined DataFrame.

    Args:
        raw_dir: Directory containing .rmt files. Defaults to Config.RAW_DATA_DIR.
        processed_dir: Directory for CSV output. Defaults to Config.PROCESSED_DATA_DIR.
        files_to_parse: Specific list of filenames to parse. If None, parses all in raw_dir.

    Returns:
        Combined DataFrame of all parsed records.
    """
    raw_dir = raw_dir or Config.RAW_DATA_DIR
    processed_dir = processed_dir or Config.PROCESSED_DATA_DIR
    processed_dir.mkdir(parents=True, exist_ok=True)

    all_records: List[Dict[str, Any]] = []
    
    if files_to_parse is not None:
        rmt_files = sorted([f for f in files_to_parse if f.lower().endswith(".rmt")])
    else:
        rmt_files = sorted([f for f in os.listdir(raw_dir) if f.lower().endswith(".rmt")])

    if not rmt_files:
        logger.warning("No .rmt files to parse.")
        return pd.DataFrame()

    logger.info("Parsing %d .rmt files", len(rmt_files))

    for filename in rmt_files:
        filepath = raw_dir / filename
        try:
            records = parse_edi_835(str(filepath), filename)
            all_records.extend(records)

            # Save per-file CSV
            if records:
                csv_name = Path(filename).stem + ".csv"
                csv_path = processed_dir / csv_name
                df_file = pd.DataFrame(records)
                df_file.to_csv(csv_path, index=False)
                logger.info("  Saved CSV: %s (%d rows)", csv_path, len(records))

        except Exception as e:
            logger.error("  ✗ Failed to parse %s: %s", filename, e, exc_info=True)
            continue

    if not all_records:
        logger.warning("No records parsed from any files.")
        return pd.DataFrame()

    df = pd.DataFrame(all_records)
    logger.info(
        "Total parsed: %d service lines from %d files",
        len(df),
        len(rmt_files),
    )
    return df
