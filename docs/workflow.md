# EDI 835 Remittance Pipeline Workflows

This document outlines the detailed workflows and data pipelines for the EDI 835 Remittance processing system, from raw file ingestion to downstream reconciliation.

## 1. End-to-End Ecosystem

The `trizetto_file` repository is the first step in a larger reconciliation ecosystem. It connects to the clearinghouse (GatewayEDI), parses raw EDI 835 files, and outputs a structured Excel report that serves as the source of truth for downstream reconciliation systems (like the `remittance_recon` tool).

```mermaid
flowchart TD
    subgraph External["External Systems"]
        SFTP["GatewayEDI SFTP Server\n(Raw .rmt files)"]
    end

    subgraph TrizettoPipeline["Trizetto File Pipeline (Extract & Parse)"]
        DL["Downloader / Scheduler"]
        Parser["EDI 835 Parser"]
        Cache[("Local SQLite / CSV Caches")]
        ReportGen["Report Builder"]
        DL --> Parser --> Cache --> ReportGen
    end

    subgraph Downstream["Downstream Applications"]
        Excel["Master_Remittance_Report.xlsx\n(Structured Excel)"]
        Recon["remittance_recon Application\n(DuckDB + Streamlit UI)"]
        Tracker["Unskilled Tracker & Reconciliation Dashboard"]
    end

    SFTP -->|"Downloads new files"| DL
    ReportGen -->|"Generates"| Excel
    Excel -->|"Ingested by"| Recon
    Recon -->|"Drives UI for"| Tracker

    classDef ext fill:#f9f9f9,stroke:#333,stroke-width:2px;
    classDef pipe fill:#e1f5fe,stroke:#0288d1,stroke-width:2px;
    classDef down fill:#e8f5e9,stroke:#388e3c,stroke-width:2px;
    class External ext;
    class TrizettoPipeline pipe;
    class Downstream down;
```

## 2. Report Builder Data Flow

The Reporting module (`report_builder.py`) contains complex logic to transform raw EDI segments into business-readable summaries. A critical component is rolling up service lines into a claim summary and handling procedure code extraction with modifiers.

```mermaid
flowchart LR
    subgraph RawData["Raw Parsed Data (CSV Cache)"]
        CLP["CLP Segments (Claim Level)"]
        SVC["SVC Segments (Service Line)"]
    end

    subgraph Transformation["Transformation & Aggregation"]
        Group["Group by TCN (Claim Identifier)"]
        Calc["Calculate Total Billed/Paid Hours"]
        Proc["Extract & Format Procedure Codes\n(e.g., 'T1005: 76')"]
    end

    subgraph Output["Master Remittance Report"]
        Sheet1["Sheet 1: Claim Summary\n- One row per Claim\n- Formatted Procedure Codes\n- Aggregated Hours"]
        Sheet2["Sheet 2: Service Line Detail\n- Granular Adjustments\n- Remark Codes"]
    end

    CLP --> Group
    SVC --> Group
    Group --> Calc
    Group --> Proc
    Calc --> Sheet1
    Proc --> Sheet1
    SVC --> Sheet2

    classDef data fill:#fff3e0,stroke:#f57c00,stroke-width:2px;
    classDef trans fill:#f3e5f5,stroke:#8e24aa,stroke-width:2px;
    classDef out fill:#e0f7fa,stroke:#0097a7,stroke-width:2px;
    class RawData data;
    class Transformation trans;
    class Output out;
```
