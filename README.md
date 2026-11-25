# COA Request Builder

Desktop and web application that generates **Certificate of Analysis (COA) requests** (`PERMINTAAN COA.xlsx`) from daily shipping plans (`Rencana Kirim.xlsx`) and SAP unrestricted stock exports (`EXPORT.xlsx`) using First-Expiry-First-Out (FEFO) allocation. Includes a **Weekly Audit** workflow to reconcile manual batch picks against automated FEFO selections.

## Features

- **FEFO Batch Allocation (`coa_engine.py`):**
  - Parses shipping plan sheets (`TANGGAL dd - mm - yyyy`), filtering valid 10-digit Sales Orders (`NOMOR SO`) and skipping cancelled (`CANCEL`) or non-customer transfer lines.
  - Prioritizes stock reserved to the matching Sales Order (`Special Stock Number`), then allocates remaining demand by earliest expiration date (`SLED/BBD`).
  - Substitutes `Supplier Batch` as the printed COA lot for trading materials (configurable storage locations such as `2002` or material codes starting with `41`).
  - Calculates remaining shelf-life percentage relative to `plan_date - ref_offset_days` and flags shortages or unit mismatches with `CEK STOK`.
- **Native Windows Desktop GUI (`coa_gui.py`):**
  - Tkinter desktop app with background worker threads, live preview table, local SQLite audit history, and `.xlsx` export.
  - Self-contained Windows `.exe` packaging via `build.ps1` and PyInstaller (`COARequestBuilder.spec`).
- **Streamlit Web UI (`app.py`):**
  - Browser interface sharing the same engine and SQLite audit store, ready for Docker deployment.
- **Weekly Audit & Compliance (`audit.py`, `audit_store.py`):**
  - Pairs customer/truck blocks between manual and automated COA sheets (sequential order matching with fuzzy customer-name fallback) and classifies rows into `MATCH`, `QTY DIFFERENT`, `MISSING IN MANUAL`, `EXTRA IN MANUAL`, `BLOCK MISSING IN MANUAL`, and `BLOCK NOT IN PLAN`.
  - Tracks 12-week compliance in SQLite and optionally locks request generation from Wednesday onward until the current ISO week's audit is saved or skipped with a supervisor PIN.

## Project Structure

```text
coa_app/
├── coa_engine.py             # Plan parser, SAP stock reader, FEFO allocator, Excel writer
├── coa_gui.py                # Native Windows desktop application (Tkinter)
├── app.py                    # Web application interface (Streamlit)
├── audit.py                  # Manual vs. automated comparison & Excel report builder
├── audit_store.py            # SQLite audit persistence and Wednesday lock rules
├── test_smoke.py             # Backend smoke test suite
├── build.ps1                 # Windows build script (venv, tests, PyInstaller, signing)
├── COARequestBuilder.spec    # PyInstaller build specification
├── Dockerfile                # Container definition for the Streamlit web UI
├── requirements.txt          # Python dependencies
└── README.md
