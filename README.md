# NDR Automation & Outbound Calls Engine

A configurable Python/FastAPI system for OMS ingestion, courier reconciliation, NDR allocation, calling feedback, RTO tracking, Excel exports, and operational reporting.

## Quick start

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e ".[dev]"
Copy-Item .env.example .env
python process_ndr.py --mode morning --input-dir data/incoming --master-path path/to/Master_NDR.xlsx
python process_ndr.py --mode evening
python process_ndr.py --mode rto-notify
python process_ndr.py --mode ndr-notify
uvicorn ndr_api.main:app --reload
pytest
```

The default database is SQLite (`data/ndr.sqlite3`). Set `DATABASE_URL` to PostgreSQL for production. Real source workbooks were not included in this workspace; see `docs/data-mapping-report.md` before configuring aliases in `config/default.yaml`.

## Processing commands

```text
python process_ndr.py --mode morning --master-path data/Master_NDR.xlsx
python process_ndr.py --mode evening
python process_ndr.py --mode rto-notify
python process_ndr.py --mode ndr-notify
```

The API exposes `/docs`, health at `/api/health`, and operational endpoints under `/api`.

## Automatic daily operation

Keep the scheduler running once (or register `start_automation.ps1` in Windows Task Scheduler at system startup):

```powershell
.\start_automation.ps1
```

It runs morning processing at 09:00, feedback sync at 18:00, NDR summary at 18:30, and RTO summaries every 48 hours on Asia/Kolkata time. Keep the Master NDR and downloaded input files in the folders described in `docs/operations.md`.
