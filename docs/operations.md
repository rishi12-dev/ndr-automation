# Operations

Use the prescribed operational folders: `1_Master_NDR/` for the workbook, `2_Input_MIS/` for OMS/courier files, `3_Config_and_Inputs/` for mappings/roster/receiving data, `4_Generated_Calling_Sheets/` for outputs, `5_Agent_Feedback/Completed_Calling_Sheets/` for filled sheets, and `6_Logs_and_Reports/` for run reports. `data/` remains a local-development fallback.

The CLI uses the newest supported OMS workbook for morning processing and recognizes courier files by filename (`MIS`, `courier`, or configured courier name). Run `python process_ndr.py --mode morning --input-dir 2_Input_MIS --master-path 1_Master_NDR/OCT - MONTH NDR.xlsx` in production. Feedback files containing `feedback` in their filename are consumed by the evening mode.

Before production use:

- Confirm column aliases against real OMS/Master NDR/courier samples and update `config/default.yaml`.
- Set `DATABASE_URL`, SMTP settings, backup storage, and authentication policy through environment configuration.
- SMTP is intentionally disabled with placeholder values (`SMTP_ENABLED=false`) until real mail credentials are supplied.
- Run the morning job at 09:00, feedback synchronization at 18:00, and notification jobs at 18:30/48-hour intervals using the host scheduler.
- Back up the SQLite file or use PostgreSQL backups. Never edit a Master workbook directly while a process is running.

The current delivery provides the core import, reconciliation, allocation, export, feedback, dashboard, safe backup, audit model, and idempotency paths. Email delivery, full RBAC, scheduler persistence, and source-specific courier mappings remain configuration/integration work pending real operational credentials and sample files; they are intentionally not fabricated.
