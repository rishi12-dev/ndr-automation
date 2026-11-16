import argparse
from datetime import date
from pathlib import Path
from ndr_api.config import load_rules, settings
from ndr_api.db import SessionLocal, init_db
from ndr_api.core import import_oms, reconcile_mis, eligible_orders, smart_allocate, export_assignments
from ndr_api.models import CallingAssignment, Order
from ndr_api.automation import import_feedback, rto_report, completion_report, update_master_ndr, process_self_collect, validate_calling_outputs, previous_working_day_girish_dockets, prepare_agent_sheet, reverse_sync_feedback
import json
from sqlalchemy import select

def latest_files(folder, patterns):
    files = [path for pattern in patterns for path in Path(folder).glob(pattern)]
    return max(files, key=lambda path: path.stat().st_mtime) if files else None

def morning(session, rules, master_path=None):
    candidates = [path for pattern in ["*.xlsx", "*.xls", "*.csv"] for path in Path(settings.input_dir).glob(pattern) if "feedback" not in path.stem.casefold() and "mis" not in path.stem.casefold() and "courier" not in path.stem.casefold()]
    oms = max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None
    if not oms: return {"warning": "No OMS workbook found", "exports": []}
    result = import_oms(session, str(oms), rules)
    if master_path:
        result["master"] = update_master_ndr(master_path, session.query(Order).all(), date.today())
    previous_girish = previous_working_day_girish_dockets(master_path, date.today()) if master_path and Path(master_path).exists() else set()
    result["previous_girish_excluded"] = len(previous_girish)
    girish, local = eligible_orders(session, date.today(), rules, previous_girish_dockets=previous_girish)
    for path in sorted(Path(settings.input_dir).glob("*")):
        if path.is_file() and any(token in path.stem.casefold() for token in ("mis", "courier", "delhivery", "dtdc", "chennai", "professional", "nandan", "pushpak")):
            try: result.setdefault("mis", []).append({path.name: reconcile_mis(session, str(path), rules)})
            except ValueError as exc: result.setdefault("validation_errors", []).append({path.name: str(exc)})
    result["self_collect"] = process_self_collect(session, master_path)
    roster_path = Path("3_Config_and_Inputs/agent_roster.json")
    roster = json.loads(roster_path.read_text(encoding="utf-8")) if roster_path.exists() else {name: {"active": True, "present": True} for name in rules["agents"]}
    assignments = smart_allocate(session, girish, local, roster, date.today())
    girish_assignments = [assignment for assignment in assignments if assignment.source == "girish"]
    local_assignments = [assignment for assignment in assignments if assignment.source != "girish"]
    outputs = []
    if girish_assignments: outputs.append(export_assignments(session, girish_assignments, f"{settings.output_dir}/Girish_NDR_{date.today():%d-%m-%Y}.xlsx"))
    if local_assignments: outputs.append(export_assignments(session, local_assignments, f"{settings.output_dir}/Local_NDR_{date.today():%d-%m-%Y}.xlsx"))
    for output in outputs: prepare_agent_sheet(output)
    girish_path = next((path for path in outputs if "Girish_NDR" in str(path)), None); local_path = next((path for path in outputs if "Local_NDR" in str(path)), None)
    result["validation"] = validate_calling_outputs(girish_path, local_path, date.today())
    return {**result, "exports": outputs}

def evening(session):
    feedback_files = [path for path in Path(settings.input_dir).glob("*") if path.is_file() and "feedback" in path.stem.casefold()]
    result = {"sync": [], "completion": []}
    master = Path("1_Master_NDR/OCT - MONTH NDR.xlsx")
    if master.exists(): result["reverse_sync"] = reverse_sync_feedback(session, str(master), "5_Agent_Feedback/Completed_Calling_Sheets", date.today())
    for path in feedback_files:
        try: result["sync"].append({path.name: import_feedback(session, str(path))})
        except ValueError as exc: result["sync"].append({path.name: {"error": str(exc)}})
    result["completion"] = completion_report(session, date.today()); Path("6_Logs_and_Reports").mkdir(parents=True, exist_ok=True); Path(f"6_Logs_and_Reports/daily_agent_completion_{date.today():%Y-%m-%d}.json").write_text(__import__("json").dumps(result["completion"], default=str, indent=2), encoding="utf-8"); return result

def main():
    parser = argparse.ArgumentParser(); parser.add_argument("--mode", choices=["morning", "evening", "rto-notify", "ndr-notify"], required=True); parser.add_argument("--input-dir", default=None); parser.add_argument("--master-path", default=None); args = parser.parse_args()
    if args.input_dir: settings.input_dir = args.input_dir
    init_db(); session = SessionLocal(); rules = load_rules()
    try:
        master_path = args.master_path or ("1_Master_NDR/OCT - MONTH NDR.xlsx" if Path("1_Master_NDR/OCT - MONTH NDR.xlsx").exists() else None)
        if args.mode == "morning": result = morning(session, rules, master_path)
        elif args.mode == "evening": result = evening(session)
        elif args.mode == "rto-notify": result = rto_report(session, date.today())
        else: result = {"type": "ndr-notify", "date": str(date.today()), "completion": completion_report(session, date.today())}
        print(result)
    finally: session.close()
if __name__ == "__main__": main()
