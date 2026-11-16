from datetime import datetime
from pathlib import Path
from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.cron import CronTrigger
from zoneinfo import ZoneInfo

from ndr_api.config import settings, load_rules
from ndr_api.db import SessionLocal, init_db
from ndr_api.automation import completion_report, rto_report, import_feedback
from process_ndr import morning

ROOT = Path(__file__).parent
INPUT_DIR = ROOT / "2_Input_MIS"
MASTER_PATH = ROOT / "1_Master_NDR" / "OCT - MONTH NDR.xlsx"
FEEDBACK_DIR = ROOT / "5_Agent_Feedback" / "Completed_Calling_Sheets"
LOG_DIR = ROOT / "6_Logs_and_Reports"

def run_job(name, callback):
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    session = SessionLocal()
    try:
        result = callback(session)
        (LOG_DIR / f"{name}_{datetime.now():%Y%m%d}.log").open("a", encoding="utf-8").write(f"{datetime.now().isoformat()} {result}\n")
    except Exception as exc:
        (LOG_DIR / f"{name}_{datetime.now():%Y%m%d}.log").open("a", encoding="utf-8").write(f"{datetime.now().isoformat()} ERROR {exc}\n")
    finally:
        session.close()

def morning_job(session):
    return morning(session, load_rules(), str(MASTER_PATH) if MASTER_PATH.exists() else None)

def evening_job(session):
    results = []
    for path in FEEDBACK_DIR.glob("*feedback*"):
        results.append({path.name: import_feedback(session, str(path))})
    return {"feedback": results, "completion": completion_report(session, datetime.now().date())}

def main():
    init_db(); scheduler = BlockingScheduler(timezone=ZoneInfo(settings.app_timezone))
    scheduler.add_job(lambda: run_job("morning", morning_job), CronTrigger(hour=9, minute=0, day_of_week="mon-sat"))
    scheduler.add_job(lambda: run_job("evening", evening_job), CronTrigger(hour=18, minute=0, day_of_week="mon-sat"))
    scheduler.add_job(lambda: run_job("ndr_notify", lambda session: completion_report(session, datetime.now().date())), CronTrigger(hour=18, minute=30, day_of_week="mon-sat"))
    scheduler.add_job(lambda: run_job("rto_notify", lambda session: rto_report(session, datetime.now().date())), CronTrigger(hour=9, minute=0, day_of_week="mon,wed,fri"))
    print("NDR scheduler running. Press Ctrl+C to stop.")
    scheduler.start()

if __name__ == "__main__": main()
