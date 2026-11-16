from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from fastapi import FastAPI, Depends, File, UploadFile, HTTPException, Header
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import RedirectResponse, FileResponse
from pydantic import BaseModel
import hashlib, secrets, shutil, json
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from sqlalchemy.orm import Session
from sqlalchemy import select, func, or_
from zoneinfo import ZoneInfo
from .config import load_rules, settings
from .db import SessionLocal, init_db
from .models import Order, CallingAssignment, CallingAttempt, RTOCase, User
from .core import import_oms, import_oms_pair, reconcile_mis, eligible_orders, smart_allocate, export_assignments, normalize_docket
from .automation import import_master_workbook, process_self_collect, reverse_sync_feedback, sync_master_assignments, update_master_ndr, previous_working_day_girish_dockets

app = FastAPI(title="NDR Automation API", version="0.1.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])
app.mount("/portal", StaticFiles(directory="frontend", html=True), name="portal")
sessions = {}
MANUAL_MODE_FILE = Path("data/manual_mode.json")

def working_date() -> date:
    """Use the date block from the active manual upload during a demo or catch-up."""
    if MANUAL_MODE_FILE.exists():
        try:
            value = json.loads(MANUAL_MODE_FILE.read_text(encoding="utf-8")).get("active_date")
            if value:
                return date.fromisoformat(str(value))
        except (OSError, ValueError, TypeError):
            pass
    return date.today()

def manual_mode_enabled() -> bool:
    if not MANUAL_MODE_FILE.exists():
        return False
    try:
        return bool(json.loads(MANUAL_MODE_FILE.read_text(encoding="utf-8")).get("enabled"))
    except (OSError, ValueError, TypeError):
        return False

def set_manual_mode(enabled: bool, filename: str = "", active_date: date | None = None) -> None:
    MANUAL_MODE_FILE.parent.mkdir(parents=True, exist_ok=True)
    MANUAL_MODE_FILE.write_text(json.dumps({"enabled": enabled, "filename": filename, "active_date": active_date.isoformat() if active_date else None, "updated_at": datetime.now().isoformat()}), encoding="utf-8")

@app.middleware("http")
async def no_cache_portal(request, call_next):
    response = await call_next(request)
    if request.url.path.startswith("/portal"): response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
    return response

def password_hash(password: str) -> str:
    salt = secrets.token_hex(16); digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 120000).hex(); return f"{salt}${digest}"

def password_ok(password: str, stored: str) -> bool:
    salt, digest = stored.split("$", 1); return secrets.compare_digest(hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 120000).hex(), digest)

def seed_users():
    session = SessionLocal()
    if not session.scalar(select(User).where(User.username == "admin")):
        users = [("admin", "admin123", "Administrator", "admin"), ("girish", "girish123", "Girish", "agent"), ("pavi", "pavi123", "Pavi", "agent"), ("zaffar", "zaffar123", "Mohammed Zaffruddin", "agent"), ("saqlain", "saqlain123", "Saqlain Muhammad", "agent"), ("rubby", "rubby123", "Rubby Srinivasan", "agent"), ("narayan", "narayan123", "Narayanaswamy K.V", "agent")]
        for username, password, name, role in users: session.add(User(username=username, password_hash=password_hash(password), display_name=name, role=role))
        session.commit()
    session.close()

def current_user(authorization: str | None, session: Session):
    user_id = sessions.get((authorization or "").removeprefix("Bearer "))
    return session.get(User, user_id) if user_id else None

@app.get("/", include_in_schema=False)
def home(): return RedirectResponse("/portal/")

@app.get("/health")
def health(): return {"status": "ok", "service": "ndr-automation"}

@app.on_event("startup")
def startup():
    init_db()
    seed_users()
    master = Path("1_Master_NDR/OCT - MONTH NDR.xlsx")
    session = SessionLocal()
    try:
        if master.exists() and (session.scalar(select(func.count()).select_from(Order)) or 0) < 100:
            import_master_workbook(session, str(master))
        if master.exists() and not manual_mode_enabled():
            sync_master_assignments(session, str(master), date.today())
    finally:
        session.close()

def db():
    session = SessionLocal()
    try: yield session
    finally: session.close()

def run_daily_automation(session: Session) -> dict:
    if manual_mode_enabled():
        raise HTTPException(409, "Automation is paused. Upload the manual NDR; only manual calls are active.")
    today = date.today(); rules = load_rules(); master = Path("1_Master_NDR/OCT - MONTH NDR.xlsx")
    previous = set()
    girish, local = eligible_orders(session, today, rules, previous_girish_dockets=previous)
    roster_path = Path("3_Config_and_Inputs/agent_roster.json")
    roster = json.loads(roster_path.read_text(encoding="utf-8")) if roster_path.exists() else {name: {"active": True, "present": True} for name in rules["agents"]}
    assignments = smart_allocate(session, girish, local, roster, today)
    return {"eligible_delhivery": len(girish), "eligible_local": len(local), "allocated": len(assignments), "previous_girish_excluded": len(previous)}

@app.get("/api/health")
def health(): return {"status": "ok", "service": "ndr-automation"}

class LoginPayload(BaseModel): username: str; password: str

@app.post("/api/auth/login")
def login(payload: LoginPayload, session: Session = Depends(db)):
    identity = payload.username.strip().casefold()
    user = session.scalar(select(User).where(or_(User.username == identity, func.lower(User.display_name) == identity), User.active.is_(True)))
    if not user or not password_ok(payload.password, user.password_hash): raise HTTPException(401, "Invalid username or password")
    token = secrets.token_urlsafe(32); sessions[token] = user.id; return {"token": token, "username": user.username, "display_name": user.display_name, "role": user.role}

@app.get("/api/me")
def me(authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user: raise HTTPException(401, "Login required")
    return {"username": user.username, "display_name": user.display_name, "role": user.role}

class UserCreate(BaseModel):
    username: str; password: str; display_name: str; role: str = "agent"

class PasswordChange(BaseModel):
    password: str

class UserUpdate(BaseModel):
    display_name: str | None = None
    role: str | None = None
    active: bool | None = None

@app.get("/api/admin/users")
def admin_users(authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or user.role != "admin": raise HTTPException(403, "Admin access required")
    return [{"username": item.username, "display_name": item.display_name, "role": item.role, "active": item.active} for item in session.scalars(select(User).order_by(User.role, User.display_name)).all()]

@app.post("/api/admin/users")
def create_user(payload: UserCreate, authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or user.role != "admin": raise HTTPException(403, "Admin access required")
    username = payload.username.strip().casefold(); display_name = payload.display_name.strip(); role = payload.role.strip().casefold()
    if not username or not display_name or len(payload.password) < 6: raise HTTPException(422, "Username, display name and a 6+ character password are required")
    if role not in {"admin", "agent"}: raise HTTPException(422, "Role must be admin or agent")
    if session.scalar(select(User).where(User.username == username)): raise HTTPException(409, "Username already exists")
    item = User(username=username, password_hash=password_hash(payload.password), display_name=display_name, role=role); session.add(item); session.commit()
    return {"username": item.username, "display_name": item.display_name, "role": item.role}

@app.put("/api/admin/users/{username}/password")
def change_user_password(username: str, payload: PasswordChange, authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or user.role != "admin": raise HTTPException(403, "Admin access required")
    if len(payload.password) < 6: raise HTTPException(422, "Password must be at least 6 characters")
    item = session.scalar(select(User).where(User.username == username.casefold()))
    if not item: raise HTTPException(404, "User not found")
    item.password_hash = password_hash(payload.password); session.commit(); return {"updated": True, "username": item.username}

@app.put("/api/admin/users/{username}")
def update_user(username: str, payload: UserUpdate, authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or user.role != "admin": raise HTTPException(403, "Admin access required")
    item = session.scalar(select(User).where(User.username == username.casefold()))
    if not item: raise HTTPException(404, "User not found")
    if payload.display_name is not None:
        display_name = payload.display_name.strip()
        if not display_name: raise HTTPException(422, "Display name cannot be empty")
        item.display_name = display_name
    if payload.role is not None:
        role = payload.role.strip().casefold()
        if role not in {"admin", "agent"}: raise HTTPException(422, "Role must be admin or agent")
        item.role = role
    if payload.active is not None:
        if item.id == user.id and not payload.active: raise HTTPException(422, "You cannot deactivate your own account")
        item.active = payload.active
    session.commit()
    return {"updated": True, "username": item.username, "display_name": item.display_name, "role": item.role, "active": item.active}

@app.get("/api/my-dashboard")
def my_dashboard(authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user: raise HTTPException(401, "Login required")
    assignments = session.scalars(select(CallingAssignment).where(CallingAssignment.agent == user.display_name, CallingAssignment.assignment_date == working_date(), CallingAssignment.status != "test")).all() if user.role != "admin" else session.scalars(select(CallingAssignment).where(CallingAssignment.assignment_date == working_date(), CallingAssignment.status != "test")).all()
    return {"assigned": len(assignments), "completed": sum(item.status == "completed" for item in assignments), "pending": sum(item.status != "completed" for item in assignments), "agent": user.display_name}

@app.get("/api/reminder-status")
def reminder_status(authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user: raise HTTPException(401, "Login required")
    if user.role == "admin": return {"due": False, "pending": 0, "reason": "admin"}
    assignments = session.scalars(select(CallingAssignment).where(CallingAssignment.agent == user.display_name, CallingAssignment.assignment_date == working_date(), CallingAssignment.status != "completed")).all()
    pending = len([item for item in assignments if item.status != "test"])
    if pending == 0: return {"due": False, "pending": 0, "reason": "complete"}
    manual = {}
    if MANUAL_MODE_FILE.exists():
        try: manual = json.loads(MANUAL_MODE_FILE.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError): manual = {}
    uploaded_at = manual.get("updated_at")
    now = datetime.now(ZoneInfo(settings.app_timezone))
    if not uploaded_at: return {"due": False, "pending": pending, "reason": "upload_time_missing"}
    try: upload_time = datetime.fromisoformat(uploaded_at).replace(tzinfo=ZoneInfo(settings.app_timezone))
    except ValueError: return {"due": False, "pending": pending, "reason": "upload_time_invalid"}
    if now < upload_time or now.hour < 14 or (now.hour == 18 and now.minute >= 20) or now.hour > 18:
        return {"due": False, "pending": pending, "reason": "outside_window"}
    if now.hour < 17:
        minutes = ((now.hour - 14) * 60 + now.minute) // 30
        slot = now.replace(hour=14, minute=0, second=0, microsecond=0) + timedelta(minutes=minutes * 30)
        interval = 30
    else:
        minutes = ((now.hour - 17) * 60 + now.minute) // 10
        slot = now.replace(hour=17, minute=0, second=0, microsecond=0) + timedelta(minutes=minutes * 10)
        interval = 10
    if slot <= upload_time: return {"due": False, "pending": pending, "reason": "waiting_for_next_slot", "interval_minutes": interval}
    return {"due": True, "pending": pending, "interval_minutes": interval, "reminder_key": slot.isoformat(), "message": f"{pending} calls pending. Please complete your allocated NDR calls."}

@app.get("/api/admin/performance")
def admin_performance(authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or user.role != "admin": raise HTTPException(403, "Admin access required")
    agents = session.scalars(select(User).where(User.role == "agent", User.active.is_(True))).all(); result = []
    for agent in agents:
        assignments = session.scalars(select(CallingAssignment).where(CallingAssignment.agent == agent.display_name, CallingAssignment.assignment_date == working_date(), CallingAssignment.status != "test")).all(); completed = sum(item.status == "completed" for item in assignments)
        result.append({"agent": agent.display_name, "assigned": len(assignments), "completed": completed, "pending": len(assignments) - completed, "completion_pct": round(completed * 100 / len(assignments), 1) if assignments else 0})
    return result

class ManualAssignment(BaseModel):
    identifier: str; agent: str

@app.post("/api/admin/assign")
def admin_assign(payload: ManualAssignment, authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or user.role != "admin": raise HTTPException(403, "Admin access required")
    identifier = payload.identifier.strip(); order = session.scalar(select(Order).where(or_(Order.order_no == identifier, Order.docket_no == identifier)))
    if not order: raise HTTPException(404, "OrderNo/Docketno not found")
    if session.scalar(select(CallingAssignment).where(CallingAssignment.order_id == order.id, CallingAssignment.assignment_date == date.today())): raise HTTPException(409, "Order already assigned today")
    sequence = (session.scalar(select(func.count()).select_from(CallingAssignment).where(CallingAssignment.assignment_date == date.today())) or 0) + 1
    assignment = CallingAssignment(assignment_id=f"CALL-{date.today():%Y%m%d}-{sequence:04d}", order_id=order.id, agent=payload.agent, assignment_date=date.today(), source="admin-manual"); session.add(assignment); session.commit(); return {"assignment_id": assignment.assignment_id, "agent": assignment.agent}

@app.get("/api/my-assignments")
def my_assignments(authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user: raise HTTPException(401, "Login required")
    if user.role == "admin": return assignments("Girish", session)
    rows = assignments(user.display_name, session)
    return [row for row in rows if str(row.get("status", "Shipped")).casefold() != "refunded"]

@app.get("/api/my-assignments/export")
def export_my_assignments(authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or user.role == "admin": raise HTTPException(403, "Agent access required")
    rows = session.scalars(select(CallingAssignment).where(CallingAssignment.agent == user.display_name, CallingAssignment.assignment_date == working_date(), CallingAssignment.status != "test")).all()
    output_dir = Path(settings.output_dir); output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / f"{user.username}_Calling_{working_date():%d-%m-%Y}.xlsx"
    export_assignments(session, rows, str(target), girish_format=user.display_name.strip().casefold() == "girish")
    return FileResponse(target, filename=target.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/api/my-completed/export")
def export_my_completed(authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or user.role == "admin": raise HTTPException(403, "Agent access required")
    if (datetime.now().hour, datetime.now().minute) < (18, 20): raise HTTPException(403, "Completed-call download opens at 6:20 PM")
    rows = session.scalars(select(CallingAssignment).where(CallingAssignment.agent == user.display_name, CallingAssignment.assignment_date == working_date(), CallingAssignment.status == "completed")).all()
    output_dir = Path(settings.output_dir); output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / f"{user.username}_Completed_Calling_{working_date():%d-%m-%Y}.xlsx"
    export_assignments(session, rows, str(target), girish_format=user.display_name.strip().casefold() == "girish", include_feedback=True)
    return FileResponse(target, filename=target.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.post("/api/my-assignments/upload")
async def upload_my_assignments(file: UploadFile = File(...), authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or user.role == "admin": raise HTTPException(403, "Agent access required")
    if (datetime.now().hour, datetime.now().minute) >= (18, 20): raise HTTPException(403, "Excel upload is closed after 6:20 PM")
    if not (file.filename or "").casefold().endswith((".xlsx", ".xls")): raise HTTPException(422, "Upload an Excel calling sheet")
    feedback_dir = Path("5_Agent_Feedback/Completed_Calling_Sheets"); feedback_dir.mkdir(parents=True, exist_ok=True)
    target = feedback_dir / f"{date.today():%Y%m%d}_{user.username}_{Path(file.filename).name}"
    target.write_bytes(await file.read())
    master = Path("1_Master_NDR/OCT - MONTH NDR.xlsx")
    if not master.exists(): raise HTTPException(404, "Master NDR workbook not found")
    result = reverse_sync_feedback(session, str(master), str(feedback_dir), date.today())
    if result.get("accepted", 0) == 0 and result.get("duplicates", 0) > 0: raise HTTPException(409, "This calling file is already uploaded")
    return {"uploaded": target.name, **result}

@app.get("/api/my-completed")
def my_completed(authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user: raise HTTPException(401, "Login required")
    query = select(CallingAssignment).where(CallingAssignment.assignment_date == working_date(), CallingAssignment.status == "completed")
    if user.role != "admin": query = query.where(CallingAssignment.agent == user.display_name)
    result = []
    for assignment in session.scalars(query).all():
        order = session.get(Order, assignment.order_id)
        attempt = session.scalar(select(CallingAttempt).where(CallingAttempt.assignment_id == assignment.assignment_id).order_by(CallingAttempt.attempted_at.desc()))
        if not order or not attempt: continue
        result.append({"assignment_id": assignment.assignment_id, "agent": assignment.agent, "docket_no": order.docket_no, "order_no": order.order_no, "customer_name": order.customer_name, "mobile": order.mobile, "disposition": attempt.disposition, "remarks": attempt.remarks, "attempted_at": attempt.attempted_at})
    return result

@app.get("/api/dashboard")
def dashboard(session: Session = Depends(db)):
    today = date.today()
    active = active_orders(session, today)
    if manual_mode_enabled():
        manual_ids = {item.order_id for item in session.scalars(select(CallingAssignment).where(CallingAssignment.assignment_date == today, CallingAssignment.source == "manual-ndr-upload")).all()}
        active = [item for item in active if item.id in manual_ids]
    assigned = session.scalar(select(func.count()).select_from(CallingAssignment).where(CallingAssignment.assignment_date == today, CallingAssignment.status != "test")) or 0
    completed = session.scalar(select(func.count()).select_from(CallingAttempt).where(func.date(CallingAttempt.attempted_at) == today)) or 0
    pending_cases = session.scalar(select(func.count()).select_from(RTOCase).where(RTOCase.store_received.is_(False))) or 0
    pending_orders = 0
    for order in session.scalars(select(Order)).all():
        status = str(order.current_status or "").casefold()
        remark = str(order.courier_remark or "").casefold()
        if any(token in status or token in remark for token in ("return", "refund", "rto")):
            pending_orders += 1
    return {"active_orders": len(active), "today_assigned": assigned, "today_completed": completed, "pending_rto": pending_cases + pending_orders}

def active_orders(session: Session, today: date):
    excluded = ("delivered", "cancelled", "returned", "refunded", "self collect", "rto")
    result = []
    for order in session.scalars(select(Order)).all():
        status = str(order.current_status or "").casefold().strip(); shipment = str(order.shipment or "").casefold().strip()
        if not status or "shipped" not in status or any(word in status for word in excluded) or order.rto_active: continue
        if shipment == "delhivery":
            if not order.pdd or order.pdd > today or order.courier_attempts <= 0: continue
        elif not order.our_edd or order.our_edd > today: continue
        if not (order.mobile or order.alternate): continue
        result.append(order)
    return result

@app.post("/api/import/oms")
async def upload_oms(file: UploadFile = File(...), session: Session = Depends(db)):
    target = Path(settings.input_dir) / file.filename; target.parent.mkdir(parents=True, exist_ok=True); target.write_bytes(await file.read())
    try: return import_oms(session, str(target), load_rules())
    except ValueError as exc: raise HTTPException(422, str(exc))

@app.post("/api/import/oms-pair")
async def upload_oms_pair(order_movement: UploadFile = File(...), order_history: UploadFile = File(...), session: Session = Depends(db)):
    input_dir = Path(settings.input_dir); input_dir.mkdir(parents=True, exist_ok=True)
    movement_path = input_dir / order_movement.filename; history_path = input_dir / order_history.filename
    movement_path.write_bytes(await order_movement.read()); history_path.write_bytes(await order_history.read())
    session.query(CallingAttempt).delete(synchronize_session=False)
    session.query(CallingAssignment).delete(synchronize_session=False)
    session.query(Order).delete(synchronize_session=False)
    session.commit()
    try: return import_oms_pair(session, str(movement_path), str(history_path), load_rules())
    except ValueError as exc: raise HTTPException(422, str(exc))

@app.post("/api/import/mis")
async def upload_mis(files: list[UploadFile] = File(...), session: Session = Depends(db)):
    input_dir = Path(settings.input_dir); input_dir.mkdir(parents=True, exist_ok=True); results = []
    for file in files:
        target = input_dir / file.filename; target.write_bytes(await file.read())
        try:
            results.append({"file": file.filename, **reconcile_mis(session, str(target), load_rules())})
        except ValueError as exc:
            raise HTTPException(422, f"{file.filename}: {exc}")
    return {"files": results, "matched": sum(item.get("matched", 0) for item in results), "unmatched": sum(item.get("unmatched", 0) for item in results), "next_step": "Click Run today's automation after all MIS files finish uploading."}

@app.post("/api/admin/import-manual-ndr")
async def upload_manual_ndr(file: UploadFile | None = File(default=None), files: list[UploadFile] = File(default=[]), authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or user.role != "admin": raise HTTPException(403, "Admin access required")
    uploads = list(files)
    if file is not None: uploads.insert(0, file)
    if not uploads: raise HTTPException(422, "Upload at least one manual NDR Excel file")
    input_dir = Path(settings.input_dir); input_dir.mkdir(parents=True, exist_ok=True)
    frames = []
    filenames = []
    for upload in uploads:
        content = await upload.read()
        if not content: continue
        filename = upload.filename or "manual-ndr.xlsx"
        (input_dir / f"manual_{filename}").write_bytes(content)
        frames.append(pd.read_excel(BytesIO(content), header=None, dtype=object))
        filenames.append(filename)
    if not frames: raise HTTPException(422, "Uploaded manual NDR files are empty")
    frame = frames[0]
    header_row = next((index for index in range(min(5, len(frame))) if {str(value).strip().casefold() for value in frame.iloc[index].tolist()} >= {"docketno", "agent for ndr"}), None)
    if header_row is None:
        header_row = next((index for index in range(min(5, len(frame))) if frame.iloc[index].astype(str).str.casefold().eq("docketno").any()), None)
    if header_row is None: raise HTTPException(422, "Manual NDR must contain a Docketno column")
    if len(frames) > 1:
        combined_rows = [frame]
        canonical_headers = [str(value).strip().casefold() for value in frame.iloc[header_row].tolist()]
        for extra in frames[1:]:
            extra_header = next((index for index in range(min(5, len(extra))) if frame.iloc[header_row].astype(str).str.strip().str.casefold().tolist() == extra.iloc[index].astype(str).str.strip().str.casefold().tolist()), None)
            extra_header = extra_header if extra_header is not None else header_row
            extra_headers = [str(value).strip().casefold() for value in extra.iloc[extra_header].tolist()]
            target_indexes = {}
            for index, value in enumerate(canonical_headers):
                target_indexes.setdefault(value, []).append(index)
            normalized_rows = []
            for _, extra_row in extra.iloc[extra_header + 1 :].iterrows():
                normalized = [None] * len(canonical_headers)
                seen_headers = {}
                for index, value in enumerate(extra_headers):
                    occurrence = seen_headers.get(value, 0)
                    seen_headers[value] = occurrence + 1
                    matches = target_indexes.get(value, [])
                    if occurrence < len(matches) and index < len(extra_row): normalized[matches[occurrence]] = extra_row.iloc[index]
                normalized_rows.append(normalized)
            combined_rows.append(pd.DataFrame(normalized_rows, columns=frame.columns))
        frame = pd.concat(combined_rows, ignore_index=True)
    headers = [str(value).strip().casefold() for value in frame.iloc[header_row].tolist()]
    base_columns = {}
    for index, value in enumerate(headers):
        base_columns.setdefault(value, index)
    def cell(row, name):
        index = base_columns.get(name)
        if index is None or index >= len(row):
            return None
        value = row.iloc[index]
        return None if pd.isna(value) else value
    def text_cell(row, name):
        value = cell(row, name)
        return "" if value is None else str(value).strip()
    def date_cell(row, name):
        value = cell(row, name)
        parsed = pd.to_datetime(value, errors="coerce") if value is not None else None
        return parsed.date() if parsed is not None and not pd.isna(parsed) else None
    def last_column(name):
        matches = [index for index, value in enumerate(headers) if value == name]
        return matches[-1] if matches else None
    docket_column = last_column("docketno")
    order_column = last_column("orderno")
    date_columns = [(index, value.date() if hasattr(value, "date") else value) for index, value in enumerate(frame.iloc[0].tolist()) if hasattr(value, "date")]
    eligible_dates = [(index, value) for index, value in date_columns if value <= date.today()]
    date_column, assignment_date = max(eligible_dates, key=lambda item: item[1]) if eligible_dates else (None, date.today())
    agent_columns = [index for index, value in enumerate(headers) if value == "agent for ndr"]
    agent_column = date_column if date_column in agent_columns else (agent_columns[-1] if agent_columns else None)
    orders = {normalize_docket(item.docket_no): item for item in session.scalars(select(Order)).all()}
    orders.update({normalize_docket(item.order_no): item for item in session.scalars(select(Order)).all() if item.order_no})
    manual_assignments = session.scalars(select(CallingAssignment).where(CallingAssignment.source == "manual-ndr-upload")).all()
    current_assignments = session.scalars(select(CallingAssignment).where(CallingAssignment.assignment_date == assignment_date)).all()
    candidate_assignments = {item.assignment_id: item for item in [*manual_assignments, *current_assignments]}
    feedback_assignment_ids = {item.assignment_id for item in session.scalars(select(CallingAttempt).where(CallingAttempt.assignment_id.in_(candidate_assignments))).all()} if candidate_assignments else set()
    assignments_to_remove = {assignment_id: item for assignment_id, item in candidate_assignments.items() if assignment_id not in feedback_assignment_ids and item.status != "completed"}
    current_assignment_ids = list(assignments_to_remove)
    if current_assignment_ids:
        session.query(CallingAttempt).filter(CallingAttempt.assignment_id.in_(current_assignment_ids)).delete(synchronize_session=False)
        session.query(CallingAssignment).filter(CallingAssignment.assignment_id.in_(current_assignment_ids)).delete(synchronize_session=False)
    session.commit()
    existing = set()
    sequence = 1
    label_map = {
        "girish": "Girish",
        "pavi": "Pavi",
        "zaffar": "Mohammed Zaffruddin",
        "narayan": "Narayanaswamy K.V",
        "rubby": "Rubby Srinivasan",
        "saqlain": "Saqlain Muhammad",
    }
    roster_path = Path("3_Config_and_Inputs/agent_roster.json")
    roster = json.loads(roster_path.read_text(encoding="utf-8")) if roster_path.exists() else {}
    fallback_agents = [name for name, config in roster.items() if config.get("active", True) and config.get("present", True)]
    if not fallback_agents: fallback_agents = ["Pavi", "Mohammed Zaffruddin", "Saqlain Muhammad", "Rubby Srinivasan", "Narayanaswamy K.V"]
    accepted = skipped = unmatched = missing_agent = 0
    for _, row in frame.iloc[header_row + 1:].iterrows():
        identifier = normalize_docket(row.iloc[docket_column]) if docket_column is not None else (normalize_docket(row.iloc[order_column]) if order_column is not None else "")
        order = orders.get(identifier)
        if not order:
            docket_value = text_cell(row, "docketno") or text_cell(row, "orderno")
            if not docket_value:
                unmatched += 1; continue
            order = Order(order_no=text_cell(row, "orderno") or docket_value, docket_no=docket_value)
            session.add(order); session.flush()
            orders[normalize_docket(docket_value)] = order
        order.warehouse = text_cell(row, "warehouse") or order.warehouse
        order.order_no = text_cell(row, "orderno") or order.order_no
        order.order_date = date_cell(row, "order date") or order.order_date
        order.customer_name = text_cell(row, "cx name") or order.customer_name
        order.dx_mail = text_cell(row, "dx mail") or order.dx_mail
        order.cx_mail = text_cell(row, "cx mail") or order.cx_mail
        order.mobile = text_cell(row, "mobile no") or order.mobile
        order.alternate = text_cell(row, "alt no") or order.alternate
        order.shipping_date = date_cell(row, "shippingdate") or order.shipping_date
        order.shipment = text_cell(row, "shipment") or order.shipment
        order.pincode_zone = text_cell(row, "pincodezone") or order.pincode_zone
        order.our_edd = date_cell(row, "our edd") or order.our_edd
        order.pdd = date_cell(row, "pdd") or order.pdd
        order.current_status = text_cell(row, "current status") or order.current_status
        order.courier_remark = text_cell(row, "courier remark") or order.courier_remark
        if order.id in existing:
            skipped += 1; continue
        agent = str(row.iloc[agent_column]).strip() if agent_column is not None and pd.notna(row.iloc[agent_column]) else ""
        agent = label_map.get(agent.casefold(), agent)
        if not agent:
            missing_agent += 1
            continue
        session.add(CallingAssignment(assignment_id=f"CALL-{assignment_date:%Y%m%d}-{sequence:04d}", order_id=order.id, agent=agent, assignment_date=assignment_date, source="manual-ndr-upload"))
        existing.add(order.id); sequence += 1; accepted += 1
    session.commit()
    set_manual_mode(True, ", ".join(filenames) or "manual NDR", assignment_date)
    return {"date": assignment_date, "selected_date_block": assignment_date, "accepted": accepted, "replaced": len(current_assignment_ids), "skipped": skipped, "unmatched": unmatched, "missing_agent": missing_agent, "files": filenames, "mode": "manual-only"}

@app.post("/api/run-automation")
def run_automation(authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or user.role != "admin": raise HTTPException(403, "Admin access required")
    return run_daily_automation(session)

@app.post("/api/import/self-collect")
async def upload_self_collect(file: UploadFile = File(...), session: Session = Depends(db)):
    target = Path("3_Config_and_Inputs/self_collect_orders.txt"); target.parent.mkdir(parents=True, exist_ok=True); target.write_bytes(await file.read())
    master = Path("1_Master_NDR/OCT - MONTH NDR.xlsx")
    return process_self_collect(session, str(master) if master.exists() else None, str(target))

class SelfCollectManual(BaseModel):
    identifier: str

@app.post("/api/self-collect/manual")
def mark_self_collect_manual(payload: SelfCollectManual, session: Session = Depends(db)):
    identifier = payload.identifier.strip()
    if not identifier: raise HTTPException(422, "OrderNo or Docketno is required")
    target = Path("3_Config_and_Inputs/self_collect_orders.txt"); target.parent.mkdir(parents=True, exist_ok=True)
    existing = target.read_text(encoding="utf-8") if target.exists() else ""
    if identifier.casefold() not in {line.strip().casefold() for line in existing.splitlines()}: target.write_text(existing.rstrip() + ("\n" if existing.strip() else "") + identifier + "\n", encoding="utf-8")
    master = Path("1_Master_NDR/OCT - MONTH NDR.xlsx")
    result = process_self_collect(session, str(master) if master.exists() else None, str(target))
    if not result["tagged"]: raise HTTPException(404, "OrderNo/Docketno not found in Master NDR")
    return result

@app.post("/api/sync/feedback")
def sync_feedback_now(session: Session = Depends(db)):
    master = Path("1_Master_NDR/OCT - MONTH NDR.xlsx"); feedback_dir = Path("5_Agent_Feedback/Completed_Calling_Sheets")
    if not master.exists(): raise HTTPException(404, "Master NDR workbook not found")
    return reverse_sync_feedback(session, str(master), str(feedback_dir), date.today())

@app.get("/api/export/master-ndr")
def download_master_ndr(authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or user.role != "admin": raise HTTPException(403, "Admin access required")
    master = Path("1_Master_NDR/OCT - MONTH NDR.xlsx")
    if not master.exists(): raise HTTPException(404, "Master NDR workbook not found")
    output_dir = Path(settings.output_dir); output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / f"Master_NDR_Check_{date.today():%d-%m-%Y}.xlsx"
    input_files = [*Path(settings.input_dir).glob("Reports_OrderMovement*.xlsx"), *Path(settings.input_dir).glob("Reports_OrderHistory*.xlsx"), *Path(settings.input_dir).glob("*.csv"), *Path(settings.input_dir).glob("Collaged MIS*.xlsx")]
    if target.exists() and (not input_files or target.stat().st_mtime >= max(path.stat().st_mtime for path in input_files)):
        return FileResponse(target, filename=target.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    shutil.copy2(master, target)
    workbook = load_workbook(target)
    sheet = workbook.active
    for row_number in range(3, sheet.max_row + 1):
        for cell in sheet[row_number]: cell.value = None
    workbook.save(target)
    update_master_ndr(str(target), session.scalars(select(Order)).all(), date.today())
    return FileResponse(target, filename=target.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/api/orders")
def orders(session: Session = Depends(db)):
    visible = active_orders(session, working_date())
    if manual_mode_enabled():
        manual_ids = {item.order_id for item in session.scalars(select(CallingAssignment).where(CallingAssignment.assignment_date == working_date(), CallingAssignment.source == "manual-ndr-upload")).all()}
        visible = [item for item in visible if item.id in manual_ids]
    return [{"docket_no": o.docket_no, "order_no": o.order_no, "warehouse": o.warehouse, "shipment": o.shipment, "status": o.current_status, "our_edd": o.our_edd, "pdd": o.pdd, "rto_active": o.rto_active} for o in visible]

@app.get("/api/assignments")
def assignments(agent: str, session: Session = Depends(db)):
    rows = []
    for assignment in session.scalars(select(CallingAssignment).where(CallingAssignment.agent == agent, CallingAssignment.assignment_date == working_date(), CallingAssignment.status == "pending")).all():
        order = session.get(Order, assignment.order_id)
        status = str(order.current_status or "").casefold()
        manual_assignment = assignment.source == "manual-ndr-upload"
        eligible_status = "shipped" in status and not any(token in status for token in ("delivered", "return", "refund", "cancel", "self collect", "rto"))
        if order and (manual_assignment or eligible_status): rows.append({"assignment_id": assignment.assignment_id, "agent": agent, "order_id": order.id, "docket_no": order.docket_no, "order_no": order.order_no, "customer_name": order.customer_name, "mobile": order.mobile, "alternate": order.alternate, "dx_mail": order.dx_mail, "cx_mail": order.cx_mail, "shipping_date": order.shipping_date, "courier_remark": order.courier_remark, "shipment": order.shipment})
    return rows

class RTOReceiveUpdate(BaseModel):
    docket: str; received: bool = False; received_date: date | None = None

def rto_access(user: User | None) -> bool:
    return bool(user and (user.role == "admin" or user.display_name.casefold() in {"pavi", "girish"}))

@app.post("/api/rto/receive")
def update_rto_receiving(payload: RTOReceiveUpdate, authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not rto_access(user): raise HTTPException(403, "RTO access is limited to Admin, Pavi and Girish")
    docket = payload.docket.strip()
    case = session.scalar(select(RTOCase).where(RTOCase.docket_no == docket))
    order = session.scalar(select(Order).where(Order.docket_no == docket))
    if not case and not order: raise HTTPException(404, "Docket not found")
    if not case:
        case = RTOCase(docket_no=docket, courier=order.shipment, warehouse=order.warehouse, initiation_date=order.shipping_date or date.today(), courier_status=order.current_status or "Returned / Refunded")
        session.add(case)
    case.store_received = payload.received
    case.store_received_date = payload.received_date or date.today()
    case.followup_remarks = "Product received" if payload.received else "Product not received"
    session.commit()
    return {"docket": docket, "received": case.store_received, "received_date": case.store_received_date}

@app.get("/api/rto")
def rto(authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not rto_access(user): raise HTTPException(403, "RTO access is limited to Admin, Pavi and Girish")
    today = date.today(); result = []
    seen = set()
    for case in session.scalars(select(RTOCase).where(RTOCase.store_received.is_(False))).all():
        pending = max(0, (today - case.initiation_date).days)
        priority = "Urgent Mismatch" if "DELIVERED" in case.courier_status.upper() else "High Priority" if pending >= 5 else "Courier Follow-Up" if pending >= 3 else "Normal Tracking"
        seen.add(case.docket_no); result.append({"docket": case.docket_no, "courier": case.courier, "warehouse": case.warehouse, "status": case.courier_status, "return_value": "Return" if "refund" in case.courier_status.casefold() or "return" in case.courier_status.casefold() else "", "pending_days": pending, "priority": priority, "store_received": case.store_received, "received_date": case.store_received_date})
    for order in session.scalars(select(Order)).all():
        status = str(order.current_status or "").casefold(); remark = str(order.courier_remark or "").casefold()
        is_return = any(token in status for token in ("return", "refund", "rto")) or any(token in remark for token in ("return", "refund", "rto"))
        if not is_return or order.docket_no in seen: continue
        start = order.shipping_date or today; pending = max(0, (today - start).days)
        priority = "Urgent Mismatch" if "rto delivered" in status and not order.rto_active else "High Priority" if pending >= 5 else "Courier Follow-Up" if pending >= 3 else "Returned / Refunded"
        result.append({"docket": order.docket_no, "courier": order.shipment, "warehouse": order.warehouse, "status": order.current_status, "return_value": "Return" if "refund" in status or "return" in status else "", "pending_days": pending, "priority": priority, "store_received": False, "received_date": None})
    return result

class Feedback(BaseModel):
    assignment_id: str; agent: str; disposition: str; remarks: str = ""; attempted_at: datetime | None = None

@app.post("/api/feedback")
def feedback(payload: Feedback, authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or (user.role != "admin" and user.display_name != payload.agent): raise HTTPException(403, "Login identity does not match agent")
    if user.role != "admin" and (datetime.now().hour, datetime.now().minute) >= (18, 20): raise HTTPException(403, "Online feedback editing is closed after 6:20 PM")
    assignment = session.scalar(select(CallingAssignment).where(CallingAssignment.assignment_id == payload.assignment_id))
    if not assignment or assignment.agent != payload.agent: raise HTTPException(403, "Assignment ownership validation failed")
    attempt_id = f"ATT-{payload.assignment_id}-{payload.attempted_at or datetime.utcnow()}"
    if session.scalar(select(CallingAttempt).where(CallingAttempt.attempt_id == attempt_id)): return {"duplicate": True}
    session.add(CallingAttempt(attempt_id=attempt_id, assignment_id=payload.assignment_id, agent=payload.agent, disposition=payload.disposition, remarks=payload.remarks, attempted_at=payload.attempted_at or datetime.utcnow())); assignment.status = "completed"; session.commit(); return {"duplicate": False, "attempt_id": attempt_id}

@app.put("/api/feedback/{assignment_id}")
def edit_feedback(assignment_id: str, payload: Feedback, authorization: str | None = Header(default=None), session: Session = Depends(db)):
    user = current_user(authorization, session)
    if not user or (user.role != "admin" and user.display_name != payload.agent): raise HTTPException(403, "Login identity does not match agent")
    if user.role != "admin" and (datetime.now().hour, datetime.now().minute) >= (18, 20): raise HTTPException(403, "Feedback editing is closed after 6:20 PM")
    assignment = session.scalar(select(CallingAssignment).where(CallingAssignment.assignment_id == assignment_id))
    if not assignment or (user.role != "admin" and assignment.agent != user.display_name): raise HTTPException(403, "Assignment ownership validation failed")
    attempt = session.scalar(select(CallingAttempt).where(CallingAttempt.assignment_id == assignment_id).order_by(CallingAttempt.attempted_at.desc()))
    if not attempt: raise HTTPException(404, "Completed feedback not found")
    attempt.disposition = payload.disposition
    attempt.remarks = payload.remarks
    if payload.attempted_at: attempt.attempted_at = payload.attempted_at
    session.commit()
    return {"updated": True, "assignment_id": assignment_id}
