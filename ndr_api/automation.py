from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import hashlib
import json
import re
import shutil
import uuid

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from sqlalchemy import select, func
from sqlalchemy.orm import Session

from .config import settings
from .core import normalize_docket, _date, canonical_status
from .models import Order, CallingAssignment, CallingAttempt, RTOCase, ImportBatch, User

STATUS_COLORS = {
    "Delivered": "006100",
    "Returned / Refunded": "FF0000",
    "Cancelled": "FFC000",
    "Self Collect": "92D050",
}

def process_self_collect(session: Session, master_path: str | None, input_path: str = "3_Config_and_Inputs/self_collect_orders.txt") -> dict:
    ids = set(); path = Path(input_path)
    if path.exists(): ids = {line.strip().casefold() for line in path.read_text(encoding="utf-8").splitlines() if line.strip() and not line.lstrip().startswith("#")}
    tagged = []
    for order in session.scalars(select(Order)).all():
        searchable = {str(order.order_no or "").strip().casefold(), str(order.docket_no or "").strip().casefold()}
        remark = str(order.courier_remark or "").casefold()
        if searchable & ids or any(token in remark for token in ("self collect", "store pickup", "customer collect")):
            order.current_status = "Self Collect"; order.rto_active = False; tagged.append(order)
    session.commit()
    backup = None
    if master_path and tagged:
        backup = safe_backup(master_path); workbook = load_workbook(master_path); sheet = workbook.active; headers = _header_map(sheet)
        docket_col = headers.get("docketno"); status_col = headers.get("current status") or headers.get("oms status"); tagged_ids = {order.docket_no for order in tagged}
        for row in range(3, sheet.max_row + 1):
            if docket_col and normalize_docket(sheet.cell(row, docket_col).value) in tagged_ids:
                for column in range(1, min(20, sheet.max_column) + 1): sheet.cell(row, column).fill = PatternFill("solid", fgColor=STATUS_COLORS["Self Collect"])
                if status_col: sheet.cell(row, status_col, "Self Collect")
        temp = Path(master_path).with_suffix(f".tmp{Path(master_path).suffix}"); workbook.save(temp); temp.replace(master_path)
    return {"tagged": len(tagged), "backup": backup}

def validate_calling_outputs(girish_path: str | None, local_path: str | None, today: date) -> dict:
    frames = []
    for path in (girish_path, local_path):
        if path and Path(path).exists(): frames.append(pd.read_excel(path, dtype=object))
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(); local_delhivery = 0
    if local_path and Path(local_path).exists():
        local = pd.read_excel(local_path, dtype=object); local_delhivery = int(local["Shipment"].astype(str).str.casefold().eq("delhivery").sum())
    if local_delhivery: raise ValueError(f"Local NDR validation failed: {local_delhivery} Delhivery rows")
    if not combined.empty and combined["OrderNo"].astype(str).duplicated().any(): raise ValueError("Duplicate OrderNo across calling files")
    if girish_path and Path(girish_path).exists():
        girish = pd.read_excel(girish_path, dtype=object); attempts = pd.to_numeric(girish["Attempts"], errors="coerce").fillna(0); pdd = pd.to_datetime(girish["PDD"], errors="coerce").dt.date
        if (attempts <= 0).any() or (pdd > today).any(): raise ValueError("Girish NDR validation failed: attempts/PDD rule")
    return {"total_active": len(combined), "girish": len(pd.read_excel(girish_path)) if girish_path and Path(girish_path).exists() else 0, "local": len(pd.read_excel(local_path)) if local_path and Path(local_path).exists() else 0, "local_delhivery": local_delhivery}

DISPOSITIONS = ["DELIV_REQ", "FUTURE_DATE", "CUST_REFUSED", "WRONG_ADDR", "FAKE_ATTEMPT", "NO_ANSWER", "BUSY", "DISCONNECTED", "SWITCHED_OFF", "SELF_COLLECT"]

def prepare_agent_sheet(path: str) -> str:
    workbook = load_workbook(path); sheet = workbook.active
    headers = {str(sheet.cell(1, column).value).strip(): column for column in range(1, sheet.max_column + 1) if sheet.cell(1, column).value}
    disposition_column = headers.get("Call Disposition") or sheet.max_column + 1; remarks_column = headers.get("Our Remarks") or sheet.max_column + 1; date_column = headers.get("Rescheduled Date") or sheet.max_column + 1
    sheet.cell(1, disposition_column, "Call Disposition"); sheet.cell(1, remarks_column, "Our Remarks"); sheet.cell(1, date_column, "Rescheduled Date")
    validation = DataValidation(type="list", formula1='"DELIV_REQ,FUTURE_DATE,CUST_REFUSED,WRONG_ADDR,FAKE_ATTEMPT,NO_ANSWER,BUSY,DISCONNECTED,SWITCHED_OFF,SELF_COLLECT"', allow_blank=True); sheet.add_data_validation(validation); validation.add(f"{sheet.cell(2, disposition_column).column_letter}2:{sheet.cell(sheet.max_row, disposition_column).column_letter}{sheet.max_row}")
    for row in range(2, sheet.max_row + 1):
        for column in range(1, sheet.max_column + 1): sheet.cell(row, column).protection = sheet.cell(row, column).protection.copy(locked=column not in {disposition_column, remarks_column, date_column})
    sheet.protection.sheet = True; sheet.protection.password = "ndr-agent"
    workbook.save(path); return path

def reverse_sync_feedback(session: Session, master_path: str, feedback_dir: str, today: date) -> dict:
    files = sorted(Path(feedback_dir).glob("*.xlsx")); if_not = Path("6_Logs_and_Reports/calling_attempt_history.csv"); if_not.parent.mkdir(parents=True, exist_ok=True)
    seen = set(if_not.read_text(encoding="utf-8").splitlines()) if if_not.exists() else set(); accepted = duplicates = rejected = 0; audit = []
    backup = safe_backup(master_path); workbook = load_workbook(master_path); sheet = workbook.active; headers = _header_map(sheet); docket_column = headers.get("docketno"); attempts_column = headers.get("attempts")
    date_column = next((column for column in range(1, sheet.max_column + 1) if getattr(sheet.cell(1, column).value, "date", lambda: None)() == today), None)
    if date_column is None:
        date_column = sheet.max_column + 1; sheet.insert_cols(date_column, 3); sheet.cell(1, date_column, today); sheet.cell(2, date_column, "Agent for NDR"); sheet.cell(2, date_column + 1, "Our Remarks"); sheet.cell(2, date_column + 2, "Call Disposition"); attempts_column = None
    agent_column = date_column; remarks_column = date_column + 1; row_by_docket = {normalize_docket(sheet.cell(row, docket_column).value): row for row in range(3, sheet.max_row + 1)}
    for path in files:
        preview = pd.read_excel(path, header=None, nrows=2, dtype=object)
        header_row = 1 if any(str(value).strip().casefold() == "orderno" for value in preview.iloc[1].tolist()) else 0
        frame = pd.read_excel(path, header=header_row, dtype=object); frame.columns = [str(column).strip().casefold().replace(" ", "_") for column in frame.columns]
        for _, record in frame.iterrows():
            assignment = str(record.get("assignment_id", "")).strip(); order_no = str(record.get("orderno", "")).strip(); remarks = str(record.get("our_remarks", "")).strip(); disposition = str(record.get("call_disposition", "")).strip().upper()
            if not disposition and remarks:
                match = re.match(r"\[([^\]]+)\]\s*(.*)", remarks); disposition = match.group(1).strip().upper() if match else "OTHER"
                remarks = match.group(2).strip() if match else remarks
            digest = hashlib.sha256(f"{assignment}|{order_no}|{disposition}|{remarks}".encode()).hexdigest()
            if digest in seen: duplicates += 1; continue
            if not assignment or not disposition or not remarks: rejected += 1; continue
            docket = normalize_docket(str(record.get("docketno", ""))); row = row_by_docket.get(docket)
            if not row: rejected += 1; continue
            assignment_row = session.scalar(select(CallingAssignment).where(CallingAssignment.assignment_id == assignment)) if assignment else None
            if not assignment_row and docket:
                order = session.scalar(select(Order).where(Order.docket_no == docket))
                if order: assignment_row = session.scalar(select(CallingAssignment).where(CallingAssignment.order_id == order.id, CallingAssignment.assignment_date == today))
            if not assignment_row: rejected += 1; continue
            rescheduled = str(record.get("rescheduled_date", "")).strip(); formatted = f"[{disposition}] {remarks}" + (f" (Rescheduled: {rescheduled})" if rescheduled and rescheduled.casefold() != "nan" else "")
            sheet.cell(row, agent_column, str(record.get("agent_for_ndr", assignment_row.agent))); sheet.cell(row, remarks_column, formatted)
            if attempts_column: sheet.cell(row, attempts_column, int(sheet.cell(row, attempts_column).value or 0) + 1)
            assignment_row.status = "completed"
            if not session.scalar(select(CallingAttempt).where(CallingAttempt.attempt_id == f"XLS-{digest}")):
                session.add(CallingAttempt(attempt_id=f"XLS-{digest}", assignment_id=assignment_row.assignment_id, agent=assignment_row.agent, disposition=disposition, remarks=remarks, attempted_at=datetime.utcnow()))
            seen.add(digest); audit.append(digest); accepted += 1
    temp = Path(master_path).with_suffix(f".tmp{Path(master_path).suffix}"); workbook.save(temp); temp.replace(master_path); session.commit(); if_not.write_text("\n".join(sorted(seen)) + "\n", encoding="utf-8")
    return {"files": len(files), "accepted": accepted, "duplicates": duplicates, "rejected": rejected, "backup": backup}

def previous_working_day_girish_dockets(path: str, today: date) -> set[str]:
    workbook = load_workbook(path, read_only=True, data_only=True); sheet = workbook.active
    date_columns = [(column, sheet.cell(1, column).value) for column in range(1, sheet.max_column + 1) if hasattr(sheet.cell(1, column).value, "date")]
    previous = [(column, value.date()) for column, value in date_columns if value.date() < today]
    if not previous: workbook.close(); return set()
    start_column, _ = max(previous, key=lambda item: item[1]); agent_column = start_column + 1; docket_column = 12
    excluded = set()
    for values in sheet.iter_rows(min_row=3, values_only=True):
        agent = values[agent_column - 1] if len(values) >= agent_column else None
        if str(agent or "").strip().casefold() == "girish":
            docket = normalize_docket(values[docket_column - 1] if len(values) >= docket_column else None)
            if docket: excluded.add(docket)
    workbook.close(); return excluded

def import_master_workbook(session: Session, path: str) -> dict:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    header_row = next(sheet.iter_rows(min_row=2, max_row=2, values_only=True))
    headers = {str(value).strip().casefold(): index for index, value in enumerate(header_row, 1) if value is not None}
    required = {"docketno", "orderno", "shipment"}
    if not required.issubset(headers): raise ValueError("Master workbook is missing Docketno, OrderNo or Shipment")
    inserted = updated = 0
    def value(row, name):
        index = headers.get(name)
        return row[index - 1] if index else None
    def safe_date(item):
        if item is None or str(item).casefold() in {"nan", "nat", "none", ""}: return None
        converted = item.date() if hasattr(item, "date") else _date(item)
        return None if str(converted).casefold() in {"nan", "nat", "none"} else converted
    def safe_int(item):
        if item is None or str(item).casefold() in {"nan", "nat", "none", ""}: return 0
        try: return int(float(item))
        except (TypeError, ValueError): return 0
    for row in sheet.iter_rows(min_row=3, values_only=True):
        docket = normalize_docket(value(row, "docketno"))
        if not docket: continue
        order = session.scalar(select(Order).where(Order.docket_no == docket))
        shipping = safe_date(value(row, "shippingdate")); edd = safe_date(value(row, "our edd")); pdd = safe_date(value(row, "pdd"))
        attempt = value(row, "attempts")
        values = dict(order_no=str(value(row, "orderno") or ""), docket_no=docket, warehouse=str(value(row, "warehouse") or ""), order_date=safe_date(value(row, "order date")), shipping_date=shipping, customer_name=str(value(row, "cx name") or ""), mobile=str(value(row, "mobile no") or ""), alternate=str(value(row, "alt no") or ""), dx_mail=str(value(row, "dx mail") or ""), cx_mail=str(value(row, "cx mail") or ""), shipment=str(value(row, "shipment") or ""), pincode_zone=str(value(row, "pincodezone") or ""), our_edd=edd, pdd=pdd, current_status=str(value(row, "current status") or value(row, "oms status") or ""), courier_remark=str(value(row, "courier remark") or ""), courier_attempts=safe_int(attempt))
        if order:
            for key, item in values.items(): setattr(order, key, item)
            updated += 1
        else: session.add(Order(**values)); inserted += 1
    session.commit(); workbook.close(); return {"inserted": inserted, "updated": updated}

def sync_master_assignments(session: Session, path: str, today: date) -> dict:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    header_row = next(sheet.iter_rows(min_row=2, max_row=2, values_only=True))
    headers = {str(value).strip().casefold(): index for index, value in enumerate(header_row, 1) if value is not None}
    docket_column = headers.get("docketno")
    if not docket_column:
        workbook.close(); return {"created": 0, "updated": 0}
    date_column = next((column for column in range(1, sheet.max_column + 1) if getattr(sheet.cell(1, column).value, "date", lambda: None)() == today), None)
    if not date_column:
        workbook.close(); return {"created": 0, "updated": 0}
    agent_offset = 0 if str(sheet.cell(2, date_column).value or "").strip().casefold() == "agent for ndr" else 1
    users = session.scalars(select(User).where(User.role == "agent", User.active.is_(True))).all()
    def resolve_agent(value):
        name = str(value or "").strip()
        folded = name.casefold()
        for user in users:
            choices = {user.username.casefold(), user.display_name.casefold(), user.display_name.split()[0].casefold()}
            if folded in choices:
                return user.display_name
        return name
    orders = {order.docket_no: order for order in session.scalars(select(Order)).all()}
    existing = {assignment.order_id: assignment for assignment in session.scalars(select(CallingAssignment).where(CallingAssignment.assignment_date == today)).all()}
    master_order_ids = set()
    sequence = len(existing) + 1
    created = updated = 0
    for values in sheet.iter_rows(min_row=3, values_only=True):
        docket = normalize_docket(values[docket_column - 1] if len(values) >= docket_column else None)
        agent_index = date_column - 1 + agent_offset
        agent = resolve_agent(values[agent_index] if len(values) > agent_index else None)
        if not docket or not agent or agent.casefold() in {"nan", "none"} or docket not in orders:
            continue
        order = orders[docket]
        master_order_ids.add(order.id)
        assignment = existing.get(order.id)
        if assignment:
            if assignment.status == "test" or assignment.source != "master-workbook" or assignment.agent != agent:
                assignment.agent = agent; assignment.source = "master-workbook"
                if assignment.status == "test": assignment.status = "pending"
                updated += 1
            continue
        assignment = CallingAssignment(assignment_id=f"MASTER-{today:%Y%m%d}-{sequence:04d}", order_id=order.id, agent=agent, assignment_date=today, source="master-workbook")
        session.add(assignment); existing[order.id] = assignment; sequence += 1
        created += 1
    for order_id, assignment in existing.items():
        if order_id not in master_order_ids and assignment.status != "test":
            assignment.status = "test"
            updated += 1
    session.commit(); workbook.close(); return {"created": created, "updated": updated}

def safe_backup(path: str, backup_dir: str | None = None) -> str:
    source = Path(path)
    if not source.exists():
        raise FileNotFoundError(path)
    target_dir = Path(backup_dir or settings.backup_dir)
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / f"{source.stem}_{datetime.now():%Y%m%d_%H%M%S}{source.suffix}"
    shutil.copy2(source, target)
    return str(target)

def _header_map(sheet):
    result = {}
    for row in sheet.iter_rows(min_row=1, max_row=min(4, sheet.max_row)):
        for cell in row:
            if cell.value is not None:
                result[str(cell.value).strip().casefold()] = cell.column
    return result

def update_master_ndr(path: str, orders: list[Order], today: date, backup_dir: str | None = None) -> dict:
    """Safely update a workbook without deleting historical columns or rows.

    The updater supports both simple header-based sheets and the common two-row
    daily-group layout. Unknown layouts are rejected rather than overwritten.
    """
    backup = safe_backup(path, backup_dir)
    workbook = load_workbook(path)
    sheet = workbook.active
    headers = _header_map(sheet)
    required = {"docketno", "orderno"}
    if not required.issubset(headers):
        raise ValueError("Master NDR requires Docketno and OrderNo headers")
    row_by_docket = {}
    available_rows = []
    template_max_row = sheet.max_row
    for row in range(2, sheet.max_row + 1):
        docket = normalize_docket(sheet.cell(row, headers["docketno"]).value)
        if docket: row_by_docket[docket] = row
        elif row >= 3: available_rows.append(row)
    today_text = today.strftime("%d-%m-%Y")
    group_labels = ["Agent for NDR", "Our Remarks", "Courier remark"]
    date_columns = {}
    for column in range(1, sheet.max_column + 1):
        date_values = [sheet.cell(row, column).value for row in (1, 2)]
        if any(today_text in str(value) or (hasattr(value, "strftime") and value.strftime("%d-%m-%Y") == today_text) for value in date_values if value is not None):
            for offset, label in enumerate(group_labels): date_columns[label.casefold()] = column + offset
            break
    if not date_columns:
        attempts_column = next((column for column in range(1, sheet.max_column + 1) if str(sheet.cell(1, column).value or "").strip().casefold() in {"attempts", "courier attempts", "global courier attempts"}), sheet.max_column + 1)
        sheet.insert_cols(attempts_column, amount=3)
        for offset, label in enumerate(group_labels):
            column = attempts_column + offset
            sheet.cell(1, column, today)
            sheet.cell(1, column).number_format = "dd/mmm/yyyy"
            sheet.cell(2, column, label)
            date_columns[label.casefold()] = column
        headers = _header_map(sheet)
    grid_side = Side(style="thin", color="B7B7B7")
    grid_border = Border(left=grid_side, right=grid_side, top=grid_side, bottom=grid_side)
    created = updated = 0
    for order in orders:
        row = row_by_docket.get(order.docket_no)
        if row is None:
            row = available_rows.pop(0) if available_rows else sheet.max_row + 1
            appended_row = row > template_max_row
            row_by_docket[order.docket_no] = row
            created += 1
            base_values = {
                "warehouse": order.warehouse, "orderno": order.order_no,
                "order date": order.order_date, "cx name": order.customer_name,
                "dx mail": order.dx_mail, "cx mail": order.cx_mail,
                "mobile no": order.mobile, "alt no": order.alternate,
                "shippingdate": order.shipping_date, "shipment": order.shipment,
                "pincodezone": order.pincode_zone, "docketno": order.docket_no,
                "our edd": order.our_edd, "pdd": order.pdd,
                "oms status": order.current_status, "current status": order.current_status,
            }
            for key, value in base_values.items():
                if key in headers and value is not None:
                    sheet.cell(row, headers[key], value)
            if appended_row:
                for column_number in range(1, sheet.max_column + 1):
                    cell = sheet.cell(row, column_number)
                    cell.border = grid_border
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                sheet.row_dimensions[row].height = 38
        else: updated += 1
        for key, value in {"pdd": order.pdd, "current status": order.current_status, "our edd": order.our_edd}.items():
            if key == "pdd" and str(order.shipment or "").casefold() != "delhivery":
                if key in headers: sheet.cell(row, headers[key], None)
                continue
            if key in headers and value is not None:
                cell = sheet.cell(row, headers[key], value)
                if key in {"pdd", "our edd"}: cell.number_format = "dd/mmm/yyyy"
        if "courier remark" in date_columns and order.courier_remark:
            sheet.cell(row, date_columns["courier remark"], order.courier_remark)
        attempts_header = next((key for key in ("attempts", "courier attempts", "global courier attempts") if key in headers), None)
        if attempts_header:
            sheet.cell(row, headers[attempts_header], (order.courier_attempts if order.courier_attempts > 0 else "#N/A") if str(order.shipment or "").casefold() == "delhivery" else None)
        normalized_status = canonical_status(str(order.current_status or ""), str(order.courier_remark or ""))
        if normalized_status in STATUS_COLORS:
            status_fill = PatternFill("solid", fgColor=STATUS_COLORS[normalized_status])
            color_columns = {headers.get("current status"), date_columns.get("courier remark")}
            for color_column in color_columns:
                if color_column:
                    sheet.cell(row, color_column).fill = status_fill
    for name in ("order date", "orderdate", "shippingdate", "shipping date", "our edd", "pdd"):
        if name in headers:
            for row in range(3, sheet.max_row + 1):
                sheet.cell(row, headers[name]).number_format = "dd/mmm/yyyy"
    for column in range(1, sheet.max_column + 1):
        if hasattr(sheet.cell(1, column).value, "date"):
            sheet.cell(1, column).number_format = "dd/mmm/yyyy"
    if "orderno" in headers:
        for row_number in range(3, sheet.max_row + 1): sheet.cell(row_number, headers["orderno"]).number_format = "@"
    if "docketno" in headers:
        for row_number in range(3, sheet.max_row + 1): sheet.cell(row_number, headers["docketno"]).number_format = "@"
    temp = Path(path).with_suffix(f".tmp{Path(path).suffix}")
    workbook.save(temp); temp.replace(path)
    return {"backup": backup, "created": created, "updated": updated, "daily_group": today_text, "daily_columns": date_columns}

def import_feedback(session: Session, path: str) -> dict:
    frame = pd.read_csv(path, dtype=object) if str(path).lower().endswith(".csv") else pd.read_excel(path, dtype=object)
    frame.columns = [str(column).strip().casefold().replace(" ", "_") for column in frame.columns]
    required = {"assignment_id", "agent_for_ndr"}
    missing = required - set(frame.columns)
    if missing: raise ValueError(f"Feedback missing columns: {', '.join(sorted(missing))}")
    digest = hashlib.sha256(Path(path).read_bytes()).hexdigest()
    if session.scalar(select(ImportBatch).where(ImportBatch.file_hash == digest)):
        return {"duplicate": True, "accepted": 0, "rejected": 0}
    accepted = rejected = 0
    for _, row in frame.iterrows():
        assignment_id = str(row.get("assignment_id", "")).strip()
        agent = str(row.get("agent_for_ndr", "")).strip()
        assignment = session.scalar(select(CallingAssignment).where(CallingAssignment.assignment_id == assignment_id))
        if not assignment or assignment.agent != agent: rejected += 1; continue
        attempt_id = str(row.get("attempt_id", "")).strip() or f"ATT-{assignment_id}-{row.get('attempt_timestamp', '')}-{row.get('disposition', '')}"
        if session.scalar(select(CallingAttempt).where(CallingAttempt.attempt_id == attempt_id)): continue
        attempted_at = pd.to_datetime(row.get("attempt_timestamp"), errors="coerce")
        session.add(CallingAttempt(attempt_id=attempt_id, assignment_id=assignment_id, agent=agent, disposition=str(row.get("disposition", "")).strip(), remarks=str(row.get("our_remarks", "")).strip(), attempted_at=attempted_at.to_pydatetime() if not pd.isna(attempted_at) else datetime.utcnow()))
        assignment.status = "completed"; accepted += 1
    session.add(ImportBatch(filename=Path(path).name, kind="feedback", file_hash=digest)); session.commit()
    return {"duplicate": False, "accepted": accepted, "rejected": rejected}

def rto_report(session: Session, today: date, output_dir: str | None = None) -> dict:
    output = Path(output_dir or settings.output_dir); output.mkdir(parents=True, exist_ok=True); rows = []
    for case in session.scalars(select(RTOCase)).all():
        pending = max(0, (today - case.initiation_date).days)
        priority = "RTO Closed" if case.store_received else ("Urgent Mismatch" if "DELIVERED" in case.courier_status.upper() else "High Priority" if pending >= 5 else "Courier Follow-Up" if pending >= 3 else "Normal Tracking")
        rows.append({"Docketno": case.docket_no, "Courier": case.courier, "Warehouse": case.warehouse, "Initiation Date": case.initiation_date, "Courier Status": case.courier_status, "Store Receiving": "Store Received" if case.store_received else "Store Not Received", "Pending Days": pending, "Priority": priority, "Owner": case.followup_owner, "Follow-up Remarks": case.followup_remarks})
    path = output / f"RTO_Followup_{today:%d-%m-%Y}.xlsx"; pd.DataFrame(rows).to_excel(path, index=False)
    return {"path": str(path), "total": len(rows), "urgent": sum(row["Priority"] == "Urgent Mismatch" for row in rows), "pending_3_plus": sum(row["Pending Days"] >= 3 and row["Priority"] != "RTO Closed" for row in rows), "pending_5_plus": sum(row["Pending Days"] >= 5 and row["Priority"] != "RTO Closed" for row in rows)}

def completion_report(session: Session, today: date) -> list[dict]:
    agents = session.scalars(select(CallingAssignment.agent).where(CallingAssignment.assignment_date == today).distinct()).all(); report = []
    for agent in agents:
        assignments = session.scalars(select(CallingAssignment).where(CallingAssignment.agent == agent, CallingAssignment.assignment_date == today)).all()
        completed = sum(assignment.status == "completed" for assignment in assignments)
        missing_remarks = session.scalar(select(func.count()).select_from(CallingAttempt).join(CallingAssignment, CallingAttempt.assignment_id == CallingAssignment.assignment_id).where(CallingAssignment.agent == agent, CallingAssignment.assignment_date == today, CallingAttempt.remarks == "")) or 0
        report.append({"agent": agent, "assigned": len(assignments), "completed": completed, "pending": len(assignments) - completed, "completion_pct": round(completed * 100 / len(assignments), 2) if assignments else 0, "missing_remarks": missing_remarks, "status": "Data Sync Pending" if not assignments else "OK"})
    return report
