from datetime import date, datetime, timedelta
from pathlib import Path
import hashlib, re, uuid
import pandas as pd
from sqlalchemy import select
from sqlalchemy.orm import Session
from .config import load_rules
from .models import Order, CallingAssignment, CallingAttempt, RTOCase, ImportBatch

def normalize_docket(value) -> str:
    if pd.isna(value): return ""
    text = format(value, ".0f") if isinstance(value, float) and value.is_integer() else str(value).strip()
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+", text):
        text = format(float(text), ".0f")
    if re.fullmatch(r"\d+\.0", text): text = text[:-2]
    return text

def canonical_status(status: str, remark: str = "") -> str:
    text = f"{status} {remark}".upper()
    if "RTO" in text and "DELIVER" in text: return "RTO Delivered"
    if "DELIVERED TO CONSIGNEE" in text or "DELIVERED (POD)" in text or text.strip() == "DELIVERED": return "Delivered"
    if "CANCEL" in text: return "Cancelled"
    if "RETURN" in text or "REFUND" in text: return "Returned / Refunded"
    if "SELF COLLECT" in text: return "Self Collect"
    return status or ""

def zone_edd(shipping_date, zone: str, rules: dict):
    if not shipping_date: return None
    normalize_zone = lambda value: " ".join(str(value or "").replace("\xa0", " ").split()).casefold()
    aliases = {normalize_zone(group): group for group in rules.get("zones", {})}
    aliases.update({normalize_zone(alias): group for group, values in rules.get("zone_aliases", {}).items() for alias in values})
    group = aliases.get(normalize_zone(zone), "Unknown")
    return shipping_date + timedelta(days=int(rules["zones"].get(group, 1))), group

def _date(value):
    if value is None or pd.isna(value) or str(value).strip() == "": return None
    return pd.to_datetime(value, errors="coerce").date()

def read_table(path: str) -> pd.DataFrame:
    return pd.read_csv(path, dtype=object) if path.lower().endswith(".csv") else pd.read_excel(path, dtype=object)

def map_columns(frame: pd.DataFrame, aliases: dict) -> pd.DataFrame:
    lower = {str(column).strip().casefold(): column for column in frame.columns}
    mapped = {}
    for canonical, names in aliases.items():
        for name in names:
            if str(name).casefold() in lower:
                mapped[canonical] = lower[str(name).casefold()]; break
    missing = [key for key in aliases if key not in mapped]
    if missing: raise ValueError(f"Missing required/expected columns: {', '.join(missing)}")
    return frame.rename(columns={source: target for target, source in mapped.items()})

def import_oms(session: Session, path: str, rules: dict, kind="monthly") -> dict:
    digest = hashlib.sha256(Path(path).read_bytes()).hexdigest()
    if session.scalar(select(ImportBatch).where(ImportBatch.file_hash == digest)): return {"duplicate": True, "inserted": 0, "updated": 0}
    oms_fields = {key: value for key, value in rules["columns"].items() if key not in {"pdd", "courier_remark", "courier_attempts"}}
    frame = map_columns(read_table(path), oms_fields)
    inserted = updated = 0
    for _, row in frame.iterrows():
        docket = normalize_docket(row.get("docket_no"))
        if not docket: continue
        order = session.scalar(select(Order).where(Order.docket_no == docket))
        shipping = _date(row.get("shipping_date")); edd, _ = zone_edd(shipping, row.get("pincode_zone"), rules)
        values = dict(order_no=str(row.get("order_no", "") or ""), docket_no=docket, warehouse=str(row.get("warehouse", "") or ""), order_date=_date(row.get("order_date")), shipping_date=shipping, customer_name=str(row.get("customer_name", "") or ""), mobile=str(row.get("mobile", "") or ""), alternate=str(row.get("alternate", "") or ""), shipment=str(row.get("shipment", "") or ""), pincode_zone=str(row.get("pincode_zone", "") or ""), our_edd=edd, current_status=str(row.get("current_status", "") or ""))
        if order:
            for key, value in values.items():
                if key not in {"current_status"} or value: setattr(order, key, value)
            updated += 1
        else: session.add(Order(**values)); inserted += 1
    session.add(ImportBatch(filename=Path(path).name, kind=kind, file_hash=digest)); session.commit()
    return {"duplicate": False, "inserted": inserted, "updated": updated}

def import_oms_pair(session: Session, movement_path: str, history_path: str, rules: dict) -> dict:
    """Import the fixed daily OMS exports used by the operations team.

    Movement columns: D,F,G,K,L,M,V,Z. History columns: C,G,AX,H,I,J.
    Movement is the source of shipment, docket, zone and status; History adds
    customer and contact fields by OrderNo.
    """
    movement = read_table(movement_path)
    history = read_table(history_path)
    if movement.shape[1] < 8 or history.shape[1] < 50:
        raise ValueError("OMS files do not match the expected Order Movement / Order History format")
    def text(value):
        if pd.isna(value): return ""
        return str(value).strip()
    def series(frame, names, fallback):
        normalized = {re.sub(r"[^a-z0-9]", "", str(name).casefold()): index for index, name in enumerate(frame.columns)}
        for name in names:
            index = normalized.get(re.sub(r"[^a-z0-9]", "", name.casefold()))
            if index is not None: return frame.iloc[:, index]
        return frame.iloc[:, fallback - 1] if fallback <= frame.shape[1] else pd.Series([""] * len(frame), index=frame.index)
    history_order = series(history, ["orderno"], 3)
    history_name = series(history, ["name", "customername"], 7)
    history_dx_mail = series(history, ["userlogin", "dxmail"], 50)
    history_cx_mail = series(history, ["emailid", "cxmail"], 8)
    history_mobile = series(history, ["mobileno", "mobile"], 9)
    history_alternate = series(history, ["alternatemobileno", "alternate"], 10)
    movement_order = series(movement, ["orderno"], 6)
    movement_docket = series(movement, ["docketno", "waybill"], 12)
    movement_zone = series(movement, ["zone", "pincodezone"], 13)
    movement_pincode_zone = series(movement, ["pincodezone", "zone"], 14)
    movement_shipping = series(movement, ["shippingdate"], 24)
    movement_status = series(movement, ["presentstatus", "currentstatus", "status"], 28)
    movement_warehouse = series(movement, ["warehouse", "warehouse"], 4)
    movement_order_date = series(movement, ["orderdate"], 7)
    movement_shipment = series(movement, ["shipment"], 11)
    history_by_order = {}
    for _, item in history.iterrows():
        key = text(history_order[_])
        if not key: continue
        history_by_order.setdefault(key, {
            "customer_name": text(history_name[_]), "dx_mail": text(history_dx_mail[_]),
            "cx_mail": text(history_cx_mail[_]), "mobile": text(history_mobile[_]),
            "alternate": text(history_alternate[_]),
        })
    orders = {order.docket_no: order for order in session.scalars(select(Order)).all()}
    inserted = updated = 0
    for _, item in movement.iterrows():
        order_no = text(movement_order[_]); docket = normalize_docket(movement_docket[_])
        if not docket: continue
        details = history_by_order.get(order_no, {})
        shipping = _date(movement_shipping[_]); zone = text(movement_zone[_])
        if not zone or zone.casefold() in {"#n/a", "nan", "none"}: zone = text(movement_pincode_zone[_])
        edd_result = zone_edd(shipping, zone, rules); edd = edd_result[0] if edd_result else None
        values = dict(order_no=order_no, docket_no=docket, warehouse=text(movement_warehouse[_]), order_date=_date(movement_order_date[_]), shipping_date=shipping, customer_name=details.get("customer_name", ""), mobile=details.get("mobile", ""), alternate=details.get("alternate", ""), dx_mail=details.get("dx_mail", ""), cx_mail=details.get("cx_mail", ""), shipment=text(movement_shipment[_]), pincode_zone=zone, our_edd=edd, current_status=text(movement_status[_]))
        order = orders.get(docket)
        if order:
            for key, value in values.items():
                if value != "" and value is not None: setattr(order, key, value)
            updated += 1
        else:
            order = Order(**values); session.add(order); orders[docket] = order; inserted += 1
    combined_hash = hashlib.sha256((Path(movement_path).read_bytes() + Path(history_path).read_bytes())).hexdigest()
    if not session.scalar(select(ImportBatch).where(ImportBatch.file_hash == combined_hash)):
        session.add(ImportBatch(filename=f"{Path(movement_path).name}+{Path(history_path).name}", kind="oms-pair", file_hash=combined_hash))
    session.commit(); return {"duplicate": False, "inserted": inserted, "updated": updated, "source": "Order Movement + Order History"}

def reconcile_mis(session: Session, path: str, rules: dict) -> dict:
    frame = read_table(path); frame.columns = [str(c).strip().casefold() for c in frame.columns]
    is_collaged = len(frame.columns) >= 2 and frame.columns[0] in {"docket no", "docketno"} and frame.columns[1] == "status"
    is_delhivery_hq = len(frame.columns) >= 24 and frame.columns[0] == "waybill"
    if is_delhivery_hq and frame.iloc[:, 0].astype(str).str.contains(r"[eE][+-]?\d+", regex=True, na=False).any():
        raise ValueError("Delhivery Waybill column is truncated in scientific notation. Export/download the file with full Waybill numbers before uploading.")
    if not is_collaged and not is_delhivery_hq:
        docket_col = next((c for c in frame.columns if c in {"docketno", "docket no", "docket", "tracking_number", "tracking number"}), None)
        if not docket_col: raise ValueError("Courier MIS requires a docket/tracking column")
    changed = unmatched = 0; matched_dockets = set(); orders = {item.docket_no: item for item in session.scalars(select(Order)).all()}
    for _, row in frame.iterrows():
        docket = normalize_docket(row.iloc[0] if (is_collaged or is_delhivery_hq) else row[docket_col]); order = orders.get(docket)
        if not order: unmatched += 1; continue
        matched_dockets.add(order.docket_no)
        if is_collaged:
            status = str(row.iloc[1] or "").strip()
            if status and status.casefold() not in {"nan", "#n/a"}:
                order.current_status = canonical_status(status)
                order.courier_remark = status
        elif is_delhivery_hq:
            pdd = _date(row.iloc[8]); remark = str(row.iloc[21] or "").strip(); attempts = row.iloc[23]
            if pdd: order.pdd = pdd
            if str(attempts).strip().casefold() not in {"", "nan", "none", "#n/a"}:
                try: order.courier_attempts = int(float(attempts))
                except (TypeError, ValueError): pass
            current = str(order.current_status or "").strip().casefold()
            if current == "shipped" and remark.casefold() == "#n/a": remark = "Waiting for status"
            elif current in {"packed", "pending", "progress"}: remark = "#N/A"
            if remark: order.courier_remark = remark
        else:
            for key, candidates in {"pdd":["pdd"], "courier_remark":["courier remark", "latest courier remark", "remark"], "courier_attempts":["courier attempts", "attempts"], "current_status":["current status", "status"], "latest_scan_timestamp":["latest scan timestamp", "scan timestamp"]}.items():
                source = next((c for c in candidates if c in frame.columns), None)
                if source and not pd.isna(row[source]):
                    value = _date(row[source]) if key == "pdd" else (int(float(row[source])) if key == "courier_attempts" and str(row[source]).replace('.','',1).isdigit() else str(row[source]))
                    setattr(order, key, value)
        order.current_status = canonical_status(order.current_status, order.courier_remark); order.rto_active = "RTO" in f"{order.current_status} {order.courier_remark}".upper() and order.current_status != "RTO Delivered"; changed += 1
    if is_collaged or is_delhivery_hq:
        courier_orders = session.scalars(select(Order)).all()
        for order in courier_orders:
            is_target = (is_delhivery_hq and str(order.shipment or "").casefold() == "delhivery") or (is_collaged and str(order.shipment or "").casefold() != "delhivery")
            if not is_target or order.docket_no in matched_dockets: continue
            current = str(order.current_status or "").casefold()
            if any(token in current for token in ("delivered", "cancel", "return", "refund")): continue
            order.courier_remark = "Waiting for status" if current == "shipped" else "#N/A"
    session.commit(); return {"matched": changed, "unmatched": unmatched}

def eligible_orders(session: Session, today: date, rules: dict, absent=None, previous_girish_dockets=None):
    absent = set(absent or []); rows = session.scalars(select(Order)).all(); blocked = {"Delivered", "Cancelled", "Returned / Refunded", "Self Collect"}
    previous_girish = set(previous_girish_dockets or [])
    local, girish = [], []
    for order in rows:
        if order.rto_active or order.current_status in blocked or "shipped" not in str(order.current_status or "").casefold(): continue
        raw_zone = str(order.pincode_zone or "").strip()
        zone_result = zone_edd(order.shipping_date or today, raw_zone, rules)
        if raw_zone and (not zone_result or zone_result[1] == "Unknown"): continue
        if order.shipment.casefold() == "delhivery" and order.docket_no not in previous_girish and order.pdd and date(today.year, today.month, 1) <= order.pdd <= today and order.courier_attempts > 0 and "Girish" not in absent: girish.append(order)
        elif order.shipment.casefold() != "delhivery" and order.our_edd == today and (order.mobile or order.alternate): local.append(order)
    return girish, local

def allocate(session: Session, orders, agents, today: date, source: str, commit: bool = True):
    todays = session.scalars(select(CallingAssignment).where(CallingAssignment.assignment_date == today)).all()
    existing_ids = {assignment.order_id for assignment in todays}
    counts = {agent: sum(assignment.agent == agent and assignment.status != "test" for assignment in todays) for agent in agents}
    sequence = len(todays) + 1
    created = []
    for order in orders:
        if order.id in existing_ids: continue
        agent = min(agents, key=lambda name: counts[name]) if agents else "Supervisor Review"; counts[agent] = counts.get(agent, 0) + 1
        assignment = CallingAssignment(assignment_id=f"CALL-{today:%Y%m%d}-{sequence:04d}", order_id=order.id, agent=agent, assignment_date=today, source=source); session.add(assignment); created.append(assignment)
        existing_ids.add(order.id); sequence += 1
    if commit: session.commit()
    return created

def smart_allocate(session: Session, girish_orders, local_orders, roster: dict, today: date):
    """Apply the work-plan's dedicated-courier priorities before general balancing."""
    def unique_orders(items):
        seen = set(); result = []
        for order in items:
            key = str(order.order_no or order.docket_no).strip().casefold()
            if key not in seen: seen.add(key); result.append(order)
        return result
    girish_orders = unique_orders(girish_orders); local_orders = unique_orders(local_orders)
    present = {name for name, config in roster.items() if config.get("active", True) and config.get("present", True)}
    result = []
    girish = ["Girish"] if "Girish" in present else [name for name in ("Pavi", "Mohammed Zaffruddin", "Saqlain Muhammad") if name in present]
    if girish_orders: result.extend(allocate(session, girish_orders, girish, today, "girish", commit=False))
    groups = {"dtdc": [], "priority": [], "general": []}
    for order in local_orders:
        courier = order.shipment.casefold()
        if courier == "dtdc": groups["dtdc"].append(order)
        elif courier in {"chennai xpress", "professional couriers"}: groups["priority"].append(order)
        else: groups["general"].append(order)
    narayan = ["Narayanaswamy K.V"] if "Narayanaswamy K.V" in present else [name for name in ("Pavi", "Saqlain Muhammad") if name in present]
    result.extend(allocate(session, groups["dtdc"], narayan, today, "local-dtdc", commit=False))
    priority = [name for name in ("Pavi", "Mohammed Zaffruddin", "Saqlain Muhammad") if name in present]
    result.extend(allocate(session, groups["priority"], priority, today, "local-priority", commit=False))
    general = [name for name in roster if name in present and name != "Girish"]
    result.extend(allocate(session, groups["general"], general, today, "local-general", commit=False))
    session.commit()
    return result

def export_assignments(session: Session, assignments, path: str, girish_format: bool = False, include_feedback: bool = False):
    rows = []
    agent_labels = {"Mohammed Zaffruddin": "Zaffar", "Narayanaswamy K.V": "Narayan"}
    rules = load_rules()
    for assignment in assignments:
        order = session.get(Order, assignment.order_id)
        raw_zone = str(order.pincode_zone or "").strip()
        zone_result = zone_edd(order.shipping_date or date.today(), raw_zone, rules)
        if raw_zone and (not zone_result or zone_result[1] == "Unknown"): continue
        canonical_zone = zone_result[1] if raw_zone and zone_result else ""
        is_delhivery = str(order.shipment or "").casefold() == "delhivery"
        common = [order.warehouse, str(order.order_no or ""), order.order_date, order.customer_name, order.dx_mail, order.cx_mail, order.mobile, order.alternate, order.shipping_date, order.shipment, canonical_zone, str(order.docket_no or ""), order.our_edd]
        feedback = None
        if include_feedback:
            attempt = session.scalar(select(CallingAttempt).where(CallingAttempt.assignment_id == assignment.assignment_id).order_by(CallingAttempt.attempted_at.desc()))
            if attempt: feedback = f"[{attempt.disposition}] {attempt.remarks}".strip()
        if girish_format:
            rows.append(common + [order.pdd, order.current_status, order.current_status, order.courier_remark, agent_labels.get(assignment.agent, assignment.agent), feedback, order.courier_attempts])
        else:
            rows.append(common + [order.current_status, order.current_status, order.courier_remark, agent_labels.get(assignment.agent, assignment.agent), feedback])
    delhivery = girish_format
    headers = ["Warehouse", "OrderNo", "Order Date", "Cx Name", "DX MAIL", "CX MAIL", "Mobile No", "Alt No", "ShippingDate", "Shipment", "PincodeZone", "Docketno", "OUR EDD"] + (["PDD", "OMS STATUS", "Current status", "Courier remark", "Agent for NDR", "Our Remarks", "Attempts"] if delhivery else ["OMS STATUS", "Current status", "Courier remark", "Agent for NDR", "Our Remarks"])
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    workbook = Workbook(); sheet = workbook.active; sheet.title = "Calling Sheet"
    agent_column = headers.index("Agent for NDR") + 1
    sheet.append([None] * (len(headers))); sheet.cell(1, agent_column, date.today())
    sheet.append(headers)
    for row in rows: sheet.append(row)
    header_map = {name: index + 1 for index, name in enumerate(headers)}
    for column_name in ("Order Date", "ShippingDate", "OUR EDD", "PDD"):
        if column_name not in header_map: continue
        for row in range(3, sheet.max_row + 1): sheet.cell(row, header_map[column_name]).number_format = "dd/mmm/yyyy"
    yellow = PatternFill("solid", fgColor="FFFF00"); thin = Side(style="thin", color="B7B7B7"); border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = border; cell.alignment = Alignment(vertical="center", wrap_text=True)
    for cell in sheet[2]:
        cell.font = Font(bold=True, color="000000"); cell.fill = yellow; cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cell in sheet[1]: cell.fill = yellow; cell.border = border
    for column_name in ("OrderNo", "Docketno"):
        column = header_map[column_name]
        for row in range(3, sheet.max_row + 1): sheet.cell(row, column).number_format = "@"
    widths = {"Warehouse": 20, "OrderNo": 15, "Order Date": 14, "Cx Name": 26, "DX MAIL": 28, "CX MAIL": 28, "Mobile No": 15, "Alt No": 15, "ShippingDate": 14, "Shipment": 24, "PincodeZone": 18, "Docketno": 18, "OUR EDD": 14, "PDD": 14, "OMS STATUS": 16, "Current status": 18, "Courier remark": 30, "Agent for NDR": 18, "Our Remarks": 34, "Attempts": 10}
    for name, width in widths.items():
        if name in header_map: sheet.column_dimensions[sheet.cell(2, header_map[name]).column_letter].width = width
    sheet.row_dimensions[1].height = 22; sheet.row_dimensions[2].height = 42
    for row in range(3, sheet.max_row + 1): sheet.row_dimensions[row].height = 38
    sheet.freeze_panes = "A3"; sheet.auto_filter.ref = f"A2:{sheet.cell(sheet.max_row, len(headers)).coordinate}"
    workbook.save(path); workbook.close(); return path
