from datetime import date
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from ndr_api.db import Base
from ndr_api.core import normalize_docket, canonical_status, zone_edd, import_oms, eligible_orders, allocate, smart_allocate
from ndr_api.models import Order, CallingAssignment
from ndr_api.models import RTOCase
from ndr_api.automation import import_feedback, rto_report

RULES = {"zones": {"Within Region": 4, "Unknown": 1}, "zone_aliases": {"Within Region": ["Zone B"]}, "columns": {"warehouse": ["Warehouse"], "order_no": ["OrderNo"], "docket_no": ["Docketno"], "order_date": ["Order Date"], "shipping_date": ["Shipping Date"], "customer_name": ["Customer Name"], "mobile": ["Mobile Number"], "alternate": ["Alternate Number"], "shipment": ["Shipment"], "pincode_zone": ["PincodeZone"], "current_status": ["Current Status"]}}

def session(tmp_path):
    engine = create_engine(f"sqlite:///{tmp_path}/test.db"); Base.metadata.create_all(engine); return sessionmaker(bind=engine)()

def test_docket_and_status():
    assert normalize_docket("12345678901234567.0") == "12345678901234567"
    assert canonical_status("Delivered", "RTO Delivered") == "RTO Delivered"

def test_edd_zone():
    assert zone_edd(date(2026, 1, 1), "Zone B", RULES)[0] == date(2026, 1, 5)

def test_oms_import_is_idempotent(tmp_path):
    frame = pd.DataFrame([{ "Warehouse":"Chennai", "OrderNo":"O1", "Docketno":"00123", "Order Date":"2026-01-01", "Shipping Date":"2026-01-01", "Customer Name":"Test", "Mobile Number":"999", "Alternate Number":"", "Shipment":"DTDC", "PincodeZone":"Zone B", "Current Status":"NDR" }]); path = tmp_path / "oms.csv"; frame.to_csv(path, index=False)
    db = session(tmp_path); assert import_oms(db, str(path), RULES)["inserted"] == 1; assert import_oms(db, str(path), RULES)["duplicate"]
    assert db.query(Order).count() == 1

def test_local_excludes_delhivery(tmp_path):
    db = session(tmp_path); db.add_all([Order(order_no="1", docket_no="1", shipment="Delhivery", current_status="Shipped", our_edd=date.today(), mobile="1"), Order(order_no="2", docket_no="2", shipment="DTDC", current_status="Shipped", our_edd=date.today(), mobile="1")]); db.commit()
    _, local = eligible_orders(db, date.today(), RULES); assert [row.docket_no for row in local] == ["2"]

def test_duplicate_assignment_prevented(tmp_path):
    db = session(tmp_path); order = Order(order_no="1", docket_no="1", shipment="DTDC"); db.add(order); db.commit(); assert len(allocate(db, [order], ["Pavi"], date.today(), "local")) == 1; assert len(allocate(db, [order], ["Pavi"], date.today(), "local")) == 0

def test_feedback_import_is_idempotent(tmp_path):
    db = session(tmp_path); order = Order(order_no="1", docket_no="1", shipment="DTDC"); db.add(order); db.commit(); assignment = allocate(db, [order], ["Pavi"], date.today(), "local")[0]
    path = tmp_path / "feedback.csv"; pd.DataFrame([{"assignment_id": assignment.assignment_id, "agent_for_ndr": "Pavi", "attempt_id": "ATT-1", "disposition": "No Answer", "our_remarks": "Retry"}]).to_csv(path, index=False)
    assert import_feedback(db, str(path))["accepted"] == 1; assert import_feedback(db, str(path))["duplicate"]

def test_rto_urgent_mismatch(tmp_path):
    db = session(tmp_path); db.add(RTOCase(docket_no="R1", courier="DTDC", warehouse="Chennai", initiation_date=date.today().replace(day=max(1, date.today().day - 5)), courier_status="RTO Delivered", store_received=False)); db.commit()
    report = rto_report(db, date.today(), str(tmp_path)); assert report["urgent"] == 1; assert report["pending_5_plus"] == 1

def test_smart_allocation_dedicated_couriers(tmp_path):
    db = session(tmp_path); orders = [Order(order_no=str(i), docket_no=str(i), shipment=courier, our_edd=date.today(), mobile="1") for i, courier in enumerate(["DTDC", "Chennai Xpress", "Nandan"], 1)]; db.add_all(orders); db.commit()
    assignments = smart_allocate(db, [], orders, {"Narayanaswamy K.V": {"active": True, "present": True}, "Pavi": {"active": True, "present": True}, "Mohammed Zaffruddin": {"active": True, "present": True}}, date.today())
    by_courier = {db.get(Order, item.order_id).shipment: item.agent for item in assignments}; assert by_courier["DTDC"] == "Narayanaswamy K.V"; assert by_courier["Chennai Xpress"] in {"Pavi", "Mohammed Zaffruddin"}
