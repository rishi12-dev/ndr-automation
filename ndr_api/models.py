from datetime import datetime, date
from sqlalchemy import String, Integer, Date, DateTime, Boolean, Text, UniqueConstraint, ForeignKey, Index
from sqlalchemy.orm import Mapped, mapped_column
from .db import Base

class Order(Base):
    __tablename__ = "orders"
    id: Mapped[int] = mapped_column(primary_key=True)
    order_no: Mapped[str] = mapped_column(String(100), index=True)
    docket_no: Mapped[str] = mapped_column(String(100), unique=True, index=True)
    warehouse: Mapped[str] = mapped_column(String(120), default="")
    order_date: Mapped[date | None] = mapped_column(Date)
    shipping_date: Mapped[date | None] = mapped_column(Date)
    customer_name: Mapped[str] = mapped_column(String(200), default="")
    mobile: Mapped[str] = mapped_column(String(40), default="")
    alternate: Mapped[str] = mapped_column(String(40), default="")
    dx_mail: Mapped[str] = mapped_column(String(255), default="")
    cx_mail: Mapped[str] = mapped_column(String(255), default="")
    shipment: Mapped[str] = mapped_column(String(120), default="")
    pincode_zone: Mapped[str] = mapped_column(String(100), default="")
    our_edd: Mapped[date | None] = mapped_column(Date)
    pdd: Mapped[date | None] = mapped_column(Date)
    current_status: Mapped[str] = mapped_column(String(120), default="")
    courier_remark: Mapped[str] = mapped_column(Text, default="")
    courier_attempts: Mapped[int] = mapped_column(Integer, default=0)
    latest_scan_timestamp: Mapped[datetime | None] = mapped_column(DateTime)
    rto_active: Mapped[bool] = mapped_column(Boolean, default=False)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class User(Base):
    __tablename__ = "users"
    id: Mapped[int] = mapped_column(primary_key=True)
    username: Mapped[str] = mapped_column(String(80), unique=True, index=True)
    password_hash: Mapped[str] = mapped_column(String(255))
    display_name: Mapped[str] = mapped_column(String(150))
    role: Mapped[str] = mapped_column(String(30), default="agent")
    active: Mapped[bool] = mapped_column(Boolean, default=True)

class CallingAssignment(Base):
    __tablename__ = "calling_assignments"
    id: Mapped[int] = mapped_column(primary_key=True)
    assignment_id: Mapped[str] = mapped_column(String(80), unique=True, index=True)
    order_id: Mapped[int] = mapped_column(ForeignKey("orders.id"))
    agent: Mapped[str] = mapped_column(String(150), index=True)
    assignment_date: Mapped[date] = mapped_column(Date, index=True)
    source: Mapped[str] = mapped_column(String(40))
    status: Mapped[str] = mapped_column(String(30), default="pending")
    __table_args__ = (UniqueConstraint("order_id", "assignment_date", name="uq_order_assignment_day"),)

class CallingAttempt(Base):
    __tablename__ = "calling_attempts"
    id: Mapped[int] = mapped_column(primary_key=True)
    attempt_id: Mapped[str] = mapped_column(String(80), unique=True, index=True)
    assignment_id: Mapped[str] = mapped_column(ForeignKey("calling_assignments.assignment_id"))
    agent: Mapped[str] = mapped_column(String(150))
    disposition: Mapped[str] = mapped_column(String(80))
    remarks: Mapped[str] = mapped_column(Text, default="")
    attempted_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class RTOCase(Base):
    __tablename__ = "rto_cases"
    id: Mapped[int] = mapped_column(primary_key=True)
    docket_no: Mapped[str] = mapped_column(String(100), unique=True, index=True)
    courier: Mapped[str] = mapped_column(String(120))
    warehouse: Mapped[str] = mapped_column(String(120))
    initiation_date: Mapped[date] = mapped_column(Date)
    courier_status: Mapped[str] = mapped_column(String(80), default="RTO Initiated")
    store_received: Mapped[bool] = mapped_column(Boolean, default=False)
    store_received_date: Mapped[date | None] = mapped_column(Date)
    followup_owner: Mapped[str] = mapped_column(String(150), default="")
    followup_remarks: Mapped[str] = mapped_column(Text, default="")

class AuditLog(Base):
    __tablename__ = "audit_logs"
    id: Mapped[int] = mapped_column(primary_key=True)
    action: Mapped[str] = mapped_column(String(100), index=True)
    entity: Mapped[str] = mapped_column(String(100), default="")
    entity_id: Mapped[str] = mapped_column(String(100), default="")
    details: Mapped[str] = mapped_column(Text, default="{}")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class ImportBatch(Base):
    __tablename__ = "import_batches"
    id: Mapped[int] = mapped_column(primary_key=True)
    filename: Mapped[str] = mapped_column(String(255))
    kind: Mapped[str] = mapped_column(String(50))
    file_hash: Mapped[str] = mapped_column(String(64), unique=True)
    imported_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    status: Mapped[str] = mapped_column(String(30), default="completed")
    errors: Mapped[str] = mapped_column(Text, default="")
