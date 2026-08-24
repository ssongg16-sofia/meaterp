import io
import os
import random
from datetime import datetime, date, timedelta
from typing import List, Optional

import bcrypt
from fastapi import FastAPI, HTTPException, Depends, UploadFile, File, Query, status
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl.styles import Font, PatternFill, Alignment
from jose import JWTError, jwt
import openpyxl
from pydantic import BaseModel, ConfigDict
from sqlalchemy import (
    Column, Integer, Float, String, DateTime, ForeignKey, create_engine, desc, func, Index, text
)
from sqlalchemy.orm import declarative_base, sessionmaker, Session
import uvicorn

# -----------------------------------------------------------------------------
# 인증 및 환경 설정
# -----------------------------------------------------------------------------
JWT_SECRET_KEY = os.environ.get("JWT_SECRET_KEY", "meatflow-enterprise-secret-key-2026")
JWT_ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 8

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/auth/login", auto_error=False)

ROLE_ADMIN = "ADMIN"
ROLE_SALES = "SALES"
ROLE_WAREHOUSE = "WAREHOUSE"
ALL_ROLES = [ROLE_ADMIN, ROLE_SALES, ROLE_WAREHOUSE]

# -----------------------------------------------------------------------------
# 1. DB 설정 (SQLite / 환경 변수 확장 가능)
# -----------------------------------------------------------------------------
DB_FILE = os.environ.get("DB_FILE", "meaterp_local.db")
DATABASE_URL = os.environ.get("DATABASE_URL", f"sqlite:///{DB_FILE}")

# Render PostgreSQL URL 호환 처리 (postgres:// -> postgresql://)
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine = create_engine(DATABASE_URL, connect_args=connect_args)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()

# -----------------------------------------------------------------------------
# 2. ORM 테이블 모델
# -----------------------------------------------------------------------------
class CompanyMaster(Base):
    __tablename__ = "companies"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String(50), unique=True, nullable=False)
    name = Column(String(100), nullable=False)
    biz_no = Column(String(30), nullable=False)
    rep_name = Column(String(50), nullable=False)
    address = Column(String(200), nullable=False)
    phone = Column(String(30), nullable=True)
    fax = Column(String(30), nullable=True)
    created_at = Column(DateTime, default=datetime.now)

class Partner(Base):
    __tablename__ = "partners"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), nullable=False)
    type = Column(String(30), nullable=False, index=True)
    biz_no = Column(String(30), nullable=True)
    contact_person = Column(String(50), nullable=True)
    phone = Column(String(30), nullable=True)
    address = Column(String(200), nullable=True)
    created_at = Column(DateTime, default=datetime.now)

class ItemMaster(Base):
    __tablename__ = "item_masters"
    id = Column(Integer, primary_key=True, index=True)
    item_code = Column(String(50), unique=True, nullable=False)
    item_name = Column(String(100), nullable=False)
    species = Column(String(30), nullable=False)
    created_at = Column(DateTime, default=datetime.now)

class CutMaster(Base):
    __tablename__ = "cut_masters"
    id = Column(Integer, primary_key=True, index=True)
    item_id = Column(Integer, ForeignKey("item_masters.id", ondelete="CASCADE"), nullable=False, index=True)
    cut_code = Column(String(50), unique=True, nullable=False)
    cut_name = Column(String(100), nullable=False)
    default_storage = Column(String(20), default="냉장")
    created_at = Column(DateTime, default=datetime.now)

class InboundRecord(Base):
    __tablename__ = "inbounds"
    id = Column(Integer, primary_key=True, index=True)
    inbound_no = Column(String(50), unique=True, index=True)
    company_name = Column(String(100), default="주식회사 티제이에프")
    inbound_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d"), index=True)
    processed_date = Column(String(20), nullable=True, index=True)
    vendor = Column(String(100), nullable=False)
    bl_no = Column(String(50), nullable=True)
    trace_no = Column(String(50), nullable=False, index=True)
    est_no = Column(String(50), nullable=True)
    brand = Column(String(50), nullable=True)
    grade = Column(String(30), default="GF")
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
    process_from_date = Column(String(20), nullable=True)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    cost_per_kg = Column(Float, nullable=False)
    total_amount = Column(Float, nullable=False)
    warehouse = Column(String(50), default="광주냉장창고")
    exp_date = Column(String(20), nullable=False, index=True)
    is_weighed = Column(String(10), default="N")
    claim_reason = Column(String(255), nullable=True)
    status = Column(String(20), default="IN_REQUEST", index=True)
    grid_no = Column(String(50), unique=True, nullable=True, index=True)
    created_at = Column(DateTime, default=datetime.now)

class InventoryLot(Base):
    __tablename__ = "inventory_lots"
    id = Column(Integer, primary_key=True, index=True)
    company_name = Column(String(100), default="주식회사 티제이에프")
    sku_code = Column(String(50), nullable=False)
    inbound_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d"))
    bl_no = Column(String(50), nullable=True)
    trace_no = Column(String(50), nullable=False, index=True)
    est_no = Column(String(50), nullable=True)
    brand = Column(String(50), nullable=True)
    grade = Column(String(30), default="GF")
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
    process_from_date = Column(String(20), nullable=True)
    initial_box_qty = Column(Integer, default=0)
    initial_weight_kg = Column(Float, default=0.0)
    avg_box_weight = Column(Float, default=0.0)
    current_box_qty = Column(Integer, nullable=False)
    current_weight_kg = Column(Float, nullable=False)
    cost_per_kg = Column(Float, nullable=False)
    warehouse = Column(String(50), nullable=False)
    exp_date = Column(String(20), nullable=False, index=True)
    is_weighed = Column(String(10), default="N")
    grid_no = Column(String(50), unique=True, nullable=False, index=True)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

class OutboundRecord(Base):
    __tablename__ = "outbounds"
    id = Column(Integer, primary_key=True, index=True)
    outbound_no = Column(String(50), unique=True, index=True)
    company_name = Column(String(100), default="주식회사 티제이에프")
    vendor_chain = Column(String(100), nullable=True)
    trade_type = Column(String(30), default="출고판매")
    is_estimated = Column(String(10), default="N")
    inbound_date = Column(String(20), nullable=True)
    outbound_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d"), index=True)
    processed_date = Column(String(20), nullable=True, index=True)
    lot_id = Column(Integer, ForeignKey("inventory_lots.id"), nullable=False, index=True)
    customer = Column(String(100), nullable=False)
    bl_no = Column(String(50), nullable=True)
    trace_no = Column(String(50), nullable=False)
    est_no = Column(String(50), nullable=True)
    brand = Column(String(50), nullable=True)
    grade = Column(String(30), default="GF")
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
    process_from_date = Column(String(20), nullable=True)
    box_qty = Column(Integer, nullable=False)
    avg_box_weight = Column(Float, default=0.0)
    weight_kg = Column(Float, nullable=False)
    actual_weight_kg = Column(Float, nullable=True)
    reconciled_amount = Column(Float, default=0.0)
    reconciled_status = Column(String(20), default="UNRECONCILED")
    unit_price_kg = Column(Float, nullable=False)
    total_amount = Column(Float, nullable=False)
    exp_date = Column(String(20), nullable=False)
    warehouse = Column(String(50), nullable=False)
    is_weighed = Column(String(10), default="N")
    claim_reason = Column(String(255), nullable=True)
    status = Column(String(20), default="OUT_REQUEST", index=True)
    grid_no = Column(String(50), nullable=True, index=True)
    remark = Column(String(255), nullable=True)
    created_at = Column(DateTime, default=datetime.now)

class DispatchRecord(Base):
    __tablename__ = "dispatches"
    id = Column(Integer, primary_key=True, index=True)
    dispatch_no = Column(String(50), unique=True, index=True)
    dispatch_type = Column(String(30), default="출고배차")
    target_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d"), index=True)
    partner_name = Column(String(100), nullable=False)
    brand = Column(String(50), nullable=True)
    grade = Column(String(30), default="GF")
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    from_warehouse = Column(String(100), nullable=False)
    to_destination = Column(String(200), nullable=False)
    pallet_loaded = Column(String(20), default="N")
    dispatch_cost = Column(Float, default=0.0)
    driver_info = Column(String(100), nullable=True)
    status = Column(String(20), default="REQUESTED")
    remark = Column(String(255), nullable=True)
    created_at = Column(DateTime, default=datetime.now)

class ReservationRecord(Base):
    __tablename__ = "reservations"
    id = Column(Integer, primary_key=True, index=True)
    res_no = Column(String(50), unique=True, index=True)
    lot_id = Column(Integer, ForeignKey("inventory_lots.id"), nullable=False, index=True)
    sales_rep = Column(String(50), nullable=False)
    customer = Column(String(100), nullable=False)
    grid_no = Column(String(50), nullable=False)
    brand = Column(String(50), nullable=True)
    grade = Column(String(30), default="GF")
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    unit_price_kg = Column(Float, nullable=False)
    total_amount = Column(Float, nullable=False)
    exp_date = Column(String(20), nullable=False)
    expire_date = Column(String(20), nullable=False, index=True)
    status = Column(String(20), default="HOLD", index=True)
    created_at = Column(DateTime, default=datetime.now)

class FinancePledge(Base):
    __tablename__ = "finance_pledges"
    id = Column(Integer, primary_key=True, index=True)
    contract_no = Column(String(50), unique=True, index=True)
    lot_id = Column(Integer, ForeignKey("inventory_lots.id"), nullable=False, index=True)
    partner_company = Column(String(100), nullable=False)
    pledge_date = Column(String(20), nullable=False)
    due_date = Column(String(20), nullable=False)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    deposit_amount = Column(Float, nullable=False)
    fee_rate = Column(Float, default=0.0)
    status = Column(String(20), default="PLEDGED", index=True)
    created_at = Column(DateTime, default=datetime.now)

class StockAdjustment(Base):
    __tablename__ = "stock_adjustments"
    id = Column(Integer, primary_key=True, index=True)
    lot_id = Column(Integer, ForeignKey("inventory_lots.id"), nullable=False, index=True)
    adj_type = Column(String(50), nullable=False)
    adj_box = Column(Integer, nullable=False)
    adj_weight = Column(Float, nullable=False)
    reason = Column(String(255), nullable=True)
    adjusted_at = Column(DateTime, default=datetime.now)

class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(50), unique=True, nullable=False, index=True)
    hashed_password = Column(String(255), nullable=False)
    full_name = Column(String(50), nullable=True)
    role = Column(String(20), nullable=False, default=ROLE_SALES)
    is_active = Column(Integer, default=1)
    created_at = Column(DateTime, default=datetime.now)

Base.metadata.create_all(bind=engine)

# -----------------------------------------------------------------------------
# 유틸리티 함수
# -----------------------------------------------------------------------------
def hash_password(plain: str) -> str:
    salt = bcrypt.gensalt()
    return bcrypt.hashpw(plain.encode('utf-8'), salt).decode('utf-8')

def verify_password(plain: str, hashed: str) -> bool:
    try:
        return bcrypt.checkpw(plain.encode('utf-8'), hashed.encode('utf-8'))
    except Exception:
        return False

def create_access_token(data: dict, expires_minutes: int = ACCESS_TOKEN_EXPIRE_MINUTES) -> str:
    to_encode = data.copy()
    expire = datetime.utcnow() + timedelta(minutes=expires_minutes)
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, JWT_SECRET_KEY, algorithm=JWT_ALGORITHM)

def get_current_user(token: Optional[str] = Depends(oauth2_scheme), db: Session = Depends(get_db)) -> "User":
    credentials_exception = HTTPException(status_code=401, detail="인증이 필요합니다.")
    if not token:
        raise credentials_exception
    try:
        payload = jwt.decode(token, JWT_SECRET_KEY, algorithms=[JWT_ALGORITHM])
        username: Optional[str] = payload.get("sub")
        if username is None:
            raise credentials_exception
    except JWTError:
        raise credentials_exception

    user = db.query(User).filter(User.username == username).first()
    if user is None or not user.is_active:
        raise credentials_exception
    return user

def require_roles(*roles: str):
    def _checker(current_user: User = Depends(get_current_user)) -> User:
        if current_user.role not in roles:
            raise HTTPException(status_code=403, detail="권한이 없습니다.")
        return current_user
    return _checker

def generate_random_grid() -> str:
    return f"GRID-{random.randint(100000, 999999)}"

def find_lot_by_grid(db: Session, grid_no: Optional[str]) -> Optional[InventoryLot]:
    if not grid_no:
        return None
    return db.query(InventoryLot).filter(InventoryLot.grid_no == grid_no).first()

# -----------------------------------------------------------------------------
# 시드 데이터 초기화
# -----------------------------------------------------------------------------
def init_sample_data():
    db = SessionLocal()
    try:
        if db.query(CompanyMaster).count() == 0:
            comps = [
                CompanyMaster(code="TJF", name="주식회사 티제이에프", biz_no="718-88-03523", rep_name="김다희", address="경기도 남양주시 별내중앙로 34, 5층 502-A59호", phone="010-5773-0619", fax="070-4758-9219"),
                CompanyMaster(code="SEOUL", name="(주)서울웰푸드", biz_no="347-81-03002", rep_name="조용훈", address="서울특별시 강동구 천중로 39길 19-25(천호동,평영빌딩)", phone="02-6958-9229", fax="070-4758-9219"),
                CompanyMaster(code="NEXUS", name="(주)넥서스트레이딩", biz_no="518-86-03633", rep_name="김다희", address="서울특별시 강동구 천중로 39길 19-25(천호동,평영빌딩),2층", phone="010-5773-0619", fax="070-4758-9219"),
            ]
            db.add_all(comps)

        if db.query(User).count() == 0:
            users = [
                User(username="admin", hashed_password=hash_password("admin1234!"), full_name="시스템관리자", role=ROLE_ADMIN),
                User(username="sales1", hashed_password=hash_password("sales1234!"), full_name="김영업", role=ROLE_SALES),
                User(username="wh1", hashed_password=hash_password("wh1234!"), full_name="최창고", role=ROLE_WAREHOUSE),
            ]
            db.add_all(users)

        if db.query(InventoryLot).count() == 0:
            lots = [
                InventoryLot(
                    company_name="주식회사 티제이에프", grid_no="GRID-771101", sku_code="BF-AMH-01", inbound_date="2026-08-10",
                    bl_no="OOLU2766442080", trace_no="802410290114", est_no="244I", brand="AMH", grade="GF",
                    item_name="우육(소)", cut_name="소일반갈비(170)", storage_type="냉동", process_from_date="2026-06-01",
                    initial_box_qty=196, initial_weight_kg=4000.0, avg_box_weight=20.41, current_box_qty=196, current_weight_kg=4000.0,
                    cost_per_kg=14500, warehouse="강동2", exp_date="2028-05-31", is_weighed="N"
                ),
                InventoryLot(
                    company_name="주식회사 티제이에프", grid_no="GRID-771102", sku_code="BF-AMH-02", inbound_date="2026-08-10",
                    bl_no="OOLU2766442080", trace_no="802410290115", est_no="244I", brand="AMH", grade="GF",
                    item_name="우육(소)", cut_name="소일반갈비(235)", storage_type="냉동", process_from_date="2026-06-01",
                    initial_box_qty=345, initial_weight_kg=7000.0, avg_box_weight=20.29, current_box_qty=345, current_weight_kg=7000.0,
                    cost_per_kg=14200, warehouse="강동2", exp_date="2028-05-31", is_weighed="N"
                )
            ]
            db.add_all(lots)
            db.flush()

        if db.query(FinancePledge).count() == 0:
            pledge = FinancePledge(
                contract_no="FIN-2026082001", lot_id=1, partner_company="(주)대한육류유통",
                pledge_date="2026-08-20", due_date="2026-09-20", box_qty=30, weight_kg=612.3,
                deposit_amount=5000000, fee_rate=2.5, status="PLEDGED"
            )
            db.add(pledge)

        db.commit()
    finally:
        db.close()

init_sample_data()

# -----------------------------------------------------------------------------
# 3. FastAPI 라우터 및 스키마
# -----------------------------------------------------------------------------
app = FastAPI(title="MeatFlow Enterprise ERP")

# Render 및 외부 접속용 CORS 허용
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class CompanyOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    code: str
    name: str
    biz_no: str
    rep_name: str
    address: str
    phone: Optional[str] = None
    fax: Optional[str] = None

class InboundCreate(BaseModel):
    company_name: str = "주식회사 티제이에프"
    inbound_date: Optional[str] = None
    vendor: str
    bl_no: Optional[str] = ""
    trace_no: str
    est_no: Optional[str] = ""
    brand: Optional[str] = ""
    grade: Optional[str] = "GF"
    item_name: str
    cut_name: str
    storage_type: str
    box_qty: int
    weight_kg: float
    cost_per_kg: float
    warehouse: str
    exp_date: str

class OutboundCreate(BaseModel):
    company_name: str = "주식회사 티제이에프"
    vendor_chain: Optional[str] = "브루니"
    trade_type: str = "출고판매"
    is_estimated: Optional[str] = "N"
    outbound_date: Optional[str] = None
    grid_no: str
    customer: str
    box_qty: int
    weight_kg: Optional[float] = None
    unit_price_kg: float
    remark: Optional[str] = ""

class DispatchCreate(BaseModel):
    dispatch_type: str = "출고배차"
    target_date: Optional[str] = None
    partner_name: str
    brand: Optional[str] = ""
    grade: Optional[str] = "GF"
    item_name: str
    cut_name: str
    box_qty: int
    weight_kg: float
    from_warehouse: str
    to_destination: str
    pallet_loaded: str = "N"
    dispatch_cost: float = 0.0
    driver_info: Optional[str] = ""
    remark: Optional[str] = ""

class WarehouseTransferReq(BaseModel):
    grid_no: str
    to_warehouse: str
    box_qty: int
    create_dispatch: bool = False
    dispatch_cost: float = 0.0
    pallet_loaded: str = "N"

class ReconcileWeightReq(BaseModel):
    outbound_no: str
    actual_weight_kg: float

class FinanceCreate(BaseModel):
    grid_no: str
    partner_company: str
    pledge_date: str
    due_date: str
    box_qty: int
    deposit_amount: float
    fee_rate: float

class WeightReconcileDirectReq(BaseModel):
    grid_no: str
    actual_weight_kg: float
    reason: Optional[str] = "실계근 중량 오차 보정"

class MessageOut(BaseModel):
    message: str

# API 엔드포인트
@app.post("/api/auth/login")
def login(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user = db.query(User).filter(User.username == form_data.username).first()
    if not user or not user.is_active or not verify_password(form_data.password, user.hashed_password):
        raise HTTPException(status_code=401, detail="아이디 또는 비밀번호 오류")
    token = create_access_token({"sub": user.username, "role": user.role})
    return {"access_token": token, "token_type": "bearer", "role": user.role, "username": user.username, "full_name": user.full_name}

@app.get("/api/companies", response_model=List[CompanyOut])
def get_companies(db: Session = Depends(get_db)):
    return db.query(CompanyMaster).all()

@app.get("/api/inbounds")
def get_inbounds(status: str, db: Session = Depends(get_db)):
    return db.query(InboundRecord).filter(InboundRecord.status == status).order_by(desc(InboundRecord.id)).all()

@app.post("/api/inbounds", response_model=MessageOut)
def create_inbound(req: InboundCreate, current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)), db: Session = Depends(get_db)):
    grid_no = generate_random_grid()
    inbound_no = f"IN-{datetime.now().strftime('%Y%m%d%H%M%S')}-{random.randint(10,99)}"
    in_date = req.inbound_date or datetime.now().strftime("%Y-%m-%d")
    total = round(req.weight_kg * req.cost_per_kg)

    rec = InboundRecord(
        inbound_no=inbound_no, company_name=req.company_name, grid_no=grid_no,
        inbound_date=in_date, vendor=req.vendor, bl_no=req.bl_no, trace_no=req.trace_no,
        est_no=req.est_no, brand=req.brand, grade=req.grade, item_name=req.item_name,
        cut_name=req.cut_name, storage_type=req.storage_type, box_qty=req.box_qty,
        weight_kg=req.weight_kg, cost_per_kg=req.cost_per_kg, total_amount=total,
        warehouse=req.warehouse, exp_date=req.exp_date, status="IN_REQUEST"
    )
    db.add(rec)
    db.commit()
    return {"message": f"입고요청 등록 완료 (GRID: {grid_no})"}

@app.post("/api/inbounds/{grid_no}/advance", response_model=MessageOut)
def advance_inbound(grid_no: str, current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)), db: Session = Depends(get_db)):
    inbound = db.query(InboundRecord).filter(InboundRecord.grid_no == grid_no).first()
    if not inbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    if inbound.status == "IN_REQUEST":
        inbound.status = "IN_CONFIRM"
        msg = "입고확정 단계로 이동되었습니다."
    elif inbound.status == "IN_CONFIRM":
        inbound.status = "IN_DONE"
        inbound.processed_date = datetime.now().strftime("%Y-%m-%d")
        avg_w = round(inbound.weight_kg / inbound.box_qty, 2) if inbound.box_qty > 0 else 0.0

        lot = find_lot_by_grid(db, grid_no)
        if lot:
            lot.current_box_qty += inbound.box_qty
            lot.current_weight_kg = round(lot.current_weight_kg + inbound.weight_kg, 2)
        else:
            sku = "SKU-" + (inbound.trace_no[:6] if len(inbound.trace_no) >= 6 else inbound.trace_no)
            new_lot = InventoryLot(
                company_name=inbound.company_name, grid_no=grid_no, sku_code=sku,
                inbound_date=inbound.inbound_date, bl_no=inbound.bl_no, trace_no=inbound.trace_no,
                est_no=inbound.est_no, brand=inbound.brand, grade=inbound.grade, item_name=inbound.item_name,
                cut_name=inbound.cut_name, storage_type=inbound.storage_type, initial_box_qty=inbound.box_qty,
                initial_weight_kg=inbound.weight_kg, avg_box_weight=avg_w, current_box_qty=inbound.box_qty,
                current_weight_kg=inbound.weight_kg, cost_per_kg=inbound.cost_per_kg, warehouse=inbound.warehouse,
                exp_date=inbound.exp_date, is_weighed=inbound.is_weighed
            )
            db.add(new_lot)
        msg = "입고완료 및 실재고 반영이 완료되었습니다."
    else:
        msg = "이미 완료된 전표입니다."
    db.commit()
    return {"message": msg}

@app.get("/api/inventory")
def get_inventory(db: Session = Depends(get_db)):
    lots = db.query(InventoryLot).filter(InventoryLot.current_box_qty > 0).order_by(InventoryLot.exp_date).all()
    
    res_rows = db.query(ReservationRecord.grid_no, func.sum(ReservationRecord.box_qty).label("box")).filter(ReservationRecord.status == "HOLD").group_by(ReservationRecord.grid_no).all()
    res_map = {r.grid_no: r.box or 0 for r in res_rows}

    pledges = db.query(FinancePledge).filter(FinancePledge.status == "PLEDGED").all()
    pledge_map = {}
    for p in pledges:
        pledge_map.setdefault(p.lot_id, []).append(f"{p.partner_company}({p.box_qty}Box)")

    result = []
    for l in lots:
        p_list = pledge_map.get(l.id, [])
        result.append({
            "grid_no": l.grid_no,
            "company_name": l.company_name,
            "bl_no": l.bl_no,
            "trace_no": l.trace_no,
            "est_no": l.est_no,
            "brand": l.brand,
            "grade": l.grade,
            "item_name": l.item_name,
            "cut_name": l.cut_name,
            "storage_type": l.storage_type,
            "current_box_qty": l.current_box_qty,
            "current_weight_kg": l.current_weight_kg,
            "cost_per_kg": l.cost_per_kg,
            "warehouse": l.warehouse,
            "exp_date": l.exp_date,
            "is_weighed": l.is_weighed,
            "reserved_box_qty": res_map.get(l.grid_no, 0),
            "pledged_desc": ", ".join(p_list) if p_list else "-",
        })
    return result

@app.post("/api/inventory/reconcile-weight-direct", response_model=MessageOut)
def reconcile_weight_direct(req: WeightReconcileDirectReq, current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)), db: Session = Depends(get_db)):
    lot = find_lot_by_grid(db, req.grid_no)
    if not lot:
        raise HTTPException(status_code=404, detail="재고를 찾을 수 없습니다.")

    diff_w = round(lot.current_weight_kg - req.actual_weight_kg, 2)
    lot.current_weight_kg = req.actual_weight_kg
    if lot.current_box_qty > 0:
        lot.avg_box_weight = round(lot.current_weight_kg / lot.current_box_qty, 2)
    lot.is_weighed = "Y"

    adj = StockAdjustment(
        lot_id=lot.id, adj_type="실계근보정", adj_box=0, adj_weight=abs(diff_w),
        reason=f"{req.reason} (변동량: {-diff_w:+.2f}kg)"
    )
    db.add(adj)
    db.commit()
    return {"message": f"실계근 중량({req.actual_weight_kg}kg, 평중 {lot.avg_box_weight}kg)으로 보정되었습니다."}

@app.get("/api/outbounds")
def get_outbounds(status: str, db: Session = Depends(get_db)):
    return db.query(OutboundRecord).filter(OutboundRecord.status == status).order_by(desc(OutboundRecord.id)).all()

@app.post("/api/outbounds/create", response_model=MessageOut)
def create_outbound(req: OutboundCreate, current_user: User = Depends(require_roles(ROLE_SALES, ROLE_ADMIN)), db: Session = Depends(get_db)):
    lot = find_lot_by_grid(db, req.grid_no)
    if not lot:
        raise HTTPException(status_code=404, detail="재고를 찾을 수 없습니다.")

    avg_w = lot.avg_box_weight if lot.avg_box_weight > 0 else 20.0
    calc_weight = req.weight_kg if (req.weight_kg and req.weight_kg > 0) else round(avg_w * req.box_qty, 2)

    lot.current_box_qty = max(0, lot.current_box_qty - req.box_qty)
    lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - calc_weight, 2))

    out_no = f"OUT-{datetime.now().strftime('%Y%m%d%H%M%S')}-{random.randint(10,99)}"
    out_date = req.outbound_date or datetime.now().strftime("%Y-%m-%d")
    total = round(calc_weight * req.unit_price_kg)

    outbound = OutboundRecord(
        outbound_no=out_no, company_name=req.company_name, vendor_chain=req.vendor_chain,
        trade_type=req.trade_type, is_estimated=req.is_estimated or "N", grid_no=lot.grid_no, lot_id=lot.id,
        inbound_date=lot.inbound_date, outbound_date=out_date, customer=req.customer, bl_no=lot.bl_no,
        trace_no=lot.trace_no, est_no=lot.est_no, brand=lot.brand, grade=lot.grade, item_name=lot.item_name,
        cut_name=lot.cut_name, storage_type=lot.storage_type, process_from_date=lot.process_from_date,
        box_qty=req.box_qty, avg_box_weight=avg_w, weight_kg=calc_weight, unit_price_kg=req.unit_price_kg,
        total_amount=total, exp_date=lot.exp_date, warehouse=lot.warehouse, status="OUT_REQUEST", remark=req.remark
    )
    db.add(outbound)
    db.commit()
    return {"message": f"[{req.trade_type}] 출고요청 등록 완료 ({out_no})"}

@app.post("/api/outbounds/{outbound_no}/advance", response_model=MessageOut)
def advance_outbound(outbound_no: str, current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)), db: Session = Depends(get_db)):
    out = db.query(OutboundRecord).filter(OutboundRecord.outbound_no == outbound_no).first()
    if not out:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    if out.status == "OUT_REQUEST":
        out.status = "OUT_CONFIRM"
        msg = f"[{outbound_no}] 출고확정 단계로 이동되었습니다. (확정 후 수정불가, 취소만 가능)"
    elif out.status == "OUT_CONFIRM":
        out.status = "OUT_DONE"
        out.processed_date = datetime.now().strftime("%Y-%m-%d")
        msg = f"[{outbound_no}] 출고완료 처리가 완료되었습니다."
    else:
        msg = "이미 완료된 전표입니다."
    db.commit()
    return {"message": msg}

@app.post("/api/outbounds/{outbound_no}/revert", response_model=MessageOut)
def revert_outbound(outbound_no: str, current_user: User = Depends(require_roles(ROLE_SALES, ROLE_WAREHOUSE, ROLE_ADMIN)), db: Session = Depends(get_db)):
    out = db.query(OutboundRecord).filter(OutboundRecord.outbound_no == outbound_no).first()
    if not out:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    if out.status == "OUT_REQUEST":
        lot = db.query(InventoryLot).filter(InventoryLot.id == out.lot_id).first()
        if lot:
            lot.current_box_qty += out.box_qty
            lot.current_weight_kg = round(lot.current_weight_kg + out.weight_kg, 2)
        db.delete(out)
        msg = "출고요청이 취소되어 재고가 원복되었습니다."
    elif out.status == "OUT_CONFIRM":
        out.status = "OUT_REQUEST"
        msg = "출고확정이 취소되어 [출고요청리스트]로 반려되었습니다."
    elif out.status == "OUT_DONE":
        out.status = "OUT_CONFIRM"
        msg = "출고완료가 취소되었습니다."
    else:
        msg = "처리할 수 없는 상태입니다."
    db.commit()
    return {"message": msg}

@app.get("/api/settlements")
def get_settlements(db: Session = Depends(get_db)):
    return db.query(OutboundRecord).filter(
        (OutboundRecord.is_estimated == "Y") | (OutboundRecord.actual_weight_kg != None)
    ).order_by(desc(OutboundRecord.id)).all()

@app.post("/api/settlements/reconcile", response_model=MessageOut)
def reconcile_settlement(req: ReconcileWeightReq, current_user: User = Depends(require_roles(ROLE_SALES, ROLE_ADMIN)), db: Session = Depends(get_db)):
    out = db.query(OutboundRecord).filter(OutboundRecord.outbound_no == req.outbound_no).first()
    if not out:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    diff_w = round(req.actual_weight_kg - out.weight_kg, 2)
    diff_amount = round(diff_w * out.unit_price_kg)

    out.actual_weight_kg = req.actual_weight_kg
    out.reconciled_amount = diff_amount
    out.reconciled_status = "RECONCILED"
    db.commit()
    return {"message": f"실계근 {req.actual_weight_kg}kg 기준 정산 완료 (정산차액: {diff_amount:+,}원)"}

@app.get("/api/dispatches")
def get_dispatches(db: Session = Depends(get_db)):
    return db.query(DispatchRecord).order_by(desc(DispatchRecord.id)).all()

@app.post("/api/dispatches", response_model=MessageOut)
def create_dispatch(req: DispatchCreate, current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)), db: Session = Depends(get_db)):
    disp_no = f"DSP-{datetime.now().strftime('%Y%m%d%H%M%S')}-{random.randint(10,99)}"
    t_date = req.target_date or datetime.now().strftime("%Y-%m-%d")

    disp = DispatchRecord(
        dispatch_no=disp_no, dispatch_type=req.dispatch_type, target_date=t_date,
        partner_name=req.partner_name, brand=req.brand, grade=req.grade,
        item_name=req.item_name, cut_name=req.cut_name, box_qty=req.box_qty,
        weight_kg=req.weight_kg, from_warehouse=req.from_warehouse,
        to_destination=req.to_destination, pallet_loaded=req.pallet_loaded,
        dispatch_cost=req.dispatch_cost, driver_info=req.driver_info,
        remark=req.remark, status="REQUESTED"
    )
    db.add(disp)
    db.commit()
    return {"message": f"배차요청 등록 완료 (전표: {disp_no})"}

@app.post("/api/inventory/transfer-warehouse", response_model=MessageOut)
def transfer_warehouse(req: WarehouseTransferReq, current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)), db: Session = Depends(get_db)):
    lot = find_lot_by_grid(db, req.grid_no)
    if not lot:
        raise HTTPException(status_code=404, detail="재고를 찾을 수 없습니다.")
    if req.to_warehouse == lot.warehouse:
        raise HTTPException(status_code=400, detail="동일 창고 이동 불가")

    avg_w = lot.avg_box_weight if lot.avg_box_weight > 0 else 20.0
    move_weight = round(avg_w * req.box_qty, 2)

    lot.current_box_qty = max(0, lot.current_box_qty - req.box_qty)
    lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - move_weight, 2))

    new_grid = generate_random_grid()
    new_lot = InventoryLot(
        company_name=lot.company_name, grid_no=new_grid, sku_code=lot.sku_code,
        inbound_date=datetime.now().strftime("%Y-%m-%d"), bl_no=lot.bl_no, trace_no=lot.trace_no,
        est_no=lot.est_no, brand=lot.brand, grade=lot.grade, item_name=lot.item_name, cut_name=lot.cut_name,
        storage_type=lot.storage_type, process_from_date=lot.process_from_date, initial_box_qty=req.box_qty,
        initial_weight_kg=move_weight, avg_box_weight=avg_w, current_box_qty=req.box_qty,
        current_weight_kg=move_weight, cost_per_kg=lot.cost_per_kg, warehouse=req.to_warehouse,
        exp_date=lot.exp_date, is_weighed=lot.is_weighed
    )
    db.add(new_lot)

    if req.create_dispatch:
        disp_no = f"DSP-TR-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        disp = DispatchRecord(
            dispatch_no=disp_no, dispatch_type="창고이동(전배)", target_date=datetime.now().strftime("%Y-%m-%d"),
            partner_name=f"자사전배({lot.company_name})", brand=lot.brand, grade=lot.grade,
            item_name=lot.item_name, cut_name=lot.cut_name, box_qty=req.box_qty, weight_kg=move_weight,
            from_warehouse=lot.warehouse, to_destination=req.to_warehouse, pallet_loaded=req.pallet_loaded,
            dispatch_cost=req.dispatch_cost, remark=f"보관료 최적화 전배 (신규 GRID: {new_grid})"
        )
        db.add(disp)

    db.commit()
    return {"message": f"{lot.warehouse} -> {req.to_warehouse}로 이동 완료 (신규 GRID: {new_grid})"}

@app.get("/api/finances")
def get_finances(db: Session = Depends(get_db)):
    return db.query(FinancePledge).order_by(desc(FinancePledge.id)).all()

@app.post("/api/finances", response_model=MessageOut)
def create_finance_pledge(req: FinanceCreate, current_user: User = Depends(require_roles(ROLE_SALES, ROLE_ADMIN)), db: Session = Depends(get_db)):
    lot = find_lot_by_grid(db, req.grid_no)
    if not lot:
        raise HTTPException(status_code=404, detail="재고를 찾을 수 없습니다.")

    avg_w = lot.avg_box_weight if lot.avg_box_weight > 0 else 20.0
    calc_w = round(avg_w * req.box_qty, 2)
    c_no = f"FIN-{datetime.now().strftime('%Y%m%d%H%M%S')}"

    pledge = FinancePledge(
        contract_no=c_no, lot_id=lot.id, partner_company=req.partner_company,
        pledge_date=req.pledge_date, due_date=req.due_date, box_qty=req.box_qty,
        weight_kg=calc_w, deposit_amount=req.deposit_amount, fee_rate=req.fee_rate,
        status="PLEDGED"
    )
    db.add(pledge)
    db.commit()
    return {"message": f"[{req.partner_company}] 대상 파이낸스 담보 등록 완료 ({c_no})"}

# -----------------------------------------------------------------------------
# 4. 프론트엔드 UI
# -----------------------------------------------------------------------------
HTML_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>MeatFlow Enterprise ERP</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    :root { --primary: #2563eb; --sidebar: #0f172a; --bg: #f8fafc; --border: #e2e8f0; --text: #1e293b; --muted: #64748b; }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Pretendard", "Segoe UI", sans-serif; }
    body { background: var(--bg); color: var(--text); display: flex; height: 100vh; overflow: hidden; }
    aside { width: 250px; background: var(--sidebar); color: #94a3b8; display: flex; flex-direction: column; flex-shrink: 0; }
    .brand { padding: 20px; font-size: 1.15rem; font-weight: 700; color: #fff; border-bottom: 1px solid #1e293b; display: flex; align-items: center; gap: 8px; }
    .brand i { color: #38bdf8; }
    .user-profile { padding: 12px 18px; background: #1e293b; color: #e2e8f0; font-size: 0.82rem; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #334155; }
    .nav-category { font-size: 0.72rem; text-transform: uppercase; font-weight: 700; color: #64748b; padding: 14px 18px 6px; }
    .nav-item { padding: 10px 18px; color: #94a3b8; cursor: pointer; display: flex; align-items: center; gap: 10px; font-size: 0.88rem; transition: 0.2s; }
    .nav-item:hover, .nav-item.active { background: #1e293b; color: #38bdf8; font-weight: 600; }
    main { flex-grow: 1; display: flex; flex-direction: column; height: 100vh; overflow-y: auto; }
    header { background: #fff; border-bottom: 1px solid var(--border); padding: 14px 28px; display: flex; justify-content: space-between; align-items: center; }
    .content { padding: 20px 28px; display: flex; flex-direction: column; gap: 14px; }
    .sub-tabs { display: flex; gap: 6px; background: #e2e8f0; padding: 4px; border-radius: 8px; width: fit-content; }
    .sub-tab { padding: 8px 16px; border-radius: 6px; font-size: 0.85rem; font-weight: 600; cursor: pointer; color: #475569; }
    .sub-tab.active { background: #fff; color: var(--primary); box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .filter-card { background: #fff; border: 1px solid var(--border); border-radius: 8px; padding: 12px 18px; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px; }
    .filter-group { display: flex; align-items: center; gap: 8px; font-size: 0.85rem; font-weight: 600; }
    .filter-input { padding: 6px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.85rem; outline: none; background: #fff; }
    .table-card { background: #fff; border: 1px solid var(--border); border-radius: 8px; overflow-x: auto; box-shadow: 0 1px 2px rgba(0,0,0,0.03); }
    table { width: 100%; border-collapse: collapse; font-size: 0.84rem; text-align: left; }
    th { background: #f8fafc; color: var(--muted); padding: 11px 12px; border-bottom: 1px solid var(--border); font-weight: 600; white-space: nowrap; }
    td { padding: 11px 12px; border-bottom: 1px solid var(--border); vertical-align: middle; white-space: nowrap; }
    tr:hover td { background: #f8fafc; }
    .btn { padding: 6px 12px; border-radius: 6px; font-size: 0.82rem; font-weight: 600; border: 1px solid transparent; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-outline { background: #fff; border-color: var(--border); color: #334155; }
    .badge { padding: 3px 6px; border-radius: 4px; font-size: 0.74rem; font-weight: 600; }
    .grid-tag { background: #312e81; color: #e0e7ff; padding: 2px 6px; border-radius: 4px; font-weight: 700; font-family: monospace; }
    .modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 100; align-items: center; justify-content: center; }
    .modal-box { background: #fff; border-radius: 10px; width: 620px; padding: 24px; max-height: 90vh; overflow-y: auto; }
    .form-group { margin-bottom: 12px; display: flex; flex-direction: column; gap: 4px; font-size: 0.82rem; font-weight: 600; }
    .form-control { padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.88rem; }

    @media print {
      body * { visibility: hidden; }
      #printWrapper, #printWrapper * { visibility: visible; }
      #printWrapper { position: absolute; left: 0; top: 0; width: 100%; }
    }
    .sheet-box { width: 780px; margin: 0 auto; background: #fff; padding: 30px; border: 1px solid #94a3b8; font-family: 'Malgun Gothic', sans-serif; color: #000; }
    .sheet-title { font-size: 26px; font-weight: 900; letter-spacing: 4px; text-align: center; margin: 15px 0; }
    .meta-table, .grid-table, .footer-table { width: 100%; border-collapse: collapse; }
    .meta-table td { padding: 6px 8px; font-size: 13px; font-weight: bold; }
    .grid-table th, .grid-table td { border: 1px solid #000; padding: 6px 4px; font-size: 12px; text-align: center; }
    .grid-table th { background: #f8fafc; font-weight: bold; }
    .footer-table td { border: 1px solid #000; padding: 6px 8px; font-size: 11px; vertical-align: middle; }
    .stamp-box { position: relative; width: 100px; height: 80px; display: flex; align-items: center; justify-content: center; margin: 0 auto; }
    .stamp-img { width: 70px; height: 70px; border: 2px solid #dc2626; border-radius: 50%; color: #dc2626; font-size: 11px; font-weight: bold; display: flex; align-items: center; justify-content: center; text-align: center; }
  </style>
</head>
<body>
  <div class="modal-overlay" id="loginModal" style="display:flex;">
    <div class="modal-box" style="width:380px;">
      <h3 style="margin-bottom:12px; text-align:center;"><i class="bi bi-shield-lock-fill" style="color:var(--primary);"></i> MeatFlow ERP</h3>
      <div class="form-group"><label>아이디</label><input type="text" id="login_user" class="form-control" value="admin" /></div>
      <div class="form-group"><label>비밀번호</label><input type="password" id="login_pw" class="form-control" value="admin1234!" /></div>
      <button class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; margin-top:8px;" onclick="handleLogin()">로그인</button>
    </div>
  </div>

  <div class="modal-overlay" id="actionModal">
    <div class="modal-box">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h3 id="modalTitle" style="font-size:1.1rem; font-weight:700;">작업</h3>
        <button class="btn btn-outline" onclick="closeModal()" style="padding:2px 8px;">✕</button>
      </div>
      <div id="modalFormBody"></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal()">취소</button>
        <button class="btn btn-primary" onclick="submitModal()">저장</button>
      </div>
    </div>
  </div>

  <aside>
    <div class="brand"><i class="bi bi-box-seam-fill"></i> MeatFlow ERP</div>
    <div class="user-profile">
      <span id="userBadge"><i class="bi bi-person-badge"></i> 미인증</span>
      <button class="btn btn-outline" style="padding:2px 6px; font-size:0.72rem;" onclick="handleLogout()">로그아웃</button>
    </div>
    <div class="nav-category">프로세스 관리</div>
    <div class="nav-item active" onclick="switchMainTab('INBOUND', this)"><i class="bi bi-box-arrow-in-down"></i> 1. 입고 관리</div>
    <div class="nav-item" onclick="switchMainTab('STOCK', this)"><i class="bi bi-grid-3x3-gap-fill"></i> 2. 현재고장 (당일조회)</div>
    <div class="nav-item" onclick="switchMainTab('OUTBOUND', this)"><i class="bi bi-box-arrow-up"></i> 3. 출고 & 이체요청 (PDF)</div>
    <div class="nav-item" onclick="switchMainTab('DISPATCH', this)"><i class="bi bi-truck" style="color:#38bdf8;"></i> 4. 배차 관리 (입·출고/전배)</div>
    <div class="nav-item" onclick="switchMainTab('SETTLEMENT', this)"><i class="bi bi-calculator-fill" style="color:#10b981;"></i> 5. 가중량 사후정산</div>
    <div class="nav-item" onclick="switchMainTab('FINANCE', this)"><i class="bi bi-bank2" style="color:#f59e0b;"></i> 6. 기업 파이낸스 (담보/재매입)</div>
  </aside>

  <main>
    <header>
      <h2 id="pageTitle" style="font-size:1.25rem;">입고 프로세스 관리</h2>
      <div id="headerActions" style="display:flex; gap:8px;"></div>
    </header>

    <div class="content">
      <div class="sub-tabs" id="subTabs" style="display:flex;"></div>
      
      <div class="filter-card">
        <div class="filter-group">
          <span><i class="bi bi-calendar-range"></i> 조회기간:</span>
          <input type="date" id="filter_start" class="filter-input" />
          <span>~</span>
          <input type="date" id="filter_end" class="filter-input" />
        </div>
        <div class="filter-group" style="flex-grow:1; max-width:460px;">
          <select id="filter_field" class="filter-input" style="width:130px; font-weight:600;">
            <option value="ALL">전체 항목</option>
            <option value="trace_no">이력번호</option>
            <option value="bl_no">BL NO</option>
            <option value="brand">브랜드</option>
            <option value="grade">등급</option>
            <option value="item_name">품목/부위</option>
            <option value="partner">거래처명</option>
            <option value="warehouse">창고</option>
          </select>
          <input type="text" id="filter_keyword" class="filter-input" style="flex-grow:1;" placeholder="검색어 입력..." onkeypress="if(event.keyCode==13){loadData();}" />
        </div>
        <div style="display:flex; gap:6px;">
          <button class="btn btn-primary" onclick="loadData()"><i class="bi bi-search"></i> 검색</button>
          <button class="btn btn-outline" onclick="resetFilters()"><i class="bi bi-arrow-counterclockwise"></i> 당일 초기화</button>
        </div>
      </div>

      <div class="table-card">
        <table>
          <thead id="tableHead"></thead>
          <tbody id="tableBody"></tbody>
        </table>
      </div>
    </div>
  </main>

  <div id="printWrapper" style="display:none;">
    <div class="sheet-box">
      <div style="display:flex; justify-content:space-between; align-items:flex-start;">
        <div id="docCompanyBadge" style="background:#dcfce7; color:#15803d; padding:4px 10px; font-weight:bold; font-size:12px; border:1px solid #86efac;">티제이에프</div>
        <table style="border:1px solid #000; border-collapse:collapse; text-align:center; font-size:12px;">
          <tr>
            <td style="border:1px solid #000; padding:4px 14px; background:#f8fafc; font-weight:bold;">매입처</td>
            <td style="border:1px solid #000; padding:4px 20px; font-weight:bold;" id="docVendorChain">브루니</td>
          </tr>
        </table>
      </div>

      <div class="sheet-title" id="docMainTitle">이 체 요 청 서</div>

      <table class="meta-table" style="margin-bottom:6px;">
        <tr>
          <td style="width:50%;">수신 : <span id="docWarehouse" style="font-weight:900; font-size:15px;">강동2</span></td>
          <td style="text-align:right;">요청일 : <span id="docDate">2026-08-24</span></td>
        </tr>
      </table>

      <table class="grid-table" style="margin-bottom:14px;">
        <thead>
          <tr>
            <th style="width:36px;">No</th>
            <th>B/L NO</th>
            <th style="width:70px;">브랜드</th>
            <th style="width:50px;">등급</th>
            <th>품목명</th>
            <th style="width:60px;">BOX</th>
            <th style="width:110px;" id="docTargetCol">이체처</th>
            <th style="width:80px;">비고</th>
          </tr>
        </thead>
        <tbody id="docGridBody"></tbody>
        <tfoot>
          <tr style="font-weight:bold; background:#f8fafc;">
            <td colspan="5">합 계</td>
            <td id="docSumBox">0</td>
            <td colspan="2"></td>
          </tr>
        </tfoot>
      </table>

      <div style="text-align:center; font-size:13px; font-weight:bold; margin-bottom:10px;" id="docBottomMsg">
        상기와 같이 요청하오니 이체하여 주시기 바랍니다.
      </div>

      <table class="footer-table">
        <tr>
          <td style="width:12%; font-weight:bold; text-align:center;">사 업</td>
          <td style="width:38%; font-weight:bold;" id="f_biz_no">718-88-03523</td>
          <td style="width:25%; font-weight:bold; text-align:center;">거래인감</td>
          <td style="width:25%; font-weight:bold; text-align:center;">연락처</td>
        </tr>
        <tr>
          <td style="font-weight:bold; text-align:center;">상</td>
          <td id="f_comp_name">주식회사 티제이에프</td>
          <td rowspan="3" style="text-align:center;">
            <div class="stamp-box">
              <div class="stamp-img"><span id="stampText">주식회사<br>티제이에프<br>인</span></div>
            </div>
          </td>
          <td style="font-size:11px;">
            TEL : <span id="f_tel">010-5773-0619</span><br>
            FAX : <span id="f_fax">070-4758-9219</span>
          </td>
        </tr>
        <tr>
          <td style="font-weight:bold; text-align:center;">주</td>
          <td id="f_address" style="font-size:10.5px;">경기도 남양주시 별내중앙로 34, 5층 502-A59호</td>
          <td rowspan="2" style="font-size:10px; color:#64748b;">* 출고/이체시 날인은 당사 거래인감으로만 가능함.</td>
        </tr>
        <tr>
          <td style="font-weight:bold; text-align:center;">대 표</td>
          <td style="font-weight:bold;" id="f_rep_name">김 다 희</td>
        </tr>
      </table>

      <div style="margin-top:10px; font-size:11px; font-weight:bold;">
        * 이체의 경우 창고보관료는 해당일자 까지만 당사 부담이며 이후는 이체처 부담으로 함.
      </div>
      <div style="text-align:center; font-size:16px; font-weight:900; margin-top:14px;" id="f_huge_title">
        주식회사 티제이에프
      </div>
    </div>
  </div>

  <script>
    let authToken = localStorage.getItem('meat_token') || '';
    let currentMain = 'INBOUND', currentSub = 'IN_REQUEST';
    let cachedList = [];
    let companiesCache = [];

    const getTodayStr = () => new Date().toISOString().slice(0, 10);

    async function apiFetch(url, options = {}) {
      options.headers = options.headers || {};
      if (authToken) options.headers['Authorization'] = 'Bearer ' + authToken;
      const res = await fetch(url, options);
      if (res.status === 401) {
        document.getElementById('loginModal').style.display = 'flex';
        throw new Error('인증 만료');
      }
      return res;
    }

    async function handleLogin() {
      const u = document.getElementById('login_user').value;
      const p = document.getElementById('login_pw').value;
      const form = new URLSearchParams({ username: u, password: p });
      const r = await fetch('/api/auth/login', { method: 'POST', body: form });
      const d = await r.json();
      if (r.ok) {
        authToken = d.access_token;
        localStorage.setItem('meat_token', authToken);
        document.getElementById('userBadge').innerText = `${d.full_name || d.username} (${d.role})`;
        document.getElementById('loginModal').style.display = 'none';
        await fetchCompanies();
        resetFilters();
      } else {
        alert(d.detail || '로그인 실패');
      }
    }

    function handleLogout() {
      localStorage.removeItem('meat_token');
      location.reload();
    }

    async function fetchCompanies() {
      try {
        const r = await apiFetch('/api/companies');
        companiesCache = await r.json();
      } catch(e) {}
    }

    function resetFilters() {
      const today = getTodayStr();
      document.getElementById('filter_start').value = today;
      document.getElementById('filter_end').value = today;
      document.getElementById('filter_keyword').value = '';
      document.getElementById('filter_field').value = 'ALL';
      loadData();
    }

    function switchMainTab(tab, el) {
      currentMain = tab;
      document.querySelectorAll('.nav-item').forEach(e => e.classList.remove('active'));
      if (el) el.classList.add('active');

      const titleMap = {
        'INBOUND': '1. 입고 프로세스 관리',
        'STOCK': '2. 현재고장 (당일조회/담보통합)',
        'OUTBOUND': '3. 출고 및 이체요청 관리 (서식 출력)',
        'DISPATCH': '4. 배차 관리 (입·출고 / 창고이동)',
        'SETTLEMENT': '5. 가중량 사후정산 관리',
        'FINANCE': '6. 기업 파이낸스 (담보/재매입)'
      };
      document.getElementById('pageTitle').innerText = titleMap[tab] || 'MeatFlow ERP';

      const sub = document.getElementById('subTabs');
      if (tab === 'INBOUND') {
        currentSub = 'IN_REQUEST';
        sub.style.display = 'flex';
        sub.innerHTML = `
          <div class="sub-tab active" onclick="setSub('IN_REQUEST', this)">1. 입고요청</div>
          <div class="sub-tab" onclick="setSub('IN_CONFIRM', this)">2. 입고확정</div>
          <div class="sub-tab" onclick="setSub('IN_DONE', this)">3. 입고완료</div>
        `;
      } else if (tab === 'OUTBOUND') {
        currentSub = 'OUT_REQUEST';
        sub.style.display = 'flex';
        sub.innerHTML = `
          <div class="sub-tab active" onclick="setSub('OUT_REQUEST', this)">1. 출고/이체 요청</div>
          <div class="sub-tab" onclick="setSub('OUT_CONFIRM', this)">2. 확정 리스트 (수정불가/취소만가능)</div>
          <div class="sub-tab" onclick="setSub('OUT_DONE', this)">3. 완료 리스트</div>
        `;
      } else {
        sub.style.display = 'none';
      }

      renderHeaderActions();
      loadData();
    }

    function setSub(s, el) {
      currentSub = s;
      document.querySelectorAll('.sub-tab').forEach(e => e.classList.remove('active'));
      if (el) el.classList.add('active');
      renderHeaderActions();
      loadData();
    }

    function renderHeaderActions() {
      const actions = document.getElementById('headerActions');
      let html = `<button class="btn btn-outline" onclick="downloadCurrentTableExcel()"><i class="bi bi-file-earmark-excel" style="color:#16a34a;"></i> 엑셀 다운로드</button>`;

      if (currentMain === 'INBOUND' && currentSub === 'IN_REQUEST') {
        html += `<button class="btn btn-primary" onclick="openModal('INBOUND')"><i class="bi bi-plus-lg"></i> 입고 등록</button>`;
      } else if (currentMain === 'STOCK') {
        html += `
          <button class="btn btn-primary" onclick="openModal('OUTBOUND')"><i class="bi bi-cart-plus"></i> 출고/이체요청</button>
          <button class="btn btn-outline" onclick="openModal('TRANSFER')"><i class="bi bi-arrow-left-right"></i> 창고이동(전배)</button>
          <button class="btn btn-outline" onclick="openModal('RECON_DIRECT')"><i class="bi bi-speedometer2"></i> 실계근 중량보정</button>
          <button class="btn btn-outline" onclick="openModal('FINANCE')"><i class="bi bi-bank"></i> 파이낸스 담보설정</button>
        `;
      } else if (currentMain === 'OUTBOUND') {
        html += `<button class="btn btn-primary" style="background:#059669;" onclick="printDocTemplate()"><i class="bi bi-printer-fill"></i> 선택 요청서 인쇄/PDF 생성</button>`;
      } else if (currentMain === 'DISPATCH') {
        html += `<button class="btn btn-primary" onclick="openModal('DISPATCH')"><i class="bi bi-truck"></i> 신규 배차요청</button>`;
      }
      actions.innerHTML = html;
    }

    function filterMatches(item, field, keyword) {
      if (!keyword) return true;
      const kw = keyword.toLowerCase().trim();
      if (field === 'trace_no') return (item.trace_no || '').toLowerCase().includes(kw);
      if (field === 'bl_no') return (item.bl_no || '').toLowerCase().includes(kw);
      if (field === 'brand') return (item.brand || '').toLowerCase().includes(kw);
      if (field === 'grade') return (item.grade || '').toLowerCase().includes(kw);
      if (field === 'item_name') return ((item.item_name || '') + (item.cut_name || '')).toLowerCase().includes(kw);
      if (field === 'partner') return ((item.vendor || '') + (item.customer || '') + (item.partner_name || '')).toLowerCase().includes(kw);
      if (field === 'warehouse') return (item.warehouse || item.from_warehouse || '').toLowerCase().includes(kw);
      return JSON.stringify(item).toLowerCase().includes(kw);
    }

    async function loadData() {
      if (!authToken) return;
      const head = document.getElementById('tableHead'), body = document.getElementById('tableBody');
      body.innerHTML = '<tr><td colspan="17" style="text-align:center;">조회 중...</td></tr>';

      const field = document.getElementById('filter_field').value;
      const keyword = document.getElementById('filter_keyword').value;

      try {
        if (currentMain === 'INBOUND') {
          head.innerHTML = '<tr><th><input type="checkbox" onchange="toggleSelectAll(this)" /></th><th>화주(법인)</th><th>입고일자</th><th>BL NO</th><th>이력번호</th><th>EST</th><th>브랜드</th><th>등급</th><th>품목명</th><th>부위명</th><th>수량(Box)</th><th>중량(kg)</th><th>창고</th><th>상태</th><th>관리</th><th>GRID 번호</th><th>전표번호</th></tr>';
          const r = await apiFetch(`/api/inbounds?status=${currentSub}`);
          let list = await r.json();
          cachedList = list.filter(i => filterMatches(i, field, keyword));
          body.innerHTML = cachedList.map(i => `
            <tr>
              <td><input type="checkbox" class="row-chk" value="${i.inbound_no}" /></td>
              <td><span class="badge" style="background:#dbeafe; color:#1e40af;">${i.company_name}</span></td>
              <td>${i.inbound_date}</td>
              <td>${i.bl_no||'-'}</td>
              <td><code>${i.trace_no}</code></td>
              <td>${i.est_no||'-'}</td>
              <td><strong>${i.brand||'-'}</strong></td>
              <td><span class="badge" style="background:#f1f5f9;">${i.grade||'GF'}</span></td>
              <td><strong>${i.item_name}</strong></td>
              <td>${i.cut_name}</td>
              <td>${i.box_qty}</td>
              <td>${i.weight_kg}kg</td>
              <td>${i.warehouse}</td>
              <td><span class="badge" style="background:#e0f2fe; color:#0369a1;">${i.status}</span></td>
              <td>${currentSub !== 'IN_DONE' ? `<button class="btn btn-primary" style="padding:2px 6px;" onclick="advanceInbound('${i.grid_no}')">다음단계</button>` : '-'}</td>
              <td><span class="grid-tag">${i.grid_no}</span></td>
              <td><code>${i.inbound_no}</code></td>
            </tr>
          `).join('') || '<tr><td colspan="17" style="text-align:center;">내역이 없습니다.</td></tr>';
        } else if (currentMain === 'STOCK') {
          head.innerHTML = '<tr><th><input type="checkbox" onchange="toggleSelectAll(this)" /></th><th>화주(법인)</th><th>BL No</th><th>이력번호</th><th>EST</th><th>브랜드</th><th>등급</th><th>품목명</th><th>부위명</th><th>현재고(Box)</th><th>현재고(kg)</th><th>예약(Box)</th><th>파이낸스 담보(매입처)</th><th>창고</th><th>계근여부</th><th>소비기한</th><th>GRID 번호</th></tr>';
          const r = await apiFetch('/api/inventory');
          let list = await r.json();
          cachedList = list.filter(l => filterMatches(l, field, keyword));
          body.innerHTML = cachedList.map(l => `
            <tr>
              <td><input type="checkbox" class="row-chk" value="${l.grid_no}" /></td>
              <td><span class="badge" style="background:#dbeafe; color:#1e40af;">${l.company_name}</span></td>
              <td>${l.bl_no||'-'}</td>
              <td><code>${l.trace_no}</code></td>
              <td>${l.est_no||'-'}</td>
              <td><strong>${l.brand||'-'}</strong></td>
              <td><span class="badge" style="background:#f1f5f9;">${l.grade||'GF'}</span></td>
              <td><strong>${l.item_name}</strong></td>
              <td>${l.cut_name}</td>
              <td><strong>${l.current_box_qty}</strong></td>
              <td>${l.current_weight_kg}kg</td>
              <td style="color:#0284c7;">${l.reserved_box_qty} Box</td>
              <td><span class="badge" style="background:#fef3c7; color:#92400e; font-weight:700;">${l.pledged_desc}</span></td>
              <td>${l.warehouse}</td>
              <td><span class="badge" style="background:${l.is_weighed==='Y'?'#dcfce7':'#f1f5f9'};">${l.is_weighed==='Y'?'계근완료':'미계근'}</span></td>
              <td>${l.exp_date}</td>
              <td><span class="grid-tag">${l.grid_no}</span></td>
            </tr>
          `).join('') || '<tr><td colspan="17" style="text-align:center;">재고가 없습니다.</td></tr>';
        } else if (currentMain === 'OUTBOUND') {
          head.innerHTML = `<tr>
            <th><input type="checkbox" onchange="toggleSelectAll(this)" /></th>
            <th>구분</th><th>가중량</th><th>화주(법인)</th><th>출고/이체처</th><th>BL No</th><th>이력번호</th><th>브랜드</th><th>등급</th><th>품목명</th><th>부위명</th><th>수량(Box)</th><th>중량(kg)</th><th>창고</th><th>상태</th><th>관리</th><th>GRID 번호</th><th>전표번호</th>
          </tr>`;
          const r = await apiFetch(`/api/outbounds?status=${currentSub}`);
          let list = await r.json();
          cachedList = list.filter(o => filterMatches(o, field, keyword));
          body.innerHTML = cachedList.map(o => `
            <tr>
              <td><input type="checkbox" class="row-chk out-chk" value="${o.outbound_no}" /></td>
              <td><span class="badge" style="background:#f3e8ff; color:#6b21a8;">${o.trade_type}</span></td>
              <td><span class="badge" style="background:${o.is_estimated==='Y'?'#fee2e2':'#f1f5f9'}; color:${o.is_estimated==='Y'?'#b91c1c':'#475569'};">${o.is_estimated==='Y'?'가중량':'실중량'}</span></td>
              <td>${o.company_name}</td>
              <td><strong>${o.customer}</strong></td>
              <td>${o.bl_no||'-'}</td>
              <td><code>${o.trace_no}</code></td>
              <td><strong>${o.brand||'-'}</strong></td>
              <td><span class="badge" style="background:#f1f5f9;">${o.grade||'GF'}</span></td>
              <td><strong>${o.item_name}</strong></td>
              <td>${o.cut_name}</td>
              <td>${o.box_qty}</td>
              <td>${o.weight_kg}kg</td>
              <td>${o.warehouse}</td>
              <td><span class="badge" style="background:#fef3c7; color:#92400e;">${o.status}</span></td>
              <td>
                ${currentSub === 'OUT_REQUEST' ? `<button class="btn btn-primary" style="padding:2px 6px;" onclick="advanceOutbound('${o.outbound_no}')">확정</button>` : ''}
                <button class="btn btn-outline" style="color:#ef4444; padding:2px 6px;" onclick="revertOutbound('${o.outbound_no}')">취소/반려</button>
              </td>
              <td><span class="grid-tag">${o.grid_no||'-'}</span></td>
              <td><code>${o.outbound_no}</code></td>
            </tr>
          `).join('') || '<tr><td colspan="18" style="text-align:center;">내역이 없습니다.</td></tr>';
        } else if (currentMain === 'DISPATCH') {
          head.innerHTML = '<tr><th><input type="checkbox" onchange="toggleSelectAll(this)" /></th><th>구분</th><th>일자</th><th>거래처/내역</th><th>브랜드</th><th>등급</th><th>품목명</th><th>부위명</th><th>수량(Box)</th><th>중량(kg)</th><th>출고창고</th><th>납품지(도착지)</th><th>파렛트적재</th><th>배차금액</th><th>상태</th><th>배차번호</th></tr>';
          const r = await apiFetch('/api/dispatches');
          let list = await r.json();
          cachedList = list.filter(d => filterMatches(d, field, keyword));
          body.innerHTML = cachedList.map(d => `
            <tr>
              <td><input type="checkbox" class="row-chk" value="${d.dispatch_no}" /></td>
              <td><span class="badge" style="background:#e0e7ff; color:#3730a3;">${d.dispatch_type}</span></td>
              <td>${d.target_date}</td>
              <td><strong>${d.partner_name}</strong></td>
              <td>${d.brand||'-'}</td>
              <td>${d.grade||'GF'}</td>
              <td><strong>${d.item_name}</strong></td>
              <td>${d.cut_name}</td>
              <td>${d.box_qty}</td>
              <td>${d.weight_kg}kg</td>
              <td>${d.from_warehouse}</td>
              <td>${d.to_destination}</td>
              <td><span class="badge" style="background:${d.pallet_loaded!=='N'?'#dcfce7':'#f1f5f9'};">${d.pallet_loaded}</span></td>
              <td><strong>${d.dispatch_cost.toLocaleString()}원</strong></td>
              <td><span class="badge" style="background:#f1f5f9;">${d.status}</span></td>
              <td><span class="grid-tag">${d.dispatch_no}</span></td>
            </tr>
          `).join('') || '<tr><td colspan="16" style="text-align:center;">배차 내역이 없습니다.</td></tr>';
        } else if (currentMain === 'SETTLEMENT') {
          head.innerHTML = '<tr><th><input type="checkbox" onchange="toggleSelectAll(this)" /></th><th>화주</th><th>거래처</th><th>브랜드</th><th>등급</th><th>품목명</th><th>부위명</th><th>가중량(kg)</th><th>실계근(kg)</th><th>차액(kg)</th><th>단가</th><th>정산차액</th><th>상태</th><th>관리</th><th>전표번호</th></tr>';
          const r = await apiFetch('/api/settlements');
          let list = await r.json();
          cachedList = list.filter(s => filterMatches(s, field, keyword));
          body.innerHTML = cachedList.map(s => {
            const diffW = s.actual_weight_kg ? (s.actual_weight_kg - s.weight_kg).toFixed(2) : '-';
            return `
              <tr>
                <td><input type="checkbox" class="row-chk" value="${s.outbound_no}" /></td>
                <td>${s.company_name}</td>
                <td>${s.customer}</td>
                <td>${s.brand||'-'}</td>
                <td>${s.grade||'GF'}</td>
                <td><strong>${s.item_name}</strong></td>
                <td>${s.cut_name}</td>
                <td>${s.weight_kg}kg</td>
                <td><strong>${s.actual_weight_kg ? s.actual_weight_kg + 'kg' : '미입력'}</strong></td>
                <td style="color:${diffW < 0 ? '#ef4444' : '#10b981'}; font-weight:700;">${diffW !== '-' ? (diffW > 0 ? '+' : '') + diffW + 'kg' : '-'}</td>
                <td>${s.unit_price_kg.toLocaleString()}원</td>
                <td><strong>${s.reconciled_amount ? s.reconciled_amount.toLocaleString() + '원' : '-'}</strong></td>
                <td><span class="badge" style="background:${s.reconciled_status==='RECONCILED'?'#dcfce7':'#fee2e2'};">${s.reconciled_status}</span></td>
                <td><button class="btn btn-outline" style="padding:2px 6px;" onclick="promptReconcile('${s.outbound_no}', ${s.weight_kg})">실계근입력</button></td>
                <td><span class="grid-tag">${s.outbound_no}</span></td>
              </tr>
            `;
          }).join('') || '<tr><td colspan="15" style="text-align:center;">정산 대상이 없습니다.</td></tr>';
        } else if (currentMain === 'FINANCE') {
          head.innerHTML = '<tr><th><input type="checkbox" onchange="toggleSelectAll(this)" /></th><th>담보처(매입회사)</th><th>담보수량(Box)</th><th>담보중량(kg)</th><th>수령 보증금</th><th>수수료율</th><th>실행일</th><th>만기일</th><th>상태</th><th>계약번호</th></tr>';
          const r = await apiFetch('/api/finances');
          let list = await r.json();
          cachedList = list.filter(f => filterMatches(f, field, keyword));
          body.innerHTML = cachedList.map(f => `
            <tr>
              <td><input type="checkbox" class="row-chk" value="${f.contract_no}" /></td>
              <td><strong>${f.partner_company}</strong></td>
              <td>${f.box_qty} Box</td>
              <td>${f.weight_kg}kg</td>
              <td>${f.deposit_amount.toLocaleString()}원</td>
              <td>${f.fee_rate}%</td>
              <td>${f.pledge_date}</td>
              <td style="color:#b91c1c;">${f.due_date}</td>
              <td><span class="badge" style="background:#fef3c7;">${f.status}</span></td>
              <td><span class="grid-tag">${f.contract_no}</span></td>
            </tr>
          `).join('') || '<tr><td colspan="10" style="text-align:center;">파이낸스 내역이 없습니다.</td></tr>';
        }
      } catch(err) {
        console.error(err);
      }
    }

    function toggleSelectAll(master) {
      document.querySelectorAll('.row-chk').forEach(c => c.checked = master.checked);
    }

    function printDocTemplate() {
      const selectedNos = Array.from(document.querySelectorAll('.out-chk:checked')).map(c => c.value);
      if (!selectedNos.length) {
        alert('출력할 전표 항목을 체크박스로 선택해주세요.');
        return;
      }

      const targets = cachedList.filter(o => selectedNos.includes(o.outbound_no));
      const first = targets[0];
      const comp = companiesCache.find(c => c.name === first.company_name) || {
        name: first.company_name, biz_no: '718-88-03523', rep_name: '김다희',
        address: '경기도 남양주시 별내중앙로 34, 5층 502-A59호', phone: '010-5773-0619', fax: '070-4758-9219'
      };

      const isTransfer = first.trade_type.includes('이체');
      document.getElementById('docMainTitle').innerText = isTransfer ? '이 체 요 청 서' : '출 고 요 청 서';
      document.getElementById('docTargetCol').innerText = isTransfer ? '이체처' : '출고처';
      document.getElementById('docCompanyBadge').innerText = comp.name.replace('(주)', '').replace('주식회사', '').trim();
      document.getElementById('docVendorChain').innerText = first.vendor_chain || '미코상사';
      document.getElementById('docWarehouse').innerText = first.warehouse;
      document.getElementById('docDate').innerText = first.outbound_date;

      document.getElementById('docBottomMsg').innerText = isTransfer
        ? '상기와 같이 요청하오니 이체하여 주시기 바랍니다.'
        : '상기와 같이 요청하오니 출고하여 주시기 바랍니다.(계근 ㅇ)';

      let rowsHtml = '';
      let sumBox = 0;
      for (let i = 0; i < 11; i++) {
        if (i < targets.length) {
          const t = targets[i];
          sumBox += t.box_qty;
          rowsHtml += `
            <tr>
              <td>${i + 1}</td>
              <td>${t.bl_no || 'NA'}</td>
              <td>${t.brand || ''}</td>
              <td>${t.grade || 'GF'}</td>
              <td style="text-align:left; padding-left:8px;">${t.cut_name || t.item_name}</td>
              <td>${t.box_qty}</td>
              <td>${t.customer}</td>
              <td>${t.remark || ''}</td>
            </tr>
          `;
        } else {
          rowsHtml += `<tr><td>${i + 1}</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>`;
        }
      }
      document.getElementById('docGridBody').innerHTML = rowsHtml;
      document.getElementById('docSumBox').innerText = sumBox;

      document.getElementById('f_biz_no').innerText = comp.biz_no;
      document.getElementById('f_comp_name').innerText = comp.name;
      document.getElementById('f_address').innerText = comp.address;
      document.getElementById('f_rep_name').innerText = comp.rep_name;
      document.getElementById('f_tel').innerText = comp.phone || '-';
      document.getElementById('f_fax').innerText = comp.fax || '-';
      document.getElementById('f_huge_title').innerText = comp.name;
      document.getElementById('stampText').innerHTML = comp.name.replace('(주)', '').trim() + '<br>인';

      const wrap = document.getElementById('printWrapper');
      wrap.style.display = 'block';
      window.print();
      wrap.style.display = 'none';
    }

    async function advanceInbound(gridNo) {
      const res = await apiFetch(`/api/inbounds/${gridNo}/advance`, { method: 'POST' });
      const d = await res.json();
      alert(d.message);
      loadData();
    }

    async function advanceOutbound(outNo) {
      const res = await apiFetch(`/api/outbounds/${outNo}/advance`, { method: 'POST' });
      const d = await res.json();
      alert(d.message);
      loadData();
    }

    async function revertOutbound(outNo) {
      if (!confirm('정말 취소/반려 처리하시겠습니까?')) return;
      const res = await apiFetch(`/api/outbounds/${outNo}/revert`, { method: 'POST' });
      const d = await res.json();
      alert(d.message);
      loadData();
    }

    async function promptReconcile(outNo, estW) {
      const act = prompt(`가중량(${estW}kg)에 대한 도착 실계근 중량(kg)을 입력하세요:`, estW);
      if (!act || isNaN(act)) return;
      const res = await apiFetch('/api/settlements/reconcile', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ outbound_no: outNo, actual_weight_kg: parseFloat(act) })
      });
      const d = await res.json();
      alert(d.message);
      loadData();
    }

    function openModal(type) {
      const modal = document.getElementById('actionModal');
      const title = document.getElementById('modalTitle');
      const body = document.getElementById('modalFormBody');

      if (type === 'INBOUND') {
        title.innerText = '신규 입고 등록';
        body.innerHTML = `
          <div class="form-group"><label>화주(당사 법인)</label>
            <select id="m_in_comp" class="form-control">
              <option value="주식회사 티제이에프">주식회사 티제이에프</option>
              <option value="(주)서울웰푸드">(주)서울웰푸드</option>
              <option value="(주)넥서스트레이딩">(주)넥서스트레이딩</option>
            </select>
          </div>
          <div class="form-group"><label>입고일자</label><input type="date" id="m_in_date" class="form-control" value="${getTodayStr()}" /></div>
          <div class="form-group"><label>매입처</label><input type="text" id="m_in_vendor" class="form-control" value="(주)글로벌미트" /></div>
          <div class="form-group"><label>BL NO</label><input type="text" id="m_in_bl" class="form-control" placeholder="예: OOLU2766442080" /></div>
          <div class="form-group"><label>이력번호(12자리)</label><input type="text" id="m_in_trace" class="form-control" placeholder="이력번호" /></div>
          <div class="form-group"><label>EST(가공장)</label><input type="text" id="m_in_est" class="form-control" value="244I" /></div>
          <div class="form-group"><label>브랜드</label><input type="text" id="m_in_brand" class="form-control" value="AMH" /></div>
          <div class="form-group"><label>등급</label><input type="text" id="m_in_grade" class="form-control" value="GF" /></div>
          <div class="form-group"><label>품목</label>
            <select id="m_in_item" class="form-control">
              <option value="우육(소)">우육(소)</option><option value="돈육(돼지)">돈육(돼지)</option>
            </select>
          </div>
          <div class="form-group"><label>부위명</label><input type="text" id="m_in_cut" class="form-control" value="소일반갈비(170)" /></div>
          <div class="form-group"><label>보관방식</label>
            <select id="m_in_storage" class="form-control">
              <option value="냉동">냉동</option><option value="냉장">냉장</option>
            </select>
          </div>
          <div class="form-group"><label>창고</label><input type="text" id="m_in_wh" class="form-control" value="강동2" /></div>
          <div class="form-group"><label>수량(Box)</label><input type="number" id="m_in_box" class="form-control" value="100" /></div>
          <div class="form-group"><label>총중량(kg)</label><input type="number" step="0.1" id="m_in_weight" class="form-control" value="2000.0" /></div>
          <div class="form-group"><label>단가(원/kg)</label><input type="number" id="m_in_cost" class="form-control" value="14500" /></div>
          <div class="form-group"><label>소비기한</label><input type="date" id="m_in_exp" class="form-control" value="${new Date(Date.now() + 730*86400000).toISOString().slice(0,10)}" /></div>
        `;
      } else if (type === 'OUTBOUND') {
        title.innerText = '출고 및 이체 요청 등록';
        body.innerHTML = `
          <div class="form-group"><label>거래 형태</label>
            <select id="m_out_type" class="form-control">
              <option value="출고판매">출고판매 (실물 배차출하)</option>
              <option value="이체판매">이체판매 (창고 내 화주변경)</option>
            </select>
          </div>
          <div class="form-group">
            <label style="color:#b91c1c;"><input type="checkbox" id="m_out_is_est" checked /> 가중량 판매 여부 (체크 시 가중량 사후정산리스트 자동 인입)</label>
          </div>
          <div class="form-group"><label>재고 GRID 번호</label><input type="text" id="m_out_grid" class="form-control" placeholder="재고장의 GRID-XXXXXX 입력" /></div>
          <div class="form-group"><label>화주(당사 법인)</label>
            <select id="m_out_comp" class="form-control">
              <option value="주식회사 티제이에프">주식회사 티제이에프</option>
              <option value="(주)서울웰푸드">(주)서울웰푸드</option>
              <option value="(주)넥서스트레이딩">(주)넥서스트레이딩</option>
            </select>
          </div>
          <div class="form-group"><label>매입처 표기 (서식 상단)</label><input type="text" id="m_out_vchain" class="form-control" value="브루니" placeholder="예: 브루니 또는 넥서스>미코상사" /></div>
          <div class="form-group"><label>출고/이체처(고객사)</label><input type="text" id="m_out_cust" class="form-control" placeholder="예: 에스티에프" /></div>
          <div class="form-group"><label>수량(Box)</label><input type="number" id="m_out_box" class="form-control" value="10" /></div>
          <div class="form-group"><label>가중량(kg) (미입력 시 평중 자동계산)</label><input type="number" step="0.1" id="m_out_weight" class="form-control" placeholder="직접 입력 시 우선 적용" /></div>
          <div class="form-group"><label>판매단가(원/kg)</label><input type="number" id="m_out_price" class="form-control" value="15500" /></div>
          <div class="form-group"><label>비고</label><input type="text" id="m_out_remark" class="form-control" /></div>
        `;
      } else if (type === 'TRANSFER') {
        title.innerText = '자사 재고 창고이동(전배) 및 배차연동';
        body.innerHTML = `
          <div class="form-group"><label>이동 대상 재고 GRID</label><input type="text" id="m_tr_grid" class="form-control" placeholder="GRID-XXXXXX" /></div>
          <div class="form-group"><label>도착 창고명</label><input type="text" id="m_tr_wh" class="form-control" placeholder="예: 삼진1냉장" /></div>
          <div class="form-group"><label>이동 수량(Box)</label><input type="number" id="m_tr_box" class="form-control" value="50" /></div>
          <div class="form-group"><label><input type="checkbox" id="m_tr_disp" checked /> 배차 관리리스트에 동시 등록</label></div>
          <div class="form-group"><label>배차 운송금액(원)</label><input type="number" id="m_tr_cost" class="form-control" value="150000" /></div>
          <div class="form-group"><label>파렛트 적재 여부</label><input type="text" id="m_tr_plt" class="form-control" value="Y (2 PLT)" /></div>
        `;
      } else if (type === 'RECON_DIRECT') {
        title.innerText = '재고장 실계근 중량 직접 보정';
        body.innerHTML = `
          <div class="form-group"><label>보정 대상 GRID</label><input type="text" id="m_rc_grid" class="form-control" placeholder="GRID-XXXXXX" /></div>
          <div class="form-group"><label>창고 실계근 총중량(kg)</label><input type="number" step="0.1" id="m_rc_weight" class="form-control" /></div>
          <div class="form-group"><label>보정 사유</label><input type="text" id="m_rc_reason" class="form-control" value="하차 계근 오차 정산" /></div>
        `;
      } else if (type === 'DISPATCH') {
        title.innerText = '신규 차량 배차 요청';
        body.innerHTML = `
          <div class="form-group"><label>배차 구분</label>
            <select id="m_dsp_type" class="form-control">
              <option value="출고배차">출고배차</option>
              <option value="입고배차">입고배차</option>
              <option value="창고이동(전배)">창고이동(전배)</option>
            </select>
          </div>
          <div class="form-group"><label>거래처명</label><input type="text" id="m_dsp_partner" class="form-control" placeholder="예: 플랜비에프에스" /></div>
          <div class="form-group"><label>브랜드</label><input type="text" id="m_dsp_brand" class="form-control" value="코엑스카" /></div>
          <div class="form-group"><label>등급</label><input type="text" id="m_dsp_grade" class="form-control" value="NA" /></div>
          <div class="form-group"><label>품목명</label><input type="text" id="m_dsp_item" class="form-control" value="돈육(돼지)" /></div>
          <div class="form-group"><label>부위명</label><input type="text" id="m_dsp_cut" class="form-control" value="돈갈매기" /></div>
          <div class="form-group"><label>수량(Box)</label><input type="number" id="m_dsp_box" class="form-control" value="10" /></div>
          <div class="form-group"><label>중량(kg)</label><input type="number" step="0.1" id="m_dsp_weight" class="form-control" value="200.0" /></div>
          <div class="form-group"><label>출고창고(출발지)</label><input type="text" id="m_dsp_from" class="form-control" value="삼진1냉장" /></div>
          <div class="form-group"><label>납품지(도착지)</label><input type="text" id="m_dsp_to" class="form-control" value="경기도 광주시 플랜비 물류센터" /></div>
          <div class="form-group"><label>파렛트 적재여부</label><input type="text" id="m_dsp_plt" class="form-control" value="Y (1 PLT)" /></div>
          <div class="form-group"><label>배차금액(원)</label><input type="number" id="m_dsp_cost" class="form-control" value="80000" /></div>
        `;
      } else if (type === 'FINANCE') {
        title.innerText = '기업 파이낸스 보증금 담보 등록';
        body.innerHTML = `
          <div class="form-group"><label>재고 GRID 번호</label><input type="text" id="m_f_grid" class="form-control" /></div>
          <div class="form-group"><label>담보처(매입회사)</label><input type="text" id="m_f_partner" class="form-control" value="(주)대한육류유통" /></div>
          <div class="form-group"><label>담보 수량(Box)</label><input type="number" id="m_f_box" class="form-control" value="50" /></div>
          <div class="form-group"><label>수령 보증금(선급금)</label><input type="number" id="m_f_deposit" class="form-control" value="5000000" /></div>
          <div class="form-group"><label>약정 수수료율(%)</label><input type="number" step="0.1" id="m_f_fee" class="form-control" value="2.5" /></div>
          <div class="form-group"><label>실행일</label><input type="date" id="m_f_pdate" class="form-control" value="${getTodayStr()}" /></div>
          <div class="form-group"><label>만기일</label><input type="date" id="m_f_ddate" class="form-control" value="${new Date(Date.now() + 30*86400000).toISOString().slice(0,10)}" /></div>
        `;
      }

      modal.style.display = 'flex';
    }

    function closeModal() {
      document.getElementById('actionModal').style.display = 'none';
    }

    async function submitModal() {
      const title = document.getElementById('modalTitle').innerText;
      try {
        if (title.includes('입고 등록')) {
          const payload = {
            company_name: document.getElementById('m_in_comp').value,
            inbound_date: document.getElementById('m_in_date').value,
            vendor: document.getElementById('m_in_vendor').value,
            bl_no: document.getElementById('m_in_bl').value,
            trace_no: document.getElementById('m_in_trace').value || 'TR-' + Math.floor(Math.random()*900000),
            est_no: document.getElementById('m_in_est').value,
            brand: document.getElementById('m_in_brand').value,
            grade: document.getElementById('m_in_grade').value,
            item_name: document.getElementById('m_in_item').value,
            cut_name: document.getElementById('m_in_cut').value,
            storage_type: document.getElementById('m_in_storage').value,
            warehouse: document.getElementById('m_in_wh').value,
            box_qty: parseInt(document.getElementById('m_in_box').value),
            weight_kg: parseFloat(document.getElementById('m_in_weight').value),
            cost_per_kg: parseFloat(document.getElementById('m_in_cost').value),
            exp_date: document.getElementById('m_in_exp').value
          };
          const res = await apiFetch('/api/inbounds', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
          const d = await res.json();
          alert(d.message);
        } else if (title.includes('출고 및 이체')) {
          const wVal = document.getElementById('m_out_weight').value;
          const payload = {
            trade_type: document.getElementById('m_out_type').value,
            is_estimated: document.getElementById('m_out_is_est').checked ? "Y" : "N",
            grid_no: document.getElementById('m_out_grid').value,
            company_name: document.getElementById('m_out_comp').value,
            vendor_chain: document.getElementById('m_out_vchain').value,
            customer: document.getElementById('m_out_cust').value,
            box_qty: parseInt(document.getElementById('m_out_box').value),
            weight_kg: wVal ? parseFloat(wVal) : null,
            unit_price_kg: parseFloat(document.getElementById('m_out_price').value),
            remark: document.getElementById('m_out_remark').value
          };
          const res = await apiFetch('/api/outbounds/create', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
          const d = await res.json();
          alert(d.message);
        } else if (title.includes('창고이동')) {
          const payload = {
            grid_no: document.getElementById('m_tr_grid').value,
            to_warehouse: document.getElementById('m_tr_wh').value,
            box_qty: parseInt(document.getElementById('m_tr_box').value),
            create_dispatch: document.getElementById('m_tr_disp').checked,
            dispatch_cost: parseFloat(document.getElementById('m_tr_cost').value),
            pallet_loaded: document.getElementById('m_tr_plt').value
          };
          const res = await apiFetch('/api/inventory/transfer-warehouse', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
          const d = await res.json();
          alert(d.message);
        } else if (title.includes('직접 보정')) {
          const payload = {
            grid_no: document.getElementById('m_rc_grid').value,
            actual_weight_kg: parseFloat(document.getElementById('m_rc_weight').value),
            reason: document.getElementById('m_rc_reason').value
          };
          const res = await apiFetch('/api/inventory/reconcile-weight-direct', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
          const d = await res.json();
          alert(d.message);
        } else if (title.includes('배차 요청')) {
          const payload = {
            dispatch_type: document.getElementById('m_dsp_type').value,
            partner_name: document.getElementById('m_dsp_partner').value,
            brand: document.getElementById('m_dsp_brand').value,
            grade: document.getElementById('m_dsp_grade').value,
            item_name: document.getElementById('m_dsp_item').value,
            cut_name: document.getElementById('m_dsp_cut').value,
            box_qty: parseInt(document.getElementById('m_dsp_box').value),
            weight_kg: parseFloat(document.getElementById('m_dsp_weight').value),
            from_warehouse: document.getElementById('m_dsp_from').value,
            to_destination: document.getElementById('m_dsp_to').value,
            pallet_loaded: document.getElementById('m_dsp_plt').value,
            dispatch_cost: parseFloat(document.getElementById('m_dsp_cost').value)
          };
          const res = await apiFetch('/api/dispatches', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
          const d = await res.json();
          alert(d.message);
        } else if (title.includes('파이낸스')) {
          const payload = {
            grid_no: document.getElementById('m_f_grid').value,
            partner_company: document.getElementById('m_f_partner').value,
            box_qty: parseInt(document.getElementById('m_f_box').value),
            deposit_amount: parseFloat(document.getElementById('m_f_deposit').value),
            fee_rate: parseFloat(document.getElementById('m_f_fee').value),
            pledge_date: document.getElementById('m_f_pdate').value,
            due_date: document.getElementById('m_f_ddate').value
          };
          const res = await apiFetch('/api/finances', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
          const d = await res.json();
          alert(d.message);
        }
        closeModal();
        loadData();
      } catch(e) {
        alert('처리 오류');
      }
    }

    function downloadCurrentTableExcel() {
      const table = document.querySelector('.table-card table');
      const wb = XLSX.utils.table_to_book(table, { sheet: "Sheet1" });
      XLSX.writeFile(wb, `MeatFlow_${currentMain}_${new Date().toISOString().slice(0,10)}.xlsx`);
    }

    window.addEventListener('DOMContentLoaded', async () => {
      if (authToken) {
        document.getElementById('loginModal').style.display = 'none';
        await fetchCompanies();
        resetFilters();
      }
    });
  </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def serve_dashboard():
    return HTML_PAGE

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port)
