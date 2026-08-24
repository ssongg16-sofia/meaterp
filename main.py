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
from jose import JWTError, jwt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pydantic import BaseModel, ConfigDict
from sqlalchemy import (
    Column, Integer, Float, String, DateTime, ForeignKey, create_engine, desc, func, Index, text
)
from sqlalchemy.orm import declarative_base, sessionmaker, Session
import uvicorn

# -----------------------------------------------------------------------------
# 1. 인증 및 환경 변수 설정
# -----------------------------------------------------------------------------
JWT_SECRET_KEY = os.environ.get("JWT_SECRET_KEY", "seoulwellfood-g3-enterprise-secret-2026")
JWT_ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 12

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/auth/login", auto_error=False)

ROLE_ADMIN = "ADMIN"
ROLE_SALES = "SALES"
ROLE_SALES_SUPPORT = "SALES_SUPPORT"
ROLE_WAREHOUSE = "WAREHOUSE"

# -----------------------------------------------------------------------------
# 2. 데이터베이스 설정 (PostgreSQL / SQLite 호환)
# -----------------------------------------------------------------------------
DB_FILE = os.environ.get("DB_FILE", "seoulwellfood_erp.db")
DATABASE_URL = os.environ.get("DATABASE_URL", f"sqlite:///{DB_FILE}")

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
# 3. 데이터 모델 정의
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

class WarehouseRate(Base):
    __tablename__ = "warehouse_rates"
    id = Column(Integer, primary_key=True, index=True)
    warehouse_name = Column(String(50), unique=True, nullable=False)
    frozen_rate = Column(Float, default=1.5)  # 원/kg/일
    chilled_rate = Column(Float, default=2.5)
    handling_in_fee = Column(Float, default=10.0)  # 입고 상하차비
    handling_out_fee = Column(Float, default=10.0)

class InventoryLot(Base):
    __tablename__ = "inventory_lots"
    id = Column(Integer, primary_key=True, index=True)
    grid_no = Column(String(50), unique=True, nullable=False, index=True)
    company_name = Column(String(100), default="(주)서울웰푸드")
    sku_code = Column(String(50), nullable=False)
    inbound_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d"))
    bl_no = Column(String(50), nullable=True)
    trace_no = Column(String(50), nullable=False, index=True)
    est_no = Column(String(50), nullable=True)
    brand = Column(String(50), nullable=True)
    grade = Column(String(30), default="GF")
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)  # 냉동, 냉장, 선적중, ❄️냉동전환
    initial_box_qty = Column(Integer, default=0)
    initial_weight_kg = Column(Float, default=0.0)
    avg_box_weight = Column(Float, default=0.0)
    current_box_qty = Column(Integer, nullable=False)
    current_weight_kg = Column(Float, nullable=False)
    cost_per_kg = Column(Float, nullable=False)
    warehouse = Column(String(50), nullable=False)
    exp_date = Column(String(20), nullable=False, index=True)
    is_weighed = Column(String(10), default="N")
    status = Column(String(20), default="IN_DONE")  # ON_WATER, IN_DONE, FREEZE_CONVERTED
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

class ReservationRecord(Base):
    __tablename__ = "reservations"
    id = Column(Integer, primary_key=True, index=True)
    res_no = Column(String(50), unique=True, index=True)
    grid_no = Column(String(50), nullable=False, index=True)
    sales_rep = Column(String(50), nullable=False)
    sales_rep_name = Column(String(50), nullable=False)
    customer = Column(String(100), nullable=False)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    unit_price_kg = Column(Float, nullable=False)
    total_amount = Column(Float, nullable=False)
    expire_date = Column(String(20), nullable=True)
    status = Column(String(20), default="HOLD", index=True)  # HOLD, PROMOTED, CANCELLED
    created_at = Column(DateTime, default=datetime.now)

class PartnerFinanceContract(Base):
    __tablename__ = "partner_finances"
    id = Column(Integer, primary_key=True, index=True)
    contract_no = Column(String(50), unique=True, index=True)
    grid_no = Column(String(50), nullable=False, index=True)
    partner_company = Column(String(100), nullable=False)  # 프로즌파트너스, 동원홈푸드, 미트박스
    finance_model = Column(String(30), default="BUY_BACK")  # BUY_BACK, IMPORT_AGENCY
    pledge_date = Column(String(20), nullable=False)
    due_date = Column(String(20), nullable=False)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    base_unit_price = Column(Float, nullable=False)
    deposit_amount = Column(Float, default=0.0)
    annual_margin_rate = Column(Float, default=7.5)
    extension_count = Column(Integer, default=0)
    status = Column(String(20), default="HOLD_BY_PARTNER", index=True)
    created_at = Column(DateTime, default=datetime.now)

class OutboundRecord(Base):
    __tablename__ = "outbounds"
    id = Column(Integer, primary_key=True, index=True)
    outbound_no = Column(String(50), unique=True, index=True)
    company_name = Column(String(100), default="(주)서울웰푸드")
    vendor_chain = Column(String(100), nullable=True)
    trade_type = Column(String(30), default="출고판매")  # 출고판매, 이체판매
    is_estimated = Column(String(10), default="Y")
    outbound_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d"))
    grid_no = Column(String(50), nullable=False, index=True)
    customer = Column(String(100), nullable=False)
    bl_no = Column(String(50), nullable=True)
    trace_no = Column(String(50), nullable=False)
    brand = Column(String(50), nullable=True)
    grade = Column(String(30), default="GF")
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)  # 가중량
    actual_weight_kg = Column(Float, nullable=True)  # 실계근
    reconciled_amount = Column(Float, default=0.0)
    reconciled_status = Column(String(20), default="UNRECONCILED")
    unit_price_kg = Column(Float, nullable=False)
    total_amount = Column(Float, nullable=False)
    payment_term = Column(String(50), default="외상 30일 회전")
    due_date = Column(String(20), nullable=True)
    warehouse = Column(String(50), nullable=False)
    status = Column(String(20), default="OUT_REQUEST", index=True)  # OUT_REQUEST, OUT_CONFIRM, OUT_DONE, OUT_HELD
    remark = Column(String(255), nullable=True)
    created_at = Column(DateTime, default=datetime.now)

class DispatchRecord(Base):
    __tablename__ = "dispatches"
    id = Column(Integer, primary_key=True, index=True)
    dispatch_no = Column(String(50), unique=True, index=True)
    dispatch_type = Column(String(30), default="출고배차")
    target_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d"))
    partner_name = Column(String(100), nullable=False)
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    from_warehouse = Column(String(100), nullable=False)
    to_destination = Column(String(200), nullable=False)
    pallet_loaded = Column(String(20), default="N")
    dispatch_cost = Column(Float, default=0.0)
    cost_bearer = Column(String(30), default="당사부담(선불)")
    status = Column(String(20), default="REQUESTED")
    created_at = Column(DateTime, default=datetime.now)

class AccountsReceivable(Base):
    __tablename__ = "accounts_receivable"
    id = Column(Integer, primary_key=True, index=True)
    ar_no = Column(String(50), unique=True, index=True)
    company_name = Column(String(100), default="(주)서울웰푸드")
    customer_name = Column(String(100), nullable=False, index=True)
    outbound_no = Column(String(50), nullable=False, index=True)
    sales_date = Column(String(20), nullable=False)
    due_date = Column(String(20), nullable=False)
    sales_amount = Column(Float, nullable=False)
    weight_diff_amount = Column(Float, default=0.0)
    final_bill_amount = Column(Float, nullable=False)
    paid_amount = Column(Float, default=0.0)
    status = Column(String(20), default="UNPAID", index=True)  # UNPAID, PARTIAL, PAID, OVERDUE

class DevIssue(Base):
    __tablename__ = "dev_issues"
    id = Column(Integer, primary_key=True, index=True)
    issue_type = Column(String(30), default="기능개선")  # 기능개선, 오류/버그, 데이터정정
    title = Column(String(200), nullable=False)
    content = Column(String(1000), nullable=False)
    related_menu = Column(String(50), default="재고관리")
    grid_link = Column(String(50), nullable=True)
    author = Column(String(50), nullable=False)
    status = Column(String(20), default="SUBMITTED")  # SUBMITTED, REVIEWING, IN_PROGRESS, RESOLVED
    created_at = Column(DateTime, default=datetime.now)

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
# 4. 유틸리티 & 비밀번호 검증
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
    credentials_exception = HTTPException(status_code=401, detail="로그인이 필요합니다.")
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

# -----------------------------------------------------------------------------
# 5. 백데이터 마이그레이션 및 시드 데이터 초기화
# -----------------------------------------------------------------------------
def init_system_data():
    db = SessionLocal()
    try:
        if db.query(CompanyMaster).count() == 0:
            comps = [
                CompanyMaster(code="SEOUL", name="(주)서울웰푸드", biz_no="347-81-03002", rep_name="조용훈", address="서울특별시 강동구 천중로 39길 19-25(천호동,평영빌딩)", phone="02-6958-9229", fax="070-4758-9219"),
                CompanyMaster(code="TJF", name="주식회사 티제이에프", biz_no="718-88-03523", rep_name="김다희", address="경기도 남양주시 별내중앙로 34, 5층 502-A59호", phone="010-5773-0619", fax="070-4758-9219"),
                CompanyMaster(code="NEXUS", name="(주)넥서스트레이딩", biz_no="518-86-03633", rep_name="김다희", address="서울특별시 강동구 천중로 39길 19-25(천호동,평영빌딩),2층", phone="010-5773-0619", fax="070-4758-9219"),
            ]
            db.add_all(comps)

        if db.query(WarehouseRate).count() == 0:
            rates = [
                WarehouseRate(warehouse_name="강동2", frozen_rate=1.5, chilled_rate=2.5),
                WarehouseRate(warehouse_name="삼진1", frozen_rate=1.5, chilled_rate=2.5),
                WarehouseRate(warehouse_name="광주냉장", frozen_rate=1.4, chilled_rate=2.4),
                WarehouseRate(warehouse_name="부산항(예정)", frozen_rate=0.0, chilled_rate=0.0),
            ]
            db.add_all(rates)

        if db.query(User).count() == 0:
            users = [
                User(username="admin", hashed_password=hash_password("admin1234!"), full_name="조용훈 대표", role=ROLE_ADMIN),
                User(username="sales_kim", hashed_password=hash_password("sales1234!"), full_name="김영업 과장", role=ROLE_SALES),
                User(username="support_lee", hashed_password=hash_password("support1234!"), full_name="이영업지원 대리", role=ROLE_SALES_SUPPORT),
                User(username="wh_park", hashed_password=hash_password("wh1234!"), full_name="박창고 반장", role=ROLE_WAREHOUSE),
            ]
            db.add_all(users)

        # 구글 스프레드시트 마이그레이션 백데이터 로드
        if db.query(InventoryLot).count() == 0:
            lots = [
                InventoryLot(
                    company_name="(주)서울웰푸드", grid_no="GRID-771101", sku_code="BF-AMH-170", inbound_date="2026-08-10",
                    bl_no="OOLU2766442080", trace_no="802410290114", est_no="244I", brand="AMH", grade="GF",
                    item_name="우육(소)", cut_name="소일반갈비(170)", storage_type="냉동",
                    initial_box_qty=196, initial_weight_kg=4000.0, avg_box_weight=20.41, current_box_qty=196, current_weight_kg=4000.0,
                    cost_per_kg=14500, warehouse="강동2", exp_date="2028-05-31", is_weighed="N"
                ),
                InventoryLot(
                    company_name="(주)서울웰푸드", grid_no="GRID-771102", sku_code="BF-AMH-235", inbound_date="2026-08-10",
                    bl_no="OOLU2766442080", trace_no="802410290115", est_no="244I", brand="AMH", grade="GF",
                    item_name="우육(소)", cut_name="소일반갈비(235)", storage_type="냉동",
                    initial_box_qty=345, initial_weight_kg=7000.0, avg_box_weight=20.29, current_box_qty=345, current_weight_kg=7000.0,
                    cost_per_kg=14200, warehouse="강동2", exp_date="2028-05-31", is_weighed="N"
                ),
                InventoryLot(
                    company_name="(주)서울웰푸드", grid_no="GRID-771103", sku_code="PK-KC-01", inbound_date="2026-07-15",
                    bl_no="COSU6319820011", trace_no="802410290330", est_no="1614", brand="킬코이", grade="GF",
                    item_name="돈육(돼지)", cut_name="돈갈매기살", storage_type="냉동",
                    initial_box_qty=100, initial_weight_kg=2000.0, avg_box_weight=20.0, current_box_qty=100, current_weight_kg=2000.0,
                    cost_per_kg=8500, warehouse="삼진1", exp_date="2028-07-14", is_weighed="N"
                ),
                InventoryLot(
                    company_name="주식회사 티제이에프", grid_no="GRID-771104", sku_code="BF-IBP-RIB", inbound_date="2026-08-18",
                    bl_no="HDMU8820194401", trace_no="802410290552", est_no="352", brand="IBP", grade="CHOICE",
                    item_name="우육(소)", cut_name="냉장 소갈비살", storage_type="냉장",
                    initial_box_qty=80, initial_weight_kg=1600.0, avg_box_weight=20.0, current_box_qty=80, current_weight_kg=1600.0,
                    cost_per_kg=17800, warehouse="삼진1", exp_date="2026-09-04", is_weighed="N"
                ),
                InventoryLot(
                    company_name="(주)서울웰푸드", grid_no="GRID-771105", sku_code="BF-AMH-TEN", inbound_date="2026-08-28",
                    bl_no="ONEY9940128840", trace_no="802410290771", est_no="244I", brand="AMH", grade="PR",
                    item_name="우육(소)", cut_name="소안심", storage_type="선적중",
                    initial_box_qty=200, initial_weight_kg=4000.0, avg_box_weight=20.0, current_box_qty=200, current_weight_kg=4000.0,
                    cost_per_kg=19500, warehouse="부산항(예정)", exp_date="2028-08-20", is_weighed="N", status="ON_WATER"
                )
            ]
            db.add_all(lots)
            db.flush()

            # 시드 예약(HOLD) 내역
            res1 = ReservationRecord(
                res_no="RES-260824-001", grid_no="GRID-771101", sales_rep="sales_kim", sales_rep_name="김영업 과장",
                customer="(주)플랜비에프에스", box_qty=20, weight_kg=408.2, unit_price_kg=15500, total_amount=6327100,
                expire_date="2026-08-27", status="HOLD"
            )
            res2 = ReservationRecord(
                res_no="RES-260824-002", grid_no="GRID-771101", sales_rep="sales_kim", sales_rep_name="김영업 과장",
                customer="에스티에프", box_qty=15, weight_kg=306.15, unit_price_kg=15800, total_amount=4837170,
                expire_date="2026-08-27", status="HOLD"
            )
            db.add_all([res1, res2])

            # 동원홈푸드 파이낸스 담보 시드
            fin1 = PartnerFinanceContract(
                contract_no="FIN-20260810-01", grid_no="GRID-771101", partner_company="동원홈푸드",
                finance_model="IMPORT_AGENCY", pledge_date="2026-08-10", due_date="2026-09-20",
                box_qty=30, weight_kg=612.3, base_unit_price=14500, deposit_amount=2663505,
                annual_margin_rate=7.5, status="HOLD_BY_PARTNER"
            )
            db.add(fin1)

        db.commit()
    finally:
        db.close()

init_system_data()

# -----------------------------------------------------------------------------
# 6. FastAPI 라우팅 및 REST API 엔진
# -----------------------------------------------------------------------------
app = FastAPI(title="SeoulWellFood G3 ERP System")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# API 요청/응답 Pydantic 스키마
class LoginReq(BaseModel):
    username: str
    password: str

class ProformaReq(BaseModel):
    grid_no: str
    customer: str
    box_qty: int
    unit_price_kg: float
    is_hold: bool = True

class OutboundDetailReq(BaseModel):
    trade_type: str = "출고판매"
    company_name: str = "(주)서울웰푸드"
    vendor_chain: str = "브루니"
    grid_no: str
    customer: str
    box_qty: int
    weight_kg: Optional[float] = None
    unit_price_kg: float
    payment_term: str = "외상 30일 회전"
    is_estimated: str = "Y"
    dispatch_cost: float = 0.0
    cost_bearer: str = "당사부담(선불)"
    destination: str = ""
    remark: Optional[str] = ""

class TransferReq(BaseModel):
    grid_no: str
    to_warehouse: str
    box_qty: int
    dispatch_cost: float = 0.0

class FreezingConvertReq(BaseModel):
    grid_no: str
    quick_freeze_cost: float = 80.0

class ReconcileReq(BaseModel):
    outbound_no: str
    actual_weight_kg: float

class FinanceExtendReq(BaseModel):
    contract_no: str
    new_due_date: str
    adjusted_margin_rate: float
    extension_fee: float = 0.0

class DevIssueReq(BaseModel):
    issue_type: str
    title: str
    content: str
    related_menu: str
    grid_link: Optional[str] = ""

@app.post("/api/auth/login")
def login(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user = db.query(User).filter(User.username == form_data.username).first()
    if not user or not verify_password(form_data.password, user.hashed_password):
        raise HTTPException(status_code=401, detail="아이디 또는 비밀번호가 일치하지 않습니다.")
    token = create_access_token({"sub": user.username, "role": user.role, "name": user.full_name})
    return {"access_token": token, "token_type": "bearer", "role": user.role, "full_name": user.full_name, "username": user.username}

@app.get("/api/companies")
def get_companies(db: Session = Depends(get_db)):
    return db.query(CompanyMaster).all()

@app.get("/api/inventory")
def get_inventory(view_type: str = "STOCK", db: Session = Depends(get_db)):
    # 양편넣기 보관료 및 실시간 가용/예약/담보 계산
    today = date.today()
    rates = {r.warehouse_name: r.frozen_rate for r in db.query(WarehouseRate).all()}

    query = db.query(InventoryLot)
    if view_type == "ON_WATER":
        lots = query.filter(InventoryLot.status == "ON_WATER").all()
    else:
        lots = query.filter(InventoryLot.status != "ON_WATER", InventoryLot.current_box_qty > 0).order_by(InventoryLot.exp_date).all()

    # 예약 집계
    res_rows = db.query(ReservationRecord).filter(ReservationRecord.status == "HOLD").all()
    res_map = {}
    res_details = {}
    for r in res_rows:
        res_map[r.grid_no] = res_map.get(r.grid_no, 0) + r.box_qty
        res_details.setdefault(r.grid_no, []).append({
            "res_no": r.res_no,
            "created_at": r.created_at.strftime("%Y-%m-%d %H:%M"),
            "sales_rep_name": r.sales_rep_name,
            "customer": r.customer,
            "box_qty": r.box_qty,
            "weight_kg": r.weight_kg,
            "unit_price_kg": r.unit_price_kg,
            "total_amount": r.total_amount
        })

    # 담보 집계
    pledges = db.query(PartnerFinanceContract).filter(PartnerFinanceContract.status == "HOLD_BY_PARTNER").all()
    pledge_map = {}
    for p in pledges:
        pledge_map[p.grid_no] = pledge_map.get(p.grid_no, 0) + p.box_qty

    result = []
    for l in lots:
        # 양편넣기 보관일수 및 보관료 계산
        try:
            in_d = datetime.strptime(l.inbound_date, "%Y-%m-%d").date()
            storage_days = (today - in_d).days + 1 if today >= in_d else 0
        except Exception:
            storage_days = 1
        
        rate = rates.get(l.warehouse, 1.5)
        storage_fee = round(l.current_weight_kg * rate * storage_days) if storage_days > 0 else 0

        reserved = res_map.get(l.grid_no, 0)
        pledged = pledge_map.get(l.grid_no, 0)
        available = max(0, l.current_box_qty - reserved - pledged)

        # 냉장 D-10 체크
        try:
            exp_d = datetime.strptime(l.exp_date, "%Y-%m-%d").date()
            days_to_exp = (exp_d - today).days
        except Exception:
            days_to_exp = 999

        result.append({
            "grid_no": l.grid_no,
            "company_name": l.company_name,
            "bl_no": l.bl_no or "-",
            "trace_no": l.trace_no,
            "est_no": l.est_no or "-",
            "brand": l.brand or "-",
            "grade": l.grade or "GF",
            "item_name": l.item_name,
            "cut_name": l.cut_name,
            "storage_type": l.storage_type,
            "current_box_qty": l.current_box_qty,
            "current_weight_kg": l.current_weight_kg,
            "avg_box_weight": l.avg_box_weight,
            "available_box_qty": available,
            "reserved_box_qty": reserved,
            "pledged_box_qty": pledged,
            "cost_per_kg": l.cost_per_kg,
            "warehouse": l.warehouse,
            "inbound_date": l.inbound_date,
            "storage_days": storage_days,
            "storage_fee": storage_fee,
            "exp_date": l.exp_date,
            "days_to_exp": days_to_exp,
            "reservation_details": res_details.get(l.grid_no, [])
        })
    return result

@app.post("/api/proforma/create")
def create_proforma(req: ProformaReq, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    lot = db.query(InventoryLot).filter(InventoryLot.grid_no == req.grid_no).first()
    if not lot:
        raise HTTPException(status_code=404, detail="해당 재고 GRID를 찾을 수 없습니다.")
    
    avg_w = lot.avg_box_weight if lot.avg_box_weight > 0 else 20.0
    weight = round(avg_w * req.box_qty, 2)
    total = round(weight * req.unit_price_kg)
    res_no = f"RES-{datetime.now().strftime('%y%m%d')}-{random.randint(100,999)}"

    if req.is_hold:
        res = ReservationRecord(
            res_no=res_no, grid_no=lot.grid_no, sales_rep=current_user.username,
            sales_rep_name=current_user.full_name or current_user.username,
            customer=req.customer, box_qty=req.box_qty, weight_kg=weight,
            unit_price_kg=req.unit_price_kg, total_amount=total,
            expire_date=(date.today() + timedelta(days=3)).strftime("%Y-%m-%d"),
            status="HOLD"
        )
        db.add(res)
        db.commit()
    return {"message": f"가전표 저장 및 예약(HOLD {req.box_qty}Box) 완료 ({res_no})"}

@app.post("/api/outbounds/detail-create")
def create_outbound_detail(req: OutboundDetailReq, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    lot = db.query(InventoryLot).filter(InventoryLot.grid_no == req.grid_no).first()
    if not lot:
        raise HTTPException(status_code=404, detail="재고를 찾을 수 없습니다.")

    avg_w = lot.avg_box_weight if lot.avg_box_weight > 0 else 20.0
    weight = req.weight_kg if (req.weight_kg and req.weight_kg > 0) else round(avg_w * req.box_qty, 2)
    total = round(weight * req.unit_price_kg)
    out_no = f"OUT-{datetime.now().strftime('%Y%m%d%H%M%S')}-{random.randint(10,99)}"

    # 결제 만기일 산출
    today = date.today()
    if "30일" in req.payment_term:
        due_d = (today + timedelta(days=30)).strftime("%Y-%m-%d")
    elif "15일" in req.payment_term:
        due_d = (today + timedelta(days=15)).strftime("%Y-%m-%d")
    elif "7일" in req.payment_term:
        due_d = (today + timedelta(days=7)).strftime("%Y-%m-%d")
    else:
        due_d = today.strftime("%Y-%m-%d")

    # 재고 차감
    lot.current_box_qty = max(0, lot.current_box_qty - req.box_qty)
    lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - weight, 2))

    # 출고 전표 생성
    out = OutboundRecord(
        outbound_no=out_no, company_name=req.company_name, vendor_chain=req.vendor_chain,
        trade_type=req.trade_type, is_estimated=req.is_estimated, outbound_date=today.strftime("%Y-%m-%d"),
        grid_no=lot.grid_no, customer=req.customer, bl_no=lot.bl_no, trace_no=lot.trace_no,
        brand=lot.brand, grade=lot.grade, item_name=lot.item_name, cut_name=lot.cut_name,
        storage_type=lot.storage_type, box_qty=req.box_qty, weight_kg=weight,
        unit_price_kg=req.unit_price_kg, total_amount=total, payment_term=req.payment_term,
        due_date=due_d, warehouse=lot.warehouse, status="OUT_REQUEST", remark=req.remark
    )
    db.add(out)

    # 배차 연동
    if req.dispatch_cost > 0 or req.destination:
        disp_no = f"DSP-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        disp = DispatchRecord(
            dispatch_no=disp_no, dispatch_type=f"{req.trade_type} 배차", target_date=today.strftime("%Y-%m-%d"),
            partner_name=req.customer, item_name=lot.item_name, cut_name=lot.cut_name,
            box_qty=req.box_qty, weight_kg=weight, from_warehouse=lot.warehouse,
            to_destination=req.destination or "거래처 지정 배송지", dispatch_cost=req.dispatch_cost,
            cost_bearer=req.cost_bearer, status="REQUESTED"
        )
        db.add(disp)

    # 미수/채권 원장 자동 연동
    ar_no = f"AR-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    ar = AccountsReceivable(
        ar_no=ar_no, company_name=req.company_name, customer_name=req.customer,
        outbound_no=out_no, sales_date=today.strftime("%Y-%m-%d"), due_date=due_d,
        sales_amount=total, final_bill_amount=total, status="UNPAID"
    )
    db.add(ar)

    db.commit()
    return {"message": f"[{req.trade_type}] 등록 완료! (전표: {out_no}, 결제만기: {due_d})"}

@app.get("/api/outbounds")
def get_outbounds(status: str = "ALL", db: Session = Depends(get_db)):
    q = db.query(OutboundRecord)
    if status != "ALL":
        q = q.filter(OutboundRecord.status == status)
    return q.order_by(desc(OutboundRecord.id)).all()

@app.post("/api/outbounds/{outbound_no}/advance")
def advance_outbound(outbound_no: str, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    out = db.query(OutboundRecord).filter(OutboundRecord.outbound_no == outbound_no).first()
    if not out:
        raise HTTPException(status_code=404, detail="전표 없음")
    if out.status == "OUT_REQUEST":
        out.status = "OUT_CONFIRM"
        msg = "상차지시(확정) 완료"
    elif out.status == "OUT_CONFIRM":
        out.status = "OUT_DONE"
        msg = "실출고 완료"
    else:
        msg = "이미 완료된 전표입니다."
    db.commit()
    return {"message": msg}

@app.post("/api/inventory/freezing-convert")
def freezing_convert(req: FreezingConvertReq, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    lot = db.query(InventoryLot).filter(InventoryLot.grid_no == req.grid_no).first()
    if not lot or lot.storage_type != "냉장":
        raise HTTPException(status_code=400, detail="냉장 재고만 냉동전환이 가능합니다.")
    
    lot.storage_type = "❄️냉동전환"
    lot.cost_per_kg += req.quick_freeze_cost
    lot.exp_date = (date.today() + timedelta(days=730)).strftime("%Y-%m-%d")
    db.commit()
    return {"message": f"[{lot.grid_no}] 냉동전환 완료 (소비기한 24개월 확장: {lot.exp_date})"}

@app.get("/api/ar")
def get_ar_ledger(db: Session = Depends(get_db)):
    today = date.today()
    records = db.query(AccountsReceivable).order_by(AccountsReceivable.due_date).all()
    result = []
    for r in records:
        try:
            due_d = datetime.strptime(r.due_date, "%Y-%m-%d").date()
            aging = (today - due_d).days
        except Exception:
            aging = 0
        
        status = "OVERDUE" if aging > 0 and r.status != "PAID" else r.status
        result.append({
            "ar_no": r.ar_no,
            "company_name": r.company_name,
            "customer_name": r.customer_name,
            "outbound_no": r.outbound_no,
            "sales_date": r.sales_date,
            "due_date": r.due_date,
            "final_bill_amount": r.final_bill_amount,
            "paid_amount": r.paid_amount,
            "outstanding_balance": r.final_bill_amount - r.paid_amount,
            "aging_days": aging,
            "status": status
        })
    return result

@app.get("/api/dev-issues")
def get_dev_issues(db: Session = Depends(get_db)):
    return db.query(DevIssue).order_by(desc(DevIssue.id)).all()

@app.post("/api/dev-issues")
def create_dev_issue(req: DevIssueReq, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    issue = DevIssue(
        issue_type=req.issue_type, title=req.title, content=req.content,
        related_menu=req.related_menu, grid_link=req.grid_link,
        author=current_user.full_name or current_user.username
    )
    db.add(issue)
    db.commit()
    return {"message": "개선안/오류 제보가 성공적으로 등록되었습니다."}

# -----------------------------------------------------------------------------
# 7. 모바일 최적화 반응형 프론트엔드 UI (SPA)
# -----------------------------------------------------------------------------
HTML_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
  <title>SeoulWellFood G3 ERP</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    :root {
      --primary: #1e40af; --sidebar: #0f172a; --bg: #f8fafc; --border: #e2e8f0;
      --text: #0f172a; --muted: #64748b; --danger: #dc2626; --warning: #d97706; --success: #16a34a;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Pretendard", "Segoe UI", Roboto, sans-serif; }
    body { background: var(--bg); color: var(--text); display: flex; height: 100vh; overflow: hidden; }

    /* 사이드바 데스크톱 */
    aside { width: 260px; background: var(--sidebar); color: #94a3b8; display: flex; flex-direction: column; flex-shrink: 0; z-index: 50; transition: transform 0.3s ease; }
    .brand { padding: 18px 20px; font-size: 1.15rem; font-weight: 800; color: #fff; border-bottom: 1px solid #1e293b; display: flex; align-items: center; justify-content: space-between; }
    .user-card { padding: 12px 18px; background: #1e293b; color: #cbd5e1; font-size: 0.82rem; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #334155; }
    .nav-list { flex-grow: 1; overflow-y: auto; padding: 8px 0; }
    .nav-item { padding: 10px 20px; color: #94a3b8; cursor: pointer; display: flex; align-items: center; gap: 10px; font-size: 0.86rem; font-weight: 500; transition: 0.15s; }
    .nav-item:hover, .nav-item.active { background: #1e293b; color: #38bdf8; font-weight: 700; border-left: 4px solid #38bdf8; }

    /* 메인 콘텐츠 영역 */
    main { flex-grow: 1; display: flex; flex-direction: column; height: 100vh; overflow-y: auto; position: relative; }
    header { background: #fff; border-bottom: 1px solid var(--border); padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; position: sticky; top: 0; z-index: 30; }
    .mobile-header-btn { display: none; background: none; border: none; font-size: 1.4rem; color: var(--text); cursor: pointer; }
    .content { padding: 18px 24px; display: flex; flex-direction: column; gap: 14px; padding-bottom: 80px; }

    /* 카드 및 테이블 */
    .card { background: #fff; border: 1px solid var(--border); border-radius: 8px; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }
    .table-container { overflow-x: auto; max-height: calc(100vh - 220px); }
    table { width: 100%; border-collapse: collapse; font-size: 0.83rem; text-align: left; }
    th { background: #f8fafc; color: var(--muted); padding: 10px 12px; border-bottom: 1px solid var(--border); font-weight: 600; white-space: nowrap; position: sticky; top: 0; z-index: 10; }
    td { padding: 10px 12px; border-bottom: 1px solid var(--border); vertical-align: middle; white-space: nowrap; }
    tr:hover td { background: #f1f5f9; }

    /* 버튼 및 배지 */
    .btn { padding: 6px 12px; border-radius: 6px; font-size: 0.82rem; font-weight: 600; border: 1px solid transparent; cursor: pointer; display: inline-flex; align-items: center; gap: 6px; }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-outline { background: #fff; border-color: var(--border); color: #334155; }
    .badge { padding: 2px 6px; border-radius: 4px; font-size: 0.72rem; font-weight: 700; }
    .badge-blue { background: #dbeafe; color: #1e40af; }
    .badge-green { background: #dcfce7; color: #166534; }
    .badge-amber { background: #fef3c7; color: #92400e; }
    .badge-red { background: #fee2e2; color: #991b1b; }
    .grid-code { font-family: monospace; font-weight: 700; background: #312e81; color: #e0e7ff; padding: 2px 6px; border-radius: 4px; }

    /* 호버 팝오버 툴팁 */
    .res-hover { color: #0284c7; font-weight: 700; text-decoration: underline dotted; cursor: pointer; position: relative; }
    .res-tooltip {
      display: none; position: absolute; bottom: 120%; left: 50%; transform: translateX(-50%);
      background: #0f172a; color: #fff; padding: 10px 14px; border-radius: 6px; font-size: 0.75rem;
      white-space: nowrap; z-index: 100; box-shadow: 0 4px 12px rgba(0,0,0,0.25);
    }
    .res-hover:hover .res-tooltip { display: block; }

    /* 모달 워크스페이스 */
    .modal-overlay { display: none; position: fixed; inset: 0; background: rgba(15,23,42,0.6); z-index: 200; align-items: center; justify-content: center; backdrop-filter: blur(2px); }
    .modal-box { background: #fff; border-radius: 12px; width: 780px; max-width: 95vw; max-height: 90vh; overflow-y: auto; padding: 24px; }
    .form-grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
    .form-group { display: flex; flex-direction: column; gap: 4px; font-size: 0.8rem; font-weight: 600; margin-bottom: 8px; }
    .form-control { padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.85rem; outline: none; }

    /* 모바일 하단 탭바 */
    .mobile-bottom-nav { display: none; position: fixed; bottom: 0; left: 0; right: 0; height: 56px; background: #fff; border-top: 1px solid var(--border); z-index: 40; justify-content: space-around; align-items: center; }
    .mobile-nav-btn { display: flex; flex-direction: column; align-items: center; gap: 2px; font-size: 0.72rem; color: var(--muted); background: none; border: none; }
    .mobile-nav-btn.active { color: var(--primary); font-weight: 700; }

    /* 반응형 모바일 미디어 쿼리 */
    @media (max-width: 768px) {
      aside { position: fixed; left: -260px; top: 0; bottom: 0; }
      aside.open { transform: translateX(260px); }
      .mobile-header-btn { display: block; }
      .mobile-bottom-nav { display: flex; }
      .form-grid-2 { grid-template-columns: 1fr; }
    }

    /* 서식 인쇄 */
    @media print {
      body * { visibility: hidden; }
      #printArea, #printArea * { visibility: visible; }
      #printArea { position: absolute; left: 0; top: 0; width: 100%; }
    }
  </style>
</head>
<body>
  <!-- 로그인 모달 -->
  <div class="modal-overlay" id="loginModal" style="display:flex;">
    <div class="modal-box" style="width:380px;">
      <h3 style="text-align:center; margin-bottom:16px;"><i class="bi bi-box-seam-fill" style="color:var(--primary);"></i> (주)서울웰푸드 ERP</h3>
      <div class="form-group"><label>아이디</label><input type="text" id="l_user" class="form-control" value="admin" /></div>
      <div class="form-group"><label>비밀번호</label><input type="password" id="l_pw" class="form-control" value="admin1234!" /></div>
      <button class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; margin-top:10px;" onclick="handleLogin()">로그인</button>
    </div>
  </div>

  <!-- 사이드바 -->
  <aside id="appSidebar">
    <div class="brand">
      <span><i class="bi bi-grid-fill"></i> 서울웰푸드 G3</span>
      <button class="mobile-header-btn" style="color:#fff;" onclick="toggleSidebar()"><i class="bi bi-x-lg"></i></button>
    </div>
    <div class="user-card">
      <span id="userBadge"><i class="bi bi-person-circle"></i> 미인증</span>
      <button class="btn btn-outline" style="padding:2px 6px; font-size:0.7rem;" onclick="handleLogout()">로그아웃</button>
    </div>
    <div class="nav-list">
      <div class="nav-item active" onclick="navTab('STOCK', this)"><i class="bi bi-boxes"></i> 4. 현재고 및 보관관리</div>
      <div class="nav-item" onclick="navTab('PROFORMA', this)"><i class="bi bi-journal-text"></i> 5-0. 영업 가전표(예약)</div>
      <div class="nav-item" onclick="navTab('OUTBOUND', this)"><i class="bi bi-truck-flatbed"></i> 5. 출고 및 배차관리</div>
      <div class="nav-item" onclick="navTab('AR', this)"><i class="bi bi-credit-card"></i> 7. 미수 및 채권관리</div>
      <div class="nav-item" onclick="navTab('DEV', this)"><i class="bi bi-chat-left-dots"></i> 9. 개발정정 게시판</div>
    </div>
  </aside>

  <!-- 메인 뷰 -->
  <main>
    <header>
      <div style="display:flex; align-items:center; gap:12px;">
        <button class="mobile-header-btn" onclick="toggleSidebar()"><i class="bi bi-list"></i></button>
        <h2 id="viewTitle" style="font-size:1.15rem; font-weight:800;">4. 현재고 및 보관관리 (서울웰푸드 실재고)</h2>
      </div>
      <div id="headerActions" style="display:flex; gap:8px;"></div>
    </header>

    <div class="content">
      <div id="subViewFilter" style="display:flex; gap:8px;"></div>
      <div class="card">
        <div class="table-container">
          <table>
            <thead id="tblHead"></thead>
            <tbody id="tblBody"></tbody>
          </table>
        </div>
      </div>
    </div>
  </main>

  <!-- 상세 작성 모달 -->
  <div class="modal-overlay" id="workModal">
    <div class="modal-box">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px;">
        <h3 id="modalHead" style="font-size:1.05rem; font-weight:800;">상세 작성</h3>
        <button class="btn btn-outline" style="padding:2px 8px;" onclick="closeModal()">✕</button>
      </div>
      <div id="modalBody"></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal()">취소</button>
        <button class="btn btn-primary" onclick="submitModal()">확정 저장</button>
      </div>
    </div>
  </div>

  <!-- 모바일 하단 퀵바 -->
  <div class="mobile-bottom-nav">
    <button class="mobile-nav-btn active" onclick="navTab('STOCK')"><i class="bi bi-boxes"></i>재고장</button>
    <button class="mobile-nav-btn" onclick="navTab('PROFORMA')"><i class="bi bi-journal-text"></i>가전표</button>
    <button class="mobile-nav-btn" onclick="navTab('OUTBOUND')"><i class="bi bi-truck"></i>출고</button>
    <button class="mobile-nav-btn" onclick="navTab('AR')"><i class="bi bi-credit-card"></i>미수</button>
    <button class="mobile-nav-btn" onclick="navTab('DEV')"><i class="bi bi-chat-dots"></i>게시판</button>
  </div>

  <!-- 인쇄 템플릿 영역 (11줄 고정) -->
  <div id="printArea" style="display:none; padding:20px; font-family:'Malgun Gothic'; color:#000;">
    <div style="text-align:center; font-size:24px; font-weight:900; margin-bottom:16px;">출 고 / 이 체 요 청 서</div>
    <div id="printContent"></div>
  </div>

  <script>
    let authToken = localStorage.getItem('swf_token') || '';
    let curTab = 'STOCK';
    let cachedStock = [];

    async function api(url, opt = {}) {
      opt.headers = opt.headers || {};
      if (authToken) opt.headers['Authorization'] = 'Bearer ' + authToken;
      const res = await fetch(url, opt);
      if (res.status === 401) {
        document.getElementById('loginModal').style.display = 'flex';
        throw new Error('인증 필요');
      }
      return res;
    }

    async function handleLogin() {
      const u = document.getElementById('l_user').value;
      const p = document.getElementById('l_pw').value;
      const form = new URLSearchParams({ username: u, password: p });
      const r = await fetch('/api/auth/login', { method: 'POST', body: form });
      const d = await r.json();
      if (r.ok) {
        authToken = d.access_token;
        localStorage.setItem('swf_token', authToken);
        document.getElementById('userBadge').innerText = `${d.full_name} (${d.role})`;
        document.getElementById('loginModal').style.display = 'none';
        loadTab();
      } else {
        alert(d.detail || '로그인 실패');
      }
    }

    function handleLogout() {
      localStorage.removeItem('swf_token');
      location.reload();
    }

    function toggleSidebar() {
      document.getElementById('appSidebar').classList.toggle('open');
    }

    function navTab(tab, el) {
      curTab = tab;
      document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
      if (el) el.classList.add('active');
      document.getElementById('appSidebar').classList.remove('open');
      loadTab();
    }

    async function loadTab() {
      const hActions = document.getElementById('headerActions');
      const sFilter = document.getElementById('subViewFilter');
      const head = document.getElementById('tblHead');
      const body = document.getElementById('tblBody');

      if (curTab === 'STOCK') {
        document.getElementById('viewTitle').innerText = '4. 현재고 및 보관관리 (양편넣기 실시간 보관료)';
        hActions.innerHTML = `
          <button class="btn btn-outline" onclick="openModal('FREEZE')"><i class="bi bi-snow"></i> 냉동전환</button>
          <button class="btn btn-primary" onclick="openModal('OUTBOUND_DETAIL')"><i class="bi bi-cart-plus"></i> 출고요청 상세작성</button>
        `;
        sFilter.innerHTML = `
          <button class="btn btn-primary" onclick="loadStock('STOCK')">창고 실재고</button>
          <button class="btn btn-outline" onclick="loadStock('ON_WATER')">선적 판매예정재고</button>
        `;
        head.innerHTML = `
          <tr>
            <th>화주</th><th>B/L NO</th><th>이력번호</th><th>브랜드</th><th>품목/부위</th>
            <th>현재고</th><th>가용Box</th><th>예약(HOLD)</th><th>담보(Box)</th>
            <th>창고</th><th>보관료(양편넣기)</th><th>소비기한</th><th>GRID</th>
          </tr>
        `;
        loadStock('STOCK');
      } else if (curTab === 'PROFORMA') {
        document.getElementById('viewTitle').innerText = '5-0. 영업 가전표 및 예약(HOLD) 등록';
        hActions.innerHTML = `<button class="btn btn-primary" onclick="openModal('PROFORMA')"><i class="bi bi-plus-lg"></i> 신규 가전표 등록</button>`;
        sFilter.innerHTML = '';
        head.innerHTML = `<tr><th>재고 GRID</th><th>품목/부위</th><th>예약처</th><th>예약수량</th><th>제안단가</th><th>합계금액</th><th>등록자</th><th>등록일시</th></tr>`;
        const r = await api('/api/inventory');
        const list = await r.json();
        let rows = '';
        list.forEach(l => {
          (l.reservation_details || []).forEach(res => {
            rows += `<tr>
              <td><span class="grid-code">${l.grid_no}</span></td>
              <td><strong>${l.item_name}</strong> ${l.cut_name}</td>
              <td><strong>${res.customer}</strong></td>
              <td style="color:#0284c7; font-weight:700;">${res.box_qty} Box (${res.weight_kg}kg)</td>
              <td>${res.unit_price_kg.toLocaleString()}원</td>
              <td>${res.total_amount.toLocaleString()}원</td>
              <td>${res.sales_rep_name}</td>
              <td>${res.created_at}</td>
            </tr>`;
          });
        });
        body.innerHTML = rows || '<tr><td colspan="8" style="text-align:center;">예약 내역이 없습니다.</td></tr>';
      } else if (curTab === 'OUTBOUND') {
        document.getElementById('viewTitle').innerText = '5. 출고 및 배차관리 (상차지시/서식)';
        hActions.innerHTML = `<button class="btn btn-primary" onclick="printStatement()"><i class="bi bi-printer"></i> 거래명세서 11줄 인쇄</button>`;
        sFilter.innerHTML = '';
        head.innerHTML = `<tr><th>전표번호</th><th>화주</th><th>거래처</th><th>품목/부위</th><th>수량</th><th>중량(가중량)</th><th>결제조건</th><th>결제만기</th><th>상태</th><th>관리</th></tr>`;
        const r = await api('/api/outbounds?status=ALL');
        const list = await r.json();
        body.innerHTML = list.map(o => `
          <tr>
            <td><span class="grid-code">${o.outbound_no}</span></td>
            <td><span class="badge badge-blue">${o.company_name}</span></td>
            <td><strong>${o.customer}</strong></td>
            <td>${o.item_name} ${o.cut_name}</td>
            <td>${o.box_qty} Box</td>
            <td>${o.weight_kg}kg</td>
            <td>${o.payment_term}</td>
            <td style="color:#b91c1c; font-weight:700;">${o.due_date||'-'}</td>
            <td><span class="badge ${o.status==='OUT_DONE'?'badge-green':'badge-amber'}">${o.status}</span></td>
            <td>${o.status!=='OUT_DONE'?`<button class="btn btn-primary" style="padding:2px 6px;" onclick="advOut('${o.outbound_no}')">승인</button>`:'-'}</td>
          </tr>
        `).join('') || '<tr><td colspan="10" style="text-align:center;">출고 내역이 없습니다.</td></tr>';
      } else if (curTab === 'AR') {
        document.getElementById('viewTitle').innerText = '7. 미수 및 채권관리 (Aging 회전율)';
        hActions.innerHTML = '';
        sFilter.innerHTML = '';
        head.innerHTML = `<tr><th>채권번호</th><th>화주</th><th>거래처</th><th>매출일</th><th>결제만기</th><th>청구금액</th><th>수금액</th><th>미수잔액</th><th>연체(Aging)</th><th>상태</th></tr>`;
        const r = await api('/api/ar');
        const list = await r.json();
        body.innerHTML = list.map(a => `
          <tr>
            <td><span class="grid-code">${a.ar_no}</span></td>
            <td>${a.company_name}</td>
            <td><strong>${a.customer_name}</strong></td>
            <td>${a.sales_date}</td>
            <td>${a.due_date}</td>
            <td>${a.final_bill_amount.toLocaleString()}원</td>
            <td>${a.paid_amount.toLocaleString()}원</td>
            <td style="font-weight:700; color:#b91c1c;">${a.outstanding_balance.toLocaleString()}원</td>
            <td><span class="badge ${a.aging_days>0?'badge-red':'badge-green'}">${a.aging_days>0?`D+${a.aging_days}일 연체`:'정상회전'}</span></td>
            <td>${a.status}</td>
          </tr>
        `).join('') || '<tr><td colspan="10" style="text-align:center;">미수금 내역이 없습니다.</td></tr>';
      } else if (curTab === 'DEV') {
        document.getElementById('viewTitle').innerText = '9. 개발정정 및 현업 개선안 게시판';
        hActions.innerHTML = `<button class="btn btn-primary" onclick="openModal('DEV')"><i class="bi bi-pencil"></i> 신규 제보/개선안</button>`;
        sFilter.innerHTML = '';
        head.innerHTML = `<tr><th>No</th><th>구분</th><th>관련메뉴</th><th>제목</th><th>작성자</th><th>상태</th><th>등록일</th></tr>`;
        const r = await api('/api/dev-issues');
        const list = await r.json();
        body.innerHTML = list.map((d, idx) => `
          <tr>
            <td>${idx+1}</td>
            <td><span class="badge badge-blue">${d.issue_type}</span></td>
            <td>${d.related_menu}</td>
            <td><strong>${d.title}</strong></td>
            <td>${d.author}</td>
            <td><span class="badge badge-green">${d.status}</span></td>
            <td>${d.created_at.slice(0,10)}</td>
          </tr>
        `).join('') || '<tr><td colspan="7" style="text-align:center;">게시글이 없습니다.</td></tr>';
      }
    }

    async function loadStock(viewType) {
      const r = await api(`/api/inventory?view_type=${viewType}`);
      cachedStock = await r.json();
      const body = document.getElementById('tblBody');
      body.innerHTML = cachedStock.map(s => {
        const tooltips = (s.reservation_details || []).map(d =>
          `• ${d.sales_rep_name} | ${d.customer} | ${d.box_qty}Box (${d.unit_price_kg.toLocaleString()}원)`
        ).join('<br>');

        return `
          <tr>
            <td><span class="badge badge-blue">${s.company_name}</span></td>
            <td>${s.bl_no}</td>
            <td><code>${s.trace_no}</code></td>
            <td><strong>${s.brand}</strong></td>
            <td><strong>${s.item_name}</strong> ${s.cut_name}</td>
            <td><strong>${s.current_box_qty}</strong> Box (${s.current_weight_kg}kg)</td>
            <td style="color:#16a34a; font-weight:800;">${s.available_box_qty} Box</td>
            <td>
              <span class="res-hover">
                ${s.reserved_box_qty} Box
                ${tooltips ? `<div class="res-tooltip">📌 예약 상세<br>${tooltips}</div>` : ''}
              </span>
            </td>
            <td>${s.pledged_box_qty ? `<span class="badge badge-amber">${s.pledged_box_qty} Box</span>` : '-'}</td>
            <td>${s.warehouse}</td>
            <td><strong>${s.storage_fee.toLocaleString()}원</strong> (${s.storage_days}일차)</td>
            <td><span class="badge ${s.days_to_exp<=15?'badge-red':'badge-green'}">${s.exp_date}</span></td>
            <td><span class="grid-code">${s.grid_no}</span></td>
          </tr>
        `;
      }).join('') || '<tr><td colspan="13" style="text-align:center;">재고가 없습니다.</td></tr>';
    }

    function openModal(type) {
      const m = document.getElementById('workModal');
      const head = document.getElementById('modalHead');
      const body = document.getElementById('modalBody');

      if (type === 'OUTBOUND_DETAIL') {
        head.innerText = '출고요청서 상세 작성 워크스페이스';
        body.innerHTML = `
          <div class="form-grid-2">
            <div class="form-group"><label>재고 GRID 번호</label><input type="text" id="m_grid" class="form-control" value="GRID-771101" /></div>
            <div class="form-group"><label>화주 법인</label>
              <select id="m_comp" class="form-control">
                <option value="(주)서울웰푸드">(주)서울웰푸드</option>
                <option value="주식회사 티제이에프">주식회사 티제이에프</option>
                <option value="(주)넥서스트레이딩">(주)넥서스트레이딩</option>
              </select>
            </div>
            <div class="form-group"><label>거래처(출고처)</label><input type="text" id="m_cust" class="form-control" value="(주)플랜비에프에스" /></div>
            <div class="form-group"><label>매입처 표기</label><input type="text" id="m_vchain" class="form-control" value="브루니" /></div>
            <div class="form-group"><label>출고 수량(Box)</label><input type="number" id="m_box" class="form-control" value="20" /></div>
            <div class="form-group"><label>판매단가(원/kg)</label><input type="number" id="m_price" class="form-control" value="15500" /></div>
            <div class="form-group"><label>💳 결제조건 설정</label>
              <select id="m_pterm" class="form-control">
                <option value="외상 30일 회전">외상 30일 회전</option>
                <option value="외상 15일 회전">외상 15일 회전</option>
                <option value="선입금 확인후 상차">선입금 확인후 상차</option>
              </select>
            </div>
            <div class="form-group"><label>배차 운송료(원)</label><input type="number" id="m_dcost" class="form-control" value="80000" /></div>
          </div>
        `;
      } else if (type === 'PROFORMA') {
        head.innerText = '영업 가전표 등록 및 재고 예약(HOLD)';
        body.innerHTML = `
          <div class="form-group"><label>재고 GRID</label><input type="text" id="pf_grid" class="form-control" value="GRID-771101" /></div>
          <div class="form-group"><label>예약 거래처</label><input type="text" id="pf_cust" class="form-control" placeholder="예: (주)플랜비에프에스" /></div>
          <div class="form-group"><label>수량(Box)</label><input type="number" id="pf_box" class="form-control" value="10" /></div>
          <div class="form-group"><label>제안단가(원/kg)</label><input type="number" id="pf_price" class="form-control" value="15500" /></div>
        `;
      } else if (type === 'FREEZE') {
        head.innerText = '냉장육 ❄️ 급속 냉동전환 신청';
        body.innerHTML = `
          <div class="form-group"><label>전환 대상 GRID (냉장)</label><input type="text" id="fz_grid" class="form-control" value="GRID-771104" /></div>
          <div class="form-group"><label>급속 동결비용 (원/kg)</label><input type="number" id="fz_cost" class="form-control" value="80" /></div>
        `;
      } else if (type === 'DEV') {
        head.innerText = '개발정정 및 기능개선 제보';
        body.innerHTML = `
          <div class="form-group"><label>구분</label>
            <select id="dev_type" class="form-control">
              <option value="기능개선">기능개선</option>
              <option value="오류/버그">오류/버그</option>
              <option value="데이터정정">데이터정정</option>
            </select>
          </div>
          <div class="form-group"><label>제목</label><input type="text" id="dev_title" class="form-control" /></div>
          <div class="form-group"><label>상세 내용</label><textarea id="dev_content" class="form-control" rows="3"></textarea></div>
        `;
      }
      m.style.display = 'flex';
    }

    function closeModal() {
      document.getElementById('workModal').style.display = 'none';
    }

    async function submitModal() {
      const head = document.getElementById('modalHead').innerText;
      try {
        if (head.includes('출고요청서')) {
          const payload = {
            grid_no: document.getElementById('m_grid').value,
            company_name: document.getElementById('m_comp').value,
            vendor_chain: document.getElementById('m_vchain').value,
            customer: document.getElementById('m_cust').value,
            box_qty: parseInt(document.getElementById('m_box').value),
            unit_price_kg: parseFloat(document.getElementById('m_price').value),
            payment_term: document.getElementById('m_pterm').value,
            dispatch_cost: parseFloat(document.getElementById('m_dcost').value)
          };
          const res = await api('/api/outbounds/detail-create', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
          const d = await res.json();
          alert(d.message);
        } else if (head.includes('가전표')) {
          const payload = {
            grid_no: document.getElementById('pf_grid').value,
            customer: document.getElementById('pf_cust').value,
            box_qty: parseInt(document.getElementById('pf_box').value),
            unit_price_kg: parseFloat(document.getElementById('pf_price').value),
            is_hold: true
          };
          const res = await api('/api/proforma/create', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
          const d = await res.json();
          alert(d.message);
        } else if (head.includes('냉동전환')) {
          const payload = {
            grid_no: document.getElementById('fz_grid').value,
            quick_freeze_cost: parseFloat(document.getElementById('fz_cost').value)
          };
          const res = await api('/api/inventory/freezing-convert', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
          const d = await res.json();
          alert(d.message);
        } else if (head.includes('개발정정')) {
          const payload = {
            issue_type: document.getElementById('dev_type').value,
            title: document.getElementById('dev_title').value,
            content: document.getElementById('dev_content').value,
            related_menu: curTab
          };
          const res = await api('/api/dev-issues', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
          const d = await res.json();
          alert(d.message);
        }
        closeModal();
        loadTab();
      } catch(e) {
        alert('처리 중 오류 발생');
      }
    }

    async function advOut(no) {
      const res = await api(`/api/outbounds/${no}/advance`, { method: 'POST' });
      const d = await res.json();
      alert(d.message);
      loadTab();
    }

    function printStatement() {
      const pArea = document.getElementById('printArea');
      const pContent = document.getElementById('printContent');
      pContent.innerHTML = `
        <table style="width:100%; border:1px solid #000; border-collapse:collapse; font-size:12px; text-align:center;">
          <tr style="background:#f8fafc; font-weight:bold;">
            <th style="border:1px solid #000; padding:6px;">No</th>
            <th style="border:1px solid #000;">B/L NO</th>
            <th style="border:1px solid #000;">브랜드</th>
            <th style="border:1px solid #000;">품목명 / 부위명</th>
            <th style="border:1px solid #000;">수량(Box)</th>
            <th style="border:1px solid #000;">중량(kg)</th>
            <th style="border:1px solid #000;">출고처</th>
          </tr>
          <tr>
            <td style="border:1px solid #000; padding:6px;">1</td>
            <td style="border:1px solid #000;">OOLU2766442080</td>
            <td style="border:1px solid #000;">AMH</td>
            <td style="border:1px solid #000;">우육 소일반갈비(170)</td>
            <td style="border:1px solid #000;">20</td>
            <td style="border:1px solid #000;">408.2kg</td>
            <td style="border:1px solid #000;">(주)플랜비에프에스</td>
          </tr>
          ${'<tr><td style="border:1px solid #000; height:24px;"></td><td style="border:1px solid #000;"></td><td style="border:1px solid #000;"></td><td style="border:1px solid #000;"></td><td style="border:1px solid #000;"></td><td style="border:1px solid #000;"></td><td style="border:1px solid #000;"></td></tr>'.repeat(10)}
        </table>
        <div style="margin-top:20px; display:flex; justify-content:space-between; align-items:center;">
          <div>발행처: (주)서울웰푸드 (대표 조용훈)</div>
          <div style="width:60px; height:60px; border:2px solid #dc2626; border-radius:50%; color:#dc2626; display:flex; align-items:center; justify-content:center; font-size:10px; font-weight:bold;">서울웰푸드<br>인</div>
        </div>
      `;
      pArea.style.display = 'block';
      window.print();
      pArea.style.display = 'none';
    }

    window.addEventListener('DOMContentLoaded', () => {
      if (authToken) {
        document.getElementById('loginModal').style.display = 'none';
        loadTab();
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
