import io
import os
import random
from datetime import datetime, date, timedelta
from typing import List, Optional

from fastapi import FastAPI, HTTPException, Depends, UploadFile, File, Query, status
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from fastapi.responses import HTMLResponse
import openpyxl
from pydantic import BaseModel, ConfigDict
from sqlalchemy import (
    Column, Integer, Float, String, DateTime, ForeignKey, create_engine, desc, func, Index
)
from sqlalchemy.orm import declarative_base, sessionmaker, Session
import uvicorn
from passlib.context import CryptContext
from jose import JWTError, jwt

# -----------------------------------------------------------------------------
# 인증 설정
# -----------------------------------------------------------------------------
JWT_SECRET_KEY = os.environ.get("JWT_SECRET_KEY", "meatflow-enterprise-secret-key-2026")
JWT_ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 8  # 8시간

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/auth/login", auto_error=False)

ROLE_ADMIN = "ADMIN"          # 전체 권한
ROLE_SALES = "SALES"          # 영업: 출고/예약/거래처
ROLE_WAREHOUSE = "WAREHOUSE"  # 창고: 입고/재고조정/전배
ALL_ROLES = [ROLE_ADMIN, ROLE_SALES, ROLE_WAREHOUSE]

# -----------------------------------------------------------------------------
# 1. DB 설정 (SQLite 파일 기반)
# -----------------------------------------------------------------------------
DB_FILE = "meaterp_local.db"
DATABASE_URL = f"sqlite:///{DB_FILE}"

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
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
class Partner(Base):
    __tablename__ = "partners"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), nullable=False)
    type = Column(String(30), nullable=False, index=True)  # VENDOR, CUSTOMER, WAREHOUSE, SHIPPING, ETC
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
    inbound_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d"), index=True)
    processed_date = Column(String(20), nullable=True, index=True)
    vendor = Column(String(100), nullable=False)
    bl_no = Column(String(50), nullable=True)
    trace_no = Column(String(50), nullable=False, index=True)
    process_from_date = Column(String(20), nullable=True)
    brand = Column(String(50), nullable=True)
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
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

    __table_args__ = (
        Index("ix_inbound_status_date", "status", "inbound_date"),
    )

class InventoryLot(Base):
    __tablename__ = "inventory_lots"
    id = Column(Integer, primary_key=True, index=True)
    sku_code = Column(String(50), nullable=False)
    inbound_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d"))
    bl_no = Column(String(50), nullable=True)
    trace_no = Column(String(50), nullable=False, index=True)
    process_from_date = Column(String(20), nullable=True)
    brand = Column(String(50), nullable=True)
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
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
    inbound_date = Column(String(20), nullable=True)
    outbound_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d"), index=True)
    processed_date = Column(String(20), nullable=True, index=True)
    lot_id = Column(Integer, ForeignKey("inventory_lots.id"), nullable=False, index=True)
    customer = Column(String(100), nullable=False)
    bl_no = Column(String(50), nullable=True)
    trace_no = Column(String(50), nullable=False)
    process_from_date = Column(String(20), nullable=True)
    brand = Column(String(50), nullable=True)
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
    box_qty = Column(Integer, nullable=False)
    avg_box_weight = Column(Float, default=0.0)
    weight_kg = Column(Float, nullable=False)
    unit_price_kg = Column(Float, nullable=False)
    total_amount = Column(Float, nullable=False)
    exp_date = Column(String(20), nullable=False)
    warehouse = Column(String(50), nullable=False)
    is_weighed = Column(String(10), default="N")
    claim_reason = Column(String(255), nullable=True)
    status = Column(String(20), default="OUT_REQUEST", index=True)
    grid_no = Column(String(50), nullable=True, index=True)
    created_at = Column(DateTime, default=datetime.now)

    __table_args__ = (
        Index("ix_outbound_status_date", "status", "outbound_date"),
    )

class ReservationRecord(Base):
    __tablename__ = "reservations"
    id = Column(Integer, primary_key=True, index=True)
    res_no = Column(String(50), unique=True, index=True)
    lot_id = Column(Integer, ForeignKey("inventory_lots.id"), nullable=False, index=True)
    sales_rep = Column(String(50), nullable=False)
    customer = Column(String(100), nullable=False)
    grid_no = Column(String(50), nullable=False)
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    unit_price_kg = Column(Float, nullable=False)
    total_amount = Column(Float, nullable=False)
    exp_date = Column(String(20), nullable=False)
    expire_date = Column(String(20), nullable=False, index=True)
    cancel_date = Column(String(20), nullable=True)
    cancel_type = Column(String(50), nullable=True)
    cancel_reason = Column(String(255), nullable=True)
    status = Column(String(20), default="HOLD", index=True)
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
# 인증 및 헬퍼 유틸리티
# -----------------------------------------------------------------------------
def hash_password(plain: str) -> str:
    # 문자열을 바이트로 변환 후 해싱하고 다시 문자열로 반환
    salt = bcrypt.gensalt()
    hashed_bytes = bcrypt.hashpw(plain.encode('utf-8'), salt)
    return hashed_bytes.decode('utf-8')

def verify_password(plain: str, hashed: str) -> bool:
    # 입력된 비밀번호와 DB의 해시된 비밀번호를 비교
    try:
        return bcrypt.checkpw(plain.encode('utf-8'), hashed.encode('utf-8'))
    except ValueError:
        return False
def create_access_token(data: dict, expires_minutes: int = ACCESS_TOKEN_EXPIRE_MINUTES) -> str:
    to_encode = data.copy()
    expire = datetime.utcnow() + timedelta(minutes=expires_minutes)
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, JWT_SECRET_KEY, algorithm=JWT_ALGORITHM)

def get_current_user(token: Optional[str] = Depends(oauth2_scheme), db: Session = Depends(get_db)) -> "User":
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="인증이 필요합니다. 로그인 후 다시 시도해주세요.",
        headers={"WWW-Authenticate": "Bearer"},
    )
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
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=f"이 작업은 {', '.join(roles)} 권한이 필요합니다. (현재: {current_user.role})",
            )
        return current_user
    return _checker

def generate_random_grid() -> str:
    return f"GRID-{random.randint(100000, 999999)}"

def calculate_meat_exp_date(from_date_str: str, storage_type: str, item_name: str) -> str:
    try:
        base_date = datetime.strptime(from_date_str, "%Y-%m-%d").date()
    except Exception:
        base_date = date.today()

    if storage_type == "냉동":
        add_days = 730
    else:
        if "우육" in item_name or "소" in item_name:
            add_days = 90
        elif "돈육" in item_name or "돼지" in item_name:
            add_days = 60
        else:
            add_days = 30

    exp = base_date + timedelta(days=add_days - 1)
    return exp.strftime("%Y-%m-%d")

def safe_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def find_lot_by_grid(db: Session, grid_no: Optional[str]) -> Optional[InventoryLot]:
    if not grid_no:
        return None
    return db.query(InventoryLot).filter(InventoryLot.grid_no == grid_no).first()

# -----------------------------------------------------------------------------
# 테스트용 시드 데이터 초기화
# -----------------------------------------------------------------------------
def init_sample_data():
    db = SessionLocal()
    try:
        if db.query(Partner).count() == 0:
            partners = [
                Partner(name="(주)글로벌미트", type="VENDOR", biz_no="105-86-12345", contact_person="김수입", phone="010-1111-2222", address="서울시 송파구"),
                Partner(name="(주)아메리칸포크", type="VENDOR", biz_no="214-88-99123", contact_person="이영업", phone="010-3333-4444", address="서울시 강남구"),
                Partner(name="(주)하남돼지집", type="CUSTOMER", biz_no="120-81-99887", contact_person="박대표", phone="010-5555-6666", address="경기도 하남시 신장동 123"),
                Partner(name="(주)명륜진사식품", type="CUSTOMER", biz_no="131-86-54321", contact_person="정구매", phone="010-7777-8888", address="서울시 광진구"),
                Partner(name="(주)대성축산유통", type="CUSTOMER", biz_no="204-85-11223", contact_person="강실장", phone="010-9999-0000", address="인천시 서구"),
                Partner(name="광주냉장창고", type="WAREHOUSE", biz_no="110-85-44332", contact_person="최창고", phone="031-760-1234", address="경기도 광주시 초월읍"),
                Partner(name="용인냉동센터", type="WAREHOUSE", biz_no="142-88-55667", contact_person="정소장", phone="031-330-5678", address="경기도 용인시 처인구"),
                Partner(name="머스크라인(Maersk)", type="SHIPPING", biz_no="101-81-33221", contact_person="선사팀", phone="02-3700-5000", address="서울시 중구"),
                Partner(name="SGS 한국시험연구원", type="ETC", biz_no="113-81-77889", contact_person="인증실", phone="031-428-5700", address="경기도 안양시")
            ]
            db.add_all(partners)

        if db.query(ItemMaster).count() == 0:
            item_pork = ItemMaster(item_code="ITM-PORK", item_name="돈육(돼지)", species="돼지")
            item_beef = ItemMaster(item_code="ITM-BEEF", item_name="우육(소)", species="소")
            db.add_all([item_pork, item_beef])
            db.flush()

            cuts = [
                CutMaster(item_id=item_pork.id, cut_code="CUT-PORK-01", cut_name="삼겹살", default_storage="냉장"),
                CutMaster(item_id=item_pork.id, cut_code="CUT-PORK-02", cut_name="목심", default_storage="냉동"),
                CutMaster(item_id=item_pork.id, cut_code="CUT-PORK-03", cut_name="앞다리(전지)", default_storage="냉동"),
                CutMaster(item_id=item_beef.id, cut_code="CUT-BEEF-01", cut_name="척아이롤", default_storage="냉동"),
                CutMaster(item_id=item_beef.id, cut_code="CUT-BEEF-02", cut_name="부채살", default_storage="냉장"),
                CutMaster(item_id=item_beef.id, cut_code="CUT-BEEF-03", cut_name="우삼겹(업진살)", default_storage="냉동")
            ]
            db.add_all(cuts)

        if db.query(InboundRecord).count() == 0:
            inbounds = [
                InboundRecord(
                    inbound_no="IN-20260824090001", grid_no="GRID-112233", inbound_date="2026-08-24",
                    vendor="(주)글로벌미트", bl_no="ONEYVAN8812001", trace_no="802408101122",
                    process_from_date="2026-08-01", brand="올리멜", item_name="돈육(돼지)", cut_name="삼겹살",
                    storage_type="냉장", box_qty=60, weight_kg=1230.0, cost_per_kg=8500, total_amount=10455000,
                    warehouse="광주냉장창고", exp_date="2026-09-29", is_weighed="N", status="IN_REQUEST"
                ),
                InboundRecord(
                    inbound_no="IN-20260823083011", grid_no="GRID-445566", inbound_date="2026-08-23",
                    vendor="(주)아메리칸포크", bl_no="MAEU772391002", trace_no="100293819201",
                    process_from_date="2026-07-15", brand="스위프트", item_name="돈육(돼지)", cut_name="목심",
                    storage_type="냉동", box_qty=100, weight_kg=2050.0, cost_per_kg=6800, total_amount=13940000,
                    warehouse="용인냉동센터", exp_date="2028-07-14", is_weighed="N", status="IN_CONFIRM"
                ),
                InboundRecord(
                    inbound_no="IN-20260810091000", grid_no="GRID-738201", inbound_date="2026-08-10",
                    processed_date="2026-08-10", vendor="(주)글로벌미트", bl_no="ONEYVAN6948200", trace_no="802410290114",
                    process_from_date="2026-07-20", brand="올리멜", item_name="돈육(돼지)", cut_name="삼겹살",
                    storage_type="냉장", box_qty=50, weight_kg=1020.5, cost_per_kg=8400, total_amount=8572200,
                    warehouse="광주냉장창고", exp_date="2026-09-17", is_weighed="N", status="IN_DONE"
                ),
                InboundRecord(
                    inbound_no="IN-20260812102000", grid_no="GRID-910482", inbound_date="2026-08-12",
                    processed_date="2026-08-12", vendor="(주)아메리칸포크", bl_no="MAEU9201948201", trace_no="100392019482",
                    process_from_date="2026-06-01", brand="엑셀", item_name="우육(소)", cut_name="척아이롤",
                    storage_type="냉동", box_qty=150, weight_kg=3050.0, cost_per_kg=14500, total_amount=44225000,
                    warehouse="용인냉동센터", exp_date="2028-05-30", is_weighed="N", status="IN_DONE"
                ),
                InboundRecord(
                    inbound_no="IN-20260818140000", grid_no="GRID-554433", inbound_date="2026-08-18",
                    processed_date="2026-08-19", vendor="(주)글로벌미트", bl_no="ONEYVAN3344112", trace_no="802407229911",
                    process_from_date="2026-07-01", brand="메이플", item_name="돈육(돼지)", cut_name="삼겹살",
                    storage_type="냉장", box_qty=20, weight_kg=410.0, cost_per_kg=8300, total_amount=3403000,
                    warehouse="광주냉장창고", exp_date="2026-08-29", is_weighed="N", status="IN_CLAIM",
                    claim_reason="창고 입고 검수 시 진공 풀림 및 갈변 확인되어 전량 반품 클레임"
                )
            ]
            db.add_all(inbounds)

        if db.query(InventoryLot).count() == 0:
            lots = [
                InventoryLot(
                    grid_no="GRID-738201", sku_code="PK-CAN-01", inbound_date="2026-08-10",
                    bl_no="ONEYVAN6948200", trace_no="802410290114", process_from_date="2026-07-20",
                    brand="올리멜", item_name="돈육(돼지)", cut_name="삼겹살", storage_type="냉장",
                    initial_box_qty=50, initial_weight_kg=1020.5, avg_box_weight=20.41,
                    current_box_qty=30, current_weight_kg=612.3, cost_per_kg=8400,
                    warehouse="광주냉장창고", exp_date="2026-09-17", is_weighed="N"
                ),
                InventoryLot(
                    grid_no="GRID-910482", sku_code="BF-USA-05", inbound_date="2026-08-12",
                    bl_no="MAEU9201948201", trace_no="100392019482", process_from_date="2026-06-01",
                    brand="엑셀", item_name="우육(소)", cut_name="척아이롤", storage_type="냉동",
                    initial_box_qty=150, initial_weight_kg=3050.0, avg_box_weight=20.33,
                    current_box_qty=110, current_weight_kg=2236.3, cost_per_kg=14500,
                    warehouse="용인냉동센터", exp_date="2028-05-30", is_weighed="N"
                ),
                InventoryLot(
                    grid_no="GRID-339912", sku_code="PK-USA-02", inbound_date="2026-08-15",
                    bl_no="CMACGM1029384", trace_no="100492819283", process_from_date="2026-05-10",
                    brand="스미스필드", item_name="돈육(돼지)", cut_name="목심", storage_type="냉동",
                    initial_box_qty=80, initial_weight_kg=1640.0, avg_box_weight=20.50,
                    current_box_qty=80, current_weight_kg=1640.0, cost_per_kg=6900,
                    warehouse="용인냉동센터", exp_date="2028-05-09", is_weighed="N"
                ),
                InventoryLot(
                    grid_no="GRID-883311", sku_code="BF-USA-09", inbound_date="2026-08-01",
                    bl_no="APLU883920192", trace_no="100994820192", process_from_date="2026-06-10",
                    brand="IBP", item_name="우육(소)", cut_name="부채살", storage_type="냉장",
                    initial_box_qty=40, initial_weight_kg=812.0, avg_box_weight=20.30,
                    current_box_qty=25, current_weight_kg=507.5, cost_per_kg=18500,
                    warehouse="광주냉장창고", exp_date="2026-09-07", is_weighed="N"
                )
            ]
            db.add_all(lots)
            db.flush()

        if db.query(OutboundRecord).count() == 0:
            outbounds = [
                OutboundRecord(
                    outbound_no="OUT-20260824093001", inbound_date="2026-08-10", outbound_date="2026-08-24",
                    lot_id=1, customer="(주)하남돼지집", bl_no="ONEYVAN6948200", trace_no="802410290114",
                    process_from_date="2026-07-20", brand="올리멜", item_name="돈육(돼지)", cut_name="삼겹살",
                    storage_type="냉장", box_qty=10, avg_box_weight=20.41, weight_kg=204.1,
                    unit_price_kg=10500, total_amount=2143050, exp_date="2026-09-17",
                    warehouse="광주냉장창고", is_weighed="N", status="OUT_REQUEST", grid_no="GRID-738201"
                ),
                OutboundRecord(
                    outbound_no="OUT-20260824091500", inbound_date="2026-08-12", outbound_date="2026-08-24",
                    lot_id=2, customer="(주)명륜진사식품", bl_no="MAEU9201948201", trace_no="100392019482",
                    process_from_date="2026-06-01", brand="엑셀", item_name="우육(소)", cut_name="척아이롤",
                    storage_type="냉동", box_qty=20, avg_box_weight=20.33, weight_kg=406.6,
                    unit_price_kg=16800, total_amount=6830880, exp_date="2028-05-30",
                    warehouse="용인냉동센터", is_weighed="N", status="OUT_CONFIRM", grid_no="GRID-910482"
                ),
                OutboundRecord(
                    outbound_no="OUT-20260820110000", inbound_date="2026-08-10", outbound_date="2026-08-20",
                    processed_date="2026-08-20", lot_id=1, customer="(주)대성축산유통", bl_no="ONEYVAN6948200",
                    trace_no="802410290114", process_from_date="2026-07-20", brand="올리멜", item_name="돈육(돼지)",
                    cut_name="삼겹살", storage_type="냉장", box_qty=10, avg_box_weight=20.41, weight_kg=204.1,
                    unit_price_kg=10400, total_amount=2122640, exp_date="2026-09-17",
                    warehouse="광주냉장창고", is_weighed="N", status="OUT_DONE", grid_no="GRID-738201"
                ),
                OutboundRecord(
                    outbound_no="OUT-20260821153000", inbound_date="2026-08-12", outbound_date="2026-08-21",
                    processed_date="2026-08-22", lot_id=2, customer="(주)명륜진사식품", bl_no="MAEU9201948201",
                    trace_no="100392019482", process_from_date="2026-06-01", brand="엑셀", item_name="우육(소)",
                    cut_name="척아이롤", storage_type="냉동", box_qty=5, avg_box_weight=20.33, weight_kg=101.65,
                    unit_price_kg=16800, total_amount=1707720, exp_date="2028-05-30",
                    warehouse="용인냉동센터", is_weighed="N", status="OUT_CLAIM", grid_no="GRID-910482",
                    claim_reason="하차 후 실계근 시 1Box 중량 과소(약 3kg 감량)로 인한 부분 반품"
                )
            ]
            db.add_all(outbounds)

        if db.query(ReservationRecord).count() == 0:
            reservations = [
                ReservationRecord(
                    res_no="RES-20260824094000", lot_id=1, sales_rep="김영업", customer="(주)하남돼지집",
                    grid_no="GRID-738201", item_name="돈육(돼지)", cut_name="삼겹살", storage_type="냉장",
                    box_qty=5, weight_kg=102.05, unit_price_kg=10800, total_amount=1102140,
                    exp_date="2026-09-17", expire_date="2026-08-30", status="HOLD"
                ),
                ReservationRecord(
                    res_no="RES-20260824095000", lot_id=2, sales_rep="이영업", customer="(주)대성축산유통",
                    grid_no="GRID-910482", item_name="우육(소)", cut_name="척아이롤", storage_type="냉동",
                    box_qty=15, weight_kg=304.95, unit_price_kg=16500, total_amount=5031675,
                    exp_date="2028-05-30", expire_date="2026-08-31", status="HOLD"
                ),
                ReservationRecord(
                    res_no="RES-20260815100000", lot_id=1, sales_rep="김영업", customer="(주)명륜진사식품",
                    grid_no="GRID-738201", item_name="돈육(돼지)", cut_name="삼겹살", storage_type="냉장",
                    box_qty=10, weight_kg=204.1, unit_price_kg=10400, total_amount=2122640,
                    exp_date="2026-09-17", expire_date="2026-08-20", cancel_date="2026-08-21",
                    cancel_type="기간만료 자동취소", cancel_reason="예약 유효 만료일 경과 자동 해제", status="CANCELLED"
                )
            ]
            db.add_all(reservations)

        if db.query(User).count() == 0:
            users = [
                User(username="admin", hashed_password=hash_password("admin1234!"), full_name="시스템관리자", role=ROLE_ADMIN),
                User(username="sales1", hashed_password=hash_password("sales1234!"), full_name="김영업", role=ROLE_SALES),
                User(username="wh1", hashed_password=hash_password("wh1234!"), full_name="최창고", role=ROLE_WAREHOUSE),
            ]
            db.add_all(users)

        db.commit()
    finally:
        db.close()

init_sample_data()

# -----------------------------------------------------------------------------
# 3. FastAPI 앱 & 스키마
# -----------------------------------------------------------------------------
app = FastAPI(title="MeatFlow Enterprise ERP")

class PartnerOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    name: str
    type: str
    biz_no: Optional[str] = None
    contact_person: Optional[str] = None
    phone: Optional[str] = None
    address: Optional[str] = None

class ItemMasterOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    item_code: str
    item_name: str
    species: str

class CutMasterOut(BaseModel):
    id: int
    item_id: int
    cut_code: str
    cut_name: str
    default_storage: Optional[str] = None
    parent_item_name: str
    species: str

class InboundOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    inbound_no: str
    inbound_date: str
    processed_date: Optional[str] = None
    vendor: str
    bl_no: Optional[str] = None
    trace_no: str
    process_from_date: Optional[str] = None
    brand: Optional[str] = None
    item_name: str
    cut_name: str
    storage_type: str
    box_qty: int
    weight_kg: float
    cost_per_kg: float
    total_amount: float
    warehouse: str
    exp_date: str
    is_weighed: str
    claim_reason: Optional[str] = None
    status: str
    grid_no: Optional[str] = None

class InventoryLotOut(BaseModel):
    id: int
    grid_no: str
    sku_code: str
    inbound_date: str
    bl_no: Optional[str] = None
    trace_no: str
    process_from_date: Optional[str] = None
    brand: Optional[str] = None
    item_name: str
    cut_name: str
    storage_type: str
    initial_box_qty: int
    initial_weight_kg: float
    avg_box_weight: float
    current_box_qty: int
    current_weight_kg: float
    cost_per_kg: float
    warehouse: str
    exp_date: str
    reserved_box_qty: int
    reserved_weight_kg: float
    reserved_customers: str

class OutboundOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    outbound_no: str
    inbound_date: Optional[str] = None
    outbound_date: str
    processed_date: Optional[str] = None
    lot_id: int
    customer: str
    bl_no: Optional[str] = None
    trace_no: str
    process_from_date: Optional[str] = None
    brand: Optional[str] = None
    item_name: str
    cut_name: str
    storage_type: str
    box_qty: int
    avg_box_weight: float
    weight_kg: float
    unit_price_kg: float
    total_amount: float
    exp_date: str
    warehouse: str
    is_weighed: str
    claim_reason: Optional[str] = None
    status: str
    grid_no: Optional[str] = None

class ReservationOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    res_no: str
    lot_id: int
    sales_rep: str
    customer: str
    grid_no: str
    item_name: str
    cut_name: str
    storage_type: str
    box_qty: int
    weight_kg: float
    unit_price_kg: float
    total_amount: float
    exp_date: str
    expire_date: str
    cancel_date: Optional[str] = None
    cancel_type: Optional[str] = None
    cancel_reason: Optional[str] = None
    status: str

class ClaimOut(BaseModel):
    id: int
    stage: str
    inbound_date: Optional[str] = None
    processed_date: Optional[str] = None
    doc_no: str
    partner_name: str
    warehouse: str
    bl_no: Optional[str] = None
    trace_no: str
    process_from_date: Optional[str] = None
    brand: Optional[str] = None
    item_name: str
    cut_name: str
    box_qty: int
    weight_kg: float
    total_amount: float
    claim_reason: Optional[str] = None
    grid_no: Optional[str] = None
    exp_date: str
    raw_type: str

class StockAdjustmentOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    lot_id: int
    adj_type: str
    adj_box: int
    adj_weight: float
    reason: Optional[str] = None
    adjusted_at: datetime

class MessageOut(BaseModel):
    message: str

class UserOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    username: str
    full_name: Optional[str] = None
    role: str
    is_active: int

class UserCreate(BaseModel):
    username: str
    password: str
    full_name: Optional[str] = ""
    role: str = ROLE_SALES

class TokenOut(BaseModel):
    access_token: str
    token_type: str = "bearer"
    role: str
    username: str
    full_name: Optional[str] = None

class PartnerCreate(BaseModel):
    name: str
    type: str
    biz_no: Optional[str] = ""
    contact_person: Optional[str] = ""
    phone: Optional[str] = ""
    address: Optional[str] = ""

class ItemMasterCreate(BaseModel):
    item_code: str
    item_name: str
    species: str

class CutMasterCreate(BaseModel):
    item_id: int
    cut_code: str
    cut_name: str
    default_storage: Optional[str] = "냉장"

class InboundCreate(BaseModel):
    inbound_date: Optional[str] = None
    vendor: str
    bl_no: Optional[str] = ""
    trace_no: str
    process_from_date: Optional[str] = None
    brand: Optional[str] = ""
    item_name: str
    cut_name: str
    storage_type: str
    box_qty: int
    weight_kg: float
    cost_per_kg: float
    warehouse: str
    exp_date: str
    is_weighed: Optional[str] = "N"

class UpdateInbound(BaseModel):
    inbound_date: str
    bl_no: Optional[str] = ""
    process_from_date: Optional[str] = None
    brand: Optional[str] = ""
    item_name: str
    cut_name: str
    box_qty: int
    weight_kg: float
    cost_per_kg: float
    warehouse: str
    is_weighed: str
    exp_date: str

class ClaimRegister(BaseModel):
    reason: str
    processed_date: Optional[str] = None

class OutboundCreate(BaseModel):
    outbound_date: Optional[str] = None
    lot_id: int
    customer: str
    box_qty: int
    unit_price_kg: float

class UpdateOutbound(BaseModel):
    outbound_date: str
    box_qty: int
    weight_kg: float
    unit_price_kg: float

class ReservationCreate(BaseModel):
    lot_id: int
    sales_rep: str
    customer: str
    box_qty: int
    unit_price_kg: float
    expire_date: str

class ReservationUpdate(BaseModel):
    box_qty: int
    unit_price_kg: float
    expire_date: str

class ReservationCancelReq(BaseModel):
    cancel_reason: str
    cancel_type: Optional[str] = "수동 요청취소"

class AdjustCreate(BaseModel):
    lot_id: int
    adj_type: str
    adj_box: int
    adj_weight: float
    reason: Optional[str] = ""

class WarehouseTransferCreate(BaseModel):
    lot_id: int
    to_warehouse: str
    box_qty: int
    reason: Optional[str] = ""

# --- [인증 API] ---
@app.post("/api/auth/login", response_model=TokenOut)
def login(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user = db.query(User).filter(User.username == form_data.username).first()
    if not user or not user.is_active or not verify_password(form_data.password, user.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="아이디 또는 비밀번호가 올바르지 않습니다.",
            headers={"WWW-Authenticate": "Bearer"},
        )
    token = create_access_token({"sub": user.username, "role": user.role})
    return {
        "access_token": token, "token_type": "bearer",
        "role": user.role, "username": user.username, "full_name": user.full_name,
    }

@app.get("/api/auth/me", response_model=UserOut)
def get_me(current_user: User = Depends(get_current_user)):
    return current_user

@app.post("/api/auth/users", response_model=MessageOut)
def create_user(req: UserCreate, current_user: User = Depends(require_roles(ROLE_ADMIN)), db: Session = Depends(get_db)):
    if req.role not in ALL_ROLES:
        raise HTTPException(status_code=400, detail=f"role은 {ALL_ROLES} 중 하나여야 합니다.")
    if db.query(User).filter(User.username == req.username).first():
        raise HTTPException(status_code=400, detail="이미 존재하는 아이디입니다.")
    user = User(username=req.username, hashed_password=hash_password(req.password), full_name=req.full_name, role=req.role)
    db.add(user)
    db.commit()
    return {"message": f"계정 '{req.username}'이(가) {req.role} 권한으로 생성되었습니다."}

@app.get("/api/auth/users", response_model=List[UserOut])
def list_users(current_user: User = Depends(require_roles(ROLE_ADMIN)), db: Session = Depends(get_db)):
    return db.query(User).order_by(User.id).all()

# --- [거래처 API] ---
@app.get("/api/partners", response_model=List[PartnerOut])
def get_partners(type: Optional[str] = None, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    q = db.query(Partner)
    if type and type != "ALL":
        q = q.filter(Partner.type == type)
    return q.order_by(desc(Partner.id)).all()

@app.post("/api/partners", response_model=MessageOut)
def create_partner(req: PartnerCreate, current_user: User = Depends(require_roles(ROLE_ADMIN, ROLE_SALES)), db: Session = Depends(get_db)):
    p = Partner(name=req.name, type=req.type, biz_no=req.biz_no, contact_person=req.contact_person, phone=req.phone, address=req.address)
    db.add(p)
    db.commit()
    return {"message": "거래처 정보가 등록되었습니다."}

@app.delete("/api/partners/{partner_id}", response_model=MessageOut)
def delete_partner(partner_id: int, current_user: User = Depends(require_roles(ROLE_ADMIN)), db: Session = Depends(get_db)):
    p = db.query(Partner).filter(Partner.id == partner_id).first()
    if not p:
        raise HTTPException(status_code=404, detail="거래처를 찾을 수 없습니다.")
    db.delete(p)
    db.commit()
    return {"message": "거래처가 삭제되었습니다."}

# --- [품목/부위 마스터 API] ---
@app.get("/api/items", response_model=List[ItemMasterOut])
def get_item_masters(current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    return db.query(ItemMaster).order_by(ItemMaster.id).all()

@app.post("/api/items", response_model=MessageOut)
def create_item_master(req: ItemMasterCreate, current_user: User = Depends(require_roles(ROLE_ADMIN)), db: Session = Depends(get_db)):
    if db.query(ItemMaster).filter(ItemMaster.item_code == req.item_code).first():
        raise HTTPException(status_code=400, detail="이미 등록된 품목코드입니다.")
    item = ItemMaster(item_code=req.item_code, item_name=req.item_name, species=req.species)
    db.add(item)
    db.commit()
    return {"message": "상위 품목 마스터가 등록되었습니다."}

@app.delete("/api/items/{item_id}", response_model=MessageOut)
def delete_item_master(item_id: int, current_user: User = Depends(require_roles(ROLE_ADMIN)), db: Session = Depends(get_db)):
    item = db.query(ItemMaster).filter(ItemMaster.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="품목을 찾을 수 없습니다.")
    db.query(CutMaster).filter(CutMaster.item_id == item_id).delete()
    db.delete(item)
    db.commit()
    return {"message": "품목 및 소속 부위가 삭제되었습니다."}

@app.get("/api/cuts", response_model=List[CutMasterOut])
def get_cut_masters(item_id: Optional[int] = None, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    q = db.query(
        CutMaster.id, CutMaster.item_id, CutMaster.cut_code, CutMaster.cut_name,
        CutMaster.default_storage, ItemMaster.item_name, ItemMaster.species
    ).join(ItemMaster, CutMaster.item_id == ItemMaster.id)
    if item_id:
        q = q.filter(CutMaster.item_id == item_id)
    results = q.order_by(CutMaster.item_id, CutMaster.id).all()
    return [
        {
            "id": r.id, "item_id": r.item_id, "cut_code": r.cut_code, "cut_name": r.cut_name,
            "default_storage": r.default_storage, "parent_item_name": r.item_name, "species": r.species
        } for r in results
    ]

@app.post("/api/cuts", response_model=MessageOut)
def create_cut_master(req: CutMasterCreate, current_user: User = Depends(require_roles(ROLE_ADMIN)), db: Session = Depends(get_db)):
    if db.query(CutMaster).filter(CutMaster.cut_code == req.cut_code).first():
        raise HTTPException(status_code=400, detail="이미 등록된 부위코드입니다.")
    cut = CutMaster(item_id=req.item_id, cut_code=req.cut_code, cut_name=req.cut_name, default_storage=req.default_storage)
    db.add(cut)
    db.commit()
    return {"message": "세부 부위 마스터가 등록되었습니다."}

@app.delete("/api/cuts/{cut_id}", response_model=MessageOut)
def delete_cut_master(cut_id: int, current_user: User = Depends(require_roles(ROLE_ADMIN)), db: Session = Depends(get_db)):
    cut = db.query(CutMaster).filter(CutMaster.id == cut_id).first()
    if not cut:
        raise HTTPException(status_code=404, detail="부위를 찾을 수 없습니다.")
    db.delete(cut)
    db.commit()
    return {"message": "부위 마스터가 삭제되었습니다."}

# --- [입고 관리 API] ---
@app.get("/api/inbounds", response_model=List[InboundOut])
def get_inbounds(
    status: str,
    target_date: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    month: Optional[str] = None,
    limit: int = Query(default=500, ge=1, le=2000),
    offset: int = Query(default=0, ge=0),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    q = db.query(InboundRecord).filter(InboundRecord.status == status)
    if target_date:
        q = q.filter(InboundRecord.inbound_date == target_date)
    elif month:
        q = q.filter(InboundRecord.inbound_date.like(f"{month}%"))
    else:
        if start_date:
            q = q.filter(InboundRecord.inbound_date >= start_date)
        if end_date:
            q = q.filter(InboundRecord.inbound_date <= end_date)
    return q.order_by(desc(InboundRecord.inbound_date), desc(InboundRecord.id)).offset(offset).limit(limit).all()

@app.post("/api/inbounds", response_model=MessageOut)
def create_inbound(req: InboundCreate, current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)), db: Session = Depends(get_db)):
    inbound_no = f"IN-{datetime.now().strftime('%Y%m%d%H%M%S')}-{random.randint(10, 99)}"
    grid_no = generate_random_grid()
    in_date = req.inbound_date if req.inbound_date else datetime.now().strftime("%Y-%m-%d")
    total = round(req.weight_kg * req.cost_per_kg)
    item = InboundRecord(
        inbound_no=inbound_no, grid_no=grid_no, inbound_date=in_date, vendor=req.vendor, bl_no=req.bl_no,
        trace_no=req.trace_no, process_from_date=req.process_from_date, brand=req.brand,
        item_name=req.item_name, cut_name=req.cut_name, storage_type=req.storage_type,
        box_qty=req.box_qty, weight_kg=req.weight_kg, cost_per_kg=req.cost_per_kg,
        total_amount=total, warehouse=req.warehouse, exp_date=req.exp_date,
        is_weighed=req.is_weighed, status="IN_REQUEST"
    )
    db.add(item)
    db.commit()
    return {"message": f"입고등록 완료 ({grid_no})"}

@app.post("/api/inbounds/upload-excel")
async def upload_inbound_excel(
    file: UploadFile = File(...),
    current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    try:
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=True)
        sheet = wb.active
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀 파일 해석 오류: {str(e)}")

    rows = list(sheet.iter_rows(values_only=True))
    if len(rows) < 2:
        raise HTTPException(status_code=400, detail="데이터가 비어있는 엑셀 파일입니다.")

    created_count = 0
    now_str = datetime.now().strftime("%Y-%m-%d")
    failed_rows = []

    for r_idx, row in enumerate(rows[1:], start=2):
        if not row or all(v is None for v in row):
            continue
        try:
            in_date = safe_str(row[0])[:10] if row[0] else now_str
            vendor = safe_str(row[1]) if len(row) > 1 and row[1] else "일괄매입처"
            bl_no = safe_str(row[2]) if len(row) > 2 and row[2] else ""
            trace_no = safe_str(row[3]) if len(row) > 3 and row[3] else f"EXCEL-{random.randint(100000,999999)}"
            process_from = safe_str(row[4])[:10] if len(row) > 4 and row[4] else in_date
            brand = safe_str(row[5]) if len(row) > 5 and row[5] else ""
            item_name = safe_str(row[6]) if len(row) > 6 and row[6] else "돈육(돼지)"
            cut_name = safe_str(row[7]) if len(row) > 7 and row[7] else "삼겹살"
            storage_type = safe_str(row[8]) if len(row) > 8 and row[8] else "냉장"
            box_qty = int(row[9]) if len(row) > 9 and row[9] else 1
            weight_kg = float(row[10]) if len(row) > 10 and row[10] else 20.0
            cost_per_kg = float(row[11]) if len(row) > 11 and row[11] else 8500.0
            warehouse = safe_str(row[12]) if len(row) > 12 and row[12] else "광주냉장창고"

            if box_qty <= 0 or weight_kg <= 0 or cost_per_kg <= 0:
                raise ValueError("박스수량/중량/단가는 0보다 커야 합니다.")

            exp_date = calculate_meat_exp_date(process_from, storage_type, item_name)
            grid_no = generate_random_grid()
            inbound_no = f"IN-XLS-{datetime.now().strftime('%Y%m%d%H%M%S')}-{r_idx}"

            inbound_rec = InboundRecord(
                inbound_no=inbound_no, grid_no=grid_no, inbound_date=in_date, vendor=vendor,
                bl_no=bl_no, trace_no=trace_no, process_from_date=process_from, brand=brand,
                item_name=item_name, cut_name=cut_name, storage_type=storage_type,
                box_qty=box_qty, weight_kg=weight_kg, cost_per_kg=cost_per_kg,
                total_amount=round(weight_kg * cost_per_kg), warehouse=warehouse,
                exp_date=exp_date, is_weighed="N", status="IN_REQUEST"
            )
            db.add(inbound_rec)
            created_count += 1
        except Exception as ex:
            failed_rows.append({"row": r_idx, "error": str(ex)})
            continue

    db.commit()
    return {
        "message": f"총 {created_count}건의 데이터가 일괄 입고 등록되었습니다."
        + (f" ({len(failed_rows)}건 실패)" if failed_rows else ""),
        "created_count": created_count,
        "failed_rows": failed_rows,
    }

@app.post("/api/inbounds/{inbound_id}/claim", response_model=MessageOut)
def register_inbound_claim(
    inbound_id: int, req: ClaimRegister,
    current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).with_for_update().first()
    if not inbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    if inbound.status == "IN_DONE":
        lot = find_lot_by_grid(db, inbound.grid_no)
        if lot:
            lot.current_box_qty = max(0, lot.current_box_qty - inbound.box_qty)
            lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - inbound.weight_kg, 2))

    inbound.status = "IN_CLAIM"
    inbound.claim_reason = req.reason
    inbound.processed_date = req.processed_date if req.processed_date else datetime.now().strftime("%Y-%m-%d")
    db.commit()
    return {"message": f"입고 클레임 등록이 완료되었습니다. (처리일자: {inbound.processed_date})"}

@app.put("/api/inbounds/{inbound_id}", response_model=MessageOut)
def update_inbound(
    inbound_id: int, req: UpdateInbound,
    current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).first()
    if not inbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    if inbound.status in ["IN_DONE", "IN_CLAIM"]:
        raise HTTPException(status_code=400, detail="완료 또는 클레임 상태의 전표는 [작업취소] 후 수정해야 합니다.")
    inbound.inbound_date = req.inbound_date
    inbound.bl_no = req.bl_no
    inbound.process_from_date = req.process_from_date
    inbound.brand = req.brand
    inbound.item_name = req.item_name
    inbound.cut_name = req.cut_name
    inbound.box_qty = req.box_qty
    inbound.weight_kg = req.weight_kg
    inbound.cost_per_kg = req.cost_per_kg
    inbound.warehouse = req.warehouse
    inbound.is_weighed = req.is_weighed
    inbound.exp_date = req.exp_date
    inbound.total_amount = round(req.weight_kg * req.cost_per_kg)
    db.commit()
    return {"message": "입고 정보가 수정되었습니다."}

@app.post("/api/inbounds/{inbound_id}/advance", response_model=MessageOut)
def advance_inbound(
    inbound_id: int,
    current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).with_for_update().first()
    if not inbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    if inbound.status == "IN_REQUEST":
        inbound.status = "IN_CONFIRM"
        msg = "입고확정 단계로 이동되었습니다."
    elif inbound.status == "IN_CONFIRM":
        inbound.status = "IN_DONE"
        inbound.processed_date = datetime.now().strftime("%Y-%m-%d")
        avg_w = round(inbound.weight_kg / inbound.box_qty, 2) if inbound.box_qty > 0 else 0.0

        lot = find_lot_by_grid(db, inbound.grid_no)
        if lot:
            lot.initial_box_qty += inbound.box_qty
            lot.initial_weight_kg = round(lot.initial_weight_kg + inbound.weight_kg, 2)
            lot.avg_box_weight = round(lot.initial_weight_kg / lot.initial_box_qty, 2) if lot.initial_box_qty > 0 else avg_w
            lot.current_box_qty += inbound.box_qty
            lot.current_weight_kg = round(lot.current_weight_kg + inbound.weight_kg, 2)
            lot.cost_per_kg = inbound.cost_per_kg
            lot.is_weighed = inbound.is_weighed
        else:
            sku = "SKU-" + (inbound.trace_no[:6] if len(inbound.trace_no) >= 6 else inbound.trace_no)
            new_lot = InventoryLot(
                grid_no=inbound.grid_no or generate_random_grid(),
                sku_code=sku, inbound_date=inbound.inbound_date, bl_no=inbound.bl_no,
                trace_no=inbound.trace_no, process_from_date=inbound.process_from_date,
                brand=inbound.brand, item_name=inbound.item_name, cut_name=inbound.cut_name,
                storage_type=inbound.storage_type, initial_box_qty=inbound.box_qty,
                initial_weight_kg=inbound.weight_kg, avg_box_weight=avg_w,
                current_box_qty=inbound.box_qty, current_weight_kg=inbound.weight_kg,
                cost_per_kg=inbound.cost_per_kg, warehouse=inbound.warehouse,
                exp_date=inbound.exp_date, is_weighed=inbound.is_weighed
            )
            db.add(new_lot)
        msg = "입고완료 처리 및 실재고 반영이 완료되었습니다."
    else:
        msg = "이미 완료된 전표입니다."
    db.commit()
    return {"message": msg}

@app.post("/api/inbounds/{inbound_id}/revert", response_model=MessageOut)
def revert_inbound(
    inbound_id: int,
    current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).with_for_update().first()
    if not inbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    if inbound.status == "IN_REQUEST":
        db.delete(inbound)
        msg = "입고요청 전표가 삭제되었습니다."
    elif inbound.status == "IN_CONFIRM":
        inbound.status = "IN_REQUEST"
        msg = "입고확정이 취소되어 [입고요청리스트]로 반려되었습니다."
    elif inbound.status == "IN_DONE":
        lot = find_lot_by_grid(db, inbound.grid_no)
        if lot:
            lot.initial_box_qty = max(0, lot.initial_box_qty - inbound.box_qty)
            lot.initial_weight_kg = max(0.0, round(lot.initial_weight_kg - inbound.weight_kg, 2))
            lot.current_box_qty = max(0, lot.current_box_qty - inbound.box_qty)
            lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - inbound.weight_kg, 2))
        inbound.status = "IN_CONFIRM"
        msg = "입고완료가 취소되어 실재고가 원복되었습니다."
    elif inbound.status == "IN_CLAIM":
        inbound.status = "IN_DONE"
        inbound.claim_reason = None
        msg = "입고 클레임이 취소되어 입고완료 상태로 복구되었습니다."
    else:
        msg = "처리할 수 없는 상태입니다."
    db.commit()
    return {"message": msg}

# --- [현 재고장 및 출고 관리 API] ---
@app.get("/api/inventory", response_model=List[InventoryLotOut])
def get_inventory(
    limit: int = Query(default=1000, ge=1, le=5000),
    offset: int = Query(default=0, ge=0),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    today_check = datetime.now().strftime("%Y-%m-%d")

    expired_count = db.query(ReservationRecord).filter(
        ReservationRecord.status == "HOLD",
        ReservationRecord.expire_date < today_check
    ).update({
        ReservationRecord.status: "CANCELLED",
        ReservationRecord.cancel_type: "기간만료 자동취소",
        ReservationRecord.cancel_reason: "예약 유효 만료일 경과 자동 해제",
        ReservationRecord.cancel_date: today_check,
    }, synchronize_session=False)
    if expired_count:
        db.commit()

    lots = (
        db.query(InventoryLot)
        .filter(InventoryLot.current_box_qty > 0)
        .order_by(InventoryLot.exp_date)
        .offset(offset)
        .limit(limit)
        .all()
    )

    res_rows = (
        db.query(
            ReservationRecord.lot_id,
            func.sum(ReservationRecord.box_qty).label("box"),
            func.sum(ReservationRecord.weight_kg).label("weight"),
        )
        .filter(ReservationRecord.status == "HOLD")
        .group_by(ReservationRecord.lot_id)
        .all()
    )
    res_totals = {r.lot_id: {"box": r.box or 0, "weight": round(r.weight or 0.0, 2)} for r in res_rows}

    lot_ids = [l.id for l in lots]
    customer_rows = (
        db.query(ReservationRecord)
        .filter(ReservationRecord.status == "HOLD", ReservationRecord.lot_id.in_(lot_ids))
        .all()
        if lot_ids else []
    )
    customer_map = {}
    for r in customer_rows:
        customer_map.setdefault(r.lot_id, []).append(f"{r.customer} ({r.sales_rep}: {r.box_qty}Box)")

    result = []
    for l in lots:
        totals = res_totals.get(l.id, {"box": 0, "weight": 0.0})
        result.append({
            "id": l.id,
            "grid_no": l.grid_no,
            "sku_code": l.sku_code,
            "inbound_date": l.inbound_date,
            "bl_no": l.bl_no,
            "trace_no": l.trace_no,
            "process_from_date": l.process_from_date,
            "brand": l.brand,
            "item_name": l.item_name,
            "cut_name": l.cut_name,
            "storage_type": l.storage_type,
            "initial_box_qty": l.initial_box_qty,
            "initial_weight_kg": l.initial_weight_kg,
            "avg_box_weight": l.avg_box_weight,
            "current_box_qty": l.current_box_qty,
            "current_weight_kg": l.current_weight_kg,
            "cost_per_kg": l.cost_per_kg,
            "warehouse": l.warehouse,
            "exp_date": l.exp_date,
            "reserved_box_qty": totals["box"],
            "reserved_weight_kg": totals["weight"],
            "reserved_customers": ", ".join(customer_map.get(l.id, [])),
        })
    return result

@app.get("/api/outbounds", response_model=List[OutboundOut])
def get_outbounds(
    status: str,
    target_date: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    month: Optional[str] = None,
    limit: int = Query(default=500, ge=1, le=2000),
    offset: int = Query(default=0, ge=0),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    q = db.query(OutboundRecord).filter(OutboundRecord.status == status)
    if target_date:
        q = q.filter(OutboundRecord.outbound_date == target_date)
    elif month:
        q = q.filter(OutboundRecord.outbound_date.like(f"{month}%"))
    else:
        if start_date:
            q = q.filter(OutboundRecord.outbound_date >= start_date)
        if end_date:
            q = q.filter(OutboundRecord.outbound_date <= end_date)
    return q.order_by(desc(OutboundRecord.outbound_date), desc(OutboundRecord.id)).offset(offset).limit(limit).all()

@app.post("/api/outbounds/create-from-stock", response_model=MessageOut)
def create_outbound_from_stock(
    req: OutboundCreate,
    current_user: User = Depends(require_roles(ROLE_SALES, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).with_for_update().first()
    if not lot:
        raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")

    reserved_boxes = db.query(func.coalesce(func.sum(ReservationRecord.box_qty), 0)).filter(
        ReservationRecord.lot_id == lot.id, ReservationRecord.status == "HOLD"
    ).scalar()
    avail_box = lot.current_box_qty - reserved_boxes

    if req.box_qty > avail_box or req.box_qty <= 0:
        raise HTTPException(status_code=400, detail=f"출하 가능 수량({avail_box} Box)을 초과했습니다. (예약 홀딩: {reserved_boxes} Box)")

    avg_weight = lot.avg_box_weight if lot.avg_box_weight > 0 else (round(lot.initial_weight_kg / lot.initial_box_qty, 2) if lot.initial_box_qty > 0 else 20.0)
    calc_weight = round(avg_weight * req.box_qty, 2)

    lot.current_box_qty -= req.box_qty
    lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - calc_weight, 2))

    outbound_no = f"OUT-{datetime.now().strftime('%Y%m%d%H%M%S')}-{random.randint(10,99)}"
    out_date = req.outbound_date if req.outbound_date else datetime.now().strftime("%Y-%m-%d")
    total = round(calc_weight * req.unit_price_kg)
    outbound = OutboundRecord(
        outbound_no=outbound_no, grid_no=lot.grid_no, inbound_date=lot.inbound_date,
        outbound_date=out_date, lot_id=lot.id, customer=req.customer,
        bl_no=lot.bl_no, trace_no=lot.trace_no, process_from_date=lot.process_from_date,
        brand=lot.brand, item_name=lot.item_name, cut_name=lot.cut_name,
        storage_type=lot.storage_type, box_qty=req.box_qty, avg_box_weight=avg_weight,
        weight_kg=calc_weight, unit_price_kg=req.unit_price_kg, total_amount=total,
        exp_date=lot.exp_date, warehouse=lot.warehouse, is_weighed=lot.is_weighed, status="OUT_REQUEST"
    )
    db.add(outbound)
    db.commit()
    return {"message": f"출고요청 등록 완료 (평중 {avg_weight}kg 기준 {calc_weight}kg 적용, 재고장 즉시 차감됨)"}

@app.post("/api/outbounds/{outbound_id}/advance", response_model=MessageOut)
def advance_outbound(
    outbound_id: int,
    current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).with_for_update().first()
    if not outbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    if outbound.status == "OUT_REQUEST":
        outbound.status = "OUT_CONFIRM"
        msg = "출고확정 단계로 이동되었습니다."
    elif outbound.status == "OUT_CONFIRM":
        outbound.status = "OUT_DONE"
        outbound.processed_date = datetime.now().strftime("%Y-%m-%d")
        msg = "배차 및 출고완료 처리가 완료되었습니다."
    else:
        msg = "이미 출고완료된 전표입니다."
    db.commit()
    return {"message": msg}

@app.post("/api/outbounds/{outbound_id}/claim", response_model=MessageOut)
def register_outbound_claim(
    outbound_id: int, req: ClaimRegister,
    current_user: User = Depends(require_roles(ROLE_SALES, ROLE_WAREHOUSE, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    outbound.status = "OUT_CLAIM"
    outbound.claim_reason = req.reason
    outbound.processed_date = req.processed_date if req.processed_date else datetime.now().strftime("%Y-%m-%d")
    db.commit()
    return {"message": f"출고 클레임 등록이 완료되었습니다. (처리일자: {outbound.processed_date})"}

@app.put("/api/outbounds/{outbound_id}", response_model=MessageOut)
def update_outbound(
    outbound_id: int, req: UpdateOutbound,
    current_user: User = Depends(require_roles(ROLE_SALES, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    if outbound.status in ["OUT_DONE", "OUT_CLAIM"]:
        raise HTTPException(status_code=400, detail="완료 또는 클레임 상태의 전표는 수정할 수 없습니다.")
    outbound.outbound_date = req.outbound_date
    outbound.box_qty = req.box_qty
    outbound.weight_kg = req.weight_kg
    outbound.unit_price_kg = req.unit_price_kg
    outbound.total_amount = round(req.weight_kg * req.unit_price_kg)
    db.commit()
    return {"message": "출고 전표가 수정되었습니다."}

@app.post("/api/outbounds/{outbound_id}/revert", response_model=MessageOut)
def revert_outbound(
    outbound_id: int,
    current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).with_for_update().first()
    if not outbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    if outbound.status == "OUT_REQUEST":
        lot = db.query(InventoryLot).filter(InventoryLot.id == outbound.lot_id).with_for_update().first()
        if lot:
            lot.current_box_qty += outbound.box_qty
            lot.current_weight_kg = round(lot.current_weight_kg + outbound.weight_kg, 2)
        msg = f"출고요청이 취소되어 [{outbound.grid_no}] 재고장으로 {outbound.box_qty}Box / {outbound.weight_kg}kg가 원복되었습니다."
        db.delete(outbound)
    elif outbound.status == "OUT_CONFIRM":
        outbound.status = "OUT_REQUEST"
        msg = "출고확정이 취소되어 [출고요청리스트]로 반려되었습니다."
    elif outbound.status == "OUT_DONE":
        outbound.status = "OUT_CONFIRM"
        msg = "출고완료가 취소되어 [출고확정리스트]로 반려되었습니다."
    elif outbound.status == "OUT_CLAIM":
        outbound.status = "OUT_DONE"
        outbound.claim_reason = None
        msg = "출고 클레임이 취소되어 출고완료 상태로 복구되었습니다."
    else:
        msg = "처리할 수 없는 상태입니다."
    db.commit()
    return {"message": msg}

# --- [예약 관리 API] ---
@app.get("/api/reservations", response_model=List[ReservationOut])
def get_reservations(status: str, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    return db.query(ReservationRecord).filter(ReservationRecord.status == status).order_by(desc(ReservationRecord.id)).all()

@app.post("/api/reservations", response_model=MessageOut)
def create_reservation(
    req: ReservationCreate,
    current_user: User = Depends(require_roles(ROLE_SALES, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).with_for_update().first()
    if not lot:
        raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")

    already_reserved = db.query(func.coalesce(func.sum(ReservationRecord.box_qty), 0)).filter(
        ReservationRecord.lot_id == lot.id, ReservationRecord.status == "HOLD"
    ).scalar()
    avail_box = lot.current_box_qty - already_reserved

    if req.box_qty > avail_box or req.box_qty <= 0:
        raise HTTPException(status_code=400, detail=f"예약 가능 잔여 박스({avail_box} Box)를 초과하여 예약할 수 없습니다.")

    avg_w = lot.avg_box_weight if lot.avg_box_weight > 0 else 20.0
    calc_weight = round(avg_w * req.box_qty, 2)
    res_no = f"RES-{datetime.now().strftime('%Y%m%d%H%M%S')}-{random.randint(10,99)}"
    total = round(calc_weight * req.unit_price_kg)

    res = ReservationRecord(
        res_no=res_no,
        lot_id=lot.id,
        sales_rep=req.sales_rep,
        customer=req.customer,
        grid_no=lot.grid_no,
        item_name=lot.item_name,
        cut_name=lot.cut_name,
        storage_type=lot.storage_type,
        box_qty=req.box_qty,
        weight_kg=calc_weight,
        unit_price_kg=req.unit_price_kg,
        total_amount=total,
        exp_date=lot.exp_date,
        expire_date=req.expire_date,
        status="HOLD"
    )
    db.add(res)
    db.commit()
    return {"message": f"[{req.sales_rep}] 담당자 예약 등록 완료 ({req.box_qty}Box 홀딩)"}

@app.put("/api/reservations/{res_id}", response_model=MessageOut)
def update_reservation(
    res_id: int, req: ReservationUpdate,
    current_user: User = Depends(require_roles(ROLE_SALES, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    res = db.query(ReservationRecord).filter(ReservationRecord.id == res_id).with_for_update().first()
    if not res:
        raise HTTPException(status_code=404, detail="예약 내역을 찾을 수 없습니다.")
    if res.status != "HOLD":
        raise HTTPException(status_code=400, detail="홀딩 상태의 예약만 수정할 수 있습니다.")

    lot = db.query(InventoryLot).filter(InventoryLot.id == res.lot_id).with_for_update().first()
    if not lot:
        raise HTTPException(status_code=404, detail="연결된 재고 로트를 찾을 수 없습니다.")

    other_reserved = db.query(func.coalesce(func.sum(ReservationRecord.box_qty), 0)).filter(
        ReservationRecord.lot_id == lot.id,
        ReservationRecord.status == "HOLD",
        ReservationRecord.id != res_id,
    ).scalar()
    avail_box = lot.current_box_qty - other_reserved

    if req.box_qty > avail_box or req.box_qty <= 0:
        raise HTTPException(status_code=400, detail=f"예약 가능 잔여 박스({avail_box} Box)를 초과합니다.")

    avg_w = lot.avg_box_weight if lot.avg_box_weight > 0 else 20.0
    res.box_qty = req.box_qty
    res.weight_kg = round(avg_w * req.box_qty, 2)
    res.unit_price_kg = req.unit_price_kg
    res.total_amount = round(res.weight_kg * req.unit_price_kg)
    res.expire_date = req.expire_date
    db.commit()
    return {"message": "예약 정보가 수정되었습니다."}

@app.post("/api/reservations/{res_id}/cancel", response_model=MessageOut)
def cancel_reservation(
    res_id: int, req: ReservationCancelReq,
    current_user: User = Depends(require_roles(ROLE_SALES, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    res = db.query(ReservationRecord).filter(ReservationRecord.id == res_id).first()
    if not res:
        raise HTTPException(status_code=404, detail="예약 내역을 찾을 수 없습니다.")

    res.status = "CANCELLED"
    res.cancel_type = req.cancel_type or "수동 요청취소"
    res.cancel_reason = req.cancel_reason
    res.cancel_date = datetime.now().strftime("%Y-%m-%d")
    db.commit()
    return {"message": "예약이 취소되어 예약요청취소리스트에 누적되었습니다."}

# --- [통합 클레임 관리 API] ---
@app.get("/api/claims", response_model=List[ClaimOut])
def get_claims(
    stage: Optional[str] = None,
    target_date: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    month: Optional[str] = None,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    claims = []

    if stage in [None, "ALL", "INBOUND"]:
        q_in = db.query(InboundRecord).filter(InboundRecord.status == "IN_CLAIM")
        if target_date:
            q_in = q_in.filter(InboundRecord.processed_date == target_date)
        elif month:
            q_in = q_in.filter(InboundRecord.processed_date.like(f"{month}%"))
        else:
            if start_date:
                q_in = q_in.filter(InboundRecord.processed_date >= start_date)
            if end_date:
                q_in = q_in.filter(InboundRecord.processed_date <= end_date)
        for r in q_in.all():
            claims.append({
                "id": r.id, "stage": "입고", "inbound_date": r.inbound_date,
                "processed_date": r.processed_date or r.inbound_date, "doc_no": r.inbound_no,
                "partner_name": r.vendor, "warehouse": r.warehouse, "bl_no": r.bl_no,
                "trace_no": r.trace_no, "process_from_date": r.process_from_date, "brand": r.brand,
                "item_name": r.item_name, "cut_name": r.cut_name, "box_qty": r.box_qty,
                "weight_kg": r.weight_kg, "total_amount": r.total_amount, "claim_reason": r.claim_reason,
                "grid_no": r.grid_no, "exp_date": r.exp_date, "raw_type": "INBOUND"
            })

    if stage in [None, "ALL", "OUTBOUND"]:
        q_out = db.query(OutboundRecord).filter(OutboundRecord.status == "OUT_CLAIM")
        if target_date:
            q_out = q_out.filter(OutboundRecord.processed_date == target_date)
        elif month:
            q_out = q_out.filter(OutboundRecord.processed_date.like(f"{month}%"))
        else:
            if start_date:
                q_out = q_out.filter(OutboundRecord.processed_date >= start_date)
            if end_date:
                q_out = q_out.filter(OutboundRecord.processed_date <= end_date)
        for r in q_out.all():
            claims.append({
                "id": r.id, "stage": "출고", "inbound_date": r.inbound_date or "-",
                "processed_date": r.processed_date or r.outbound_date, "doc_no": r.outbound_no,
                "partner_name": r.customer, "warehouse": r.warehouse, "bl_no": r.bl_no,
                "trace_no": r.trace_no, "process_from_date": r.process_from_date, "brand": r.brand,
                "item_name": r.item_name, "cut_name": r.cut_name, "box_qty": r.box_qty,
                "weight_kg": r.weight_kg, "total_amount": r.total_amount, "claim_reason": r.claim_reason,
                "grid_no": r.grid_no, "exp_date": r.exp_date, "raw_type": "OUTBOUND"
            })

    claims.sort(key=lambda x: x["processed_date"] or "", reverse=True)
    return claims

# --- [재고 조정 및 창고 전배 API] ---
@app.post("/api/inventory/adjust", response_model=MessageOut)
def adjust_stock(
    req: AdjustCreate,
    current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).with_for_update().first()
    if not lot:
        raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")
    if req.adj_box > lot.current_box_qty:
        raise HTTPException(status_code=400, detail=f"조정 수량이 현재고({lot.current_box_qty} Box)를 초과합니다.")
    lot.current_box_qty = max(0, lot.current_box_qty - req.adj_box)
    lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - req.adj_weight, 2))
    log = StockAdjustment(lot_id=lot.id, adj_type=req.adj_type, adj_box=req.adj_box, adj_weight=req.adj_weight, reason=req.reason)
    db.add(log)
    db.commit()
    return {"message": f"재고 조정({req.adj_type})이 완료되었습니다."}

@app.post("/api/inventory/transfer", response_model=MessageOut)
def transfer_warehouse(
    req: WarehouseTransferCreate,
    current_user: User = Depends(require_roles(ROLE_WAREHOUSE, ROLE_ADMIN)),
    db: Session = Depends(get_db),
):
    src_lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).with_for_update().first()
    if not src_lot:
        raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")
    if req.to_warehouse == src_lot.warehouse:
        raise HTTPException(status_code=400, detail="출발 창고와 도착 창고가 동일합니다.")

    reserved_boxes = db.query(func.coalesce(func.sum(ReservationRecord.box_qty), 0)).filter(
        ReservationRecord.lot_id == src_lot.id, ReservationRecord.status == "HOLD"
    ).scalar()
    avail_box = src_lot.current_box_qty - reserved_boxes
    if req.box_qty > avail_box or req.box_qty <= 0:
        raise HTTPException(status_code=400, detail=f"전배 가능 수량({avail_box} Box)을 초과했습니다. (예약 홀딩: {reserved_boxes} Box)")

    avg_w = src_lot.avg_box_weight if src_lot.avg_box_weight > 0 else 20.0
    move_weight = round(avg_w * req.box_qty, 2)

    src_lot.current_box_qty -= req.box_qty
    src_lot.current_weight_kg = max(0.0, round(src_lot.current_weight_kg - move_weight, 2))

    new_grid = generate_random_grid()
    dest_lot = InventoryLot(
        grid_no=new_grid,
        sku_code=src_lot.sku_code,
        inbound_date=datetime.now().strftime("%Y-%m-%d"),
        bl_no=src_lot.bl_no,
        trace_no=src_lot.trace_no,
        process_from_date=src_lot.process_from_date,
        brand=src_lot.brand,
        item_name=src_lot.item_name,
        cut_name=src_lot.cut_name,
        storage_type=src_lot.storage_type,
        initial_box_qty=req.box_qty,
        initial_weight_kg=move_weight,
        avg_box_weight=avg_w,
        current_box_qty=req.box_qty,
        current_weight_kg=move_weight,
        cost_per_kg=src_lot.cost_per_kg,
        warehouse=req.to_warehouse,
        exp_date=src_lot.exp_date,
        is_weighed=src_lot.is_weighed,
    )
    db.add(dest_lot)

    log = StockAdjustment(
        lot_id=src_lot.id,
        adj_type="창고간전배",
        adj_box=req.box_qty,
        adj_weight=move_weight,
        reason=f"[{src_lot.warehouse} -> {req.to_warehouse}] 전배 이동 (신규 GRID: {new_grid}). {req.reason or ''}".strip()
    )
    db.add(log)
    db.commit()
    return {"message": f"{src_lot.warehouse}에서 {req.to_warehouse}(으)로 {req.box_qty}Box 전배 완료되었습니다. (신규 GRID: {new_grid})"}

# -----------------------------------------------------------------------------
# 4. 프론트엔드 UI (JWT 인증 연동)
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
    aside { width: 240px; background: var(--sidebar); color: #94a3b8; display: flex; flex-direction: column; flex-shrink: 0; }
    .brand { padding: 20px; font-size: 1.15rem; font-weight: 700; color: #fff; border-bottom: 1px solid #1e293b; display: flex; align-items: center; gap: 8px; }
    .brand i { color: #38bdf8; }
    .user-profile { padding: 12px 18px; background: #1e293b; color: #e2e8f0; font-size: 0.82rem; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #334155; }
    .nav-category { font-size: 0.72rem; text-transform: uppercase; font-weight: 700; color: #64748b; padding: 14px 18px 6px; }
    .nav-item { padding: 10px 18px; color: #94a3b8; cursor: pointer; display: flex; align-items: center; gap: 10px; font-size: 0.88rem; transition: 0.2s ease; }
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
    .table-card { background: #fff; border: 1px solid var(--border); border-radius: 8px; overflow: hidden; box-shadow: 0 1px 2px rgba(0,0,0,0.03); }
    table { width: 100%; border-collapse: collapse; font-size: 0.84rem; }
    th { background: #f8fafc; color: var(--muted); padding: 11px 12px; border-bottom: 1px solid var(--border); text-align: left; font-weight: 600; white-space: nowrap; }
    td { padding: 11px 12px; border-bottom: 1px solid var(--border); vertical-align: middle; }
    tr:hover td { background: #f8fafc; }
    .btn { padding: 5px 10px; border-radius: 6px; font-size: 0.8rem; font-weight: 600; border: 1px solid transparent; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-outline { background: #fff; border-color: var(--border); color: #334155; }
    .btn-outline:hover { background: #f1f5f9; }
    .badge { padding: 3px 6px; border-radius: 4px; font-size: 0.74rem; font-weight: 600; }
    .grid-tag { background: #312e81; color: #e0e7ff; padding: 2px 6px; border-radius: 4px; font-weight: 700; font-family: monospace; font-size: 0.78rem; }
    .modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 100; align-items: center; justify-content: center; }
    .modal-box { background: #fff; border-radius: 10px; width: 560px; padding: 24px; max-height: 90vh; overflow-y: auto; }
    .form-group { margin-bottom: 12px; display: flex; flex-direction: column; gap: 4px; font-size: 0.82rem; font-weight: 600; }
    .form-control { padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.88rem; }
  </style>
</head>
<body>
  <!-- 로그인 모달 -->
  <div class="modal-overlay" id="loginModal" style="display:flex;">
    <div class="modal-box" style="width:380px;">
      <h3 style="margin-bottom:12px; text-align:center;"><i class="bi bi-shield-lock-fill" style="color:var(--primary);"></i> ERP 로그인</h3>
      <div class="form-group"><label>아이디</label><input type="text" id="login_user" class="form-control" value="admin" /></div>
      <div class="form-group"><label>비밀번호</label><input type="password" id="login_pw" class="form-control" value="admin1234!" /></div>
      <button class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; margin-top:8px;" onclick="handleLogin()">로그인</button>
      <div style="font-size:0.75rem; color:#64748b; margin-top:10px; text-align:center;">테스트 계정: admin / sales1 / wh1 (비번: 아이디+1234!)</div>
    </div>
  </div>

  <aside>
    <div class="brand"><i class="bi bi-box-seam-fill"></i> MeatFlow ERP</div>
    <div class="user-profile">
      <span id="userBadge"><i class="bi bi-person-badge"></i> 미인증</span>
      <button class="btn btn-outline" style="padding:2px 6px; font-size:0.72rem;" onclick="handleLogout()">로그아웃</button>
    </div>
    <div class="nav-category">프로세스 관리</div>
    <div class="nav-item active" onclick="switchMainTab('INBOUND', this)"><i class="bi bi-box-arrow-in-down"></i> 입고 관리</div>
    <div class="nav-item" onclick="switchMainTab('OUTBOUND', this)"><i class="bi bi-box-arrow-up"></i> 출고 관리 (재고장)</div>
    <div class="nav-item" onclick="switchMainTab('RESERVATION', this)"><i class="bi bi-calendar-check-fill" style="color:#38bdf8;"></i> 예약 관리</div>
    <div class="nav-item" onclick="switchMainTab('CLAIM', this)"><i class="bi bi-exclamation-octagon-fill" style="color:#ef4444;"></i> 클레임 관리</div>
    <div class="nav-category">기본정보 마스터</div>
    <div class="nav-item" onclick="switchMainTab('PARTNER', this)"><i class="bi bi-building"></i> 거래처 정보 관리</div>
    <div class="nav-item" onclick="switchMainTab('ITEM_CUT_MASTER', this)"><i class="bi bi-diagram-3"></i> 품목/부위 마스터 관리</div>
  </aside>

  <main>
    <header>
      <h2 id="pageTitle" style="font-size:1.25rem;">입고 프로세스 관리</h2>
      <div id="headerActions" style="display:flex; gap:8px;"></div>
    </header>

    <div class="content">
      <div class="sub-tabs" id="subTabs" style="display:flex;"></div>
      <div class="filter-card" id="dateFilterCard">
        <div class="filter-group">
          <span id="dateFilterLabel"><i class="bi bi-calendar-range"></i> 기준 일자:</span>
          <input type="date" id="filter_start" class="filter-input" />
          <span>~</span>
          <input type="date" id="filter_end" class="filter-input" />
        </div>
        <div class="filter-group" style="flex-grow:1; max-width:380px;">
          <input type="text" id="filter_keyword" class="filter-input" style="width:100%;" placeholder="품목, 부위, 거래처, 이력번호 검색..." onkeypress="if(event.keyCode==13){loadData();}" />
        </div>
        <button class="btn btn-primary" onclick="loadData()"><i class="bi bi-search"></i> 조회</button>
      </div>

      <div class="table-card">
        <div style="overflow-x: auto;">
          <table>
            <thead id="tableHead"></thead>
            <tbody id="tableBody"></tbody>
          </table>
        </div>
      </div>
    </div>
  </main>

  <script>
    let authToken = localStorage.getItem('meat_erp_token') || '';
    let currentRole = localStorage.getItem('meat_erp_role') || '';
    let currentMain = 'INBOUND', currentSub = 'IN_REQUEST';

    async function apiFetch(url, options = {}) {
      options.headers = options.headers || {};
      if (authToken) {
        options.headers['Authorization'] = 'Bearer ' + authToken;
      }
      const res = await fetch(url, options);
      if (res.status === 401) {
        document.getElementById('loginModal').style.display = 'flex';
        throw new Error('인증이 만료되었습니다.');
      }
      return res;
    }

    async function handleLogin() {
      const u = document.getElementById('login_user').value;
      const p = document.getElementById('login_pw').value;
      const form = new URLSearchParams();
      form.append('username', u);
      form.append('password', p);

      const r = await fetch('/api/auth/login', { method: 'POST', body: form });
      const d = await r.json();
      if (r.ok) {
        authToken = d.access_token;
        currentRole = d.role;
        localStorage.setItem('meat_erp_token', authToken);
        localStorage.setItem('meat_erp_role', currentRole);
        document.getElementById('userBadge').innerText = `${d.full_name || d.username} (${d.role})`;
        document.getElementById('loginModal').style.display = 'none';
        renderSubTabs();
        loadData();
      } else {
        alert(d.detail || '로그인에 실패했습니다.');
      }
    }

    function handleLogout() {
      localStorage.removeItem('meat_erp_token');
      localStorage.removeItem('meat_erp_role');
      authToken = '';
      currentRole = '';
      location.reload();
    }

    function switchMainTab(tab, el) {
      currentMain = tab;
      document.querySelectorAll('.nav-item').forEach(e => e.classList.remove('active'));
      if (el) el.classList.add('active');
      renderSubTabs();
      loadData();
    }

    function renderSubTabs() {
      const container = document.getElementById('subTabs');
      container.innerHTML = '';
      if (currentMain === 'INBOUND') {
        const steps = [{k:'IN_REQUEST', n:'1. 입고요청'}, {k:'IN_CONFIRM', n:'2. 입고확정'}, {k:'IN_DONE', n:'3. 입고완료'}];
        steps.forEach(s => { container.innerHTML += `<div class="sub-tab ${currentSub===s.k?'active':''}" onclick="setSubTab('${s.k}')">${s.n}</div>`; });
      } else if (currentMain === 'OUTBOUND') {
        const steps = [{k:'OUT_STOCK', n:'1. 재고장 (출하대기)'}, {k:'OUT_REQUEST', n:'2. 출고요청'}, {k:'OUT_CONFIRM', n:'3. 출고확정'}, {k:'OUT_DONE', n:'4. 출고완료'}];
        steps.forEach(s => { container.innerHTML += `<div class="sub-tab ${currentSub===s.k?'active':''}" onclick="setSubTab('${s.k}')">${s.n}</div>`; });
      }
    }

    function setSubTab(k) { currentSub = k; renderSubTabs(); loadData(); }

    async function loadData() {
      if (!authToken) return;
      const head = document.getElementById('tableHead'), body = document.getElementById('tableBody');
      body.innerHTML = '<tr><td colspan="10" style="text-align:center; padding:15px;">데이터 조회 중...</td></tr>';
      
      try {
        if (currentMain === 'INBOUND') {
          head.innerHTML = '<tr><th>전표번호</th><th>입고일자</th><th>창고</th><th>매입처</th><th>품목/부위</th><th>수량(Box)</th><th>중량(kg)</th><th>상태</th><th>GRID</th></tr>';
          const r = await apiFetch(`/api/inbounds?status=${currentSub}`);
          const list = await r.json();
          body.innerHTML = list.length ? list.map(i => `
            <tr><td><strong>${i.inbound_no}</strong></td><td>${i.inbound_date}</td><td>${i.warehouse}</td><td>${i.vendor}</td><td>${i.item_name} ${i.cut_name}</td><td>${i.box_qty}</td><td>${i.weight_kg}kg</td><td>${i.status}</td><td><span class="grid-tag">${i.grid_no||'-'}</span></td></tr>
          `).join('') : '<tr><td colspan="9" style="text-align:center; padding:20px;">내역이 없습니다.</td></tr>';
        } else if (currentMain === 'OUTBOUND') {
          if (currentSub === 'OUT_STOCK') {
            head.innerHTML = '<tr><th>B/L</th><th>이력번호</th><th>품목/부위</th><th>잔여(Box)</th><th>잔여(kg)</th><th>예약(Box)</th><th>원가</th><th>창고</th><th>소비기한</th><th>GRID</th></tr>';
            const r = await apiFetch('/api/inventory');
            const list = await r.json();
            body.innerHTML = list.length ? list.map(l => `
              <tr><td>${l.bl_no||'-'}</td><td><code>${l.trace_no}</code></td><td>${l.item_name} ${l.cut_name}</td><td><strong>${l.current_box_qty}</strong></td><td>${l.current_weight_kg}kg</td><td style="color:#0284c7; font-weight:700;">${l.reserved_box_qty}</td><td>${l.cost_per_kg.toLocaleString()}원</td><td>${l.warehouse}</td><td>${l.exp_date}</td><td><span class="grid-tag">${l.grid_no}</span></td></tr>
            `).join('') : '<tr><td colspan="10" style="text-align:center; padding:20px;">보유 재고가 없습니다.</td></tr>';
          } else {
            head.innerHTML = '<tr><th>출고번호</th><th>출고일자</th><th>매출처</th><th>품목/부위</th><th>수량(Box)</th><th>중량(kg)</th><th>단가</th><th>총금액</th><th>상태</th></tr>';
            const r = await apiFetch(`/api/outbounds?status=${currentSub}`);
            const list = await r.json();
            body.innerHTML = list.length ? list.map(o => `
              <tr><td><strong>${o.outbound_no}</strong></td><td>${o.outbound_date}</td><td>${o.customer}</td><td>${o.item_name} ${o.cut_name}</td><td>${o.box_qty}</td><td>${o.weight_kg}kg</td><td>${o.unit_price_kg.toLocaleString()}원</td><td>${o.total_amount.toLocaleString()}원</td><td>${o.status}</td></tr>
            `).join('') : '<tr><td colspan="9" style="text-align:center; padding:20px;">내역이 없습니다.</td></tr>';
          }
        } else if (currentMain === 'RESERVATION') {
          head.innerHTML = '<tr><th>예약번호</th><th>영업사원</th><th>매출처</th><th>품목/부위</th><th>수량(Box)</th><th>중량(kg)</th><th>만료일</th><th>상태</th></tr>';
          const r = await apiFetch('/api/reservations?status=HOLD');
          const list = await r.json();
          body.innerHTML = list.length ? list.map(res => `
            <tr><td><strong>${res.res_no}</strong></td><td>${res.sales_rep}</td><td>${res.customer}</td><td>${res.item_name} ${res.cut_name}</td><td>${res.box_qty}</td><td>${res.weight_kg}kg</td><td style="color:#b91c1c;">${res.expire_date}</td><td>${res.status}</td></tr>
          `).join('') : '<tr><td colspan="8" style="text-align:center; padding:20px;">대기 중인 예약이 없습니다.</td></tr>';
        } else if (currentMain === 'CLAIM') {
          head.innerHTML = '<tr><th>발생</th><th>처리일자</th><th>전표번호</th><th>거래처</th><th>품목/부위</th><th>수량(Box)</th><th>사유</th></tr>';
          const r = await apiFetch('/api/claims');
          const list = await r.json();
          body.innerHTML = list.length ? list.map(c => `
            <tr><td>${c.stage}</td><td>${c.processed_date}</td><td><strong>${c.doc_no}</strong></td><td>${c.partner_name}</td><td>${c.item_name} ${c.cut_name}</td><td>${c.box_qty}</td><td style="color:#b91c1c;">${c.claim_reason||'-'}</td></tr>
          `).join('') : '<tr><td colspan="7" style="text-align:center; padding:20px;">클레임 내역이 없습니다.</td></tr>';
        } else if (currentMain === 'PARTNER') {
          head.innerHTML = '<tr><th>ID</th><th>구분</th><th>상호명</th><th>사업자번호</th><th>담당자</th><th>연락처</th></tr>';
          const r = await apiFetch('/api/partners');
          const list = await r.json();
          body.innerHTML = list.length ? list.map(p => `
            <tr><td>#${p.id}</td><td>${p.type}</td><td><strong>${p.name}</strong></td><td>${p.biz_no||'-'}</td><td>${p.contact_person||'-'}</td><td>${p.phone||'-'}</td></tr>
          `).join('') : '<tr><td colspan="6" style="text-align:center; padding:20px;">거래처가 없습니다.</td></tr>';
        } else if (currentMain === 'ITEM_CUT_MASTER') {
          head.innerHTML = '<tr><th>부위ID</th><th>상위품목</th><th>축종</th><th>부위코드</th><th>부위명</th><th>보관</th></tr>';
          const r = await apiFetch('/api/cuts');
          const list = await r.json();
          body.innerHTML = list.length ? list.map(c => `
            <tr><td>#${c.id}</td><td>${c.parent_item_name}</td><td>${c.species}</td><td><code>${c.cut_code}</code></td><td><strong>${c.cut_name}</strong></td><td>${c.default_storage}</td></tr>
          `).join('') : '<tr><td colspan="6" style="text-align:center; padding:20px;">부위 마스터가 없습니다.</td></tr>';
        }
      } catch (err) {
        console.error(err);
      }
    }

    window.addEventListener('DOMContentLoaded', () => {
      if (authToken) {
        document.getElementById('loginModal').style.display = 'none';
        document.getElementById('userBadge').innerText = `접속중 (${currentRole})`;
        renderSubTabs();
        loadData();
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
    uvicorn.run(app, host="0.0.0.0", port=port)
