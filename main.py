import io
import os
import random
from datetime import datetime, date, timedelta
from typing import List, Optional

from fastapi import FastAPI, HTTPException, Depends, UploadFile, File, Query
from fastapi.responses import HTMLResponse
import openpyxl
from pydantic import BaseModel, ConfigDict
from sqlalchemy import (
    Column, Integer, Float, String, DateTime, ForeignKey, create_engine, desc, func, Index
)
from sqlalchemy.orm import declarative_base, sessionmaker, Session
import uvicorn

# -----------------------------------------------------------------------------
# 1. DB 설정 (SQLite 파일 기반)
# -----------------------------------------------------------------------------
DB_FILE = "meaterp_local.db"
DATABASE_URL = f"sqlite:///{DB_FILE}"

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


# [개선 #6] 세션 예외처리 공통화: 라우터에서 예외가 발생하면 자동 롤백,
# 정상 종료 시 커밋 누락분까지 안전하게 커밋합니다.
# (각 라우터 내부의 개별 db.commit() 호출은 그대로 둬도 무방 - 중복 커밋은 무해)
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
    status = Column(String(20), default="IN_REQUEST", index=True)  # [개선] 인덱스 추가
    grid_no = Column(String(50), unique=True, nullable=True, index=True)  # [개선 #3] unique 제약 추가
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
    exp_date = Column(String(20), nullable=False, index=True)  # [개선] 유통기한 임박조회용 인덱스
    is_weighed = Column(String(10), default="N")
    grid_no = Column(String(50), unique=True, nullable=False, index=True)  # [개선 #3] unique 제약 추가
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
    status = Column(String(20), default="OUT_REQUEST", index=True)  # [개선] 인덱스 추가
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
    status = Column(String(20), default="HOLD", index=True)  # [개선] 인덱스 추가
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


Base.metadata.create_all(bind=engine)


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
    """[개선 #4] 엑셀 숫자 셀이 이력번호/문자열 컬럼에 들어올 때 '123.0' 형태로
    깨지는 문제 방지. openpyxl은 숫자로 보이는 셀을 float으로 반환하므로
    정수형이면 소수점을 제거하고 문자열로 변환한다."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def find_lot_by_grid(db: Session, grid_no: Optional[str]) -> Optional[InventoryLot]:
    """[개선 #3] trace_no + warehouse + exp_date 조합 매칭 대신
    입고 시점에 발급된 grid_no로 로트를 매칭한다. 훨씬 안정적이고 중복 로트 생성을 방지한다."""
    if not grid_no:
        return None
    return db.query(InventoryLot).filter(InventoryLot.grid_no == grid_no).first()


# -----------------------------------------------------------------------------
# 테스트용 시드 데이터 초기화 (입고, 재고, 출고, 예약, 클레임)
# -----------------------------------------------------------------------------
def init_sample_data():
    db = SessionLocal()
    try:
        # 1. 거래처 등록
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
                Partner(name="SGS 한국시험연구원", type="ETC", biz_no="113-81-77889", contact_person="인증실", phone="031-428-5700", address="경기도 안양시"),
            ]
            db.add_all(partners)

        # 2. 품목 / 부위 마스터 등록
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
                CutMaster(item_id=item_beef.id, cut_code="CUT-BEEF-03", cut_name="우삼겹(업진살)", default_storage="냉동"),
            ]
            db.add_all(cuts)

        # 3. 입고 전표 (요청, 확정, 완료, 클레임 상태별)
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
                ),
            ]
            db.add_all(inbounds)

        # 4. 현 재고장 (Inventory Lots)
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
                ),
            ]
            db.add_all(lots)
            db.flush()

        # 5. 출고 전표 (요청, 확정, 완료, 클레임 상태별)
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
                ),
            ]
            db.add_all(outbounds)

        # 6. 예약 관리 (홀딩 및 취소 누적건)
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
                ),
            ]
            db.add_all(reservations)

        db.commit()
    finally:
        db.close()


init_sample_data()

# -----------------------------------------------------------------------------
# 3. FastAPI 앱
# -----------------------------------------------------------------------------
app = FastAPI(title="MeatFlow Enterprise ERP")


# =============================================================================
# [1단계] 응답 스키마 (Pydantic Response Models)
# -----------------------------------------------------------------------------
# 기존 코드는 SQLAlchemy ORM 객체를 그대로 return 했습니다. FastAPI가 이를
# 직렬화하려고 하면 내부적으로 vars(obj)를 사용하는데, ORM 인스턴스에는
# _sa_instance_state(세션 및 매퍼에 대한 내부 참조를 포함하는 객체)가 포함되어
# 있어서 요청마다 불필요하게 무거운 직렬화가 발생하고, 데이터가 많아지면
# 응답 지연이나 드물게 직렬화 오류로 이어질 수 있습니다.
# response_model을 명시하면 (1) 에러 위험 제거 (2) 응답 속도 개선
# (3) 프론트에 필요한 필드만 노출 을 동시에 얻을 수 있습니다.
# =============================================================================

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


class ExpiringLotOut(BaseModel):
    id: int
    grid_no: str
    item_name: str
    cut_name: str
    storage_type: str
    warehouse: str
    exp_date: str
    days_left: int
    current_box_qty: int
    current_weight_kg: float


class DashboardSummaryOut(BaseModel):
    total_inventory_box: int
    total_inventory_weight_kg: float
    total_inventory_value: float
    warehouse_breakdown: List[dict]
    pending_inbound_count: int
    pending_outbound_count: int
    active_claim_count: int
    expiring_soon_count: int  # 7일 이내
    recent_7days_inbound_weight_kg: float
    recent_7days_outbound_weight_kg: float


# -----------------------------------------------------------------------------
# 요청(Request) 스키마
# -----------------------------------------------------------------------------
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
    # [3단계 부가기능] 예약 수정 - 기존에는 취소만 가능했음
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
    # [3단계 부가기능] 창고 간 재고 이동(전배)
    lot_id: int
    to_warehouse: str
    box_qty: int
    reason: Optional[str] = ""


# --- [거래처 API] ---
@app.get("/api/partners", response_model=List[PartnerOut])
def get_partners(type: Optional[str] = None, db: Session = Depends(get_db)):
    q = db.query(Partner)
    if type and type != "ALL":
        q = q.filter(Partner.type == type)
    return q.order_by(desc(Partner.id)).all()


@app.post("/api/partners", response_model=MessageOut)
def create_partner(req: PartnerCreate, db: Session = Depends(get_db)):
    p = Partner(name=req.name, type=req.type, biz_no=req.biz_no, contact_person=req.contact_person, phone=req.phone, address=req.address)
    db.add(p)
    db.commit()
    return {"message": "거래처 정보가 등록되었습니다."}


@app.delete("/api/partners/{partner_id}", response_model=MessageOut)
def delete_partner(partner_id: int, db: Session = Depends(get_db)):
    p = db.query(Partner).filter(Partner.id == partner_id).first()
    if not p:
        raise HTTPException(status_code=404, detail="거래처를 찾을 수 없습니다.")
    db.delete(p)
    db.commit()
    return {"message": "거래처가 삭제되었습니다."}


# --- [품목/부위 마스터 API] ---
@app.get("/api/items", response_model=List[ItemMasterOut])
def get_item_masters(db: Session = Depends(get_db)):
    return db.query(ItemMaster).order_by(ItemMaster.id).all()


@app.post("/api/items", response_model=MessageOut)
def create_item_master(req: ItemMasterCreate, db: Session = Depends(get_db)):
    if db.query(ItemMaster).filter(ItemMaster.item_code == req.item_code).first():
        raise HTTPException(status_code=400, detail="이미 등록된 품목코드입니다.")
    item = ItemMaster(item_code=req.item_code, item_name=req.item_name, species=req.species)
    db.add(item)
    db.commit()
    return {"message": "상위 품목 마스터가 등록되었습니다."}


@app.delete("/api/items/{item_id}", response_model=MessageOut)
def delete_item_master(item_id: int, db: Session = Depends(get_db)):
    item = db.query(ItemMaster).filter(ItemMaster.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="품목을 찾을 수 없습니다.")
    db.query(CutMaster).filter(CutMaster.item_id == item_id).delete()
    db.delete(item)
    db.commit()
    return {"message": "품목 및 소속 부위가 삭제되었습니다."}


@app.get("/api/cuts", response_model=List[CutMasterOut])
def get_cut_masters(item_id: Optional[int] = None, db: Session = Depends(get_db)):
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
def create_cut_master(req: CutMasterCreate, db: Session = Depends(get_db)):
    if db.query(CutMaster).filter(CutMaster.cut_code == req.cut_code).first():
        raise HTTPException(status_code=400, detail="이미 등록된 부위코드입니다.")
    cut = CutMaster(item_id=req.item_id, cut_code=req.cut_code, cut_name=req.cut_name, default_storage=req.default_storage)
    db.add(cut)
    db.commit()
    return {"message": "세부 부위 마스터가 등록되었습니다."}


@app.delete("/api/cuts/{cut_id}", response_model=MessageOut)
def delete_cut_master(cut_id: int, db: Session = Depends(get_db)):
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
    # [2단계 성능개선] 페이지네이션 파라미터 추가
    limit: int = Query(default=200, ge=1, le=1000),
    offset: int = Query(default=0, ge=0),
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
def create_inbound(req: InboundCreate, db: Session = Depends(get_db)):
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
async def upload_inbound_excel(file: UploadFile = File(...), db: Session = Depends(get_db)):
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
    failed_rows = []  # [개선 #4] 실패 행을 사용자에게 알려주기 위해 수집

    for r_idx, row in enumerate(rows[1:], start=2):
        if not row or all(v is None for v in row):
            continue
        try:
            # 엑셀 헤더 규격 매핑: 입고일자(0), 매입처(1), BL(2), 이력번호(3), 가공일(4), 브랜드(5), 품목(6), 부위(7), 보관(8), 박스(9), 실중량(10), 단가(11), 창고(12)
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
            # [개선 #4] 조용히 버리지 않고 실패 사유를 기록
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
def register_inbound_claim(inbound_id: int, req: ClaimRegister, db: Session = Depends(get_db)):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).with_for_update().first()
    if not inbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    if inbound.status == "IN_DONE":
        # [개선 #3] grid_no 기반 매칭으로 통일
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
def update_inbound(inbound_id: int, req: UpdateInbound, db: Session = Depends(get_db)):
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
def advance_inbound(inbound_id: int, db: Session = Depends(get_db)):
    # [2단계 개선] with_for_update로 행 잠금하여 동시 처리 시 중복 반영 방지
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

        # [개선 #3] grid_no 기반 로트 매칭 (기존: trace_no+warehouse+exp_date 조합 - 취약함)
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
def revert_inbound(inbound_id: int, db: Session = Depends(get_db)):
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
        lot = find_lot_by_grid(db, inbound.grid_no)  # [개선 #3]
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
    limit: int = Query(default=500, ge=1, le=2000),
    offset: int = Query(default=0, ge=0),
    db: Session = Depends(get_db),
):
    today_check = datetime.now().strftime("%Y-%m-%d")

    # 1. 예약 만료일 경과 건 자동 취소 체크
    # 참고: 조회(GET) API에서 부수효과(자동취소+commit)를 갖는 것은 REST 원칙상
    # 바람직하지 않습니다(캐싱 불가, 예측 불가능한 동작). 운영 환경에서는
    # APScheduler 등으로 별도 배치 잡으로 분리하는 것을 권장합니다.
    # 우선은 기존 요구사항을 지키되 아래처럼 SQL bulk update로 성능만 개선합니다.
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

    # [개선 #5] Python 루프 집계 대신 SQL GROUP BY로 한 번에 예약 합계 조회
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

    # 고객명 리스트는 요약 목적이라 그대로 유지하되, 필요한 lot_id만 조회
    lot_ids = [l.id for l in lots]
    customer_rows = (
        db.query(ReservationRecord)
        .filter(ReservationRecord.status == "HOLD", ReservationRecord.lot_id.in_(lot_ids))
        .all()
        if lot_ids else []
    )
    customer_map: dict = {}
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
    limit: int = Query(default=200, ge=1, le=1000),
    offset: int = Query(default=0, ge=0),
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
def create_outbound_from_stock(req: OutboundCreate, db: Session = Depends(get_db)):
    # [2단계 개선 #2] with_for_update()로 행 잠금 -> 동시 요청 시 오버셀 방지
    lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).with_for_update().first()
    if not lot:
        raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")

    # [개선 #5] SQL 집계로 예약 홀딩 수량 계산
    reserved_boxes = db.query(func.coalesce(func.sum(ReservationRecord.box_qty), 0)).filter(
        ReservationRecord.lot_id == lot.id, ReservationRecord.status == "HOLD"
    ).scalar()
    avail_box = lot.current_box_qty - reserved_boxes

    if req.box_qty > avail_box or req.box_qty <= 0:
        raise HTTPException(status_code=400, detail=f"출하 가능 수량({avail_box} Box)을 초과했습니다. (예약 홀딩: {reserved_boxes} Box)")

    avg_weight = lot.avg_box_weight if lot.avg_box_weight > 0 else (round(lot.initial_weight_kg / lot.initial_box_qty, 2) if lot.initial_box_qty > 0 else 20.0)
    calc_weight = round(avg_weight * req.box_qty, 2)

    # 행 잠금 상태에서 다시 한 번 재검증 (트랜잭션 내 최종 방어선)
    if lot.current_box_qty - reserved_boxes < req.box_qty:
        raise HTTPException(status_code=409, detail="재고 상태가 변경되어 처리할 수 없습니다. 다시 시도해주세요.")

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
def advance_outbound(outbound_id: int, db: Session = Depends(get_db)):
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
def register_outbound_claim(outbound_id: int, req: ClaimRegister, db: Session = Depends(get_db)):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    outbound.status = "OUT_CLAIM"
    outbound.claim_reason = req.reason
    outbound.processed_date = req.processed_date if req.processed_date else datetime.now().strftime("%Y-%m-%d")
    db.commit()
    return {"message": f"출고 클레임 등록이 완료되었습니다. (처리일자: {outbound.processed_date})"}


@app.put("/api/outbounds/{outbound_id}", response_model=MessageOut)
def update_outbound(outbound_id: int, req: UpdateOutbound, db: Session = Depends(get_db)):
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
def revert_outbound(outbound_id: int, db: Session = Depends(get_db)):
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
def get_reservations(status: str, db: Session = Depends(get_db)):
    return db.query(ReservationRecord).filter(ReservationRecord.status == status).order_by(desc(ReservationRecord.id)).all()


@app.post("/api/reservations", response_model=MessageOut)
def create_reservation(req: ReservationCreate, db: Session = Depends(get_db)):
    # [2단계 개선] 행 잠금으로 동시 예약 시 초과예약 방지
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
def update_reservation(res_id: int, req: ReservationUpdate, db: Session = Depends(get_db)):
    """[3단계 부가기능] 예약 수량/단가/만료일 수정.
    기존에는 취소 후 재등록만 가능했음 - 영업 담당자 편의를 위해 추가."""
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
def cancel_reservation(res_id: int, req: ReservationCancelReq, db: Session = Depends(get_db)):
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


@app.post("/api/inventory/adjust", response_model=MessageOut)
def adjust_stock(req: AdjustCreate, db: Session = Depends(get_db)):
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


# =============================================================================
# [3단계] 부가 기능
# =============================================================================

@app.get("/api/inventory/adjustments", response_model=List[StockAdjustmentOut])
def get_stock_adjustments(lot_id: Optional[int] = None, db: Session = Depends(get_db)):
    """재고조정 이력 조회 - 기존에는 POST만 있고 조회 API가 없어
    누가 언제 왜 조정했는지 화면에서 확인할 수 없었음."""
    q = db.query(StockAdjustment)
    if lot_id:
        q = q.filter(StockAdjustment.lot_id == lot_id)
    return q.order_by(desc(StockAdjustment.adjusted_at)).limit(500).all()


@app.get("/api/inventory/expiring", response_model=List[ExpiringLotOut])
def get_expiring_inventory(days: int = Query(default=7, ge=1, le=90), db: Session = Depends(get_db)):
    """유통기한 임박 재고 조회 (기본 7일 이내)."""
    today = date.today()
    cutoff = (today + timedelta(days=days)).strftime("%Y-%m-%d")
    today_str = today.strftime("%Y-%m-%d")

    lots = db.query(InventoryLot).filter(
        InventoryLot.current_box_qty > 0,
        InventoryLot.exp_date <= cutoff,
    ).order_by(InventoryLot.exp_date).all()

    result = []
    for l in lots:
        try:
            exp_d = datetime.strptime(l.exp_date, "%Y-%m-%d").date()
            days_left = (exp_d - today).days
        except Exception:
            days_left = 0
        result.append({
            "id": l.id, "grid_no": l.grid_no, "item_name": l.item_name, "cut_name": l.cut_name,
            "storage_type": l.storage_type, "warehouse": l.warehouse, "exp_date": l.exp_date,
            "days_left": days_left, "current_box_qty": l.current_box_qty,
            "current_weight_kg": l.current_weight_kg,
        })
    return result


@app.get("/api/dashboard/summary", response_model=DashboardSummaryOut)
def get_dashboard_summary(db: Session = Depends(get_db)):
    """대시보드 요약 지표 - 총 재고 금액, 창고별 재고량, 최근 7일 입출고 추이 등.
    기존에는 리스트 API만 있고 요약 지표가 전혀 없었음."""
    today = date.today()
    week_ago = (today - timedelta(days=7)).strftime("%Y-%m-%d")
    expiring_cutoff = (today + timedelta(days=7)).strftime("%Y-%m-%d")

    total_box = db.query(func.coalesce(func.sum(InventoryLot.current_box_qty), 0)).scalar()
    total_weight = db.query(func.coalesce(func.sum(InventoryLot.current_weight_kg), 0.0)).scalar()
    total_value = db.query(
        func.coalesce(func.sum(InventoryLot.current_weight_kg * InventoryLot.cost_per_kg), 0.0)
    ).scalar()

    warehouse_rows = (
        db.query(
            InventoryLot.warehouse,
            func.sum(InventoryLot.current_box_qty).label("box"),
            func.sum(InventoryLot.current_weight_kg).label("weight"),
        )
        .filter(InventoryLot.current_box_qty > 0)
        .group_by(InventoryLot.warehouse)
        .all()
    )
    warehouse_breakdown = [
        {"warehouse": r.warehouse, "box_qty": r.box or 0, "weight_kg": round(r.weight or 0.0, 2)}
        for r in warehouse_rows
    ]

    pending_inbound = db.query(func.count(InboundRecord.id)).filter(
        InboundRecord.status.in_(["IN_REQUEST", "IN_CONFIRM"])
    ).scalar()
    pending_outbound = db.query(func.count(OutboundRecord.id)).filter(
        OutboundRecord.status.in_(["OUT_REQUEST", "OUT_CONFIRM"])
    ).scalar()
    active_claims = (
        db.query(func.count(InboundRecord.id)).filter(InboundRecord.status == "IN_CLAIM").scalar()
        + db.query(func.count(OutboundRecord.id)).filter(OutboundRecord.status == "OUT_CLAIM").scalar()
    )
    expiring_soon = db.query(func.count(InventoryLot.id)).filter(
        InventoryLot.current_box_qty > 0, InventoryLot.exp_date <= expiring_cutoff
    ).scalar()

    recent_inbound_weight = db.query(func.coalesce(func.sum(InboundRecord.weight_kg), 0.0)).filter(
        InboundRecord.inbound_date >= week_ago
    ).scalar()
    recent_outbound_weight = db.query(func.coalesce(func.sum(OutboundRecord.weight_kg), 0.0)).filter(
        OutboundRecord.outbound_date >= week_ago
    ).scalar()

    return {
        "total_inventory_box": total_box,
        "total_inventory_weight_kg": round(total_weight, 2),
        "total_inventory_value": round(total_value, 2),
        "warehouse_breakdown": warehouse_breakdown,
        "pending_inbound_count": pending_inbound,
        "pending_outbound_count": pending_outbound,
        "active_claim_count": active_claims,
        "expiring_soon_count": expiring_soon,
        "recent_7days_inbound_weight_kg": round(recent_inbound_weight, 2),
        "recent_7days_outbound_weight_kg": round(recent_outbound_weight, 2),
    }


@app.post("/api/inventory/transfer", response_model=MessageOut)
def transfer_warehouse(req: WarehouseTransferCreate, db: Session = Depends(get_db)):
    """창고 간 재고 이동(전배). 기존 로트에서 수량을 차감하고, 대상 창고에
    동일 조건(grid_no 기반 신규 grid 발급)의 새 로트를 생성하거나 합산한다."""
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
        raise HTTPException(status_code=400, detail=f"이동 가능 수량({avail_box} Box)을 초과했습니다. (예약 홀딩 제외)")

    avg_w = src_lot.avg_box_weight if src_lot.avg_box_weight > 0 else 20.0
    move_weight = round(avg_w * req.box_qty, 2)

    src_lot.current_box_qty -= req.box_qty
    src_lot.current_weight_kg = max(0.0, round(src_lot.current_weight_kg - move_weight, 2))

    dest_lot = db.query(InventoryLot).filter(
        InventoryLot.trace_no == src_lot.trace_no,
        InventoryLot.warehouse == req.to_warehouse,
        InventoryLot.exp_date == src_lot.exp_date,
    ).with_for_update().first()

    if dest_lot:
        dest_lot.current_box_qty += req.box_qty
        dest_lot.current_weight_kg = round(dest_lot.current_weight_kg + move_weight, 2)
        dest_lot.initial_box_qty += req.box_qty
        dest_lot.initial_weight_kg = round(dest_lot.initial_weight_kg + move_weight, 2)
    else:
        dest_lot = InventoryLot(
            grid_no=generate_random_grid(),
            sku_code=src_lot.sku_code, inbound_date=src_lot.inbound_date, bl_no=src_lot.bl_no,
            trace_no=src_lot.trace_no, process_from_date=src_lot.process_from_date,
            brand=src_lot.brand, item_name=src_lot.item_name, cut_name=src_lot.cut_name,
            storage_type=src_lot.storage_type, initial_box_qty=req.box_qty,
            initial_weight_kg=move_weight, avg_box_weight=avg_w,
            current_box_qty=req.box_qty, current_weight_kg=move_weight,
            cost_per_kg=src_lot.cost_per_kg, warehouse=req.to_warehouse,
            exp_date=src_lot.exp_date, is_weighed=src_lot.is_weighed,
        )
        db.add(dest_lot)

    log = StockAdjustment(
        lot_id=src_lot.id, adj_type="WAREHOUSE_TRANSFER", adj_box=req.box_qty, adj_weight=move_weight,
        reason=f"{src_lot.warehouse} → {req.to_warehouse} 이동. {req.reason or ''}".strip()
    )
    db.add(log)
    db.commit()
    return {"message": f"{src_lot.warehouse} → {req.to_warehouse} 로 {req.box_qty}Box 이동 완료되었습니다."}


@app.get("/", response_class=HTMLResponse)
def serve_dashboard():
    return HTML_PAGE


HTML_PAGE = """<!DOCTYPE html><html><head><meta charset="utf-8"><title>MeatFlow ERP</title></head>
<body><h1>MeatFlow Enterprise ERP</h1><p>API 문서: <a href="/docs">/docs</a></p></body></html>"""


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
