import io
import os
import random
from datetime import datetime, date, timedelta
from typing import List, Optional

from fastapi import FastAPI, HTTPException, Depends, UploadFile, File
from fastapi.responses import HTMLResponse
import openpyxl
from pydantic import BaseModel
from sqlalchemy import (
    Column, Integer, Float, String, DateTime, ForeignKey, create_engine, desc
)
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
import uvicorn

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
    finally:
        db.close()

# -----------------------------------------------------------------------------
# 2. ORM 테이블 모델
# -----------------------------------------------------------------------------
class Partner(Base):
    __tablename__ = "partners"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), nullable=False)
    type = Column(String(30), nullable=False)  # VENDOR, CUSTOMER, WAREHOUSE, SHIPPING, ETC
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
    item_id = Column(Integer, ForeignKey("item_masters.id", ondelete="CASCADE"), nullable=False)
    cut_code = Column(String(50), unique=True, nullable=False)
    cut_name = Column(String(100), nullable=False)
    default_storage = Column(String(20), default="냉장")
    created_at = Column(DateTime, default=datetime.now)

class InboundRecord(Base):
    __tablename__ = "inbounds"
    id = Column(Integer, primary_key=True, index=True)
    inbound_no = Column(String(50), unique=True, index=True)
    inbound_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d"))
    processed_date = Column(String(20), nullable=True)
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
    exp_date = Column(String(20), nullable=False)
    is_weighed = Column(String(10), default="N")
    claim_reason = Column(String(255), nullable=True)
    status = Column(String(20), default="IN_REQUEST")
    grid_no = Column(String(50), nullable=True, index=True)
    created_at = Column(DateTime, default=datetime.now)

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
    exp_date = Column(String(20), nullable=False)
    is_weighed = Column(String(10), default="N")
    grid_no = Column(String(50), nullable=False, index=True)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

class OutboundRecord(Base):
    __tablename__ = "outbounds"
    id = Column(Integer, primary_key=True, index=True)
    outbound_no = Column(String(50), unique=True, index=True)
    inbound_date = Column(String(20), nullable=True)
    outbound_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d"))
    processed_date = Column(String(20), nullable=True)
    lot_id = Column(Integer, nullable=False)
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
    status = Column(String(20), default="OUT_REQUEST")
    grid_no = Column(String(50), nullable=True)
    created_at = Column(DateTime, default=datetime.now)

# 예약 관리 테이블
class ReservationRecord(Base):
    __tablename__ = "reservations"
    id = Column(Integer, primary_key=True, index=True)
    res_no = Column(String(50), unique=True, index=True)
    lot_id = Column(Integer, nullable=False)
    sales_rep = Column(String(50), nullable=False)      # 영업사원명
    customer = Column(String(100), nullable=False)     # 예약 매출처
    grid_no = Column(String(50), nullable=False)
    item_name = Column(String(100), nullable=False)
    cut_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    unit_price_kg = Column(Float, nullable=False)
    total_amount = Column(Float, nullable=False)
    exp_date = Column(String(20), nullable=False)
    expire_date = Column(String(20), nullable=False)   # 예약 유효 만료일
    cancel_date = Column(String(20), nullable=True)    # 취소/만료 처리일
    cancel_type = Column(String(50), nullable=True)    # '기간만료 자동취소', '수동 요청취소'
    cancel_reason = Column(String(255), nullable=True)
    status = Column(String(20), default="HOLD")        # HOLD(예약중), CANCELLED(취소완료)
    created_at = Column(DateTime, default=datetime.now)

class StockAdjustment(Base):
    __tablename__ = "stock_adjustments"
    id = Column(Integer, primary_key=True, index=True)
    lot_id = Column(Integer, nullable=False)
    adj_type = Column(String(50), nullable=False)
    adj_box = Column(Integer, nullable=False)
    adj_weight = Column(Float, nullable=False)
    reason = Column(String(255), nullable=True)
    adjusted_at = Column(DateTime, default=datetime.now)

Base.metadata.create_all(bind=engine)

def generate_random_grid() -> str:
    return f"GRID-{random.randint(100000, 999999)}"

# 초기 시드 데이터
def init_sample_data():
    db = SessionLocal()
    if db.query(Partner).count() == 0:
        partners = [
            Partner(name="(주)글로벌미트", type="VENDOR", biz_no="105-86-12345", contact_person="김수입", phone="010-1111-2222", address="서울시 송파구"),
            Partner(name="(주)하남돼지집", type="CUSTOMER", biz_no="120-81-99887", contact_person="박대표", phone="010-5555-6666", address="경기도 하남시 신장동 123"),
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
            CutMaster(item_id=item_beef.id, cut_code="CUT-BEEF-01", cut_name="척아이롤", default_storage="냉동")
        ]
        db.add_all(cuts)

    if db.query(InventoryLot).count() == 0:
        lots = [
            InventoryLot(grid_no="GRID-738201", sku_code="PK-CAN-01", inbound_date="2026-08-10", bl_no="ONEYVAN6948200", trace_no="802410290114", process_from_date="2026-07-20", brand="올리멜", item_name="돈육(돼지)", cut_name="삼겹살", storage_type="냉장", initial_box_qty=50, initial_weight_kg=1020.5, avg_box_weight=20.41, current_box_qty=45, current_weight_kg=918.45, cost_per_kg=8400, warehouse="광주냉장창고", exp_date="2026-09-17", is_weighed="N"),
            InventoryLot(grid_no="GRID-910482", sku_code="BF-USA-05", inbound_date="2026-08-12", bl_no="MAEU9201948201", trace_no="100392019482", process_from_date="2026-06-01", brand="엑셀", item_name="우육(소)", cut_name="척아이롤", storage_type="냉동", initial_box_qty=150, initial_weight_kg=3050.0, avg_box_weight=20.33, current_box_qty=120, current_weight_kg=2439.6, cost_per_kg=14500, warehouse="용인냉동센터", exp_date="2028-05-30", is_weighed="N")
        ]
        db.add_all(lots)
    db.commit()
    db.close()

init_sample_data()

# -----------------------------------------------------------------------------
# 3. FastAPI 라우터
# -----------------------------------------------------------------------------
app = FastAPI(title="MeatFlow Enterprise ERP")

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

class ReservationCancelReq(BaseModel):
    cancel_reason: str
    cancel_type: Optional[str] = "수동 요청취소"

class AdjustCreate(BaseModel):
    lot_id: int
    adj_type: str
    adj_box: int
    adj_weight: float
    reason: Optional[str] = ""

# 거래처 API
@app.get("/api/partners")
def get_partners(type: Optional[str] = None, db: Session = Depends(get_db)):
    q = db.query(Partner)
    if type and type != "ALL": q = q.filter(Partner.type == type)
    return q.order_by(desc(Partner.id)).all()

@app.post("/api/partners")
def create_partner(req: PartnerCreate, db: Session = Depends(get_db)):
    p = Partner(name=req.name, type=req.type, biz_no=req.biz_no, contact_person=req.contact_person, phone=req.phone, address=req.address)
    db.add(p)
    db.commit()
    return {"message": "거래처 정보가 등록되었습니다."}

@app.delete("/api/partners/{partner_id}")
def delete_partner(partner_id: int, db: Session = Depends(get_db)):
    p = db.query(Partner).filter(Partner.id == partner_id).first()
    if not p: raise HTTPException(status_code=404, detail="거래처를 찾을 수 없습니다.")
    db.delete(p)
    db.commit()
    return {"message": "거래처가 삭제되었습니다."}

# 품목/부위 마스터 API
@app.get("/api/items")
def get_item_masters(db: Session = Depends(get_db)):
    return db.query(ItemMaster).order_by(ItemMaster.id).all()

@app.post("/api/items")
def create_item_master(req: ItemMasterCreate, db: Session = Depends(get_db)):
    if db.query(ItemMaster).filter(ItemMaster.item_code == req.item_code).first():
        raise HTTPException(status_code=400, detail="이미 등록된 품목코드입니다.")
    item = ItemMaster(item_code=req.item_code, item_name=req.item_name, species=req.species)
    db.add(item)
    db.commit()
    return {"message": "상위 품목 마스터가 등록되었습니다."}

@app.delete("/api/items/{item_id}")
def delete_item_master(item_id: int, db: Session = Depends(get_db)):
    item = db.query(ItemMaster).filter(ItemMaster.id == item_id).first()
    if not item: raise HTTPException(status_code=404, detail="품목을 찾을 수 없습니다.")
    db.query(CutMaster).filter(CutMaster.item_id == item_id).delete()
    db.delete(item)
    db.commit()
    return {"message": "품목 및 소속 부위가 삭제되었습니다."}

@app.get("/api/cuts")
def get_cut_masters(item_id: Optional[int] = None, db: Session = Depends(get_db)):
    q = db.query(CutMaster.id, CutMaster.item_id, CutMaster.cut_code, CutMaster.cut_name, CutMaster.default_storage, ItemMaster.item_name, ItemMaster.species).join(ItemMaster, CutMaster.item_id == ItemMaster.id)
    if item_id: q = q.filter(CutMaster.item_id == item_id)
    results = q.order_by(CutMaster.item_id, CutMaster.id).all()
    return [
        {
            "id": r.id, "item_id": r.item_id, "cut_code": r.cut_code, "cut_name": r.cut_name,
            "default_storage": r.default_storage, "parent_item_name": r.item_name, "species": r.species
        } for r in results
    ]

@app.post("/api/cuts")
def create_cut_master(req: CutMasterCreate, db: Session = Depends(get_db)):
    if db.query(CutMaster).filter(CutMaster.cut_code == req.cut_code).first():
        raise HTTPException(status_code=400, detail="이미 등록된 부위코드입니다.")
    cut = CutMaster(item_id=req.item_id, cut_code=req.cut_code, cut_name=req.cut_name, default_storage=req.default_storage)
    db.add(cut)
    db.commit()
    return {"message": "세부 부위 마스터가 등록되었습니다."}

@app.delete("/api/cuts/{cut_id}")
def delete_cut_master(cut_id: int, db: Session = Depends(get_db)):
    cut = db.query(CutMaster).filter(CutMaster.id == cut_id).first()
    if not cut: raise HTTPException(status_code=404, detail="부위를 찾을 수 없습니다.")
    db.delete(cut)
    db.commit()
    return {"message": "부위 마스터가 삭제되었습니다."}

# 입고 API
@app.get("/api/inbounds")
def get_inbounds(
    status: str,
    target_date: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    month: Optional[str] = None,
    db: Session = Depends(get_db)
):
    q = db.query(InboundRecord).filter(InboundRecord.status == status)
    if target_date:
        q = q.filter(InboundRecord.inbound_date == target_date)
    elif month:
        q = q.filter(InboundRecord.inbound_date.like(f"{month}%"))
    else:
        if start_date: q = q.filter(InboundRecord.inbound_date >= start_date)
        if end_date: q = q.filter(InboundRecord.inbound_date <= end_date)
    return q.order_by(desc(InboundRecord.inbound_date), desc(InboundRecord.id)).all()

@app.post("/api/inbounds")
def create_inbound(req: InboundCreate, db: Session = Depends(get_db)):
    inbound_no = f"IN-{datetime.now().strftime('%Y%m%d%H%M%S')}"
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
    return {"message": f"입고등록 완료 ({grid_no})", "inbound_no": inbound_no}

@app.post("/api/inbounds/{inbound_id}/claim")
def register_inbound_claim(inbound_id: int, req: ClaimRegister, db: Session = Depends(get_db)):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).first()
    if not inbound: raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    
    if inbound.status == "IN_DONE":
        lot = db.query(InventoryLot).filter(
            InventoryLot.trace_no == inbound.trace_no,
            InventoryLot.warehouse == inbound.warehouse,
            InventoryLot.exp_date == inbound.exp_date
        ).first()
        if lot:
            lot.current_box_qty = max(0, lot.current_box_qty - inbound.box_qty)
            lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - inbound.weight_kg, 2))

    inbound.status = "IN_CLAIM"
    inbound.claim_reason = req.reason
    inbound.processed_date = req.processed_date if req.processed_date else datetime.now().strftime("%Y-%m-%d")
    db.commit()
    return {"message": f"입고 클레임 등록이 완료되었습니다. (처리일자: {inbound.processed_date})"}

@app.put("/api/inbounds/{inbound_id}")
def update_inbound(inbound_id: int, req: UpdateInbound, db: Session = Depends(get_db)):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).first()
    if not inbound: raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
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

@app.post("/api/inbounds/{inbound_id}/advance")
def advance_inbound(inbound_id: int, db: Session = Depends(get_db)):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).first()
    if not inbound: raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    if inbound.status == "IN_REQUEST":
        inbound.status = "IN_CONFIRM"
        msg = "입고확정 단계로 이동되었습니다."
    elif inbound.status == "IN_CONFIRM":
        inbound.status = "IN_DONE"
        inbound.processed_date = datetime.now().strftime("%Y-%m-%d")
        avg_w = round(inbound.weight_kg / inbound.box_qty, 2) if inbound.box_qty > 0 else 0.0
        lot = db.query(InventoryLot).filter(
            InventoryLot.trace_no == inbound.trace_no,
            InventoryLot.warehouse == inbound.warehouse,
            InventoryLot.exp_date == inbound.exp_date
        ).first()
        if lot:
            lot.initial_box_qty += inbound.box_qty
            lot.initial_weight_kg = round(lot.initial_weight_kg + inbound.weight_kg, 2)
            lot.avg_box_weight = round(lot.initial_weight_kg / lot.initial_box_qty, 2) if lot.initial_box_qty > 0 else avg_w
            lot.current_box_qty += inbound.box_qty
            lot.current_weight_kg = round(lot.current_weight_kg + inbound.weight_kg, 2)
            lot.cost_per_kg = inbound.cost_per_kg
            lot.is_weighed = inbound.is_weighed
        else:
            sku = "SKU-" + inbound.trace_no[:6]
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

@app.post("/api/inbounds/{inbound_id}/revert")
def revert_inbound(inbound_id: int, db: Session = Depends(get_db)):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).first()
    if not inbound: raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    if inbound.status == "IN_REQUEST":
        db.delete(inbound)
        msg = "입고요청 전표가 삭제되었습니다."
    elif inbound.status == "IN_CONFIRM":
        inbound.status = "IN_REQUEST"
        msg = "입고확정이 취소되어 [입고요청리스트]로 반려되었습니다."
    elif inbound.status == "IN_DONE":
        lot = db.query(InventoryLot).filter(
            InventoryLot.trace_no == inbound.trace_no,
            InventoryLot.warehouse == inbound.warehouse,
            InventoryLot.exp_date == inbound.exp_date
        ).first()
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

# 현 재고장 & 출고 관리 (예약 수량 맵핑 포함)
@app.get("/api/inventory")
def get_inventory(db: Session = Depends(get_db)):
    today_check = datetime.now().strftime("%Y-%m-%d")
    
    # 1. 예약 만료일 경과 건 자동 취소 체크 (AUTO_EXPIRED)
    expired_reservations = db.query(ReservationRecord).filter(
        ReservationRecord.status == "HOLD",
        ReservationRecord.expire_date < today_check
    ).all()
    for res in expired_reservations:
        res.status = "CANCELLED"
        res.cancel_type = "기간만료 자동취소"
        res.cancel_reason = "예약 유효 만료일 경과 자동 해제"
        res.cancel_date = today_check
    if expired_reservations:
        db.commit()

    lots = db.query(InventoryLot).filter(InventoryLot.current_box_qty > 0).order_by(InventoryLot.exp_date).all()
    
    # 활성 예약 건 집계 (Lot별 예약 박스, 중량, 예약처 목록)
    active_res = db.query(ReservationRecord).filter(ReservationRecord.status == "HOLD").all()
    res_map = {}
    for r in active_res:
        if r.lot_id not in res_map:
            res_map[r.lot_id] = {"box": 0, "weight": 0.0, "customers": []}
        res_map[r.lot_id]["box"] += r.box_qty
        res_map[r.lot_id]["weight"] = round(res_map[r.lot_id]["weight"] + r.weight_kg, 2)
        res_map[r.lot_id]["customers"].append(f"{r.customer} ({r.sales_rep}: {r.box_qty}Box)")

    result = []
    for l in lots:
        info = res_map.get(l.id, {"box": 0, "weight": 0.0, "customers": []})
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
            "reserved_box_qty": info["box"],
            "reserved_weight_kg": info["weight"],
            "reserved_customers": ", ".join(info["customers"]) if info["customers"] else ""
        })
    return result

@app.get("/api/outbounds")
def get_outbounds(
    status: str,
    target_date: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    month: Optional[str] = None,
    db: Session = Depends(get_db)
):
    q = db.query(OutboundRecord).filter(OutboundRecord.status == status)
    if target_date:
        q = q.filter(OutboundRecord.outbound_date == target_date)
    elif month:
        q = q.filter(OutboundRecord.outbound_date.like(f"{month}%"))
    else:
        if start_date: q = q.filter(OutboundRecord.outbound_date >= start_date)
        if end_date: q = q.filter(OutboundRecord.outbound_date <= end_date)
    return q.order_by(desc(OutboundRecord.outbound_date), desc(OutboundRecord.id)).all()

@app.post("/api/outbounds/create-from-stock")
def create_outbound_from_stock(req: OutboundCreate, db: Session = Depends(get_db)):
    lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).first()
    if not lot: raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")
    if req.box_qty > lot.current_box_qty or req.box_qty <= 0:
        raise HTTPException(status_code=400, detail="보유 잔여 박스 수량보다 많이 출하요청할 수 없습니다.")

    avg_weight = lot.avg_box_weight if lot.avg_box_weight > 0 else (round(lot.initial_weight_kg / lot.initial_box_qty, 2) if lot.initial_box_qty > 0 else 20.0)
    calc_weight = round(avg_weight * req.box_qty, 2)
    
    lot.current_box_qty -= req.box_qty
    lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - calc_weight, 2))

    outbound_no = f"OUT-{datetime.now().strftime('%Y%m%d%H%M%S')}"
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
    return {"message": f"출고요청 등록 완료 (평중 {avg_weight}kg 기준 {calc_weight}kg 적용, 재고장 즉시 차감됨)", "outbound_no": outbound_no}

@app.post("/api/outbounds/{outbound_id}/claim")
def register_outbound_claim(outbound_id: int, req: ClaimRegister, db: Session = Depends(get_db)):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound: raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    outbound.status = "OUT_CLAIM"
    outbound.claim_reason = req.reason
    outbound.processed_date = req.processed_date if req.processed_date else datetime.now().strftime("%Y-%m-%d")
    db.commit()
    return {"message": f"출고 클레임 등록이 완료되었습니다. (처리일자: {outbound.processed_date})"}

@app.post("/api/outbounds/{outbound_id}/revert")
def revert_outbound(outbound_id: int, db: Session = Depends(get_db)):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound: raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    if outbound.status == "OUT_REQUEST":
        lot = db.query(InventoryLot).filter(InventoryLot.id == outbound.lot_id).first()
        if lot:
            lot.current_box_qty += outbound.box_qty
            lot.current_weight_kg = round(lot.current_weight_kg + outbound.weight_kg, 2)
        db.delete(outbound)
        msg = f"출고요청이 취소되어 [{outbound.grid_no}] 재고장으로 {outbound.box_qty}Box / {outbound.weight_kg}kg가 원복되었습니다."
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

# --- [예약 관리 API (예약요청 / 취소 누적)] ---
@app.get("/api/reservations")
def get_reservations(status: str, db: Session = Depends(get_db)):
    return db.query(ReservationRecord).filter(ReservationRecord.status == status).order_by(desc(ReservationRecord.id)).all()

@app.post("/api/reservations")
def create_reservation(req: ReservationCreate, db: Session = Depends(get_db)):
    lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).first()
    if not lot: raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")
    if req.box_qty > lot.current_box_qty or req.box_qty <= 0:
        raise HTTPException(status_code=400, detail="보유 잔여 박스 수량보다 많이 예약할 수 없습니다.")

    avg_w = lot.avg_box_weight if lot.avg_box_weight > 0 else 20.0
    calc_weight = round(avg_w * req.box_qty, 2)
    res_no = f"RES-{datetime.now().strftime('%Y%m%d%H%M%S')}"
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

@app.post("/api/reservations/{res_id}/cancel")
def cancel_reservation(res_id: int, req: ReservationCancelReq, db: Session = Depends(get_db)):
    res = db.query(ReservationRecord).filter(ReservationRecord.id == res_id).first()
    if not res: raise HTTPException(status_code=404, detail="예약 내역을 찾을 수 없습니다.")
    
    res.status = "CANCELLED"
    res.cancel_type = req.cancel_type or "수동 요청취소"
    res.cancel_reason = req.cancel_reason
    res.cancel_date = datetime.now().strftime("%Y-%m-%d")
    db.commit()
    return {"message": "예약이 취소되어 예약요청취소리스트에 누적되었습니다."}

# 통합 클레임 관리 API
@app.get("/api/claims")
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
        if target_date: q_in = q_in.filter(InboundRecord.processed_date == target_date)
        elif month: q_in = q_in.filter(InboundRecord.processed_date.like(f"{month}%"))
        else:
            if start_date: q_in = q_in.filter(InboundRecord.processed_date >= start_date)
            if end_date: q_in = q_in.filter(InboundRecord.processed_date <= end_date)
        for r in q_in.all():
            claims.append({
                "id": r.id,
                "stage": "입고",
                "inbound_date": r.inbound_date,
                "processed_date": r.processed_date or r.inbound_date,
                "doc_no": r.inbound_no,
                "partner_name": r.vendor,
                "warehouse": r.warehouse,
                "bl_no": r.bl_no,
                "trace_no": r.trace_no,
                "process_from_date": r.process_from_date,
                "brand": r.brand,
                "item_name": r.item_name,
                "cut_name": r.cut_name,
                "box_qty": r.box_qty,
                "weight_kg": r.weight_kg,
                "total_amount": r.total_amount,
                "claim_reason": r.claim_reason,
                "grid_no": r.grid_no,
                "exp_date": r.exp_date,
                "raw_type": "INBOUND"
            })

    if stage in [None, "ALL", "OUTBOUND"]:
        q_out = db.query(OutboundRecord).filter(OutboundRecord.status == "OUT_CLAIM")
        if target_date: q_out = q_out.filter(OutboundRecord.processed_date == target_date)
        elif month: q_out = q_out.filter(OutboundRecord.processed_date.like(f"{month}%"))
        else:
            if start_date: q_out = q_out.filter(OutboundRecord.processed_date >= start_date)
            if end_date: q_out = q_out.filter(OutboundRecord.processed_date <= end_date)
        for r in q_out.all():
            claims.append({
                "id": r.id,
                "stage": "출고",
                "inbound_date": r.inbound_date or "-",
                "processed_date": r.processed_date or r.outbound_date,
                "doc_no": r.outbound_no,
                "partner_name": r.customer,
                "warehouse": r.warehouse,
                "bl_no": r.bl_no,
                "trace_no": r.trace_no,
                "process_from_date": r.process_from_date,
                "brand": r.brand,
                "item_name": r.item_name,
                "cut_name": r.cut_name,
                "box_qty": r.box_qty,
                "weight_kg": r.weight_kg,
                "total_amount": r.total_amount,
                "claim_reason": r.claim_reason,
                "grid_no": r.grid_no,
                "exp_date": r.exp_date,
                "raw_type": "OUTBOUND"
            })

    claims.sort(key=lambda x: x["processed_date"] or "", reverse=True)
    return claims

@app.post("/api/inventory/adjust")
def adjust_stock(req: AdjustCreate, db: Session = Depends(get_db)):
    lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).first()
    if not lot: raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")
    lot.current_box_qty = max(0, lot.current_box_qty - req.adj_box)
    lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - req.adj_weight, 2))
    log = StockAdjustment(lot_id=lot.id, adj_type=req.adj_type, adj_box=req.adj_box, adj_weight=req.adj_weight, reason=req.reason)
    db.add(log)
    db.commit()
    return {"message": f"재고 조정({req.adj_type})이 완료되었습니다."}

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
    :root {
      --primary: #2563eb;
      --sidebar: #0f172a;
      --bg: #f8fafc;
      --border: #e2e8f0;
      --text: #1e293b;
      --muted: #64748b;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Pretendard", "Segoe UI", sans-serif; }
    body { background: var(--bg); color: var(--text); display: flex; height: 100vh; overflow: hidden; }
    aside { width: 240px; background: var(--sidebar); color: #94a3b8; display: flex; flex-direction: column; flex-shrink: 0; }
    .brand { padding: 20px; font-size: 1.15rem; font-weight: 700; color: #fff; border-bottom: 1px solid #1e293b; display: flex; align-items: center; gap: 8px; }
    .brand i { color: #38bdf8; }
    .nav-category { font-size: 0.72rem; text-transform: uppercase; font-weight: 700; color: #64748b; padding: 14px 18px 6px; }
    .nav-item { padding: 10px 18px; color: #94a3b8; cursor: pointer; display: flex; align-items: center; gap: 10px; font-size: 0.88rem; transition: 0.2s ease; }
    .nav-item:hover, .nav-item.active { background: #1e293b; color: #38bdf8; font-weight: 600; }
    main { flex-grow: 1; display: flex; flex-direction: column; height: 100vh; overflow-y: auto; }
    header { background: #fff; border-bottom: 1px solid var(--border); padding: 14px 28px; display: flex; justify-content: space-between; align-items: center; }
    .content { padding: 20px 28px; display: flex; flex-direction: column; gap: 14px; }
    .sub-tabs { display: flex; gap: 6px; background: #e2e8f0; padding: 4px; border-radius: 8px; width: fit-content; }
    .sub-tab { padding: 8px 16px; border-radius: 6px; font-size: 0.85rem; font-weight: 600; cursor: pointer; color: #475569; }
    .sub-tab.active { background: #fff; color: var(--primary); box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .sub-tab-create { background: #dbeafe; color: #1d4ed8; font-weight: 700; border: 1px dashed #2563eb; }
    .sub-tab-create:hover { background: #bfdbfe; }
    .filter-card { background: #fff; border: 1px solid var(--border); border-radius: 8px; padding: 12px 18px; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px; }
    .filter-group { display: flex; align-items: center; gap: 8px; font-size: 0.85rem; font-weight: 600; }
    .filter-input { padding: 6px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.85rem; outline: none; background: #fff; }
    .filter-btn-group { display: flex; gap: 4px; }
    .filter-btn { padding: 6px 10px; border: 1px solid var(--border); background: #f8fafc; border-radius: 6px; font-size: 0.8rem; font-weight: 600; cursor: pointer; }
    .filter-btn.active { background: var(--primary); color: #fff; border-color: var(--primary); }
    .table-card { background: #fff; border: 1px solid var(--border); border-radius: 8px; overflow: hidden; box-shadow: 0 1px 2px rgba(0,0,0,0.03); }
    table { width: 100%; border-collapse: collapse; font-size: 0.84rem; }
    th { background: #f8fafc; color: var(--muted); padding: 11px 12px; border-bottom: 1px solid var(--border); text-align: left; font-weight: 600; white-space: nowrap; }
    td { padding: 11px 12px; border-bottom: 1px solid var(--border); vertical-align: middle; }
    tr:hover td { background: #f8fafc; }
    .btn { padding: 5px 10px; border-radius: 6px; font-size: 0.8rem; font-weight: 600; border: 1px solid transparent; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-success { background: #16a34a; color: #fff; }
    .btn-warning { background: #d97706; color: #fff; }
    .btn-danger { background: #dc2626; color: #fff; }
    .btn-outline { background: #fff; border-color: var(--border); color: #334155; }
    .btn-outline:hover { background: #f1f5f9; }
    .badge { padding: 3px 6px; border-radius: 4px; font-size: 0.74rem; font-weight: 600; }
    .badge-chilled { background: #e0f2fe; color: #0284c7; }
    .badge-frozen { background: #e0e7ff; color: #4f46e5; }
    .badge-vendor { background: #fef3c7; color: #b45309; }
    .badge-customer { background: #dcfce7; color: #15803d; }
    .badge-warehouse { background: #f1f5f9; color: #475569; }
    .badge-shipping { background: #ede9fe; color: #6d28d9; }
    .badge-etc { background: #f3f4f6; color: #374151; }
    .badge-stage-in { background: #e0f2fe; color: #0369a1; border: 1px solid #bae6fd; }
    .badge-stage-out { background: #fae8ff; color: #86198f; border: 1px solid #f5d0fe; }
    .grid-tag { background: #312e81; color: #e0e7ff; padding: 2px 6px; border-radius: 4px; font-weight: 700; font-family: monospace; font-size: 0.78rem; }
    .bl-tag { font-family: monospace; font-weight: 700; color: #0284c7; }
    .brand-tag { background: #f1f5f9; color: #334155; padding: 2px 5px; border-radius: 4px; font-weight: 600; font-size: 0.76rem; }
    .item-tag { background: #f0fdf4; color: #166534; padding: 2px 5px; border-radius: 4px; font-weight: 600; font-size: 0.76rem; border: 1px solid #bbf7d0; }
    .amount-highlight { font-weight: 700; color: #0f172a; }
    .remain-badge { padding: 2px 5px; border-radius: 4px; font-size: 0.72rem; font-weight: 700; }
    .remain-normal { background: #e0f2fe; color: #0369a1; }
    .remain-warning { background: #fee2e2; color: #b91c1c; }
    
    /* 예약재고 강조 링크 & 툴팁 */
    .res-link { color: #2563eb; font-weight: 700; cursor: pointer; text-decoration: underline; }
    .res-link:hover { color: #1d4ed8; }

    .modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 100; align-items: center; justify-content: center; }
    .modal-box { background: #fff; border-radius: 10px; width: 560px; padding: 24px; max-height: 90vh; overflow-y: auto; }
    .form-group { margin-bottom: 12px; display: flex; flex-direction: column; gap: 4px; font-size: 0.82rem; font-weight: 600; }
    .form-control { padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.88rem; }
    
    @media print {
      body * { visibility: hidden; }
      #printArea, #printArea * { visibility: visible; }
      #printArea { position: absolute; left: 0; top: 0; width: 100%; padding: 20px; font-family: 'Pretendard', sans-serif; color: #000; }
      .no-print { display: none !important; }
    }
  </style>
</head>
<body>
  <aside>
    <div class="brand"><i class="bi bi-box-seam-fill"></i> MeatFlow ERP</div>
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
      <div class="sub-tabs" id="subTabs" style="display:none;"></div>

      <div class="filter-card" id="dateFilterCard">
        <div class="filter-group">
          <span id="dateFilterLabel"><i class="bi bi-calendar-range"></i> 기준 일자:</span>
          <input type="date" id="filter_start" class="filter-input" />
          <span>~</span>
          <input type="date" id="filter_end" class="filter-input" />
        </div>
        
        <div class="filter-group" style="flex-grow:1; max-width:380px; display:flex; gap:6px;">
          <input type="text" id="filter_keyword" class="filter-input" style="width:100%; border-radius:6px;" placeholder="다중검색 (예: 삼겹살 AND 냉장, 광주 OR 용인, 삼겹, 목심)..." onkeypress="if(event.keyCode==13){loadData();}" />
          <button class="btn btn-primary" style="border-radius:6px; padding:6px 14px; white-space:nowrap;" onclick="loadData()"><i class="bi bi-search"></i> 조회</button>
        </div>

        <div style="display:flex; gap:6px;">
          <button class="btn btn-outline" style="color:#16a34a; border-color:#86efac; background:#f0fdf4;" onclick="downloadCurrentTableExcel()"><i class="bi bi-file-earmark-spreadsheet-fill"></i> 엑셀 다운로드</button>
        </div>

        <div class="filter-btn-group">
          <button class="filter-btn active" id="btn_today" onclick="setDateFilter('TODAY')">오늘 (당일)</button>
          <button class="filter-btn" id="btn_yesterday" onclick="setDateFilter('YESTERDAY')">전일</button>
          <button class="filter-btn" id="btn_this_month" onclick="setDateFilter('THIS_MONTH')">당월</button>
          <button class="filter-btn" id="btn_all" onclick="setDateFilter('ALL')">전체</button>
        </div>
      </div>

      <div id="summaryBar" style="display:flex; gap:16px; background:#eff6ff; border:1px solid #bfdbfe; border-radius:8px; padding:10px 18px; font-size:0.86rem; color:#1e40af; font-weight:600;">
        <span>조회 건수: <strong id="sum_count" style="color:#1d4ed8;">0</strong>건</span>
        <span>|</span>
        <span>총 박스 수량: <strong id="sum_box" style="color:#1d4ed8;">0</strong> Box</span>
        <span>|</span>
        <span>총 실중량: <strong id="sum_weight" style="color:#b91c1c;">0.0</strong> kg</span>
        <span>|</span>
        <span>총 금액: <strong id="sum_amount" style="color:#b91c1c;">0</strong> 원</span>
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

  <!-- 1. 신규 입고등록 모달 -->
  <div class="modal-overlay" id="inboundModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px;">신규 입고등록</h3>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>입고 일자 (KST 기준)</label><input type="date" id="in_date" class="form-control" /></div>
        <div class="form-group" style="flex:1;"><label>매입처 상호</label><input type="text" id="in_vendor" class="form-control" value="(주)글로벌미트" /></div>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>B/L No.</label><input type="text" id="in_bl_no" class="form-control" value="ONEYVAN6948200" /></div>
        <div class="form-group" style="flex:1;"><label>이력번호 (12자리)</label><input type="text" id="in_trace" class="form-control" value="802410290114" /></div>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>브랜드 / 가공장</label><input type="text" id="in_brand" class="form-control" value="올리멜" /></div>
        <div class="form-group" style="flex:1;">
          <label>상위 품목명 (축종)</label>
          <select id="in_item" class="form-control" onchange="autoCalculateExpDate()">
            <option value="돈육(돼지)">돈육(돼지)</option>
            <option value="우육(소)">우육(소)</option>
            <option value="계육(닭)">계육(닭)</option>
            <option value="기타축산물">기타축산물</option>
          </select>
        </div>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>세부 부위명</label><input type="text" id="in_cut" class="form-control" value="삼겹살" /></div>
        <div class="form-group" style="flex:1;">
          <label>보관 형태</label>
          <select id="in_storage" class="form-control" onchange="autoCalculateExpDate()">
            <option value="냉장">냉장 (Chilled)</option>
            <option value="냉동">냉동 (Frozen)</option>
          </select>
        </div>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;">
          <label style="color:#2563eb;">미트와치 가공일자 (FROM)</label>
          <input type="date" id="in_process_from" class="form-control" onchange="autoCalculateExpDate()" />
        </div>
        <div class="form-group" style="flex:1;">
          <label style="color:#b91c1c;">소비기한 (D-1 적용)</label>
          <input type="date" id="in_exp" class="form-control" />
        </div>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>입고 박스 수(Box)</label><input type="number" id="in_box" class="form-control" value="50" /></div>
        <div class="form-group" style="flex:1;"><label>총 실중량(kg)</label><input type="number" step="0.1" id="in_weight" class="form-control" value="1020.5" /></div>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>매입단가(원/kg)</label><input type="number" id="in_cost" class="form-control" value="8400" /></div>
        <div class="form-group" style="flex:1;">
          <label>보관창고 (선택)</label>
          <select id="in_warehouse" class="form-control"></select>
        </div>
      </div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal('inboundModal')">취소</button>
        <button class="btn btn-primary" onclick="submitInbound()">신규 입고등록 완료</button>
      </div>
    </div>
  </div>

  <!-- 2. 예약 등록 모달 -->
  <div class="modal-overlay" id="reserveModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px; color:#0284c7;">출고 예약 등록 (재고 홀딩)</h3>
      <input type="hidden" id="res_lot_id" />
      <div class="form-group">
        <label>담당 영업사원명</label>
        <input type="text" id="res_sales_rep" class="form-control" placeholder="예: 김영업 대리" value="김영업" />
      </div>
      <div class="form-group">
        <label>예약 매출처 상호명</label>
        <input type="text" id="res_customer" class="form-control" placeholder="예: (주)하남돼지집" value="(주)하남돼지집" />
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>예약 박스 수량(Box)</label><input type="number" id="res_box_qty" class="form-control" value="10" /></div>
        <div class="form-group" style="flex:1;"><label>예정 판매단가(원/kg)</label><input type="number" id="res_unit_price" class="form-control" value="11500" /></div>
      </div>
      <div class="form-group">
        <label style="color:#b91c1c;">예약 유효 만료일자 (해당일 경과 시 자동취소)</label>
        <input type="date" id="res_expire_date" class="form-control" />
      </div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal('reserveModal')">취소</button>
        <button class="btn btn-primary" onclick="submitReservation()">예약 등록</button>
      </div>
    </div>
  </div>

  <!-- 3. 거래처 등록 모달 -->
  <div class="modal-overlay" id="partnerModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px;">신규 거래처 등록</h3>
      <div class="form-group"><label>거래처/상호명</label><input type="text" id="p_name" class="form-control" placeholder="예: (주)한국축산유통" /></div>
      <div class="form-group"><label>대분류 구분</label>
        <select id="p_type" class="form-control">
          <option value="VENDOR">매입처 (Vendor)</option>
          <option value="CUSTOMER">매출처 (Customer)</option>
          <option value="WAREHOUSE">보관창고 (Warehouse)</option>
          <option value="SHIPPING">선사 / 포워딩 (Shipping Line)</option>
          <option value="ETC">기타 (시험검사기관/관공서 등)</option>
        </select>
      </div>
      <div class="form-group"><label>사업자등록번호</label><input type="text" id="p_biz_no" class="form-control" placeholder="000-00-00000" /></div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>담당자명</label><input type="text" id="p_contact" class="form-control" /></div>
        <div class="form-group" style="flex:1;"><label>연락처</label><input type="text" id="p_phone" class="form-control" placeholder="010-0000-0000" /></div>
      </div>
      <div class="form-group"><label>사업장 / 창고 주소</label><input type="text" id="p_address" class="form-control" /></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal('partnerModal')">취소</button>
        <button class="btn btn-primary" onclick="submitPartner()">등록 저장</button>
      </div>
    </div>
  </div>

  <!-- 4. 품목 등록 모달 -->
  <div class="modal-overlay" id="itemMasterModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px;">신규 품목(대분류) 등록</h3>
      <div class="form-group"><label>품목 코드</label><input type="text" id="m_item_code" class="form-control" placeholder="예: ITM-PORK" /></div>
      <div class="form-group"><label>품목명 (대분류)</label><input type="text" id="m_item_name" class="form-control" placeholder="예: 돈육(돼지), 우육(소)" /></div>
      <div class="form-group"><label>축종</label><select id="m_species" class="form-control"><option>돼지</option><option>소</option><option>닭</option><option>기타</option></select></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal('itemMasterModal')">취소</button>
        <button class="btn btn-primary" onclick="submitItemMaster()">품목 등록</button>
      </div>
    </div>
  </div>

  <!-- 5. 부위 등록 모달 -->
  <div class="modal-overlay" id="cutMasterModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px;">신규 부위(하위 세부분류) 등록</h3>
      <div class="form-group"><label>소속 상위 품목</label><select id="m_cut_parent_item" class="form-control"></select></div>
      <div class="form-group"><label>부위 코드</label><input type="text" id="m_cut_code" class="form-control" placeholder="예: CUT-PORK-04" /></div>
      <div class="form-group"><label>세부 부위명</label><input type="text" id="m_cut_name" class="form-control" placeholder="예: 삼겹살, 목심, 안심" /></div>
      <div class="form-group"><label>기본 보관형태</label><select id="m_cut_storage" class="form-control"><option value="냉장">냉장</option><option value="냉동">냉동</option></select></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal('cutMasterModal')">취소</button>
        <button class="btn btn-primary" onclick="submitCutMaster()">부위 등록</button>
      </div>
    </div>
  </div>

  <!-- 6. 출고요청서 팩스/PDF 미리보기 모달 -->
  <div class="modal-overlay" id="pdfModal">
    <div class="modal-box" style="width:780px;">
      <div class="no-print" style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; border-bottom:1px solid #e2e8f0; padding-bottom:10px;">
        <h3 style="font-size:1.1rem;"><i class="bi bi-file-earmark-pdf-fill" style="color:#dc2626;"></i> 출고요청서 (창고 팩스 발송용)</h3>
        <div style="display:flex; gap:8px;">
          <button class="btn btn-primary" onclick="window.print()"><i class="bi bi-printer-fill"></i> PDF 저장 / 팩스 인쇄</button>
          <button class="btn btn-outline" onclick="closeModal('pdfModal')">닫기</button>
        </div>
      </div>

      <div id="printArea" style="background:#fff; padding:20px; border:1px solid #ccc; font-size:12px; line-height:1.5;">
        <div style="text-align:center; margin-bottom:20px;">
          <h1 style="font-size:22px; text-decoration:underline; font-weight:800; letter-spacing:4px;">출 고 요 청 서 (배차/출하)</h1>
        </div>
        
        <table style="width:100%; border:1px solid #000; border-collapse:collapse; margin-bottom:15px;">
          <tr>
            <td style="width:15%; background:#f2f2f2; font-weight:bold; border:1px solid #000; padding:6px;">발 신 (본 사)</td>
            <td style="width:35%; border:1px solid #000; padding:6px;" id="pdf_sender">(주)한국축산유통</td>
            <td style="width:15%; background:#f2f2f2; font-weight:bold; border:1px solid #000; padding:6px;">수 신 (보관창고)</td>
            <td style="width:35%; border:1px solid #000; padding:6px;" id="pdf_warehouse">광주냉장창고 귀중</td>
          </tr>
          <tr>
            <td style="background:#f2f2f2; font-weight:bold; border:1px solid #000; padding:6px;">출고 요청일자</td>
            <td style="border:1px solid #000; padding:6px;" id="pdf_out_date">2026-08-20</td>
            <td style="background:#f2f2f2; font-weight:bold; border:1px solid #000; padding:6px;">출하처 (매출처)</td>
            <td style="border:1px solid #000; padding:6px;" id="pdf_customer">(주)하남돼지집</td>
          </tr>
          <tr>
            <td style="background:#f2f2f2; font-weight:bold; border:1px solid #000; padding:6px;">전표 관리번호</td>
            <td style="border:1px solid #000; padding:6px;" id="pdf_out_no">OUT-20260820-001</td>
            <td style="background:#f2f2f2; font-weight:bold; border:1px solid #000; padding:6px;">가공일(FROM)</td>
            <td style="border:1px solid #000; padding:6px;" id="pdf_process_from">-</td>
          </tr>
        </table>

        <div style="font-weight:bold; margin-bottom:6px;">■ 출하 대상 품목 및 수량 명세</div>
        <table style="width:100%; border:1px solid #000; border-collapse:collapse; text-align:center; margin-bottom:20px;">
          <thead>
            <tr style="background:#f2f2f2;">
              <th style="border:1px solid #000; padding:6px;">B/L No.</th>
              <th style="border:1px solid #000; padding:6px;">이력번호</th>
              <th style="border:1px solid #000; padding:6px;">품목/부위</th>
              <th style="border:1px solid #000; padding:6px;">보관</th>
              <th style="border:1px solid #000; padding:6px;">기준평중</th>
              <th style="border:1px solid #000; padding:6px;">출하 박스</th>
              <th style="border:1px solid #000; padding:6px;">출하 실중량(kg)</th>
              <th style="border:1px solid #000; padding:6px;">소비기한</th>
              <th style="border:1px solid #000; padding:6px;">GRID 매칭코드</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style="border:1px solid #000; padding:8px;" id="pdf_bl">-</td>
              <td style="border:1px solid #000; padding:8px;" id="pdf_trace">-</td>
              <td style="border:1px solid #000; padding:8px;" id="pdf_item_cut">-</td>
              <td style="border:1px solid #000; padding:8px;" id="pdf_storage">-</td>
              <td style="border:1px solid #000; padding:8px;" id="pdf_avg">-</td>
              <td style="border:1px solid #000; padding:8px; font-weight:bold;" id="pdf_box">-</td>
              <td style="border:1px solid #000; padding:8px; font-weight:bold;" id="pdf_weight">-</td>
              <td style="border:1px solid #000; padding:8px;" id="pdf_exp">-</td>
              <td style="border:1px solid #000; padding:8px; font-family:monospace; font-weight:bold;" id="pdf_grid">-</td>
            </tr>
          </tbody>
        </table>

        <div style="border:1px solid #000; padding:12px; margin-bottom:30px;">
          <strong>[출하 시 주의사항]</strong><br/>
          1. 상기 지정된 이력번호 및 B/L No., GRID 번호의 재고 로트에서 정확히 피킹하여 상차 바랍니다.<br/>
          2. 냉장/냉동 탑차 온도(-18℃ 이하 / 냉장 0~2℃) 확인 후 상차 요망.<br/>
          3. 상차 완료 즉시 계근표 및 인수증을 본사 FAX 또는 전산으로 회신 부탁드립니다.
        </div>

        <div style="display:flex; justify-content:space-between; align-items:flex-end; padding:0 20px;">
          <div>작성일시: <span id="pdf_today"></span></div>
          <div style="font-size:14px; font-weight:bold;">담당자 확인 : ________________ (인)</div>
        </div>
      </div>
    </div>
  </div>

  <!-- 7. 클레임 등록 모달 -->
  <div class="modal-overlay" id="claimModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px; color:#dc2626;" id="claimModalTitle">클레임 등록 및 전표 이관</h3>
      <input type="hidden" id="claim_type" />
      <input type="hidden" id="claim_id" />
      <div class="form-group">
        <label>클레임 처리일자 (KST 기준)</label>
        <input type="date" id="claim_date" class="form-control" />
      </div>
      <div class="form-group">
        <label>클레임 발생 사유</label>
        <select id="claim_reason_select" class="form-control" onchange="if(this.value!='직접입력'){document.getElementById('claim_reason_text').value=this.value;}">
          <option value="표면 갈변 발생 및 변질">표면 갈변 발생 및 변질</option>
          <option value="중량 부족 / 감량 오차">중량 부족 / 감량 오차</option>
          <option value="포장 파손 및 진공 풀림">포장 파손 및 진공 풀림</option>
          <option value="소비기한 경과 / 임박 반품">소비기한 경과 / 임박 반품</option>
          <option value="직접입력">직접입력</option>
        </select>
      </div>
      <div class="form-group">
        <label>상세 사유 기록</label>
        <input type="text" id="claim_reason_text" class="form-control" value="표면 갈변 발생 및 변질" placeholder="상세 사유를 입력하세요" />
      </div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal('claimModal')">취소</button>
        <button class="btn btn-danger" onclick="submitClaim()">클레임 이관 확정</button>
      </div>
    </div>
  </div>

  <!-- 8. 수정 모달 -->
  <div class="modal-overlay" id="editModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px;" id="editModalTitle">정보 수정</h3>
      <input type="hidden" id="edit_type" />
      <input type="hidden" id="edit_id" />
      <div class="form-group"><label id="edit_date_label">일자</label><input type="date" id="edit_date" class="form-control" /></div>
      <div id="edit_inbound_fields" style="display:none;">
        <div style="display:flex; gap:10px;">
          <div class="form-group" style="flex:1;"><label>B/L No.</label><input type="text" id="edit_bl_no" class="form-control" /></div>
          <div class="form-group" style="flex:1;"><label>가공일자(FROM)</label><input type="date" id="edit_process_from" class="form-control" /></div>
        </div>
        <div style="display:flex; gap:10px;">
          <div class="form-group" style="flex:1;"><label>브랜드</label><input type="text" id="edit_brand" class="form-control" /></div>
          <div class="form-group" style="flex:1;"><label>품목명</label><input type="text" id="edit_item_name" class="form-control" /></div>
        </div>
        <div style="display:flex; gap:10px;">
          <div class="form-group" style="flex:1;"><label>부위명</label><input type="text" id="edit_cut_name" class="form-control" /></div>
          <div class="form-group" style="flex:1;"><label>보관창고</label><select id="edit_warehouse" class="form-control"></select></div>
        </div>
        <div class="form-group"><label>소비기한</label><input type="date" id="edit_exp_date" class="form-control" /></div>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>박스 수량(Box)</label><input type="number" id="edit_box" class="form-control" /></div>
        <div class="form-group" style="flex:1;"><label>실중량(kg)</label><input type="number" step="0.01" id="edit_weight" class="form-control" /></div>
      </div>
      <div class="form-group"><label id="edit_price_label">단가(원/kg)</label><input type="number" id="edit_price" class="form-control" /></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal('editModal')">취소</button>
        <button class="btn btn-primary" onclick="submitEdit()">수정 반영</button>
      </div>
    </div>
  </div>

  <!-- 9. 재고 감량 모달 -->
  <div class="modal-overlay" id="adjustModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px; color:#dc2626;">재고 감량 및 폐기 조정</h3>
      <input type="hidden" id="adj_lot_id" />
      <div class="form-group"><label>조정 사유 구분</label>
        <select id="adj_type" class="form-control">
          <option value="갈변/변질폐기">갈변/변질 폐기</option>
          <option value="감량/드립로스">자연 감량 / 드립 로스</option>
          <option value="실사조정">재고 실사 차이 조정</option>
        </select>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>차감 박스</label><input type="number" id="adj_box" class="form-control" value="1" /></div>
        <div class="form-group" style="flex:1;"><label>차감 중량(kg)</label><input type="number" step="0.1" id="adj_weight" class="form-control" value="20.4" /></div>
      </div>
      <div class="form-group"><label>상세 사유</label><input type="text" id="adj_reason" class="form-control" placeholder="예: 표면 갈변 발생 폐기" /></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal('adjustModal')">취소</button>
        <button class="btn btn-danger" onclick="submitAdjust()">조정 적용</button>
      </div>
    </div>
  </div>

  <input type="file" id="excelFileInput" style="display:none;" onchange="uploadExcelFile(event)" accept=".xlsx, .xls" />

  <script>
    let currentMain = 'INBOUND', currentSub = 'IN_REQUEST';
    let partnerSubTab = 'ALL';
    let masterSubTab = 'ITEM_LIST';
    let claimSubTab = 'ALL';
    let resSubTab = 'HOLD'; // HOLD | CANCELLED

    const nowKST = new Date();
    const todayStr = nowKST.toISOString().split('T')[0];
    document.getElementById('in_date').value = todayStr;
    document.getElementById('in_process_from').value = todayStr;

    let targetExactDate = todayStr;
    let activeMonthFilter = null;
    document.getElementById('filter_start').value = todayStr;
    document.getElementById('filter_end').value = todayStr;

    // 미트와치 가공일(FROM) 기준 소비기한 계산 함수 (D-1 적용)
    function autoCalculateExpDate() {
      const fromDateVal = document.getElementById('in_process_from').value || document.getElementById('in_date').value || todayStr;
      const storage = document.getElementById('in_storage').value;
      const item = document.getElementById('in_item').value;

      const baseDate = new Date(fromDateVal);
      if (isNaN(baseDate.getTime())) return;

      let addDays = 60; // 기본 냉장돈육 60일
      if (storage === '냉동') {
        addDays = 730; // 냉동 730일
      } else {
        if (item.includes('우육') || item.includes('소')) {
          addDays = 90; // 냉장우육 90일
        } else if (item.includes('돈육') || item.includes('돼지')) {
          addDays = 60; // 냉장돈육 60일
        } else {
          addDays = 30;
        }
      }

      // D-1 일자 감산 적용
      baseDate.setDate(baseDate.getDate() + addDays - 1);
      document.getElementById('in_exp').value = baseDate.toISOString().split('T')[0];
    }

    // 잔존일수 계산기
    function getRemainingDays(expDateStr) {
      if (!expDateStr) return 9999;
      const today = new Date(todayStr);
      const exp = new Date(expDateStr);
      const diffTime = exp.getTime() - today.getTime();
      return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    }

    function getRemainingDaysHtml(expDateStr) {
      if (!expDateStr) return '-';
      const diffDays = getRemainingDays(expDateStr);

      if (diffDays < 0) {
        return `${expDateStr} <span class="remain-badge remain-warning">만료 (D+${Math.abs(diffDays)})</span>`;
      } else if (diffDays <= 30) {
        return `${expDateStr} <span class="remain-badge remain-warning">D-${diffDays}일</span>`;
      } else {
        return `${expDateStr} <span class="remain-badge remain-normal">D-${diffDays}일</span>`;
      }
    }

    async function loadWarehouseOptions() {
      try {
        const res = await fetch('/api/partners?type=WAREHOUSE');
        const list = await res.json();
        const select = document.getElementById('in_warehouse');
        const editSelect = document.getElementById('edit_warehouse');
        
        let html = '';
        if (list.length > 0) {
          html = list.map(w => `<option value="${w.name}">${w.name}</option>`).join('');
        } else {
          html = '<option value="광주냉장창고">광주냉장창고 (기본)</option>';
        }
        if (select) select.innerHTML = html;
        if (editSelect) editSelect.innerHTML = html;
      } catch (e) {
        console.error("창고 목록 로드 실패", e);
      }
    }

    function evaluateMultiCondition(targetStr, rawQuery) {
      if (!rawQuery) return true;
      targetStr = (targetStr || '').toLowerCase();
      rawQuery = rawQuery.trim().toLowerCase();

      if (rawQuery.includes(' or ')) {
        const orTokens = rawQuery.split(' or ').map(t => t.trim()).filter(Boolean);
        return orTokens.some(token => evaluateMultiCondition(targetStr, token));
      }

      if (rawQuery.includes(',')) {
        const inTokens = rawQuery.split(',').map(t => t.trim()).filter(Boolean);
        return inTokens.some(token => targetStr.includes(token));
      }

      const andTokens = rawQuery.split(/ and | /).map(t => t.trim()).filter(Boolean);
      return andTokens.every(token => targetStr.includes(token));
    }

    function setDateFilter(mode) {
      ['btn_today', 'btn_yesterday', 'btn_this_month', 'btn_all'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.classList.remove('active');
      });

      const startInput = document.getElementById('filter_start');
      const endInput = document.getElementById('filter_end');
      targetExactDate = null;
      activeMonthFilter = null;

      if (mode === 'TODAY') {
        document.getElementById('btn_today').classList.add('active');
        targetExactDate = todayStr;
        startInput.value = todayStr;
        endInput.value = todayStr;
      } else if (mode === 'YESTERDAY') {
        document.getElementById('btn_yesterday').classList.add('active');
        const yest = new Date();
        yest.setDate(yest.getDate() - 1);
        const yestStr = yest.toISOString().split('T')[0];
        targetExactDate = yestStr;
        startInput.value = yestStr;
        endInput.value = yestStr;
      } else if (mode === 'THIS_MONTH') {
        document.getElementById('btn_this_month').classList.add('active');
        const y = nowKST.getFullYear();
        const m = String(nowKST.getMonth() + 1).padStart(2, '0');
        activeMonthFilter = `${y}-${m}`;
        startInput.value = `${y}-${m}-01`;
        endInput.value = new Date(y, nowKST.getMonth() + 1, 0).toISOString().split('T')[0];
      } else if (mode === 'ALL') {
        document.getElementById('btn_all').classList.add('active');
        startInput.value = '';
        endInput.value = '';
      }
      loadData();
    }

    function switchMainTab(tab, el) {
      currentMain = tab;
      document.querySelectorAll('.nav-item').forEach(e => e.classList.remove('active'));
      if (el) el.classList.add('active');

      const subTabs = document.getElementById('subTabs');
      const headerActions = document.getElementById('headerActions');
      const pageTitle = document.getElementById('pageTitle');
      const dateCard = document.getElementById('dateFilterCard');
      const summaryBar = document.getElementById('summaryBar');

      if (tab === 'INBOUND') {
        currentSub = 'IN_REQUEST';
        pageTitle.innerText = '입고 프로세스 관리';
        subTabs.style.display = 'flex';
        dateCard.style.display = 'flex';
        summaryBar.style.display = 'flex';
        document.getElementById('dateFilterLabel').innerHTML = '<i class="bi bi-calendar-range"></i> 입고 일자:';
        headerActions.innerHTML = `
          <button class="btn btn-outline" onclick="document.getElementById('excelFileInput').click()"><i class="bi bi-file-earmark-excel"></i> 엑셀 일괄입고</button>
          <button class="btn btn-primary" onclick="openInboundCreateModal()"><i class="bi bi-plus-lg"></i> + 신규 입고등록</button>
        `;
        renderSubTabs();
      } else if (tab === 'OUTBOUND') {
        currentSub = 'OUT_STOCK';
        pageTitle.innerText = '출고 프로세스 관리 (현 재고장 연동)';
        subTabs.style.display = 'flex';
        dateCard.style.display = 'flex';
        summaryBar.style.display = 'flex';
        document.getElementById('dateFilterLabel').innerHTML = '<i class="bi bi-calendar-range"></i> 출고 일자:';
        headerActions.innerHTML = '';
        renderSubTabs();
      } else if (tab === 'RESERVATION') {
        pageTitle.innerText = '출고 예약 관리 (영업사원 홀딩)';
        subTabs.style.display = 'flex';
        dateCard.style.display = 'none';
        summaryBar.style.display = 'flex';
        headerActions.innerHTML = '';
        resSubTab = 'HOLD';
        renderReservationSubTabs();
        return;
      } else if (tab === 'CLAIM') {
        pageTitle.innerText = '통합 클레임 관리 (입고 / 출고 구분)';
        subTabs.style.display = 'flex';
        dateCard.style.display = 'flex';
        summaryBar.style.display = 'flex';
        document.getElementById('dateFilterLabel').innerHTML = '<i class="bi bi-calendar-range"></i> 처리 일자:';
        headerActions.innerHTML = '';
        claimSubTab = 'ALL';
        renderClaimSubTabs();
        return;
      } else if (tab === 'PARTNER') {
        pageTitle.innerText = '거래처 정보 관리 (대분류 구분)';
        subTabs.style.display = 'flex';
        dateCard.style.display = 'none';
        summaryBar.style.display = 'none';
        headerActions.innerHTML = `<button class="btn btn-primary" onclick="openModal('partnerModal')"><i class="bi bi-plus-lg"></i> 신규 거래처 등록</button>`;
        renderPartnerSubTabs();
        return;
      } else if (tab === 'ITEM_CUT_MASTER') {
        pageTitle.innerText = '품목 및 부위 마스터 계층 관리';
        subTabs.style.display = 'flex';
        dateCard.style.display = 'none';
        summaryBar.style.display = 'none';
        masterSubTab = 'ITEM_LIST';
        renderMasterSubTabs();
        return;
      }
      loadData();
    }

    function renderSubTabs() {
      const container = document.getElementById('subTabs');
      container.innerHTML = '';
      if (currentMain === 'INBOUND') {
        const steps = [
          {k:'IN_REQUEST', n:'1. 입고요청리스트'},
          {k:'IN_CONFIRM', n:'2. 입고확정리스트'},
          {k:'IN_DONE', n:'3. 입고완료리스트'}
        ];
        steps.forEach(s => {
          container.innerHTML += `<div class="sub-tab ${currentSub===s.k?'active':''}" onclick="setSubTab('${s.k}')">${s.n}</div>`;
        });
        container.innerHTML += `<div class="sub-tab sub-tab-create" onclick="openInboundCreateModal()"><i class="bi bi-plus-circle-fill"></i> + 신규 입고등록</div>`;
      } else if (currentMain === 'OUTBOUND') {
        const steps = [
          {k:'OUT_STOCK', n:'1. 출하요청리스트 (현 재고장)'},
          {k:'OUT_REQUEST', n:'2. 출고요청리스트'},
          {k:'OUT_CONFIRM', n:'3. 출고확정리스트'},
          {k:'OUT_DONE', n:'4. 출고완료리스트'}
        ];
        steps.forEach(s => {
          container.innerHTML += `<div class="sub-tab ${currentSub===s.k?'active':''}" onclick="setSubTab('${s.k}')">${s.n}</div>`;
        });
      }
    }

    function renderReservationSubTabs() {
      const container = document.getElementById('subTabs');
      container.innerHTML = `
        <div class="sub-tab ${resSubTab==='HOLD'?'active':''}" onclick="setResSubTab('HOLD')">1. 예약요청리스트 (홀딩)</div>
        <div class="sub-tab ${resSubTab==='CANCELLED'?'active':''}" onclick="setResSubTab('CANCELLED')">2. 예약요청취소리스트 (만료/취소 누적)</div>
      `;
      loadData();
    }
    function setResSubTab(t) { resSubTab = t; renderReservationSubTabs(); }

    function renderClaimSubTabs() {
      const container = document.getElementById('subTabs');
      const tabs = [
        {k:'ALL', n:'전체 클레임 리스트'},
        {k:'INBOUND', n:'입고 클레임'},
        {k:'OUTBOUND', n:'출고 클레임'}
      ];
      container.innerHTML = tabs.map(t => `<div class="sub-tab ${claimSubTab===t.k?'active':''}" onclick="setClaimSubTab('${t.k}')">${t.n}</div>`).join('');
      loadData();
    }
    function setClaimSubTab(t) { claimSubTab = t; renderClaimSubTabs(); }

    function renderPartnerSubTabs() {
      const container = document.getElementById('subTabs');
      const tabs = [{k:'ALL',n:'전체 거래처'},{k:'VENDOR',n:'매입처'},{k:'CUSTOMER',n:'매출처'},{k:'WAREHOUSE',n:'보관창고'},{k:'SHIPPING',n:'선사/포워딩'},{k:'ETC',n:'기타'}];
      container.innerHTML = tabs.map(t => `<div class="sub-tab ${partnerSubTab===t.k?'active':''}" onclick="setPartnerSubTab('${t.k}')">${t.n}</div>`).join('');
      loadData();
    }
    function setPartnerSubTab(t) { partnerSubTab = t; renderPartnerSubTabs(); }

    function renderMasterSubTabs() {
      const container = document.getElementById('subTabs');
      container.innerHTML = `
        <div class="sub-tab ${masterSubTab==='ITEM_LIST'?'active':''}" onclick="setMasterSubTab('ITEM_LIST')">1. 상위 품목(대분류) 관리</div>
        <div class="sub-tab ${masterSubTab==='CUT_LIST'?'active':''}" onclick="setMasterSubTab('CUT_LIST')">2. 세부 부위(하위분류) 관리</div>
      `;
      const headerActions = document.getElementById('headerActions');
      if (masterSubTab === 'ITEM_LIST') {
        headerActions.innerHTML = `<button class="btn btn-primary" onclick="openModal('itemMasterModal')"><i class="bi bi-plus-lg"></i> 신규 품목 등록</button>`;
      } else {
        headerActions.innerHTML = `<button class="btn btn-primary" onclick="openCutModal()"><i class="bi bi-plus-lg"></i> 신규 부위 등록</button>`;
      }
      loadData();
    }
    function setMasterSubTab(tab) { masterSubTab = tab; renderMasterSubTabs(); }
    function setSubTab(k) { currentSub = k; renderSubTabs(); loadData(); }

    function downloadCurrentTableExcel() {
      const table = document.querySelector('.table-card table');
      if (!table) return;
      const cloneTable = table.cloneNode(true);
      const wb = XLSX.utils.table_to_book(cloneTable, { sheet: "Sheet1" });
      const tabName = currentMain === 'INBOUND' ? `입고_${currentSub}` : currentMain === 'OUTBOUND' ? `출고_${currentSub}` : currentMain === 'RESERVATION' ? `예약_${resSubTab}` : currentMain === 'CLAIM' ? `클레임_${claimSubTab}` : currentMain;
      const dateStr = new Date().toISOString().split('T')[0];
      XLSX.writeFile(wb, `MeatERP_${tabName}_${dateStr}.xlsx`);
    }

    async function loadData() {
      const head = document.getElementById('tableHead'), body = document.getElementById('tableBody');
      const keyword = (document.getElementById('filter_keyword') ? document.getElementById('filter_keyword').value : '').trim();
      body.innerHTML = '<tr><td colspan="18" style="text-align:center; padding:15px;">데이터 조회 중...</td></tr>';

      const sDate = document.getElementById('filter_start').value;
      const eDate = document.getElementById('filter_end').value;
      let dateQuery = '';

      if (targetExactDate) {
        dateQuery = `&target_date=${targetExactDate}`;
      } else if (activeMonthFilter) {
        dateQuery = `&month=${activeMonthFilter}`;
      } else {
        if (sDate) dateQuery += `&start_date=${sDate}`;
        if (eDate) dateQuery += `&end_date=${eDate}`;
      }

      // 1. 입고 관리 뷰
      if (currentMain === 'INBOUND') {
        head.innerHTML = `
          <tr>
            <th>전표번호</th>
            <th>입고일자</th>
            <th>보관창고</th>
            <th>매입처</th>
            <th>B/L No.</th>
            <th>이력번호</th>
            <th>가공일(FROM)</th>
            <th>브랜드</th>
            <th>품목명</th>
            <th>부위명</th>
            <th>보관</th>
            <th style="text-align:right;">박스</th>
            <th style="text-align:right;">실중량(kg)</th>
            <th style="text-align:right;">단가(kg)</th>
            <th style="text-align:right;">총 매입금액</th>
            <th>소비기한 (잔존일)</th>
            <th style="text-align:center;">작업 관리</th>
            <th>GRID코드</th>
          </tr>
        `;
        const res = await fetch(`/api/inbounds?status=${currentSub}${dateQuery}`);
        let list = await res.json();

        if (keyword) {
          list = list.filter(i => {
            const rowStr = `${i.item_name} ${i.cut_name} ${i.brand} ${i.vendor} ${i.bl_no} ${i.trace_no} ${i.warehouse} ${i.grid_no} ${i.storage_type} ${i.process_from_date}`;
            return evaluateMultiCondition(rowStr, keyword);
          });
        }

        let sumBox = 0, sumWeight = 0.0, sumAmount = 0;
        list.forEach(i => {
          sumBox += i.box_qty;
          sumWeight += i.weight_kg;
          sumAmount += i.total_amount;
        });

        document.getElementById('sum_count').innerText = list.length.toLocaleString();
        document.getElementById('sum_box').innerText = sumBox.toLocaleString();
        document.getElementById('sum_weight').innerText = sumWeight.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1});
        document.getElementById('sum_amount').innerText = Math.round(sumAmount).toLocaleString();

        body.innerHTML = list.length ? list.map(i => {
          let btns = '';
          if (currentSub === 'IN_REQUEST') {
            btns = `
              <button class="btn btn-primary" onclick="advIn(${i.id})">확정 <i class="bi bi-arrow-right"></i></button>
              <button class="btn btn-outline" onclick="openEditModal('INBOUND', ${i.id}, '${i.inbound_date}', '${i.bl_no||''}', '${i.brand||''}', '${i.item_name}', '${i.cut_name}', '${i.warehouse||''}', ${i.box_qty}, ${i.weight_kg}, ${i.cost_per_kg}, '${i.is_weighed}', '${i.exp_date}', '${i.process_from_date||''}')"><i class="bi bi-pencil"></i></button>
              <button class="btn btn-outline" style="color:#dc2626;" onclick="openClaimModal('INBOUND', ${i.id})"><i class="bi bi-exclamation-triangle"></i> 클레임</button>
              <button class="btn btn-outline" style="color:#dc2626;" onclick="revertIn(${i.id}, '요청 취소(삭제)')"><i class="bi bi-trash"></i></button>
            `;
          } else if (currentSub === 'IN_CONFIRM') {
            btns = `
              <button class="btn btn-success" onclick="advIn(${i.id})"><i class="bi bi-check-lg"></i> 입고완료</button>
              <button class="btn btn-outline" onclick="openEditModal('INBOUND', ${i.id}, '${i.inbound_date}', '${i.bl_no||''}', '${i.brand||''}', '${i.item_name}', '${i.cut_name}', '${i.warehouse||''}', ${i.box_qty}, ${i.weight_kg}, ${i.cost_per_kg}, '${i.is_weighed}', '${i.exp_date}', '${i.process_from_date||''}')"><i class="bi bi-pencil"></i></button>
              <button class="btn btn-outline" style="color:#dc2626;" onclick="openClaimModal('INBOUND', ${i.id})"><i class="bi bi-exclamation-triangle"></i> 클레임</button>
              <button class="btn btn-outline" style="color:#d97706;" onclick="revertIn(${i.id}, '이전 단계 반려')"><i class="bi bi-arrow-counterclockwise"></i> 반려</button>
            `;
          } else {
            btns = `
              <span class="badge" style="background:#dcfce7; color:#15803d; margin-right:4px;">완료됨</span>
              <button class="btn btn-outline" style="color:#dc2626;" onclick="openClaimModal('INBOUND', ${i.id})"><i class="bi bi-exclamation-triangle"></i> 클레임</button>
              <button class="btn btn-outline" style="color:#dc2626;" onclick="revertIn(${i.id}, '입고 취소 및 재고 원복')"><i class="bi bi-arrow-counterclockwise"></i> 작업취소</button>
            `;
          }

          return `
            <tr>
              <td><strong>${i.inbound_no}</strong></td>
              <td>${i.inbound_date}</td>
              <td><span class="badge badge-warehouse">${i.warehouse || '광주냉장창고'}</span></td>
              <td>${i.vendor}</td>
              <td><span class="bl-tag">${i.bl_no || '-'}</span></td>
              <td><code>${i.trace_no}</code></td>
              <td>${i.process_from_date || '-'}</td>
              <td><span class="brand-tag">${i.brand || '-'}</span></td>
              <td><span class="item-tag">${i.item_name}</span></td>
              <td><strong>${i.cut_name}</strong></td>
              <td><span class="badge ${i.storage_type==='냉장'?'badge-chilled':'badge-frozen'}">${i.storage_type}</span></td>
              <td style="text-align:right;"><strong>${i.box_qty}</strong> Box</td>
              <td style="text-align:right;"><strong>${i.weight_kg.toLocaleString()}</strong> kg</td>
              <td style="text-align:right;">${i.cost_per_kg.toLocaleString()}원</td>
              <td style="text-align:right;" class="amount-highlight">${i.total_amount.toLocaleString()}원</td>
              <td>${getRemainingDaysHtml(i.exp_date)}</td>
              <td style="text-align:center; white-space:nowrap;">${btns}</td>
              <td style="text-align:center;"><span class="grid-tag">${i.grid_no || '-'}</span></td>
            </tr>
          `;
        }).join('') : '<tr><td colspan="18" style="text-align:center; padding:20px; color:#94a3b8;">조회된 입고 전표가 없습니다.</td></tr>';

      // 2. 출고 관리 뷰 (파란색 예약재고 표기 & 마우스오버 툴팁 & 탭 이동 링크)
      } else if (currentMain === 'OUTBOUND') {
        if (currentSub === 'OUT_STOCK') {
          head.innerHTML = `
            <tr>
              <th>B/L No.</th>
              <th>이력번호</th>
              <th>가공일(FROM)</th>
              <th>브랜드</th>
              <th>품목/부위</th>
              <th>보관</th>
              <th style="text-align:right; background:#f1f5f9;">기준 평중(kg/Box)</th>
              <th style="text-align:right; background:#eff6ff;">최초 입고(Box/kg)</th>
              <th style="text-align:right; background:#fef2f2; color:#b91c1c;">현재고 잔여(Box/kg)</th>
              <th style="text-align:right; background:#e0f2fe; color:#0284c7;">예약재고 (홀딩)</th>
              <th style="text-align:right;">원가단가</th>
              <th style="text-align:right;">재고 평가금액</th>
              <th>창고</th>
              <th>소비기한 (잔존일)</th>
              <th style="text-align:center;">관리 작업</th>
              <th>GRID코드</th>
            </tr>
          `;
          const res = await fetch('/api/inventory');
          let list = await res.json();

          if (keyword) {
            list = list.filter(l => {
              const rowStr = `${l.item_name} ${l.cut_name} ${l.brand} ${l.bl_no} ${l.trace_no} ${l.warehouse} ${l.grid_no} ${l.storage_type} ${l.process_from_date}`;
              return evaluateMultiCondition(rowStr, keyword);
            });
          }

          let sumBox = 0, sumWeight = 0.0, sumAmount = 0;
          list.forEach(l => {
            sumBox += l.current_box_qty;
            sumWeight += l.current_weight_kg;
            sumAmount += (l.current_weight_kg * l.cost_per_kg);
          });

          document.getElementById('sum_count').innerText = list.length.toLocaleString();
          document.getElementById('sum_box').innerText = sumBox.toLocaleString();
          document.getElementById('sum_weight').innerText = sumWeight.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1});
          document.getElementById('sum_amount').innerText = Math.round(sumAmount).toLocaleString();

          body.innerHTML = list.length ? list.map(l => {
            const evalAmount = Math.round(l.current_weight_kg * l.cost_per_kg);
            
            // 파란색 글씨 예약재고 링크 & 마우스오버 툴팁(title)
            let reserveHtml = '<span style="color:#94a3b8;">0 Box</span>';
            if (l.reserved_box_qty > 0) {
              reserveHtml = `
                <span class="res-link" title="예약 업체 내역: ${l.reserved_customers}" onclick="switchMainTab('RESERVATION', document.querySelectorAll('.nav-item')[2])">
                  ${l.reserved_box_qty} Box / ${l.reserved_weight_kg} kg
                </span>
              `;
            }

            return `
              <tr>
                <td><span class="bl-tag">${l.bl_no || '-'}</span></td>
                <td><code>${l.trace_no}</code></td>
                <td>${l.process_from_date || '-'}</td>
                <td><span class="brand-tag">${l.brand || '-'}</span></td>
                <td><span class="item-tag">${l.item_name}</span> <strong>${l.cut_name}</strong></td>
                <td><span class="badge ${l.storage_type==='냉장'?'badge-chilled':'badge-frozen'}">${l.storage_type}</span></td>
                <td style="text-align:right; background:#f1f5f9; font-weight:bold; color:#0284c7;">${l.avg_box_weight} kg/Box</td>
                <td style="text-align:right; background:#eff6ff;">${l.initial_box_qty} Box / ${l.initial_weight_kg.toLocaleString()} kg</td>
                <td style="text-align:right; background:#fef2f2; font-weight:700; color:#b91c1c;">${l.current_box_qty} Box / ${l.current_weight_kg.toLocaleString()} kg</td>
                <td style="text-align:right; background:#e0f2fe;">${reserveHtml}</td>
                <td style="text-align:right;">${l.cost_per_kg.toLocaleString()}원</td>
                <td style="text-align:right;" class="amount-highlight">${evalAmount.toLocaleString()}원</td>
                <td>${l.warehouse}</td>
                <td>${getRemainingDaysHtml(l.exp_date)}</td>
                <td style="text-align:center; white-space:nowrap;">
                  <button class="btn btn-warning" onclick="reqOutWithWarning(${l.id},'${l.grid_no}',${l.current_box_qty},${l.avg_box_weight},${l.cost_per_kg},'${l.storage_type}','${l.item_name}','${l.exp_date}')"><i class="bi bi-cart-plus"></i> 출하요청</button>
                  <button class="btn btn-outline" style="color:#0284c7;" onclick="openReserveModal(${l.id},${l.current_box_qty},${l.cost_per_kg})"><i class="bi bi-calendar-plus"></i> 예약등록</button>
                  <button class="btn btn-outline" style="color:#dc2626;" onclick="openAdjustModal(${l.id})"><i class="bi bi-scissors"></i> 감량/폐기</button>
                </td>
                <td style="text-align:center;"><span class="grid-tag">${l.grid_no}</span></td>
              </tr>
            `;
          }).join('') : '<tr><td colspan="16" style="text-align:center; padding:20px; color:#94a3b8;">보유 재고가 없습니다.</td></tr>';
        } else {
          head.innerHTML = `
            <tr>
              <th>출고번호</th>
              <th>출고일자</th>
              <th>매출처</th>
              <th>보관창고</th>
              <th>B/L No.</th>
              <th>이력번호</th>
              <th>가공일(FROM)</th>
              <th>브랜드</th>
              <th>품목/부위</th>
              <th style="text-align:right;">기준평중</th>
              <th style="text-align:right;">출고 박스</th>
              <th style="text-align:right;">출고 실중량(kg)</th>
              <th style="text-align:right;">판매단가</th>
              <th style="text-align:right;">총 매출금액</th>
              <th>소비기한 (잔존일)</th>
              <th style="text-align:center;">작업 / 팩스발송</th>
              <th>GRID코드</th>
            </tr>
          `;
          const res = await fetch(`/api/outbounds?status=${currentSub}${dateQuery}`);
          let list = await res.json();

          if (keyword) {
            list = list.filter(o => {
              const rowStr = `${o.item_name} ${o.cut_name} ${o.customer} ${o.brand} ${o.bl_no} ${o.trace_no} ${o.warehouse} ${o.grid_no} ${o.process_from_date}`;
              return evaluateMultiCondition(rowStr, keyword);
            });
          }

          let sumBox = 0, sumWeight = 0.0, sumAmount = 0;
          list.forEach(o => {
            sumBox += o.box_qty;
            sumWeight += o.weight_kg;
            sumAmount += o.total_amount;
          });

          document.getElementById('sum_count').innerText = list.length.toLocaleString();
          document.getElementById('sum_box').innerText = sumBox.toLocaleString();
          document.getElementById('sum_weight').innerText = sumWeight.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1});
          document.getElementById('sum_amount').innerText = Math.round(sumAmount).toLocaleString();

          body.innerHTML = list.length ? list.map(o => {
            let btns = '';
            if (currentSub === 'OUT_REQUEST') {
              btns = `
                <button class="btn btn-primary" onclick="advOut(${o.id})">확정 <i class="bi bi-arrow-right"></i></button>
                <button class="btn btn-outline" onclick="openEditModal('OUTBOUND', ${o.id}, '${o.outbound_date}', '', '', '', '', '', ${o.box_qty}, ${o.weight_kg}, ${o.unit_price_kg}, '', '')"><i class="bi bi-pencil"></i></button>
                <button class="btn btn-outline" style="color:#dc2626;" onclick="openClaimModal('OUTBOUND', ${o.id})"><i class="bi bi-exclamation-triangle"></i> 클레임</button>
                <button class="btn btn-outline" style="color:#dc2626;" onclick="revertOut(${o.id}, '요청 취소(재고복구)')"><i class="bi bi-arrow-counterclockwise"></i> 요청취소</button>
              `;
            } else if (currentSub === 'OUT_CONFIRM') {
              btns = `
                <button class="btn btn-success" onclick="advOut(${o.id})"><i class="bi bi-truck"></i> 출고완료</button>
                <button class="btn btn-outline" onclick="openEditModal('OUTBOUND', ${o.id}, '${o.outbound_date}', '', '', '', '', '', ${o.box_qty}, ${o.weight_kg}, ${o.unit_price_kg}, '', '')"><i class="bi bi-pencil"></i></button>
                <button class="btn btn-outline" style="color:#dc2626;" onclick="openClaimModal('OUTBOUND', ${o.id})"><i class="bi bi-exclamation-triangle"></i> 클레임</button>
                <button class="btn btn-outline" style="color:#d97706;" onclick="revertOut(${o.id}, '이전 단계 반려')"><i class="bi bi-arrow-counterclockwise"></i> 반려</button>
              `;
            } else {
              btns = `
                <span class="badge" style="background:#dbeafe; color:#1d4ed8; margin-right:4px;">출고완료</span>
                <button class="btn btn-outline" style="color:#dc2626;" onclick="openClaimModal('OUTBOUND', ${o.id})"><i class="bi bi-exclamation-triangle"></i> 클레임</button>
                <button class="btn btn-outline" style="color:#dc2626;" onclick="revertOut(${o.id}, '출고완료 취소')"><i class="bi bi-arrow-counterclockwise"></i> 작업취소</button>
              `;
            }

            return `
              <tr>
                <td><strong>${o.outbound_no}</strong></td>
                <td>${o.outbound_date}</td>
                <td><strong>${o.customer}</strong></td>
                <td><span class="badge badge-warehouse">${o.warehouse || '광주냉장창고'}</span></td>
                <td><span class="bl-tag">${o.bl_no || '-'}</span></td>
                <td><code>${o.trace_no}</code></td>
                <td>${o.process_from_date || '-'}</td>
                <td><span class="brand-tag">${o.brand || '-'}</span></td>
                <td><span class="item-tag">${o.item_name}</span> <strong>${o.cut_name}</strong></td>
                <td style="text-align:right; color:#0284c7; font-weight:600;">${o.avg_box_weight} kg/Box</td>
                <td style="text-align:right;"><strong>${o.box_qty}</strong> Box</td>
                <td style="text-align:right; font-weight:700; color:#b91c1c;">${o.weight_kg.toLocaleString()} kg</td>
                <td style="text-align:right;">${o.unit_price_kg.toLocaleString()}원</td>
                <td style="text-align:right;" class="amount-highlight">${o.total_amount.toLocaleString()}원</td>
                <td>${getRemainingDaysHtml(o.exp_date)}</td>
                <td style="text-align:center; white-space:nowrap;">
                  <button class="btn btn-outline" style="color:#dc2626;" onclick='openPdfPreview(${JSON.stringify(o)})'><i class="bi bi-file-earmark-pdf-fill"></i> 출고요청서(팩스)</button>
                  ${btns}
                </td>
                <td style="text-align:center;"><span class="grid-tag">${o.grid_no || '-'}</span></td>
              </tr>
            `;
          }).join('') : '<tr><td colspan="17" style="text-align:center; padding:20px; color:#94a3b8;">대기 중인 출고 건이 없습니다.</td></tr>';
        }

      // 3. 예약 관리 뷰 (영업사원 컬럼 포함, 만료/취소 누적)
      } else if (currentMain === 'RESERVATION') {
        if (resSubTab === 'HOLD') {
          head.innerHTML = `
            <tr>
              <th>예약번호</th>
              <th>영업사원</th>
              <th>예약 매출처</th>
              <th>GRID코드</th>
              <th>품목/부위</th>
              <th>보관</th>
              <th style="text-align:right;">예약 박스</th>
              <th style="text-align:right;">예약 실중량(kg)</th>
              <th style="text-align:right;">예정 단가</th>
              <th style="text-align:right;">총 예정금액</th>
              <th>소비기한 (잔존일)</th>
              <th style="color:#b91c1c;">예약 만료일</th>
              <th style="text-align:center;">관리 작업</th>
            </tr>
          `;
          const res = await fetch('/api/reservations?status=HOLD');
          let list = await res.json();

          let sumBox = 0, sumWeight = 0.0, sumAmount = 0;
          list.forEach(r => { sumBox += r.box_qty; sumWeight += r.weight_kg; sumAmount += r.total_amount; });
          document.getElementById('sum_count').innerText = list.length.toLocaleString();
          document.getElementById('sum_box').innerText = sumBox.toLocaleString();
          document.getElementById('sum_weight').innerText = sumWeight.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1});
          document.getElementById('sum_amount').innerText = Math.round(sumAmount).toLocaleString();

          body.innerHTML = list.length ? list.map(r => `
            <tr>
              <td><strong>${r.res_no}</strong></td>
              <td><span class="badge" style="background:#e0e7ff; color:#3730a3; font-weight:700;">${r.sales_rep}</span></td>
              <td><strong>${r.customer}</strong></td>
              <td><span class="grid-tag">${r.grid_no}</span></td>
              <td><span class="item-tag">${r.item_name}</span> <strong>${r.cut_name}</strong></td>
              <td><span class="badge ${r.storage_type==='냉장'?'badge-chilled':'badge-frozen'}">${r.storage_type}</span></td>
              <td style="text-align:right;"><strong>${r.box_qty}</strong> Box</td>
              <td style="text-align:right;"><strong>${r.weight_kg.toLocaleString()}</strong> kg</td>
              <td style="text-align:right;">${r.unit_price_kg.toLocaleString()}원</td>
              <td style="text-align:right;" class="amount-highlight">${r.total_amount.toLocaleString()}원</td>
              <td>${getRemainingDaysHtml(r.exp_date)}</td>
              <td style="color:#b91c1c; font-weight:700;">${r.expire_date}</td>
              <td style="text-align:center; white-space:nowrap;">
                <button class="btn btn-outline" style="color:#dc2626;" onclick="cancelReservationPrompt(${r.id})"><i class="bi bi-x-circle"></i> 예약취소</button>
              </td>
            </tr>
          `).join('') : '<tr><td colspan="13" style="text-align:center; padding:20px; color:#94a3b8;">현재 대기 중인 예약 건이 없습니다.</td></tr>';
        } else {
          // CANCELLED
          head.innerHTML = `
            <tr>
              <th>취소일자</th>
              <th>취소 구분</th>
              <th>원예약번호</th>
              <th>영업사원</th>
              <th>매출처</th>
              <th>GRID코드</th>
              <th>품목/부위</th>
              <th style="text-align:right;">취소 박스</th>
              <th style="text-align:right;">취소 실중량(kg)</th>
              <th>취소/만료 상세 사유</th>
            </tr>
          `;
          const res = await fetch('/api/reservations?status=CANCELLED');
          let list = await res.json();

          let sumBox = 0, sumWeight = 0.0, sumAmount = 0;
          list.forEach(r => { sumBox += r.box_qty; sumWeight += r.weight_kg; sumAmount += r.total_amount; });
          document.getElementById('sum_count').innerText = list.length.toLocaleString();
          document.getElementById('sum_box').innerText = sumBox.toLocaleString();
          document.getElementById('sum_weight').innerText = sumWeight.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1});
          document.getElementById('sum_amount').innerText = Math.round(sumAmount).toLocaleString();

          body.innerHTML = list.length ? list.map(r => `
            <tr>
              <td><strong>${r.cancel_date || '-'}</strong></td>
              <td><span class="badge ${r.cancel_type==='기간만료 자동취소'?'badge-danger':'badge-etc'}" style="background:${r.cancel_type==='기간만료 자동취소'?'#fee2e2':'#f1f5f9'}; color:${r.cancel_type==='기간만료 자동취소'?'#b91c1c':'#475569'};">${r.cancel_type}</span></td>
              <td>${r.res_no}</td>
              <td><strong>${r.sales_rep}</strong></td>
              <td>${r.customer}</td>
              <td><span class="grid-tag">${r.grid_no}</span></td>
              <td><span class="item-tag">${r.item_name}</span> <strong>${r.cut_name}</strong></td>
              <td style="text-align:right;"><strong>${r.box_qty}</strong> Box</td>
              <td style="text-align:right;"><strong>${r.weight_kg.toLocaleString()}</strong> kg</td>
              <td style="color:#b91c1c; font-weight:600;">${r.cancel_reason || '사유 미기재'}</td>
            </tr>
          `).join('') : '<tr><td colspan="10" style="text-align:center; padding:20px; color:#94a3b8;">취소된 예약 이력이 없습니다.</td></tr>';
        }

      // 4. 통합 클레임 관리 뷰
      } else if (currentMain === 'CLAIM') {
        head.innerHTML = `
          <tr>
            <th style="background:#fef2f2; color:#b91c1c; text-align:center;">발생시점</th>
            <th style="background:#fef2f2; color:#b91c1c;">입고일자</th>
            <th style="background:#fef2f2; color:#b91c1c;">클레임 처리일자</th>
            <th>전표번호</th>
            <th>거래처 (매입/매출처)</th>
            <th>보관창고</th>
            <th>B/L No.</th>
            <th>이력번호</th>
            <th>브랜드</th>
            <th>품목/부위</th>
            <th style="text-align:right;">박스</th>
            <th style="text-align:right;">실중량(kg)</th>
            <th style="text-align:right;">총 금액</th>
            <th>소비기한 (잔존일)</th>
            <th>클레임 상세 사유</th>
            <th style="text-align:center;">상태 관리</th>
            <th>GRID코드</th>
          </tr>
        `;
        const res = await fetch(`/api/claims?stage=${claimSubTab}${dateQuery}`);
        let list = await res.json();

        if (keyword) {
          list = list.filter(c => {
            const rowStr = `${c.item_name} ${c.cut_name} ${c.partner_name} ${c.brand} ${c.bl_no} ${c.trace_no} ${c.warehouse} ${c.grid_no} ${c.claim_reason}`;
            return evaluateMultiCondition(rowStr, keyword);
          });
        }

        let sumBox = 0, sumWeight = 0.0, sumAmount = 0;
        list.forEach(c => {
          sumBox += c.box_qty;
          sumWeight += c.weight_kg;
          sumAmount += c.total_amount;
        });

        document.getElementById('sum_count').innerText = list.length.toLocaleString();
        document.getElementById('sum_box').innerText = sumBox.toLocaleString();
        document.getElementById('sum_weight').innerText = sumWeight.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1});
        document.getElementById('sum_amount').innerText = Math.round(sumAmount).toLocaleString();

        body.innerHTML = list.length ? list.map(c => {
          const revertFn = c.raw_type === 'INBOUND' ? `revertIn(${c.id}, '클레임 취소 및 입고완료 복구')` : `revertOut(${c.id}, '클레임 취소 및 출고완료 복구')`;
          const stageBadge = c.stage === '입고' ? `<span class="badge badge-stage-in"><i class="bi bi-box-arrow-in-down"></i> 입고</span>` : `<span class="badge badge-stage-out"><i class="bi bi-box-arrow-up"></i> 출고</span>`;

          return `
            <tr>
              <td style="text-align:center;">${stageBadge}</td>
              <td style="background:#fff1f2; font-weight:bold;">${c.inbound_date}</td>
              <td style="background:#fff1f2; font-weight:bold; color:#b91c1c;">${c.processed_date}</td>
              <td><strong>${c.doc_no}</strong></td>
              <td><strong>${c.partner_name}</strong></td>
              <td><span class="badge badge-warehouse">${c.warehouse || '광주냉장창고'}</span></td>
              <td><span class="bl-tag">${c.bl_no || '-'}</span></td>
              <td><code>${c.trace_no}</code></td>
              <td><span class="brand-tag">${c.brand || '-'}</span></td>
              <td><span class="item-tag">${c.item_name}</span> <strong>${c.cut_name}</strong></td>
              <td style="text-align:right;"><strong>${c.box_qty}</strong> Box</td>
              <td style="text-align:right;"><strong>${c.weight_kg.toLocaleString()}</strong> kg</td>
              <td style="text-align:right;" class="amount-highlight">${c.total_amount.toLocaleString()}원</td>
              <td>${getRemainingDaysHtml(c.exp_date)}</td>
              <td style="color:#b91c1c; font-weight:600;">${c.claim_reason || '사유 미기재'}</td>
              <td style="text-align:center; white-space:nowrap;">
                <span class="badge" style="background:#fee2e2; color:#b91c1c; margin-right:4px;">클레임 처리됨</span>
                <button class="btn btn-outline" style="color:#2563eb;" onclick="${revertFn}"><i class="bi bi-arrow-counterclockwise"></i> 원복</button>
              </td>
              <td style="text-align:center;"><span class="grid-tag">${c.grid_no || '-'}</span></td>
            </tr>
          `;
        }).join('') : '<tr><td colspan="17" style="text-align:center; padding:20px; color:#94a3b8;">등록된 클레임 내역이 없습니다.</td></tr>';

      // 5. 거래처 마스터
      } else if (currentMain === 'PARTNER') {
        head.innerHTML = '<tr><th>ID</th><th>대분류 구분</th><th>상호명</th><th>사업자번호</th><th>담당자</th><th>연락처</th><th>사업장/창고 주소</th><th style="text-align:center;">관리</th></tr>';
        const typeQuery = partnerSubTab === 'ALL' ? '' : `?type=${partnerSubTab}`;
        const res = await fetch(`/api/partners${typeQuery}`), list = await res.json();
        body.innerHTML = list.length ? list.map(p => {
          let badgeClass = 'badge-etc', label = '기타';
          if (p.type === 'VENDOR') { badgeClass = 'badge-vendor'; label = '매입처'; }
          else if (p.type === 'CUSTOMER') { badgeClass = 'badge-customer'; label = '매출처'; }
          else if (p.type === 'WAREHOUSE') { badgeClass = 'badge-warehouse'; label = '보관창고'; }
          else if (p.type === 'SHIPPING') { badgeClass = 'badge-shipping'; label = '선사/포워딩'; }

          return `
            <tr>
              <td>#${p.id}</td>
              <td><span class="badge ${badgeClass}">${label}</span></td>
              <td><strong>${p.name}</strong></td>
              <td>${p.biz_no || '-'}</td>
              <td>${p.contact_person || '-'}</td>
              <td>${p.phone || '-'}</td>
              <td>${p.address || '-'}</td>
              <td style="text-align:center;"><button class="btn btn-outline" style="color:#dc2626;" onclick="deletePartner(${p.id})"><i class="bi bi-trash"></i> 삭제</button></td>
            </tr>
          `;
        }).join('') : '<tr><td colspan="8" style="text-align:center; padding:20px; color:#94a3b8;">해당 분류의 거래처가 없습니다.</td></tr>';

      // 6. 품목/부위 마스터
      } else if (currentMain === 'ITEM_CUT_MASTER') {
        if (masterSubTab === 'ITEM_LIST') {
          head.innerHTML = '<tr><th>품목 ID</th><th>품목코드</th><th>품목명 (대분류)</th><th>축종</th><th style="text-align:center;">관리</th></tr>';
          const res = await fetch('/api/items'), list = await res.json();
          body.innerHTML = list.length ? list.map(it => `
            <tr>
              <td>#${it.id}</td>
              <td><code>${it.item_code}</code></td>
              <td><strong>${it.item_name}</strong></td>
              <td><span class="brand-tag">${it.species}</span></td>
              <td style="text-align:center;"><button class="btn btn-outline" style="color:#dc2626;" onclick="deleteItemMaster(${it.id})"><i class="bi bi-trash"></i> 삭제</button></td>
            </tr>
          `).join('') : '<tr><td colspan="5" style="text-align:center; padding:20px; color:#94a3b8;">등록된 품목(대분류)이 없습니다.</td></tr>';
        } else {
          head.innerHTML = '<tr><th>부위 ID</th><th>소속 품목(대분류)</th><th>축종</th><th>부위코드</th><th>세부 부위명</th><th>기본 보관형태</th><th style="text-align:center;">관리</th></tr>';
          const res = await fetch('/api/cuts'), list = await res.json();
          body.innerHTML = list.length ? list.map(c => `
            <tr>
              <td>#${c.id}</td>
              <td><span class="item-tag">${c.parent_item_name}</span></td>
              <td>${c.species}</td>
              <td><code>${c.cut_code}</code></td>
              <td><strong>${c.cut_name}</strong></td>
              <td><span class="badge ${c.default_storage==='냉장'?'badge-chilled':'badge-frozen'}">${c.default_storage}</span></td>
              <td style="text-align:center;"><button class="btn btn-outline" style="color:#dc2626;" onclick="deleteCutMaster(${c.id})"><i class="bi bi-trash"></i> 삭제</button></td>
            </tr>
          `).join('') : '<tr><td colspan="7" style="text-align:center; padding:20px; color:#94a3b8;">등록된 부위(하위분류)가 없습니다.</td></tr>';
        }
      }
    }

    async function advIn(id) { const r = await fetch(`/api/inbounds/${id}/advance`,{method:'POST'}); const d = await r.json(); alert(d.message); loadData(); }
    async function advOut(id) { const r = await fetch(`/api/outbounds/${id}/advance`,{method:'POST'}); const d = await r.json(); alert(d.message); loadData(); }
    async function revertIn(id, msg) { if(!confirm(`정말 [${msg}] 작업을 진행하시겠습니까?`)) return; const r = await fetch(`/api/inbounds/${id}/revert`,{method:'POST'}); const d = await r.json(); alert(d.message); loadData(); }
    async function revertOut(id, msg) { if(!confirm(`정말 [${msg}] 작업을 진행하시겠습니까?`)) return; const r = await fetch(`/api/outbounds/${id}/revert`,{method:'POST'}); const d = await r.json(); alert(d.message); loadData(); }

    // 출하요청 잔여 소비기한 경고 팝업 로직
    async function reqOutWithWarning(lotId, gridNo, maxBox, avgBoxWeight, costPrice, storage, item, expDate) {
      const diffDays = getRemainingDays(expDate);

      // 경고 조건 체크
      if (storage === '냉장') {
        if ((item.includes('돈육') || item.includes('돼지')) && diffDays <= 15) {
          if (!confirm(`⚠️ [냉장 돈육 경고]\n잔여 소비기한이 ${diffDays}일 남았습니다. (15일 이하 임박)\n계속 출하요청을 진행하시겠습니까?`)) return;
        } else if ((item.includes('우육') || item.includes('소')) && diffDays <= 30) {
          if (!confirm(`⚠️ [냉장 우육 경고]\n잔여 소비기한이 ${diffDays}일 남았습니다. (30일 이하 임박)\n계속 출하요청을 진행하시겠습니까?`)) return;
        }
      } else if (storage === '냉동') {
        if (diffDays <= 200) {
          if (!confirm(`🚨 [냉동육 적체 경고]\n잔여 소비기한이 ${diffDays}일 남았습니다. (200일 이하)\n적체해소방안 필요 대상입니다. 출하를 진행하시겠습니까?`)) return;
        } else if (diffDays <= 300) {
          if (!confirm(`⚠️ [냉동육 준적체 경고]\n잔여 소비기한이 ${diffDays}일 남았습니다. (300일 이하 준적체 재고)\n출하를 진행하시겠습니까?`)) return;
        }
      }

      const cust = prompt(`[${gridNo}] 출하할 매출처 상호명을 입력하세요:`, "(주)하남돼지집"); if(!cust) return;
      const box = prompt(`출하 요청 박스 수를 입력하세요 (현재 잔여: ${maxBox} Box, 적용 평중: ${avgBoxWeight} kg/Box):`, "10"); if(!box || isNaN(box) || Number(box)<=0 || Number(box)>maxBox) { alert("박스 수가 올바르지 않습니다."); return; }
      const price = prompt("kg당 판매단가(원)를 입력하세요:", String(costPrice + 2000)); if(!price) return;
      const r = await fetch('/api/outbounds/create-from-stock', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({lot_id:lotId, customer:cust, box_qty:Number(box), unit_price_kg:Number(price)})});
      const d = await r.json(); alert(d.message); setSubTab('OUT_REQUEST');
    }

    // 예약 등록 모달 열기 & 처리
    function openReserveModal(lotId, maxBox, costPrice) {
      document.getElementById('res_lot_id').value = lotId;
      document.getElementById('res_box_qty').value = Math.min(10, maxBox);
      document.getElementById('res_unit_price').value = costPrice + 2000;
      
      const expDate = new Date();
      expDate.setDate(expDate.getDate() + 7); // 기본 7일 후 만료
      document.getElementById('res_expire_date').value = expDate.toISOString().split('T')[0];

      openModal('reserveModal');
    }

    async function submitReservation() {
      const payload = {
        lot_id: Number(document.getElementById('res_lot_id').value),
        sales_rep: document.getElementById('res_sales_rep').value,
        customer: document.getElementById('res_customer').value,
        box_qty: Number(document.getElementById('res_box_qty').value),
        unit_price_kg: Number(document.getElementById('res_unit_price').value),
        expire_date: document.getElementById('res_expire_date').value
      };
      if (!payload.sales_rep || !payload.customer) { alert("영업사원명과 매출처명을 입력하세요."); return; }
      const r = await fetch('/api/reservations', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
      const d = await r.json();
      alert(d.message);
      closeModal('reserveModal');
      loadData();
    }

    async function cancelReservationPrompt(resId) {
      const reason = prompt("예약 취소 사유를 입력하세요:", "거래처 발주 취소");
      if (!reason) return;
      const r = await fetch(`/api/reservations/${resId}/cancel`, {
        method: 'POST',
        headers: {'Content-Type':'application/json'},
        body: JSON.stringify({ cancel_reason: reason, cancel_type: "수동 요청취소" })
      });
      const d = await r.json();
      alert(d.message);
      loadData();
    }

    function openClaimModal(type, id) {
      document.getElementById('claim_type').value = type;
      document.getElementById('claim_id').value = id;
      document.getElementById('claim_date').value = todayStr;
      document.getElementById('claimModalTitle').innerText = `${type==='INBOUND'?'입고':'출고'} 클레임 등록 및 전표 이관 (ID: ${id})`;
      openModal('claimModal');
    }

    async function submitClaim() {
      const type = document.getElementById('claim_type').value;
      const id = document.getElementById('claim_id').value;
      const reason = document.getElementById('claim_reason_text').value;
      const processed_date = document.getElementById('claim_date').value;

      const url = type === 'INBOUND' ? `/api/inbounds/${id}/claim` : `/api/outbounds/${id}/claim`;
      const r = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ reason: reason, processed_date: processed_date })
      });
      const d = await r.json();
      alert(d.message);
      closeModal('claimModal');
      loadData();
    }

    function openPdfPreview(item) {
      document.getElementById('pdf_warehouse').innerText = (item.warehouse || '보관창고') + ' 담당자 귀중';
      document.getElementById('pdf_out_date').innerText = item.outbound_date;
      document.getElementById('pdf_customer').innerText = item.customer;
      document.getElementById('pdf_out_no').innerText = item.outbound_no;
      document.getElementById('pdf_process_from').innerText = item.process_from_date || '-';
      
      document.getElementById('pdf_grid').innerText = item.grid_no || '-';
      document.getElementById('pdf_bl').innerText = item.bl_no || '-';
      document.getElementById('pdf_trace').innerText = item.trace_no;
      document.getElementById('pdf_item_cut').innerText = `[${item.brand||''}] ${item.item_name} ${item.cut_name}`;
      document.getElementById('pdf_storage').innerText = item.storage_type;
      document.getElementById('pdf_avg').innerText = item.avg_box_weight + ' kg/Box';
      document.getElementById('pdf_box').innerText = item.box_qty + ' Box';
      document.getElementById('pdf_weight').innerText = item.weight_kg.toLocaleString() + ' kg';
      document.getElementById('pdf_exp').innerText = item.exp_date;
      document.getElementById('pdf_today').innerText = new Date().toISOString().split('T')[0];

      openModal('pdfModal');
    }

    async function openInboundCreateModal() {
      await loadWarehouseOptions();
      document.getElementById('in_process_from').value = todayStr;
      autoCalculateExpDate();
      openModal('inboundModal');
    }

    async function openEditModal(type, id, dt, bl_no, brand, item_name, cut_name, warehouse, box, weight, price, is_weighed, exp_date, process_from_date) {
      await loadWarehouseOptions();
      document.getElementById('edit_type').value = type;
      document.getElementById('edit_id').value = id;
      document.getElementById('edit_date').value = dt;
      document.getElementById('edit_box').value = box;
      document.getElementById('edit_weight').value = weight;
      document.getElementById('edit_price').value = price;
      document.getElementById('edit_date_label').innerText = type === 'INBOUND' ? '입고 일자' : '출고 일자';
      document.getElementById('edit_price_label').innerText = type === 'INBOUND' ? '매입단가(원/kg)' : '판매단가(원/kg)';

      const inFields = document.getElementById('edit_inbound_fields');
      if (type === 'INBOUND') {
        inFields.style.display = 'block';
        document.getElementById('edit_bl_no').value = bl_no;
        document.getElementById('edit_process_from').value = process_from_date || '';
        document.getElementById('edit_brand').value = brand;
        document.getElementById('edit_item_name').value = item_name;
        document.getElementById('edit_cut_name').value = cut_name;
        document.getElementById('edit_warehouse').value = warehouse || '광주냉장창고';
        document.getElementById('edit_exp_date').value = exp_date || '';
      } else {
        inFields.style.display = 'none';
      }
      openModal('editModal');
    }

    async function submitEdit() {
      const type = document.getElementById('edit_type').value;
      const id = document.getElementById('edit_id').value;
      let payload = {};

      if (type === 'INBOUND') {
        payload = {
          inbound_date: document.getElementById('edit_date').value,
          bl_no: document.getElementById('edit_bl_no').value,
          process_from_date: document.getElementById('edit_process_from').value,
          brand: document.getElementById('edit_brand').value,
          item_name: document.getElementById('edit_item_name').value,
          cut_name: document.getElementById('edit_cut_name').value,
          warehouse: document.getElementById('edit_warehouse').value,
          box_qty: Number(document.getElementById('edit_box').value),
          weight_kg: Number(document.getElementById('edit_weight').value),
          cost_per_kg: Number(document.getElementById('edit_price').value),
          is_weighed: "N",
          exp_date: document.getElementById('edit_exp_date').value
        };
      } else {
        payload = {
          outbound_date: document.getElementById('edit_date').value,
          box_qty: Number(document.getElementById('edit_box').value),
          weight_kg: Number(document.getElementById('edit_weight').value),
          unit_price_kg: Number(document.getElementById('edit_price').value)
        };
      }

      const url = type === 'INBOUND' ? `/api/inbounds/${id}` : `/api/outbounds/${id}`;
      const r = await fetch(url, { method: 'PUT', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
      const d = await r.json();
      if(r.ok) { alert(d.message); closeModal('editModal'); loadData(); } else { alert(d.detail); }
    }

    async function submitPartner() {
      const p = {
        name: document.getElementById('p_name').value,
        type: document.getElementById('p_type').value,
        biz_no: document.getElementById('p_biz_no').value,
        contact_person: document.getElementById('p_contact').value,
        phone: document.getElementById('p_phone').value,
        address: document.getElementById('p_address').value
      };
      if(!p.name) { alert("상호명을 입력하세요."); return; }
      const r = await fetch('/api/partners', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p)});
      const d = await r.json(); alert(d.message); closeModal('partnerModal'); loadData();
    }
    async function deletePartner(id) {
      if(!confirm("정말 이 거래처를 삭제하시겠습니까?")) return;
      const r = await fetch(`/api/partners/${id}`, {method:'DELETE'});
      const d = await r.json(); alert(d.message); loadData();
    }

    async function submitItemMaster() {
      const p = {
        item_code: document.getElementById('m_item_code').value,
        item_name: document.getElementById('m_item_name').value,
        species: document.getElementById('m_species').value
      };
      if(!p.item_code || !p.item_name) { alert("품목코드와 품목명은 필수입니다."); return; }
      const r = await fetch('/api/items', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p)});
      const d = await r.json();
      if(r.ok) { alert(d.message); closeModal('itemMasterModal'); loadData(); } else { alert(d.detail); }
    }
    async function deleteItemMaster(id) {
      if(!confirm("품목을 삭제하면 하위 소속 부위도 함께 삭제됩니다. 계속하시겠습니까?")) return;
      const r = await fetch(`/api/items/${id}`, {method:'DELETE'});
      const d = await r.json(); alert(d.message); loadData();
    }

    async function openCutModal() {
      const res = await fetch('/api/items');
      const items = await res.json();
      const select = document.getElementById('m_cut_parent_item');
      select.innerHTML = items.map(it => `<option value="${it.id}">[${it.species}] ${it.item_name} (${it.item_code})</option>`).join('');
      openModal('cutMasterModal');
    }
    async function submitCutMaster() {
      const p = {
        item_id: Number(document.getElementById('m_cut_parent_item').value),
        cut_code: document.getElementById('m_cut_code').value,
        cut_name: document.getElementById('m_cut_name').value,
        default_storage: document.getElementById('m_cut_storage').value
      };
      if(!p.cut_code || !p.cut_name) { alert("부위코드와 부위명은 필수입니다."); return; }
      const r = await fetch('/api/cuts', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p)});
      const d = await r.json();
      if(r.ok) { alert(d.message); closeModal('cutMasterModal'); loadData(); } else { alert(d.detail); }
    }
    async function deleteCutMaster(id) {
      if(!confirm("정말 이 세부 부위를 삭제하시겠습니까?")) return;
      const r = await fetch(`/api/cuts/${id}`, {method:'DELETE'});
      const d = await r.json(); alert(d.message); loadData();
    }

    function openModal(id) { document.getElementById(id).style.display = 'flex'; }
    function closeModal(id) { document.getElementById(id).style.display = 'none'; }

    async function submitInbound() {
      const p = {
        inbound_date: document.getElementById('in_date').value,
        vendor: document.getElementById('in_vendor').value,
        bl_no: document.getElementById('in_bl_no').value,
        trace_no: document.getElementById('in_trace').value,
        process_from_date: document.getElementById('in_process_from').value,
        brand: document.getElementById('in_brand').value,
        item_name: document.getElementById('in_item').value,
        cut_name: document.getElementById('in_cut').value,
        storage_type: document.getElementById('in_storage').value,
        box_qty: Number(document.getElementById('in_box').value),
        weight_kg: Number(document.getElementById('in_weight').value),
        cost_per_kg: Number(document.getElementById('in_cost').value),
        warehouse: document.getElementById('in_warehouse').value,
        exp_date: document.getElementById('in_exp').value,
        is_weighed: "N"
      };
      const r = await fetch('/api/inbounds', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p)});
      const d = await r.json(); alert(d.message); closeModal('inboundModal'); setSubTab('IN_REQUEST');
    }

    function openAdjustModal(lotId) { document.getElementById('adj_lot_id').value = lotId; openModal('adjustModal'); }
    async function submitAdjust() {
      const p = { lot_id: Number(document.getElementById('adj_lot_id').value), adj_type: document.getElementById('adj_type').value, adj_box: Number(document.getElementById('adj_box').value), adj_weight: Number(document.getElementById('adj_weight').value), reason: document.getElementById('adj_reason').value };
      const r = await fetch('/api/inventory/adjust', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p)});
      const d = await r.json(); alert(d.message); closeModal('adjustModal'); loadData();
    }

    async function uploadExcelFile(e) {
      const file = e.target.files[0]; if (!file) return;
      const formData = new FormData(); formData.append('file', file);
      const r = await fetch('/api/inbounds/upload-excel', { method: 'POST', body: formData });
      const d = await r.json(); alert(d.message || d.detail); e.target.value = ''; setSubTab('IN_REQUEST');
    }

    window.addEventListener('DOMContentLoaded', () => { 
      loadWarehouseOptions();
      renderSubTabs(); 
      loadData(); 
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
