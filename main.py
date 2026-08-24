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

class ReservationRecord(Base):
    __tablename__ = "reservations"
    id = Column(Integer, primary_key=True, index=True)
    res_no = Column(String(50), unique=True, index=True)
    lot_id = Column(Integer, nullable=False)
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
    expire_date = Column(String(20), nullable=False)
    cancel_date = Column(String(20), nullable=True)
    cancel_type = Column(String(50), nullable=True)
    cancel_reason = Column(String(255), nullable=True)
    status = Column(String(20), default="HOLD")
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

# 초기 시드 데이터
def init_sample_data():
    db = SessionLocal()
    try:
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
    finally:
        db.close()

init_sample_data()

# -----------------------------------------------------------------------------
# 3. FastAPI 라우터 및 요청 스키마
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

# --- [거래처 API] ---
@app.get("/api/partners")
def get_partners(type: Optional[str] = None, db: Session = Depends(get_db)):
    q = db.query(Partner)
    if type and type != "ALL":
        q = q.filter(Partner.type == type)
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
    if not p:
        raise HTTPException(status_code=404, detail="거래처를 찾을 수 없습니다.")
    db.delete(p)
    db.commit()
    return {"message": "거래처가 삭제되었습니다."}

# --- [품목/부위 마스터 API] ---
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
    if not item:
        raise HTTPException(status_code=404, detail="품목을 찾을 수 없습니다.")
    db.query(CutMaster).filter(CutMaster.item_id == item_id).delete()
    db.delete(item)
    db.commit()
    return {"message": "품목 및 소속 부위가 삭제되었습니다."}

@app.get("/api/cuts")
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
    if not cut:
        raise HTTPException(status_code=404, detail="부위를 찾을 수 없습니다.")
    db.delete(cut)
    db.commit()
    return {"message": "부위 마스터가 삭제되었습니다."}

# --- [입고 관리 API] ---
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
        if start_date:
            q = q.filter(InboundRecord.inbound_date >= start_date)
        if end_date:
            q = q.filter(InboundRecord.inbound_date <= end_date)
    return q.order_by(desc(InboundRecord.inbound_date), desc(InboundRecord.id)).all()

@app.post("/api/inbounds")
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
    return {"message": f"입고등록 완료 ({grid_no})", "inbound_no": inbound_no}

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

    for r_idx, row in enumerate(rows[1:], start=2):
        if not row or all(v is None for v in row):
            continue
        try:
            # 엑셀 헤더 규격 매핑: 입고일자(0), 매입처(1), BL(2), 이력번호(3), 가공일(4), 브랜드(5), 품목(6), 부위(7), 보관(8), 박스(9), 실중량(10), 단가(11), 창고(12)
            in_date = str(row[0])[:10] if row[0] else now_str
            vendor = str(row[1]).strip() if len(row) > 1 and row[1] else "일괄매입처"
            bl_no = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            trace_no = str(row[3]).strip() if len(row) > 3 and row[3] else f"EXCEL-{random.randint(100000,999999)}"
            process_from = str(row[4])[:10] if len(row) > 4 and row[4] else in_date
            brand = str(row[5]).strip() if len(row) > 5 and row[5] else ""
            item_name = str(row[6]).strip() if len(row) > 6 and row[6] else "돈육(돼지)"
            cut_name = str(row[7]).strip() if len(row) > 7 and row[7] else "삼겹살"
            storage_type = str(row[8]).strip() if len(row) > 8 and row[8] else "냉장"
            box_qty = int(row[9]) if len(row) > 9 and row[9] else 1
            weight_kg = float(row[10]) if len(row) > 10 and row[10] else 20.0
            cost_per_kg = float(row[11]) if len(row) > 11 and row[11] else 8500.0
            warehouse = str(row[12]).strip() if len(row) > 12 and row[12] else "광주냉장창고"
            
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
            continue

    db.commit()
    return {"message": f"총 {created_count}건의 데이터가 일괄 입고 등록되었습니다."}

@app.post("/api/inbounds/{inbound_id}/claim")
def register_inbound_claim(inbound_id: int, req: ClaimRegister, db: Session = Depends(get_db)):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).first()
    if not inbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    
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

@app.post("/api/inbounds/{inbound_id}/advance")
def advance_inbound(inbound_id: int, db: Session = Depends(get_db)):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).first()
    if not inbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
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

@app.post("/api/inbounds/{inbound_id}/revert")
def revert_inbound(inbound_id: int, db: Session = Depends(get_db)):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).first()
    if not inbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
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

# --- [현 재고장 및 출고 관리 API] ---
@app.get("/api/inventory")
def get_inventory(db: Session = Depends(get_db)):
    today_check = datetime.now().strftime("%Y-%m-%d")
    
    # 1. 예약 만료일 경과 건 자동 취소 체크
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
        if start_date:
            q = q.filter(OutboundRecord.outbound_date >= start_date)
        if end_date:
            q = q.filter(OutboundRecord.outbound_date <= end_date)
    return q.order_by(desc(OutboundRecord.outbound_date), desc(OutboundRecord.id)).all()

@app.post("/api/outbounds/create-from-stock")
def create_outbound_from_stock(req: OutboundCreate, db: Session = Depends(get_db)):
    lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).first()
    if not lot:
        raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")
    
    # 예약 홀딩 수량 계산 및 가용 재고 검증
    reserved_boxes = sum([r.box_qty for r in db.query(ReservationRecord).filter(ReservationRecord.lot_id == lot.id, ReservationRecord.status == "HOLD").all()])
    avail_box = lot.current_box_qty - reserved_boxes

    if req.box_qty > avail_box or req.box_qty <= 0:
        raise HTTPException(status_code=400, detail=f"출하 가능 수량({avail_box} Box)을 초과했습니다. (예약 홀딩: {reserved_boxes} Box)")

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

@app.post("/api/outbounds/{outbound_id}/advance")
def advance_outbound(outbound_id: int, db: Session = Depends(get_db)):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
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

@app.post("/api/outbounds/{outbound_id}/claim")
def register_outbound_claim(outbound_id: int, req: ClaimRegister, db: Session = Depends(get_db)):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    outbound.status = "OUT_CLAIM"
    outbound.claim_reason = req.reason
    outbound.processed_date = req.processed_date if req.processed_date else datetime.now().strftime("%Y-%m-%d")
    db.commit()
    return {"message": f"출고 클레임 등록이 완료되었습니다. (처리일자: {outbound.processed_date})"}

@app.put("/api/outbounds/{outbound_id}")
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

@app.post("/api/outbounds/{outbound_id}/revert")
def revert_outbound(outbound_id: int, db: Session = Depends(get_db)):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

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

# --- [예약 관리 API] ---
@app.get("/api/reservations")
def get_reservations(status: str, db: Session = Depends(get_db)):
    return db.query(ReservationRecord).filter(ReservationRecord.status == status).order_by(desc(ReservationRecord.id)).all()

@app.post("/api/reservations")
def create_reservation(req: ReservationCreate, db: Session = Depends(get_db)):
    lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).first()
    if not lot:
        raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")
    
    # 활성 예약 수량 합산 및 예약 가능 한도 점검
    already_reserved = sum([r.box_qty for r in db.query(ReservationRecord).filter(ReservationRecord.lot_id == lot.id, ReservationRecord.status == "HOLD").all()])
    avail_box = lot.current_box_qty - already_reserved

    if req.box_qty > avail_box or req.box_qty <= 0:
        raise HTTPException(status_code=400, detail=f"예약 가능 잔여 박스({avail_box} Box)를 초과하여 예약할 수 없습니다.")

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
    if not res:
        raise HTTPException(status_code=404, detail="예약 내역을 찾을 수 없습니다.")
    
    res.status = "CANCELLED"
    res.cancel_type = req.cancel_type or "수동 요청취소"
    res.cancel_reason = req.cancel_reason
    res.cancel_date = datetime.now().strftime("%Y-%m-%d")
    db.commit()
    return {"message": "예약이 취소되어 예약요청취소리스트에 누적되었습니다."}

# --- [통합 클레임 관리 API] ---
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
    if not lot:
        raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")
    lot.current_box_qty = max(0, lot.current_box_qty - req.adj_box)
    lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - req.adj_weight, 2))
    log = StockAdjustment(lot_id=lot.id, adj_type=req.adj_type, adj_box=req.adj_box, adj_weight=req.adj_weight, reason=req.reason)
    db.add(log)
    db.commit()
    return {"message": f"재고 조정({req.adj_type})이 완료되었습니다."}

# -----------------------------------------------------------------------------
# 4. 프론트엔드 HTML
# -----------------------------------------------------------------------------
# (동일한 UI 코드베이스 유지)
