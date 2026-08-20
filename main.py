import io
import os
from datetime import datetime, date
from typing import List, Optional

from fastapi import FastAPI, HTTPException, Depends, UploadFile, File, Query
from fastapi.responses import HTMLResponse
import openpyxl
from pydantic import BaseModel
from sqlalchemy import (
    Column, Integer, Float, String, DateTime, create_engine, desc
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
# 2. ORM 테이블 정의 (브랜드 컬럼 분리 적용)
# -----------------------------------------------------------------------------
class Partner(Base):
    __tablename__ = "partners"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), nullable=False)
    type = Column(String(20), nullable=False)  # 'VENDOR', 'CUSTOMER', 'WAREHOUSE'
    biz_no = Column(String(30), nullable=True)
    contact_person = Column(String(50), nullable=True)
    phone = Column(String(30), nullable=True)
    address = Column(String(200), nullable=True)
    created_at = Column(DateTime, default=datetime.now)

class Product(Base):
    __tablename__ = "products"
    id = Column(Integer, primary_key=True, index=True)
    sku_code = Column(String(50), unique=True, nullable=False)
    species = Column(String(30), nullable=False)  # 소, 돼지 등
    cut_name = Column(String(50), nullable=False)  # 삼겹살, 목심 등
    brand = Column(String(50), nullable=True)      # 올리멜, 엑셀 등 (분리)
    origin = Column(String(50), nullable=False)    # 캐나다, 미국, 국내산 등
    storage_type = Column(String(20), nullable=False) # 냉장, 냉동
    created_at = Column(DateTime, default=datetime.now)

class InboundRecord(Base):
    __tablename__ = "inbounds"
    id = Column(Integer, primary_key=True, index=True)
    inbound_no = Column(String(50), unique=True, index=True)
    inbound_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d")) # 입고일자
    vendor = Column(String(100), nullable=False)
    brand = Column(String(50), nullable=True)       # 브랜드/가공장 (분리)
    item_name = Column(String(100), nullable=False) # 품목/부위명
    storage_type = Column(String(20), nullable=False)
    trace_no = Column(String(50), nullable=False, index=True)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    cost_per_kg = Column(Float, nullable=False)
    total_amount = Column(Float, nullable=False)
    warehouse = Column(String(50), default="광주냉장창고")
    exp_date = Column(String(20), nullable=False)
    status = Column(String(20), default="IN_REQUEST")
    created_at = Column(DateTime, default=datetime.now)

class InventoryLot(Base):
    __tablename__ = "inventory_lots"
    id = Column(Integer, primary_key=True, index=True)
    sku_code = Column(String(50), nullable=False)
    brand = Column(String(50), nullable=True)       # 브랜드 (분리)
    item_name = Column(String(100), nullable=False) # 품목명
    storage_type = Column(String(20), nullable=False)
    trace_no = Column(String(50), nullable=False, index=True)
    current_box_qty = Column(Integer, nullable=False)
    current_weight_kg = Column(Float, nullable=False)
    cost_per_kg = Column(Float, nullable=False)
    warehouse = Column(String(50), nullable=False)
    exp_date = Column(String(20), nullable=False)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

class OutboundRecord(Base):
    __tablename__ = "outbounds"
    id = Column(Integer, primary_key=True, index=True)
    outbound_no = Column(String(50), unique=True, index=True)
    outbound_date = Column(String(20), default=lambda: datetime.now().strftime("%Y-%m-%d")) # 출고일자
    lot_id = Column(Integer, nullable=False)
    customer = Column(String(100), nullable=False)
    brand = Column(String(50), nullable=True)       # 브랜드 (분리)
    item_name = Column(String(100), nullable=False) # 품목명
    storage_type = Column(String(20), nullable=False)
    trace_no = Column(String(50), nullable=False)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    unit_price_kg = Column(Float, nullable=False)
    total_amount = Column(Float, nullable=False)
    exp_date = Column(String(20), nullable=False)
    warehouse = Column(String(50), nullable=False)
    status = Column(String(20), default="OUT_REQUEST")
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

# 초기 시드 데이터
def init_sample_data():
    db = SessionLocal()
    if db.query(Partner).count() == 0:
        partners = [
            Partner(name="(주)글로벌미트", type="VENDOR", biz_no="105-86-12345", contact_person="김수입", phone="010-1111-2222", address="서울시 송파구"),
            Partner(name="도드람양돈농협", type="VENDOR", biz_no="214-82-67890", contact_person="이한돈", phone="010-3333-4444", address="경기도 이천시"),
            Partner(name="(주)하남돼지집", type="CUSTOMER", biz_no="120-81-99887", contact_person="박대표", phone="010-5555-6666", address="경기도 하남시"),
            Partner(name="광주냉장창고", type="WAREHOUSE", biz_no="110-85-44332", contact_person="최창고", phone="031-760-1234", address="경기도 광주시"),
            Partner(name="용인냉동센터", type="WAREHOUSE", biz_no="142-88-55667", contact_person="정소장", phone="031-330-5678", address="경기도 용인시")
        ]
        db.add_all(partners)
    
    if db.query(Product).count() == 0:
        products = [
            Product(sku_code="PK-CAN-01", species="돼지", cut_name="삼겹살", brand="올리멜", origin="캐나다", storage_type="냉장"),
            Product(sku_code="BF-USA-05", species="소", cut_name="척아이롤", brand="엑셀", origin="미국", storage_type="냉동"),
            Product(sku_code="PK-KOR-01", species="돼지", cut_name="생삼겹", brand="도드람", origin="국내산", storage_type="냉장"),
            Product(sku_code="PK-USA-02", species="돼지", cut_name="목심", brand="스위프트", origin="미국", storage_type="냉동")
        ]
        db.add_all(products)

    if db.query(InventoryLot).count() == 0:
        lots = [
            InventoryLot(sku_code="PK-CAN-01", brand="올리멜", item_name="돼지 삼겹살", storage_type="냉장", trace_no="802410290114", current_box_qty=45, current_weight_kg=912.4, cost_per_kg=8400, warehouse="광주냉장창고", exp_date="2026-09-02"),
            InventoryLot(sku_code="BF-USA-05", brand="엑셀", item_name="소 척아이롤", storage_type="냉동", trace_no="100392019482", current_box_qty=120, current_weight_kg=2450.0, cost_per_kg=14500, warehouse="용인냉동센터", exp_date="2027-04-15"),
            InventoryLot(sku_code="PK-KOR-01", brand="도드람", item_name="국내산 생삼겹", storage_type="냉장", trace_no="002194820192", current_box_qty=15, current_weight_kg=305.5, cost_per_kg=16800, warehouse="광주냉장창고", exp_date="2026-08-25")
        ]
        db.add_all(lots)
    db.commit()
    db.close()

init_sample_data()

# -----------------------------------------------------------------------------
# 3. FastAPI API 엔드포인트
# -----------------------------------------------------------------------------
app = FastAPI(title="MeatFlow Enterprise ERP")

# Pydantic 모델
class PartnerCreate(BaseModel):
    name: str
    type: str
    biz_no: Optional[str] = ""
    contact_person: Optional[str] = ""
    phone: Optional[str] = ""
    address: Optional[str] = ""

class ProductCreate(BaseModel):
    sku_code: str
    species: str
    cut_name: str
    brand: Optional[str] = ""
    origin: str
    storage_type: str

class InboundCreate(BaseModel):
    inbound_date: Optional[str] = None
    vendor: str
    brand: Optional[str] = ""
    item_name: str
    storage_type: str
    trace_no: str
    box_qty: int
    weight_kg: float
    cost_per_kg: float
    warehouse: str
    exp_date: str

class UpdateInbound(BaseModel):
    inbound_date: str
    brand: Optional[str] = ""
    item_name: str
    box_qty: int
    weight_kg: float
    cost_per_kg: float

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

class AdjustCreate(BaseModel):
    lot_id: int
    adj_type: str
    adj_box: int
    adj_weight: float
    reason: Optional[str] = ""

# --- [1. 거래처 API] ---
@app.get("/api/partners")
def get_partners(type: Optional[str] = None, db: Session = Depends(get_db)):
    query = db.query(Partner)
    if type and type != "ALL": query = query.filter(Partner.type == type)
    return query.order_by(desc(Partner.id)).all()

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

# --- [2. 상품 마스터 API] ---
@app.get("/api/products")
def get_products(db: Session = Depends(get_db)):
    return db.query(Product).order_by(desc(Product.id)).all()

@app.post("/api/products")
def create_product(req: ProductCreate, db: Session = Depends(get_db)):
    if db.query(Product).filter(Product.sku_code == req.sku_code).first():
        raise HTTPException(status_code=400, detail="이미 등록된 품목코드(SKU)입니다.")
    p = Product(sku_code=req.sku_code, species=req.species, cut_name=req.cut_name, brand=req.brand, origin=req.origin, storage_type=req.storage_type)
    db.add(p)
    db.commit()
    return {"message": "상품 마스터가 등록되었습니다."}

@app.delete("/api/products/{product_id}")
def delete_product(product_id: int, db: Session = Depends(get_db)):
    p = db.query(Product).filter(Product.id == product_id).first()
    if not p: raise HTTPException(status_code=404, detail="상품을 찾을 수 없습니다.")
    db.delete(p)
    db.commit()
    return {"message": "상품 마스터가 삭제되었습니다."}

# --- [3. 입고 API (기간 필터 포함)] ---
@app.get("/api/inbounds")
def get_inbounds(
    status: str,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    month: Optional[str] = None,
    db: Session = Depends(get_db)
):
    query = db.query(InboundRecord).filter(InboundRecord.status == status)
    if month:
        query = query.filter(InboundRecord.inbound_date.like(f"{month}%"))
    else:
        if start_date: query = query.filter(InboundRecord.inbound_date >= start_date)
        if end_date: query = query.filter(InboundRecord.inbound_date <= end_date)
    return query.order_by(desc(InboundRecord.inbound_date), desc(InboundRecord.id)).all()

@app.post("/api/inbounds")
def create_inbound(req: InboundCreate, db: Session = Depends(get_db)):
    inbound_no = f"IN-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    in_date = req.inbound_date if req.inbound_date else datetime.now().strftime("%Y-%m-%d")
    total = round(req.weight_kg * req.cost_per_kg)
    item = InboundRecord(
        inbound_no=inbound_no, inbound_date=in_date, vendor=req.vendor, brand=req.brand,
        item_name=req.item_name, storage_type=req.storage_type, trace_no=req.trace_no,
        box_qty=req.box_qty, weight_kg=req.weight_kg, cost_per_kg=req.cost_per_kg,
        total_amount=total, warehouse=req.warehouse, exp_date=req.exp_date, status="IN_REQUEST"
    )
    db.add(item)
    db.commit()
    return {"message": "입고요청 등록 완료", "inbound_no": inbound_no}

@app.put("/api/inbounds/{inbound_id}")
def update_inbound(inbound_id: int, req: UpdateInbound, db: Session = Depends(get_db)):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).first()
    if not inbound: raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    if inbound.status == "IN_DONE": raise HTTPException(status_code=400, detail="입고완료 전표는 [작업취소] 후 수정해야 합니다.")
    inbound.inbound_date = req.inbound_date
    inbound.brand = req.brand
    inbound.item_name = req.item_name
    inbound.box_qty = req.box_qty
    inbound.weight_kg = req.weight_kg
    inbound.cost_per_kg = req.cost_per_kg
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
        lot = db.query(InventoryLot).filter(
            InventoryLot.trace_no == inbound.trace_no,
            InventoryLot.warehouse == inbound.warehouse,
            InventoryLot.exp_date == inbound.exp_date
        ).first()
        if lot:
            lot.current_box_qty += inbound.box_qty
            lot.current_weight_kg = round(lot.current_weight_kg + inbound.weight_kg, 2)
            lot.cost_per_kg = inbound.cost_per_kg
            lot.brand = inbound.brand
        else:
            sku = "SKU-" + inbound.trace_no[:6]
            new_lot = InventoryLot(
                sku_code=sku, brand=inbound.brand, item_name=inbound.item_name,
                storage_type=inbound.storage_type, trace_no=inbound.trace_no,
                current_box_qty=inbound.box_qty, current_weight_kg=inbound.weight_kg,
                cost_per_kg=inbound.cost_per_kg, warehouse=inbound.warehouse, exp_date=inbound.exp_date
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
            lot.current_box_qty = max(0, lot.current_box_qty - inbound.box_qty)
            lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - inbound.weight_kg, 2))
        inbound.status = "IN_CONFIRM"
        msg = "입고완료가 취소되어 실재고가 원복(차감)되었습니다."
    else:
        msg = "처리할 수 없는 상태입니다."
    db.commit()
    return {"message": msg}

@app.post("/api/inbounds/upload-excel")
async def upload_inbound_excel(file: UploadFile = File(...), db: Session = Depends(get_db)):
    if not file.filename.endswith(('.xlsx', '.xls')): raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")
    try:
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=True)
        sheet = wb.active
        count = 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None: continue
            # [0:입고일자, 1:매입처, 2:브랜드, 3:품목명, 4:보관(냉장/냉동), 5:이력번호, 6:박스, 7:중량kg, 8:단가, 9:창고, 10:유통기한]
            in_date_raw = row[0] if row[0] else datetime.now().strftime("%Y-%m-%d")
            in_date = in_date_raw.strftime('%Y-%m-%d') if isinstance(in_date_raw, (datetime, date)) else str(in_date_raw).strip()
            vendor = str(row[1]).strip()
            brand = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            item_name = str(row[3]).strip() if len(row) > 3 else ""
            storage_type = str(row[4]).strip() if len(row) > 4 else "냉동"
            trace_no = str(row[5]).strip() if len(row) > 5 else ""
            box_qty = int(row[6]) if len(row) > 6 else 0
            weight_kg = float(row[7]) if len(row) > 7 else 0.0
            cost_per_kg = float(row[8]) if len(row) > 8 else 0.0
            warehouse = str(row[9]).strip() if len(row) > 9 and row[9] else "광주냉장창고"
            exp_date_raw = row[10] if len(row) > 10 else "2026-12-31"
            exp_date_str = exp_date_raw.strftime('%Y-%m-%d') if isinstance(exp_date_raw, (datetime, date)) else str(exp_date_raw).strip()

            inbound_no = f"IN-{datetime.now().strftime('%Y%m%d%H%M%S')}-{count+1}"
            inbound = InboundRecord(
                inbound_no=inbound_no, inbound_date=in_date, vendor=vendor, brand=brand,
                item_name=item_name, storage_type=storage_type, trace_no=trace_no,
                box_qty=box_qty, weight_kg=weight_kg, cost_per_kg=cost_per_kg,
                total_amount=round(weight_kg * cost_per_kg), warehouse=warehouse, exp_date=exp_date_str, status="IN_REQUEST"
            )
            db.add(inbound)
            count += 1
        db.commit()
        return {"message": f"총 {count}건의 엑셀 입고 요청이 등록되었습니다."}
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"엑셀 처리 오류: {str(e)}")

# --- [4. 출고 API (일자별/월별 기간 필터 포함)] ---
@app.get("/api/inventory")
def get_inventory(db: Session = Depends(get_db)):
    return db.query(InventoryLot).filter(InventoryLot.current_box_qty > 0).order_by(InventoryLot.exp_date).all()

@app.get("/api/outbounds")
def get_outbounds(
    status: str,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    month: Optional[str] = None,
    db: Session = Depends(get_db)
):
    query = db.query(OutboundRecord).filter(OutboundRecord.status == status)
    if month:
        query = query.filter(OutboundRecord.outbound_date.like(f"{month}%"))
    else:
        if start_date: query = query.filter(OutboundRecord.outbound_date >= start_date)
        if end_date: query = query.filter(OutboundRecord.outbound_date <= end_date)
    return query.order_by(desc(OutboundRecord.outbound_date), desc(OutboundRecord.id)).all()

@app.post("/api/outbounds/create-from-stock")
def create_outbound_from_stock(req: OutboundCreate, db: Session = Depends(get_db)):
    lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).first()
    if not lot: raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")
    if req.box_qty > lot.current_box_qty or req.box_qty <= 0: raise HTTPException(status_code=400, detail="보유 박스 수량보다 많이 요청할 수 없습니다.")
    unit_weight = lot.current_weight_kg / lot.current_box_qty
    calc_weight = round(unit_weight * req.box_qty, 2)
    outbound_no = f"OUT-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    out_date = req.outbound_date if req.outbound_date else datetime.now().strftime("%Y-%m-%d")
    total = round(calc_weight * req.unit_price_kg)
    outbound = OutboundRecord(
        outbound_no=outbound_no, outbound_date=out_date, lot_id=lot.id, customer=req.customer,
        brand=lot.brand, item_name=lot.item_name, storage_type=lot.storage_type, trace_no=lot.trace_no,
        box_qty=req.box_qty, weight_kg=calc_weight, unit_price_kg=req.unit_price_kg, total_amount=total,
        exp_date=lot.exp_date, warehouse=lot.warehouse, status="OUT_REQUEST"
    )
    db.add(outbound)
    db.commit()
    return {"message": "출고요청 등록 완료", "outbound_no": outbound_no}

@app.put("/api/outbounds/{outbound_id}")
def update_outbound(outbound_id: int, req: UpdateOutbound, db: Session = Depends(get_db)):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound: raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    if outbound.status == "OUT_DONE": raise HTTPException(status_code=400, detail="출고완료 전표는 [작업취소] 후 수정해야 합니다.")
    outbound.outbound_date = req.outbound_date
    outbound.box_qty = req.box_qty
    outbound.weight_kg = req.weight_kg
    outbound.unit_price_kg = req.unit_price_kg
    outbound.total_amount = round(req.weight_kg * req.unit_price_kg)
    db.commit()
    return {"message": "출고 정보가 수정되었습니다."}

@app.post("/api/outbounds/{outbound_id}/advance")
def advance_outbound(outbound_id: int, db: Session = Depends(get_db)):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound: raise HTTPException(status_code=404, detail="출고 전표를 찾을 수 없습니다.")
    if outbound.status == "OUT_REQUEST":
        outbound.status = "OUT_CONFIRM"
        msg = "출고확정 단계로 이동되었습니다."
    elif outbound.status == "OUT_CONFIRM":
        outbound.status = "OUT_DONE"
        lot = db.query(InventoryLot).filter(InventoryLot.id == outbound.lot_id).first()
        if lot:
            lot.current_box_qty = max(0, lot.current_box_qty - outbound.box_qty)
            lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - outbound.weight_kg, 2))
        msg = "출고완료 처리 및 실재고 차감이 완료되었습니다."
    else:
        msg = "이미 완료된 전표입니다."
    db.commit()
    return {"message": msg}

@app.post("/api/outbounds/{outbound_id}/revert")
def revert_outbound(outbound_id: int, db: Session = Depends(get_db)):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound: raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    if outbound.status == "OUT_REQUEST":
        db.delete(outbound)
        msg = "출고요청 전표가 삭제되었습니다."
    elif outbound.status == "OUT_CONFIRM":
        outbound.status = "OUT_REQUEST"
        msg = "출고확정이 취소되어 [출고요청리스트]로 반려되었습니다."
    elif outbound.status == "OUT_DONE":
        lot = db.query(InventoryLot).filter(InventoryLot.id == outbound.lot_id).first()
        if lot:
            lot.current_box_qty += outbound.box_qty
            lot.current_weight_kg = round(lot.current_weight_kg + outbound.weight_kg, 2)
        outbound.status = "OUT_CONFIRM"
        msg = "출고완료가 취소되어 실재고가 복구(가산)되었습니다."
    else:
        msg = "처리할 수 없는 상태입니다."
    db.commit()
    return {"message": msg}

# --- [5. 재고 감량/클레임] ---
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
# 4. 프론트엔드 UI (HTML/CSS/JS)
# -----------------------------------------------------------------------------
HTML_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>MeatFlow Enterprise ERP</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
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
    /* 기간 필터 바 */
    .filter-card { background: #fff; border: 1px solid var(--border); border-radius: 8px; padding: 12px 18px; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px; }
    .filter-group { display: flex; align-items: center; gap: 8px; font-size: 0.85rem; font-weight: 600; }
    .filter-input { padding: 6px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.85rem; outline: none; background: #fff; }
    .filter-btn-group { display: flex; gap: 4px; }
    .filter-btn { padding: 6px 10px; border: 1px solid var(--border); background: #f8fafc; border-radius: 6px; font-size: 0.8rem; font-weight: 600; cursor: pointer; }
    .filter-btn.active { background: var(--primary); color: #fff; border-color: var(--primary); }
    /* 테이블 */
    .table-card { background: #fff; border: 1px solid var(--border); border-radius: 8px; overflow: hidden; box-shadow: 0 1px 2px rgba(0,0,0,0.03); }
    table { width: 100%; border-collapse: collapse; font-size: 0.86rem; }
    th { background: #f8fafc; color: var(--muted); padding: 12px 14px; border-bottom: 1px solid var(--border); text-align: left; font-weight: 600; white-space: nowrap; }
    td { padding: 12px 14px; border-bottom: 1px solid var(--border); vertical-align: middle; }
    tr:hover td { background: #f8fafc; }
    .btn { padding: 5px 10px; border-radius: 6px; font-size: 0.8rem; font-weight: 600; border: 1px solid transparent; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-success { background: #16a34a; color: #fff; }
    .btn-warning { background: #d97706; color: #fff; }
    .btn-danger { background: #dc2626; color: #fff; }
    .btn-outline { background: #fff; border-color: var(--border); color: #334155; }
    .btn-outline:hover { background: #f1f5f9; }
    .badge { padding: 3px 8px; border-radius: 4px; font-size: 0.75rem; font-weight: 600; }
    .badge-chilled { background: #e0f2fe; color: #0284c7; }
    .badge-frozen { background: #e0e7ff; color: #4f46e5; }
    .badge-vendor { background: #fef3c7; color: #b45309; }
    .badge-customer { background: #dcfce7; color: #15803d; }
    .badge-warehouse { background: #f1f5f9; color: #475569; }
    .amount-highlight { font-weight: 700; color: #0f172a; }
    .brand-tag { background: #f1f5f9; color: #334155; padding: 2px 6px; border-radius: 4px; font-weight: 600; font-size: 0.78rem; margin-right: 4px; }
    /* 모달 */
    .modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 100; align-items: center; justify-content: center; }
    .modal-box { background: #fff; border-radius: 10px; width: 520px; padding: 24px; max-height: 90vh; overflow-y: auto; }
    .form-group { margin-bottom: 12px; display: flex; flex-direction: column; gap: 4px; font-size: 0.82rem; font-weight: 600; }
    .form-control { padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.88rem; }
  </style>
</head>
<body>
  <aside>
    <div class="brand"><i class="bi bi-box-seam-fill"></i> MeatFlow ERP</div>
    <div class="nav-category">프로세스 관리</div>
    <div class="nav-item active" onclick="switchMainTab('INBOUND', this)"><i class="bi bi-box-arrow-in-down"></i> 입고 관리</div>
    <div class="nav-item" onclick="switchMainTab('OUTBOUND', this)"><i class="bi bi-box-arrow-up"></i> 출고 관리 (재고장)</div>

    <div class="nav-category">기본정보 마스터</div>
    <div class="nav-item" onclick="switchMainTab('PARTNER', this)"><i class="bi bi-building"></i> 거래처 정보 관리</div>
    <div class="nav-item" onclick="switchMainTab('PRODUCT', this)"><i class="bi bi-tags"></i> 상품 마스터 관리</div>
  </aside>

  <main>
    <header>
      <h2 id="pageTitle" style="font-size:1.25rem;">입고 프로세스 관리</h2>
      <div id="headerActions" style="display:flex; gap:8px;"></div>
    </header>

    <div class="content">
      <div class="sub-tabs" id="subTabs" style="display:none;"></div>

      <!-- 일자별 / 월별 선택 필터 카드 -->
      <div class="filter-card" id="dateFilterCard">
        <div class="filter-group">
          <span id="dateFilterLabel"><i class="bi bi-calendar-range"></i> 기준 일자:</span>
          <input type="date" id="filter_start" class="filter-input" onchange="loadData()" />
          <span>~</span>
          <input type="date" id="filter_end" class="filter-input" onchange="loadData()" />
        </div>
        <div class="filter-btn-group">
          <button class="filter-btn active" id="btn_all" onclick="setDateFilter('ALL')">전체</button>
          <button class="filter-btn" id="btn_this_month" onclick="setDateFilter('THIS_MONTH')">당월</button>
          <button class="filter-btn" id="btn_last_month" onclick="setDateFilter('LAST_MONTH')">전월</button>
        </div>
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

  <!-- 1. 입고 등록 모달 (브랜드/품목명 분리) -->
  <div class="modal-overlay" id="inboundModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px;">신규 축산물 입고요청</h3>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>입고 일자</label><input type="date" id="in_date" class="form-control" /></div>
        <div class="form-group" style="flex:1;"><label>매입처 상호</label><input type="text" id="in_vendor" class="form-control" value="(주)글로벌미트" /></div>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>브랜드 / 가공장</label><input type="text" id="in_brand" class="form-control" placeholder="예: 올리멜, 엑셀, 도드람" value="올리멜" /></div>
        <div class="form-group" style="flex:1;"><label>품목 및 부위명</label><input type="text" id="in_item" class="form-control" placeholder="예: 삼겹살, 목심" value="삼겹살" /></div>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>보관 형태</label><select id="in_storage" class="form-control"><option value="냉장">냉장</option><option value="냉동">냉동</option></select></div>
        <div class="form-group" style="flex:1;"><label>보관창고</label><input type="text" id="in_warehouse" class="form-control" value="광주냉장창고" /></div>
      </div>
      <div class="form-group"><label>이력번호 (12자리)</label><input type="text" id="in_trace" class="form-control" value="802410290114" /></div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>박스 수량</label><input type="number" id="in_box" class="form-control" value="50" /></div>
        <div class="form-group" style="flex:1;"><label>실중량(kg)</label><input type="number" step="0.1" id="in_weight" class="form-control" value="1020.5" /></div>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>매입단가(원/kg)</label><input type="number" id="in_cost" class="form-control" value="8400" /></div>
        <div class="form-group" style="flex:1;"><label>유통기한</label><input type="date" id="in_exp" class="form-control" value="2026-09-15" /></div>
      </div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal('inboundModal')">취소</button>
        <button class="btn btn-primary" onclick="submitInbound()">저장</button>
      </div>
    </div>
  </div>

  <!-- 2. 거래처 등록 모달 -->
  <div class="modal-overlay" id="partnerModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px;">신규 거래처 등록</h3>
      <div class="form-group"><label>거래처/상호명</label><input type="text" id="p_name" class="form-control" placeholder="예: (주)한국축산유통" /></div>
      <div class="form-group"><label>분류 구분</label>
        <select id="p_type" class="form-control">
          <option value="VENDOR">매입처 (Vendor)</option>
          <option value="CUSTOMER">매출처 (Customer)</option>
          <option value="WAREHOUSE">보관창고 (Warehouse)</option>
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

  <!-- 3. 상품 마스터 등록 모달 (브랜드/품목명 분리) -->
  <div class="modal-overlay" id="productModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px;">신규 상품 마스터 등록</h3>
      <div class="form-group"><label>품목 코드 (SKU)</label><input type="text" id="prod_sku" class="form-control" placeholder="예: PK-CAN-02" /></div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>축종</label><select id="prod_species" class="form-control"><option>돼지</option><option>소</option><option>닭</option><option>기타</option></select></div>
        <div class="form-group" style="flex:1;"><label>브랜드 / 가공장</label><input type="text" id="prod_brand" class="form-control" placeholder="예: 올리멜, 도드람" /></div>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>부위명</label><input type="text" id="prod_cut" class="form-control" placeholder="예: 삼겹살, 갈비" /></div>
        <div class="form-group" style="flex:1;"><label>원산지</label><input type="text" id="prod_origin" class="form-control" placeholder="예: 캐나다, 국내산" /></div>
      </div>
      <div class="form-group"><label>보관 형태</label><select id="prod_storage" class="form-control"><option value="냉장">냉장 (Chilled)</option><option value="냉동">냉동 (Frozen)</option></select></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal('productModal')">취소</button>
        <button class="btn btn-primary" onclick="submitProduct()">상품 등록</button>
      </div>
    </div>
  </div>

  <!-- 4. 수량/중량/단가/일자 수정 모달 -->
  <div class="modal-overlay" id="editModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px;" id="editModalTitle">정보 수정</h3>
      <input type="hidden" id="edit_type" />
      <input type="hidden" id="edit_id" />
      <div class="form-group"><label id="edit_date_label">일자</label><input type="date" id="edit_date" class="form-control" /></div>
      <div id="edit_inbound_fields" style="display:none;">
        <div style="display:flex; gap:10px;">
          <div class="form-group" style="flex:1;"><label>브랜드</label><input type="text" id="edit_brand" class="form-control" /></div>
          <div class="form-group" style="flex:1;"><label>품목명</label><input type="text" id="edit_item_name" class="form-control" /></div>
        </div>
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

  <!-- 5. 재고 감량/폐기 모달 -->
  <div class="modal-overlay" id="adjustModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px; color:#dc2626;">재고 감량 및 클레임 조정</h3>
      <input type="hidden" id="adj_lot_id" />
      <div class="form-group"><label>조정 사유 구분</label>
        <select id="adj_type" class="form-control">
          <option value="갈변/변질폐기">갈변/변질 폐기</option>
          <option value="감량/드립로스">자연 감량 / 드립 로스</option>
          <option value="클레임반품">고객사 클레임 반품</option>
          <option value="실사조정">재고 실사 차이 조정</option>
        </select>
      </div>
      <div style="display:flex; gap:10px;">
        <div class="form-group" style="flex:1;"><label>차감 박스</label><input type="number" id="adj_box" class="form-control" value="1" /></div>
        <div class="form-group" style="flex:1;"><label>차감 중량(kg)</label><input type="number" step="0.1" id="adj_weight" class="form-control" value="20.4" /></div>
      </div>
      <div class="form-group"><label>상세 사유</label><input type="text" id="adj_reason" class="form-control" placeholder="예: 표면 갈변으로 폐기 처리" /></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeModal('adjustModal')">취소</button>
        <button class="btn btn-danger" onclick="submitAdjust()">조정 적용</button>
      </div>
    </div>
  </div>

  <input type="file" id="excelFileInput" style="display:none;" onchange="uploadExcelFile(event)" accept=".xlsx, .xls" />

  <script>
    let currentMain = 'INBOUND', currentSub = 'IN_REQUEST';
    let activeMonthFilter = null;

    // 기본 오늘 날짜 세팅
    document.getElementById('in_date').value = new Date().toISOString().split('T')[0];

    function setDateFilter(mode) {
      document.getElementById('btn_all').classList.remove('active');
      document.getElementById('btn_this_month').classList.remove('active');
      document.getElementById('btn_last_month').classList.remove('active');

      const startInput = document.getElementById('filter_start');
      const endInput = document.getElementById('filter_end');

      if (mode === 'ALL') {
        document.getElementById('btn_all').classList.add('active');
        activeMonthFilter = null;
        startInput.value = '';
        endInput.value = '';
      } else if (mode === 'THIS_MONTH') {
        document.getElementById('btn_this_month').classList.add('active');
        const now = new Date();
        const y = now.getFullYear();
        const m = String(now.getMonth() + 1).padStart(2, '0');
        activeMonthFilter = `${y}-${m}`;
        startInput.value = `${y}-${m}-01`;
        endInput.value = new Date(y, now.getMonth() + 1, 0).toISOString().split('T')[0];
      } else if (mode === 'LAST_MONTH') {
        document.getElementById('btn_last_month').classList.add('active');
        const now = new Date();
        now.setMonth(now.getMonth() - 1);
        const y = now.getFullYear();
        const m = String(now.getMonth() + 1).padStart(2, '0');
        activeMonthFilter = `${y}-${m}`;
        startInput.value = `${y}-${m}-01`;
        endInput.value = new Date(y, now.getMonth() + 1, 0).toISOString().split('T')[0];
      }
      loadData();
    }

    function switchMainTab(tab, el) {
      currentMain = tab;
      document.querySelectorAll('.nav-item').forEach(e => e.classList.remove('active'));
      el.classList.add('active');

      const subTabs = document.getElementById('subTabs');
      const headerActions = document.getElementById('headerActions');
      const pageTitle = document.getElementById('pageTitle');
      const dateCard = document.getElementById('dateFilterCard');

      if (tab === 'INBOUND') {
        currentSub = 'IN_REQUEST';
        pageTitle.innerText = '입고 프로세스 관리';
        subTabs.style.display = 'flex';
        dateCard.style.display = 'flex';
        document.getElementById('dateFilterLabel').innerHTML = '<i class="bi bi-calendar-range"></i> 입고 일자:';
        headerActions.innerHTML = `
          <button class="btn btn-outline" onclick="document.getElementById('excelFileInput').click()"><i class="bi bi-file-earmark-excel"></i> 엑셀 일괄입고</button>
          <button class="btn btn-primary" onclick="openModal('inboundModal')"><i class="bi bi-plus-lg"></i> 신규 입고요청</button>
        `;
      } else if (tab === 'OUTBOUND') {
        currentSub = 'OUT_STOCK';
        pageTitle.innerText = '출고 프로세스 관리 (현 재고장 연동)';
        subTabs.style.display = 'flex';
        dateCard.style.display = 'flex';
        document.getElementById('dateFilterLabel').innerHTML = '<i class="bi bi-calendar-range"></i> 출고 일자:';
        headerActions.innerHTML = '';
      } else if (tab === 'PARTNER') {
        pageTitle.innerText = '거래처 정보 관리 (매입처 / 매출처 / 창고)';
        subTabs.style.display = 'none';
        dateCard.style.display = 'none';
        headerActions.innerHTML = `<button class="btn btn-primary" onclick="openModal('partnerModal')"><i class="bi bi-plus-lg"></i> 신규 거래처 등록</button>`;
      } else if (tab === 'PRODUCT') {
        pageTitle.innerText = '상품 마스터 관리 (축종 / 브랜드 / 부위 / 원산지)';
        subTabs.style.display = 'none';
        dateCard.style.display = 'none';
        headerActions.innerHTML = `<button class="btn btn-primary" onclick="openModal('productModal')"><i class="bi bi-plus-lg"></i> 신규 상품 등록</button>`;
      }
      renderSubTabs();
      loadData();
    }

    function renderSubTabs() {
      const container = document.getElementById('subTabs');
      container.innerHTML = '';
      if (currentMain === 'INBOUND') {
        const steps = [{k:'IN_REQUEST',n:'1. 입고요청리스트'},{k:'IN_CONFIRM',n:'2. 입고확정리스트'},{k:'IN_DONE',n:'3. 입고완료리스트'}];
        steps.forEach(s => { container.innerHTML += `<div class="sub-tab ${currentSub===s.k?'active':''}" onclick="setSubTab('${s.k}')">${s.n}</div>`; });
      } else if (currentMain === 'OUTBOUND') {
        const steps = [{k:'OUT_STOCK',n:'1. 출하요청리스트 (현 재고장)'},{k:'OUT_REQUEST',n:'2. 출고요청리스트'},{k:'OUT_CONFIRM',n:'3. 출고확정리스트'},{k:'OUT_DONE',n:'4. 출고완료리스트'}];
        steps.forEach(s => { container.innerHTML += `<div class="sub-tab ${currentSub===s.k?'active':''}" onclick="setSubTab('${s.k}')">${s.n}</div>`; });
      }
    }

    function setSubTab(k) { currentSub = k; renderSubTabs(); loadData(); }

    async function loadData() {
      const head = document.getElementById('tableHead'), body = document.getElementById('tableBody');
      body.innerHTML = '<tr><td colspan="12" style="text-align:center; padding:15px;">데이터 조회 중...</td></tr>';

      const sDate = document.getElementById('filter_start').value;
      const eDate = document.getElementById('filter_end').value;
      let dateQuery = '';
      if (activeMonthFilter) {
        dateQuery = `&month=${activeMonthFilter}`;
      } else {
        if (sDate) dateQuery += `&start_date=${sDate}`;
        if (eDate) dateQuery += `&end_date=${eDate}`;
      }

      // --- [1. 입고 관리 뷰] ---
      if (currentMain === 'INBOUND') {
        head.innerHTML = '<tr><th>전표번호</th><th>입고일자</th><th>매입처</th><th>브랜드</th><th>품목/부위</th><th>보관</th><th>이력번호</th><th style="text-align:right;">박스</th><th style="text-align:right;">실중량(kg)</th><th style="text-align:right;">단가(kg)</th><th style="text-align:right;">총 매입금액</th><th>유통기한</th><th style="text-align:center;">작업 관리</th></tr>';
        const res = await fetch(`/api/inbounds?status=${currentSub}${dateQuery}`), list = await res.json();
        body.innerHTML = list.length ? list.map(i => {
          let btns = '';
          if (currentSub === 'IN_REQUEST') {
            btns = `
              <button class="btn btn-primary" onclick="advIn(${i.id})">확정 <i class="bi bi-arrow-right"></i></button>
              <button class="btn btn-outline" onclick="openEditModal('INBOUND', ${i.id}, '${i.inbound_date}', '${i.brand||''}', '${i.item_name}', ${i.box_qty}, ${i.weight_kg}, ${i.cost_per_kg})"><i class="bi bi-pencil"></i></button>
              <button class="btn btn-outline" style="color:#dc2626;" onclick="revertIn(${i.id}, '요청 취소(삭제)')"><i class="bi bi-trash"></i></button>
            `;
          } else if (currentSub === 'IN_CONFIRM') {
            btns = `
              <button class="btn btn-success" onclick="advIn(${i.id})"><i class="bi bi-check-lg"></i> 입고완료</button>
              <button class="btn btn-outline" onclick="openEditModal('INBOUND', ${i.id}, '${i.inbound_date}', '${i.brand||''}', '${i.item_name}', ${i.box_qty}, ${i.weight_kg}, ${i.cost_per_kg})"><i class="bi bi-pencil"></i></button>
              <button class="btn btn-outline" style="color:#d97706;" onclick="revertIn(${i.id}, '이전 단계 반려')"><i class="bi bi-arrow-counterclockwise"></i> 반려</button>
            `;
          } else {
            btns = `
              <span class="badge" style="background:#dcfce7; color:#15803d; margin-right:4px;">완료</span>
              <button class="btn btn-outline" style="color:#dc2626;" onclick="revertIn(${i.id}, '입고 취소 및 재고 원복')"><i class="bi bi-arrow-counterclockwise"></i> 작업취소</button>
            `;
          }
          return `
            <tr>
              <td><strong>${i.inbound_no}</strong></td>
              <td>${i.inbound_date}</td>
              <td>${i.vendor}</td>
              <td><span class="brand-tag">${i.brand || '-'}</span></td>
              <td><strong>${i.item_name}</strong></td>
              <td><span class="badge ${i.storage_type==='냉장'?'badge-chilled':'badge-frozen'}">${i.storage_type}</span></td>
              <td><code>${i.trace_no}</code></td>
              <td style="text-align:right;"><strong>${i.box_qty}</strong> Box</td>
              <td style="text-align:right;"><strong>${i.weight_kg.toLocaleString()}</strong> kg</td>
              <td style="text-align:right;">${i.cost_per_kg.toLocaleString()}원</td>
              <td style="text-align:right;" class="amount-highlight">${i.total_amount.toLocaleString()}원</td>
              <td>${i.exp_date}</td>
              <td style="text-align:center; white-space:nowrap;">${btns}</td>
            </tr>
          `;
        }).join('') : '<tr><td colspan="13" style="text-align:center; padding:20px; color:#94a3b8;">대기 중인 전표가 없습니다.</td></tr>';

      // --- [2. 출고 및 현 재고장 뷰] ---
      } else if (currentMain === 'OUTBOUND') {
        if (currentSub === 'OUT_STOCK') {
          head.innerHTML = '<tr><th>로트ID</th><th>브랜드</th><th>품목/부위</th><th>보관</th><th>이력번호</th><th style="text-align:right;">보유 박스</th><th style="text-align:right;">보유 중량(kg)</th><th style="text-align:right;">원가 단가</th><th style="text-align:right;">재고 평가금액</th><th>창고</th><th>유통기한</th><th style="text-align:center;">관리 작업</th></tr>';
          const res = await fetch('/api/inventory'), list = await res.json();
          body.innerHTML = list.length ? list.map(l => {
            const evalAmount = Math.round(l.current_weight_kg * l.cost_per_kg);
            return `
              <tr>
                <td>#${l.id}</td>
                <td><span class="brand-tag">${l.brand || '-'}</span></td>
                <td><strong>${l.item_name}</strong></td>
                <td><span class="badge ${l.storage_type==='냉장'?'badge-chilled':'badge-frozen'}">${l.storage_type}</span></td>
                <td><code>${l.trace_no}</code></td>
                <td style="text-align:right;"><strong>${l.current_box_qty}</strong> Box</td>
                <td style="text-align:right;"><strong>${l.current_weight_kg.toLocaleString()}</strong> kg</td>
                <td style="text-align:right;">${l.cost_per_kg.toLocaleString()}원</td>
                <td style="text-align:right;" class="amount-highlight">${evalAmount.toLocaleString()}원</td>
                <td>${l.warehouse}</td>
                <td>${l.exp_date}</td>
                <td style="text-align:center; white-space:nowrap;">
                  <button class="btn btn-warning" onclick="reqOut(${l.id},${l.current_box_qty},${l.cost_per_kg})"><i class="bi bi-cart-plus"></i> 출하요청</button>
                  <button class="btn btn-outline" style="color:#dc2626;" onclick="openAdjustModal(${l.id})"><i class="bi bi-scissors"></i> 감량/폐기</button>
                </td>
              </tr>
            `;
          }).join('') : '<tr><td colspan="12" style="text-align:center; padding:20px; color:#94a3b8;">보유 재고가 없습니다.</td></tr>';
        } else {
          head.innerHTML = '<tr><th>출고번호</th><th>출고일자</th><th>매출처</th><th>브랜드</th><th>품목명</th><th>이력번호</th><th style="text-align:right;">출고 박스</th><th style="text-align:right;">실중량(kg)</th><th style="text-align:right;">판매 단가(kg)</th><th style="text-align:right;">총 매출금액</th><th>유통기한</th><th style="text-align:center;">작업 관리</th></tr>';
          const res = await fetch(`/api/outbounds?status=${currentSub}${dateQuery}`), list = await res.json();
          body.innerHTML = list.length ? list.map(o => {
            let btns = '';
            if (currentSub === 'OUT_REQUEST') {
              btns = `
                <button class="btn btn-primary" onclick="advOut(${o.id})">확정 <i class="bi bi-arrow-right"></i></button>
                <button class="btn btn-outline" onclick="openEditModal('OUTBOUND', ${o.id}, '${o.outbound_date}', '${o.brand||''}', '${o.item_name}', ${o.box_qty}, ${o.weight_kg}, ${o.unit_price_kg})"><i class="bi bi-pencil"></i></button>
                <button class="btn btn-outline" style="color:#dc2626;" onclick="revertOut(${o.id}, '요청 취소(삭제)')"><i class="bi bi-trash"></i></button>
              `;
            } else if (currentSub === 'OUT_CONFIRM') {
              btns = `
                <button class="btn btn-success" onclick="advOut(${o.id})"><i class="bi bi-truck"></i> 출고완료</button>
                <button class="btn btn-outline" onclick="openEditModal('OUTBOUND', ${o.id}, '${o.outbound_date}', '${o.brand||''}', '${o.item_name}', ${o.box_qty}, ${o.weight_kg}, ${o.unit_price_kg})"><i class="bi bi-pencil"></i></button>
                <button class="btn btn-outline" style="color:#d97706;" onclick="revertOut(${o.id}, '이전 단계 반려')"><i class="bi bi-arrow-counterclockwise"></i> 반려</button>
              `;
            } else {
              btns = `
                <span class="badge" style="background:#dbeafe; color:#1d4ed8; margin-right:4px;">출고완료</span>
                <button class="btn btn-outline" style="color:#dc2626;" onclick="revertOut(${o.id}, '출고 취소 및 재고 원복')"><i class="bi bi-arrow-counterclockwise"></i> 작업취소</button>
              `;
            }
            return `
              <tr>
                <td><strong>${o.outbound_no}</strong></td>
                <td>${o.outbound_date}</td>
                <td>${o.customer}</td>
                <td><span class="brand-tag">${o.brand || '-'}</span></td>
                <td><strong>${o.item_name}</strong></td>
                <td><code>${o.trace_no}</code></td>
                <td style="text-align:right;"><strong>${o.box_qty}</strong> Box</td>
                <td style="text-align:right;"><strong>${o.weight_kg.toLocaleString()}</strong> kg</td>
                <td style="text-align:right;">${o.unit_price_kg.toLocaleString()}원</td>
                <td style="text-align:right;" class="amount-highlight">${o.total_amount.toLocaleString()}원</td>
                <td>${o.exp_date}</td>
                <td style="text-align:center; white-space:nowrap;">${btns}</td>
              </tr>
            `;
          }).join('') : '<tr><td colspan="12" style="text-align:center; padding:20px; color:#94a3b8;">대기 중인 출고 건이 없습니다.</td></tr>';
        }

      // --- [3. 거래처 마스터 뷰] ---
      } else if (currentMain === 'PARTNER') {
        head.innerHTML = '<tr><th>ID</th><th>상호명</th><th>구분</th><th>사업자번호</th><th>담당자</th><th>연락처</th><th>사업장/창고 주소</th><th style="text-align:center;">관리</th></tr>';
        const res = await fetch('/api/partners'), list = await res.json();
        body.innerHTML = list.length ? list.map(p => `
          <tr>
            <td>#${p.id}</td>
            <td><strong>${p.name}</strong></td>
            <td><span class="badge ${p.type==='VENDOR'?'badge-vendor':p.type==='CUSTOMER'?'badge-customer':'badge-warehouse'}">${p.type==='VENDOR'?'매입처':p.type==='CUSTOMER'?'매출처':'창고'}</span></td>
            <td>${p.biz_no || '-'}</td>
            <td>${p.contact_person || '-'}</td>
            <td>${p.phone || '-'}</td>
            <td>${p.address || '-'}</td>
            <td style="text-align:center;"><button class="btn btn-outline" style="color:#dc2626;" onclick="deletePartner(${p.id})"><i class="bi bi-trash"></i> 삭제</button></td>
          </tr>
        `).join('') : '<tr><td colspan="8" style="text-align:center; padding:20px; color:#94a3b8;">등록된 거래처가 없습니다.</td></tr>';

      // --- [4. 상품 마스터 뷰] ---
      } else if (currentMain === 'PRODUCT') {
        head.innerHTML = '<tr><th>ID</th><th>품목코드(SKU)</th><th>축종</th><th>브랜드/가공장</th><th>부위명</th><th>원산지</th><th>보관형태</th><th style="text-align:center;">관리</th></tr>';
        const res = await fetch('/api/products'), list = await res.json();
        body.innerHTML = list.length ? list.map(pr => `
          <tr>
            <td>#${pr.id}</td>
            <td><code>${pr.sku_code}</code></td>
            <td>${pr.species}</td>
            <td><span class="brand-tag">${pr.brand || '-'}</span></td>
            <td><strong>${pr.cut_name}</strong></td>
            <td>${pr.origin}</td>
            <td><span class="badge ${pr.storage_type==='냉장'?'badge-chilled':'badge-frozen'}">${pr.storage_type}</span></td>
            <td style="text-align:center;"><button class="btn btn-outline" style="color:#dc2626;" onclick="deleteProduct(${pr.id})"><i class="bi bi-trash"></i> 삭제</button></td>
          </tr>
        `).join('') : '<tr><td colspan="8" style="text-align:center; padding:20px; color:#94a3b8;">등록된 상품 마스터가 없습니다.</td></tr>';
      }
    }

    // 단계 전이 및 롤백
    async function advIn(id) { const r = await fetch(`/api/inbounds/${id}/advance`,{method:'POST'}); const d = await r.json(); alert(d.message); loadData(); }
    async function advOut(id) { const r = await fetch(`/api/outbounds/${id}/advance`,{method:'POST'}); const d = await r.json(); alert(d.message); loadData(); }
    async function revertIn(id, msg) { if(!confirm(`정말 [${msg}] 작업을 진행하시겠습니까?`)) return; const r = await fetch(`/api/inbounds/${id}/revert`,{method:'POST'}); const d = await r.json(); alert(d.message); loadData(); }
    async function revertOut(id, msg) { if(!confirm(`정말 [${msg}] 작업을 진행하시겠습니까?`)) return; const r = await fetch(`/api/outbounds/${id}/revert`,{method:'POST'}); const d = await r.json(); alert(d.message); loadData(); }

    // 수정 모달
    function openEditModal(type, id, dt, brand, item_name, box, weight, price) {
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
        document.getElementById('edit_brand').value = brand;
        document.getElementById('edit_item_name').value = item_name;
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
          brand: document.getElementById('edit_brand').value,
          item_name: document.getElementById('edit_item_name').value,
          box_qty: Number(document.getElementById('edit_box').value),
          weight_kg: Number(document.getElementById('edit_weight').value),
          cost_per_kg: Number(document.getElementById('edit_price').value)
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

    // 출하요청 등록
    async function reqOut(lotId, maxBox, costPrice) {
      const cust = prompt("출하할 매출처 상호명을 입력하세요:", "(주)하남돼지집"); if(!cust) return;
      const box = prompt(`출하 요청 박스 수를 입력하세요 (현재 보유: ${maxBox} Box):`, "10"); if(!box || isNaN(box) || Number(box)<=0 || Number(box)>maxBox) { alert("박스 수가 올바르지 않습니다."); return; }
      const price = prompt("kg당 판매단가(원)를 입력하세요:", String(costPrice + 2000)); if(!price) return;
      const r = await fetch('/api/outbounds/create-from-stock', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({lot_id:lotId, customer:cust, box_qty:Number(box), unit_price_kg:Number(price)})});
      const d = await r.json(); alert(d.message); setSubTab('OUT_REQUEST');
    }

    // 거래처 등록 / 삭제
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

    // 상품 등록 / 삭제
    async function submitProduct() {
      const p = {
        sku_code: document.getElementById('prod_sku').value,
        species: document.getElementById('prod_species').value,
        brand: document.getElementById('prod_brand').value,
        cut_name: document.getElementById('prod_cut').value,
        origin: document.getElementById('prod_origin').value,
        storage_type: document.getElementById('prod_storage').value
      };
      if(!p.sku_code || !p.cut_name) { alert("품목코드와 부위명은 필수입니다."); return; }
      const r = await fetch('/api/products', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p)});
      const d = await r.json();
      if(r.ok) { alert(d.message); closeModal('productModal'); loadData(); } else { alert(d.detail); }
    }
    async function deleteProduct(id) {
      if(!confirm("정말 이 상품 마스터를 삭제하시겠습니까?")) return;
      const r = await fetch(`/api/products/${id}`, {method:'DELETE'});
      const d = await r.json(); alert(d.message); loadData();
    }

    function openModal(id) { document.getElementById(id).style.display = 'flex'; }
    function closeModal(id) { document.getElementById(id).style.display = 'none'; }

    // 입고 개별 저장
    async function submitInbound() {
      const p = {
        inbound_date: document.getElementById('in_date').value,
        vendor: document.getElementById('in_vendor').value,
        brand: document.getElementById('in_brand').value,
        item_name: document.getElementById('in_item').value,
        storage_type: document.getElementById('in_storage').value,
        trace_no: document.getElementById('in_trace').value,
        box_qty: Number(document.getElementById('in_box').value),
        weight_kg: Number(document.getElementById('in_weight').value),
        cost_per_kg: Number(document.getElementById('in_cost').value),
        warehouse: document.getElementById('in_warehouse').value,
        exp_date: document.getElementById('in_exp').value
      };
      const r = await fetch('/api/inbounds', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p)});
      const d = await r.json(); alert(d.message); closeModal('inboundModal'); setSubTab('IN_REQUEST');
    }

    // 재고 감량/폐기 저장
    function openAdjustModal(lotId) { document.getElementById('adj_lot_id').value = lotId; openModal('adjustModal'); }
    async function submitAdjust() {
      const p = { lot_id: Number(document.getElementById('adj_lot_id').value), adj_type: document.getElementById('adj_type').value, adj_box: Number(document.getElementById('adj_box').value), adj_weight: Number(document.getElementById('adj_weight').value), reason: document.getElementById('adj_reason').value };
      const r = await fetch('/api/inventory/adjust', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p)});
      const d = await r.json(); alert(d.message); closeModal('adjustModal'); loadData();
    }

    // 엑셀 업로드
    async function uploadExcelFile(e) {
      const file = e.target.files[0]; if (!file) return;
      const formData = new FormData(); formData.append('file', file);
      const r = await fetch('/api/inbounds/upload-excel', { method: 'POST', body: formData });
      const d = await r.json(); alert(d.message || d.detail); e.target.value = ''; setSubTab('IN_REQUEST');
    }

    window.addEventListener('DOMContentLoaded', () => { renderSubTabs(); loadData(); });
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
