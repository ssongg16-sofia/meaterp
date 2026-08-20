import io
import os
from datetime import datetime, date
from typing import List, Optional

from fastapi import FastAPI, HTTPException, Depends, UploadFile, File
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
# 1. DB 설정 (SQLite 파일 기반 영속성)
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
# 2. ORM 테이블 정의
# -----------------------------------------------------------------------------
class InboundRecord(Base):
    __tablename__ = "inbounds"
    id = Column(Integer, primary_key=True, index=True)
    inbound_no = Column(String(50), unique=True, index=True)
    vendor = Column(String(100), nullable=False)
    item_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
    trace_no = Column(String(50), nullable=False, index=True)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    cost_per_kg = Column(Float, nullable=False)
    warehouse = Column(String(50), default="광주냉장창고")
    exp_date = Column(String(20), nullable=False)
    status = Column(String(20), default="IN_REQUEST")  # IN_REQUEST -> IN_CONFIRM -> IN_DONE
    created_at = Column(DateTime, default=datetime.now)

class InventoryLot(Base):
    __tablename__ = "inventory_lots"
    id = Column(Integer, primary_key=True, index=True)
    sku_code = Column(String(50), nullable=False)
    item_name = Column(String(100), nullable=False)
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
    lot_id = Column(Integer, nullable=False)
    customer = Column(String(100), nullable=False)
    item_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
    trace_no = Column(String(50), nullable=False)
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    unit_price_kg = Column(Float, nullable=False)
    exp_date = Column(String(20), nullable=False)
    warehouse = Column(String(50), nullable=False)
    status = Column(String(20), default="OUT_REQUEST")  # OUT_REQUEST -> OUT_CONFIRM -> OUT_DONE
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

# 초기 샘플 데이터 시딩
def init_sample_data():
    db = SessionLocal()
    if db.query(InventoryLot).count() == 0:
        lots = [
            InventoryLot(sku_code="PK-CAN-01", item_name="돼지 삼겹살(올리멜)", storage_type="냉장", trace_no="802410290114", current_box_qty=45, current_weight_kg=912.4, cost_per_kg=8400, warehouse="광주냉장창고", exp_date="2026-09-02"),
            InventoryLot(sku_code="BF-USA-05", item_name="소 척아이롤(엑셀)", storage_type="냉동", trace_no="100392019482", current_box_qty=120, current_weight_kg=2450.0, cost_per_kg=14500, warehouse="용인냉동센터", exp_date="2027-04-15"),
            InventoryLot(sku_code="PK-KOR-01", item_name="국내산 생삼겹(도드람)", storage_type="냉장", trace_no="002194820192", current_box_qty=15, current_weight_kg=305.5, cost_per_kg=16800, warehouse="광주냉장창고", exp_date="2026-08-25")
        ]
        db.add_all(lots)
        db.commit()
    db.close()

init_sample_data()

# -----------------------------------------------------------------------------
# 3. FastAPI 라우터 및 상태 전이 / 수정 / 롤백 API
# -----------------------------------------------------------------------------
app = FastAPI(title="MeatFlow Enterprise ERP")

class InboundCreate(BaseModel):
    vendor: str
    item_name: str
    storage_type: str
    trace_no: str
    box_qty: int
    weight_kg: float
    cost_per_kg: float
    warehouse: str
    exp_date: str

class UpdateQtyWeight(BaseModel):
    box_qty: int
    weight_kg: float

class OutboundCreate(BaseModel):
    lot_id: int
    customer: str
    box_qty: int
    unit_price_kg: float

class AdjustCreate(BaseModel):
    lot_id: int
    adj_type: str
    adj_box: int
    adj_weight: float
    reason: Optional[str] = ""

# 입고 데이터 목록 조회
@app.get("/api/inbounds")
def get_inbounds(status: str, db: Session = Depends(get_db)):
    return db.query(InboundRecord).filter(InboundRecord.status == status).order_by(desc(InboundRecord.id)).all()

# 신규 입고요청 등록
@app.post("/api/inbounds")
def create_inbound(req: InboundCreate, db: Session = Depends(get_db)):
    inbound_no = f"IN-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    item = InboundRecord(
        inbound_no=inbound_no,
        vendor=req.vendor,
        item_name=req.item_name,
        storage_type=req.storage_type,
        trace_no=req.trace_no,
        box_qty=req.box_qty,
        weight_kg=req.weight_kg,
        cost_per_kg=req.cost_per_kg,
        warehouse=req.warehouse,
        exp_date=req.exp_date,
        status="IN_REQUEST"
    )
    db.add(item)
    db.commit()
    return {"message": "입고요청 등록 완료", "inbound_no": inbound_no}

# [기능추가] 입고 전표 수량/중량 수정
@app.put("/api/inbounds/{inbound_id}")
def update_inbound(inbound_id: int, req: UpdateQtyWeight, db: Session = Depends(get_db)):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).first()
    if not inbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    if inbound.status == "IN_DONE":
        raise HTTPException(status_code=400, detail="이미 입고 완료된 전표는 [작업 취소] 후 수정해야 합니다.")

    inbound.box_qty = req.box_qty
    inbound.weight_kg = req.weight_kg
    db.commit()
    return {"message": "입고 수량 및 중량이 수정되었습니다."}

# [기능추가] 입고 단계 취소/반려/원복
@app.post("/api/inbounds/{inbound_id}/revert")
def revert_inbound(inbound_id: int, db: Session = Depends(get_db)):
    inbound = db.query(InboundRecord).filter(InboundRecord.id == inbound_id).first()
    if not inbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    if inbound.status == "IN_REQUEST":
        # 1단계 취소 시 삭제
        db.delete(inbound)
        msg = "입고요청 전표가 취소(삭제)되었습니다."
    elif inbound.status == "IN_CONFIRM":
        # 2단계 취소 시 1단계로 반려
        inbound.status = "IN_REQUEST"
        msg = "입고확정이 취소되어 [입고요청리스트]로 반려되었습니다."
    elif inbound.status == "IN_DONE":
        # 3단계 취소 시 실재고 차감 후 2단계로 복구
        lot = db.query(InventoryLot).filter(
            InventoryLot.trace_no == inbound.trace_no,
            InventoryLot.warehouse == inbound.warehouse,
            InventoryLot.exp_date == inbound.exp_date
        ).first()
        if lot:
            lot.current_box_qty = max(0, lot.current_box_qty - inbound.box_qty)
            lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - inbound.weight_kg, 2))
        inbound.status = "IN_CONFIRM"
        msg = "입고완료가 취소되어 실재고가 원복(차감)되었으며 [입고확정리스트]로 복구되었습니다."
    else:
        msg = "처리할 수 없는 상태입니다."

    db.commit()
    return {"message": msg}

# 엑셀 일괄 업로드
@app.post("/api/inbounds/upload-excel")
async def upload_inbound_excel(file: UploadFile = File(...), db: Session = Depends(get_db)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀(.xlsx) 파일만 업로드할 수 있습니다.")
    try:
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=True)
        sheet = wb.active

        count = 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            vendor = str(row[0]).strip()
            item_name = str(row[1]).strip()
            storage_type = str(row[2]).strip()
            trace_no = str(row[3]).strip()
            box_qty = int(row[4])
            weight_kg = float(row[5])
            cost_per_kg = float(row[6])
            warehouse = str(row[7]).strip() if len(row) > 7 and row[7] else "광주냉장창고"
            exp_date_raw = row[8] if len(row) > 8 else "2026-12-31"

            if isinstance(exp_date_raw, (datetime, date)):
                exp_date_str = exp_date_raw.strftime('%Y-%m-%d')
            else:
                exp_date_str = str(exp_date_raw).strip()

            inbound_no = f"IN-{datetime.now().strftime('%Y%m%d%H%M%S')}-{count+1}"
            inbound = InboundRecord(
                inbound_no=inbound_no,
                vendor=vendor,
                item_name=item_name,
                storage_type=storage_type,
                trace_no=trace_no,
                box_qty=box_qty,
                weight_kg=weight_kg,
                cost_per_kg=cost_per_kg,
                warehouse=warehouse,
                exp_date=exp_date_str,
                status="IN_REQUEST"
            )
            db.add(inbound)
            count += 1

        db.commit()
        return {"message": f"총 {count}건의 엑셀 입고 요청이 등록되었습니다."}
    except Exception as e:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"엑셀 처리 오류: {str(e)}")

# 입고 단계 이동 (IN_REQUEST -> IN_CONFIRM -> IN_DONE: 재고 합산)
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
        lot = db.query(InventoryLot).filter(
            InventoryLot.trace_no == inbound.trace_no,
            InventoryLot.warehouse == inbound.warehouse,
            InventoryLot.exp_date == inbound.exp_date
        ).first()

        if lot:
            lot.current_box_qty += inbound.box_qty
            lot.current_weight_kg = round(lot.current_weight_kg + inbound.weight_kg, 2)
            lot.cost_per_kg = inbound.cost_per_kg
        else:
            sku = "SKU-" + inbound.trace_no[:6]
            new_lot = InventoryLot(
                sku_code=sku,
                item_name=inbound.item_name,
                storage_type=inbound.storage_type,
                trace_no=inbound.trace_no,
                current_box_qty=inbound.box_qty,
                current_weight_kg=inbound.weight_kg,
                cost_per_kg=inbound.cost_per_kg,
                warehouse=inbound.warehouse,
                exp_date=inbound.exp_date
            )
            db.add(new_lot)
        msg = "입고완료 처리 및 실재고 반영이 완료되었습니다."
    else:
        msg = "이미 완료된 전표입니다."

    db.commit()
    return {"message": msg}

# 현 재고장 (출하요청리스트) 조회
@app.get("/api/inventory")
def get_inventory(db: Session = Depends(get_db)):
    return db.query(InventoryLot).filter(InventoryLot.current_box_qty > 0).order_by(InventoryLot.exp_date).all()

# 출고 목록 조회
@app.get("/api/outbounds")
def get_outbounds(status: str, db: Session = Depends(get_db)):
    return db.query(OutboundRecord).filter(OutboundRecord.status == status).order_by(desc(OutboundRecord.id)).all()

# 출하요청(재고장) -> 출고요청 생성
@app.post("/api/outbounds/create-from-stock")
def create_outbound_from_stock(req: OutboundCreate, db: Session = Depends(get_db)):
    lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).first()
    if not lot:
        raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")
    if req.box_qty > lot.current_box_qty or req.box_qty <= 0:
        raise HTTPException(status_code=400, detail="보유 박스 수량보다 많이 요청할 수 없습니다.")

    unit_weight = lot.current_weight_kg / lot.current_box_qty
    calc_weight = round(unit_weight * req.box_qty, 2)
    outbound_no = f"OUT-{datetime.now().strftime('%Y%m%d%H%M%S')}"

    outbound = OutboundRecord(
        outbound_no=outbound_no,
        lot_id=lot.id,
        customer=req.customer,
        item_name=lot.item_name,
        storage_type=lot.storage_type,
        trace_no=lot.trace_no,
        box_qty=req.box_qty,
        weight_kg=calc_weight,
        unit_price_kg=req.unit_price_kg,
        exp_date=lot.exp_date,
        warehouse=lot.warehouse,
        status="OUT_REQUEST"
    )
    db.add(outbound)
    db.commit()
    return {"message": "출고요청 등록 완료", "outbound_no": outbound_no}

# [기능추가] 출고 전표 수량/중량 수정
@app.put("/api/outbounds/{outbound_id}")
def update_outbound(outbound_id: int, req: UpdateQtyWeight, db: Session = Depends(get_db)):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")
    if outbound.status == "OUT_DONE":
        raise HTTPException(status_code=400, detail="이미 출고 완료된 전표는 [작업 취소] 후 수정해야 합니다.")

    outbound.box_qty = req.box_qty
    outbound.weight_kg = req.weight_kg
    db.commit()
    return {"message": "출고 수량 및 실중량이 수정되었습니다."}

# [기능추가] 출고 단계 취소/반려/원복
@app.post("/api/outbounds/{outbound_id}/revert")
def revert_outbound(outbound_id: int, db: Session = Depends(get_db)):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound:
        raise HTTPException(status_code=404, detail="전표를 찾을 수 없습니다.")

    if outbound.status == "OUT_REQUEST":
        # 1단계 취소 시 삭제
        db.delete(outbound)
        msg = "출고요청 전표가 취소(삭제)되었습니다."
    elif outbound.status == "OUT_CONFIRM":
        # 2단계 취소 시 1단계로 반려
        outbound.status = "OUT_REQUEST"
        msg = "출고확정이 취소되어 [출고요청리스트]로 반려되었습니다."
    elif outbound.status == "OUT_DONE":
        # 3단계 취소 시 실재고 가산 후 2단계로 복구
        lot = db.query(InventoryLot).filter(InventoryLot.id == outbound.lot_id).first()
        if lot:
            lot.current_box_qty += outbound.box_qty
            lot.current_weight_kg = round(lot.current_weight_kg + outbound.weight_kg, 2)
        outbound.status = "OUT_CONFIRM"
        msg = "출고완료가 취소되어 실재고가 복구(가산)되었으며 [출고확정리스트]로 복구되었습니다."
    else:
        msg = "처리할 수 없는 상태입니다."

    db.commit()
    return {"message": msg}

# 출고 단계 이동 (OUT_REQUEST -> OUT_CONFIRM -> OUT_DONE: 재고 차감)
@app.post("/api/outbounds/{outbound_id}/advance")
def advance_outbound(outbound_id: int, db: Session = Depends(get_db)):
    outbound = db.query(OutboundRecord).filter(OutboundRecord.id == outbound_id).first()
    if not outbound:
        raise HTTPException(status_code=404, detail="출고 전표를 찾을 수 없습니다.")

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

# 재고 감량/클레임 조정
@app.post("/api/inventory/adjust")
def adjust_stock(req: AdjustCreate, db: Session = Depends(get_db)):
    lot = db.query(InventoryLot).filter(InventoryLot.id == req.lot_id).first()
    if not lot:
        raise HTTPException(status_code=404, detail="재고 로트를 찾을 수 없습니다.")

    lot.current_box_qty = max(0, lot.current_box_qty - req.adj_box)
    lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - req.adj_weight, 2))

    log = StockAdjustment(
        lot_id=lot.id,
        adj_type=req.adj_type,
        adj_box=req.adj_box,
        adj_weight=req.adj_weight,
        reason=req.reason
    )
    db.add(log)
    db.commit()
    return {"message": f"재고 조정({req.adj_type})이 완료되었습니다."}

# -----------------------------------------------------------------------------
# 4. 프론트엔드 웹 UI
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
    .nav-item { padding: 12px 18px; color: #94a3b8; cursor: pointer; display: flex; align-items: center; gap: 10px; font-size: 0.9rem; }
    .nav-item:hover, .nav-item.active { background: #1e293b; color: #38bdf8; font-weight: 600; }
    main { flex-grow: 1; display: flex; flex-direction: column; height: 100vh; overflow-y: auto; }
    header { background: #fff; border-bottom: 1px solid var(--border); padding: 14px 28px; display: flex; justify-content: space-between; align-items: center; }
    .content { padding: 24px 28px; display: flex; flex-direction: column; gap: 16px; }
    .sub-tabs { display: flex; gap: 6px; background: #e2e8f0; padding: 4px; border-radius: 8px; width: fit-content; }
    .sub-tab { padding: 8px 16px; border-radius: 6px; font-size: 0.85rem; font-weight: 600; cursor: pointer; color: #475569; }
    .sub-tab.active { background: #fff; color: var(--primary); box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .table-card { background: #fff; border: 1px solid var(--border); border-radius: 8px; overflow: hidden; box-shadow: 0 1px 2px rgba(0,0,0,0.03); }
    table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
    th { background: #f8fafc; color: var(--muted); padding: 12px 16px; border-bottom: 1px solid var(--border); text-align: left; font-weight: 600; }
    td { padding: 12px 16px; border-bottom: 1px solid var(--border); vertical-align: middle; }
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
    .modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 100; align-items: center; justify-content: center; }
    .modal-box { background: #fff; border-radius: 10px; width: 480px; padding: 24px; }
    .form-group { margin-bottom: 12px; display: flex; flex-direction: column; gap: 4px; font-size: 0.82rem; font-weight: 600; }
    .form-control { padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.88rem; }
  </style>
</head>
<body>
  <aside>
    <div class="brand"><i class="bi bi-box-seam-fill"></i> MeatFlow ERP</div>
    <div class="nav-item active" onclick="switchMainTab('INBOUND', this)"><i class="bi bi-box-arrow-in-down"></i> 입고 관리</div>
    <div class="nav-item" onclick="switchMainTab('OUTBOUND', this)"><i class="bi bi-box-arrow-up"></i> 출고 관리 (재고장)</div>
  </aside>

  <main>
    <header>
      <h2 id="pageTitle" style="font-size:1.25rem;">입고 프로세스 관리</h2>
      <div id="headerActions" style="display:flex; gap:8px;">
        <input type="file" id="excelFileInput" style="display:none;" onchange="uploadExcelFile(event)" accept=".xlsx, .xls" />
        <button class="btn btn-outline" onclick="document.getElementById('excelFileInput').click()"><i class="bi bi-file-earmark-excel"></i> 엑셀 일괄입고</button>
        <button class="btn btn-primary" onclick="openInboundModal()"><i class="bi bi-plus-lg"></i> 신규 입고요청</button>
      </div>
    </header>

    <div class="content">
      <div class="sub-tabs" id="subTabs"></div>
      <div class="table-card">
        <table>
          <thead id="tableHead"></thead>
          <tbody id="tableBody"></tbody>
        </table>
      </div>
    </div>
  </main>

  <!-- 입고 등록 모달 -->
  <div class="modal-overlay" id="inboundModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px;">신규 축산물 입고요청</h3>
      <div class="form-group"><label>매입처 상호</label><input type="text" id="in_vendor" class="form-control" value="(주)글로벌미트" /></div>
      <div class="form-group"><label>품목 및 부위명</label><input type="text" id="in_item" class="form-control" value="돼지 삼겹살(올리멜)" /></div>
      <div class="form-group"><label>보관 형태</label><select id="in_storage" class="form-control"><option value="냉장">냉장</option><option value="냉동">냉동</option></select></div>
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
        <button class="btn btn-outline" onclick="closeModal()">취소</button>
        <button class="btn btn-primary" onclick="submitInbound()">저장</button>
      </div>
    </div>
  </div>

  <!-- 수량/중량 수정 모달 -->
  <div class="modal-overlay" id="editQtyModal">
    <div class="modal-box">
      <h3 style="margin-bottom:14px;" id="editModalTitle">수량 및 중량 수정</h3>
      <input type="hidden" id="edit_target_id" />
      <input type="hidden" id="edit_target_type" />
      <div class="form-group">
        <label>수정할 박스 수량(Box)</label>
        <input type="number" id="edit_box_qty" class="form-control" />
      </div>
      <div class="form-group">
        <label>수정할 총 실중량(kg)</label>
        <input type="number" step="0.01" id="edit_weight_kg" class="form-control" />
      </div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeEditModal()">취소</button>
        <button class="btn btn-primary" onclick="submitEditQtyWeight()">수정 완료</button>
      </div>
    </div>
  </div>

  <!-- 재고 감량 모달 -->
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
        <div class="form-group" style="flex:1;"><label>차감 박스 수</label><input type="number" id="adj_box" class="form-control" value="1" /></div>
        <div class="form-group" style="flex:1;"><label>차감 실중량(kg)</label><input type="number" step="0.1" id="adj_weight" class="form-control" value="20.4" /></div>
      </div>
      <div class="form-group"><label>상세 사유 기록</label><input type="text" id="adj_reason" class="form-control" placeholder="예: 표면 갈변 발생 폐기" /></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:16px;">
        <button class="btn btn-outline" onclick="closeAdjustModal()">취소</button>
        <button class="btn btn-danger" onclick="submitAdjust()">조정 적용</button>
      </div>
    </div>
  </div>

  <script>
    let currentMain = 'INBOUND', currentSub = 'IN_REQUEST';

    function switchMainTab(tab, el) {
      currentMain = tab;
      document.querySelectorAll('.nav-item').forEach(e => e.classList.remove('active'));
      el.classList.add('active');
      currentSub = tab === 'INBOUND' ? 'IN_REQUEST' : 'OUT_STOCK';
      document.getElementById('pageTitle').innerText = tab === 'INBOUND' ? '입고 프로세스 관리' : '출고 프로세스 관리 (현 재고장 연동)';
      document.getElementById('headerActions').style.display = tab === 'INBOUND' ? 'flex' : 'none';
      renderSubTabs();
      loadData();
    }

    function renderSubTabs() {
      const container = document.getElementById('subTabs'); container.innerHTML = '';
      const steps = currentMain === 'INBOUND' 
        ? [{k:'IN_REQUEST',n:'1. 입고요청리스트'},{k:'IN_CONFIRM',n:'2. 입고확정리스트'},{k:'IN_DONE',n:'3. 입고완료리스트'}]
        : [{k:'OUT_STOCK',n:'1. 출하요청리스트 (현 재고장)'},{k:'OUT_REQUEST',n:'2. 출고요청리스트'},{k:'OUT_CONFIRM',n:'3. 출고확정리스트'},{k:'OUT_DONE',n:'4. 출고완료리스트'}];
      steps.forEach(s => {
        container.innerHTML += `<div class="sub-tab ${currentSub===s.k?'active':''}" onclick="setSubTab('${s.k}')">${s.n}</div>`;
      });
    }

    function setSubTab(k) { currentSub = k; renderSubTabs(); loadData(); }

    async function loadData() {
      const head = document.getElementById('tableHead'), body = document.getElementById('tableBody');
      body.innerHTML = '<tr><td colspan="9" style="text-align:center; padding:15px;">데이터 조회 중...</td></tr>';
      
      if (currentMain === 'INBOUND') {
        head.innerHTML = '<tr><th>전표번호</th><th>매입처</th><th>품목명</th><th>보관</th><th>이력번호</th><th>박스</th><th>중량(kg)</th><th>유통기한</th><th style="text-align:center;">작업 관리</th></tr>';
        const res = await fetch(`/api/inbounds?status=${currentSub}`), list = await res.json();
        body.innerHTML = list.length ? list.map(i => {
          let actionBtns = '';
          if (currentSub === 'IN_REQUEST') {
            actionBtns = `
              <button class="btn btn-primary" onclick="advIn(${i.id})">확정 <i class="bi bi-arrow-right"></i></button>
              <button class="btn btn-outline" onclick="openEditModal('INBOUND', ${i.id}, ${i.box_qty}, ${i.weight_kg})"><i class="bi bi-pencil"></i></button>
              <button class="btn btn-outline" style="color:#dc2626;" onclick="revertIn(${i.id}, '요청 취소(삭제)')"><i class="bi bi-trash"></i></button>
            `;
          } else if (currentSub === 'IN_CONFIRM') {
            actionBtns = `
              <button class="btn btn-success" onclick="advIn(${i.id})"><i class="bi bi-check-lg"></i> 입고완료</button>
              <button class="btn btn-outline" onclick="openEditModal('INBOUND', ${i.id}, ${i.box_qty}, ${i.weight_kg})"><i class="bi bi-pencil"></i></button>
              <button class="btn btn-outline" style="color:#d97706;" onclick="revertIn(${i.id}, '이전 단계 반려')"><i class="bi bi-arrow-counterclockwise"></i> 반려</button>
            `;
          } else {
            actionBtns = `
              <span class="badge" style="background:#dcfce7; color:#15803d; margin-right:6px;">완료됨</span>
              <button class="btn btn-outline" style="color:#dc2626;" onclick="revertIn(${i.id}, '입고 취소 및 재고 원복')"><i class="bi bi-arrow-counterclockwise"></i> 작업취소</button>
            `;
          }
          return `
            <tr>
              <td><strong>${i.inbound_no}</strong></td>
              <td>${i.vendor}</td>
              <td>${i.item_name}</td>
              <td><span class="badge ${i.storage_type === '냉장' ? 'badge-chilled' : 'badge-frozen'}">${i.storage_type}</span></td>
              <td><code>${i.trace_no}</code></td>
              <td><strong>${i.box_qty}</strong> Box</td>
              <td><strong>${i.weight_kg.toLocaleString()}</strong> kg</td>
              <td>${i.exp_date}</td>
              <td style="text-align:center; white-space:nowrap;">${actionBtns}</td>
            </tr>
          `;
        }).join('') : '<tr><td colspan="9" style="text-align:center; padding:20px; color:#94a3b8;">대기 중인 전표가 없습니다.</td></tr>';
      } else {
        if (currentSub === 'OUT_STOCK') {
          head.innerHTML = '<tr><th>로트ID</th><th>품목명</th><th>보관</th><th>이력번호</th><th>보유 박스</th><th>보유 중량(kg)</th><th>유통기한</th><th>창고</th><th style="text-align:center;">관리 작업</th></tr>';
          const res = await fetch('/api/inventory'), list = await res.json();
          body.innerHTML = list.length ? list.map(l => `
            <tr>
              <td>#${l.id}</td>
              <td><strong>${l.item_name}</strong></td>
              <td><span class="badge ${l.storage_type === '냉장' ? 'badge-chilled' : 'badge-frozen'}">${l.storage_type}</span></td>
              <td><code>${l.trace_no}</code></td>
              <td><strong>${l.current_box_qty}</strong> Box</td>
              <td><strong>${l.current_weight_kg.toLocaleString()}</strong> kg</td>
              <td>${l.exp_date}</td>
              <td>${l.warehouse}</td>
              <td style="text-align:center; white-space:nowrap;">
                <button class="btn btn-warning" onclick="reqOut(${l.id},${l.current_box_qty})"><i class="bi bi-cart-plus"></i> 출하요청</button>
                <button class="btn btn-outline" style="color:#dc2626;" onclick="openAdjustModal(${l.id})"><i class="bi bi-scissors"></i> 감량/폐기</button>
              </td>
            </tr>
          `).join('') : '<tr><td colspan="9" style="text-align:center; padding:20px; color:#94a3b8;">보유 재고가 없습니다.</td></tr>';
        } else {
          head.innerHTML = '<tr><th>출고번호</th><th>매출처</th><th>품목명</th><th>이력번호</th><th>박스</th><th>중량(kg)</th><th>단가</th><th>유통기한</th><th style="text-align:center;">작업 관리</th></tr>';
          const res = await fetch(`/api/outbounds?status=${currentSub}`), list = await res.json();
          body.innerHTML = list.length ? list.map(o => {
            let actionBtns = '';
            if (currentSub === 'OUT_REQUEST') {
              actionBtns = `
                <button class="btn btn-primary" onclick="advOut(${o.id})">확정 <i class="bi bi-arrow-right"></i></button>
                <button class="btn btn-outline" onclick="openEditModal('OUTBOUND', ${o.id}, ${o.box_qty}, ${o.weight_kg})"><i class="bi bi-pencil"></i></button>
                <button class="btn btn-outline" style="color:#dc2626;" onclick="revertOut(${o.id}, '요청 취소(삭제)')"><i class="bi bi-trash"></i></button>
              `;
            } else if (currentSub === 'OUT_CONFIRM') {
              actionBtns = `
                <button class="btn btn-success" onclick="advOut(${o.id})"><i class="bi bi-truck"></i> 출고완료</button>
                <button class="btn btn-outline" onclick="openEditModal('OUTBOUND', ${o.id}, ${o.box_qty}, ${o.weight_kg})"><i class="bi bi-pencil"></i></button>
                <button class="btn btn-outline" style="color:#d97706;" onclick="revertOut(${o.id}, '이전 단계 반려')"><i class="bi bi-arrow-counterclockwise"></i> 반려</button>
              `;
            } else {
              actionBtns = `
                <span class="badge" style="background:#dbeafe; color:#1d4ed8; margin-right:6px;">출고완료</span>
                <button class="btn btn-outline" style="color:#dc2626;" onclick="revertOut(${o.id}, '출고 취소 및 재고 원복')"><i class="bi bi-arrow-counterclockwise"></i> 작업취소</button>
              `;
            }
            return `
              <tr>
                <td><strong>${o.outbound_no}</strong></td>
                <td>${o.customer}</td>
                <td>${o.item_name}</td>
                <td><code>${o.trace_no}</code></td>
                <td><strong>${o.box_qty}</strong> Box</td>
                <td><strong>${o.weight_kg.toLocaleString()}</strong> kg</td>
                <td>${o.unit_price_kg.toLocaleString()}원</td>
                <td>${o.exp_date}</td>
                <td style="text-align:center; white-space:nowrap;">${actionBtns}</td>
              </tr>
            `;
          }).join('') : '<tr><td colspan="9" style="text-align:center; padding:20px; color:#94a3b8;">대기 중인 출고 건이 없습니다.</td></tr>';
        }
      }
    }

    async function advIn(id) { const r = await fetch(`/api/inbounds/${id}/advance`,{method:'POST'}); const d = await r.json(); alert(d.message); loadData(); }
    async function advOut(id) { const r = await fetch(`/api/outbounds/${id}/advance`,{method:'POST'}); const d = await r.json(); alert(d.message); loadData(); }
    
    async function revertIn(id, actionName) {
      if(!confirm(`정말 [${actionName}] 작업을 진행하시겠습니까?`)) return;
      const r = await fetch(`/api/inbounds/${id}/revert`,{method:'POST'});
      const d = await r.json(); alert(d.message); loadData();
    }

    async function revertOut(id, actionName) {
      if(!confirm(`정말 [${actionName}] 작업을 진행하시겠습니까?`)) return;
      const r = await fetch(`/api/outbounds/${id}/revert`,{method:'POST'});
      const d = await r.json(); alert(d.message); loadData();
    }

    function openEditModal(type, id, box, weight) {
      document.getElementById('edit_target_type').value = type;
      document.getElementById('edit_target_id').value = id;
      document.getElementById('edit_box_qty').value = box;
      document.getElementById('edit_weight_kg').value = weight;
      document.getElementById('editModalTitle').innerText = `${type === 'INBOUND' ? '입고' : '출고'} 수량/중량 수정 (ID: ${id})`;
      document.getElementById('editQtyModal').style.display = 'flex';
    }
    function closeEditModal() { document.getElementById('editQtyModal').style.display = 'none'; }

    async function submitEditQtyWeight() {
      const type = document.getElementById('edit_target_type').value;
      const id = document.getElementById('edit_target_id').value;
      const box = Number(document.getElementById('edit_box_qty').value);
      const weight = Number(document.getElementById('edit_weight_kg').value);

      const url = type === 'INBOUND' ? `/api/inbounds/${id}` : `/api/outbounds/${id}`;
      const r = await fetch(url, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ box_qty: box, weight_kg: weight })
      });
      const d = await r.json();
      if(r.ok) {
        alert(d.message);
        closeEditModal();
        loadData();
      } else {
        alert("수정 실패: " + d.detail);
      }
    }

    async function reqOut(lotId, maxBox) {
      const cust = prompt("출하할 매출처 상호명을 입력하세요:", "(주)신규거래처"); if(!cust) return;
      const box = prompt(`출하 요청 박스 수를 입력하세요 (보유: ${maxBox} Box):`, "10"); if(!box || isNaN(box) || Number(box)<=0 || Number(box)>maxBox) { alert("수량이 올바르지 않습니다."); return; }
      const price = prompt("kg당 판매단가(원)를 입력하세요:", "11500"); if(!price) return;
      const r = await fetch('/api/outbounds/create-from-stock', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({lot_id:lotId, customer:cust, box_qty:Number(box), unit_price_kg:Number(price)})});
      const d = await r.json(); alert(d.message); setSubTab('OUT_REQUEST');
    }

    function openInboundModal() { document.getElementById('inboundModal').style.display = 'flex'; }
    function closeModal() { document.getElementById('inboundModal').style.display = 'none'; }
    
    function openAdjustModal(lotId) { document.getElementById('adj_lot_id').value = lotId; document.getElementById('adjustModal').style.display = 'flex'; }
    function closeAdjustModal() { document.getElementById('adjustModal').style.display = 'none'; }

    async function submitInbound() {
      const p = { vendor: document.getElementById('in_vendor').value, item_name: document.getElementById('in_item').value, storage_type: document.getElementById('in_storage').value, trace_no: document.getElementById('in_trace').value, box_qty: Number(document.getElementById('in_box').value), weight_kg: Number(document.getElementById('in_weight').value), cost_per_kg: Number(document.getElementById('in_cost').value), warehouse: "광주냉장창고", exp_date: document.getElementById('in_exp').value };
      const r = await fetch('/api/inbounds', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p)});
      const d = await r.json(); alert(d.message); closeModal(); setSubTab('IN_REQUEST');
    }

    async function submitAdjust() {
      const p = { lot_id: Number(document.getElementById('adj_lot_id').value), adj_type: document.getElementById('adj_type').value, adj_box: Number(document.getElementById('adj_box').value), adj_weight: Number(document.getElementById('adj_weight').value), reason: document.getElementById('adj_reason').value };
      const r = await fetch('/api/inventory/adjust', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p)});
      const d = await r.json(); alert(d.message); closeAdjustModal(); loadData();
    }

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
