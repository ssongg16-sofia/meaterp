pip install fastapi uvicorn sqlalchemy pydantic jinja2

python app.py

import os
from datetime import date, datetime
from decimal import Decimal
from typing import List, Optional

from fastapi import FastAPI, HTTPException, Depends
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from sqlalchemy import (
    Column, Integer, Float, String, Date, DateTime, Enum,
    create_engine, desc
)
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
import uvicorn

# -----------------------------------------------------------------------------
# 1. 데이터베이스 설정 (SQLite 자동 생성 및 연동)
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
class InboundRecord(Base):
    __tablename__ = "inbounds"
    id = Column(Integer, primary_key=True, index=True)
    inbound_no = Column(String(50), unique=True, index=True)
    vendor = Column(String(100), nullable=False)
    item_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)  # '냉장', '냉동'
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
# 3. FastAPI 라우터 및 상태 전이 비즈니스 로직
# -----------------------------------------------------------------------------
app = FastAPI(title="MeatFlow ERP API")

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

class OutboundCreate(BaseModel):
    lot_id: int
    customer: str
    box_qty: int
    unit_price_kg: float

# 입고 데이터 조회
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

# 입고 상태 전이 (IN_REQUEST -> IN_CONFIRM -> IN_DONE: 재고 합산)
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
        # 실재고 로트 반영
        lot = db.query(InventoryLot).filter(
            InventoryLot.trace_no == inbound.trace_no,
            InventoryLot.warehouse == inbound.warehouse,
            InventoryLot.exp_date == inbound.exp_date
        ).first()

        if lot:
            lot.current_box_qty += inbound.box_qty
            lot.current_weight_kg += inbound.weight_kg
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

# 출고 데이터 조회
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

# 출고 상태 전이 (OUT_REQUEST -> OUT_CONFIRM -> OUT_DONE: 재고 차감)
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
        # 실재고 차감 처리
        lot = db.query(InventoryLot).filter(InventoryLot.id == outbound.lot_id).first()
        if lot:
            lot.current_box_qty = max(0, lot.current_box_qty - outbound.box_qty)
            lot.current_weight_kg = max(0.0, round(lot.current_weight_kg - outbound.weight_kg, 2))
        msg = "출고완료 처리 및 실재고 차감이 완료되었습니다."
    else:
        msg = "이미 완료된 전표입니다."

    db.commit()
    return {"message": msg}

# -----------------------------------------------------------------------------
# 4. 실시간 통합 대시보드 HTML 프론트엔드 서빙
# -----------------------------------------------------------------------------
HTML_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>MeatFlow ERP - 실운영 시스템</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
  <style>
    :root {
      --primary: #2563eb;
      --sidebar-bg: #0f172a;
      --bg-body: #f8fafc;
      --card-bg: #ffffff;
      --border: #e2e8f0;
      --text: #1e293b;
      --muted: #64748b;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Pretendard", sans-serif; }
    body { background: var(--bg-body); color: var(--text); display: flex; height: 100vh; overflow: hidden; }
    aside { width: 240px; background: var(--sidebar-bg); color: #94a3b8; display: flex; flex-direction: column; flex-shrink: 0; }
    .brand { padding: 20px; font-size: 1.15rem; font-weight: 700; color: #fff; border-bottom: 1px solid #1e293b; display: flex; align-items: center; gap: 8px; }
    .brand i { color: #38bdf8; }
    .nav-item { padding: 12px 18px; color: #94a3b8; font-size: 0.9rem; font-weight: 500; cursor: pointer; display: flex; align-items: center; gap: 10px; }
    .nav-item:hover, .nav-item.active { background: #1e293b; color: #38bdf8; font-weight: 600; }
    main { flex-grow: 1; display: flex; flex-direction: column; height: 100vh; overflow-y: auto; }
    header { background: #fff; border-bottom: 1px solid var(--border); padding: 14px 28px; display: flex; justify-content: space-between; align-items: center; }
    .page-title { font-size: 1.25rem; font-weight: 700; }
    .content { padding: 24px 28px; display: flex; flex-direction: column; gap: 18px; }
    .sub-tabs { display: flex; gap: 6px; background: #e2e8f0; padding: 4px; border-radius: 8px; width: fit-content; }
    .sub-tab { padding: 8px 16px; border-radius: 6px; font-size: 0.85rem; font-weight: 600; color: #475569; cursor: pointer; }
    .sub-tab.active { background: #fff; color: var(--primary); box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .table-card { background: #fff; border: 1px solid var(--border); border-radius: 8px; overflow: hidden; }
    .table-header { padding: 14px 18px; border-bottom: 1px solid var(--border); font-weight: 700; display: flex; justify-content: space-between; align-items: center; }
    table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
    th { background: #f8fafc; color: var(--muted); font-weight: 600; padding: 12px 16px; border-bottom: 1px solid var(--border); text-align: left; }
    td { padding: 12px 16px; border-bottom: 1px solid var(--border); vertical-align: middle; }
    tr:hover td { background: #f8fafc; }
    .btn { padding: 6px 12px; border-radius: 6px; font-size: 0.82rem; font-weight: 600; border: none; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; }
    .btn-primary { background: var(--primary); color: white; }
    .btn-success { background: #16a34a; color: white; }
    .btn-warning { background: #d97706; color: white; }
    .badge { padding: 3px 8px; border-radius: 4px; font-size: 0.75rem; font-weight: 600; }
    .badge-chilled { background: #e0f2fe; color: #0284c7; }
    .badge-frozen { background: #e0e7ff; color: #4f46e5; }
    /* 모달 */
    .modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 100; align-items: center; justify-content: center; }
    .modal-box { background: white; border-radius: 10px; width: 500px; padding: 24px; }
    .form-group { margin-bottom: 12px; display: flex; flex-direction: column; gap: 4px; font-size: 0.85rem; font-weight: 600; }
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
      <div class="page-title" id="pageTitle">입고 프로세스 관리</div>
      <div id="headerActions">
        <button class="btn btn-primary" onclick="openInboundModal()"><i class="bi bi-plus-lg"></i> 신규 입고요청 등록</button>
      </div>
    </header>

    <div class="content">
      <div class="sub-tabs" id="subTabs"></div>
      <div class="table-card">
        <div class="table-header" id="tableHeader">목록</div>
        <div style="overflow-x: auto;">
          <table id="mainTable">
            <thead id="tableHead"></thead>
            <tbody id="tableBody"></tbody>
          </table>
        </div>
      </div>
    </div>
  </main>

  <!-- 입고 등록 모달 -->
  <div class="modal-overlay" id="inboundModal">
    <div class="modal-box">
      <h3 style="margin-bottom: 16px;">신규 축산물 입고요청 등록</h3>
      <div class="form-group">
        <label>매입처 상호</label>
        <input type="text" id="in_vendor" class="form-control" value="(주)글로벌미트" />
      </div>
      <div class="form-group">
        <label>품목 및 부위명</label>
        <input type="text" id="in_item" class="form-control" value="돼지 삼겹살(올리멜)" />
      </div>
      <div class="form-group">
        <label>보관 형태</label>
        <select id="in_storage" class="form-control">
          <option value="냉장">냉장 (Chilled)</option>
          <option value="냉동">냉동 (Frozen)</option>
        </select>
      </div>
      <div class="form-group">
        <label>축산물 이력번호 (12자리)</label>
        <input type="text" id="in_trace" class="form-control" value="802410290114" />
      </div>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">
        <div class="form-group">
          <label>입고 박스 수(Box)</label>
          <input type="number" id="in_box" class="form-control" value="50" />
        </div>
        <div class="form-group">
          <label>총 실중량(kg)</label>
          <input type="number" step="0.1" id="in_weight" class="form-control" value="1020.5" />
        </div>
      </div>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">
        <div class="form-group">
          <label>kg당 매입단가(원)</label>
          <input type="number" id="in_cost" class="form-control" value="8400" />
        </div>
        <div class="form-group">
          <label>유통기한</label>
          <input type="date" id="in_exp" class="form-control" value="2026-09-15" />
        </div>
      </div>
      <div style="display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px;">
        <button class="btn" style="background:#e2e8f0;" onclick="closeModal()">취소</button>
        <button class="btn btn-primary" onclick="submitInbound()">입고요청 저장</button>
      </div>
    </div>
  </div>

  <script>
    let currentMain = 'INBOUND';
    let currentSub = 'IN_REQUEST';

    function switchMainTab(tab, el) {
      currentMain = tab;
      document.querySelectorAll('.nav-item').forEach(e => e.classList.remove('active'));
      el.classList.add('active');

      if (tab === 'INBOUND') {
        currentSub = 'IN_REQUEST';
        document.getElementById('pageTitle').innerText = '입고 프로세스 관리';
        document.getElementById('headerActions').innerHTML = '<button class="btn btn-primary" onclick="openInboundModal()"><i class="bi bi-plus-lg"></i> 신규 입고요청 등록</button>';
      } else {
        currentSub = 'OUT_STOCK';
        document.getElementById('pageTitle').innerText = '출고 프로세스 관리 (현 재고장 연동)';
        document.getElementById('headerActions').innerHTML = '';
      }
      renderSubTabs();
      loadData();
    }

    function renderSubTabs() {
      const container = document.getElementById('subTabs');
      container.innerHTML = '';
      if (currentMain === 'INBOUND') {
        const steps = [
          { key: 'IN_REQUEST', name: '1. 입고요청리스트' },
          { key: 'IN_CONFIRM', name: '2. 입고확정리스트' },
          { key: 'IN_DONE', name: '3. 입고리스트(완료)' }
        ];
        steps.forEach(s => {
          const active = currentSub === s.key ? 'active' : '';
          container.innerHTML += `<div class="sub-tab ${active}" onclick="setSubTab('${s.key}')">${s.name}</div>`;
        });
      } else {
        const steps = [
          { key: 'OUT_STOCK', name: '1. 출하요청리스트 (현 재고장)' },
          { key: 'OUT_REQUEST', name: '2. 출고요청리스트' },
          { key: 'OUT_CONFIRM', name: '3. 출고확정리스트' },
          { key: 'OUT_DONE', name: '4. 출고리스트(완료)' }
        ];
        steps.forEach(s => {
          const active = currentSub === s.key ? 'active' : '';
          container.innerHTML += `<div class="sub-tab ${active}" onclick="setSubTab('${s.key}')">${s.name}</div>`;
        });
      }
    }

    function setSubTab(sub) {
      currentSub = sub;
      renderSubTabs();
      loadData();
    }

    async function loadData() {
      const head = document.getElementById('tableHead');
      const body = document.getElementById('tableBody');
      body.innerHTML = '<tr><td colspan="9" style="text-align:center; padding:20px;">데이터 로딩 중...</td></tr>';

      if (currentMain === 'INBOUND') {
        head.innerHTML = `
          <tr>
            <th>전표번호</th>
            <th>매입처</th>
            <th>품목명</th>
            <th>보관</th>
            <th>이력번호</th>
            <th style="text-align:right;">박스 수</th>
            <th style="text-align:right;">실중량(kg)</th>
            <th>유통기한</th>
            <th style="text-align:center;">상태 처리</th>
          </tr>
        `;
        const res = await fetch(`/api/inbounds?status=${currentSub}`);
        const list = await res.json();
        body.innerHTML = '';
        if (list.length === 0) {
          body.innerHTML = '<tr><td colspan="9" style="text-align:center; padding:30px; color:#94a3b8;">대기 중인 전표가 없습니다.</td></tr>';
          return;
        }
        list.forEach(item => {
          let btn = '';
          if (currentSub === 'IN_REQUEST') {
            btn = `<button class="btn btn-primary" onclick="advanceInbound(${item.id})">입고확정 <i class="bi bi-arrow-right"></i></button>`;
          } else if (currentSub === 'IN_CONFIRM') {
            btn = `<button class="btn btn-success" onclick="advanceInbound(${item.id})"><i class="bi bi-check-lg"></i> 입고완료(재고반영)</button>`;
          } else {
            btn = `<span class="badge" style="background:#dcfce7; color:#15803d;">입고 완료됨</span>`;
          }
          body.innerHTML += `
            <tr>
              <td><strong>${item.inbound_no}</strong></td>
              <td>${item.vendor}</td>
              <td>${item.item_name}</td>
              <td><span class="badge ${item.storage_type === '냉장' ? 'badge-chilled' : 'badge-frozen'}">${item.storage_type}</span></td>
              <td><code>${item.trace_no}</code></td>
              <td style="text-align:right; font-weight:600;">${item.box_qty} Box</td>
              <td style="text-align:right; font-weight:600;">${item.weight_kg.toLocaleString()} kg</td>
              <td>${item.exp_date}</td>
              <td style="text-align:center;">${btn}</td>
            </tr>
          `;
        });
      } else {
        if (currentSub === 'OUT_STOCK') {
          head.innerHTML = `
            <tr>
              <th>로트 ID</th>
              <th>품목명</th>
              <th>보관</th>
              <th>이력번호</th>
              <th style="text-align:right;">보유 박스</th>
              <th style="text-align:right;">보유 중량(kg)</th>
              <th>보관창고</th>
              <th>유통기한</th>
              <th style="text-align:center;">출하 등록</th>
            </tr>
          `;
          const res = await fetch('/api/inventory');
          const list = await res.json();
          body.innerHTML = '';
          if (list.length === 0) {
            body.innerHTML = '<tr><td colspan="9" style="text-align:center; padding:30px; color:#94a3b8;">보유 재고가 없습니다.</td></tr>';
            return;
          }
          list.forEach(lot => {
            body.innerHTML += `
              <tr>
                <td><strong>#${lot.id}</strong></td>
                <td>${lot.item_name}</td>
                <td><span class="badge ${lot.storage_type === '냉장' ? 'badge-chilled' : 'badge-frozen'}">${lot.storage_type}</span></td>
                <td><code>${lot.trace_no}</code></td>
                <td style="text-align:right; font-weight:600;">${lot.current_box_qty} Box</td>
                <td style="text-align:right; font-weight:600;">${lot.current_weight_kg.toLocaleString()} kg</td>
                <td>${lot.warehouse}</td>
                <td>${lot.exp_date}</td>
                <td style="text-align:center;">
                  <button class="btn btn-warning" onclick="requestOutbound(${lot.id}, ${lot.current_box_qty})"><i class="bi bi-cart-plus"></i> 출하요청</button>
                </td>
              </tr>
            `;
          });
        } else {
          head.innerHTML = `
            <tr>
              <th>출고번호</th>
              <th>매출처</th>
              <th>품목명</th>
              <th>이력번호</th>
              <th style="text-align:right;">출고 박스</th>
              <th style="text-align:right;">출고 중량(kg)</th>
              <th style="text-align:right;">판매 단가(kg)</th>
              <th>유통기한</th>
              <th style="text-align:center;">상태 처리</th>
            </tr>
          `;
          const res = await fetch(`/api/outbounds?status=${currentSub}`);
          const list = await res.json();
          body.innerHTML = '';
          if (list.length === 0) {
            body.innerHTML = '<tr><td colspan="9" style="text-align:center; padding:30px; color:#94a3b8;">대기 중인 출고 건이 없습니다.</td></tr>';
            return;
          }
          list.forEach(item => {
            let btn = '';
            if (currentSub === 'OUT_REQUEST') {
              btn = `<button class="btn btn-primary" onclick="advanceOutbound(${item.id})">출고확정 <i class="bi bi-arrow-right"></i></button>`;
            } else if (currentSub === 'OUT_CONFIRM') {
              btn = `<button class="btn btn-success" onclick="advanceOutbound(${item.id})"><i class="bi bi-truck"></i> 출고완료(재고차감)</button>`;
            } else {
              btn = `<span class="badge" style="background:#dbeafe; color:#1d4ed8;">출고 완료됨</span>`;
            }
            body.innerHTML += `
              <tr>
                <td><strong>${item.outbound_no}</strong></td>
                <td>${item.customer}</td>
                <td>${item.item_name}</td>
                <td><code>${item.trace_no}</code></td>
                <td style="text-align:right; font-weight:600;">${item.box_qty} Box</td>
                <td style="text-align:right; font-weight:600;">${item.weight_kg.toLocaleString()} kg</td>
                <td style="text-align:right;">${item.unit_price_kg.toLocaleString()}원</td>
                <td>${item.exp_date}</td>
                <td style="text-align:center;">${btn}</td>
              </tr>
            `;
          });
        }
      }
    }

    async function advanceInbound(id) {
      const res = await fetch(`/api/inbounds/${id}/advance`, { method: 'POST' });
      const data = await res.json();
      alert(data.message);
      loadData();
    }

    async function advanceOutbound(id) {
      const res = await fetch(`/api/outbounds/${id}/advance`, { method: 'POST' });
      const data = await res.json();
      alert(data.message);
      loadData();
    }

    async function requestOutbound(lotId, maxBox) {
      const customer = prompt("출하할 거래처 상호명을 입력하세요:", "(주)신규식당");
      if (!customer) return;
      const box = prompt(`출하 요청 박스 수를 입력하세요 (보유: ${maxBox} Box):`, "10");
      if (!box || isNaN(box) || Number(box) <= 0 || Number(box) > maxBox) {
        alert("올바른 박스 수량을 입력하세요.");
        return;
      }
      const price = prompt("kg당 판매단가(원)를 입력하세요:", "11500");
      if (!price) return;

      const res = await fetch('/api/outbounds/create-from-stock', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ lot_id: lotId, customer: customer, box_qty: Number(box), unit_price_kg: Number(price) })
      });
      const data = await res.json();
      alert(data.message);
      setSubTab('OUT_REQUEST');
    }

    function openInboundModal() { document.getElementById('inboundModal').style.display = 'flex'; }
    function closeModal() { document.getElementById('inboundModal').style.display = 'none'; }

    async function submitInbound() {
      const payload = {
        vendor: document.getElementById('in_vendor').value,
        item_name: document.getElementById('in_item').value,
        storage_type: document.getElementById('in_storage').value,
        trace_no: document.getElementById('in_trace').value,
        box_qty: Number(document.getElementById('in_box').value),
        weight_kg: Number(document.getElementById('in_weight').value),
        cost_per_kg: Number(document.getElementById('in_cost').value),
        warehouse: "광주냉장창고",
        exp_date: document.getElementById('in_exp').value
      };
      const res = await fetch('/api/inbounds', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      alert(data.message);
      closeModal();
      setSubTab('IN_REQUEST');
    }

    window.addEventListener('DOMContentLoaded', () => {
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
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)