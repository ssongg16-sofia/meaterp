import os
from datetime import datetime
from typing import List

from fastapi import FastAPI, HTTPException, Depends
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from sqlalchemy import (
    Column, Integer, Float, String, DateTime, create_engine, desc
)
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
import uvicorn

# DB 설정 (SQLite 파일 기반)
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

# ORM 모델 정의
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
    status = Column(String(20), default="IN_REQUEST")
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
    status = Column(String(20), default="OUT_REQUEST")
    created_at = Column(DateTime, default=datetime.now)

Base.metadata.create_all(bind=engine)

# 초기 샘플 데이터
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

@app.get("/api/inbounds")
def get_inbounds(status: str, db: Session = Depends(get_db)):
    return db.query(InboundRecord).filter(InboundRecord.status == status).order_by(desc(InboundRecord.id)).all()

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

@app.get("/api/inventory")
def get_inventory(db: Session = Depends(get_db)):
    return db.query(InventoryLot).filter(InventoryLot.current_box_qty > 0).order_by(InventoryLot.exp_date).all()

@app.get("/api/outbounds")
def get_outbounds(status: str, db: Session = Depends(get_db)):
    return db.query(OutboundRecord).filter(OutboundRecord.status == status).order_by(desc(OutboundRecord.id)).all()

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

HTML_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>MeatFlow ERP</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
  <style>
    :root { --primary: #2563eb; --sidebar: #0f172a; --bg: #f8fafc; --border: #e2e8f0; }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Pretendard", sans-serif; }
    body { background: var(--bg); display: flex; height: 100vh; overflow: hidden; }
    aside { width: 240px; background: var(--sidebar); color: #94a3b8; display: flex; flex-direction: column; }
    .brand { padding: 20px; font-size: 1.15rem; font-weight: 700; color: #fff; border-bottom: 1px solid #1e293b; }
    .nav-item { padding: 12px 18px; color: #94a3b8; cursor: pointer; display: flex; align-items: center; gap: 8px; }
    .nav-item:hover, .nav-item.active { background: #1e293b; color: #38bdf8; font-weight: 600; }
    main { flex-grow: 1; display: flex; flex-direction: column; height: 100vh; overflow-y: auto; }
    header { background: #fff; border-bottom: 1px solid var(--border); padding: 14px 28px; display: flex; justify-content: space-between; align-items: center; }
    .content { padding: 24px 28px; display: flex; flex-direction: column; gap: 16px; }
    .sub-tabs { display: flex; gap: 6px; background: #e2e8f0; padding: 4px; border-radius: 8px; width: fit-content; }
    .sub-tab { padding: 8px 16px; border-radius: 6px; font-size: 0.85rem; font-weight: 600; cursor: pointer; }
    .sub-tab.active { background: #fff; color: var(--primary); }
    .table-card { background: #fff; border: 1px solid var(--border); border-radius: 8px; overflow: hidden; }
    table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
    th { background: #f8fafc; color: #64748b; padding: 12px 16px; border-bottom: 1px solid var(--border); text-align: left; }
    td { padding: 12px 16px; border-bottom: 1px solid var(--border); }
    .btn { padding: 6px 12px; border-radius: 6px; font-size: 0.82rem; font-weight: 600; border: none; cursor: pointer; }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-success { background: #16a34a; color: #fff; }
    .btn-warning { background: #d97706; color: #fff; }
    .modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 100; align-items: center; justify-content: center; }
    .modal-box { background: #fff; border-radius: 10px; width: 480px; padding: 24px; }
    .form-group { margin-bottom: 10px; display: flex; flex-direction: column; gap: 4px; font-size: 0.82rem; font-weight: 600; }
    .form-control { padding: 8px; border: 1px solid var(--border); border-radius: 6px; }
  </style>
</head>
<body>
  <aside>
    <div class="brand"><i class="bi bi-box-seam-fill" style="color:#38bdf8;"></i> MeatFlow ERP</div>
    <div class="nav-item active" onclick="switchMainTab('INBOUND', this)"><i class="bi bi-box-arrow-in-down"></i> 입고 관리</div>
    <div class="nav-item" onclick="switchMainTab('OUTBOUND', this)"><i class="bi bi-box-arrow-up"></i> 출고 관리 (재고장)</div>
  </aside>
  <main>
    <header>
      <h2 id="pageTitle" style="font-size:1.2rem;">입고 프로세스 관리</h2>
      <div id="headerActions"><button class="btn btn-primary" onclick="openInboundModal()">+ 신규 입고요청</button></div>
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
  <div class="modal-overlay" id="inboundModal">
    <div class="modal-box">
      <h3 style="margin-bottom:12px;">신규 입고요청</h3>
      <div class="form-group"><label>매입처</label><input type="text" id="in_vendor" class="form-control" value="(주)글로벌미트" /></div>
      <div class="form-group"><label>품목명</label><input type="text" id="in_item" class="form-control" value="돼지 삼겹살(올리멜)" /></div>
      <div class="form-group"><label>보관형태</label><select id="in_storage" class="form-control"><option>냉장</option><option>냉동</option></select></div>
      <div class="form-group"><label>이력번호</label><input type="text" id="in_trace" class="form-control" value="802410290114" /></div>
      <div style="display:flex; gap:10px;"><div class="form-group" style="flex:1;"><label>박스</label><input type="number" id="in_box" class="form-control" value="50" /></div><div class="form-group" style="flex:1;"><label>중량(kg)</label><input type="number" step="0.1" id="in_weight" class="form-control" value="1020.5" /></div></div>
      <div style="display:flex; gap:10px;"><div class="form-group" style="flex:1;"><label>단가/kg</label><input type="number" id="in_cost" class="form-control" value="8400" /></div><div class="form-group" style="flex:1;"><label>유통기한</label><input type="date" id="in_exp" class="form-control" value="2026-09-15" /></div></div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:14px;"><button class="btn" style="background:#e2e8f0;" onclick="closeModal()">취소</button><button class="btn btn-primary" onclick="submitInbound()">저장</button></div>
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
      document.getElementById('headerActions').innerHTML = tab === 'INBOUND' ? '<button class="btn btn-primary" onclick="openInboundModal()">+ 신규 입고요청</button>' : '';
      renderSubTabs(); loadData();
    }
    function renderSubTabs() {
      const container = document.getElementById('subTabs'); container.innerHTML = '';
      const steps = currentMain === 'INBOUND' ? [{k:'IN_REQUEST',n:'1. 입고요청'},{k:'IN_CONFIRM',n:'2. 입고확정'},{k:'IN_DONE',n:'3. 입고완료'}] : [{k:'OUT_STOCK',n:'1. 출하요청(재고장)'},{k:'OUT_REQUEST',n:'2. 출고요청'},{k:'OUT_CONFIRM',n:'3. 출고확정'},{k:'OUT_DONE',n:'4. 출고완료'}];
      steps.forEach(s => { container.innerHTML += `<div class="sub-tab ${currentSub===s.k?'active':''}" onclick="setSubTab('${s.k}')">${s.n}</div>`; });
    }
    function setSubTab(k) { currentSub = k; renderSubTabs(); loadData(); }
    async function loadData() {
      const head = document.getElementById('tableHead'), body = document.getElementById('tableBody');
      body.innerHTML = '<tr><td colspan="8" style="text-align:center; padding:15px;">로딩 중...</td></tr>';
      if (currentMain === 'INBOUND') {
        head.innerHTML = '<tr><th>전표번호</th><th>매입처</th><th>품목명</th><th>보관</th><th>이력번호</th><th>박스</th><th>중량(kg)</th><th>처리</th></tr>';
        const res = await fetch(`/api/inbounds?status=${currentSub}`), list = await res.json();
        body.innerHTML = list.length ? list.map(i => `<tr><td><strong>${i.inbound_no}</strong></td><td>${i.vendor}</td><td>${i.item_name}</td><td>${i.storage_type}</td><td><code>${i.trace_no}</code></td><td>${i.box_qty} Box</td><td>${i.weight_kg} kg</td><td>${currentSub==='IN_REQUEST'?`<button class="btn btn-primary" onclick="advIn(${i.id})">확정</button>`:currentSub==='IN_CONFIRM'?`<button class="btn btn-success" onclick="advIn(${i.id})">입고완료</button>`:'완료됨'}</td></tr>`).join('') : '<tr><td colspan="8" style="text-align:center; padding:20px; color:#94a3b8;">내역 없음</td></tr>';
      } else {
        if (currentSub === 'OUT_STOCK') {
          head.innerHTML = '<tr><th>ID</th><th>품목명</th><th>보관</th><th>이력번호</th><th>보유박스</th><th>보유중량(kg)</th><th>창고</th><th>출하</th></tr>';
          const res = await fetch('/api/inventory'), list = await res.json();
          body.innerHTML = list.length ? list.map(l => `<tr><td>#${l.id}</td><td>${l.item_name}</td><td>${l.storage_type}</td><td><code>${l.trace_no}</code></td><td>${l.current_box_qty} Box</td><td>${l.current_weight_kg} kg</td><td>${l.warehouse}</td><td><button class="btn btn-warning" onclick="reqOut(${l.id},${l.current_box_qty})">출하요청</button></td></tr>`).join('') : '<tr><td colspan="8" style="text-align:center; padding:20px; color:#94a3b8;">재고 없음</td></tr>';
        } else {
          head.innerHTML = '<tr><th>출고번호</th><th>매출처</th><th>품목명</th><th>이력번호</th><th>박스</th><th>중량(kg)</th><th>단가</th><th>처리</th></tr>';
          const res = await fetch(`/api/outbounds?status=${currentSub}`), list = await res.json();
          body.innerHTML = list.length ? list.map(o => `<tr><td><strong>${o.outbound_no}</strong></td><td>${o.customer}</td><td>${o.item_name}</td><td><code>${o.trace_no}</code></td><td>${o.box_qty} Box</td><td>${o.weight_kg} kg</td><td>${o.unit_price_kg}원</td><td>${currentSub==='OUT_REQUEST'?`<button class="btn btn-primary" onclick="advOut(${o.id})">확정</button>`:currentSub==='OUT_CONFIRM'?`<button class="btn btn-success" onclick="advOut(${o.id})">출고완료</button>`:'완료됨'}</td></tr>`).join('') : '<tr><td colspan="8" style="text-align:center; padding:20px; color:#94a3b8;">내역 없음</td></tr>';
        }
      }
    }
    async function advIn(id) { const r = await fetch(`/api/inbounds/${id}/advance`,{method:'POST'}); const d = await r.json(); alert(d.message); loadData(); }
    async function advOut(id) { const r = await fetch(`/api/outbounds/${id}/advance`,{method:'POST'}); const d = await r.json(); alert(d.message); loadData(); }
    async function reqOut(lotId, maxBox) {
      const cust = prompt("매출처명 입력:", "(주)신규거래처"); if(!cust) return;
      const box = prompt(`박스 수 (보유: ${maxBox}):`, "10"); if(!box) return;
      const r = await fetch('/api/outbounds/create-from-stock', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({lot_id:lotId, customer:cust, box_qty:Number(box), unit_price_kg:11000})});
      const d = await r.json(); alert(d.message); setSubTab('OUT_REQUEST');
    }
    function openInboundModal() { document.getElementById('inboundModal').style.display = 'flex'; }
    function closeModal() { document.getElementById('inboundModal').style.display = 'none'; }
    async function submitInbound() {
      const p = { vendor: document.getElementById('in_vendor').value, item_name: document.getElementById('in_item').value, storage_type: document.getElementById('in_storage').value, trace_no: document.getElementById('in_trace').value, box_qty: Number(document.getElementById('in_box').value), weight_kg: Number(document.getElementById('in_weight').value), cost_per_kg: Number(document.getElementById('in_cost').value), warehouse: "광주냉장창고", exp_date: document.getElementById('in_exp').value };
      const r = await fetch('/api/inbounds', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p)});
      const d = await r.json(); alert(d.message); closeModal(); setSubTab('IN_REQUEST');
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
