"""
================================================================================
(주)서울웰푸드 미트플로우(MeatFlow) ERP - [모듈 1. 기준정보 마스터 엔진]
- 기본 법인: (주)서울웰푸드 (사업자번호: 347-81-03002 / 대표이사: 조용훈)
- 주소: 서울특별시 강동구 천중로 39길 19-25 (천호동, 평영빌딩)
- 기능: 
  1. 당사 멀티법인 마스터 (서울웰푸드, 티제이에프, 넥서스트레이딩)
  2. 6대 거래처 마스터 및 여신 한도 / 결제조건 / 여신 락 통제
  3. 외부 보관창고별 일할 보관료(양편넣기) 및 급속동결/상하차비 요율
  4. 축종(우육/돈육), EST(가공장), 브랜드, 부위 표준 규격
- 실행: uvicorn module1_master:app --host 0.0.0.0 --port 8000 --reload
================================================================================
"""

import os
from datetime import datetime
from typing import List, Optional

import bcrypt
from fastapi import FastAPI, HTTPException, Depends, status, Query
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from sqlalchemy import (
    Column, Integer, Float, String, DateTime, create_engine, desc, text
)
from sqlalchemy.orm import declarative_base, sessionmaker, Session

# -----------------------------------------------------------------------------
# 1. DB 연결 및 엔진 설정
# -----------------------------------------------------------------------------
DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///./meatflow_m1_master.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

connect_args = {"check_same_thread": False} if "sqlite" in DATABASE_URL else {}
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
# 2. SQLAlchemy ORM 마스터 데이터 모델
# -----------------------------------------------------------------------------
class CompanyMaster(Base):
    """1-1. 당사 법인 마스터 (기본 화주: 서울웰푸드 / 관계사: 티제이에프, 넥서스트레이딩)"""
    __tablename__ = "companies"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String(50), unique=True, nullable=False, index=True)
    name = Column(String(100), nullable=False)
    biz_no = Column(String(30), nullable=False)
    rep_name = Column(String(50), nullable=False)
    address = Column(String(200), nullable=False)
    phone = Column(String(30), nullable=True)
    fax = Column(String(30), nullable=True)
    bank_name = Column(String(50), nullable=True)
    bank_account = Column(String(50), nullable=True)
    stamp_text = Column(String(50), default="서울웰푸드인")
    is_default = Column(Integer, default=0) # 1: 기본 화주
    created_at = Column(DateTime, default=datetime.now)

class PartnerMaster(Base):
    """1-2. 6대 거래처 마스터 (매입처, 매출처, 매입매출처, 창고, 파이낸스사, 운송주선사)"""
    __tablename__ = "partners"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String(50), unique=True, nullable=False, index=True)
    name = Column(String(100), unique=True, nullable=False, index=True)
    partner_type = Column(String(30), nullable=False, index=True) 
    biz_no = Column(String(30), nullable=True)
    rep_name = Column(String(50), nullable=True)
    contact_person = Column(String(50), nullable=True)
    phone = Column(String(30), nullable=True)
    email = Column(String(100), nullable=True)
    address = Column(String(200), nullable=True)
    # 여신 및 결제조건 통제 필드
    credit_limit = Column(Float, default=0.0)        # 여신한도 (원)
    collateral_amount = Column(Float, default=0.0)   # 담보설정액 (원)
    payment_term = Column(String(50), default="외상 30일 회전") # 결제조건
    is_credit_locked = Column(Integer, default=0)    # 1: 여신초과/미수연체 출고차단 락
    remark = Column(String(255), nullable=True)
    created_at = Column(DateTime, default=datetime.now)

class WarehouseRateMaster(Base):
    """1-3. 창고별 일할 보관료(양편넣기) 및 하역/동결 작업비 요율 마스터"""
    __tablename__ = "warehouse_rates"
    id = Column(Integer, primary_key=True, index=True)
    warehouse_name = Column(String(100), nullable=False, index=True)
    storage_type = Column(String(20), nullable=False) # 냉동, 냉장
    daily_rate_kg = Column(Float, default=1.5)        # 원/kg/일 (양편넣기 기본 요율)
    freezing_fee_kg = Column(Float, default=80.0)     # 급속동결비 원/kg (냉장육 냉동전환 시 산입)
    handling_in_kg = Column(Float, default=5.0)       # 입고 하차비 원/kg
    handling_out_kg = Column(Float, default=5.0)      # 출고 상차비 원/kg
    pallet_storage_daily = Column(Float, default=1500.0)
    remark = Column(String(255), nullable=True)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

class ItemCutMaster(Base):
    """1-4. 축종(우육/돈육), EST(가공장), 브랜드, 부위 표준 규격 마스터"""
    __tablename__ = "item_cuts"
    id = Column(Integer, primary_key=True, index=True)
    item_code = Column(String(50), unique=True, nullable=False, index=True)
    species = Column(String(30), nullable=False)      # 우육, 돈육, 계육
    cut_name = Column(String(100), nullable=False)    # 소일반갈비(170), 돈갈매기살 등
    est_no = Column(String(50), nullable=False)       # 244I, EST-717, 1614 등
    brand = Column(String(50), nullable=False)        # AMH, 킬코이, SWIFT, IBP 등
    default_grade = Column(String(30), default="GF")  # GF, CHOICE, PRIME
    origin = Column(String(50), default="호주")       # 호주, 미국, 캐나다
    default_storage = Column(String(20), default="냉동") # 냉동, 냉장
    shelf_life_months = Column(Integer, default=24)   # 유통기한 기본 개월수
    created_at = Column(DateTime, default=datetime.now)

class UserMaster(Base):
    """1-5. 시스템 사용자 및 RBAC 역할 모델"""
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(50), unique=True, nullable=False, index=True)
    hashed_password = Column(String(255), nullable=False)
    full_name = Column(String(50), nullable=False)
    role = Column(String(20), default="SALES") # ADMIN, SALES, SALES_SUPPORT, WAREHOUSE
    is_active = Column(Integer, default=1)
    created_at = Column(DateTime, default=datetime.now)

Base.metadata.create_all(bind=engine)

# -----------------------------------------------------------------------------
# 3. 초기 마스터 데이터 시딩
# -----------------------------------------------------------------------------
def hash_pw(plain: str) -> str:
    return bcrypt.hashpw(plain.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def init_master_seed():
    db = SessionLocal()
    try:
        if not db.query(CompanyMaster).first():
            comps = [
                CompanyMaster(
                    code="SEOUL", name="(주)서울웰푸드", biz_no="347-81-03002", rep_name="조용훈",
                    address="서울특별시 강동구 천중로 39길 19-25(천호동,평영빌딩)", phone="02-6958-9229", fax="070-4758-9219",
                    bank_name="기업은행", bank_account="0347-81030-02-011", stamp_text="서울웰푸드인", is_default=1
                ),
                CompanyMaster(
                    code="TJF", name="주식회사 티제이에프", biz_no="718-88-03523", rep_name="김다희",
                    address="경기도 남양주시 별내중앙로 34, 5층 502-A59호", phone="010-5773-0619", fax="070-4758-9219",
                    bank_name="국민은행", bank_account="71888-03523-01", stamp_text="티제이에프인", is_default=0
                ),
                CompanyMaster(
                    code="NEXUS", name="(주)넥서스트레이딩", biz_no="518-86-03633", rep_name="김다희",
                    address="서울특별시 강동구 천중로 39길 19-25(천호동,평영빌딩),2층", phone="010-5773-0619", fax="070-4758-9219",
                    bank_name="하나은행", bank_account="51886-03633-01", stamp_text="넥서스인", is_default=0
                )
            ]
            db.add_all(comps)

        if not db.query(PartnerMaster).first():
            partners = [
                PartnerMaster(code="P-STF", name="에스티에프", partner_type="매입매출처", biz_no="128-86-99102", rep_name="이대표", contact_person="이영업 부장", phone="010-2211-3344", credit_limit=150000000, payment_term="외상 30일 회전"),
                PartnerMaster(code="P-PLANB", name="(주)플랜비에프에스", partner_type="매출처", biz_no="214-88-77192", rep_name="박플랜", contact_person="최과장", phone="010-9988-1122", credit_limit=120000000, payment_term="외상 30일 회전"),
                PartnerMaster(code="P-BRUNI", name="브루니", partner_type="매입매출처", biz_no="134-87-00192", rep_name="정대표", contact_person="김대리", phone="010-4433-2211", credit_limit=100000000, payment_term="외상 15일 회전"),
                PartnerMaster(code="P-MEATKO", name="미코상사", partner_type="매입처", biz_no="110-81-22991", rep_name="강대표", contact_person="강실장", phone="02-2299-0011", credit_limit=0, payment_term="선입금 확인후 상차"),
                PartnerMaster(code="P-KOREAMEAT", name="대한육류유통", partner_type="매출처", biz_no="215-81-44910", rep_name="윤대표", contact_person="조과장", phone="010-3344-5566", credit_limit=80000000, payment_term="외상 30일 회전"),
                PartnerMaster(code="P-WH-KD2", name="강동2", partner_type="창고", biz_no="211-85-11029", contact_person="창고반장", phone="031-768-1100", credit_limit=0),
                PartnerMaster(code="P-WH-SJ1", name="삼진1", partner_type="창고", biz_no="211-85-44919", contact_person="하역반장", phone="031-768-2200", credit_limit=0),
                PartnerMaster(code="P-FIN-DONGWON", name="동원홈푸드", partner_type="파이낸스사", biz_no="107-86-09192", contact_person="금융팀장", phone="02-589-3000", credit_limit=0),
                PartnerMaster(code="P-TR-24CALL", name="24시콜화물", partner_type="운송주선사", biz_no="119-86-77881", contact_person="배차실장", phone="1588-0024", credit_limit=0)
            ]
            db.add_all(partners)

        if not db.query(WarehouseRateMaster).first():
            rates = [
                WarehouseRateMaster(warehouse_name="강동2", storage_type="냉동", daily_rate_kg=1.5, freezing_fee_kg=80.0, handling_in_kg=5.0, handling_out_kg=5.0, remark="기본 냉동요율 1.5원"),
                WarehouseRateMaster(warehouse_name="강동2", storage_type="냉장", daily_rate_kg=2.0, freezing_fee_kg=80.0, handling_in_kg=5.0, handling_out_kg=5.0, remark="냉장 일할 2.0원"),
                WarehouseRateMaster(warehouse_name="삼진1", storage_type="냉동", daily_rate_kg=1.5, freezing_fee_kg=80.0, handling_in_kg=5.0, handling_out_kg=5.0, remark="삼진 냉동 1.5원"),
                WarehouseRateMaster(warehouse_name="삼진1", storage_type="냉장", daily_rate_kg=2.0, freezing_fee_kg=80.0, handling_in_kg=5.0, handling_out_kg=5.0, remark="삼진 냉장 2.0원"),
                WarehouseRateMaster(warehouse_name="광주냉장", storage_type="냉동", daily_rate_kg=1.4, freezing_fee_kg=80.0, handling_in_kg=4.5, handling_out_kg=4.5, remark="광주 물류센터")
            ]
            db.add_all(rates)

        if not db.query(ItemCutMaster).first():
            items = [
                ItemCutMaster(item_code="BF-AMH-170", species="우육", cut_name="소일반갈비(170)", est_no="244I", brand="AMH", default_grade="GF", origin="호주", default_storage="냉동"),
                ItemCutMaster(item_code="BF-AMH-235", species="우육", cut_name="소일반갈비(235)", est_no="244I", brand="AMH", default_grade="GF", origin="호주", default_storage="냉동"),
                ItemCutMaster(item_code="PK-KC-GUL", species="돈육", cut_name="돈갈매기살", est_no="1614", brand="킬코이", default_grade="GF", origin="호주", default_storage="냉동"),
                ItemCutMaster(item_code="BF-IBP-RIB", species="우육", cut_name="냉장 소갈비살", est_no="352", brand="IBP", default_grade="CHOICE", origin="미국", default_storage="냉장", shelf_life_months=3),
                ItemCutMaster(item_code="BF-AMH-TEN", species="우육", cut_name="소안심", est_no="244I", brand="AMH", default_grade="PRIME", origin="호주", default_storage="냉동")
            ]
            db.add_all(items)

        if not db.query(UserMaster).first():
            users = [
                UserMaster(username="admin", hashed_password=hash_pw("admin1234!"), full_name="조용훈 대표", role="ADMIN"),
                UserMaster(username="sales_kim", hashed_password=hash_pw("sales1234!"), full_name="김영업 과장", role="SALES"),
                UserMaster(username="support_lee", hashed_password=hash_pw("support1234!"), full_name="이영업지원 대리", role="SALES_SUPPORT"),
                UserMaster(username="wh_park", hashed_password=hash_pw("wh1234!"), full_name="박창고 반장", role="WAREHOUSE")
            ]
            db.add_all(users)

        db.commit()
    finally:
        db.close()

init_master_seed()

# -----------------------------------------------------------------------------
# 4. FastAPI Pydantic DTO 스키마
# -----------------------------------------------------------------------------
app = FastAPI(title="(주)서울웰푸드 미트플로우(MeatFlow) ERP - 모듈 1. 기준정보", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class CompanyDTO(BaseModel):
    code: str
    name: str
    biz_no: str
    rep_name: str
    address: str
    phone: Optional[str] = ""
    fax: Optional[str] = ""
    bank_name: Optional[str] = ""
    bank_account: Optional[str] = ""
    stamp_text: Optional[str] = "서울웰푸드인"
    is_default: Optional[int] = 0

class PartnerDTO(BaseModel):
    code: str
    name: str
    partner_type: str
    biz_no: Optional[str] = ""
    rep_name: Optional[str] = ""
    contact_person: Optional[str] = ""
    phone: Optional[str] = ""
    email: Optional[str] = ""
    address: Optional[str] = ""
    credit_limit: float = 0.0
    collateral_amount: float = 0.0
    payment_term: str = "외상 30일 회전"
    remark: Optional[str] = ""

class WarehouseRateDTO(BaseModel):
    warehouse_name: str
    storage_type: str
    daily_rate_kg: float = 1.5
    freezing_fee_kg: float = 80.0
    handling_in_kg: float = 5.0
    handling_out_kg: float = 5.0
    remark: Optional[str] = ""

class ItemCutDTO(BaseModel):
    item_code: str
    species: str
    cut_name: str
    est_no: str
    brand: str
    default_grade: str = "GF"
    origin: str = "호주"
    default_storage: str = "냉동"
    shelf_life_months: int = 24

# -----------------------------------------------------------------------------
# 5. REST API 엔드포인트
# -----------------------------------------------------------------------------
@app.get("/api/masters/companies")
def get_companies(db: Session = Depends(get_db)):
    return db.query(CompanyMaster).order_by(desc(CompanyMaster.is_default), CompanyMaster.id).all()

@app.post("/api/masters/companies")
def create_company(req: CompanyDTO, db: Session = Depends(get_db)):
    if db.query(CompanyMaster).filter(CompanyMaster.code == req.code).first():
        raise HTTPException(status_code=400, detail="이미 등록된 법인 코드입니다.")
    comp = CompanyMaster(**req.model_dump())
    db.add(comp)
    db.commit()
    return {"message": f"법인 '{req.name}' 등록 완료"}

@app.get("/api/masters/partners")
def get_partners(partner_type: Optional[str] = None, db: Session = Depends(get_db)):
    q = db.query(PartnerMaster)
    if partner_type and partner_type != "ALL":
        q = q.filter(PartnerMaster.partner_type == partner_type)
    return q.order_by(PartnerMaster.name).all()

@app.post("/api/masters/partners")
def create_partner(req: PartnerDTO, db: Session = Depends(get_db)):
    if db.query(PartnerMaster).filter(PartnerMaster.name == req.name).first():
        raise HTTPException(status_code=400, detail="이미 존재하는 거래처명입니다.")
    partner = PartnerMaster(**req.model_dump())
    db.add(partner)
    db.commit()
    return {"message": f"거래처 '{req.name}' ({req.partner_type}) 등록 완료"}

@app.patch("/api/masters/partners/{partner_id}/toggle-lock")
def toggle_credit_lock(partner_id: int, db: Session = Depends(get_db)):
    partner = db.query(PartnerMaster).filter(PartnerMaster.id == partner_id).first()
    if not partner:
        raise HTTPException(status_code=404, detail="거래처를 찾을 수 없습니다.")
    partner.is_credit_locked = 0 if partner.is_credit_locked == 1 else 1
    db.commit()
    status_str = "여신 락(출고차단) 설정" if partner.is_credit_locked == 1 else "여신 정상 해제"
    return {"message": f"'{partner.name}' {status_str} 완료", "is_credit_locked": partner.is_credit_locked}

@app.get("/api/masters/warehouse-rates")
def get_warehouse_rates(db: Session = Depends(get_db)):
    return db.query(WarehouseRateMaster).order_by(WarehouseRateMaster.warehouse_name, WarehouseRateMaster.storage_type).all()

@app.post("/api/masters/warehouse-rates")
def save_warehouse_rate(req: WarehouseRateDTO, db: Session = Depends(get_db)):
    rate = db.query(WarehouseRateMaster).filter(
        WarehouseRateMaster.warehouse_name == req.warehouse_name,
        WarehouseRateMaster.storage_type == req.storage_type
    ).first()
    if rate:
        rate.daily_rate_kg = req.daily_rate_kg
        rate.freezing_fee_kg = req.freezing_fee_kg
        rate.handling_in_kg = req.handling_in_kg
        rate.handling_out_kg = req.handling_out_kg
        rate.remark = req.remark
        msg = f"'{req.warehouse_name} ({req.storage_type})' 요율 수정 완료"
    else:
        rate = WarehouseRateMaster(**req.model_dump())
        db.add(rate)
        msg = f"'{req.warehouse_name} ({req.storage_type})' 요율 신규 등록 완료"
    db.commit()
    return {"message": msg}

@app.get("/api/masters/item-cuts")
def get_item_cuts(species: Optional[str] = None, db: Session = Depends(get_db)):
    q = db.query(ItemCutMaster)
    if species and species != "ALL":
        q = q.filter(ItemCutMaster.species == species)
    return q.order_by(ItemCutMaster.species, ItemCutMaster.cut_name).all()

@app.post("/api/masters/item-cuts")
def create_item_cut(req: ItemCutDTO, db: Session = Depends(get_db)):
    if db.query(ItemCutMaster).filter(ItemCutMaster.item_code == req.item_code).first():
        raise HTTPException(status_code=400, detail="이미 등록된 품목 코드입니다.")
    item = ItemCutMaster(**req.model_dump())
    db.add(item)
    db.commit()
    return {"message": f"품목 '{req.cut_name}' ({req.brand}) 등록 완료"}

# -----------------------------------------------------------------------------
# 6. 미트플로우 기준정보 전용 SPA 프론트엔드 UI
# -----------------------------------------------------------------------------
HTML_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>(주)서울웰푸드 미트플로우 ERP - 1. 기준정보 관리</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
  <style>
    :root {
      --primary: #1e3a8a; --primary-hover: #1e40af; --sidebar: #0f172a; --bg: #f8fafc;
      --border: #cbd5e1; --text: #0f172a; --muted: #64748b; --danger: #dc2626; --success: #16a34a;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Noto Sans KR", sans-serif; }
    body { background: var(--bg); color: var(--text); display: flex; height: 100vh; overflow: hidden; }

    aside { width: 250px; background: var(--sidebar); color: #fff; display: flex; flex-direction: column; flex-shrink: 0; }
    .brand-box { padding: 18px 16px; border-bottom: 1px solid #1e293b; }
    .brand-title { font-size: 15px; font-weight: 800; color: #38bdf8; display: flex; align-items: center; gap: 8px; }
    .brand-sub { font-size: 10.5px; color: #94a3b8; margin-top: 4px; }
    .nav-list { list-style: none; padding: 12px 8px; flex: 1; overflow-y: auto; }
    .nav-header { font-size: 10px; font-weight: 700; color: #64748b; padding: 8px 10px 4px; text-transform: uppercase; }
    .nav-item { display: flex; align-items: center; gap: 10px; padding: 9px 12px; border-radius: 6px; font-size: 12.5px; font-weight: 600; color: #cbd5e1; cursor: pointer; margin-bottom: 2px; }
    .nav-item:hover, .nav-item.active { background: var(--primary); color: #fff; }

    main { flex: 1; display: flex; flex-direction: column; overflow: hidden; }
    header { background: #fff; border-bottom: 1px solid var(--border); padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; }
    .content-area { flex: 1; padding: 20px 24px; overflow-y: auto; }

    .sub-tabs { display: flex; gap: 8px; margin-bottom: 16px; border-bottom: 2px solid var(--border); padding-bottom: 8px; }
    .tab-btn { padding: 8px 16px; border: none; background: #e2e8f0; border-radius: 6px; font-size: 12px; font-weight: 700; cursor: pointer; color: var(--muted); }
    .tab-btn.active { background: var(--primary); color: #fff; }

    .card { background: #fff; border-radius: 8px; border: 1px solid var(--border); overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.02); }
    .toolbar { padding: 12px 16px; background: #f8fafc; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; text-align: left; }
    th { background: #f1f5f9; padding: 9px 12px; font-weight: 700; color: #475569; border-bottom: 1px solid var(--border); white-space: nowrap; }
    td { padding: 9px 12px; border-bottom: 1px solid var(--border); vertical-align: middle; }
    tr:hover { background: #f8fafc; }

    .badge { padding: 2px 6px; border-radius: 4px; font-size: 10.5px; font-weight: 700; display: inline-block; }
    .badge-blue { background: #dbeafe; color: #1e40af; }
    .badge-green { background: #dcfce7; color: #166534; }
    .badge-orange { background: #ffedd5; color: #c2410c; }
    .badge-red { background: #fee2e2; color: #991b1b; }
    .badge-gray { background: #f1f5f9; color: #475569; }

    .btn { padding: 7px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; border: none; display: inline-flex; align-items: center; gap: 4px; }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-outline { background: #fff; border: 1px solid var(--border); color: var(--text); }

    .modal-overlay { position: fixed; inset: 0; background: rgba(15,23,42,0.6); display: none; align-items: center; justify-content: center; z-index: 1000; }
    .modal-box { background: #fff; width: 680px; max-width: 95%; border-radius: 10px; overflow: hidden; }
    .modal-header { padding: 12px 18px; background: #0f172a; color: #fff; font-weight: 700; font-size: 13px; display: flex; justify-content: space-between; align-items: center; }
    .modal-body { padding: 18px; max-height: 75vh; overflow-y: auto; }
    .form-group { margin-bottom: 12px; }
    .form-group label { display: block; font-size: 11.5px; font-weight: 700; color: #334155; margin-bottom: 4px; }
    .form-control { width: 100%; padding: 7px 10px; border: 1px solid var(--border); border-radius: 5px; font-size: 12px; }
    .modal-footer { padding: 10px 18px; background: #f8fafc; border-top: 1px solid var(--border); display: flex; justify-content: flex-end; gap: 6px; }
  </style>
</head>
<body>
  <aside>
    <div class="brand-box">
      <div class="brand-title"><i class="bi bi-box-seam-fill"></i> 미트플로우 ERP</div>
      <div class="brand-sub">(주)서울웰푸드 맞춤형</div>
    </div>
    <ul class="nav-list">
      <div class="nav-header">기본 관제</div>
      <li class="nav-item" onclick="alert('모듈 0은 대시보드 뷰입니다.')"><i class="bi bi-speedometer2"></i> 0. 경영 대시보드</li>
      <div class="nav-header">기준정보 (모듈 1)</div>
      <li class="nav-item active"><i class="bi bi-gear-fill"></i> 1. 기준정보 관리</li>
      <div class="nav-header">수입 / 물류 / 통관</div>
      <li class="nav-item" onclick="alert('모듈 2에서 제공됩니다.')"><i class="bi bi-water"></i> 2. 무역 / 선적일정</li>
      <li class="nav-item" onclick="alert('모듈 3에서 제공됩니다.')"><i class="bi bi-shield-check"></i> 3. 보세 / 하차검수</li>
      <div class="nav-header">재고 / 영업 / 정산</div>
      <li class="nav-item" onclick="alert('모듈 4에서 제공됩니다.')"><i class="bi bi-grid-3x3-gap-fill"></i> 4. 재고 / 보관관리</li>
      <li class="nav-item" onclick="alert('모듈 5에서 제공됩니다.')"><i class="bi bi-truck"></i> 5. 영업 / 출고배차</li>
    </ul>
  </aside>

  <main>
    <header>
      <div style="font-size:16px; font-weight:800; color:var(--text);"><i class="bi bi-gear-fill"></i> 1. 기준정보 마스터 관리 (MeatFlow Core Master)</div>
      <div id="headerActions"></div>
    </header>

    <div class="content-area">
      <div class="sub-tabs">
        <button class="tab-btn active" onclick="switchSubTab('PARTNER', this)"><i class="bi bi-people-fill"></i> 6대 거래처 & 여신/결제조건</button>
        <button class="tab-btn" onclick="switchSubTab('COMPANY', this)"><i class="bi bi-building"></i> 당사 법인 마스터</button>
        <button class="tab-btn" onclick="switchSubTab('RATE', this)"><i class="bi bi-snow2"></i> 창고 일할요율 & 작업비</button>
        <button class="tab-btn" onclick="switchSubTab('ITEM', this)"><i class="bi bi-tag-fill"></i> 품목 / 부위 / EST 규격</button>
      </div>

      <div class="card">
        <div class="toolbar" id="toolbarArea"></div>
        <div style="overflow-x:auto;">
          <table>
            <thead id="tblHead"></thead>
            <tbody id="tblBody"></tbody>
          </table>
        </div>
      </div>
    </div>
  </main>

  <div class="modal-overlay" id="masterModal">
    <div class="modal-box">
      <div class="modal-header">
        <span id="modalTitle">마스터 등록</span>
        <i class="bi bi-x-lg" style="cursor:pointer;" onclick="closeModal()"></i>
      </div>
      <div class="modal-body" id="modalBody"></div>
      <div class="modal-footer">
        <button class="btn btn-outline" onclick="closeModal()">취소</button>
        <button class="btn btn-primary" onclick="submitModal()">저장하기</button>
      </div>
    </div>
  </div>

  <script>
    let curSubTab = 'PARTNER';

    async function apiFetch(url, opt={}) {
      const res = await fetch(url, opt);
      if (!res.ok) {
        const err = await res.json().catch(()=>({detail:'오류 발생'}));
        throw new Error(err.detail || '서버 오류');
      }
      return res.json();
    }

    function switchSubTab(tab, el) {
      curSubTab = tab;
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      if (el) el.classList.add('active');
      renderView();
    }

    async function renderView() {
      const hActions = document.getElementById('headerActions');
      const tb = document.getElementById('toolbarArea');
      const th = document.getElementById('tblHead');
      const tbBody = document.getElementById('tblBody');

      if (curSubTab === 'PARTNER') {
        hActions.innerHTML = `<button class="btn btn-primary" onclick="openCreateModal('PARTNER')"><i class="bi bi-plus-lg"></i> 신규 거래처 등록</button>`;
        tb.innerHTML = `
          <strong>📋 6대 거래처 (매입처, 매출처, 매입매출처, 창고, 파이낸스사, 운송사) 및 여신 한도</strong>
          <select class="form-control" style="width:160px;" onchange="filterPartner(this.value)">
            <option value="ALL">전체 구분</option>
            <option value="매입처">매입처</option>
            <option value="매출처">매출처</option>
            <option value="매입매출처">매입매출처</option>
            <option value="창고">보관창고</option>
            <option value="파이낸스사">파이낸스사</option>
            <option value="운송주선사">운송주선사</option>
          </select>
        `;
        th.innerHTML = `<tr><th>코드</th><th>거래처명</th><th>구분</th><th>사업자번호</th><th>대표/담당자</th><th>연락처</th><th>여신한도</th><th>결제조건</th><th>여신통제</th></tr>`;
        loadPartners('ALL');
      } else if (curSubTab === 'COMPANY') {
        hActions.innerHTML = `<button class="btn btn-primary" onclick="openCreateModal('COMPANY')"><i class="bi bi-plus-lg"></i> 신규 법인 등록</button>`;
        tb.innerHTML = `<strong>🏢 당사 화주 법인 마스터 (기본 발행처: (주)서울웰푸드)</strong>`;
        th.innerHTML = `<tr><th>코드</th><th>법인명</th><th>사업자번호</th><th>대표자</th><th>주소</th><th>전화 / 팩스</th><th>기본화주</th></tr>`;
        const list = await apiFetch('/api/masters/companies');
        tbBody.innerHTML = list.map(c => `
          <tr>
            <td><strong>${c.code}</strong></td>
            <td><strong>${c.name}</strong></td>
            <td>${c.biz_no}</td>
            <td>${c.rep_name}</td>
            <td>${c.address}</td>
            <td>${c.phone||'-'} / ${c.fax||'-'}</td>
            <td>${c.is_default ? '<span class="badge badge-green">기본화주</span>' : '<span class="badge badge-gray">관계사</span>'}</td>
          </tr>
        `).join('');
      } else if (curSubTab === 'RATE') {
        hActions.innerHTML = `<button class="btn btn-primary" onclick="openCreateModal('RATE')"><i class="bi bi-plus-lg"></i> 창고 요율 등록/수정</button>`;
        tb.innerHTML = `<strong>❄️ 외부 보관창고별 일할 보관료 (양편넣기) 및 급속동결/상하차비</strong>`;
        th.innerHTML = `<tr><th>창고명</th><th>보관온도</th><th>일할보관료 (원/kg/일)</th><th>급속동결비 (원/kg)</th><th>입고하차비 (원/kg)</th><th>출고상차비 (원/kg)</th><th>비고</th></tr>`;
        const list = await apiFetch('/api/masters/warehouse-rates');
        tbBody.innerHTML = list.map(r => `
          <tr>
            <td><strong>${r.warehouse_name}</strong></td>
            <td><span class="badge ${r.storage_type==='냉장'?'badge-orange':'badge-blue'}">${r.storage_type}</span></td>
            <td style="color:#1d4ed8; font-weight:700;">${r.daily_rate_kg} 원</td>
            <td>${r.freezing_fee_kg} 원</td>
            <td>${r.handling_in_kg} 원</td>
            <td>${r.handling_out_kg} 원</td>
            <td>${r.remark||'-'}</td>
          </tr>
        `).join('');
      } else if (curSubTab === 'ITEM') {
        hActions.innerHTML = `<button class="btn btn-primary" onclick="openCreateModal('ITEM')"><i class="bi bi-plus-lg"></i> 신규 품목 등록</button>`;
        tb.innerHTML = `<strong>🥩 축종(우육/돈육), EST(가공장), 브랜드, 부위 표준 규격 마스터</strong>`;
        th.innerHTML = `<tr><th>품목코드</th><th>축종</th><th>부위명</th><th>가공장(EST)</th><th>브랜드</th><th>기본등급</th><th>원산지</th><th>기본보관</th></tr>`;
        const list = await apiFetch('/api/masters/item-cuts');
        tbBody.innerHTML = list.map(i => `
          <tr>
            <td><strong>${i.item_code}</strong></td>
            <td><span class="badge ${i.species==='우육'?'badge-red':'badge-green'}">${i.species}</span></td>
            <td><strong>${i.cut_name}</strong></td>
            <td>${i.est_no}</td>
            <td>${i.brand}</td>
            <td>${i.default_grade}</td>
            <td>${i.origin}</td>
            <td>${i.default_storage}</td>
          </tr>
        `).join('');
      }
    }

    async function loadPartners(type) {
      const list = await apiFetch(`/api/masters/partners?partner_type=${type}`);
      const tbBody = document.getElementById('tblBody');
      tbBody.innerHTML = list.map(p => `
        <tr>
          <td><strong>${p.code}</strong></td>
          <td><strong>${p.name}</strong></td>
          <td><span class="badge badge-gray">${p.partner_type}</span></td>
          <td>${p.biz_no||'-'}</td>
          <td>${p.rep_name||'-'} / ${p.contact_person||'-'}</td>
          <td>${p.phone||'-'}</td>
          <td style="color:#1d4ed8; font-weight:700;">${p.credit_limit.toLocaleString()} 원</td>
          <td>${p.payment_term}</td>
          <td>
            <button class="btn ${p.is_credit_locked ? 'btn-outline' : 'btn-primary'}" style="padding:2px 6px; font-size:11px; ${p.is_credit_locked?'border-color:red; color:red;':''}" onclick="toggleLock(${p.id})">
              ${p.is_credit_locked ? '🚨 여신 락(차단)' : '정상'}
            </button>
          </td>
        </tr>
      `).join('') || '<tr><td colspan="9" style="text-align:center;">거래처 내역이 없습니다.</td></tr>';
    }

    function filterPartner(val) { loadPartners(val); }

    async function toggleLock(id) {
      const res = await apiFetch(`/api/masters/partners/${id}/toggle-lock`, { method: 'PATCH' });
      alert(res.message);
      renderView();
    }

    function openCreateModal(type) {
      const modal = document.getElementById('masterModal');
      const title = document.getElementById('modalTitle');
      const body = document.getElementById('modalBody');

      if (type === 'PARTNER') {
        title.innerText = '신규 거래처 및 여신/결제조건 등록';
        body.innerHTML = `
          <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
            <div class="form-group"><label>거래처 코드</label><input type="text" id="p_code" class="form-control" placeholder="P-NEW" /></div>
            <div class="form-group"><label>거래처 구분</label>
              <select id="p_type" class="form-control">
                <option value="매출처">매출처</option>
                <option value="매입처">매입처</option>
                <option value="매입매출처">매입매출처</option>
                <option value="창고">보관창고</option>
                <option value="파이낸스사">파이낸스사</option>
                <option value="운송주선사">운송주선사</option>
              </select>
            </div>
          </div>
          <div class="form-group"><label>상호(거래처명)</label><input type="text" id="p_name" class="form-control" placeholder="예: (주)플랜비에프에스" /></div>
          <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
            <div class="form-group"><label>사업자번호</label><input type="text" id="p_biz" class="form-control" placeholder="000-00-00000" /></div>
            <div class="form-group"><label>담당자명 / 연락처</label><input type="text" id="p_contact" class="form-control" placeholder="홍길동 과장 / 010-0000-0000" /></div>
            <div class="form-group"><label>여신한도 (원)</label><input type="number" id="p_credit" class="form-control" value="50000000" /></div>
            <div class="form-group"><label>기본 결제조건</label>
              <select id="p_term" class="form-control">
                <option value="외상 30일 회전">외상 30일 회전</option>
                <option value="외상 15일 회전">외상 15일 회전</option>
                <option value="선입금 확인후 상차">선입금 확인후 상차</option>
                <option value="전자어음 60일">전자어음 60일</option>
              </select>
            </div>
          </div>
        `;
      } else if (type === 'COMPANY') {
        title.innerText = '당사 법인 마스터 등록';
        body.innerHTML = `
          <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
            <div class="form-group"><label>법인 코드</label><input type="text" id="c_code" class="form-control" placeholder="SEOUL" /></div>
            <div class="form-group"><label>법인명</label><input type="text" id="c_name" class="form-control" placeholder="(주)서울웰푸드" /></div>
            <div class="form-group"><label>사업자등록번호</label><input type="text" id="c_biz" class="form-control" placeholder="347-81-03002" /></div>
            <div class="form-group"><label>대표자명</label><input type="text" id="c_rep" class="form-control" placeholder="조용훈" /></div>
          </div>
          <div class="form-group"><label>사업장 주소</label><input type="text" id="c_addr" class="form-control" placeholder="서울특별시 강동구 천중로 39길 19-25" /></div>
        `;
      } else if (type === 'RATE') {
        title.innerText = '창고 요율 및 작업비 마스터 등록';
        body.innerHTML = `
          <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
            <div class="form-group"><label>창고명</label><input type="text" id="r_wh" class="form-control" placeholder="강동2" /></div>
            <div class="form-group"><label>보관 조건</label>
              <select id="r_type" class="form-control">
                <option value="냉동">냉동</option>
                <option value="냉장">냉장</option>
              </select>
            </div>
            <div class="form-group"><label>일할 보관료 (원/kg/일)</label><input type="number" step="0.1" id="r_daily" class="form-control" value="1.5" /></div>
            <div class="form-group"><label>급속 동결비 (원/kg)</label><input type="number" id="r_freeze" class="form-control" value="80.0" /></div>
            <div class="form-group"><label>입고 하차비 (원/kg)</label><input type="number" id="r_in" class="form-control" value="5.0" /></div>
            <div class="form-group"><label>출고 상차비 (원/kg)</label><input type="number" id="r_out" class="form-control" value="5.0" /></div>
          </div>
        `;
      } else if (type === 'ITEM') {
        title.innerText = '품목 / 부위 / EST 마스터 등록';
        body.innerHTML = `
          <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
            <div class="form-group"><label>품목 코드</label><input type="text" id="i_code" class="form-control" placeholder="BF-AMH-170" /></div>
            <div class="form-group"><label>축종</label>
              <select id="i_species" class="form-control">
                <option value="우육">우육</option>
                <option value="돈육">돈육</option>
                <option value="계육">계육</option>
              </select>
            </div>
            <div class="form-group"><label>부위명</label><input type="text" id="i_cut" class="form-control" placeholder="소일반갈비(170)" /></div>
            <div class="form-group"><label>가공장(EST)</label><input type="text" id="i_est" class="form-control" placeholder="244I" /></div>
            <div class="form-group"><label>브랜드</label><input type="text" id="i_brand" class="form-control" placeholder="AMH" /></div>
            <div class="form-group"><label>기본 등급</label><input type="text" id="i_grade" class="form-control" value="GF" /></div>
          </div>
        `;
      }
      modal.style.display = 'flex';
    }

    function closeModal() { document.getElementById('masterModal').style.display = 'none'; }

    async function submitModal() {
      const title = document.getElementById('modalTitle').innerText;
      try {
        if (title.includes('거래처')) {
          const payload = {
            code: document.getElementById('p_code').value,
            name: document.getElementById('p_name').value,
            partner_type: document.getElementById('p_type').value,
            biz_no: document.getElementById('p_biz').value,
            contact_person: document.getElementById('p_contact').value,
            credit_limit: parseFloat(document.getElementById('p_credit').value),
            payment_term: document.getElementById('p_term').value
          };
          const res = await apiFetch('/api/masters/partners', {
            method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)
          });
          alert(res.message);
        } else if (title.includes('당사 법인')) {
          const payload = {
            code: document.getElementById('c_code').value,
            name: document.getElementById('c_name').value,
            biz_no: document.getElementById('c_biz').value,
            rep_name: document.getElementById('c_rep').value,
            address: document.getElementById('c_addr').value
          };
          const res = await apiFetch('/api/masters/companies', {
            method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)
          });
          alert(res.message);
        } else if (title.includes('창고 요율')) {
          const payload = {
            warehouse_name: document.getElementById('r_wh').value,
            storage_type: document.getElementById('r_type').value,
            daily_rate_kg: parseFloat(document.getElementById('r_daily').value),
            freezing_fee_kg: parseFloat(document.getElementById('r_freeze').value),
            handling_in_kg: parseFloat(document.getElementById('r_in').value),
            handling_out_kg: parseFloat(document.getElementById('r_out').value)
          };
          const res = await apiFetch('/api/masters/warehouse-rates', {
            method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)
          });
          alert(res.message);
        } else if (title.includes('품목')) {
          const payload = {
            item_code: document.getElementById('i_code').value,
            species: document.getElementById('i_species').value,
            cut_name: document.getElementById('i_cut').value,
            est_no: document.getElementById('i_est').value,
            brand: document.getElementById('i_brand').value,
            default_grade: document.getElementById('i_grade').value
          };
          const res = await apiFetch('/api/masters/item-cuts', {
            method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)
          });
          alert(res.message);
        }
        closeModal();
        renderView();
      } catch (err) {
        alert(err.message);
      }
    }

    renderView();
  </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def serve_master_ui():
    return HTML_PAGE

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))

    uvicorn.run("module1_master:app", host="0.0.0.0", port=port, reload=True)
    """
================================================================================
(주)서울웰푸드 미트플로우(MeatFlow) ERP - [모듈 2. 수입/무역 & 디머리지/창고 파이프라인]
- 전제조건: 모듈 1 기준정보(법인, 6대 거래처, 창고별 냉동/냉장 일할요율, 품목규격) 가동
- 기본 법인: (주)서울웰푸드 (대표이사: 조용훈 / 사업자번호: 347-81-03002)
- 핵심 기능:
  1. Deal No. 오퍼 계약 & L/C, USANCE 유전스 이자 원가 자동 가산
  2. 4대 핵심 일정 (ETD 선적, ETA 부산입항, 통관예정, 내륙창고 입고예정)
  3. 선사 Free Time 디머리지(체선료) 카운트다운 (D-3, D-1 긴급반출) & 실시간 손실액 계산
  4. 선사 D/O 승인 & 내륙 보관창고(강동2, 삼진1 등) 보세운송 파이프라인
  5. 선적중(ON_WATER) 판매예정재고 및 입항 전 '실재고 무차감' HOLD 선점 예약
- 실행: uvicorn module2_trade:app --host 0.0.0.0 --port 8000 --reload
================================================================================
"""

import os
from datetime import datetime, date, timedelta
from typing import List, Optional, Dict, Any

from fastapi import FastAPI, HTTPException, Depends, status, Query
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from sqlalchemy import (
    Column, Integer, Float, String, DateTime, create_engine, desc, text
)
from sqlalchemy.orm import declarative_base, sessionmaker, Session

# -----------------------------------------------------------------------------
# 1. 데이터베이스 엔진 설정 (모듈 1 기준정보 연계)
# -----------------------------------------------------------------------------
DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///./meatflow_core.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

connect_args = {"check_same_thread": False} if "sqlite" in DATABASE_URL else {}
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
# 2. SQLAlchemy ORM 데이터 모델 (모듈 1 마스터 참조 및 모듈 2 무역 테이블)
# -----------------------------------------------------------------------------
class CompanyMaster(Base):
    """모듈 1 연계: 당사 법인 마스터"""
    __tablename__ = "companies"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String(50), unique=True, nullable=False)
    name = Column(String(100), nullable=False)
    biz_no = Column(String(30), nullable=False)
    rep_name = Column(String(50), nullable=False)
    address = Column(String(200), nullable=False)
    phone = Column(String(30), nullable=True)
    fax = Column(String(30), nullable=True)

class PartnerMaster(Base):
    """모듈 1 연계: 6대 거래처 (공급사, 선사, 창고, 운송사 등)"""
    __tablename__ = "partners"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String(50), unique=True, nullable=False)
    name = Column(String(100), unique=True, nullable=False)
    partner_type = Column(String(30), nullable=False)
    credit_limit = Column(Float, default=0.0)
    payment_term = Column(String(50), default="외상 30일 회전")

class WarehouseRateMaster(Base):
    """모듈 1 연계: 창고별 일할 보관료 및 상하차/동결 요율"""
    __tablename__ = "warehouse_rates"
    id = Column(Integer, primary_key=True, index=True)
    warehouse_name = Column(String(100), nullable=False)
    storage_type = Column(String(20), nullable=False)
    daily_rate_kg = Column(Float, default=1.5) # 양편넣기 기본 요율
    freezing_fee_kg = Column(Float, default=80.0)
    handling_in_kg = Column(Float, default=5.0)
    handling_out_kg = Column(Float, default=5.0)

class TradeDeal(Base):
    """2-1. Deal No. 중심의 수입 오퍼 계약, 외환 결제 및 선적 파이프라인"""
    __tablename__ = "trade_deals"
    id = Column(Integer, primary_key=True, index=True)
    deal_no = Column(String(50), unique=True, nullable=False, index=True)
    company_name = Column(String(100), default="(주)서울웰푸드") # 화주 법인
    supplier = Column(String(100), nullable=False) # 해외 공급사 (AMH, JBS, IBP 등)
    carrier_line = Column(String(50), default="OOCL") # OOCL, ONE, MAERSK, HMM, MSC 등
    bl_no = Column(String(50), unique=True, nullable=False, index=True)
    container_no = Column(String(50), nullable=True)
    seal_no = Column(String(50), nullable=True)
    
    # 품목 및 규격 정보 (우육/돈육 표준 명칭 적용)
    species = Column(String(30), default="우육")
    cut_name = Column(String(100), nullable=False) # 소일반갈비(170), 소안심 등
    brand = Column(String(50), nullable=False) # AMH, 킬코이 등
    grade = Column(String(30), default="GF")
    est_no = Column(String(50), default="244I")
    box_qty = Column(Integer, nullable=False)
    weight_kg = Column(Float, nullable=False)
    contract_unit_price = Column(Float, nullable=False) # 오퍼 계약 매입단가 (원/kg)
    adjusted_cost_per_kg = Column(Float, nullable=False) # 유전스이자 가산 수입원가
    
    # 외환 결제 관리
    pay_type = Column(String(30), default="USANCE") # L/C, USANCE, T/T_PRE, T/T_POST
    lc_no = Column(String(50), nullable=True)
    due_date = Column(String(20), nullable=True) # 외환 결제 만기일
    usance_interest_rate = Column(Float, default=5.5) # 연이율 %
    is_tt_receipt_attached = Column(String(10), default="Y")
    
    # 4대 핵심 일정 (ETD, ETA 부산, 통관예정, 내륙창고 입고예정)
    etd_date = Column(String(20), nullable=False) # 선적일
    eta_busan_date = Column(String(20), nullable=False) # 부산항 입항예정일
    customs_est_date = Column(String(20), nullable=True) # 통관예정일
    eta_warehouse_date = Column(String(20), nullable=True) # 내륙창고 입고예정일
    discharge_date = Column(String(20), nullable=True) # 실제 양하일자
    
    # 선사 Free Time 및 체선료(디머리지)
    free_time_days = Column(Integer, default=10) # 선사 제공 무료보관일 (7~10일)
    demurrage_daily_cost_usd = Column(Float, default=180.0) # 일일 체선료 과징금 ($180)
    
    # 선사 D/O(화물인도지시서) & 내륙 보관창고
    do_status = Column(String(30), default="DO_WAITING") # DO_WAITING, DO_ISSUED, DO_HOLD
    do_number = Column(String(50), nullable=True)
    do_valid_date = Column(String(20), nullable=True)
    bonded_wh = Column(String(50), default="강동2") # 반입 내륙 보관창고 (강동2, 삼진1 등)
    
    # 진행 상태: ON_WATER(선적중), PORT_IN(입항양하), CUSTOMS_CLEAR(통관수리), WAREHOUSE_IN(창고입고)
    status = Column(String(30), default="ON_WATER", index=True)
    created_at = Column(DateTime, default=datetime.now)

class TradePreSaleReservation(Base):
    """2-2. 선적 판매예정재고 입항 전 '실재고 무차감' 선점 예약(HOLD) 원장"""
    __tablename__ = "trade_presale_reservations"
    id = Column(Integer, primary_key=True, index=True)
    res_no = Column(String(50), unique=True, nullable=False, index=True)
    deal_no = Column(String(50), nullable=False, index=True)
    bl_no = Column(String(50), nullable=False)
    sales_rep = Column(String(50), default="조용훈 대표")
    customer = Column(String(100), nullable=False) # 예약 고객사 (예: (주)플랜비에프에스, 에스티에프)
    reserved_box_qty = Column(Integer, nullable=False)
    reserved_weight_kg = Column(Float, nullable=False)
    target_unit_price = Column(Float, nullable=False)
    total_amount = Column(Float, nullable=False)
    status = Column(String(20), default="HOLD") # HOLD, CONVERTED_TO_STOCK, CANCELLED
    created_at = Column(DateTime, default=datetime.now)

Base.metadata.create_all(bind=engine)

# -----------------------------------------------------------------------------
# 3. 비즈니스 로직 연산 엔진 (디머리지, 유전스이자, 가용선적재고)
# -----------------------------------------------------------------------------
def calculate_usance_interest(price: float, rate_pct: float, etd_str: str, due_str: str) -> float:
    """유전스(USANCE) 일할 외화이자 계산 후 수입원가에 가산"""
    try:
        d1 = datetime.strptime(etd_str[:10], "%Y-%m-%d").date()
        d2 = datetime.strptime(due_str[:10], "%Y-%m-%d").date()
        days = max(1, (d2 - d1).days)
        daily_rate = (rate_pct / 100.0) / 365.0
        interest_cost = price * daily_rate * days
        return round(price + interest_cost, 1)
    except Exception:
        return price

def calculate_demurrage_info(deal: TradeDeal) -> Dict[str, Any]:
    """선사 Free Time 및 디머리지(체선료) 초과 카운트다운 & 손실액 연산"""
    if not deal.discharge_date or deal.status == "ON_WATER":
        return {
            "status_code": "ON_WATER",
            "d_day_text": "선적 항해중",
            "badge_class": "badge-blue",
            "overdue_days": 0,
            "demurrage_loss_krw": 0
        }
    
    try:
        dis_d = datetime.strptime(deal.discharge_date[:10], "%Y-%m-%d").date()
        today = date.today()
        elapsed_days = (today - dis_d).days + 1 # 양하 당일 포함
        remaining_free_days = deal.free_time_days - elapsed_days
        
        if remaining_free_days > 3:
            return {
                "status_code": "FREE_TIME_SAFE",
                "d_day_text": f"Free Time 잔여 {remaining_free_days}일",
                "badge_class": "badge-green",
                "overdue_days": 0,
                "demurrage_loss_krw": 0
            }
        elif 1 <= remaining_free_days <= 3:
            return {
                "status_code": "DEMURRAGE_IMMINENT",
                "d_day_text": f"⚠️ 디머리지 임박 D-{remaining_free_days}",
                "badge_class": "badge-orange",
                "overdue_days": 0,
                "demurrage_loss_krw": 0
            }
        elif remaining_free_days == 0:
            return {
                "status_code": "DEMURRAGE_URGENT",
                "d_day_text": "🚨 긴급 반출 요망 D-Day",
                "badge_class": "badge-red",
                "overdue_days": 0,
                "demurrage_loss_krw": 0
            }
        else:
            overdue = abs(remaining_free_days)
            # 환율 1,350원 기준 일일 과징금 실시간 계산
            loss_krw = round(overdue * deal.demurrage_daily_cost_usd * 1350)
            return {
                "status_code": "DEMURRAGE_PENALTY",
                "d_day_text": f"🚨 체선료 발생 ({overdue}일 초과)",
                "badge_class": "badge-danger-blink",
                "overdue_days": overdue,
                "demurrage_loss_krw": loss_krw
            }
    except Exception:
        return {
            "status_code": "UNKNOWN",
            "d_day_text": "-",
            "badge_class": "badge-gray",
            "overdue_days": 0,
            "demurrage_loss_krw": 0
        }

# -----------------------------------------------------------------------------
# 4. 초기 기준정보 및 실무 테스트 데이터 시딩
# -----------------------------------------------------------------------------
def init_core_data():
    db = SessionLocal()
    try:
        # 모듈 1 연계 기준정보
        if not db.query(CompanyMaster).first():
            comps = [
                CompanyMaster(code="SEOUL", name="(주)서울웰푸드", biz_no="347-81-03002", rep_name="조용훈", address="서울특별시 강동구 천중로 39길 19-25", phone="02-6958-9229", fax="070-4758-9219"),
                CompanyMaster(code="TJF", name="주식회사 티제이에프", biz_no="718-88-03523", rep_name="김다희", address="경기도 남양주시 별내중앙로 34", phone="010-5773-0619", fax="070-4758-9219"),
                CompanyMaster(code="NEXUS", name="(주)넥서스트레이딩", biz_no="518-86-03633", rep_name="김다희", address="서울특별시 강동구 천중로 39길 19-25", phone="02-2299-1111", fax="02-2299-1112")
            ]
            db.add_all(comps)

        if not db.query(WarehouseRateMaster).first():
            rates = [
                WarehouseRateMaster(warehouse_name="강동2", storage_type="냉동", daily_rate_kg=1.5, freezing_fee_kg=80.0),
                WarehouseRateMaster(warehouse_name="강동2", storage_type="냉장", daily_rate_kg=2.0, freezing_fee_kg=80.0),
                WarehouseRateMaster(warehouse_name="삼진1", storage_type="냉동", daily_rate_kg=1.5, freezing_fee_kg=80.0),
                WarehouseRateMaster(warehouse_name="삼진1", storage_type="냉장", daily_rate_kg=2.0, freezing_fee_kg=80.0),
                WarehouseRateMaster(warehouse_name="광주냉장", storage_type="냉동", daily_rate_kg=1.4, freezing_fee_kg=80.0)
            ]
            db.add_all(rates)

        if not db.query(TradeDeal).first():
            # Deal 1: 부산 입항 후 Free Time D-2 임박 건 (강동2 창고 반입 예정)
            deal1 = TradeDeal(
                deal_no="DEAL-2026-088", company_name="(주)서울웰푸드", supplier="AMH AUSTRALIA", carrier_line="OOCL",
                bl_no="OOLU2766442080", container_no="OOLU9918201", seal_no="SL-99201",
                species="우육", cut_name="소일반갈비(170)", brand="AMH", grade="GF", est_no="244I",
                box_qty=541, weight_kg=11000.0, contract_unit_price=14500.0, adjusted_cost_per_kg=14620.0,
                pay_type="USANCE", lc_no="LC-2607-0091", due_date="2026-09-20", usance_interest_rate=5.5,
                etd_date="2026-07-15", eta_busan_date="2026-08-16", customs_est_date="2026-08-26",
                eta_warehouse_date="2026-08-27", discharge_date="2026-08-17",
                free_time_days=10, demurrage_daily_cost_usd=180.0,
                do_status="DO_ISSUED", do_number="DO-OOCL-20260818-09", do_valid_date="2026-08-28",
                bonded_wh="강동2", status="PORT_IN"
            )

            # Deal 2: 해상 항해 중 판매예정 가용재고 건 (삼진1 창고 반입 예정)
            deal2 = TradeDeal(
                deal_no="DEAL-2026-092", company_name="(주)서울웰푸드", supplier="JBS USA", carrier_line="ONE",
                bl_no="ONEY9940128840", container_no="ONEU8810294", seal_no="SL-88192",
                species="우육", cut_name="소안심", brand="AMH", grade="PRIME", est_no="244I",
                box_qty=200, weight_kg=4000.0, contract_unit_price=19500.0, adjusted_cost_per_kg=19680.0,
                pay_type="L/C", lc_no="LC-2608-0112", due_date="2026-10-10", usance_interest_rate=5.0,
                etd_date="2026-08-10", eta_busan_date="2026-08-28", customs_est_date="2026-09-02",
                eta_warehouse_date="2026-09-04", discharge_date=None,
                free_time_days=10, demurrage_daily_cost_usd=200.0,
                do_status="DO_WAITING", do_number=None, do_valid_date=None,
                bonded_wh="삼진1", status="ON_WATER"
            )

            # Deal 3: 체선료 초과 발생 건 (삼진1 창고 반입 예정)
            deal3 = TradeDeal(
                deal_no="DEAL-2026-075", company_name="주식회사 티제이에프", supplier="KILCOY AUSTRALIA", carrier_line="MAERSK",
                bl_no="MAEU8819201102", container_no="MAEU1029910", seal_no="SL-10291",
                species="돈육", cut_name="돈갈매기살", brand="킬코이", grade="GF", est_no="1614",
                box_qty=300, weight_kg=6000.0, contract_unit_price=8500.0, adjusted_cost_per_kg=8580.0,
                pay_type="T/T_POST", due_date="2026-08-30", usance_interest_rate=0.0,
                etd_date="2026-07-10", eta_busan_date="2026-08-05", customs_est_date="2026-08-25",
                eta_warehouse_date="2026-08-26", discharge_date="2026-08-06",
                free_time_days=10, demurrage_daily_cost_usd=180.0,
                do_status="DO_HOLD", do_number=None, do_valid_date=None,
                bonded_wh="삼진1", status="PORT_IN"
            )

            db.add_all([deal1, deal2, deal3])
            db.flush()

            # Deal 2(선적중 재고)에 대한 입항 전 무차감 예약 1건 시딩 (50 Box 예약)
            res1 = TradePreSaleReservation(
                res_no="PRESALE-260824-001", deal_no="DEAL-2026-092", bl_no="ONEY9940128840",
                sales_rep="조용훈 대표", customer="(주)플랜비에프에스",
                reserved_box_qty=50, reserved_weight_kg=1000.0,
                target_unit_price=20500.0, total_amount=20500000.0, status="HOLD"
            )
            db.add(res1)
            db.commit()
    finally:
        db.close()

init_core_data()

# -----------------------------------------------------------------------------
# 5. FastAPI Pydantic DTO 스키마
# -----------------------------------------------------------------------------
app = FastAPI(title="(주)서울웰푸드 미트플로우(MeatFlow) ERP - 모듈 2. 수입/무역관리", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class TradeDealCreateReq(BaseModel):
    deal_no: str
    company_name: str = "(주)서울웰푸드"
    supplier: str
    carrier_line: str = "OOCL"
    bl_no: str
    container_no: Optional[str] = ""
    seal_no: Optional[str] = ""
    species: str = "우육"
    cut_name: str
    brand: str
    grade: str = "GF"
    est_no: str = "244I"
    box_qty: int
    weight_kg: float
    contract_unit_price: float
    pay_type: str = "USANCE"
    lc_no: Optional[str] = None
    due_date: Optional[str] = None
    usance_interest_rate: float = 5.5
    etd_date: str
    eta_busan_date: str
    customs_est_date: Optional[str] = None
    eta_warehouse_date: Optional[str] = None
    free_time_days: int = 10
    bonded_wh: str = "강동2"

class PreSaleHoldReq(BaseModel):
    deal_no: str
    customer: str
    reserved_box_qty: int
    target_unit_price: float
    sales_rep: str = "조용훈 대표"

class DOApproveReq(BaseModel):
    do_number: str
    do_valid_date: str

# -----------------------------------------------------------------------------
# 6. REST API 엔드포인트
# -----------------------------------------------------------------------------
@app.get("/api/trades/summary")
def get_trade_summary(db: Session = Depends(get_db)):
    deals = db.query(TradeDeal).all()
    on_water_boxes = sum(d.box_qty for d in deals if d.status == "ON_WATER")
    on_water_weight = sum(d.weight_kg for d in deals if d.status == "ON_WATER")
    
    total_demurrage_loss = 0
    imminent_count = 0
    overdue_count = 0
    
    for d in deals:
        info = calculate_demurrage_info(d)
        total_demurrage_loss += info["demurrage_loss_krw"]
        if info["status_code"] in ["DEMURRAGE_IMMINENT", "DEMURRAGE_URGENT"]:
            imminent_count += 1
        elif info["status_code"] == "DEMURRAGE_PENALTY":
            overdue_count += 1
            
    return {
        "on_water_boxes": on_water_boxes,
        "on_water_weight": round(on_water_weight, 1),
        "imminent_count": imminent_count,
        "overdue_count": overdue_count,
        "total_demurrage_loss_krw": total_demurrage_loss
    }

@app.get("/api/trades/warehouses")
def get_bonded_warehouses(db: Session = Depends(get_db)):
    """모듈 1에 등록된 보관창고 목록 호출"""
    return db.query(WarehouseRateMaster).all()

@app.get("/api/trades/deals")
def get_trade_deals(status_filter: Optional[str] = None, db: Session = Depends(get_db)):
    q = db.query(TradeDeal)
    if status_filter and status_filter != "ALL":
        q = q.filter(TradeDeal.status == status_filter)
    deals = q.order_by(desc(TradeDeal.id)).all()
    
    presales = db.query(TradePreSaleReservation).filter(TradePreSaleReservation.status == "HOLD").all()
    res_map = {}
    for p in presales:
        res_map[p.deal_no] = res_map.get(p.deal_no, 0) + p.reserved_box_qty
        
    result = []
    for d in deals:
        dem_info = calculate_demurrage_info(d)
        reserved_box = res_map.get(d.deal_no, 0)
        avail_presale_box = max(0, d.box_qty - reserved_box)
        
        result.append({
            "id": d.id,
            "deal_no": d.deal_no,
            "company_name": d.company_name,
            "supplier": d.supplier,
            "carrier_line": d.carrier_line,
            "bl_no": d.bl_no,
            "container_no": d.container_no or "-",
            "species": d.species,
            "cut_name": d.cut_name,
            "brand": d.brand,
            "grade": d.grade,
            "est_no": d.est_no,
            "box_qty": d.box_qty,
            "weight_kg": d.weight_kg,
            "contract_unit_price": d.contract_unit_price,
            "adjusted_cost_per_kg": d.adjusted_cost_per_kg,
            "pay_type": d.pay_type,
            "due_date": d.due_date or "-",
            "etd_date": d.etd_date,
            "eta_busan_date": d.eta_busan_date,
            "customs_est_date": d.customs_est_date or "-",
            "eta_warehouse_date": d.eta_warehouse_date or "-",
            "discharge_date": d.discharge_date or "-",
            "free_time_days": d.free_time_days,
            "demurrage_info": dem_info,
            "do_status": d.do_status,
            "do_number": d.do_number or "-",
            "bonded_wh": d.bonded_wh, # 반입 보관창고
            "status": d.status,
            "reserved_box_qty": reserved_box,
            "avail_presale_box": avail_presale_box
        })
    return result

@app.post("/api/trades/deals")
def create_trade_deal(req: TradeDealCreateReq, db: Session = Depends(get_db)):
    if db.query(TradeDeal).filter(TradeDeal.deal_no == req.deal_no).first():
        raise HTTPException(status_code=400, detail="이미 존재하는 Deal No.입니다.")
    if db.query(TradeDeal).filter(TradeDeal.bl_no == req.bl_no).first():
        raise HTTPException(status_code=400, detail="이미 등록된 B/L 번호입니다.")
        
    adjusted_cost = req.contract_unit_price
    if req.pay_type == "USANCE" and req.due_date:
        adjusted_cost = calculate_usance_interest(
            req.contract_unit_price, req.usance_interest_rate, req.etd_date, req.due_date
        )
        
    deal = TradeDeal(
        deal_no=req.deal_no, company_name=req.company_name, supplier=req.supplier,
        carrier_line=req.carrier_line, bl_no=req.bl_no, container_no=req.container_no,
        seal_no=req.seal_no, species=req.species, cut_name=req.cut_name, brand=req.brand,
        grade=req.grade, est_no=req.est_no, box_qty=req.box_qty, weight_kg=req.weight_kg,
        contract_unit_price=req.contract_unit_price, adjusted_cost_per_kg=adjusted_cost,
        pay_type=req.pay_type, lc_no=req.lc_no, due_date=req.due_date,
        usance_interest_rate=req.usance_interest_rate, etd_date=req.etd_date,
        eta_busan_date=req.eta_busan_date, customs_est_date=req.customs_est_date,
        eta_warehouse_date=req.eta_warehouse_date, free_time_days=req.free_time_days,
        bonded_wh=req.bonded_wh, status="ON_WATER"
    )
    db.add(deal)
    db.commit()
    return {"message": f"수입 오퍼 계약 '{req.deal_no}' 등록 완료 (보관창고: {req.bonded_wh} 지정)"}

@app.post("/api/trades/presale-hold")
def create_presale_hold(req: PreSaleHoldReq, db: Session = Depends(get_db)):
    deal = db.query(TradeDeal).filter(TradeDeal.deal_no == req.deal_no).first()
    if not deal:
        raise HTTPException(status_code=404, detail="해당 수입 계약을 찾을 수 없습니다.")
        
    curr_res = db.query(TradePreSaleReservation).filter(
        TradePreSaleReservation.deal_no == req.deal_no,
        TradePreSaleReservation.status == "HOLD"
    ).all()
    total_res_box = sum(r.reserved_box_qty for r in curr_res)
    avail_box = deal.box_qty - total_res_box
    
    if req.reserved_box_qty > avail_box:
        raise HTTPException(status_code=400, detail=f"판매예정 가용수량({avail_box} Box)을 초과했습니다.")
        
    avg_weight = deal.weight_kg / deal.box_qty if deal.box_qty > 0 else 20.0
    res_w = round(avg_weight * req.reserved_box_qty, 2)
    tot_amt = round(res_w * req.target_unit_price)
    res_no = f"PRESALE-{datetime.now().strftime('%y%m%d%H%M%S')}"
    
    presale = TradePreSaleReservation(
        res_no=res_no, deal_no=deal.deal_no, bl_no=deal.bl_no,
        sales_rep=req.sales_rep, customer=req.customer,
        reserved_box_qty=req.reserved_box_qty, reserved_weight_kg=res_w,
        target_unit_price=req.target_unit_price, total_amount=tot_amt, status="HOLD"
    )
    db.add(presale)
    db.commit()
    return {"message": f"입항 전 무차감 예약(HOLD) 등록 완료 ({res_no}: {req.reserved_box_qty} Box 선점)"}

@app.post("/api/trades/deals/{deal_no}/approve-do")
def approve_do(deal_no: str, req: DOApproveReq, db: Session = Depends(get_db)):
    deal = db.query(TradeDeal).filter(TradeDeal.deal_no == deal_no).first()
    if not deal:
        raise HTTPException(status_code=404, detail="수입 계약을 찾을 수 없습니다.")
        
    deal.do_status = "DO_ISSUED"
    deal.do_number = req.do_number
    deal.do_valid_date = req.do_valid_date
    db.commit()
    return {"message": f"선사 D/O 발급 승인 완료 (D/O: {req.do_number}, {deal.bonded_wh} 창고 보세운송 가능)"}

@app.post("/api/trades/deals/{deal_no}/discharge")
def register_discharge(deal_no: str, discharge_date_str: str = Query(...), db: Session = Depends(get_db)):
    deal = db.query(TradeDeal).filter(TradeDeal.deal_no == deal_no).first()
    if not deal:
        raise HTTPException(status_code=404, detail="수입 계약을 찾을 수 없습니다.")
        
    deal.discharge_date = discharge_date_str
    deal.status = "PORT_IN"
    db.commit()
    return {"message": f"부산항 양하 등록 완료 ({discharge_date_str}), 선사 Free Time 카운트다운 가동"}

# -----------------------------------------------------------------------------
# 7. 모듈 2 SPA 프론트엔드 UI
# -----------------------------------------------------------------------------
HTML_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>(주)서울웰푸드 미트플로우 ERP - 2. 무역선적 및 디머리지/창고 관리</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
  <style>
    :root {
      --primary: #1e3a8a; --primary-hover: #1e40af; --sidebar: #0f172a; --bg: #f8fafc;
      --border: #cbd5e1; --text: #0f172a; --muted: #64748b; --danger: #dc2626; --warning: #d97706; --success: #16a34a;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Noto Sans KR", sans-serif; }
    body { background: var(--bg); color: var(--text); display: flex; height: 100vh; overflow: hidden; }

    aside { width: 250px; background: var(--sidebar); color: #fff; display: flex; flex-direction: column; flex-shrink: 0; }
    .brand-box { padding: 18px 16px; border-bottom: 1px solid #1e293b; }
    .brand-title { font-size: 15px; font-weight: 800; color: #38bdf8; display: flex; align-items: center; gap: 8px; }
    .brand-sub { font-size: 10.5px; color: #94a3b8; margin-top: 4px; }
    .nav-list { list-style: none; padding: 12px 8px; flex: 1; overflow-y: auto; }
    .nav-header { font-size: 10px; font-weight: 700; color: #64748b; padding: 8px 10px 4px; text-transform: uppercase; }
    .nav-item { display: flex; align-items: center; gap: 10px; padding: 9px 12px; border-radius: 6px; font-size: 12.5px; font-weight: 600; color: #cbd5e1; cursor: pointer; margin-bottom: 2px; }
    .nav-item:hover, .nav-item.active { background: var(--primary); color: #fff; }

    main { flex: 1; display: flex; flex-direction: column; overflow: hidden; }
    header { background: #fff; border-bottom: 1px solid var(--border); padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; }
    .content-area { flex: 1; padding: 18px 24px; overflow-y: auto; }

    .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 16px; }
    .kpi-card { background: #fff; padding: 14px 16px; border-radius: 8px; border: 1px solid var(--border); box-shadow: 0 1px 2px rgba(0,0,0,0.03); }
    .kpi-title { font-size: 11.5px; font-weight: 600; color: var(--muted); margin-bottom: 4px; }
    .kpi-val { font-size: 18px; font-weight: 800; color: var(--text); }

    .card { background: #fff; border-radius: 8px; border: 1px solid var(--border); overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.02); }
    .toolbar { padding: 12px 16px; background: #f8fafc; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; text-align: left; }
    th { background: #f1f5f9; padding: 9px 12px; font-weight: 700; color: #475569; border-bottom: 1px solid var(--border); white-space: nowrap; }
    td { padding: 9px 12px; border-bottom: 1px solid var(--border); vertical-align: middle; }
    tr:hover { background: #f8fafc; }

    .badge { padding: 3px 7px; border-radius: 4px; font-size: 11px; font-weight: 700; display: inline-block; }
    .badge-blue { background: #dbeafe; color: #1e40af; }
    .badge-green { background: #dcfce7; color: #166534; }
    .badge-orange { background: #ffedd5; color: #c2410c; }
    .badge-red { background: #fee2e2; color: #991b1b; }
    .badge-gray { background: #f1f5f9; color: #475569; }
    .badge-danger-blink { background: #dc2626; color: #fff; animation: blink 1.2s infinite; }
    @keyframes blink { 0% { opacity: 1; } 50% { opacity: 0.4; } 100% { opacity: 1; } }

    .btn { padding: 7px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; border: none; display: inline-flex; align-items: center; gap: 4px; }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-outline { background: #fff; border: 1px solid var(--border); color: var(--text); }

    .modal-overlay { position: fixed; inset: 0; background: rgba(15,23,42,0.6); display: none; align-items: center; justify-content: center; z-index: 1000; }
    .modal-box { background: #fff; width: 780px; max-width: 95%; border-radius: 10px; overflow: hidden; }
    .modal-header { padding: 12px 18px; background: #0f172a; color: #fff; font-weight: 700; font-size: 13px; display: flex; justify-content: space-between; align-items: center; }
    .modal-body { padding: 18px; max-height: 75vh; overflow-y: auto; }
    .form-group { margin-bottom: 12px; }
    .form-group label { display: block; font-size: 11.5px; font-weight: 700; color: #334155; margin-bottom: 4px; }
    .form-control { width: 100%; padding: 7px 10px; border: 1px solid var(--border); border-radius: 5px; font-size: 12px; }
    .modal-footer { padding: 10px 18px; background: #f8fafc; border-top: 1px solid var(--border); display: flex; justify-content: flex-end; gap: 6px; }
  </style>
</head>
<body>
  <aside>
    <div class="brand-box">
      <div class="brand-title"><i class="bi bi-box-seam-fill"></i> 미트플로우 ERP</div>
      <div class="brand-sub">(주)서울웰푸드 맞춤형</div>
    </div>
    <ul class="nav-list">
      <div class="nav-header">기준정보</div>
      <li class="nav-item" onclick="alert('모듈 1에서 제공됩니다.')"><i class="bi bi-gear-fill"></i> 1. 기준정보 관리</li>
      <div class="nav-header">수입 / 무역 / 물류 (모듈 2)</div>
      <li class="nav-item active"><i class="bi bi-water"></i> 2. 무역 / 선적일정</li>
      <div class="nav-header">입고 / 검수 / 보세</div>
      <li class="nav-item" onclick="alert('모듈 3에서 제공됩니다.')"><i class="bi bi-shield-check"></i> 3. 보세 / 하차검수</li>
      <div class="nav-header">재고 / 영업 / 정산</div>
      <li class="nav-item" onclick="alert('모듈 4에서 제공됩니다.')"><i class="bi bi-grid-3x3-gap-fill"></i> 4. 재고 / 보관관리</li>
      <li class="nav-item" onclick="alert('모듈 5에서 제공됩니다.')"><i class="bi bi-truck"></i> 5. 영업 / 출고배차</li>
    </ul>
  </aside>

  <main>
    <header>
      <div style="font-size:16px; font-weight:800; color:var(--text);"><i class="bi bi-water"></i> 2. 수입/무역 선적 관리 & 디머리지 방어 엔진 (MeatFlow Trade)</div>
      <div id="headerActions" style="display:flex; gap:8px;"></div>
    </header>

    <div class="content-area">
      <!-- 상단 무역/디머리지 관제 KPI -->
      <div class="kpi-grid">
        <div class="kpi-card">
          <div class="kpi-title">선적 항해중 판매예정재고</div>
          <div class="kpi-val" style="color:var(--primary);" id="kpi_onwater">0 Box</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-title">디머리지 임박 건수 (D-3 ~ D-Day)</div>
          <div class="kpi-val" style="color:var(--warning);" id="kpi_imminent">0 건</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-title">🚨 체선료 초과 발생 건수</div>
          <div class="kpi-val" style="color:var(--danger);" id="kpi_overdue">0 건</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-title">누적 체선료 발생 손실액</div>
          <div class="kpi-val" style="color:var(--danger);" id="kpi_loss">0 원</div>
        </div>
      </div>

      <div class="card">
        <div class="toolbar">
          <strong>🚢 Deal No. 오퍼 계약 & 선사 Free Time 디머리지 실시간 타임라인</strong>
          <div style="display:flex; gap:8px;">
            <select class="form-control" style="width:160px;" onchange="loadDeals(this.value)">
              <option value="ALL">전체 상태 보기</option>
              <option value="ON_WATER">선적 항해중 (ON_WATER)</option>
              <option value="PORT_IN">부산 입항/양하 (PORT_IN)</option>
            </select>
          </div>
        </div>
        <div style="overflow-x:auto;">
          <table>
            <thead>
              <tr>
                <th>Deal No</th><th>화주</th><th>공급사 / 선사</th><th>B/L NO</th><th>품목 / 부위</th>
                <th>선적수량</th><th>판매예정가용</th><th>예약(HOLD)</th><th>외환/결제</th><th>수입원가</th>
                <th>선적일(ETD)</th><th>입항일(ETA)</th><th>반입 보관창고</th><th>디머리지 방어 타임라인</th><th>D/O 상태</th><th>작업</th>
              </tr>
            </thead>
            <tbody id="dealTableBody"></tbody>
          </table>
        </div>
      </div>
    </div>
  </main>

  <!-- 수입 계약 등록 모달 (창고 선택 포함) -->
  <div class="modal-overlay" id="dealModal">
    <div class="modal-box">
      <div class="modal-header">
        <span>신규 수입 계약(Deal No.) 등록 및 판매예정재고 생성</span>
        <i class="bi bi-x-lg" style="cursor:pointer;" onclick="closeModal('dealModal')"></i>
      </div>
      <div class="modal-body">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
          <div class="form-group"><label>Deal No.</label><input type="text" id="d_deal_no" class="form-control" placeholder="DEAL-2026-XXX" /></div>
          <div class="form-group"><label>화주 법인</label>
            <select id="d_comp" class="form-control">
              <option value="(주)서울웰푸드">(주)서울웰푸드</option>
              <option value="주식회사 티제이에프">주식회사 티제이에프</option>
              <option value="(주)넥서스트레이딩">(주)넥서스트레이딩</option>
            </select>
          </div>
          <div class="form-group"><label>해외 공급사</label><input type="text" id="d_supp" class="form-control" placeholder="AMH AUSTRALIA" /></div>
          <div class="form-group"><label>선사 (Shipping Line)</label>
            <select id="d_line" class="form-control">
              <option value="OOCL">OOCL (Free Time 10일)</option>
              <option value="ONE">ONE (Free Time 10일)</option>
              <option value="MAERSK">MAERSK (Free Time 7일)</option>
              <option value="HMM">HMM (Free Time 10일)</option>
              <option value="MSC">MSC (Free Time 8일)</option>
            </select>
          </div>
          <div class="form-group"><label>B/L NO</label><input type="text" id="d_bl" class="form-control" placeholder="OOLUXXXXXXXX" /></div>
          <div class="form-group"><label>컨테이너 번호</label><input type="text" id="d_cntr" class="form-control" placeholder="OOLU1234567" /></div>
        </div>
        <hr style="margin:8px 0; border:0; border-top:1px solid var(--border);" />
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
          <div class="form-group"><label>축종 / 부위명</label><input type="text" id="d_cut" class="form-control" placeholder="우육 소일반갈비(170)" /></div>
          <div class="form-group"><label>브랜드 / 가공장(EST)</label><input type="text" id="d_brand_est" class="form-control" placeholder="AMH / 244I" /></div>
          <div class="form-group"><label>선적 수량 (Box)</label><input type="number" id="d_box" class="form-control" value="500" /></div>
          <div class="form-group"><label>선적 총중량 (kg)</label><input type="number" id="d_weight" class="form-control" value="10000.0" /></div>
          <div class="form-group"><label>계약 매입단가 (원/kg)</label><input type="number" id="d_price" class="form-control" value="14500" /></div>
          <div class="form-group"><label>결제방식</label>
            <select id="d_pay_type" class="form-control">
              <option value="USANCE">USANCE (유전스 외화이자 가산)</option>
              <option value="L/C">L/C (신용장 결제)</option>
              <option value="T/T_PRE">T/T 사전송금 (영수증 필수)</option>
              <option value="T/T_POST">T/T 사후송금</option>
            </select>
          </div>
        </div>
        <hr style="margin:8px 0; border:0; border-top:1px solid var(--border);" />
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
          <div class="form-group"><label>선적일자 (ETD)</label><input type="date" id="d_etd" class="form-control" /></div>
          <div class="form-group"><label>부산 입항예정일 (ETA)</label><input type="date" id="d_eta" class="form-control" /></div>
          <div class="form-group"><label>외환 결제 만기일 (Due Date)</label><input type="date" id="d_due" class="form-control" /></div>
          <div class="form-group"><label>반입 보관창고 (모듈 1 연계)</label>
            <select id="d_wh" class="form-control">
              <option value="강동2">강동2 (냉동 1.5원 / 냉장 2.0원)</option>
              <option value="삼진1">삼진1 (냉동 1.5원 / 냉장 2.0원)</option>
              <option value="광주냉장">광주냉장 (냉동 1.4원)</option>
            </select>
          </div>
        </div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-outline" onclick="closeModal('dealModal')">취소</button>
        <button class="btn btn-primary" onclick="submitTradeDeal()">오퍼 계약 등록</button>
      </div>
    </div>
  </div>

  <!-- 입항 전 무차감 예약(HOLD) 모달 -->
  <div class="modal-overlay" id="presaleModal">
    <div class="modal-box" style="width:520px;">
      <div class="modal-header">
        <span>선적 판매예정재고 입항 전 무차감 예약(HOLD)</span>
        <i class="bi bi-x-lg" style="cursor:pointer;" onclick="closeModal('presaleModal')"></i>
      </div>
      <div class="modal-body">
        <input type="hidden" id="ps_deal_no" />
        <div class="form-group"><label>선택 계약 / B/L</label><input type="text" id="ps_deal_disp" class="form-control" readonly /></div>
        <div class="form-group"><label>예약 고객사(매출처)</label><input type="text" id="ps_customer" class="form-control" placeholder="(주)플랜비에프에스" /></div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
          <div class="form-group"><label>선점 예약 수량 (Box)</label><input type="number" id="ps_box" class="form-control" value="50" /></div>
          <div class="form-group"><label>제안 견적단가 (원/kg)</label><input type="number" id="ps_price" class="form-control" value="15500" /></div>
        </div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-outline" onclick="closeModal('presaleModal')">취소</button>
        <button class="btn btn-primary" onclick="submitPreSaleHold()">입항 전 선점 예약(HOLD)</button>
      </div>
    </div>
  </div>

  <!-- D/O 승인 모달 -->
  <div class="modal-overlay" id="doModal">
    <div class="modal-box" style="width:480px;">
      <div class="modal-header">
        <span>선사 D/O(화물인도지시서) 발급 승인 등록</span>
        <i class="bi bi-x-lg" style="cursor:pointer;" onclick="closeModal('doModal')"></i>
      </div>
      <div class="modal-body">
        <input type="hidden" id="do_deal_no" />
        <div class="form-group"><label>D/O 발급 승인 번호</label><input type="text" id="do_number" class="form-control" placeholder="DO-OOCL-2026XXXX-XX" /></div>
        <div class="form-group"><label>D/O 유효기한</label><input type="date" id="do_valid_date" class="form-control" /></div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-outline" onclick="closeModal('doModal')">취소</button>
        <button class="btn btn-primary" onclick="submitDOApprove()">D/O 승인 완료</button>
      </div>
    </div>
  </div>

  <script>
    let cachedDeals = [];

    async function apiFetch(url, opt={}) {
      const res = await fetch(url, opt);
      if (!res.ok) {
        const err = await res.json().catch(()=>({detail:'오류 발생'}));
        throw new Error(err.detail || '서버 오류');
      }
      return res.json();
    }

    async function initPage() {
      document.getElementById('headerActions').innerHTML = `
        <button class="btn btn-primary" onclick="openDealModal()"><i class="bi bi-plus-lg"></i> 신규 오퍼 계약 등록</button>
      `;
      loadSummary();
      loadDeals('ALL');
    }

    async function loadSummary() {
      const s = await apiFetch('/api/trades/summary');
      document.getElementById('kpi_onwater').innerText = `${s.on_water_boxes.toLocaleString()} Box (${s.on_water_weight.toLocaleString()}kg)`;
      document.getElementById('kpi_imminent').innerText = `${s.imminent_count} 건`;
      document.getElementById('kpi_overdue').innerText = `${s.overdue_count} 건`;
      document.getElementById('kpi_loss').innerText = `${s.total_demurrage_loss_krw.toLocaleString()} 원`;
    }

    async function loadDeals(statusFilter) {
      const list = await apiFetch(`/api/trades/deals?status_filter=${statusFilter}`);
      cachedDeals = list;
      const tb = document.getElementById('dealTableBody');
      tb.innerHTML = list.map(d => `
        <tr>
          <td><strong>${d.deal_no}</strong></td>
          <td><span class="badge badge-blue">${d.company_name}</span></td>
          <td>${d.supplier} / <strong>${d.carrier_line}</strong></td>
          <td><code>${d.bl_no}</code></td>
          <td><strong>${d.species} ${d.cut_name}</strong> (${d.brand})</td>
          <td>${d.box_qty} Box</td>
          <td style="color:#166534; font-weight:800;">${d.avail_presale_box} Box</td>
          <td style="color:#1e40af; font-weight:700;">${d.reserved_box_qty > 0 ? d.reserved_box_qty + ' Box' : '-'}</td>
          <td><span class="badge badge-gray">${d.pay_type}</span><div style="font-size:10px; color:#64748b;">만기: ${d.due_date}</div></td>
          <td>${d.adjusted_cost_per_kg.toLocaleString()}원</td>
          <td>${d.etd_date}</td>
          <td>${d.eta_busan_date}</td>
          <td><strong style="color:#0284c7;">${d.bonded_wh}</strong></td>
          <td>
            <span class="badge ${d.demurrage_info.badge_class}">
              ${d.demurrage_info.d_day_text}
            </span>
            ${d.demurrage_info.demurrage_loss_krw > 0 ? `<div style="font-size:10.5px; color:red; font-weight:700;">손실: ${d.demurrage_info.demurrage_loss_krw.toLocaleString()}원</div>` : ''}
          </td>
          <td>
            <span class="badge ${d.do_status==='DO_ISSUED'?'badge-green':'badge-orange'}">${d.do_status}</span>
          </td>
          <td>
            ${d.status === 'ON_WATER' ? `<button class="btn btn-primary" style="padding:2px 6px; font-size:11px;" onclick="openPreSaleModal('${d.deal_no}')">선점예약</button>` : ''}
            ${d.status === 'PORT_IN' && d.do_status !== 'DO_ISSUED' ? `<button class="btn btn-outline" style="padding:2px 6px; font-size:11px;" onclick="openDOModal('${d.deal_no}')">D/O승인</button>` : ''}
            ${d.status === 'ON_WATER' ? `<button class="btn btn-outline" style="padding:2px 6px; font-size:11px;" onclick="doDischarge('${d.deal_no}')">양하등록</button>` : ''}
          </td>
        </tr>
      `).join('') || '<tr><td colspan="16" style="text-align:center;">수입 계약 내역이 없습니다.</td></tr>';
    }

    function openDealModal() { document.getElementById('dealModal').style.display = 'flex'; }
    function closeModal(id) { document.getElementById(id).style.display = 'none'; }

    async function submitTradeDeal() {
      const dealNo = document.getElementById('d_deal_no').value.trim();
      const blNo = document.getElementById('d_bl').value.trim();
      if (!dealNo || !blNo) { alert('Deal No와 B/L NO를 입력하세요.'); return; }

      const payload = {
        deal_no: dealNo,
        company_name: document.getElementById('d_comp').value,
        supplier: document.getElementById('d_supp').value,
        carrier_line: document.getElementById('d_line').value,
        bl_no: blNo,
        container_no: document.getElementById('d_cntr').value,
        cut_name: document.getElementById('d_cut').value,
        brand: document.getElementById('d_brand_est').value.split('/')[0].trim(),
        box_qty: parseInt(document.getElementById('d_box').value),
        weight_kg: parseFloat(document.getElementById('d_weight').value),
        contract_unit_price: parseFloat(document.getElementById('d_price').value),
        pay_type: document.getElementById('d_pay_type').value,
        etd_date: document.getElementById('d_etd').value,
        eta_busan_date: document.getElementById('d_eta').value,
        due_date: document.getElementById('d_due').value,
        bonded_wh: document.getElementById('d_wh').value
      };

      const res = await apiFetch('/api/trades/deals', {
        method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)
      });
      alert(res.message);
      closeModal('dealModal');
      initPage();
    }

    function openPreSaleModal(dealNo) {
      const deal = cachedDeals.find(d => d.deal_no === dealNo);
      if (!deal) return;
      document.getElementById('ps_deal_no').value = deal.deal_no;
      document.getElementById('ps_deal_disp').value = `${deal.deal_no} | ${deal.cut_name} (가용: ${deal.avail_presale_box} Box)`;
      document.getElementById('ps_box').value = deal.avail_presale_box;
      document.getElementById('presaleModal').style.display = 'flex';
    }

    async function submitPreSaleHold() {
      const dealNo = document.getElementById('ps_deal_no').value;
      const cust = document.getElementById('ps_customer').value.trim();
      const box = parseInt(document.getElementById('ps_box').value);
      const price = parseFloat(document.getElementById('ps_price').value);

      if (!cust || !box) { alert('고객사 및 예약 수량을 입력하세요.'); return; }

      const res = await apiFetch('/api/trades/presale-hold', {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ deal_no: dealNo, customer: cust, reserved_box_qty: box, target_unit_price: price })
      });
      alert(res.message);
      closeModal('presaleModal');
      initPage();
    }

    function openDOModal(dealNo) {
      document.getElementById('do_deal_no').value = dealNo;
      document.getElementById('do_number').value = `DO-OOCL-${new Date().toISOString().slice(0,10).replace(/-/g,'')}-01`;
      document.getElementById('do_valid_date').value = new Date(Date.now() + 10*86400000).toISOString().slice(0,10);
      document.getElementById('doModal').style.display = 'flex';
    }

    async function submitDOApprove() {
      const dealNo = document.getElementById('do_deal_no').value;
      const doNum = document.getElementById('do_number').value.trim();
      const doVal = document.getElementById('do_valid_date').value;

      const res = await apiFetch(`/api/trades/deals/${dealNo}/approve-do`, {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ do_number: doNum, do_valid_date: doVal })
      });
      alert(res.message);
      closeModal('doModal');
      initPage();
    }

    async function doDischarge(dealNo) {
      const dt = prompt('부산항 양하일자를 입력하세요 (YYYY-MM-DD):', new Date().toISOString().slice(0,10));
      if (!dt) return;
      const res = await apiFetch(`/api/trades/deals/${dealNo}/discharge?discharge_date_str=${dt}`, { method: 'POST' });
      alert(res.message);
      initPage();
    }

    initPage();
  </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def serve_trade_ui():
    return HTML_PAGE

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("module2_trade:app", host="0.0.0.0", port=port, reload=True)
