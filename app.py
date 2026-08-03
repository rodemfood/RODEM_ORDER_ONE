import io
import json
import os
import re
import secrets
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, make_response, redirect, render_template, request, send_file
from sqlalchemy import Column, Integer, String, Text, create_engine, select, update
from sqlalchemy.orm import declarative_base, sessionmaker
import xlwt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

BASE_DIR = Path(__file__).resolve().parent
DATABASE_URL = os.environ.get("DATABASE_URL", f"sqlite:///{BASE_DIR / 'data' / 'rodem_order_one.db'}")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+psycopg://", 1)
elif DATABASE_URL.startswith("postgresql://") and "+psycopg" not in DATABASE_URL:
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+psycopg://", 1)

engine_kwargs = {"pool_pre_ping": True}
if DATABASE_URL.startswith("sqlite"):
    engine_kwargs["connect_args"] = {"check_same_thread": False}
engine = create_engine(DATABASE_URL, **engine_kwargs)
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)
Base = declarative_base()

app = Flask(__name__)
app.config["JSON_AS_ASCII"] = False
app.config["MAX_CONTENT_LENGTH"] = 1024 * 1024


class Customer(Base):
    __tablename__ = "customers"
    id = Column(Integer, primary_key=True)
    token = Column(String(128), unique=True, nullable=False, index=True)
    company = Column(String(100), nullable=False)
    receiver = Column(String(50), nullable=False)
    phone = Column(String(30), nullable=False)
    postal_code = Column(String(20), default="")
    address = Column(String(250), nullable=False)
    created_at = Column(String(19), nullable=False)


class Order(Base):
    __tablename__ = "orders"
    id = Column(Integer, primary_key=True)
    order_no = Column(String(40), unique=True, nullable=False, index=True)
    request_id = Column(String(80), unique=True, nullable=False, index=True)
    customer_token = Column(String(128), nullable=False, index=True)
    company = Column(String(100), nullable=False)
    receiver = Column(String(50), nullable=False)
    phone = Column(String(30), nullable=False)
    postal_code = Column(String(20), default="")
    address = Column(String(250), nullable=False)
    memo = Column(String(500), default="")
    items_json = Column(Text, nullable=False)
    total_qty = Column(Integer, nullable=False)
    status = Column(String(20), nullable=False, default="신규")
    created_at = Column(String(19), nullable=False)
    invoiced_at = Column(String(19), default="")


Base.metadata.create_all(engine)


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def make_order_no():
    return "R" + datetime.now().strftime("%y%m%d%H%M%S%f")[:18]


def clean_text(value, max_len=200):
    return re.sub(r"[\x00-\x1f\x7f]", " ", str(value or "")).strip()[:max_len]


def clean_phone(value):
    return clean_text(value, 30)


def customer_to_dict(row):
    return {
        "company": row.company,
        "receiver": row.receiver,
        "phone": row.phone,
        "postal_code": row.postal_code or "",
        "address": row.address,
    }


def order_to_dict(row):
    return {
        "id": row.id,
        "order_no": row.order_no,
        "company": row.company,
        "receiver": row.receiver,
        "phone": row.phone,
        "postal_code": row.postal_code or "",
        "address": row.address,
        "memo": row.memo or "",
        "items": json.loads(row.items_json),
        "total_qty": row.total_qty,
        "status": row.status,
        "created_at": row.created_at,
        "invoiced_at": row.invoiced_at or "",
    }


def get_customer():
    token = request.cookies.get("rodem_customer", "")
    if not token:
        return None
    with SessionLocal() as session:
        return session.scalar(select(Customer).where(Customer.token == token))


def logen_product_text(items):
    # 로젠 필수 규칙: 첫 제품부터 모든 제품 앞에 # 삽입
    return "".join(f"#{clean_text(item['name'], 100).replace('#', '')}{int(item['qty'])}" for item in items)


@app.get("/")
def home():
    return redirect("/order")


@app.get("/health")
def health():
    return jsonify(ok=True, service="RODEM ORDER ONE", time=now_text())


@app.get("/order")
def order_page():
    return render_template("customer.html")


@app.get("/staff")
def staff_page():
    return render_template("staff.html")


@app.get("/api/customer/me")
def customer_me():
    customer = get_customer()
    return jsonify(customer=customer_to_dict(customer) if customer else None)


@app.post("/api/customer/register")
def register_customer():
    data = request.get_json(silent=True) or {}
    values = {
        "company": clean_text(data.get("company"), 100),
        "receiver": clean_text(data.get("receiver"), 50),
        "phone": clean_phone(data.get("phone")),
        "postal_code": clean_text(data.get("postal_code"), 20),
        "address": clean_text(data.get("address"), 250),
    }
    if not values["company"] or not values["receiver"] or not values["phone"] or not values["address"]:
        return jsonify(error="업체명, 받는 분, 연락처, 배송지 주소를 입력해 주세요."), 400

    token = secrets.token_urlsafe(32)
    customer = Customer(token=token, created_at=now_text(), **values)
    with SessionLocal() as session:
        session.add(customer)
        session.commit()

    response = make_response(jsonify(customer=customer_to_dict(customer)))
    response.set_cookie(
        "rodem_customer", token, max_age=31536000 * 3,
        httponly=True, samesite="Lax", secure=request.is_secure,
    )
    return response


@app.post("/api/orders")
def create_order():
    customer = get_customer()
    if not customer:
        return jsonify(error="배송지 정보를 먼저 저장해 주세요."), 401

    data = request.get_json(silent=True) or {}
    request_id = clean_text(data.get("request_id"), 80) or secrets.token_urlsafe(24)
    items = []
    for item in (data.get("items") or [])[:8]:
        name = clean_text(item.get("name"), 100).replace("#", "")
        try:
            qty = int(item.get("qty", 0))
        except (TypeError, ValueError):
            qty = 0
        if name and 0 < qty <= 999999:
            items.append({"name": name, "qty": qty})
    if not items:
        return jsonify(error="제품명과 낱개 수량을 한 개 이상 입력해 주세요."), 400

    with SessionLocal() as session:
        existing = session.scalar(select(Order).where(Order.request_id == request_id))
        if existing:
            return jsonify(order_no=existing.order_no, total_qty=existing.total_qty, duplicate=True)

        total_qty = sum(item["qty"] for item in items)
        row = Order(
            order_no=make_order_no(), request_id=request_id,
            customer_token=customer.token, company=customer.company,
            receiver=customer.receiver, phone=customer.phone,
            postal_code=customer.postal_code or "", address=customer.address,
            memo=clean_text(data.get("memo"), 500),
            items_json=json.dumps(items, ensure_ascii=False), total_qty=total_qty,
            status="신규", created_at=now_text(), invoiced_at="",
        )
        session.add(row)
        session.commit()
        return jsonify(order_no=row.order_no, total_qty=row.total_qty, duplicate=False)


@app.get("/api/staff/orders")
def staff_orders():
    with SessionLocal() as session:
        rows = session.scalars(select(Order).order_by(Order.id.desc())).all()
        return jsonify(orders=[order_to_dict(row) for row in rows])


@app.post("/api/staff/mark-new")
def mark_new():
    data = request.get_json(silent=True) or {}
    try:
        order_id = int(data.get("id"))
    except (TypeError, ValueError):
        return jsonify(error="주문번호가 올바르지 않습니다."), 400
    with SessionLocal() as session:
        session.execute(update(Order).where(Order.id == order_id).values(status="신규", invoiced_at=""))
        session.commit()
    return jsonify(ok=True)


def selected_orders(ids):
    with SessionLocal() as session:
        return session.scalars(select(Order).where(Order.id.in_(ids)).order_by(Order.id)).all()


@app.post("/api/staff/export-logen")
def export_logen():
    data = request.get_json(silent=True) or {}
    try:
        ids = sorted({int(v) for v in data.get("ids", [])})
    except (TypeError, ValueError):
        ids = []
    if not ids:
        return jsonify(error="송장을 생성할 주문을 선택해 주세요."), 400

    rows = selected_orders(ids)
    if not rows:
        return jsonify(error="선택한 주문을 찾을 수 없습니다."), 404

    # 로젠 공식 열 순서에 맞춘 .xls 출력
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Sheet4")
    headers = [
        "수화주명", "전화번호", "휴대폰", "우편번호", "주소", "상품명",
        "박스수량", "배송메시지", "선착분", "운임", "", "주문자명"
    ]
    header_style = xlwt.easyxf(
        "font: bold on, color white; pattern: pattern solid, fore_colour green;"
        "align: horiz center, vert center; borders: left thin, right thin, top thin, bottom thin;"
    )
    cell_style = xlwt.easyxf(
        "align: vert top, wrap on; borders: left thin, right thin, top thin, bottom thin;"
    )
    for col, header in enumerate(headers):
        ws.write(0, col, header, header_style)
    widths = [18, 18, 18, 12, 50, 70, 10, 40, 10, 10, 5, 22]
    for i, width in enumerate(widths):
        ws.col(i).width = width * 256

    for r, row in enumerate(rows, start=1):
        items = json.loads(row.items_json)
        values = [
            row.receiver,
            row.phone,
            row.phone,
            row.postal_code or "",
            row.address,
            logen_product_text(items),
            1,
            row.memo or "",
            "",
            "",
            "",
            row.company,
        ]
        for c, value in enumerate(values):
            ws.write(r, c, value, cell_style)

    with SessionLocal() as session:
        session.execute(
            update(Order).where(Order.id.in_(ids)).values(status="송장생성", invoiced_at=now_text())
        )
        session.commit()

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output, as_attachment=True,
        download_name=f"RODEM_LOGEN_{datetime.now():%Y%m%d_%H%M}.xls",
        mimetype="application/vnd.ms-excel",
    )


def style_xlsx_header(ws, headers):
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="188754")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


@app.post("/api/staff/export-backup")
def export_backup():
    with SessionLocal() as session:
        rows = session.scalars(select(Order).order_by(Order.id.desc())).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "주문백업"
    style_xlsx_header(ws, [
        "상태", "주문번호", "접수시간", "업체명", "받는 분", "연락처",
        "우편번호", "주소", "주문상품", "로젠상품명", "총 낱개수량",
        "배송요청사항", "송장생성시간"
    ])
    for row in rows:
        items = json.loads(row.items_json)
        products = " / ".join(f"{i['name']} {i['qty']}개" for i in items)
        ws.append([
            row.status, row.order_no, row.created_at, row.company, row.receiver,
            row.phone, row.postal_code or "", row.address, products,
            logen_product_text(items), row.total_qty, row.memo or "", row.invoiced_at or ""
        ])
    for col, width in {
        "A":12,"B":24,"C":20,"D":22,"E":16,"F":18,"G":12,"H":45,
        "I":65,"J":65,"K":14,"L":38,"M":20
    }.items():
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output, as_attachment=True,
        download_name=f"RODEM_ORDER_BACKUP_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.errorhandler(413)
def too_large(_):
    return jsonify(error="입력 데이터가 너무 큽니다."), 413


if __name__ == "__main__":
    from waitress import serve
    port = int(os.environ.get("PORT", "5000"))
    serve(app, host="0.0.0.0", port=port)
