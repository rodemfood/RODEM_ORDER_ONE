import io
import json
import os
import secrets
import sqlite3
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, make_response, redirect, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("DATABASE_PATH", BASE_DIR / "data" / "rodem_order_one.db"))
DB_PATH.parent.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.config["JSON_AS_ASCII"] = False
app.config["MAX_CONTENT_LENGTH"] = 1024 * 1024


def db():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with db() as con:
        con.execute("""
            CREATE TABLE IF NOT EXISTS customers (
                token TEXT PRIMARY KEY,
                company TEXT NOT NULL,
                receiver TEXT NOT NULL,
                phone TEXT NOT NULL,
                postal_code TEXT DEFAULT '',
                address TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        con.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_no TEXT UNIQUE NOT NULL,
                customer_token TEXT NOT NULL,
                company TEXT NOT NULL,
                receiver TEXT NOT NULL,
                phone TEXT NOT NULL,
                postal_code TEXT DEFAULT '',
                address TEXT NOT NULL,
                memo TEXT DEFAULT '',
                items_json TEXT NOT NULL,
                total_qty INTEGER NOT NULL,
                status TEXT NOT NULL DEFAULT '신규',
                created_at TEXT NOT NULL,
                invoiced_at TEXT DEFAULT ''
            )
        """)


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def order_no():
    return "R" + datetime.now().strftime("%y%m%d%H%M%S%f")[:18]


def customer_from_cookie():
    token = request.cookies.get("rodem_customer", "")
    if not token:
        return None
    with db() as con:
        row = con.execute("SELECT * FROM customers WHERE token=?", (token,)).fetchone()
    return dict(row) if row else None


def clean_text(value, max_len=200):
    return str(value or "").strip()[:max_len]


@app.get("/")
def home():
    return redirect("/order")


@app.get("/order")
def order_page():
    return render_template("customer.html")


@app.get("/staff")
def staff_page():
    return render_template("staff.html")


@app.get("/api/customer/me")
def customer_me():
    return jsonify(customer=customer_from_cookie())


@app.post("/api/customer/register")
def register_customer():
    data = request.get_json(silent=True) or {}
    required = {
        "company": clean_text(data.get("company"), 100),
        "receiver": clean_text(data.get("receiver"), 50),
        "phone": clean_text(data.get("phone"), 30),
        "address": clean_text(data.get("address"), 250),
    }
    if not all(required.values()):
        return jsonify(error="업체명, 받는 분, 연락처, 배송지 주소를 입력해 주세요."), 400

    token = secrets.token_urlsafe(32)
    postal_code = clean_text(data.get("postal_code"), 20)
    with db() as con:
        con.execute(
            "INSERT INTO customers(token,company,receiver,phone,postal_code,address,created_at) VALUES(?,?,?,?,?,?,?)",
            (token, required["company"], required["receiver"], required["phone"], postal_code, required["address"], now_text()),
        )
        row = con.execute("SELECT * FROM customers WHERE token=?", (token,)).fetchone()

    response = make_response(jsonify(customer=dict(row)))
    response.set_cookie("rodem_customer", token, max_age=31536000, httponly=True, samesite="Lax", secure=request.is_secure)
    return response


@app.post("/api/orders")
def create_order():
    customer = customer_from_cookie()
    if not customer:
        return jsonify(error="배송지 정보를 먼저 저장해 주세요."), 401

    data = request.get_json(silent=True) or {}
    items = []
    for item in data.get("items", [])[:8]:
        name = clean_text(item.get("name"), 100)
        try:
            qty = int(item.get("qty", 0))
        except (TypeError, ValueError):
            qty = 0
        if name and qty > 0:
            items.append({"name": name, "qty": qty})
    if not items:
        return jsonify(error="제품명과 낱개 수량을 한 개 이상 입력해 주세요."), 400

    total_qty = sum(item["qty"] for item in items)
    no = order_no()
    memo = clean_text(data.get("memo"), 500)
    with db() as con:
        con.execute(
            """INSERT INTO orders(
                order_no,customer_token,company,receiver,phone,postal_code,address,memo,
                items_json,total_qty,status,created_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                no, customer["token"], customer["company"], customer["receiver"], customer["phone"],
                customer.get("postal_code", ""), customer["address"], memo,
                json.dumps(items, ensure_ascii=False), total_qty, "신규", now_text(),
            ),
        )
    return jsonify(order_no=no, total_qty=total_qty)


@app.get("/api/staff/orders")
def staff_orders():
    with db() as con:
        rows = con.execute("SELECT * FROM orders ORDER BY id DESC").fetchall()
    orders = []
    for row in rows:
        obj = dict(row)
        obj["items"] = json.loads(obj.pop("items_json"))
        orders.append(obj)
    return jsonify(orders=orders)


def styled_header(ws, headers):
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="188754")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


@app.post("/api/staff/export-logen")
def export_logen():
    data = request.get_json(silent=True) or {}
    try:
        ids = sorted({int(v) for v in data.get("ids", [])})
    except (TypeError, ValueError):
        ids = []
    if not ids:
        return jsonify(error="송장을 생성할 주문을 선택해 주세요."), 400

    placeholders = ",".join("?" for _ in ids)
    with db() as con:
        rows = con.execute(f"SELECT * FROM orders WHERE id IN ({placeholders}) ORDER BY id", ids).fetchall()
        if not rows:
            return jsonify(error="선택한 주문을 찾을 수 없습니다."), 404

        wb = Workbook()
        ws = wb.active
        ws.title = "로젠택배 업로드"
        styled_header(ws, [
            "받는분성명", "받는분전화번호", "받는분기타연락처", "받는분우편번호",
            "받는분주소(전체,분할)", "품목명", "박스수량", "배송메세지", "주문번호"
        ])
        for row in rows:
            items = json.loads(row["items_json"])
            product_text = "".join(f"#{item['name']}{item['qty']}" for item in items)
            ws.append([
                row["receiver"], row["phone"], "", row["postal_code"], row["address"],
                product_text, 1, row["memo"], row["order_no"]
            ])
        widths = {"A":18,"B":18,"C":18,"D":14,"E":48,"F":65,"G":12,"H":38,"I":24}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        invoice_time = now_text()
        con.execute(f"UPDATE orders SET status='송장생성', invoiced_at=? WHERE id IN ({placeholders})", [invoice_time, *ids])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"RODEM_LOGEN_{datetime.now():%Y%m%d_%H%M}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/staff/export-backup")
def export_backup():
    with db() as con:
        rows = con.execute("SELECT * FROM orders ORDER BY id DESC").fetchall()
    wb = Workbook()
    ws = wb.active
    ws.title = "주문백업"
    styled_header(ws, [
        "상태", "주문번호", "접수시간", "업체명", "받는분", "연락처", "우편번호",
        "주소", "주문상품", "총수량", "배송요청사항", "송장생성시간"
    ])
    for row in rows:
        items = json.loads(row["items_json"])
        products = " / ".join(f"{item['name']} {item['qty']}개" for item in items)
        ws.append([
            row["status"], row["order_no"], row["created_at"], row["company"], row["receiver"],
            row["phone"], row["postal_code"], row["address"], products, row["total_qty"],
            row["memo"], row["invoiced_at"]
        ])
    for col, width in {"A":12,"B":24,"C":20,"D":22,"E":16,"F":18,"G":12,"H":45,"I":65,"J":12,"K":38,"L":20}.items():
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"RODEM_ORDER_BACKUP_{datetime.now():%Y%m%d_%H%M}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/health")
def health():
    return jsonify(ok=True)


init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8765")), debug=False)
