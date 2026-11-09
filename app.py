# -*- coding: utf-8 -*-
"""창고 재고 불출 웹페이지 - 모든 데이터는 data/ 폴더의 Excel 파일에 저장됩니다."""
import io
import os
import re
import secrets
import threading
from datetime import datetime, timedelta

from flask import Flask, abort, flash, g, jsonify, redirect, render_template, request, send_file, url_for
from markupsafe import Markup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.environ.get("DATA_DIR") or os.path.join(BASE, "data")  # 테스트 시 DATA_DIR로 다른 폴더 지정 가능
STOCK_FILE = os.path.join(DATA, "재고현황.xlsx")
LOG_FILE = os.path.join(DATA, "변경로그.xlsx")
IMG_DIR = os.path.join(DATA, "images")
THUMB_DIR = os.path.join(IMG_DIR, "thumbs")
DEVICE_FILE = os.path.join(DATA, "접속기기.xlsx")
ORDER_FILE = os.path.join(DATA, "발주요청.xlsx")

STOCK_COLS = ["제품코드", "제품명", "제조사", "적재위치", "현재고수량", "안전재고", "긴급도", "최종변경일시", "비고"]
LOG_COLS = ["일시", "구분", "제품명", "제조사", "적재위치", "변경전 재고", "변경수량", "변경후 재고", "담당자", "비고", "제품코드"]
TEMPLATE_COLS = ["제품명", "제조사", "적재위치", "현재고수량"]          # 필수
TEMPLATE_OPT = ["안전재고", "긴급도"]                                  # 선택 (비우면 유지/기본값)
URGENCY = ("A", "B", "C")
DEVICE_COLS = ["기기코드", "이름", "최초IP", "최근IP", "최초등록일시", "최근접속일시", "브라우저"]
ORDER_COLS = ["요청번호", "요청일시", "제품코드", "제품명", "제조사", "적재위치", "요청시 재고", "안전재고",
              "요청수량", "요청자", "비고", "상태", "처리일시", "처리자", "기기코드", "발주묶음"]
ORDER_STATES = ("진행중", "완료", "반려")
OLD_STATE = {"요청": "진행중", "취소": "반려"}  # 예전 상태 이름 → 새 이름
DEVICE_COOKIE = "wh_device"

SEED = [  # 제품명, 제조사, 적재위치, 현재고
    ("A", "Molex", "A-01-01", 120),
    ("B", "HRS", "A-01-02", 80),
    ("C", "TE", "B-02-01", 45),
]
# 임시 기준값: 제품명 → (안전재고, 긴급도). 그 외 신규 제품 기본값은 (0, "C")
DEFAULT_CRITERIA = {"A": (100, "A"), "B": (50, "B"), "C": (30, "C")}
NEW_DEFAULT = (0, "C")

app = Flask(__name__)
app.secret_key = "warehouse-local"
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 휴대폰 사진 업로드 허용
app.config["TEMPLATES_AUTO_RELOAD"] = True  # 화면(templates) 파일을 고치면 서버 재시작 없이 바로 반영
lock = threading.Lock()

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")


class ExcelBusy(Exception):
    pass


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def style_header(ws, widths):
    for i, w in enumerate(widths, start=1):
        c = ws.cell(row=1, column=i)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[c.column_letter].width = w
    ws.freeze_panes = "A2"


def save(wb, path):
    try:
        wb.save(path)
    except PermissionError:
        raise ExcelBusy(f"'{os.path.basename(path)}' 파일이 Excel에서 열려 있어 저장하지 못했습니다. 파일을 닫고 다시 시도해 주세요.")


def to_int(v):
    try:
        return int(float(str(v).replace(",", "").strip()))
    except (TypeError, ValueError):
        return None


def default_criteria(name):
    return DEFAULT_CRITERIA.get(name, NEW_DEFAULT)


# ---------- Excel 저장소 ----------
def ensure_book(path, title, cols, widths):
    if not os.path.exists(path):
        wb = Workbook()
        ws = wb.active
        ws.title = title
        ws.append(cols)
        style_header(ws, widths)
        wb.save(path)


def init_files():
    os.makedirs(DATA, exist_ok=True)
    os.makedirs(THUMB_DIR, exist_ok=True)
    ensure_book(DEVICE_FILE, "기기목록", DEVICE_COLS, [14, 12, 16, 16, 20, 20, 50])
    ensure_book(ORDER_FILE, "발주요청", ORDER_COLS, [10, 20, 10, 14, 12, 12, 11, 10, 10, 10, 24, 8, 20, 10, 12, 10])
    wb = load_workbook(ORDER_FILE)  # 예전 상태 이름(요청/취소)을 새 이름(진행중/반려)으로 변환
    changed = False
    ws = wb["발주요청"]
    if ws.cell(row=1, column=len(ORDER_COLS)).value != ORDER_COLS[-1]:  # '발주묶음' 열 추가
        ws.cell(row=1, column=len(ORDER_COLS), value=ORDER_COLS[-1])
        style_header(ws, [10, 20, 10, 14, 12, 12, 11, 10, 10, 10, 24, 8, 20, 10, 12, 10])
        changed = True
    for row in wb["발주요청"].iter_rows(min_row=2):
        if row[11].value in OLD_STATE:
            row[11].value = OLD_STATE[row[11].value]
            changed = True
    if changed:
        wb.save(ORDER_FILE)
    if not os.path.exists(LOG_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "변경로그"
        ws.append(LOG_COLS)
        style_header(ws, [20, 10, 14, 12, 12, 12, 10, 12, 10, 30, 10])
        wb.save(LOG_FILE)
    else:  # 기존 로그 파일에 '제품코드' 열이 없으면 머리글만 추가 (기존 행은 그대로)
        wb = load_workbook(LOG_FILE)
        ws = wb["변경로그"]
        if ws.cell(row=1, column=len(LOG_COLS)).value != LOG_COLS[-1]:
            ws.cell(row=1, column=len(LOG_COLS), value=LOG_COLS[-1])
            style_header(ws, [20, 10, 14, 12, 12, 12, 10, 12, 10, 30, 10])
            wb.save(LOG_FILE)
    if not os.path.exists(STOCK_FILE):
        items, logs = [], []
        for i, (name, maker, loc, qty) in enumerate(SEED, start=1):
            safety, urg = default_criteria(name)
            it = dict(code=f"P{i:05d}", name=name, maker=maker, loc=loc, qty=qty, safety=safety, urg=urg, updated=now())
            items.append(it)
            logs.append(log_row("초기등록", it, 0, qty, "시스템", "샘플 데이터"))
        write_stock(items)
        append_logs(logs)
    else:
        migrate_stock()


def migrate_stock():
    """예전 형식(제품코드/안전재고/긴급도 없음, 입고·출고 누적 있음)을 새 형식으로 변환. 재고 값은 그대로 유지."""
    wb = load_workbook(STOCK_FILE)
    ws = wb["재고"]
    header = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    if header[:8] == STOCK_COLS[:8]:  # '비고' 열은 없어도 됨 (처음 저장할 때 붙음)
        return
    items = read_stock()  # 머리글 이름으로 읽으므로 예전 형식도 읽힘. 빠진 값은 여기서 채워짐
    write_stock(items)
    append_logs([log_row("형식변환", it, it["qty"], 0, "시스템",
                         f"제품코드 {it['code']} 부여, 안전재고 {it['safety']}, 긴급도 {it['urg']}") for it in items])


def read_stock():
    wb = load_workbook(STOCK_FILE)
    ws = wb["재고"]
    header = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    col = {h: i for i, h in enumerate(header) if h}
    items = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        row = {h: (r[i] if i < len(r) else None) for h, i in col.items()}
        g = row.get
        if not g("제품명"):
            continue
        code = str(g("제품코드") or "").strip().upper()
        items.append({"code": code, "name": str(g("제품명")), "maker": str(g("제조사") or ""), "loc": str(g("적재위치") or ""),
                      "qty": to_int(g("현재고수량")) or 0, "safety": g("안전재고"), "urg": str(g("긴급도") or "").strip().upper(),
                      "updated": g("최종변경일시") or "", "memo": str(g("비고") or "")})
    for it in items:  # 코드 없는 행에는 순서대로 코드 부여, 빈 기준값은 기본값
        if not it["code"]:
            it["code"] = next_code(items)
        d_safety, d_urg = default_criteria(it["name"])
        s = to_int(it["safety"])
        it["safety"] = s if s is not None and s >= 0 else d_safety
        if it["urg"] not in URGENCY:
            it["urg"] = d_urg
    return items


def next_code(items):
    nums = [int(m.group(1)) for it in items if (m := re.fullmatch(r"P(\d+)", it.get("code") or ""))]
    return f"P{(max(nums) + 1) if nums else 1:05d}"


def write_stock(items):
    wb = Workbook()
    ws = wb.active
    ws.title = "재고"
    ws.append(STOCK_COLS)
    style_header(ws, [12, 16, 12, 12, 12, 10, 8, 20, 30])
    for it in items:
        ws.append([it["code"], it["name"], it["maker"], it["loc"], it["qty"], it["safety"], it["urg"], it["updated"], it.get("memo") or None])
    save(wb, STOCK_FILE)


def log_row(kind, it, before, delta, worker, memo):
    return [now(), kind, it["name"], it["maker"], it["loc"], before, delta, before + delta, worker, memo, it["code"]]


def append_logs(rows):
    wb = load_workbook(LOG_FILE)
    ws = wb["변경로그"]
    for r in rows:
        ws.append(r)
    save(wb, LOG_FILE)


def read_logs(limit=200):
    wb = load_workbook(LOG_FILE, read_only=True)
    rows = [tuple(r) + (None,) * (len(LOG_COLS) - len(r)) for r in wb["변경로그"].iter_rows(min_row=2, values_only=True)]
    wb.close()
    rows = [r[:len(LOG_COLS)] for r in reversed(rows)]
    return rows if limit is None else rows[:limit]


def find(items, name, maker):
    key = (name.strip().lower(), maker.strip().lower())
    for it in items:
        if (it["name"].lower(), it["maker"].lower()) == key:
            return it
    return None


def find_code(items, code):
    code = (code or "").strip().upper()
    return next((it for it in items if it["code"] == code), None)


def check_writable(*paths):
    for p in paths:  # 저장 전에 파일 잠금(Excel에서 열림) 여부 확인
        try:
            open(p, "a").close()
        except PermissionError:
            raise ExcelBusy(f"'{os.path.basename(p)}' 파일이 Excel에서 열려 있어 저장하지 못했습니다. 파일을 닫고 다시 시도해 주세요.")


def commit(items, logs):
    """재고와 로그를 함께 저장. 로그 파일이 열려 있으면 재고도 저장하지 않음."""
    check_writable(STOCK_FILE, LOG_FILE)
    write_stock(items)
    if logs:
        append_logs(logs)


# ---------- 사진 ----------
def photo_path(code, thumb=False):
    return os.path.join(THUMB_DIR if thumb else IMG_DIR, f"{code}.jpg")


def photo_ver(code):
    p = photo_path(code)
    return int(os.path.getmtime(p)) if os.path.exists(p) else 0


_use_cache = {"mtime": None, "counts": {}}


def usage_counts(days=30):
    """최근 30일 입고·출고 횟수 (사용빈도순 정렬용). 로그 파일이 바뀔 때만 다시 읽음"""
    try:
        mt = os.path.getmtime(LOG_FILE)
    except OSError:
        return {}
    if _use_cache["mtime"] == mt:
        return _use_cache["counts"]
    since = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    counts = {}
    try:
        for r in read_logs(limit=None):
            if r[1] in ("입고", "출고") and str(r[0] or "")[:10] >= since:
                key = str(r[10] or "").upper() or ("@" + str(r[2] or "").lower() + "|" + str(r[3] or "").lower())
                counts[key] = counts.get(key, 0) + 1
    except Exception:
        return _use_cache["counts"]
    _use_cache.update(mtime=mt, counts=counts)
    return counts


def public(it, ordered=None):
    d = dict(it)
    d["short"] = d["qty"] < d["safety"]
    # 상태 이모지: 부족 🔴 / 안전재고 근접(안전재고 이상 ~ 120% 미만) 🟡 / 정상 🟢
    d["mood"] = "🔴" if d["short"] else ("🟡" if d["safety"] > 0 and d["qty"] < d["safety"] * 1.2 else "🟢")
    d["photo"] = photo_ver(it["code"])
    d["ordered"] = (open_order_qty() if ordered is None else ordered).get(it["code"], 0)  # 발주 진행중 수량
    return d


# ---------- 접속 기기 등록 (비밀번호 없이 기기별 이름 1회 입력) ----------
_dev_cache = {"mtime": None, "rows": {}}


def read_devices():
    m = os.path.getmtime(DEVICE_FILE)
    if _dev_cache["mtime"] != m:
        wb = load_workbook(DEVICE_FILE, read_only=True)
        rows = {}
        for r in wb["기기목록"].iter_rows(min_row=2, values_only=True):
            r = list(r) + [None] * (len(DEVICE_COLS) - len(r))
            if r[0]:
                rows[str(r[0]).strip()] = r[:len(DEVICE_COLS)]
        wb.close()
        _dev_cache.update(mtime=m, rows=rows)
    return _dev_cache["rows"]


def client_ip():
    return request.remote_addr or ""


def update_device_row(code, **vals):
    wb = load_workbook(DEVICE_FILE)
    ws = wb["기기목록"]
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value or "").strip() == code:
            for k, v in vals.items():
                ws.cell(row=row[0].row, column=DEVICE_COLS.index(k) + 1, value=v)
            break
    save(wb, DEVICE_FILE)


def touch_device(code, row):
    """IP가 바뀌었거나 오늘 처음 접속이면 최근IP/최근접속일시만 갱신 (매 요청마다 Excel을 쓰지 않음)"""
    ip, last = client_ip(), str(row[5] or "")
    if row[3] == ip and last[:10] == now()[:10]:
        return
    try:
        with lock:
            update_device_row(code, **{"최근IP": ip, "최근접속일시": now()})
    except (ExcelBusy, OSError):
        pass


PAGE_ENDPOINTS = {"index", "shortage", "orders", "upload", "log", "labels", "label"}  # 미등록 기기에는 흐리게 + 이름 입력창


@app.before_request
def device_gate():
    g.device = None
    code = (request.cookies.get(DEVICE_COOKIE) or "").strip()
    if code:
        row = read_devices().get(code)
        if row:
            g.device = {"code": code, "name": str(row[1] or "")}
            touch_device(code, row)
    if g.device or request.endpoint in ("register", "static", "photo"):
        return None
    if request.method == "GET" and request.endpoint in PAGE_ENDPOINTS:
        return None
    if request.path.startswith("/api/"):
        return jsonify(ok=False, msg="먼저 이 기기를 등록해 주세요. (첫 화면에서 이름 입력)"), 403
    return redirect(url_for("index"))


@app.post("/register")
def register():
    d = request.get_json(silent=True) or request.form
    name = (d.get("name") or "").strip()
    if not name or len(name) > 30:
        return jsonify(ok=False, msg="이름을 입력해 주세요. (30자 이내)"), 400
    if g.device:
        return jsonify(ok=True, name=g.device["name"], msg="이미 등록된 기기입니다.")
    with lock:
        rows = read_devices()
        code = "D" + secrets.token_hex(4).upper()
        while code in rows:
            code = "D" + secrets.token_hex(4).upper()
        ip, t = client_ip(), now()
        try:
            wb = load_workbook(DEVICE_FILE)
            wb["기기목록"].append([code, name, ip, ip, t, t, request.headers.get("User-Agent", "")[:200]])
            save(wb, DEVICE_FILE)
        except ExcelBusy as e:
            return jsonify(ok=False, msg=str(e)), 409
    resp = jsonify(ok=True, name=name, code=code, msg=f"{name}님, 기기 등록이 완료되었습니다.")
    resp.set_cookie(DEVICE_COOKIE, code, max_age=10 * 365 * 24 * 3600, samesite="Lax", httponly=True)
    return resp


def who(worker):
    """담당자 칸이 비어 있으면 이 기기에 등록된 이름을 사용"""
    worker = (worker or "").strip()
    return worker or (g.device["name"] if g.get("device") else "")


@app.context_processor
def nav_counts():
    try:
        items = read_stock()
        return {"nav_short": sum(1 for i in items if i["qty"] < i["safety"]),
                "nav_orders": sum(1 for o in read_orders() if o["state"] == "진행중")}
    except Exception:
        return {"nav_short": 0, "nav_orders": 0}


# ---------- 발주요청 (data/발주요청.xlsx + 변경로그에 '발주요청' 기록) ----------
def read_orders():
    wb = load_workbook(ORDER_FILE, read_only=True)
    out = []
    for r in wb["발주요청"].iter_rows(min_row=2, values_only=True):
        r = list(r) + [None] * (len(ORDER_COLS) - len(r))
        if not r[0]:
            continue
        out.append({"no": str(r[0]), "at": str(r[1] or ""), "code": str(r[2] or ""), "name": str(r[3] or ""),
                    "maker": str(r[4] or ""), "loc": str(r[5] or ""), "stock": to_int(r[6]) or 0,
                    "safety": to_int(r[7]) or 0, "qty": to_int(r[8]) or 0, "worker": str(r[9] or ""),
                    "memo": str(r[10] or ""), "state": OLD_STATE.get(str(r[11] or "진행중"), str(r[11] or "진행중")), "done_at": str(r[12] or ""),
                    "done_by": str(r[13] or ""), "batch": str(r[15] or "")})
    wb.close()
    return out


def open_order_qty():
    m = {}
    for o in read_orders():
        if o["state"] == "진행중":
            m[o["code"]] = m.get(o["code"], 0) + o["qty"]
    return m


def next_order_no(orders):
    nums = [int(m.group(1)) for o in orders if (m := re.fullmatch(r"R(\d+)", o["no"]))]
    return f"R{(max(nums) + 1) if nums else 1:05d}"


# ---------- 화면 ----------
def stock_view(only_short):
    om = open_order_qty()
    all_items = [public(i, om) for i in read_stock()]
    uc = usage_counts()
    for i in all_items:
        i["use"] = uc.get(i["code"], 0) + uc.get("@" + i["name"].lower() + "|" + i["maker"].lower(), 0)
    # 부족 먼저 → 긴급도 A,B,C → 제품코드 순
    all_items.sort(key=lambda i: (not i["short"], i["urg"], i["code"]))
    items = [i for i in all_items if i["short"]] if only_short else all_items
    short = sum(1 for i in all_items if i["short"])
    return render_template("index.html", items=items, all_items=all_items, short=short, only_short=only_short)


@app.route("/")
def index():
    return stock_view(False)


@app.get("/shortage")
def shortage():
    return stock_view(True)


@app.get("/orders")
def orders():
    rows = list(reversed(read_orders()))
    stock = {it["code"]: it for it in read_stock()}
    for o in rows:  # 목록에는 지금 현재고/안전재고를 보여줌 (요청 당시 값은 상세 팝업에)
        it = stock.get(o["code"])
        o["cur_qty"], o["cur_safety"] = (it["qty"], it["safety"]) if it else (None, None)
    # 목록 한 줄 = 묶음. 완료 건은 같이 처리한 발주묶음끼리 '○○ 외 n건' 한 줄, 나머지는 한 건씩
    groups, by_key = [], {}
    for o in rows:
        key = o["batch"] if o["state"] == "완료" and o["batch"] else o["no"]
        if key not in by_key:
            by_key[key] = {"key": key, "state": o["state"], "items": []}
            groups.append(by_key[key])
        by_key[key]["items"].append(o)
    for g in groups:
        g["items"].reverse()  # 묶음 안에서는 요청 순서대로
        f = g["items"][0]
        g["n"] = len(g["items"])
        g["nos"] = [o["no"] for o in g["items"]]
        g["date"] = (f["done_at"] if g["state"] == "완료" and f["done_at"] else f["at"])[2:10]
        g["qty"] = sum(o["qty"] for o in g["items"])
    return render_template("orders.html", orders=rows, groups=groups)


@app.post("/api/order")
def order():
    """발주요청 등록 (수량 입력). 재고 수량은 바뀌지 않음"""
    d = request.get_json(force=True)
    qty = to_int(d.get("qty"))
    worker, memo = who(d.get("worker")), (d.get("memo") or "").strip()
    if not qty or qty <= 0:
        return jsonify(ok=False, msg="발주 수량을 1 이상으로 입력해 주세요."), 400
    with lock:
        it = find_code(read_stock(), d.get("code"))
        if not it:
            return jsonify(ok=False, msg="제품을 찾을 수 없습니다."), 404
        try:
            check_writable(ORDER_FILE, LOG_FILE)
            orders_ = read_orders()
            no = next_order_no(orders_)
            wb = load_workbook(ORDER_FILE)
            wb["발주요청"].append([no, now(), it["code"], it["name"], it["maker"], it["loc"], it["qty"], it["safety"],
                                qty, worker, memo, "진행중", "", "", g.device["code"] if g.device else ""])
            save(wb, ORDER_FILE)
            append_logs([log_row("발주요청", it, it["qty"], 0, worker,
                                 f"{no} 발주 {qty}개 요청" + (f" ({memo})" if memo else ""))])
        except ExcelBusy as e:
            return jsonify(ok=False, msg=str(e)), 409
    return jsonify(ok=True, item=public(it), msg=f"{it['name']} [{it['maker']}] 발주요청 {qty}개 등록 ({no})")


def set_order_states(nos, state, worker):
    """발주요청 상태 변경 (여러 건 한 번에). 완료는 한 번에 처리한 건을 같은 발주묶음 번호(B00001…)로 묶음.
    반환: (변경된 건수, 발주묶음 번호)"""
    check_writable(ORDER_FILE, LOG_FILE)
    wb = load_workbook(ORDER_FILE)
    ws = wb["발주요청"]
    stock = {it["code"]: it for it in read_stock()}
    bcol = len(ORDER_COLS)
    batch = ""
    if state == "완료":
        nums = [int(m.group(1)) for r in range(2, ws.max_row + 1)
                if (m := re.fullmatch(r"B(\d+)", str(ws.cell(row=r, column=bcol).value or "")))]
        batch = f"B{(max(nums) + 1) if nums else 1:05d}"
    logs, nos = [], set(nos)
    for row in ws.iter_rows(min_row=2):
        no = str(row[0].value or "")
        if no not in nos:
            continue
        old = str(row[11].value or "진행중")
        old = OLD_STATE.get(old, old)
        if old == state:
            continue
        row[11].value = state
        row[12].value = now() if state != "진행중" else ""
        row[13].value = worker if state != "진행중" else ""
        ws.cell(row=row[0].row, column=bcol, value=batch or None)
        vals = [c.value for c in row]
        it = {"name": vals[3], "maker": vals[4], "loc": vals[5], "code": vals[2]}
        cur = stock.get(vals[2])
        q = cur["qty"] if cur else (to_int(vals[6]) or 0)
        logs.append(log_row("발주" + {"완료": "완료", "반려": "반려", "진행중": "되돌림"}[state], it, q, 0, worker,
                            f"{no} 발주 {vals[8]}개: {old} → {state}" + (f" (묶음 {batch})" if batch else "")))
    if logs:
        save(wb, ORDER_FILE)
        append_logs(logs)
    return len(logs), (batch if logs else "")


@app.post("/api/orders/state")
def orders_state():
    """발주요청 상태 변경: 완료 / 반려 / 진행중(되돌리기). nos 에 요청번호 여러 개 → 일괄 처리"""
    d = request.get_json(force=True)
    state, nos = (d.get("state") or "").strip(), [str(n) for n in (d.get("nos") or [])]
    if state not in ORDER_STATES or not nos:
        return jsonify(ok=False, msg="처리할 발주요청과 상태를 선택해 주세요."), 400
    with lock:
        try:
            n, batch = set_order_states(nos, state, who(d.get("worker")))
        except ExcelBusy as e:
            return jsonify(ok=False, msg=str(e)), 409
    return jsonify(ok=True, n=n, batch=batch, msg=f"{n}건을 '{state}'(으)로 변경했습니다." if n else "변경된 내용이 없습니다.")


@app.post("/api/orders/memo")
def orders_memo():
    """발주건별 비고(메모) 저장. items: [{no, memo}, ...]"""
    d = request.get_json(force=True)
    want = {str(i.get("no")): str(i.get("memo") or "").strip()[:500] for i in (d.get("items") or []) if i.get("no")}
    if not want:
        return jsonify(ok=False, msg="저장할 발주건이 없습니다."), 400
    with lock:
        try:
            check_writable(ORDER_FILE)
            wb = load_workbook(ORDER_FILE)
            ws = wb["발주요청"]
            n = 0
            for row in ws.iter_rows(min_row=2):
                no = str(row[0].value or "")
                if no in want and str(row[10].value or "") != want[no]:
                    row[10].value = want[no] or None
                    n += 1
            if n:
                wb.save(ORDER_FILE)
            wb.close()
        except ExcelBusy as e:
            return jsonify(ok=False, msg=str(e)), 409
    return jsonify(ok=True, n=n, msg="비고를 저장했습니다." if n else "바뀐 내용이 없습니다.")


@app.post("/api/move")
def move():
    """입고/출고 처리"""
    d = request.get_json(force=True)
    kind, qty = d.get("kind"), to_int(d.get("qty"))
    worker, memo = who(d.get("worker")), (d.get("memo") or "").strip()
    if kind not in ("입고", "출고") or not qty or qty <= 0:
        return jsonify(ok=False, msg="수량을 1 이상으로 입력해 주세요."), 400
    with lock:
        items = read_stock()
        it = find_code(items, d.get("code"))
        if not it:
            return jsonify(ok=False, msg="제품을 찾을 수 없습니다."), 404
        before = it["qty"]
        if kind == "출고" and qty > before:
            return jsonify(ok=False, msg=f"현재고({before})보다 많이 출고할 수 없습니다."), 400
        delta = qty if kind == "입고" else -qty
        it["qty"] += delta
        it["updated"] = now()
        try:
            commit(items, [log_row(kind, it, before, delta, worker, memo)])
        except ExcelBusy as e:
            return jsonify(ok=False, msg=str(e)), 409
    return jsonify(ok=True, item=public(it), msg=f"{it['name']} [{it['maker']}] {kind} {qty}개 완료 (현재고 {it['qty']})")


@app.post("/api/memo")
def item_memo():
    """제품 비고(메모) 자동 저장. 변경로그에는 남기지 않음"""
    d = request.get_json(force=True)
    memo = str(d.get("memo") or "").strip()[:500]
    with lock:
        items = read_stock()
        it = find_code(items, d.get("code"))
        if not it:
            return jsonify(ok=False, msg="제품을 찾을 수 없습니다."), 404
        if it.get("memo", "") != memo:
            it["memo"] = memo
            try:
                check_writable(STOCK_FILE)
                write_stock(items)
            except ExcelBusy as e:
                return jsonify(ok=False, msg=str(e)), 409
    return jsonify(ok=True, item=public(it), msg="비고를 저장했습니다.")


@app.post("/api/location")
def change_location():
    d = request.get_json(force=True)
    loc = (d.get("loc") or "").strip()
    if not loc:
        return jsonify(ok=False, msg="적재위치를 입력해 주세요."), 400
    with lock:
        items = read_stock()
        it = find_code(items, d.get("code"))
        if not it:
            return jsonify(ok=False, msg="제품을 찾을 수 없습니다."), 404
        old = it["loc"]
        it["loc"], it["updated"] = loc, now()
        try:
            commit(items, [log_row("위치변경", it, it["qty"], 0, (d.get("worker") or "").strip(), f"{old} → {loc}")])
        except ExcelBusy as e:
            return jsonify(ok=False, msg=str(e)), 409
    return jsonify(ok=True, item=public(it), msg=f"적재위치 변경: {old} → {loc}")


@app.post("/api/rename")
def rename():
    """제품명 · 제조사 수정 (제품코드, 사진, 재고는 그대로)"""
    d = request.get_json(force=True)
    name, maker = (d.get("name") or "").strip(), (d.get("maker") or "").strip()
    if not name or not maker:
        return jsonify(ok=False, msg="제품명과 제조사를 모두 입력해 주세요."), 400
    with lock:
        items = read_stock()
        it = find_code(items, d.get("code"))
        if not it:
            return jsonify(ok=False, msg="제품을 찾을 수 없습니다."), 404
        dup = find(items, name, maker)
        if dup and dup["code"] != it["code"]:
            return jsonify(ok=False, msg=f"{name} [{maker}] 은(는) 이미 등록된 제품입니다. ({dup['code']})"), 409
        old_name, old_maker = it["name"], it["maker"]
        if (old_name, old_maker) == (name, maker):
            return jsonify(ok=True, item=public(it), msg="변경된 내용이 없습니다.")
        it["name"], it["maker"], it["updated"] = name, maker, now()
        memo = ", ".join(x for x in [f"제품명 {old_name} → {name}" if old_name != name else "",
                                      f"제조사 {old_maker} → {maker}" if old_maker != maker else ""] if x)
        try:
            commit(items, [log_row("정보수정", it, it["qty"], 0, who(d.get("worker")), memo)])
        except ExcelBusy as e:
            return jsonify(ok=False, msg=str(e)), 409
    return jsonify(ok=True, item=public(it), msg=f"수정 완료: {memo}")


@app.post("/api/criteria")
def change_criteria():
    """기준수정: 안전재고 / 긴급도"""
    d = request.get_json(force=True)
    safety, urg = to_int(d.get("safety")), (d.get("urg") or "").strip().upper()
    if safety is None or safety < 0 or urg not in URGENCY:
        return jsonify(ok=False, msg="안전재고는 0 이상, 긴급도는 A/B/C 중 하나여야 합니다."), 400
    with lock:
        items = read_stock()
        it = find_code(items, d.get("code"))
        if not it:
            return jsonify(ok=False, msg="제품을 찾을 수 없습니다."), 404
        old = (it["safety"], it["urg"])
        if old == (safety, urg):
            return jsonify(ok=True, item=public(it), msg="변경된 내용이 없습니다.")
        it["safety"], it["urg"], it["updated"] = safety, urg, now()
        try:
            commit(items, [log_row("기준수정", it, it["qty"], 0, (d.get("worker") or "").strip(),
                                   f"안전재고 {old[0]} → {safety}, 긴급도 {old[1]} → {urg}")])
        except ExcelBusy as e:
            return jsonify(ok=False, msg=str(e)), 409
    return jsonify(ok=True, item=public(it), msg=f"기준 변경: 안전재고 {safety}, 긴급도 {urg}")


@app.post("/add")
def add():
    f = request.form
    name, maker, loc = f.get("name", "").strip(), f.get("maker", "").strip(), f.get("loc", "").strip()
    qty = to_int(f.get("qty") or 0)
    d_safety, d_urg = default_criteria(name)
    safety = to_int(f.get("safety")) if (f.get("safety") or "").strip() else d_safety
    urg = (f.get("urg") or d_urg).strip().upper()
    if not name or not maker or not loc or qty is None or qty < 0 or safety is None or safety < 0 or urg not in URGENCY:
        flash("제품명, 제조사, 적재위치, 수량(0 이상), 안전재고(0 이상), 긴급도(A/B/C)를 확인해 주세요.", "error")
        return redirect(url_for("index"))
    with lock:
        items = read_stock()
        if find(items, name, maker):
            flash(f"{name} [{maker}] 은(는) 이미 등록된 제품입니다. 목록에서 눌러 입고 처리해 주세요.", "error")
            return redirect(url_for("index"))
        it = dict(code=next_code(items), name=name, maker=maker, loc=loc, qty=qty, safety=safety, urg=urg, updated=now())
        items.append(it)
        try:
            commit(items, [log_row("신규등록", it, 0, qty, f.get("worker", "").strip(),
                                   f"웹 개별 추가 (안전재고 {safety}, 긴급도 {urg})")])
        except ExcelBusy as e:
            flash(str(e), "error")
            return redirect(url_for("index"))
    flash(f"{name} [{maker}] 제품이 추가되었습니다. 제품코드 {it['code']}", "ok")
    return redirect(url_for("index"))


# ---------- 사진 업로드 ----------
@app.post("/api/photo/<code>")
def upload_photo(code):
    from PIL import Image, ImageOps
    f = request.files.get("photo")
    if not f or not f.filename:
        return jsonify(ok=False, msg="사진 파일을 선택해 주세요."), 400
    with lock:
        items = read_stock()
        it = find_code(items, code)
        if not it:
            return jsonify(ok=False, msg="제품을 찾을 수 없습니다."), 404
        try:
            img = Image.open(f.stream)
            img = ImageOps.exif_transpose(img).convert("RGB")  # 휴대폰 사진 방향 보정
        except Exception:
            return jsonify(ok=False, msg="이미지 파일을 읽지 못했습니다. (JPG/PNG 사진을 올려 주세요)"), 400
        img.thumbnail((1200, 1200))
        img.save(photo_path(it["code"]), "JPEG", quality=85)
        th = img.copy()
        th.thumbnail((200, 200))
        th.save(photo_path(it["code"], thumb=True), "JPEG", quality=80)
        try:
            append_logs([log_row("사진등록", it, it["qty"], 0, (request.form.get("worker") or "").strip(), f"{it['code']}.jpg")])
        except ExcelBusy:
            pass  # 사진은 저장됨. 로그 파일이 열려 있으면 로그만 생략
    return jsonify(ok=True, item=public(it), msg="사진이 저장되었습니다.")


@app.post("/api/photo/<code>/delete")
def delete_photo(code):
    with lock:
        it = find_code(read_stock(), code)
        if not it:
            return jsonify(ok=False, msg="제품을 찾을 수 없습니다."), 404
        removed = False
        for p in (photo_path(it["code"]), photo_path(it["code"], thumb=True)):
            if os.path.exists(p):
                os.remove(p)
                removed = True
        if not removed:
            return jsonify(ok=False, msg="등록된 사진이 없습니다."), 400
        try:
            append_logs([log_row("사진삭제", it, it["qty"], 0, who((request.get_json(silent=True) or {}).get("worker")), f"{it['code']}.jpg 삭제")])
        except ExcelBusy:
            pass
    return jsonify(ok=True, item=public(it), msg="사진이 삭제되었습니다.")


@app.get("/photo/<code>.jpg")
def photo(code):
    thumb = request.args.get("thumb") == "1"
    p = photo_path(code.upper(), thumb)
    if thumb and not os.path.exists(p):  # 작은 사진이 없으면 원본으로
        p = photo_path(code.upper())
    if not re.fullmatch(r"P\d+", code.upper()) or not os.path.exists(p):
        abort(404)
    return send_file(p, mimetype="image/jpeg", max_age=0)


# ---------- 바코드 라벨 ----------
def barcode_svg(code):
    import barcode
    from barcode.writer import SVGWriter
    buf = io.BytesIO()
    barcode.Code128(code, writer=SVGWriter()).write(
        buf, options={"module_width": 0.33, "module_height": 14, "quiet_zone": 3, "write_text": False})
    svg = buf.getvalue().decode("utf-8")
    return Markup(svg[svg.index("<svg"):])


@app.get("/label/<code>")
def label(code):
    it = find_code(read_stock(), code)
    if not it:
        abort(404)
    return render_template("labels.html", labels=[(it, barcode_svg(it["code"]))], single=True)


@app.get("/labels")
def labels():
    items = sorted(read_stock(), key=lambda i: i["code"])
    return render_template("labels.html", labels=[(it, barcode_svg(it["code"])) for it in items], single=False)


# ---------- Excel 일괄 업로드 ----------
def parse_upload(file):
    wb = load_workbook(file, data_only=True)
    ws = wb.active
    header = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    missing = [c for c in TEMPLATE_COLS if c not in header]
    if missing:
        raise ValueError(f"양식의 머리글이 다릅니다. 빠진 항목: {', '.join(missing)}  (양식을 다시 내려받아 사용해 주세요)")
    idx = {c: header.index(c) for c in TEMPLATE_COLS + TEMPLATE_OPT if c in header}
    items = read_stock()
    rows, errors, seen = [], [], set()
    for n, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        get = lambda c: r[idx[c]] if c in idx and idx[c] < len(r) else None
        name, maker, loc = (str(get(c) or "").strip() for c in ("제품명", "제조사", "적재위치"))
        raw_qty, raw_safety, raw_urg = get("현재고수량"), get("안전재고"), str(get("긴급도") or "").strip().upper()
        if not any([name, maker, loc, raw_qty not in (None, "")]):
            continue
        qty = to_int(raw_qty)
        if not name or not maker:
            errors.append(f"{n}행: 제품명과 제조사는 필수입니다."); continue
        if qty is None or qty < 0:
            errors.append(f"{n}행: 현재고수량은 0 이상의 숫자여야 합니다."); continue
        safety = None
        if raw_safety not in (None, ""):
            safety = to_int(raw_safety)
            if safety is None or safety < 0:
                errors.append(f"{n}행: 안전재고는 0 이상의 숫자여야 합니다. (비워두면 유지)"); continue
        if raw_urg and raw_urg not in URGENCY:
            errors.append(f"{n}행: 긴급도는 A, B, C 중 하나여야 합니다. (비워두면 유지)"); continue
        key = (name.lower(), maker.lower())
        if key in seen:
            errors.append(f"{n}행: {name} [{maker}] 이(가) 양식에 중복되어 있습니다."); continue
        seen.add(key)
        cur = find(items, name, maker)
        if cur:
            loc = loc or cur["loc"]
            safety = cur["safety"] if safety is None else safety
            urg = raw_urg or cur["urg"]
            same = (cur["qty"], cur["loc"], cur["safety"], cur["urg"]) == (qty, loc, safety, urg)
            rows.append(dict(code=cur["code"], name=cur["name"], maker=cur["maker"], loc=loc, qty=qty, before=cur["qty"],
                             old_loc=cur["loc"], safety=safety, urg=urg, old_safety=cur["safety"], old_urg=cur["urg"],
                             status="변경없음" if same else "수정"))
        else:
            if not loc:
                errors.append(f"{n}행: 신규 제품은 적재위치가 필요합니다."); continue
            d_safety, d_urg = default_criteria(name)
            rows.append(dict(code="", name=name, maker=maker, loc=loc, qty=qty, before=0, old_loc="",
                             safety=d_safety if safety is None else safety, urg=raw_urg or d_urg,
                             old_safety=None, old_urg=None, status="신규"))
    return rows, errors


@app.route("/upload", methods=["GET", "POST"])
def upload():
    if request.method == "GET":
        return render_template("upload.html", rows=None)
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        flash(".xlsx 형식의 Excel 파일을 선택해 주세요.", "error")
        return redirect(url_for("upload"))
    try:
        rows, errors = parse_upload(f)
    except Exception as e:  # 손상된 파일, 머리글 오류 등
        flash(f"파일을 읽지 못했습니다: {e}", "error")
        return redirect(url_for("upload"))
    return render_template("upload.html", rows=rows, errors=errors, filename=f.filename)


@app.post("/upload/apply")
def upload_apply():
    f = request.form
    cols = zip(f.getlist("name"), f.getlist("maker"), f.getlist("loc"), f.getlist("qty"), f.getlist("safety"), f.getlist("urg"))
    worker, filename = f.get("worker", "").strip(), f.get("filename", "")
    added = changed = 0
    with lock:
        items = read_stock()
        logs = []
        for name, maker, loc, q, s, u in cols:
            qty, safety, urg = to_int(q), to_int(s), (u or "").upper()
            if qty is None or qty < 0 or safety is None or safety < 0 or urg not in URGENCY:
                continue
            it = find(items, name, maker)
            if it:
                if (it["qty"], it["loc"], it["safety"], it["urg"]) == (qty, loc, safety, urg):
                    continue
                before, old_loc, old_s, old_u = it["qty"], it["loc"], it["safety"], it["urg"]
                it.update(qty=qty, loc=loc, safety=safety, urg=urg, updated=now())
                memo = f"양식 업로드 ({filename})"
                if old_loc != loc:
                    memo += f", 위치 {old_loc} → {loc}"
                if old_s != safety:
                    memo += f", 안전재고 {old_s} → {safety}"
                if old_u != urg:
                    memo += f", 긴급도 {old_u} → {urg}"
                logs.append(log_row("일괄수정", it, before, qty - before, worker, memo))
                changed += 1
            else:
                it = dict(code=next_code(items), name=name, maker=maker, loc=loc, qty=qty, safety=safety, urg=urg, updated=now())
                items.append(it)
                logs.append(log_row("일괄신규", it, 0, qty, worker, f"양식 업로드 ({filename})"))
                added += 1
        if logs:
            try:
                commit(items, logs)
            except ExcelBusy as e:
                flash(str(e), "error")
                return redirect(url_for("upload"))
    flash(f"일괄 적용 완료: 신규 {added}건, 수정 {changed}건", "ok")
    return redirect(url_for("index"))


@app.get("/template")
def template():
    wb = Workbook()
    ws = wb.active
    ws.title = "재고입력"
    ws.append(TEMPLATE_COLS + TEMPLATE_OPT)
    style_header(ws, [16, 14, 14, 14, 12, 10])
    for it in read_stock():  # 현재 재고를 미리 채워두면 수정만 하면 됨
        ws.append([it["name"], it["maker"], it["loc"], it["qty"], it["safety"], it["urg"]])
    dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0",
                        showErrorMessage=True, error="0 이상의 정수를 입력하세요.")
    ws.add_data_validation(dv)
    dv.add("D2:E1000")
    dv2 = DataValidation(type="list", formula1='"A,B,C"', allow_blank=True,
                         showErrorMessage=True, error="A, B, C 중 하나를 고르세요.")
    ws.add_data_validation(dv2)
    dv2.add("F2:F1000")
    guide = wb.create_sheet("작성방법")
    for line in [
        "■ 재고 일괄 입력 양식 작성방법",
        "1. '재고입력' 시트의 2행부터 한 줄에 제품 하나씩 입력합니다.",
        "2. 제품명 + 제조사가 같으면 같은 제품으로 봅니다.",
        "3. 이미 있는 제품: '현재고수량'으로 재고가 맞춰집니다. (변경 내용은 로그에 기록)",
        "4. 없는 제품: 신규 등록되고 제품코드(P00001…)가 자동으로 붙습니다. 신규 제품은 적재위치가 꼭 필요합니다.",
        "5. 기존 제품의 적재위치를 비워두면 기존 위치가 유지됩니다.",
        "6. 안전재고 · 긴급도(A/B/C)는 선택 항목입니다. 비워두면 기존 값 유지 (신규 제품은 기본값).",
        "7. 업로드하면 적용 전에 미리보기 화면에서 변경 내용을 확인할 수 있습니다.",
        "※ 머리글(1행)은 수정하지 마세요.",
    ]:
        guide.append([line])
    guide.column_dimensions["A"].width = 100
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="재고_일괄입력_양식.xlsx")


# ---------- 로그 / 다운로드 ----------
@app.get("/log")
def log():
    return render_template("log.html", logs=read_logs(), cols=LOG_COLS)


@app.get("/download/po")
def download_po():
    """완료된 발주묶음(PO)만 엑셀로. 파일명 = PO 번호 (여러 개면 B00001_B00002.xlsx)"""
    want = [b.strip().upper() for b in request.args.get("b", "").split(",") if re.fullmatch(r"B\d+", b.strip().upper())]
    rows = [o for o in read_orders() if o["batch"] in want]
    if not rows:
        return redirect(url_for("orders"))
    pos = [b for b in want if any(o["batch"] == b for o in rows)]
    wb = Workbook(); ws = wb.active; ws.title = "PO"
    head = ["PO 번호", "요청번호", "제품코드", "제품명", "제조사", "수량", "요청일시", "요청자", "완료일시", "처리자", "비고"]
    ws.append(head)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="F97316"); c.alignment = Alignment(horizontal="center")
    for o in sorted(rows, key=lambda o: (pos.index(o["batch"]), o["no"])):
        ws.append([o["batch"], o["no"], o["code"], o["name"], o["maker"], o["qty"], o["at"], o["worker"], o["done_at"], o["done_by"], o["memo"]])
    ws.append([]); ws.append(["", "", "", "합계", "", sum(o["qty"] for o in rows)])
    ws.cell(row=ws.max_row, column=4).font = Font(bold=True); ws.cell(row=ws.max_row, column=6).font = Font(bold=True)
    for i, w in enumerate([10, 10, 10, 22, 14, 8, 18, 10, 18, 10, 20], 1):
        ws.column_dimensions[chr(64 + i)].width = w
    ws.freeze_panes = "A2"
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="_".join(pos) + ".xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/download/<kind>")
def download(kind):
    path = {"stock": STOCK_FILE, "log": LOG_FILE, "order": ORDER_FILE, "device": DEVICE_FILE}.get(kind)
    if not path:
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


init_files()

if __name__ == "__main__":
    if os.name == "nt":  # 모니터가 꺼져도(덮개 닫힘) Windows가 대기 상태로 들어가지 않고 서버가 계속 응답하도록 요청
        import ctypes
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000001)  # ES_CONTINUOUS | ES_SYSTEM_REQUIRED
    port = int(os.environ.get("PORT", 8080))
    print(f"\n  창고 재고 불출 페이지: http://localhost:{port}\n")
    # use_reloader: app.py 를 고치면 서버가 스스로 다시 시작됨 (수동 재시작 불필요)
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=True)
