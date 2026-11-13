"""Web app local: python -m app.main  (hoặc chạy run.bat)."""
from __future__ import annotations

import base64
import threading
from datetime import date
from pathlib import Path
from typing import Literal

import uvicorn
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, RedirectResponse
from pydantic import BaseModel

from . import db
from .config import settings
from .gdt.client import GDTClient, GDTError
from .providers import load_providers
from .sync import DONE_PDF, JOBS, _storage_base, process_pdf, start_job

app = FastAPI(title="Đồng bộ hóa đơn điện tử")
TEMPLATES = Path(__file__).resolve().parent / "templates"

# Phiên đăng nhập cổng thuế chỉ giữ trong RAM (không lưu mật khẩu / token xuống đĩa).
PENDING: dict[str, GDTClient] = {}   # captcha key -> client
SESSIONS: dict[str, GDTClient] = {}  # MST -> client đã đăng nhập
_lock = threading.Lock()


def _new_client() -> GDTClient:
    return GDTClient(settings.gdt_base_url, settings.gdt_verify_ssl, settings.gdt_delay)


@app.exception_handler(db.DBNotConfigured)
async def _db_missing(_, exc):  # noqa: ANN001
    from fastapi.responses import JSONResponse
    return JSONResponse({"detail": str(exc)}, status_code=500)


@app.get("/")
def index():
    return FileResponse(TEMPLATES / "index.html")


# ---------- đăng nhập cổng thuế ----------
@app.get("/api/captcha")
def captcha():
    gdt = _new_client()
    try:
        c = gdt.get_captcha()
    except GDTError as e:
        gdt.close()
        raise HTTPException(502, str(e))
    with _lock:
        PENDING[c["key"]] = gdt
        while len(PENDING) > 20:
            PENDING.pop(next(iter(PENDING))).close()
    img = "data:image/svg+xml;base64," + base64.b64encode(c["svg"].encode("utf-8")).decode()
    return {"key": c["key"], "image": img}


class LoginIn(BaseModel):
    username: str
    password: str
    ckey: str
    cvalue: str


@app.post("/api/login")
def login(body: LoginIn):
    gdt = PENDING.pop(body.ckey, None) or _new_client()
    try:
        gdt.login(body.username.strip(), body.password, body.ckey, body.cvalue.strip())
    except GDTError as e:
        gdt.close()
        raise HTTPException(401, str(e))
    mst = body.username.strip()
    old = SESSIONS.pop(mst, None)
    if old:
        old.close()
    SESSIONS[mst] = gdt
    profile = gdt.profile()
    name = profile.get("name") or profile.get("tennnt")
    try:
        db.upsert_company(mst, name)
    except Exception:  # noqa: BLE001 - chưa cấu hình DB vẫn cho đăng nhập
        pass
    return {"ok": True, "mst": mst, "name": name}


@app.get("/api/sessions")
def sessions():
    return {"sessions": list(SESSIONS.keys())}


# ---------- đồng bộ ----------
class SyncIn(BaseModel):
    mst: str
    direction: Literal["purchase", "sold"]
    date_from: date
    date_to: date
    fetch_pdf: bool = True


@app.post("/api/sync")
def sync(body: SyncIn):
    gdt = SESSIONS.get(body.mst)
    if not gdt:
        raise HTTPException(401, "Chưa đăng nhập cổng thuế cho MST này")
    if body.date_from > body.date_to:
        raise HTTPException(400, "Từ ngày phải trước đến ngày")
    db.client()  # báo lỗi sớm nếu chưa cấu hình Supabase
    job = start_job(gdt, body.mst, body.direction, body.date_from, body.date_to, body.fetch_pdf)
    return job.public()


# ---------- tự động tải qua Chrome (người dùng tự đăng nhập) ----------
@app.post("/api/browser/open")
def browser_open():
    from .gdt.browser import BROWSER, BrowserError
    try:
        return BROWSER.open()
    except BrowserError as e:
        raise HTTPException(500, str(e))
    except ImportError:
        raise HTTPException(500, "Thiếu thư viện playwright. Chạy: .venv\\Scripts\\pip install playwright")


@app.get("/api/browser/status")
def browser_status():
    from .auto_download import running_job
    from .gdt.browser import BROWSER
    st = BROWSER.status()
    job = running_job()
    st["job_id"] = job.id if job else None
    st["output_dir"] = settings.output_dir
    return st


class AutoIn(BaseModel):
    date_from: date
    date_to: date
    kinds: list[str]
    excel: bool = True
    xml: bool = True
    pdf: bool = True
    easy_captcha: bool = False


@app.post("/api/auto/start")
def auto_start(body: AutoIn):
    from .auto_download import KINDS, running_job, start
    if running_job():
        raise HTTPException(409, "Đang có một lượt tải chạy, hãy chờ xong hoặc bấm Dừng")
    if body.date_from > body.date_to:
        raise HTTPException(400, "Từ ngày phải trước đến ngày")
    kinds = [k for k in KINDS if k in body.kinds]
    if not kinds:
        raise HTTPException(400, "Chọn ít nhất một loại hóa đơn")
    return start(body.date_from, body.date_to, kinds, body.excel, body.xml, body.pdf, body.easy_captcha).public()


def _auto_job(job_id: str):
    from .auto_download import JOBS
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Không có lượt tải này")
    return job


@app.get("/api/auto/{job_id}")
def auto_status(job_id: str):
    return _auto_job(job_id).public()


class CaptchaIn(BaseModel):
    value: str | None = None  # để trống = bỏ qua loại hóa đơn đang chờ captcha


@app.post("/api/auto/{job_id}/captcha")
def auto_captcha(job_id: str, body: CaptchaIn):
    _auto_job(job_id).submit_captcha(body.value)
    return {"ok": True}


@app.post("/api/auto/{job_id}/stop")
def auto_stop(job_id: str):
    job = _auto_job(job_id)
    job.cancel = True
    job.submit_captcha(None)
    return {"ok": True}


# ---------- đối chiếu tờ khai 01/GTGT ----------
@app.post("/api/reconcile")
async def reconcile(file: UploadFile = File(...), mst: str = Form("")):
    from . import reports
    from .gdt.browser import BROWSER

    data = await file.read()
    vals, period = reports.parse_tax_return(data, file.filename or "")
    if not vals:
        raise HTTPException(400, "Không đọc được chỉ tiêu nào trong file tờ khai (cần file XML eTax hoặc Excel HTKK của tờ khai 01/GTGT)")
    if not period:
        raise HTTPException(400, "Không xác định được kỳ tính thuế trong file tờ khai")
    mst = (mst or BROWSER.status().get("mst") or "").strip()
    base = Path(settings.output_dir)
    if not mst:  # chưa đăng nhập: lấy thư mục MST duy nhất nếu có
        dirs = [d.name for d in base.iterdir() if d.is_dir()] if base.is_dir() else []
        mst = dirs[0] if len(dirs) == 1 else ""
    if not mst:
        raise HTTPException(400, "Không biết MST nào – hãy đăng nhập cổng thuế hoặc nhập MST")
    root = base / mst
    months = reports.period_months(*period)
    rows = reports.load_period_data(root, months)
    missing = [f"{m:02d}/{y}" for y, m in months
               if not (root / f"{y:04d}-{m:02d}" / f".dulieu_{y:04d}-{m:02d}.json").exists()]
    if len(missing) == len(months):
        raise HTTPException(400, f"Chưa có dữ liệu hóa đơn của kỳ này ({', '.join(missing)}). "
                                 "Hãy tải hóa đơn trọn các tháng đó trước.")
    dest, result = reports.write_reconcile(root, vals, period, rows, file.filename or "")
    return {"file": str(dest), "period": period, "missing_months": missing, "rows": len(rows), "result": result}


# ---------- đọc chứng từ: hóa đơn hàng loạt, sao kê ngân hàng ----------
def _doc_root(mst: str) -> tuple[str, Path]:
    from .gdt.browser import BROWSER
    base = Path(settings.output_dir)
    mst = (mst or "").strip() or (BROWSER.status().get("mst") or "")
    if not mst and base.is_dir():
        dirs = [d.name for d in base.iterdir() if d.is_dir()]
        mst = dirs[0] if len(dirs) == 1 else ""
    if not mst:
        raise HTTPException(400, "Không biết MST nào – hãy đăng nhập cổng thuế hoặc nhập MST")
    return mst, base / mst


async def _read_uploads(files: list[UploadFile]) -> list[tuple[str, bytes]]:
    out, total = [], 0
    for f in files:
        data = await f.read()
        total += len(data)
        if total > MAX_UPLOAD:
            raise HTTPException(413, "Tổng dung lượng file quá 300 MB, hãy chia nhỏ")
        out.append((f.filename or "file", data))
    return out


@app.get("/api/ocr/status")
def ocr_status():
    from . import ocr
    return {"language": ocr.ocr_language()}


@app.post("/api/docs/invoices")
async def docs_invoices(files: list[UploadFile] = File(...), mst: str = Form("")):
    """Đọc hàng loạt XML/PDF/ZIP hóa đơn -> Excel + kiểm tra với dữ liệu đã tải từ cổng thuế."""
    import asyncio
    from datetime import datetime

    from . import docreader
    mst, root = _doc_root(mst)
    payload = await _read_uploads(files)
    recs = await asyncio.to_thread(docreader.read_many, payload)
    idx = docreader.load_portal_index(root)
    dest = docreader.write_invoice_report(root / "DocChungTu" / f"DocHoaDon_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                                          recs, idx)
    stat: dict[str, int] = {}
    for r in recs:
        k = "Lỗi đọc" if r.get("loi") and not r.get("shdon") else r.get("_check", "")
        k = ("Khớp với cổng thuế" if k.startswith("Khớp") else "LỆCH tổng tiền" if k.startswith("LỆCH")
             else "Chưa có dữ liệu cổng tháng đó" if k.startswith("Chưa có dữ liệu") else k)
        stat[k] = stat.get(k, 0) + 1
    return {"file": str(dest), "count": len(recs), "stat": stat,
            "sample": [{k: r.get(k) for k in ("file", "khhdon", "shdon", "ngay_lap", "nbten", "tgtttbso", "_check", "loi")}
                       for r in recs[:200]]}


@app.post("/api/docs/bank")
async def docs_bank(files: list[UploadFile] = File(...), mst: str = Form("")):
    """Đọc sao kê (Excel/CSV/PDF) và ghép với hóa đơn đã tải trong khoảng ngày của sao kê."""
    from datetime import datetime

    from . import bank
    mst, root = _doc_root(mst)
    payload = await _read_uploads(files)
    txns, errors = [], []
    for name, data in payload:
        tx, err = bank.read_statement(name, data)
        txns += tx
        if err:
            errors.append(f"{name}: {err}")
    if not txns:
        raise HTTPException(400, "Không đọc được giao dịch nào. " + " | ".join(errors))
    txns.sort(key=lambda t: t["ngay"])
    d1, d2 = txns[0]["ngay"], txns[-1]["ngay"]
    invoices = bank.invoices_for_range(root, d1, d2)
    pairs, unpaid, unmatched = bank.match(invoices, txns)
    dest = bank.write_report(root / "DocChungTu" / f"DoiChieu_SaoKe_{d1:%Y%m%d}_{d2:%Y%m%d}_{datetime.now():%H%M%S}.xlsx",
                             pairs, unpaid, unmatched, txns, errors)
    return {"file": str(dest), "from": d1.isoformat(), "to": d2.isoformat(), "txns": len(txns),
            "invoices": len(invoices), "matched": len(pairs), "unpaid": len(unpaid), "unmatched": len(unmatched),
            "tien_vao": sum(t["co"] for t in txns), "tien_ra": sum(t["no"] for t in txns), "errors": errors}


# ---------- chạy tự động theo lịch ----------
@app.get("/api/schedule")
def schedule_get():
    from .scheduler import SCHEDULER
    return SCHEDULER.status()


class ScheduleIn(BaseModel):
    enabled: bool
    day: int = 5
    hour: int = 8
    kinds: list[str] = ["mua-vao", "mua-vao-khong-ma", "ban-ra"]


@app.post("/api/schedule")
def schedule_set(body: ScheduleIn):
    from .scheduler import SCHEDULER
    if not 1 <= body.day <= 28 or not 0 <= body.hour <= 23:
        raise HTTPException(400, "Ngày chạy 1–28, giờ 0–23")
    SCHEDULER.configure(body.enabled, body.day, body.hour, body.kinds)
    return SCHEDULER.status()


@app.post("/api/schedule/run-now")
def schedule_run_now():
    from .scheduler import SCHEDULER
    SCHEDULER.trigger()
    return SCHEDULER.status()


@app.post("/api/open-folder")
def open_folder(path: str | None = None):
    import os
    target = Path(path or settings.output_dir)
    base = Path(settings.output_dir).resolve()
    if not target.resolve().is_relative_to(base):
        raise HTTPException(400, "Chỉ mở được thư mục lưu hóa đơn")
    if target.suffix.lower() != ".xlsx":
        target.mkdir(parents=True, exist_ok=True)
    os.startfile(target)  # noqa: S606 - app chạy trên máy người dùng
    return {"ok": True}


# ---------- danh sách MST khách hàng + theo dõi thư mục Downloads ----------
@app.on_event("startup")
def _start_watcher():
    from .scheduler import SCHEDULER
    from .watcher import WATCHER
    WATCHER.ensure_running()
    SCHEDULER.ensure_running()


@app.get("/api/companies")
def companies():
    return {"items": db.list_companies()}


class CompanyIn(BaseModel):
    mst: str
    name: str | None = None


@app.post("/api/companies")
def add_company(body: CompanyIn):
    mst = body.mst.strip().replace("-", "")
    if not mst.isdigit() or len(mst) not in (10, 13):
        raise HTTPException(400, "MST phải gồm 10 số (hoặc 13 số với chi nhánh)")
    db.upsert_company(mst, (body.name or "").strip() or None)
    return {"ok": True}


@app.delete("/api/companies/{mst}")
def delete_company(mst: str):
    db.client().table("companies").delete().eq("mst", mst).execute()
    return {"ok": True}


@app.get("/api/watcher")
def watcher_status():
    from .watcher import WATCHER
    return WATCHER.status()


class WatcherIn(BaseModel):
    enabled: bool
    folder: str | None = None
    fetch_pdf: bool | None = None


@app.post("/api/watcher")
def watcher_config(body: WatcherIn):
    from .watcher import WATCHER
    WATCHER.configure(body.enabled, body.folder, body.fetch_pdf)
    return WATCHER.status()


# ---------- nhập file tải từ cổng thuế ----------
MAX_UPLOAD = 300 * 1024 * 1024  # 300 MB / lần


@app.post("/api/import")
async def import_files(mst: str = Form(...), direction: str = Form("purchase"),
                       fetch_pdf: bool = Form(True), files: list[UploadFile] = File(...)):
    from .sync import start_import_job

    mst = mst.strip()
    if not mst.isdigit() or len(mst) not in (10, 13, 14):
        raise HTTPException(400, "MST không hợp lệ (10 hoặc 13 số, chi nhánh dạng 0123456789-001 thì bỏ dấu '-')")
    if direction not in ("purchase", "sold", "auto"):
        raise HTTPException(400, "Loại phải là mua vào hoặc bán ra")
    db.client()
    payload, total = [], 0
    for f in files:
        data = await f.read()
        total += len(data)
        if total > MAX_UPLOAD:
            raise HTTPException(413, "Tổng dung lượng file quá 300 MB, hãy chia nhỏ")
        payload.append((f.filename or "file", data))
    job = start_import_job(mst, direction, payload, fetch_pdf)
    return job.public()


# ---------- nguồn TConnect ----------
@app.get("/api/tconnect/status")
def tconnect_status():
    configured = all([settings.tconnect_partner_code, settings.tconnect_client_id,
                      settings.tconnect_client_secret, settings.tconnect_aes_key,
                      settings.tconnect_username, settings.tconnect_password])
    companies = []
    try:
        companies = [c for c in db.list_companies() if c.get("tconnect_account_id")]
    except Exception:  # noqa: BLE001
        pass
    return {"configured": configured, "auth_base": settings.tconnect_auth_base,
            "companies": [{"mst": c["mst"], "name": c.get("name")} for c in companies]}


class TConnectAccountIn(BaseModel):
    mst: str
    password: str


@app.post("/api/tconnect/accounts")
def tconnect_connect(body: TConnectAccountIn):
    """Đăng ký (hoặc cập nhật mật khẩu) tài khoản cổng thuế với TConnect."""
    from .sources.tconnect import TConnectError
    from .sync import tconnect_client

    mst = body.mst.strip()
    try:
        tc = tconnect_client()
    except TConnectError as e:
        raise HTTPException(400, str(e))
    try:
        existing = (db.get_company(mst) or {}).get("tconnect_account_id")
        if existing:
            tc.update_tax_password(existing, body.password)
            return {"ok": True, "mst": mst, "account_id": existing, "updated": True}
        try:
            acc = tc.register_tax_account(mst, body.password, initial_sync=False)
        except TConnectError as e:
            # Có thể MST đã đăng ký trước đó -> tìm lại theo maskedMst (2 số đầu + 3 số cuối)
            match = [a for a in tc.list_tax_accounts()
                     if str(a.get("maskedMst", "")).startswith(mst[:2])
                     and str(a.get("maskedMst", "")).endswith(mst[-3:])]
            if len(match) != 1:
                raise HTTPException(400, str(e))
            acc = match[0]
            tc.update_tax_password(acc["id"], body.password)
        db.set_tconnect_account(mst, acc["id"])
        return {"ok": True, "mst": mst, "account_id": acc["id"], "updated": False}
    except TConnectError as e:
        raise HTTPException(400, str(e))
    finally:
        tc.close()


class TConnectSyncIn(BaseModel):
    mst: str
    direction: Literal["purchase", "sold", "both"]
    date_from: date
    date_to: date
    fetch_pdf: bool = True


@app.post("/api/tconnect/sync")
def tconnect_sync(body: TConnectSyncIn):
    from .sync import start_tconnect_job
    if body.date_from > body.date_to:
        raise HTTPException(400, "Từ ngày phải trước đến ngày")
    if not (db.get_company(body.mst) or {}).get("tconnect_account_id"):
        raise HTTPException(400, "MST này chưa kết nối TConnect")
    job = start_tconnect_job(body.mst, body.direction, body.date_from, body.date_to, body.fetch_pdf)
    return job.public()


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Không có job")
    if job.status == "auth_error":
        SESSIONS.pop(job.owner_mst, None)  # vô hại với job TConnect (không dùng SESSIONS)
    return job.public()


# ---------- hóa đơn ----------
@app.get("/api/invoices")
def invoices(mst: str | None = None, direction: str | None = None, date_from: str | None = None,
             date_to: str | None = None, pdf_status: str | None = None, q: str | None = None,
             limit: int = 200, offset: int = 0):
    return {"items": db.list_invoices(mst, direction, date_from, date_to, pdf_status, q,
                                      min(limit, 1000), offset)}


def _get(invoice_id: str) -> dict:
    row = db.get_invoice(invoice_id)
    if not row:
        raise HTTPException(404, "Không có hóa đơn")
    return row


@app.get("/api/invoices/{invoice_id}")
def invoice_detail(invoice_id: str):
    return _get(invoice_id)


@app.get("/api/invoices/{invoice_id}/file/{kind}")
def invoice_file(invoice_id: str, kind: Literal["xml", "pdf"]):
    row = _get(invoice_id)
    path = row.get(f"{kind}_path")
    if not path:
        raise HTTPException(404, f"Hóa đơn chưa có file {kind.upper()}")
    return RedirectResponse(db.signed_url(path))


@app.post("/api/invoices/{invoice_id}/refetch-pdf")
def refetch_pdf(invoice_id: str):
    row = _get(invoice_id)
    if not row.get("parsed"):
        raise HTTPException(400, "Hóa đơn chưa có XML, hãy đồng bộ lại trước")
    upd = process_pdf(row, row["parsed"], load_providers())
    return db.update_invoice(invoice_id, upd)


@app.post("/api/invoices/{invoice_id}/upload-pdf")
async def upload_pdf(invoice_id: str, file: UploadFile = File(...)):
    row = _get(invoice_id)
    data = await file.read()
    if data[:5] != b"%PDF-":
        raise HTTPException(400, "File tải lên không phải PDF")
    path = _storage_base(row) + ".pdf"
    db.upload(path, data, "application/pdf")
    return db.update_invoice(invoice_id, {
        "pdf_path": path, "pdf_status": "manual", "pdf_is_original": True,
        "pdf_source_url": None, "pdf_error": None,
    })


@app.post("/api/retry-failed-pdf")
def retry_failed(mst: str, direction: str | None = None):
    """Thử lấy lại PDF gốc cho các hóa đơn đang 'rendered' hoặc 'failed'."""
    providers = load_providers()
    items = [r for r in db.list_invoices(mst, direction, limit=1000)
             if r["pdf_status"] not in DONE_PDF]
    ok = 0
    for r in items:
        full = db.get_invoice(r["id"])
        if not full or not full.get("parsed"):
            continue
        upd = process_pdf(full, full["parsed"], providers)
        db.update_invoice(r["id"], upd)
        ok += upd.get("pdf_status") == "original"
    return {"checked": len(items), "original": ok}


@app.get("/api/providers")
def providers():
    return {"providers": [p.__dict__ for p in load_providers()]}


def run() -> None:
    uvicorn.run(app, host=settings.host, port=settings.port)


if __name__ == "__main__":
    run()
