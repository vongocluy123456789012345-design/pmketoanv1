"""Web app local: python -m app.main  (hoặc chạy run.bat)."""
from __future__ import annotations

import base64
import os
import re
import threading
from datetime import date
from pathlib import Path
from typing import Literal

import uvicorn
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from pydantic import BaseModel

from . import db
from .config import settings
from .gdt.client import GDTClient, GDTError
from .providers import load_providers
from .sync import DONE_PDF, JOBS, _storage_base, process_pdf, start_job

app = FastAPI(title="AWBook")
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
    # Không cho cửa sổ phần mềm giữ giao diện cũ trong bộ nhớ đệm (sau khi cập nhật/đóng gói lại vẫn hiện bản mới)
    return FileResponse(TEMPLATES / "index.html",
                        headers={"Cache-Control": "no-store, no-cache, must-revalidate", "Pragma": "no-cache"})


# ---------- logo AWBook (icon cửa sổ / thanh taskbar thay cho biểu tượng trình duyệt)
ASSETS = Path(__file__).resolve().parent / "assets"


@app.get("/favicon.ico")
def favicon():
    return FileResponse(ASSETS / "awbook.ico", media_type="image/x-icon")


@app.get("/assets/{name}")
def asset(name: str):
    if name not in ("awbook.ico", "awbook_192.png", "awbook_512.png"):
        raise HTTPException(404)
    return FileResponse(ASSETS / name)


@app.get("/manifest.webmanifest")
def manifest():
    from fastapi.responses import JSONResponse
    return JSONResponse({
        "name": "AWBook", "short_name": "AWBook", "start_url": "/", "display": "standalone",
        "background_color": "#ffffff", "theme_color": "#0b3d91",
        "icons": [{"src": "/assets/awbook_192.png", "sizes": "192x192", "type": "image/png"},
                  {"src": "/assets/awbook_512.png", "sizes": "512x512", "type": "image/png"}]},
        media_type="application/manifest+json")


# Nhịp tim từ giao diện: bản desktop tự tắt khi cửa sổ giao diện đóng (không còn nhịp tim).
PING = {"last": 0.0}


@app.get("/api/ping")
def ping():
    import time
    PING["last"] = time.time()
    return {"ok": True, "app": "hoadon"}


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


@app.get("/api/browser/diag")
def browser_diag():
    from .gdt.browser import BROWSER
    return BROWSER.diagnose()


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
    easy_captcha: bool = True


@app.post("/api/auto/start")
def auto_start(body: AutoIn):
    from .auto_download import KINDS, busy, start
    if busy():
        raise HTTPException(409, "Đang có một lượt tải chạy (hoặc đang tải nhiều công ty), hãy chờ xong hoặc bấm Dừng")
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


@app.get("/api/auto-running")
def auto_running():
    """Lượt đọc/tải đang chạy (kể cả đọc nhanh cho dự báo VAT do AI gọi) – để giao diện hiện ô captcha."""
    from . import auto_download as ad
    job = ad.running_job()
    return {"id": job.id, **job.public()} if job else {}


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


# ---------- tải nhiều công ty: mỗi "trang HĐĐT" là 1 cửa sổ Chrome riêng, đăng nhập 1 công ty ----------
SLOT_JOBS: dict[int, str] = {}     # số trang -> lượt tải gần nhất


def _slot_job(n: int):
    from .auto_download import JOBS
    return JOBS.get(SLOT_JOBS.get(n, ""))


@app.get("/api/slots")
def slots_status(ns: str = "1,2"):
    from .gdt.browser import BROWSER
    nums = sorted({int(x) for x in ns.split(",") if x.strip().isdigit() and 0 < int(x) <= 30}) or [1]
    st = BROWSER.slots_status(nums)
    out = []
    for n in nums:
        job = _slot_job(n)
        out.append({"n": n, **st.get(n, {}), "job": job.public() if job else None})
    return {"items": out}


class SlotOpenIn(BaseModel):
    user: str = ""      # tài khoản đã lưu (tùy chọn) -> điền sẵn tài khoản + mật khẩu


@app.post("/api/slots/{n}/open")
def slot_open(n: int, body: SlotOpenIn):
    from .gdt import credentials
    from .gdt.browser import BROWSER, BrowserError
    try:
        st = BROWSER.open_slot(n)
    except BrowserError as e:
        raise HTTPException(500, str(e))
    msg = ""
    if body.user and not st.get("logged_in"):
        creds = credentials.load(body.user)
        if creds:
            try:
                BROWSER.prefill_login(*creds, slot=n)
                msg = "Đã điền sẵn tài khoản + mật khẩu – nhập captcha rồi bấm Đăng nhập trên cửa sổ Chrome"
            except Exception as e:  # noqa: BLE001
                msg = f"Chưa điền sẵn được tài khoản: {e}"
        else:
            msg = "Chưa lưu mật khẩu công ty này – tự nhập trên cửa sổ Chrome"
    return {**st, "msg": msg}


@app.post("/api/slots/{n}/close")
def slot_close(n: int):
    from .gdt.browser import BROWSER
    job = _slot_job(n)
    if job and job.status == "running":
        raise HTTPException(409, "Trang này đang tải – bấm Dừng trước")
    BROWSER.close_slot(n)
    return {"ok": True}


def _start_slot(n: int, body: AutoIn, kinds: list[str], taken: set[str]) -> dict:
    from .auto_download import running_job, start
    from .gdt.browser import BROWSER
    job = _slot_job(n)
    if job and job.status == "running":
        return {"n": n, "ok": False, "msg": "Đang tải"}
    if n == 1 and running_job():
        return {"n": n, "ok": False, "msg": "Trang Tải hóa đơn đang tải bằng cửa sổ này – chờ xong"}
    st = BROWSER.slots_status([n]).get(n, {})
    if not st.get("logged_in"):
        return {"n": n, "ok": False, "msg": "Chưa đăng nhập"}
    mst = str(st.get("mst") or "")
    if mst in taken:
        return {"n": n, "ok": False, "msg": f"MST {mst} đã tải ở trang khác"}
    token = BROWSER.slot_token(n)
    if not token:
        return {"n": n, "ok": False, "msg": "Chưa đăng nhập"}
    taken.add(mst)
    job = start(body.date_from, body.date_to, kinds, body.excel, body.xml, body.pdf, body.easy_captcha,
                token=token, user=mst, slot=n)
    SLOT_JOBS[n] = job.id
    return {"n": n, "ok": True, "mst": mst, "job_id": job.id}


class SlotsStartIn(AutoIn):
    slots: list[int]


@app.post("/api/slots/start")
def slots_start(body: SlotsStartIn):
    from .auto_download import JOBS, KINDS
    if body.date_from > body.date_to:
        raise HTTPException(400, "Từ ngày phải trước đến ngày")
    kinds = [k for k in KINDS if k in body.kinds]
    if not kinds:
        raise HTTPException(400, "Chọn ít nhất một loại hóa đơn")
    taken = {j.mst for j in JOBS.values() if j.status == "running" and j.mst}
    return {"items": [_start_slot(n, body, kinds, taken) for n in sorted(set(body.slots))]}


@app.post("/api/slots/stop")
def slots_stop():
    from .auto_download import JOBS
    for job in JOBS.values():
        if job.slot and job.status == "running":
            job.cancel = True
            job.submit_captcha(None)
    return {"ok": True}


# ---------- nơi lưu hóa đơn ----------
@app.get("/api/settings")
def settings_get():
    from . import config
    return {"output_dir": settings.output_dir, "default_dir": config._DEFAULT_OUT,
            "custom": config.custom_output_dir(), "from_env": bool(os.getenv("OUTPUT_DIR"))}


class OutDirIn(BaseModel):
    path: str | None = None  # None / rỗng = về mặc định


@app.post("/api/settings/output-dir")
def settings_output_dir(body: OutDirIn):
    from . import config
    from .auto_download import busy
    if busy():
        raise HTTPException(409, "Đang có một lượt tải chạy, hãy đợi xong hoặc bấm Dừng rồi mới đổi nơi lưu")
    try:
        config.set_output_dir((body.path or "").strip().strip('"') or None)
    except ValueError as e:
        raise HTTPException(400, str(e))
    return settings_get()


@app.post("/api/settings/pick-folder")
def settings_pick_folder():
    """Mở hộp chọn thư mục của Windows (hiện trên màn hình người dùng). Trả về thư mục đã chọn hoặc null nếu bấm Hủy."""
    import subprocess
    cur = settings.output_dir.replace("'", "''")
    ps = ("[Console]::OutputEncoding=[Text.Encoding]::UTF8; Add-Type -AssemblyName System.Windows.Forms; "
          "$f=New-Object System.Windows.Forms.Form; $f.TopMost=$true; $f.ShowInTaskbar=$false; "
          "$f.StartPosition='CenterScreen'; $f.Size=New-Object System.Drawing.Size(1,1); $f.Show(); $f.Activate(); "
          "$d=New-Object System.Windows.Forms.FolderBrowserDialog; "
          "$d.Description='Chọn thư mục lưu hóa đơn, bảng kê, XML, PDF'; $d.ShowNewFolderButton=$true; "
          f"if(Test-Path -LiteralPath '{cur}'){{$d.SelectedPath='{cur}'}}; "
          "if($d.ShowDialog($f) -eq 'OK'){ $d.SelectedPath }; $f.Close()")
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-STA", "-Command", ps], capture_output=True, timeout=600,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception as e:  # noqa: BLE001
        raise HTTPException(500, f"Không mở được hộp chọn thư mục: {e}")
    path = r.stdout.decode("utf-8", "ignore").strip().splitlines()
    return {"path": path[-1].strip() if path else None}


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
        dirs = [d.name for d in base.iterdir() if d.is_dir() and d.name[:1].isdigit()] if base.is_dir() else []
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


# ---------- dự báo thuế GTGT (tờ khai 01/GTGT) ----------
def _vat_mst(mst: str | None) -> str:
    from .gdt.browser import BROWSER
    from .vat import _msts
    mst = (mst or "").strip()
    if not mst:
        mst = str(BROWSER.status().get("mst") or "")
    if not mst:
        dirs = _msts()
        mst = dirs[0] if len(dirs) == 1 else ""
    if not mst:
        raise HTTPException(400, "Chọn MST cần dự báo")
    return mst


@app.get("/api/vat/msts")
def vat_msts():
    from .gdt.browser import BROWSER
    from .vat import _msts, get_cfg
    cur = str(BROWSER.status().get("mst") or "")
    return {"items": [{"mst": m, **{k: v for k, v in get_cfg(m).items() if k in ("kieu", "auto")}} for m in _msts()],
            "current": cur}


@app.get("/api/vat/info")
def vat_info(mst: str, kieu: Literal["M", "Q"], ky: int, nam: int):
    from . import vat
    mst = _vat_mst(mst)
    key = vat.period_key(kieu, ky, nam)
    cfg = vat.get_cfg(mst)
    return {"mst": mst, "cfg": {"kieu": cfg["kieu"], "auto": cfg["auto"]}, "inputs": cfg["inputs"].get(key, {}),
            "ct22": vat.suggest_ct22(mst, kieu, ky, nam), "result": vat.saved_result(mst, key),
            "history": vat.history(mst), "deadline": vat.deadline(kieu, ky, nam).isoformat(),
            "form": vat.form_for("TT89" if vat.uses_tt89(kieu, ky, nam) else "TT80"),
            "mau": "TT89" if vat.uses_tt89(kieu, ky, nam) else "TT80"}


class VatIn(BaseModel):
    mst: str | None = None
    kieu: Literal["M", "Q"] = "M"
    ky: int
    nam: int
    inputs: dict[str, float] = {}


@app.post("/api/vat/forecast")
def vat_forecast(body: VatIn):
    from . import vat
    if not (1 <= body.ky <= (4 if body.kieu == "Q" else 12)) or not 2020 <= body.nam <= 2100:
        raise HTTPException(400, "Kỳ tính thuế không hợp lệ")
    mst = _vat_mst(body.mst)
    return vat.forecast(mst, body.kieu, body.ky, body.nam, body.inputs, source="manual")


@app.post("/api/vat/quick-read")
def vat_quick_read(body: VatIn):
    """Đọc nhanh danh sách hóa đơn kỳ này trên cổng (Chrome đã đăng nhập đúng MST) rồi dự báo ngay."""
    from . import quick_read
    mst = _vat_mst(body.mst)
    st = quick_read.read_period(mst, body.kieu, body.ky, body.nam)
    if st.get("loi"):
        raise HTTPException(400, st["loi"])
    return st


@app.get("/api/vat/quick-read/{mst}")
def vat_quick_status(mst: str):
    from . import quick_read
    return quick_read.STATE.get(mst) or {}


class VatCfgIn(BaseModel):
    mst: str
    kieu: Literal["M", "Q"] = "M"
    auto: bool = True


@app.post("/api/vat/config")
def vat_config(body: VatCfgIn):
    from . import vat
    return vat.set_cfg(_vat_mst(body.mst), kieu=body.kieu, auto=body.auto)


class VatLichIn(BaseModel):
    bat: bool = True
    ngay: str | list[int] = ""
    gio: str = "08:00"


@app.get("/api/vat/lich")
def vat_lich_get():
    from . import vat_lich
    d = vat_lich.get()
    return {**d, "mo_ta": vat_lich.describe(d), "windows_ok": vat_lich.task_exists()}


@app.post("/api/vat/lich")
def vat_lich_save(body: VatLichIn):
    from . import vat_lich
    try:
        d = vat_lich.save(body.bat, body.ngay, body.gio)
    except ValueError as e:
        raise HTTPException(400, str(e)) from e
    return {**d, "mo_ta": vat_lich.describe(d), "windows_ok": vat_lich.task_exists()}


class VatClientIn(BaseModel):
    mst: str
    key: str
    sender: str = "ATAX"
    ten: str | None = None
    ten_ngan: str | None = None
    ngay: str | None = None  # YYYY-MM-DD; trống = ngày lập dự báo
    write: bool = True       # False: chỉ lấy thông tin tự điền + lời nhắn


@app.post("/api/vat/client-report")
def vat_client_report(body: VatClientIn):
    """Xuất file Excel 'Dự báo thuế GTGT' gửi khách hàng + lời nhắn gửi kèm, từ kết quả dự báo đã lập."""
    from datetime import date as _date

    from . import vat, vat_client
    if not re.fullmatch(r"\d{4}-(Q[1-4]|\d{2})", body.key or ""):
        raise HTTPException(400, "Kỳ không hợp lệ")
    res = vat.saved_result(_vat_mst(body.mst), body.key)
    if not res:
        raise HTTPException(404, "Chưa có kết quả dự báo kỳ này – bấm Dự báo VAT trước")
    try:
        ngay = _date.fromisoformat(body.ngay) if body.ngay else None
    except ValueError:
        raise HTTPException(400, "Ngày dự báo không hợp lệ")
    return vat_client.export(res, write=body.write, sender=body.sender or "ATAX", ten=body.ten or None,
                             ten_ngan=body.ten_ngan or None, ngay=ngay)


@app.get("/api/vat/print")
def vat_print(mst: str, key: str):
    """Mở bản HTML tờ khai để in (Ctrl+P)."""
    from fastapi.responses import HTMLResponse

    from . import vat
    if not re.fullmatch(r"\d{4}-(Q[1-4]|\d{2})", key or ""):
        raise HTTPException(400, "Kỳ không hợp lệ")
    res = vat.saved_result(_vat_mst(mst), key)
    if not res or not res.get("files", {}).get("html"):
        raise HTTPException(404, "Chưa có bản dự báo kỳ này")
    return HTMLResponse(Path(res["files"]["html"]).read_text(encoding="utf-8"))


@app.get("/api/notifications")
def notifications_get():
    from . import vat
    items = vat.notifications()
    return {"items": items[:60], "unread": sum(1 for n in items if not n.get("read"))}


class ReadIn(BaseModel):
    ids: list[str] | None = None  # None = đánh dấu tất cả đã đọc


@app.post("/api/notifications/read")
def notifications_read(body: ReadIn):
    from . import vat
    vat.mark_read(body.ids)
    return {"ok": True}


# ---------- trợ lý AI ----------
@app.get("/api/ai/status")
def ai_status():
    from . import assistant
    return assistant.status()


class AiKeyIn(BaseModel):
    provider: str = "gemini"
    key: str | None = None    # None = giữ khóa cũ; "" = xóa khóa đã lưu
    model: str | None = None


@app.post("/api/ai/key")
def ai_key(body: AiKeyIn):
    from . import assistant
    k = body.key.strip() if body.key is not None else None
    if k and body.provider == "claude" and not k.startswith("sk-ant-"):
        raise HTTPException(400, "Khóa API Anthropic phải bắt đầu bằng sk-ant-")
    try:
        assistant.configure(body.provider, k, body.model)
    except ValueError as e:
        raise HTTPException(400, str(e))
    return assistant.status()


@app.get("/api/ai/models")
def ai_models(provider: str):
    from . import assistant
    if provider not in assistant.PROVIDERS:
        raise HTTPException(400, "Nhà cung cấp AI không hợp lệ")
    try:
        return {"items": assistant.list_models(provider)}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(400, f"Không lấy được danh sách model: {e}")


class AiChatIn(BaseModel):
    chat_id: str | None = None
    text: str
    context: str = ""  # trang đang mở + trạng thái màn hình (giao diện gửi kèm)
    group: bool = False  # nhóm chat nhiều AI cùng soát câu trả lời


class AiBuiltinIn(BaseModel):
    provider: str = "gemini"
    key: str = ""


@app.post("/api/ai/builtin")
def ai_builtin(body: AiBuiltinIn):
    """Nhà phát hành gắn sẵn khóa vào bản đóng gói (chỉ ở bản chạy từ mã nguồn)."""
    from . import assistant
    from .config import FROZEN
    if FROZEN:
        raise HTTPException(403, "Bản đã đóng gói không đổi được khóa dựng sẵn")
    if body.provider not in assistant.PROVIDERS:
        raise HTTPException(400, "Nhà cung cấp AI không hợp lệ")
    assistant.save_builtin(body.provider, body.key.strip())
    return assistant.status()


@app.post("/api/ai/chat")
def ai_chat(body: AiChatIn):
    from . import assistant
    if not body.text.strip():
        raise HTTPException(400, "Nhập câu hỏi")
    try:
        fn = assistant.group_chat if body.group else assistant.chat
        r = fn(body.chat_id, body.text.strip()[:4000], body.context)
        r.pop("evidence", None)
        return r
    except RuntimeError as e:
        raise HTTPException(400, str(e))


# ---------- lấy PDF gốc từ trang tra cứu (không cần đăng nhập cổng) ----------
class LookupIn(BaseModel):
    url: str = ""
    code: str = ""
    seller_mst: str = ""
    khhdon: str = ""
    shdon: str = ""


@app.post("/api/pdf/lookup")
def pdf_lookup(body: LookupIn):
    from . import pdf_tools
    if not body.url.strip() and not body.code.strip():
        raise HTTPException(400, "Nhập link trang tra cứu và/hoặc mã tra cứu")
    return pdf_tools.lookup_one(body.url, body.code, body.seller_mst, body.khhdon, body.shdon)


class FillIn(BaseModel):
    mst: str | None = None
    month_from: str  # YYYY-MM
    month_to: str


@app.post("/api/pdf/fill-missing")
def pdf_fill(body: FillIn):
    from . import pdf_tools
    from .auto_download import month_list
    if any(j.status == "running" for j in pdf_tools.FILL_JOBS.values()):
        raise HTTPException(409, "Đang có một lượt lấy bù PDF chạy")
    try:
        months = [f"{a:%Y-%m}" for a, _ in month_list(body.month_from, body.month_to)]
    except ValueError:
        raise HTTPException(400, "Tháng không hợp lệ")
    if not months or len(months) > 36:
        raise HTTPException(400, "Chọn từ 1 đến 36 tháng")
    return pdf_tools.start_fill(_vat_mst(body.mst), months).public()


@app.get("/api/pdf/fill-missing/{job_id}")
def pdf_fill_status(job_id: str):
    from . import pdf_tools
    job = pdf_tools.FILL_JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Không có lượt này")
    return job.public()


@app.post("/api/pdf/fill-missing/{job_id}/stop")
def pdf_fill_stop(job_id: str):
    from . import pdf_tools
    job = pdf_tools.FILL_JOBS.get(job_id)
    if job:
        job.cancel = True
    return {"ok": True}


# ---------- chuyển hóa đơn PDF/XML vào file mẫu nhập khẩu MISA ----------
@app.post("/api/misa/convert")
async def misa_convert(files: list[UploadFile] = File(...), template: UploadFile | None = File(None),
                       direction: str = Form("purchase"), mst: str = Form(""), accounts: str = Form("{}")):
    import json as _json

    from . import misa_import
    if direction not in ("purchase", "sold", "auto"):
        raise HTTPException(400, "Chọn hóa đơn đầu vào, đầu ra hoặc tự nhận")
    payload = await _read_uploads(files)
    tpl = await template.read() if template is not None else b""
    tpl_name = (template.filename or "") if template is not None and tpl else ""
    if not tpl:
        if direction == "auto":
            raise HTTPException(400, "Chế độ tự nhận cần file mẫu của anh/chị (mẫu cơ bản chỉ có loại mua hoặc bán)")
        tpl = misa_import.basic_template(direction)
    try:
        misa_import.read_template(tpl)
        acc = _json.loads(accounts or "{}")
    except ValueError as e:
        raise HTTPException(400, str(e))
    return misa_import.start(payload, tpl, direction, mst.strip(), acc, tpl_name).public()


@app.get("/api/misa/{job_id}")
def misa_status(job_id: str):
    from . import misa_import
    job = misa_import.JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Không có lượt chuyển này")
    return job.public()


@app.post("/api/inventory/run")
async def inventory_run(opening: UploadFile = File(...), tpl: UploadFile | None = File(None),
                        buy_file: UploadFile | None = File(None), buy_inv: list[UploadFile] | None = File(None),
                        sell_file: UploadFile | None = File(None), sell_inv: list[UploadFile] | None = File(None),
                        ret_file: UploadFile | None = File(None), ret_inv: list[UploadFile] | None = File(None),
                        catalog_tpl: UploadFile | None = File(None), mst: str = Form(""), ky_text: str = Form(""),
                        ngay: str = Form(""), accounts: str = Form("{}"), ten_dn: str = Form("")):
    """Tổng hợp tồn kho theo mẫu MISA: đầu kỳ + mua vào + bán ra (+ hàng bán trả lại), ghép mã hàng tự động."""
    import json as _json

    from . import inventory

    async def one(f: UploadFile | None) -> tuple[str, bytes] | None:
        if f is None:
            return None
        data = await f.read()
        return (f.filename or "file", data) if data else None

    inp: dict = {"opening": await one(opening), "tpl": await one(tpl), "catalog_tpl": await one(catalog_tpl),
                 "mst": mst.strip(), "ky_text": ky_text.strip(), "ngay": ngay.strip(), "ten_dn": ten_dn.strip()}
    try:
        inp["acc"] = _json.loads(accounts or "{}")
    except ValueError:
        inp["acc"] = {}
    if not inp["opening"]:
        raise HTTPException(400, "Chọn file tồn kho đầu kỳ / danh mục hàng hóa")
    for key, f, invs in (("buy", buy_file, buy_inv), ("sell", sell_file, sell_inv), ("ret", ret_file, ret_inv)):
        src = await one(f)
        inv_files = await _read_uploads([x for x in (invs or []) if x is not None]) if invs else []
        inv_files = [(n, d) for n, d in inv_files if d]
        if src or inv_files:
            inp[key] = {"file": None if inv_files else src, "mau": src if inv_files else None, "invoices": inv_files}
    if not any(k in inp for k in ("buy", "sell")):
        raise HTTPException(400, "Cần file mua vào hoặc bán ra (hoặc hóa đơn để chuyển vào mẫu MISA)")
    try:
        inventory.read_opening(inp["opening"][1])
    except ValueError as e:
        raise HTTPException(400, str(e))
    return inventory.start(inp).public()


@app.get("/api/inventory/{job_id}")
def inventory_status(job_id: str):
    from . import inventory
    job = inventory.JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Không có lượt tổng hợp này")
    return job.public()


@app.get("/api/misa-template/{direction}")
def misa_template(direction: Literal["purchase", "sold"]):
    from fastapi.responses import Response

    from . import misa_import
    name = "Mau_MISA_MuaHang.xlsx" if direction == "purchase" else "Mau_MISA_BanHang.xlsx"
    return Response(misa_import.basic_template(direction),
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="{name}"'})


# ---------- ☰ Tiện ích: bảng điều khiển, hạn nộp, rủi ro HĐ, lãi lỗ, biên bản công nợ, định khoản, XML tờ khai ----------
def _d(s: str, default: date) -> date:
    try:
        return date.fromisoformat(str(s)[:10])
    except ValueError:
        return default


def _ti_range(tu: str, den: str) -> tuple[date, date]:
    t = date.today()
    d1 = _d(tu, date(t.year, t.month, 1))
    return d1, max(d1, _d(den, t))


def _xlsx(data: bytes, name: str):
    from fastapi.responses import Response
    return Response(data, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="{name}"'})


@app.get("/api/ti/dashboard")
def ti_dashboard():
    from . import tienich
    return {"rows": tienich.dashboard()}


@app.get("/api/ti/han-nop")
def ti_han_nop(mst: str = "", ahead: int = 60):
    from . import tienich
    rows = tienich.deadlines(mst, ahead=ahead) if mst else tienich.all_deadlines(ahead=ahead)
    return {"rows": rows}


class TiMarkIn(BaseModel):
    mst: str
    key: str
    done: bool = True


@app.post("/api/ti/han-nop/mark")
def ti_han_mark(body: TiMarkIn):
    from . import tienich
    return tienich.mark_done(body.mst, body.key, body.done)


class TiRangeIn(BaseModel):
    mst: str
    tu: str = ""
    den: str = ""


@app.post("/api/ti/rui-ro")
def ti_rui_ro(body: TiRangeIn):
    from . import tienich
    return tienich.risk_check(body.mst, *_ti_range(body.tu, body.den))


@app.get("/api/ti/rui-ro.xlsx")
def ti_rui_ro_xlsx(mst: str, tu: str = "", den: str = ""):
    from . import tienich
    r = tienich.risk_check(mst, *_ti_range(tu, den))
    return _xlsx(tienich.risk_excel(r), f"KiemTraRuiRo_{mst}_{r['tu']}_{r['den']}.xlsx")


class TiTextIn(BaseModel):
    text: str = ""


@app.get("/api/ti/mst-rui-ro")
def ti_mst_rui_ro():
    from . import tienich
    return {"list": tienich.risky_list()}


@app.post("/api/ti/mst-rui-ro")
def ti_mst_rui_ro_save(body: TiTextIn):
    from . import tienich
    return {"list": tienich.set_risky_list(body.text)}


class TiProfitIn(BaseModel):
    mst: str
    tu: str = ""
    den: str = ""
    chi_phi_khac: float = 0
    luong: float = 0
    nhom_hkd: str = "phan_phoi"
    thue_suat: float = 0


@app.post("/api/ti/lai-lo")
def ti_lai_lo(body: TiProfitIn):
    from . import tienich
    return tienich.profit(body.mst, *_ti_range(body.tu, body.den), body.chi_phi_khac, body.luong, body.nhom_hkd,
                          body.thue_suat or None)


@app.get("/api/ti/nhom-hkd")
def ti_nhom_hkd():
    from . import tienich
    return {"rows": [{"ma": k, "ten": v[0], "gtgt": v[1], "tncn": v[2]} for k, v in tienich.HKD_RATES.items()]}


def _ti_out(mst: str, sub: str) -> Path:
    p = Path(settings.output_dir) / (mst or "_chung") / sub
    p.mkdir(parents=True, exist_ok=True)
    return p


@app.get("/api/ti/tai")
def ti_tai(p: str):
    f = Path(p).resolve()
    base = Path(settings.output_dir).resolve()
    if base not in f.parents or not f.is_file():
        raise HTTPException(404, "Không tìm thấy file")
    return FileResponse(f, filename=f.name)


@app.post("/api/ti/bien-ban")
async def ti_bien_ban(file: UploadFile = File(...), loai: str = Form("KH"), mst: str = Form(""),
                      ngay: str = Form(""), chon: str = Form("")):
    from starlette.concurrency import run_in_threadpool

    from . import bank_smart, tienich
    data = await file.read()
    name = file.filename or "congno.xlsx"
    loai = "NCC" if loai.upper().startswith("N") else "KH"
    led = await run_in_threadpool(bank_smart.read_ledger, name, data, loai)
    if not led or not led.get("partners"):
        ps = await run_in_threadpool(bank_smart.read_partners, name, data, loai)
        if not ps:
            raise HTTPException(400, "Không đọc được danh sách công nợ trong file (cần cột Tên đối tượng và Số dư / Còn phải thu, trả)")
        led = {"partners": ps, "entries": []}
    own = tienich.company(mst) if mst else {}
    d = _d(ngay, date.today())
    pick = [x.strip() for x in chon.split("\n") if x.strip()] or None
    xb, msgs = await run_in_threadpool(tienich.debt_letters, own, led, loai, d, pick)
    out = _ti_out(mst, "BienBanDoiChieuCongNo") / f"BBDC_{'PhaiThu' if loai == 'KH' else 'PhaiTra'}_{d:%Y%m%d}.xlsx"
    out.write_bytes(xb)
    return {"file": str(out), "so_doi_tuong": len(msgs), "msgs": msgs,
            "co_chi_tiet": bool(led.get("entries"))}


class TiTkIn(BaseModel):
    mst: str
    tu: str = ""
    den: str = ""
    che_do: str = "TT200"
    mac_dinh: str = ""


@app.post("/api/ti/dinh-khoan")
def ti_dinh_khoan(body: TiTkIn):
    from . import tienich
    return {"rows": tienich.purchase_lines(body.mst, *_ti_range(body.tu, body.den), body.che_do, body.mac_dinh)}


class TiLearnIn(BaseModel):
    pairs: list[list[str]]


@app.post("/api/ti/dinh-khoan/hoc")
def ti_dinh_khoan_hoc(body: TiLearnIn):
    from . import tienich
    return {"n": tienich.learn_tk([(p[0], p[1]) for p in body.pairs if len(p) >= 2])}


@app.post("/api/ti/dinh-khoan/hoc-file")
async def ti_dinh_khoan_hoc_file(file: UploadFile = File(...)):
    from . import tienich
    try:
        return {"n": tienich.learn_tk_from_file(file.filename or "", await file.read())}
    except ValueError as e:
        raise HTTPException(400, str(e)) from e


@app.post("/api/ti/dinh-khoan.xlsx")
def ti_dinh_khoan_xlsx(body: TiTkIn):
    import io as _io

    from openpyxl import Workbook

    from . import tienich
    rows = tienich.purchase_lines(body.mst, *_ti_range(body.tu, body.den), body.che_do, body.mac_dinh)
    wb = Workbook()
    ws = wb.active
    ws.append(["Ngày", "Hóa đơn", "Người bán", "MST", "Tên hàng", "ĐVT", "SL", "Đơn giá", "Thành tiền", "Thuế", "TK Nợ", "Lý do"])
    for r in rows:
        ws.append([r["ngay"], r["hd"], r["nguoi_ban"], r["mst_ban"], r["ten"], r["dvt"], r["sl"], r["don_gia"],
                   r["thanh_tien"], r["thue"], r["tk"], r["ly_do"]])
    b = _io.BytesIO()
    wb.save(b)
    return _xlsx(b.getvalue(), f"DinhKhoan_MuaVao_{body.mst}.xlsx")


class TiXmlIn(BaseModel):
    mst: str
    key: str
    nguoi_ky: str = ""


@app.post("/api/ti/xml-tk")
def ti_xml_tk(body: TiXmlIn):
    from . import tienich
    try:
        f, note = tienich.vat_xml(body.mst, body.key, body.nguoi_ky)
    except ValueError as e:
        raise HTTPException(400, str(e)) from e
    return {"file": str(f), "ghi_chu": note}


@app.get("/api/ti/kiem-xuat/kho")
def ti_kiem_xuat_kho(mst: str):
    from . import kiemxuat
    f = kiemxuat.last_stock(mst)
    return {"file": str(f) if f else "", "ten": f.name.split("_", 2)[-1] if f else "",
            "luc": f.name[:15] if f else ""}


@app.post("/api/ti/kiem-xuat")
async def ti_kiem_xuat(mst: str = Form(...), tu: str = Form(""), den: str = Form(""), tin_nhan: str = Form(""),
                       dung_ai: bool = Form(True), dung_kho_cu: bool = Form(True),
                       ton_kho: UploadFile | None = File(None), yeu_cau: list[UploadFile] = File(default=[])):
    from datetime import datetime as _dt

    from starlette.concurrency import run_in_threadpool

    from . import kiemxuat
    files = [(f.filename or "file", await f.read()) for f in yeu_cau if f is not None]
    files = [(n, d) for n, d in files if d]
    if not files and not tin_nhan.strip():
        raise HTTPException(400, "Tải lên ảnh / PDF / Excel yêu cầu của khách hoặc dán tin nhắn")
    notes: list[str] = []
    stock = None
    if ton_kho is not None and ton_kho.filename:
        data = await ton_kho.read()
        if data:
            try:
                stock = await run_in_threadpool(kiemxuat.read_stock, data)
            except ValueError as e:
                raise HTTPException(400, str(e)) from e
            kiemxuat.save_stock(mst, ton_kho.filename, data)
            notes.append(f"Tồn kho: {ton_kho.filename} – {len(stock)} mặt hàng (đã lưu cho lần sau)")
    elif dung_kho_cu and kiemxuat.last_stock(mst):
        f = kiemxuat.last_stock(mst)
        try:
            stock = await run_in_threadpool(kiemxuat.read_stock, f.read_bytes())
            notes.append(f"Tồn kho: dùng file đã lưu {f.name} – {len(stock)} mặt hàng")
        except ValueError:
            stock = None
    try:
        req, n2 = await run_in_threadpool(kiemxuat.read_request, files, tin_nhan, dung_ai)
    except RuntimeError as e:
        raise HTTPException(400, str(e)) from e
    notes += n2
    if not req:
        raise HTTPException(400, "Không đọc được mặt hàng nào từ yêu cầu. " + " · ".join(notes))
    t = date.today()
    d1 = _d(tu, date(t.year - 1, t.month, 1))
    d2 = max(d1, _d(den, t))
    res = await run_in_threadpool(kiemxuat.check, mst, req, stock, d1, d2)
    out = _ti_out(mst, "KiemTraXuatHD") / f"KiemTraDauVao_{_dt.now():%Y%m%d_%H%M%S}.xlsx"
    out.write_bytes(await run_in_threadpool(kiemxuat.write_excel, res))
    return {**res, "file": str(out), "notes": notes, "loi_nhan": kiemxuat.reply_text(res)}


@app.get("/api/ti/ky-vat")
def ti_ky_vat(mst: str):
    from . import vat
    return {"rows": vat.history(mst)}


# ---------- đọc chứng từ: hóa đơn hàng loạt, sao kê ngân hàng ----------
def _doc_root(mst: str) -> tuple[str, Path]:
    from .gdt.browser import BROWSER
    base = Path(settings.output_dir)
    mst = (mst or "").strip() or (BROWSER.status().get("mst") or "")
    if not mst and base.is_dir():
        dirs = [d.name for d in base.iterdir() if d.is_dir() and d.name[:1].isdigit()]  # bỏ MauNhapKhau...
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


@app.post("/api/docs/bank-book")
async def docs_bank_book(stmt: list[UploadFile] = File(...), book: list[UploadFile] = File(...), mst: str = Form("")):
    """Đối chiếu sao kê gốc (PDF/Excel của ngân hàng) với sổ tiền gửi đã nhập trên phần mềm kế toán."""
    from datetime import datetime

    from starlette.concurrency import run_in_threadpool

    from . import bank, bank_book
    mst, root = _doc_root(mst)
    errors: list[str] = []
    txns: list[dict] = []
    for name, data in await _read_uploads(stmt):
        tx, err = bank.read_statement(name, data)
        txns += tx
        if err:
            errors.append(f"Sao kê {name}: {err}")
    if not txns:
        raise HTTPException(400, "Không đọc được giao dịch nào trong sao kê gốc. " + " | ".join(errors))
    entries: list[dict] = []
    binfo: dict = {}
    for name, data in await _read_uploads(book):
        e, info, err = bank_book.read_book(name, data)
        entries += e
        for k, v in info.items():
            binfo.setdefault(k, v)
        if err:
            errors.append(f"Sổ {name}: {err}")
    if not entries:
        raise HTTPException(400, "Không đọc được dòng nào trong sổ kế toán. " + " | ".join(errors))
    res = await run_in_threadpool(bank_book.reconcile, txns, entries)
    summ = bank_book.summary(res, txns, binfo)
    dest = bank_book.write_report(
        root / "DocChungTu" / f"DoiChieu_SaoKe_So_{res['d1']:%Y%m%d}_{res['d2']:%Y%m%d}_{datetime.now():%H%M%S}.xlsx",
        res, summ, errors)
    top = [{"trang_thai": r["trang_thai"], "ghi_chu": r["ghi_chu"], "chenh_lech": r["chenh_lech"],
            "ngay": str(min(x["ngay"] for x in r["sk"] + r["so"])),
            "sk": "; ".join(f"{x['co'] or -x['no']:,.0f} {x['mo_ta'][:60]}" for x in r["sk"]),
            "so": "; ".join(f"{x['vao'] or -x['ra']:,.0f} {x['so_ct']} {x['mo_ta'][:60]}" for x in r["so"])}
           for r in res["rows"] if not r["trang_thai"].startswith("Khớp") or r["trang_thai"] == "Khớp – lệch ngày"][:300]
    return {"mst": mst, "file": str(dest), "tong_hop": summ, "can_xu_ly": top, "so_sk": len(txns), "so_dong_so": len(entries),
            "errors": errors}


@app.post("/api/docs/bank")
async def docs_bank(files: list[UploadFile] = File(...), mst: str = Form(""), months_back: int = Form(12),
                    che_do: str = Form("TT200"), lay_cong: bool = Form(True), so_ct_bat_dau: int = Form(1),
                    phai_thu: list[UploadFile] | None = File(None), phai_tra: list[UploadFile] | None = File(None),
                    so_ban: list[UploadFile] | None = File(None), so_mua: list[UploadFile] | None = File(None),
                    so_thu_chi: list[UploadFile] | None = File(None), mau_thu: UploadFile | None = File(None),
                    mau_chi: UploadFile | None = File(None), dung_ai: bool = Form(True)):
    """Đọc sao kê, dò từng giao dịch với: hóa đơn HĐĐT (5 loại, lùi các tháng trước), chi tiết công nợ phải thu / phải
    trả (đối tượng + khoản đã thu / đã trả), sổ chi tiết bán hàng / mua hàng, sổ thu / chi tiền gửi đã nhập (bỏ dòng đã
    hạch toán + học cách kế toán đã ghi); AI cho dòng còn chưa rõ. Xuất đúng file mẫu nhập khẩu Phiếu thu / chi tiền gửi."""
    from datetime import datetime, timedelta

    from starlette.concurrency import run_in_threadpool

    from . import bank, bank_post, bank_smart, quick_read
    mst, root = _doc_root(mst)
    payload = await _read_uploads(files)
    txns, errors, info = [], [], {}
    for name, data in payload:
        tx, err = await run_in_threadpool(bank.read_statement, name, data)   # bản scan: OCR
        txns += tx
        if err:
            errors.append(f"{name}: {err}")
        if tx and not info.get("so_tk"):
            info = await run_in_threadpool(bank_post.statement_info, name, data)
    if not txns:
        raise HTTPException(400, "Không đọc được giao dịch nào. " + " | ".join(errors))
    info["chu_tk"] = re.split(r"\b(?:DIA CHI|ĐỊA CHỈ)\b", info.get("chu_tk", ""))[0].strip()
    txns.sort(key=lambda t: t["ngay"])
    d1, d2 = txns[0]["ngay"], txns[-1]["ngay"]
    months_back = max(0, min(int(months_back or 0), 36))
    start = (d1.replace(day=1) - timedelta(days=31 * months_back)).replace(day=1)
    sources, notes = [], []

    # 1) hóa đơn HĐĐT: dữ liệu trên máy (đủ 5 loại nếu đã tải / đọc nhanh) + lấy thêm từ cổng (loại không captcha)
    local = bank.invoices_for_range(root, start + timedelta(days=62), d2)
    portal, note = [], ""
    if lay_cong:
        portal, perr = await run_in_threadpool(bank_post.portal_invoices, root, start, d2, lambda m: None)
        if perr:
            errors.append(perr)
        note = f"Hóa đơn lấy thêm từ cổng HĐĐT {start:%m/%Y} – {d2:%m/%Y}: {len(portal)}"
    invoices = bank_post.merge_invoices(local, portal)
    by_kind: dict[str, int] = {}
    for i in invoices:
        k = i.get("loai") or i.get("huong") or "?"
        by_kind[k] = by_kind.get(k, 0) + 1
    sources.append(f"Hóa đơn HĐĐT {start:%m/%Y} – {d2:%m/%Y}: {len(invoices)} HĐ" +
                   (" (" + ", ".join(f"{k}: {v}" for k, v in sorted(by_kind.items())) + ")" if by_kind else ""))
    try:
        miss = [x for x in quick_read.coverage(root, bank_post._months(start, d2), invoices)
                if " CHƯA đọc:" in x or "THIẾU" in x]
        lab = sorted({re.search(r"Tháng (\d{2}/\d{4})", x).group(1) for x in miss if re.search(r"Tháng \d{2}/\d{4}", x)},
                     key=lambda m: m[3:] + m[:2])
        if lab:
            notes.append(f"{len(lab)} tháng chưa đọc đủ hóa đơn trên HĐĐT (nhất là máy tính tiền): {', '.join(lab)} – vào "
                         "Tải hóa đơn hoặc Dự báo VAT bấm ⚡ Đọc nhanh từ cổng các tháng này để dò chính xác hơn")
    except Exception:  # noqa: BLE001
        pass

    # 2) công nợ phải thu / phải trả: danh bạ đối tượng + khoản đã thu / đã trả (chi tiết công nợ MISA)
    partners, ledger = [], []
    for loai, ups in (("KH", phai_thu), ("NCC", phai_tra)):
        lab = "Phải thu (khách hàng)" if loai == "KH" else "Phải trả (nhà cung cấp)"
        for name, data in await _read_uploads(ups or []):
            try:
                led = bank_smart.read_ledger(name, data, loai)
                if led:
                    partners += led["partners"]
                    ledger += led["entries"]
                    sources.append(f"{lab} – {name}: {len(led['partners'])} đối tượng, {len(led['entries'])} dòng chứng từ")
                else:
                    ps = bank_smart.read_partners(name, data, loai)
                    partners += ps
                    sources.append(f"{lab} – {name}: {len(ps)} đối tượng")
            except Exception as e:  # noqa: BLE001
                errors.append(f"{name}: {e}")
    dirx = bank_smart.Directory(partners) if partners else None

    # 3) sổ thu / chi tiền gửi đã nhập: bỏ dòng đã hạch toán, học đối tượng / lý do / diễn giải kế toán đã dùng
    book = []
    for name, data in await _read_uploads(so_thu_chi or []):
        try:
            bb = bank_smart.read_bank_book(name, data)
            if bb:
                book += bb["entries"]
                sources.append(f"Sổ {'thu' if bb['huong'] == 'thu' else 'chi'} tiền gửi đã nhập – {name}: {len(bb['entries'])} chứng từ")
            else:
                errors.append(f"{name}: không nhận ra sổ thu / chi tiền gửi (cần cột Số chứng từ, Số tiền)")
        except Exception as e:  # noqa: BLE001
            errors.append(f"{name}: {e}")
    n_booked = bank_smart.mark_booked(txns, book, ledger) if (book or ledger) else 0
    if n_booked:
        sources.append(f"{n_booked} giao dịch đã có chứng từ trên sổ – không xuất lại vào file nhập")
    n_hoc = bank_smart.learn(txns) if n_booked else 0
    if n_hoc:
        sources.append(f"Gợi ý theo lịch sử hạch toán cho {n_hoc} giao dịch chưa hạch toán")
    n_mem = bank_smart.apply_memory(txns, bank_smart.load_memory(info.get("so_tk") or info.get("chu_tk", "")))
    if n_mem:
        sources.append(f"Gợi ý theo bộ nhớ các kỳ trước cho {n_mem} giao dịch")
    if dirx:
        own = info.get("chu_tk", "")
        for t in txns:
            t["dt"] = dirx.identify(t, own)
            hoc = t.get("hoc")
            if not t["dt"] and hoc and hoc["doi_tuong"]:
                p = dirx.by_partner("", hoc["doi_tuong"])
                t["dt"] = {"ma": (p or {}).get("ma", ""), "ten": hoc["doi_tuong"], "mst": (p or {}).get("mst", ""),
                           "loai": (p or {}).get("loai", "KH" if t.get("co") else "NCC"), "why": "theo lịch sử hạch toán",
                           "diem": 80, "nghi_ngo": []}

    # 4) ghép hóa đơn, sổ bán hàng (khoản thu), chứng từ mua hàng / công nợ phải trả (khoản chi)
    bank_post.allocate(invoices, txns)
    sales, purchases = [], []
    for name, data in await _read_uploads(so_ban or []):
        try:
            vs = bank_smart.read_sales(name, data)
            sales += vs
            sources.append(f"Sổ chi tiết bán hàng – {name}: {len(vs)} chứng từ")
        except Exception as e:  # noqa: BLE001
            errors.append(f"{name}: {e}")
    for name, data in await _read_uploads(so_mua or []):
        try:
            ps = bank_smart.read_purchases(name, data)
            purchases += ps
            sources.append(f"Sổ chi tiết mua hàng – {name}: {len(ps)} chứng từ")
        except Exception as e:  # noqa: BLE001
            errors.append(f"{name}: {e}")
    if sales:
        sources.append(f"Ghép thêm {bank_smart.apply_sales(txns, sales, dirx)} khoản thu theo sổ chi tiết bán hàng")
    docs = bank_smart.payable_docs(ledger, purchases)
    if docs:
        sources.append(f"Ghép thêm {bank_smart.apply_payables(txns, docs, dirx)} khoản chi theo chứng từ mua hàng / công nợ phải trả")
    for t in txns:                      # đối tượng lấy theo hóa đơn đã ghép (nếu chưa nhận diện được)
        if not t.get("dt") and t.get("alloc"):
            inv = t["alloc"][0]["inv"]
            pm, pn = bank_post.partner(inv)
            p = dirx.by_partner(pm, pn) if dirx else None
            t["dt"] = {"ma": (p or {}).get("ma") or pm, "ten": (p or {}).get("ten") or pn, "mst": pm,
                       "loai": "NCC" if inv.get("huong") == "purchase" else "KH", "why": "theo hóa đơn đã ghép",
                       "diem": 80, "nghi_ngo": []}

    # 5) định khoản lượt 1 (phí NH, lương, thuế...), AI cho dòng còn chưa rõ, định khoản lại
    start_no = max(1, int(so_ct_bat_dau or 1))
    bank_post.build_vouchers(txns, che_do, info.get("chu_tk", ""), start_no)
    if dung_ai:
        try:
            todo = [t for t in txns if not t.get("booked") and not t.get("hoc")]
            n_ai = await run_in_threadpool(bank_smart.ai_classify, todo, dirx, lambda m: errors.append(m))
            if n_ai:
                sources.append(f"AI gợi ý cho {n_ai} giao dịch chưa rõ (đều đánh dấu cần kiểm tra)")
        except Exception as e:  # noqa: BLE001
            errors.append(f"AI phân loại lỗi: {e}")
    vouchers = bank_post.build_vouchers(txns, che_do, info.get("chu_tk", ""), start_no)
    bank_smart.fill_codes(vouchers, dirx)
    to_import = bank_smart.prepare_export(vouchers, book)

    tag = f"{d1:%Y%m%d}_{d2:%Y%m%d}_{datetime.now():%H%M%S}"
    out_dir = root / "DocChungTu"
    tpl_dir = Path(settings.output_dir) / "MauNhapKhau"
    tpl_dir.mkdir(parents=True, exist_ok=True)
    codes = bank_post.load_codes(tpl_dir)
    dest = bank_post.write_report(out_dir / f"DoiChieu_SaoKe_{tag}.xlsx", txns, invoices, vouchers, info, errors, note)
    ket_qua = out_dir / f"KetQua_GhepThuChi_{tag}.xlsx"
    stat = bank_smart.write_result(ket_qua, vouchers, info, notes, sources, invoices, (d1, d2))
    acct = info.get("so_tk") or info.get("chu_tk", "")
    n_mem = bank_smart.save_memory(acct, txns)
    if n_mem:
        sources.append(f"Đã ghi nhớ {n_mem} cách hạch toán của tài khoản này cho các kỳ sau")
    addr = {p["ma"]: p.get("dia_chi", "") for p in partners if p.get("ma")}
    misa = []
    for up, loai, label in ((mau_thu, "thu", "PhieuThuTienGui"), (mau_chi, "chi", "PhieuChiTienGui")):
        vs = [v for v in to_import if v["loai"] == loai]
        if up is None or not vs:
            continue
        try:
            data, ext = await run_in_threadpool(bank_smart.fill_misa_template, await up.read(), up.filename or "", vs,
                                                info, addr)
            f = out_dir / f"{label}_{tag}{ext}"
            f.write_bytes(data)
            misa.append(f)
        except Exception as e:  # noqa: BLE001
            errors.append(f"Điền mẫu {up.filename}: {e}")
    if not misa:
        misa = bank_post.export_misa(tpl_dir, out_dir, tag, to_import, info, codes)
    easy = bank_post.export_easybooks(tpl_dir, out_dir, tag, to_import, info, codes)
    matched = sum(1 for t in txns if t.get("alloc"))
    try:  # tóm tắt cho Bảng điều khiển / Trợ lý AI
        import json as _json
        ket_qua.with_suffix(".json").write_text(_json.dumps({
            "tu": d1.isoformat(), "den": d2.isoformat(), "luc": datetime.now().isoformat(timespec="seconds"),
            "tai_khoan": info.get("so_tk", ""), "ngan_hang": info.get("ngan_hang", ""), "giao_dich": len(txns),
            "khop_hoa_don": matched, "stat": stat, "file": str(ket_qua)}, ensure_ascii=False, default=str), encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    return {"file": str(dest), "ket_qua": str(ket_qua), "from": d1.isoformat(), "to": d2.isoformat(), "txns": len(txns),
            "invoices": len(invoices), "matched": matched, "unmatched": len(txns) - matched,
            "unpaid": sum(1 for i in invoices if i["_con_lai"] > 1 and str(i.get("ngay_lap"))[:10] >= d1.isoformat()),
            "kiem_tra": sum(1 for v in vouchers if v.get("kiem_tra") and not v["tx"].get("booked")),
            "tien_vao": sum(t["co"] for t in txns), "tien_ra": sum(t["no"] for t in txns), "errors": errors,
            "note": note, "misa": [str(p) for p in misa], "easybooks": [str(p) for p in easy],
            "tai_khoan": info, "mau_nhap_khau": str(tpl_dir), "stat": stat, "sources": sources, "notes": notes,
            "doi_tuong": len(partners), "co_doi_tuong": sum(1 for t in txns if t.get("dt")),
            "da_co": n_booked, "can_nhap": len(to_import)}


# ---------- chạy tự động theo lịch ----------
@app.get("/api/schedule")
def schedule_get():
    from .scheduler import SCHEDULER
    return SCHEDULER.status()


class ScheduleIn(BaseModel):
    enabled: bool
    day: int = 5
    hour: int = 8
    minute: int = 0
    ky: Literal["truoc", "nay", "quy_truoc", "quy_nay", "thang", "quy", "khoang"] = "truoc"
    tu: str = ""      # YYYY-MM-DD (ky = thang / quy / khoang)
    den: str = ""
    kinds: list[str] = ["mua-vao", "mua-vao-khong-ma", "ban-ra"]


@app.post("/api/schedule")
def schedule_set(body: ScheduleIn):
    from .scheduler import SCHEDULER
    if not 1 <= body.day <= 31 or not 0 <= body.hour <= 23 or not 0 <= body.minute <= 59:
        raise HTTPException(400, "Ngày chạy 1–31, giờ 0–23, phút 0–59")
    if not body.kinds:
        raise HTTPException(400, "Chọn ít nhất một loại hóa đơn")
    if body.ky in ("thang", "quy", "khoang"):
        from datetime import date as _d
        try:
            d1, d2 = _d.fromisoformat(body.tu), _d.fromisoformat(body.den)
        except ValueError:
            raise HTTPException(400, "Chọn tháng / quý / khoảng ngày cần tải") from None
        if d1 > d2:
            raise HTTPException(400, "Từ ngày phải trước đến ngày")
    SCHEDULER.configure(body.enabled, body.day, body.hour, body.kinds, body.minute, body.ky, body.tu, body.den)
    return SCHEDULER.status()


@app.post("/api/schedule/run-now")
def schedule_run_now():
    """Bấm Chạy ngay: tải như lượt theo lịch (lần lượt các công ty đã lưu, cửa sổ tiến trình) nhưng không tính là
    lượt của tháng."""
    from .scheduler import SCHEDULER
    if SCHEDULER.busy():
        raise HTTPException(409, "Đang có lượt tải tự động chạy")
    SCHEDULER.trigger(scheduled=True, mark=False)
    return SCHEDULER.status()


@app.post("/api/schedule/run-scheduled")
def schedule_run_scheduled():
    """Windows gọi đúng giờ đã đặt khi phần mềm đang mở: chạy lượt theo lịch (có cửa sổ tiến trình)."""
    from .scheduler import SCHEDULER
    if not SCHEDULER.busy():
        SCHEDULER.trigger(scheduled=True)
    return SCHEDULER.status()


class SchedLoginIn(BaseModel):
    value: str | None = None     # captcha đăng nhập; trống = bỏ qua


@app.post("/api/schedule/skip")
def schedule_skip():
    from .scheduler import SCHEDULER
    return SCHEDULER.skip_company()


@app.post("/api/schedule/refill")
def schedule_refill():
    from .scheduler import SCHEDULER
    return SCHEDULER.refill()


class GdtAccountIn(BaseModel):
    username: str
    password: str | None = None   # sửa công ty đã có: để trống = giữ mật khẩu cũ
    ten: str = ""
    auto: bool = True


@app.get("/api/gdt-accounts")
def gdt_accounts():
    from .gdt import credentials
    return {"items": credentials.accounts()}          # không bao giờ trả mật khẩu


@app.post("/api/gdt-accounts")
def gdt_account_save(body: GdtAccountIn):
    from .gdt import credentials
    try:
        return {"items": credentials.save(body.username, body.password, body.ten, body.auto)}
    except (ValueError, OSError) as e:
        raise HTTPException(400, str(e)) from e


class GdtAutoIn(BaseModel):
    auto: bool


@app.post("/api/gdt-accounts/{username}/auto")
def gdt_account_auto(username: str, body: GdtAutoIn):
    from .gdt import credentials
    return {"items": credentials.set_auto(username, body.auto)}


@app.delete("/api/gdt-accounts/{username}")
def gdt_account_delete(username: str):
    from .gdt import credentials
    return {"items": credentials.delete(username)}


@app.post("/api/schedule/open-app")
def schedule_open_app():
    """Nút "Vào phần mềm kiểm tra" trên cửa sổ tiến trình."""
    from .scheduler import open_window
    open_window("/#tai")
    return {"ok": True}


@app.get("/tien-trinh", response_class=HTMLResponse)
def tien_trinh_page():
    return HTMLResponse((TEMPLATES / "tien_trinh.html").read_text(encoding="utf-8"),
                        headers={"Cache-Control": "no-store, no-cache, must-revalidate", "Pragma": "no-cache"})


@app.post("/api/open-folder")
def open_folder(path: str | None = None):
    import os
    target = Path(path or settings.output_dir)
    base = Path(settings.output_dir).resolve()
    if not target.resolve().is_relative_to(base):
        raise HTTPException(400, "Chỉ mở được thư mục lưu hóa đơn")
    if target.is_file():
        pass  # mở file kết quả (.xlsx, .xls, .pdf, .zip...)
    elif target.suffix.lower() in (".xlsx", ".xls", ".xlsm", ".csv", ".pdf", ".html", ".xml", ".zip", ".json"):
        raise HTTPException(404, "Không còn file này")
    else:
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
    # Chạy bằng module "app.main" đã import đúng cách (không dùng bản "__main__"): khi khởi động kiểu
    # runpy / python -c, các model dữ liệu (AutoIn...) trong "__main__" không phân giải được kiểu -> lỗi 500.
    from app.main import run as _run
    _run()
