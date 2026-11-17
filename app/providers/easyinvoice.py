"""Tải PDF gốc EasyInvoice (SoftDreams) qua trang tra cứu <MST>hd.easyinvoice.com.vn.

Trang tra cứu bắt nhập captcha cho MỖI lần tra cứu. Phần mềm KHÔNG tự giải captcha:
ảnh captcha được chụp lại và hiện cho người dùng nhập (qua hàm ask_captcha), sau đó phần mềm
tự điền mã tra cứu (Fkey), bấm "Tra cứu" rồi bấm "Tải PDF & đính kèm" ngay trên Chrome.
"""
from __future__ import annotations

import base64
import io
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Callable
from urllib.parse import urlparse

from ..gdt.browser import BROWSER


class CaptchaRejected(Exception):
    pass


def is_easyinvoice(url: str | None) -> bool:
    host = (urlparse(url or "").hostname or "").lower()
    return host.endswith("easyinvoice.com.vn") or host.endswith("easyinvoice.vn")


def search_url(lookup_url: str) -> str:
    u = urlparse(lookup_url)
    return f"{u.scheme or 'http'}://{u.hostname}/Search/Index"


def _open_search(url: str, fkey: str) -> bytes:
    """(Chạy trên luồng trình duyệt) mở trang tra cứu, điền mã, chụp ảnh captcha."""
    p = BROWSER.aux_page("easyinvoice")
    p.goto(url, wait_until="domcontentloaded", timeout=60000)
    p.fill("#iFkey", fkey)
    # Lấy ảnh captcha trực tiếp (tab có thể đang chạy nền nên không chụp màn hình được).
    # Lần tải này tạo mã mới trong phiên, đúng mã sẽ dùng khi gửi form.
    b64 = p.evaluate("""async () => { const r = await fetch('/Captcha/Show?_dc=' + Date.now(), {credentials: 'include'});
        const b = new Uint8Array(await r.arrayBuffer()); let s = '';
        for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
        document.querySelector('#captcha').src = 'data:image/png;base64,' + btoa(s); return btoa(s); }""")
    return base64.b64decode(b64)


def _submit_and_download(captcha: str) -> tuple[bytes, str]:
    """(Chạy trên luồng trình duyệt) gửi captcha, chờ kết quả, bấm tải PDF."""
    p = BROWSER.aux_page("easyinvoice")
    p.fill("#Capcha", captcha)
    with p.expect_navigation(wait_until="domcontentloaded", timeout=60000):
        p.click("#btnSubmit")
    btn = p.locator("button[name='downloadPdfAndFileAttach']")
    try:
        btn.first.wait_for(state="visible", timeout=20000)
    except Exception:  # noqa: BLE001 - không ra hóa đơn -> đọc thông báo lỗi của trang
        msg = p.evaluate("() => (document.querySelector('#msg') || {}).value || "
                         "(document.querySelector('.sweet-alert p') || {}).innerText || ''")
        raise CaptchaRejected(msg or "Không tra cứu được hóa đơn") from None
    with p.expect_download(timeout=90000) as dl:
        btn.first.click()
    d = dl.value
    with tempfile.TemporaryDirectory() as tmp:
        dest = Path(tmp) / (d.suggested_filename or "invoice.pdf")
        d.save_as(dest)
        return dest.read_bytes(), d.suggested_filename or ""


def _pdf_from(data: bytes) -> bytes | None:
    if data[:5] == b"%PDF-":
        return data
    if data[:2] == b"PK":  # zip gồm PDF + file đính kèm
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            pdfs = [n for n in zf.namelist() if n.lower().endswith(".pdf")]
            if pdfs:
                return zf.read(pdfs[0])
    return None


# ======================================================================================
# Hóa đơn BÁN RA của chính doanh nghiệp: tải hàng loạt qua trang quản lý <MST>.easyinvoice.com.vn
# (người dùng tự đăng nhập 1 lần trên Chrome của phần mềm, không cần captcha từng hóa đơn).
#   - Danh sách: /EInvoice?Pattern=..&FromDate=..&ToDate=..&Pagesize=500&page=N  -> mã nội bộ (invId)
#   - Tải PDF gốc: /EInvoice/DownloadPdf?invId=..&pattern=..
# ======================================================================================
MGMT_PAGE_SIZE = 500
MGMT_BATCH = 10

_GET_TEXT_JS = """async (url) => { const r = await fetch(url, {credentials: 'include'});
  return {status: r.status, url: r.url, text: await r.text()}; }"""

_GET_MANY_JS = """async (urls) => Promise.all(urls.map(async (u) => {
  try { const r = await fetch(u, {credentials: 'include'}); const b = new Uint8Array(await r.arrayBuffer());
    let s = ''; for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
    return {status: r.status, b64: btoa(s)};
  } catch (e) { return {status: 0, b64: '', error: String(e)}; } }))"""


def mgmt_origin(mst: str) -> str:
    return f"https://{mst}.easyinvoice.com.vn"


def _mgmt_page(mst: str, show: bool = False):
    p = BROWSER.aux_page("easy-mgmt")
    if not p.url.startswith(mgmt_origin(mst)):
        p.goto(mgmt_origin(mst) + "/", wait_until="domcontentloaded", timeout=60000)
    if show:
        p.bring_to_front()
    return p


def mgmt_logged_in_quiet(mst: str) -> bool:
    """Kiểm tra đã đăng nhập trang quản lý EasyInvoice chưa mà KHÔNG mở tab / không đổi trang nào
    (gửi 1 yêu cầu bằng cookie sẵn có của Chrome phần mềm)."""
    def run() -> bool:
        ctx = BROWSER._context()
        r = ctx.request.get(mgmt_origin(mst) + "/EInvoice/Index?Pagesize=10", max_redirects=0, timeout=20000)
        loc = r.headers.get("location") or ""
        return r.status == 200 and "/Account/LogOn" not in r.url and "LogOn" not in loc
    try:
        return BROWSER.call(run, 40)
    except Exception:  # noqa: BLE001
        return False


def mgmt_logged_in(mst: str, show_login: bool = False) -> bool:
    """Đã đăng nhập trang quản lý EasyInvoice trên Chrome của phần mềm chưa."""
    def run() -> bool:
        p = _mgmt_page(mst)
        r = p.evaluate(_GET_TEXT_JS, "/EInvoice/Index?Pagesize=10")
        ok = r["status"] == 200 and "/Account/LogOn" not in r["url"]
        if not ok and show_login:
            p.goto(mgmt_origin(mst) + "/", wait_until="domcontentloaded", timeout=60000)
            p.bring_to_front()
        return ok
    try:
        return BROWSER.call(run, 90)
    except Exception:  # noqa: BLE001
        return False


def mgmt_index(mst: str, pattern: str, d1, d2) -> dict[str, str]:
    """Số hóa đơn -> mã nội bộ EasyInvoice (invId) của 1 mẫu/ký hiệu trong khoảng ngày."""
    from urllib.parse import urlencode

    def run() -> dict[str, str]:
        p = _mgmt_page(mst)
        out: dict[str, str] = {}
        for page in range(1, 200):
            q = urlencode({"Pattern": pattern, "Status": -1, "FromDate": f"{d1:%d/%m/%Y}", "ToDate": f"{d2:%d/%m/%Y}",
                           "Store": -1, "typeInvoice": -1, "SortOption": 0, "Pagesize": MGMT_PAGE_SIZE,
                           "TCTCheckStatus": 0, "page": page})
            r = p.evaluate(_GET_TEXT_JS, "/EInvoice?" + q)
            if "/Account/LogOn" in r["url"]:
                raise RuntimeError("Phiên đăng nhập trang quản lý EasyInvoice đã hết hạn")
            rows = re.findall(r'class="no_hd_(\d+)" type="hidden" value="(\d+)"', r["text"])
            new = {n: i for i, n in rows if n != "0" and n not in out}
            out.update(new)
            if len(rows) < MGMT_PAGE_SIZE or not new:
                return out
        return out
    return BROWSER.call(run, 300)


def mgmt_download_many(mst: str, jobs: list[tuple[str, str]]) -> list[bytes | None]:
    """jobs: [(invId, pattern)] -> PDF gốc (hoặc None) theo đúng thứ tự, tải song song theo lô."""
    from urllib.parse import urlencode

    def run(batch: list[tuple[str, str]]) -> list[bytes | None]:
        p = _mgmt_page(mst)
        res = p.evaluate(_GET_MANY_JS, ["/EInvoice/DownloadPdf?" + urlencode({"invId": i, "pattern": pt})
                                        for i, pt in batch])
        out = []
        for x in res:
            data = base64.b64decode(x.get("b64") or "")
            out.append(data if data[:5] == b"%PDF-" else None)
        return out

    result: list[bytes | None] = []
    for k in range(0, len(jobs), MGMT_BATCH):
        result += BROWSER.call(lambda b=jobs[k:k + MGMT_BATCH]: run(b), 300)
    return result


def fetch(lookup_url: str, fkey: str, ask_captcha: Callable[[str, str], str], label: str) -> bytes:
    """Trả về PDF gốc. ask_captcha(image_data_uri, reason) -> mã người dùng nhập
    (hàm này raise nếu người dùng bỏ qua)."""
    url = search_url(lookup_url)
    last = ""
    for attempt in range(3):
        png = BROWSER.call(lambda: _open_search(url, fkey), 120)
        img = "data:image/png;base64," + base64.b64encode(png).decode()
        reason = f"EasyInvoice – {label}" + (f" (mã trước bị từ chối: {last})" if last else "")
        value = ask_captcha(img, reason)
        try:
            data, _name = BROWSER.call(lambda: _submit_and_download(value), 200)
        except CaptchaRejected as e:
            last = str(e)[:120]
            continue
        pdf = _pdf_from(data)
        if not pdf:
            raise RuntimeError("EasyInvoice trả về file không phải PDF")
        return pdf
    raise RuntimeError(f"EasyInvoice từ chối 3 lần: {last}")
