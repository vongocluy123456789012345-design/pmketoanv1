"""Connector viết riêng bằng Python cho NCC cần quy trình phức tạp hơn 1 URL.

Ví dụ:

    def fetch_misa(http, ref):
        r = http.get("https://.../api/...", params={"code": ref.lookup_code})
        if r.content[:5] == b"%PDF-":
            return PdfResult(r.content, str(r.url), "misa")
        return None

    CUSTOM_FETCHERS["misa"] = fetch_misa

Khóa là `code` của NCC trong providers.json. Hàm trả về PdfResult hoặc None.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Callable

import httpx

if TYPE_CHECKING:
    from . import InvoiceRef, PdfResult

CUSTOM_FETCHERS: dict[str, Callable[["httpx.Client", "InvoiceRef"], "PdfResult | None"]] = {}

# Header giống trình duyệt thật: một số trang NCC có tường lửa trả về rỗng nếu thiếu.
BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/130.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8",
}


def fetch_misa(_http: "httpx.Client", ref: "InvoiceRef") -> "PdfResult | None":
    """MISA meInvoice (đã kiểm chứng 11/2026).

    Trang https://www.meinvoice.vn/tra-cuu/ làm 2 bước:
      1. POST tra-cuu/GetInvoiceDataByTransactionID  (transactionID = mã tra cứu trong XML)
      2. GET  tra-cuu/DownloadHandler.ashx?Type=pdf&Code=<mã>   -> PDF gốc
    Chỉ hiện captcha khi tra sai nhiều lần.
    """
    from . import PdfResult

    code = (ref.lookup_code or "").strip()
    if not code:
        return None
    base = "https://www.meinvoice.vn/tra-cuu/"
    with httpx.Client(follow_redirects=True, timeout=httpx.Timeout(60.0, connect=15.0),
                      headers=BROWSER_HEADERS) as c:
        c.get(base)
        r = c.post(base + "GetInvoiceDataByTransactionID", data={"transactionID": code},
                   headers={"X-Requested-With": "XMLHttpRequest", "Referer": base,
                            "Accept": "application/json, text/javascript, */*; q=0.01"})
        try:
            d = r.json()
        except ValueError:
            raise RuntimeError(f"MISA trả về không phải JSON (HTTP {r.status_code})") from None
        if not d.get("success"):
            raise RuntimeError(f"MISA báo lỗi: {d.get('errorCode') or d.get('errors')}")
        x = c.get(base + "DownloadHandler.ashx", params={"Type": "pdf", "Code": code},
                  headers={"Referer": base})
        if x.content[:5] == b"%PDF-":
            return PdfResult(x.content, str(x.url), "misa")
    return None


CUSTOM_FETCHERS["misa"] = fetch_misa
