"""Client cho API của cổng hoadondientu.gdt.gov.vn.

Cổng thuế không công bố tài liệu API chính thức; các endpoint dưới đây là endpoint
mà giao diện web của cổng đang gọi. Nếu cổng thay đổi, chỉ cần sửa QUERY_SETS / đường dẫn.

Captcha: phần mềm KHÔNG tự giải captcha. Ảnh captcha được hiển thị trên giao diện để
người dùng tự nhập, đúng như khi đăng nhập trên web của cổng thuế.
"""
from __future__ import annotations

import io
import time
import zipfile
from calendar import monthrange
from datetime import date
from typing import Any, Iterator

import httpx

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

# Mỗi nhóm truy vấn: endpoint danh sách + endpoint xuất XML + điều kiện lọc trạng thái xử lý.
#   ttxly==5: đã cấp mã | ttxly==6: CQT đã nhận không mã | ttxly==8: máy tính tiền (có mã)
# Đường dẫn đã đối chiếu với mã JS của cổng (11/2026): base https://hoadondientu.gdt.gov.vn/api,
#   /{query|sco-query}/invoices/{purchase|sold|detail|export-xml}
QUERY_SETS: dict[str, list[dict[str, Any]]] = {
    "purchase": [
        {"key": "purchase-5", "list": "/query/invoices/purchase",
         "xml": "/query/invoices/export-xml", "detail": "/query/invoices/detail", "filter": "ttxly==5"},
        {"key": "purchase-6", "list": "/query/invoices/purchase",
         "xml": "/query/invoices/export-xml", "detail": "/query/invoices/detail", "filter": "ttxly==6"},
        {"key": "purchase-sco", "list": "/sco-query/invoices/purchase",
         "xml": "/sco-query/invoices/export-xml", "detail": "/sco-query/invoices/detail",
         "filter": "ttxly==8"},
    ],
    "sold": [
        {"key": "sold", "list": "/query/invoices/sold",
         "xml": "/query/invoices/export-xml", "detail": "/query/invoices/detail", "filter": None},
        {"key": "sold-sco", "list": "/sco-query/invoices/sold",
         "xml": "/sco-query/invoices/export-xml", "detail": "/sco-query/invoices/detail",
         "filter": None},
    ],
}
QUERY_BY_KEY = {q["key"]: q for qs in QUERY_SETS.values() for q in qs}


class GDTError(Exception):
    pass


class GDTAuthError(GDTError):
    pass


def month_chunks(start: date, end: date) -> Iterator[tuple[date, date]]:
    """Cổng thuế chỉ cho tra cứu tối đa ~1 tháng mỗi lần -> chia theo tháng."""
    cur = start
    while cur <= end:
        last = date(cur.year, cur.month, monthrange(cur.year, cur.month)[1])
        yield cur, min(last, end)
        cur = date(last.year + (last.month == 12), last.month % 12 + 1, 1)


def invoice_params(inv: dict[str, Any]) -> dict[str, Any]:
    return {
        "nbmst": inv["nbmst"],
        "khhdon": inv["khhdon"],
        "shdon": inv["shdon"],
        "khmshdon": inv["khmshdon"],
    }


class GDTClient:
    def __init__(self, base_url: str, verify: bool = True, delay: float = 0.6):
        self.delay = delay
        self.token: str | None = None
        self.username: str | None = None
        self.http = httpx.Client(
            base_url=base_url,
            verify=verify,
            timeout=httpx.Timeout(60.0, connect=20.0),
            headers={"User-Agent": UA, "Accept": "application/json, text/plain, */*"},
            follow_redirects=True,
        )

    # ---------- low level ----------
    def _request(self, method: str, path: str, **kw) -> httpx.Response:
        headers = kw.pop("headers", {})
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"
        for attempt in range(3):
            try:
                r = self.http.request(method, path, headers=headers, **kw)
            except httpx.TransportError as e:
                if attempt == 2:
                    raise GDTError(f"Không kết nối được cổng thuế: {e}") from e
                time.sleep(2 * (attempt + 1))
                continue
            if r.status_code in (401, 403):
                raise GDTAuthError("Phiên đăng nhập cổng thuế đã hết hạn, hãy đăng nhập lại.")
            if r.status_code == 429 or r.status_code >= 500:
                if attempt == 2:
                    raise GDTError(f"Cổng thuế lỗi {r.status_code}: {r.text[:300]}")
                time.sleep(3 * (attempt + 1))
                continue
            if r.status_code >= 400:
                raise GDTError(f"Cổng thuế trả lỗi {r.status_code}: {r.text[:300]}")
            return r
        raise GDTError("Không gọi được cổng thuế")

    # ---------- auth ----------
    def get_captcha(self) -> dict[str, str]:
        data = self._request("GET", "/captcha").json()
        return {"key": data["key"], "svg": data["content"]}

    def login(self, username: str, password: str, captcha_key: str, captcha_value: str) -> str:
        r = self.http.post(
            "/security-taxpayer/authenticate",
            json={"username": username, "password": password, "ckey": captcha_key, "cvalue": captcha_value},
        )
        if r.status_code != 200:
            try:
                msg = r.json().get("message") or r.text
            except ValueError:
                msg = r.text
            raise GDTAuthError(f"Đăng nhập thất bại: {msg[:300]}")
        token = r.json().get("token")
        if not token:
            raise GDTAuthError("Đăng nhập thất bại: không nhận được token")
        self.token = token
        self.username = username
        return token

    def profile(self) -> dict[str, Any]:
        try:
            return self._request("GET", "/security-taxpayer/profile").json()
        except GDTError:
            return {}

    # ---------- invoices ----------
    def iter_invoices(
        self, direction: str, start: date, end: date, page_size: int = 50
    ) -> Iterator[dict[str, Any]]:
        if direction not in QUERY_SETS:
            raise ValueError("direction phải là 'purchase' hoặc 'sold'")
        for qs in QUERY_SETS[direction]:
            for s, e in month_chunks(start, end):
                search = f"tdlap=ge={s:%d/%m/%Y}T00:00:00;tdlap=le={e:%d/%m/%Y}T23:59:59"
                if qs["filter"]:
                    search += ";" + qs["filter"]
                state = None
                while True:
                    params: dict[str, Any] = {
                        "sort": "tdlap:desc",
                        "size": page_size,
                        "search": search,
                    }
                    if state:
                        params["state"] = state
                    data = self._request("GET", qs["list"], params=params).json()
                    rows = data.get("datas") or []
                    for row in rows:
                        row["_query"] = qs["key"]
                        yield row
                    state = data.get("state")
                    time.sleep(self.delay)
                    if not state or not rows:
                        break

    def get_detail(self, inv: dict[str, Any]) -> dict[str, Any]:
        qs = QUERY_BY_KEY.get(inv.get("_query") or "", QUERY_SETS["purchase"][0])
        return self._request("GET", qs["detail"], params=invoice_params(inv)).json()

    def export_xml(self, inv: dict[str, Any]) -> bytes:
        """Tải XML hóa đơn (cổng trả về file zip chứa invoice.xml)."""
        qs = QUERY_BY_KEY.get(inv.get("_query") or "", QUERY_SETS["purchase"][0])
        r = self._request("GET", qs["xml"], params=invoice_params(inv))
        time.sleep(self.delay)
        content = r.content
        if content[:2] == b"PK":
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                names = [n for n in zf.namelist() if n.lower().endswith(".xml")]
                if not names:
                    raise GDTError("File zip từ cổng thuế không có XML")
                names.sort(key=lambda n: (n.lower() != "invoice.xml", n))
                return zf.read(names[0])
        if content.lstrip()[:1] == b"<":
            return content
        raise GDTError("Cổng thuế không trả về XML (có thể hóa đơn chưa có dữ liệu XML)")

    def close(self) -> None:
        self.http.close()
