"""Điều khiển Chrome thật để làm việc với cổng hoadondientu.gdt.gov.vn.

Quy trình:
  1. Phần mềm mở Chrome (hồ sơ riêng trong data/chrome-profile) tại trang cổng thuế.
  2. Người dùng TỰ đăng nhập trên Chrome (tự nhập mật khẩu + captcha).
  3. Phần mềm nối vào Chrome đó qua cổng debug (CDP) và gọi API của cổng NGAY TRONG TRANG,
     dùng đúng cookie/token của phiên đăng nhập và đúng các header mà web cổng thuế gửi
     (Authorization, Action, End-Point, request-id).

Playwright (sync) chỉ được dùng trên 1 luồng -> mọi thao tác đi qua BrowserWorker.call().
"""
from __future__ import annotations

import base64
import json
import logging
import os
import queue
import subprocess
import threading
import time
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable
from urllib.parse import quote, urlencode

import httpx

from ..config import ROOT, settings

log = logging.getLogger(__name__)
PORTAL = "https://hoadondientu.gdt.gov.vn"
API = PORTAL + "/api"
END_POINT = "/tra-cuu/tra-cuu-hoa-don"  # header End-Point = trang tra cứu hóa đơn của cổng
PROFILE_DIR = ROOT / "data" / "chrome-profile"


class BrowserError(Exception):
    pass


class BrowserAuthError(BrowserError):
    """Chưa đăng nhập hoặc phiên đăng nhập đã hết hạn."""


@dataclass
class FetchResult:
    status: int
    ctype: str
    body: bytes

    @property
    def ok(self) -> bool:
        return 200 <= self.status < 300

    def json(self) -> Any:
        return json.loads(self.body.decode("utf-8"))

    def message(self) -> str:
        """Thông báo lỗi dễ đọc từ phản hồi của cổng."""
        try:
            d = self.json()
            if isinstance(d, dict):
                return str(d.get("message") or d.get("error") or d)[:300]
        except Exception:  # noqa: BLE001
            pass
        return self.body[:300].decode("utf-8", "ignore")


def encode_action(action: str | None) -> str:
    """Giống encodeURIComponent của JS."""
    return quote(action or "", safe="-_.!~*'()")


def find_browser() -> str:
    cands = [settings.chrome_path] if settings.chrome_path else []
    for base in (os.environ.get("PROGRAMFILES"), os.environ.get("PROGRAMFILES(X86)"),
                 os.environ.get("LOCALAPPDATA")):
        if base:
            cands.append(str(Path(base) / "Google/Chrome/Application/chrome.exe"))
    for base in (os.environ.get("PROGRAMFILES(X86)"), os.environ.get("PROGRAMFILES")):
        if base:
            cands.append(str(Path(base) / "Microsoft/Edge/Application/msedge.exe"))
    for c in cands:
        if c and Path(c).is_file():
            return c
    raise BrowserError("Không tìm thấy Chrome hoặc Edge. Điền đường dẫn vào CHROME_PATH trong file .env")


def _jwt_payload(token: str) -> dict[str, Any]:
    try:
        part = token.split(".")[1]
        part += "=" * (-len(part) % 4)
        return json.loads(base64.urlsafe_b64decode(part))
    except Exception:  # noqa: BLE001
        return {}


# JS chạy trong trang cổng thuế: gọi API bằng fetch, trả nội dung dạng base64.
_FETCH_JS = """
async ({url, headers, timeoutMs}) => {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const r = await fetch(url, {headers, credentials: 'include', signal: ctrl.signal});
    const buf = new Uint8Array(await r.arrayBuffer());
    let s = '';
    for (let i = 0; i < buf.length; i += 0x8000) s += String.fromCharCode.apply(null, buf.subarray(i, i + 0x8000));
    return {status: r.status, ctype: r.headers.get('content-type') || '', b64: btoa(s)};
  } catch (e) {
    return {status: 0, ctype: '', b64: '', error: String(e)};
  } finally { clearTimeout(timer); }
}
"""


# Gửi nhiều lệnh song song trong 1 lần (dùng cho tải XML hàng loạt).
_FETCH_MANY_JS = """
async ({reqs, timeoutMs}) => {
  const one = async ({url, headers}) => {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), timeoutMs);
    try {
      const r = await fetch(url, {headers, credentials: 'include', signal: ctrl.signal});
      const buf = new Uint8Array(await r.arrayBuffer());
      let s = '';
      for (let i = 0; i < buf.length; i += 0x8000) s += String.fromCharCode.apply(null, buf.subarray(i, i + 0x8000));
      return {status: r.status, ctype: r.headers.get('content-type') || '', b64: btoa(s)};
    } catch (e) {
      return {status: 0, ctype: '', b64: '', error: String(e)};
    } finally { clearTimeout(timer); }
  };
  return Promise.all(reqs.map(one));
}
"""


class BrowserWorker:
    def __init__(self) -> None:
        self._q: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._pw = None
        self._browser = None
        self._lock = threading.Lock()
        self.profile_name: str | None = None

    # ---------- luồng riêng cho Playwright ----------
    def _ensure_thread(self) -> None:
        with self._lock:
            if self._thread and self._thread.is_alive():
                return
            self._thread = threading.Thread(target=self._loop, daemon=True, name="chrome-worker")
            self._thread.start()

    def _loop(self) -> None:
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        while True:
            fn, box, ev = self._q.get()
            try:
                box["result"] = fn()
            except BaseException as e:  # noqa: BLE001 - chuyển lỗi về luồng gọi
                box["error"] = e
            ev.set()

    def call(self, fn: Callable[[], Any], timeout: float = 180) -> Any:
        self._ensure_thread()
        box: dict[str, Any] = {}
        ev = threading.Event()
        self._q.put((fn, box, ev))
        if not ev.wait(timeout):
            raise BrowserError("Trình duyệt không phản hồi (quá thời gian chờ)")
        if "error" in box:
            raise box["error"]
        return box["result"]

    # ---------- mở / kết nối Chrome ----------
    def _connected(self) -> bool:
        try:
            return bool(self._browser and self._browser.is_connected())
        except Exception:  # noqa: BLE001
            return False

    def _cdp_alive(self) -> bool:
        try:
            return httpx.get(f"http://127.0.0.1:{settings.cdp_port}/json/version", timeout=2).status_code == 200
        except httpx.HTTPError:
            return False

    def _context(self):
        if not self._connected():
            raise BrowserError("Chưa mở trình duyệt. Bấm 'Mở trang hóa đơn điện tử' trước.")
        ctxs = self._browser.contexts
        if not ctxs:
            raise BrowserError("Trình duyệt đã đóng. Bấm 'Mở trang hóa đơn điện tử' lại.")
        return ctxs[0]

    def _portal_page(self, create: bool = True):
        ctx = self._context()
        for p in ctx.pages:
            try:
                if p.url.startswith(PORTAL):
                    return p
            except Exception:  # noqa: BLE001
                continue
        if not create:
            return None
        p = ctx.new_page()
        # cổng thuế hay ngắt kết nối (ERR_HTTP2_PROTOCOL_ERROR, timeout lúc quá tải): thử lại vài lần;
        # vẫn lỗi thì giữ tab cho người dùng tự bấm tải lại, không báo lỗi 500
        for i in range(3):
            try:
                p.goto(PORTAL + "/", wait_until="domcontentloaded", timeout=45000)
                break
            except Exception as e:  # noqa: BLE001
                log.warning("Mở cổng thuế lỗi (lần %d): %s", i + 1, str(e).splitlines()[0])
                time.sleep(2 + 3 * i)
        return p

    def _logged_page(self):
        """Tab cổng thuế người dùng đã đăng nhập. Không bao giờ mở tab/trang mới lúc đồng bộ
        (mở trang cổng lần 2 có thể làm mất phiên đăng nhập)."""
        p = self._portal_page(create=False)
        if p is None:
            raise BrowserAuthError("Không thấy tab cổng thuế đã đăng nhập (tab bị đóng?). "
                                   "Bấm 'Mở trang hóa đơn điện tử' và đăng nhập lại.")
        return p

    def aux_page(self, key: str):
        """Tab phụ dùng riêng (vd trang tra cứu EasyInvoice). Chỉ gọi bên trong call()."""
        aux = self.__dict__.setdefault("_aux", {})
        p = aux.get(key)
        if p is None or p.is_closed():
            p = self._context().new_page()
            aux[key] = p
        return p

    def _open(self) -> dict[str, Any]:
        if not self._connected():
            if not self._cdp_alive():
                exe = find_browser()
                self.profile_name = Path(exe).stem
                PROFILE_DIR.mkdir(parents=True, exist_ok=True)
                subprocess.Popen([
                    exe, f"--remote-debugging-port={settings.cdp_port}",
                    f"--user-data-dir={PROFILE_DIR}", "--no-first-run", "--no-default-browser-check",
                    PORTAL + "/",
                ])
                for _ in range(60):
                    if self._cdp_alive():
                        break
                    time.sleep(0.5)
                else:
                    raise BrowserError("Mở được trình duyệt nhưng không kết nối được (cổng debug). "
                                       "Hãy tắt hết cửa sổ Chrome do phần mềm mở rồi thử lại.")
            self._browser = self._pw.chromium.connect_over_cdp(f"http://127.0.0.1:{settings.cdp_port}")
        page = self._portal_page()
        try:
            page.bring_to_front()
        except Exception:  # noqa: BLE001
            pass
        return self._status()

    def open(self) -> dict[str, Any]:
        return self.call(self._open, 120)

    # ---------- trạng thái đăng nhập ----------
    def _jwt(self) -> str | None:
        for c in self._context().cookies(PORTAL):
            if c.get("name") == "jwt" and c.get("value"):
                return c["value"]
        return None

    def _status(self) -> dict[str, Any]:
        if not self._connected():
            return {"open": False, "logged_in": False}
        try:
            token = self._jwt()
        except BrowserError:
            return {"open": False, "logged_in": False}
        if not token:
            return {"open": True, "logged_in": False}
        p = _jwt_payload(token)
        exp = p.get("exp")
        if exp and exp < time.time():
            return {"open": True, "logged_in": False, "expired": True}
        return {"open": True, "logged_in": True, "mst": p.get("sub") or p.get("username"),
                "expires": exp}

    def status(self) -> dict[str, Any]:
        if not (self._thread and self._thread.is_alive()):
            return {"open": False, "logged_in": False}
        try:
            return self.call(self._status, 20)
        except Exception as e:  # noqa: BLE001
            return {"open": False, "logged_in": False, "error": str(e)}

    # ---------- gọi API cổng thuế trong trang ----------
    def fetch(self, path: str, query: str | dict[str, Any] | None = None, action: str | None = None,
              extra_headers: dict[str, str] | None = None, auth: bool = True,
              timeout: float = 120, raise_401: bool = True, retries: int = 3) -> FetchResult:
        """GET {API}{path}?{query}. query là chuỗi (giữ nguyên như web cổng) hoặc dict (sẽ được mã hóa)."""
        if isinstance(query, dict):
            query = urlencode(query)
        url = API + path + (("?" + query) if query else "")

        def run() -> FetchResult:
            page = self._logged_page()
            headers = {
                "Accept": "application/json, text/plain, */*",
                "Accept-Language": "vi",
                "request-id": str(uuid.uuid4()),
                "Action": encode_action(action),
                "End-Point": END_POINT,
            }
            if auth:
                token = self._jwt()
                if not token:
                    raise BrowserAuthError("Chưa đăng nhập cổng thuế trên trình duyệt (hoặc phiên đã hết hạn).")
                headers["Authorization"] = f"Bearer {token}"
            headers.update(extra_headers or {})
            args = {"url": url, "headers": headers, "timeoutMs": int(timeout * 1000)}
            res: dict[str, Any] = {}
            for attempt in range(retries):
                try:
                    res = page.evaluate(_FETCH_JS, args)
                except Exception as e:  # noqa: BLE001 - trang đang chuyển trang -> chờ rồi thử lại
                    if attempt == retries - 1:
                        raise BrowserError(f"Lỗi chạy lệnh trong trang: {e}") from e
                    res = {"error": str(e)}
                if not res.get("error"):
                    break
                # "Failed to fetch" thoáng qua (trang đang tải / mạng chập chờn) -> chờ rồi thử lại
                time.sleep(2 * (attempt + 1))
                try:
                    page.wait_for_load_state("domcontentloaded", timeout=15000)
                except Exception:  # noqa: BLE001
                    page = self._logged_page()
            if res.get("error"):
                raise BrowserError(f"Không gọi được cổng thuế: {res['error']}")
            return FetchResult(res["status"], res["ctype"], base64.b64decode(res["b64"] or ""))

        r = self.call(run, timeout + 30)
        if r.status == 401 and raise_401:
            raise BrowserAuthError("Phiên đăng nhập cổng thuế đã hết hạn, hãy đăng nhập lại trên trình duyệt.")
        return r

    def fetch_many(self, items: list[tuple[str, dict[str, Any] | str | None, str | None]],
                   timeout: float = 120) -> list[FetchResult]:
        """Gọi song song nhiều lệnh GET (path, query, action) trong trang cổng thuế."""
        def run() -> list[FetchResult]:
            page = self._logged_page()
            token = self._jwt()
            if not token:
                raise BrowserAuthError("Chưa đăng nhập cổng thuế trên trình duyệt (hoặc phiên đã hết hạn).")
            reqs = []
            for path, query, action in items:
                q = urlencode(query) if isinstance(query, dict) else query
                reqs.append({"url": API + path + (("?" + q) if q else ""), "headers": {
                    "Accept": "application/json, text/plain, */*", "Accept-Language": "vi",
                    "request-id": str(uuid.uuid4()), "Action": encode_action(action),
                    "End-Point": END_POINT, "Authorization": f"Bearer {token}"}})
            try:
                res = page.evaluate(_FETCH_MANY_JS, {"reqs": reqs, "timeoutMs": int(timeout * 1000)})
            except Exception as e:  # noqa: BLE001 - trang đang tải lại -> trả lỗi để bên gọi thử lại
                res = [{"status": 0, "ctype": "", "b64": "", "error": str(e)}] * len(reqs)
            return [FetchResult(x["status"], x["ctype"], base64.b64decode(x["b64"] or "")) for x in res]

        out = self.call(run, timeout + 30)
        if any(r.status == 401 for r in out):
            raise BrowserAuthError("Phiên đăng nhập cổng thuế đã hết hạn, hãy đăng nhập lại trên trình duyệt.")
        return out


BROWSER = BrowserWorker()
