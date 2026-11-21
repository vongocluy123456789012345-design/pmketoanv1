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


def _net_error(err: str) -> str:
    """Lỗi mạng khi gọi cổng -> câu báo dễ hiểu cho người dùng."""
    if "Failed to fetch" in err or "NetworkError" in err:
        return ("Không gọi được cổng thuế (Failed to fetch): máy chủ hoadondientu.gdt.gov.vn đang ngắt kết nối – cổng quá "
                "tải hoặc tạm chặn do truy cập dồn dập. Chờ 10–15 phút, bấm F5 tải lại trang cổng trong Chrome (vào được "
                "và vẫn đăng nhập) rồi chạy lại; hóa đơn đã tải sẽ được bỏ qua.")
    if "AbortError" in err or "aborted" in err:
        return "Cổng thuế phản hồi quá lâu (hết thời gian chờ) – cổng đang chậm, thử lại sau ít phút."
    return f"Không gọi được cổng thuế: {err}"


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


def find_browsers() -> list[str]:
    """Mọi Chrome / Edge có trên máy (Chrome trước). Chrome không mở được cổng debug thì thử Edge."""
    cands = [settings.chrome_path] if settings.chrome_path else []
    for base in (os.environ.get("PROGRAMFILES"), os.environ.get("PROGRAMFILES(X86)"),
                 os.environ.get("LOCALAPPDATA")):
        if base:
            cands.append(str(Path(base) / "Google/Chrome/Application/chrome.exe"))
    for base in (os.environ.get("PROGRAMFILES(X86)"), os.environ.get("PROGRAMFILES"), os.environ.get("LOCALAPPDATA")):
        if base:
            cands.append(str(Path(base) / "Microsoft/Edge/Application/msedge.exe"))
    try:  # Chrome cài ở chỗ khác: lấy theo đăng ký của Windows (App Paths)
        import winreg
        for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            for exe in ("chrome.exe", "msedge.exe"):
                try:
                    with winreg.OpenKey(hive, rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe}") as k:
                        cands.append(str(winreg.QueryValue(k, None)).strip('"'))
                except OSError:
                    pass
    except ImportError:
        pass
    out: list[str] = []
    for c in cands:
        if c and Path(c).is_file() and c.lower() not in (x.lower() for x in out):
            out.append(c)
    return out


def find_browser() -> str:
    found = find_browsers()
    if found:
        return found[0]
    raise BrowserError("Không tìm thấy Chrome hoặc Edge trên máy. Cài Google Chrome (chrome.com) rồi mở lại phần mềm, "
                       "hoặc điền đường dẫn chrome.exe vào CHROME_PATH trong file .env cạnh AWBook.exe")


def _pw_error(e: BaseException) -> str:
    return ("Bộ điều khiển trình duyệt (Playwright) trong phần mềm không chạy được: "
            f"{str(e).splitlines()[0][:300] if str(e) else type(e).__name__}. Thường do phần mềm diệt virus / Windows "
            "Defender chặn hoặc xóa file _internal\\playwright\\driver\\node.exe – hãy thêm cả thư mục AWBook vào danh sách "
            "ngoại lệ (Exclusions) của diệt virus, giải nén lại bản đầy đủ rồi mở lại phần mềm.")


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


# Bắt đầu các lệnh fetch rồi TRẢ VỀ NGAY (không giữ luồng trình duyệt); kết quả đặt ở window.__awb[id].
_START_JS = """
({id, reqs, timeoutMs}) => {
  const W = window.__awb || (window.__awb = {});
  W[id] = {done: false};
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
  Promise.all(reqs.map(one)).then(out => { W[id] = {done: true, out}; });
  return true;
}
"""
_POLL_JS = """(id) => { const W = window.__awb; if (!W || !W[id]) return null;
  const r = W[id]; if (r.done) delete W[id]; return r; }"""


RATE_GAP = 0.12        # khoảng cách tối thiểu giữa 2 lần gọi cổng (mọi công ty cộng lại)
# Cổng chỉ nhận khoảng 5 lệnh cùng lúc từ 1 máy; gửi dồn 10–30 lệnh thì gần như tất cả bị trả 429 (đo 21/11/2026)
# -> mọi luồng (mọi công ty) dùng chung tối đa 4 lệnh đang chờ cổng trả lời.
PORTAL_CONC = 4
RATE_GAP_MAX = 2.5
BUSY_RETRIES = 5       # cổng báo quá tải (429) -> chờ rồi thử lại tối đa 5 lần


class BrowserWorker:
    def __init__(self) -> None:
        self._q: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._pw = None
        self._browser = None
        self._lock = threading.Lock()
        self.profile_name: str | None = None
        self.hidden = False
        self.start_error: BaseException | None = None
        self.last_error = ""
        self._tl = threading.local()          # token riêng của luồng (tải nhiều công ty cùng lúc)
        self._slots: dict[int, dict[str, Any]] = {}   # trang HĐĐT số 2, 3...: {"ctx", "page"} – cookie riêng từng cửa sổ
        # nhịp gọi cổng dùng chung mọi luồng: cổng giới hạn theo máy (HTTP 429) -> gặp 429 thì tất cả cùng giãn ra
        self._rate_lock = threading.Lock()
        self._next_slot = 0.0
        self._gap = RATE_GAP
        self._cool_until = 0.0
        self._inflight = 0
        self._conc = threading.Condition()

    # ---------- luồng riêng cho Playwright ----------
    def _ensure_thread(self) -> None:
        with self._lock:
            if self._thread and self._thread.is_alive():
                return
            self._thread = threading.Thread(target=self._loop, daemon=True, name="chrome-worker")
            self._thread.start()

    def _loop(self) -> None:
        try:
            from playwright.sync_api import sync_playwright
            self._pw = sync_playwright().start()
            self.start_error = None
        except BaseException as e:  # noqa: BLE001 - node.exe bị diệt virus chặn, thiếu file... -> báo ngay, không treo
            log.exception("Không khởi động được Playwright")
            self.start_error = e
            self.last_error = _pw_error(e)
            while not self._q.empty():               # trả lỗi cho các lệnh đang chờ
                _, box, ev = self._q.get()
                box["error"] = BrowserError(self.last_error)
                ev.set()
            return
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
        end = time.time() + timeout
        while not ev.wait(0.5):
            if self.start_error is not None and not (self._thread and self._thread.is_alive()):
                raise BrowserError(_pw_error(self.start_error))
            if time.time() > end:
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
            # trust_env=False: máy có cài proxy (công ty, VPN) thì vẫn gọi thẳng 127.0.0.1, không đi qua proxy
            return httpx.get(f"http://127.0.0.1:{settings.cdp_port}/json/version", timeout=2,
                             trust_env=False).status_code == 200
        except httpx.HTTPError:
            return False

    def _context(self):
        if not self._connected():
            raise BrowserError("Chưa mở trình duyệt. Bấm 'Mở trang hóa đơn điện tử' trước.")
        ctxs = self._browser.contexts
        if not ctxs:
            raise BrowserError("Trình duyệt đã đóng. Bấm 'Mở trang hóa đơn điện tử' lại.")
        mine = [s["ctx"] for s in self._slots.values()]
        main = [c for c in ctxs if c not in mine]   # cửa sổ cổng thuế chính, không phải các trang HĐĐT 2, 3...
        if not main:
            raise BrowserError("Trình duyệt đã đóng. Bấm 'Mở trang hóa đơn điện tử' lại.")
        return main[0]

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

    def _kill_profile(self) -> None:
        """Chrome cũ (do phần mềm mở lần trước, bị treo / mở không có cổng debug) đang giữ hồ sơ cổng thuế: Chrome
        mới sẽ chỉ chuyển sang cửa sổ cũ và cổng debug không bao giờ mở -> dừng hẳn trước khi mở lại."""
        try:
            import psutil
            key = str(PROFILE_DIR).lower()
            for proc in psutil.process_iter(["cmdline"]):
                if any(key in str(x).lower() for x in (proc.info["cmdline"] or [])):
                    proc.kill()
            time.sleep(1)
        except Exception:  # noqa: BLE001
            pass

    def _launch(self, hidden: bool) -> None:
        try:
            PROFILE_DIR.mkdir(parents=True, exist_ok=True)
            (PROFILE_DIR / ".ghi_thu").write_text("ok", encoding="utf-8")
        except OSError as e:
            raise BrowserError(f"Không ghi được thư mục {PROFILE_DIR} ({e}). Giải nén phần mềm vào thư mục có quyền ghi, "
                               "ví dụ C:\\AWBook hoặc Documents (không để trong C:\\Program Files, không chạy thẳng trong "
                               "file .zip).") from e
        browsers = find_browsers()
        if not browsers:
            find_browser()                               # báo lỗi không có Chrome / Edge
        self._kill_profile()
        errs = []
        for exe in browsers:
            # hidden: tải tự động theo lịch khi chưa mở phần mềm – Chrome chạy ẩn, dùng lại phiên đăng nhập đã lưu
            try:
                proc = subprocess.Popen([
                    exe, f"--remote-debugging-port={settings.cdp_port}", "--remote-allow-origins=*",
                    f"--user-data-dir={PROFILE_DIR}", "--no-first-run", "--no-default-browser-check",
                    *(["--headless=new", "--window-size=1366,900"] if hidden else []),
                    PORTAL + "/",
                ])
            except OSError as e:
                errs.append(f"{Path(exe).name}: không chạy được ({e})")
                continue
            for _ in range(180 if hidden else 70):     # lần đầu tạo hồ sơ mới / máy chậm: khởi động lâu
                if self._cdp_alive():
                    self.profile_name = Path(exe).stem
                    self.hidden = hidden
                    return
                time.sleep(0.5)
            code = proc.poll()
            errs.append(f"{Path(exe).name}: " + ("đã tự tắt ngay (mã " + str(code) + ")" if code is not None else
                                                 f"mở nhưng không bật cổng debug {settings.cdp_port}"))
            log.warning("Mở %s không kết nối được cổng debug %s (poll=%s)", exe, settings.cdp_port, code)
            self._kill_profile()
        raise BrowserError("Không mở được trình duyệt để kết nối cổng thuế – " + "; ".join(errs) + ". Cách xử lý: tắt hết "
                           "cửa sổ Chrome / Edge rồi bấm lại; nếu là máy công ty có chính sách chặn chế độ gỡ lỗi "
                           "(RemoteDebuggingAllowed) thì cài thêm Google Chrome bản thường; kiểm tra phần mềm diệt virus "
                           "có chặn AWBook không. Bấm 'Kiểm tra lỗi' để xem chi tiết.")

    def _open(self, hidden: bool = False) -> dict[str, Any]:
        if not self._connected():
            if not self._cdp_alive():
                self._launch(hidden)
            try:
                self._browser = self._pw.chromium.connect_over_cdp(f"http://127.0.0.1:{settings.cdp_port}", timeout=60000)
            except Exception as e:  # noqa: BLE001
                raise BrowserError(f"Trình duyệt đã mở nhưng phần mềm không nối vào được: {str(e).splitlines()[0][:300]}. "
                                   "Tắt hết cửa sổ Chrome do phần mềm mở rồi bấm lại.") from e
        page = self._portal_page()
        try:
            page.bring_to_front()
        except Exception:  # noqa: BLE001
            pass
        return self._status()

    def open(self, hidden: bool = False) -> dict[str, Any]:
        try:
            if not hidden and self.hidden:
                r = self.show_for_login()        # người dùng bấm mở cổng mà Chrome đang chạy ẩn -> hiện lên
            else:
                r = self.call(lambda: self._open(hidden), 240)
            self.last_error = ""
            return r
        except Exception as e:
            log.exception("Mở trình duyệt lỗi")
            self.last_error = str(e)
            raise

    def diagnose(self) -> dict[str, Any]:
        """Thông tin để tìm lỗi không mở được Chrome trên máy khách."""
        import sys
        drv = Path(getattr(sys, "_MEIPASS", ROOT)) / "playwright" / "driver" / "node.exe"
        port_owner = ""
        try:
            import psutil
            for c in psutil.net_connections("tcp"):
                if c.laddr and c.laddr.port == settings.cdp_port and c.status == "LISTEN":
                    port_owner = psutil.Process(c.pid).name() if c.pid else "?"
        except Exception:  # noqa: BLE001
            pass
        return {"trinh_duyet_tim_thay": find_browsers(), "playwright_driver": str(drv),
                "driver_con_file": drv.exists() if getattr(sys, "frozen", False) else True,
                "playwright_loi": _pw_error(self.start_error) if self.start_error else "",
                "cong_debug": settings.cdp_port, "cong_debug_dang_mo": self._cdp_alive(), "chuong_trinh_giu_cong": port_owner,
                "thu_muc_ho_so": str(PROFILE_DIR), "proxy": {k: v for k, v in os.environ.items() if "proxy" in k.lower() and k.lower() != "no_proxy"},
                "loi_gan_nhat": self.last_error, "thu_muc_phan_mem": str(ROOT)}

    def _shutdown(self) -> None:
        """Tắt hẳn Chrome cổng thuế do phần mềm mở (để mở lại ở chế độ khác: ẩn <-> hiện)."""
        try:
            if not self._connected() and self._cdp_alive():
                self._browser = self._pw.chromium.connect_over_cdp(f"http://127.0.0.1:{settings.cdp_port}")
            if self._connected():
                self._browser.new_browser_cdp_session().send("Browser.close")
        except Exception:  # noqa: BLE001
            pass
        self._browser = None
        for _ in range(40):
            if not self._cdp_alive():
                break
            time.sleep(0.25)
        else:   # không tắt được qua cổng debug: dừng hẳn các tiến trình Chrome dùng hồ sơ cổng thuế của phần mềm
            try:
                import psutil
                key = str(PROFILE_DIR).lower()
                for proc in psutil.process_iter(["cmdline"]):
                    if any(key in str(x).lower() for x in (proc.info["cmdline"] or [])):
                        proc.kill()
            except Exception:  # noqa: BLE001
                pass
        self.hidden = False
        time.sleep(1.5)

    def show_for_login(self) -> dict[str, Any]:
        """Chrome đang chạy ẩn mà cần đăng nhập: mở lại thành cửa sổ thường để người dùng tự đăng nhập."""
        def run() -> dict[str, Any]:
            if getattr(self, "hidden", False):
                self._shutdown()
            return self._open(False)
        return self.call(run, 120)

    def close(self) -> None:
        try:
            self.call(self._shutdown, 30)
        except Exception:  # noqa: BLE001
            pass

    # ---------- đăng nhập: người dùng tự gõ captcha và tự bấm "Đăng nhập" trên trang cổng ----------
    def prefill_login(self, username: str, password: str, slot: int = 0) -> bool:
        """Mở form đăng nhập trên trang cổng (cửa sổ Chrome thật), điền sẵn tài khoản + mật khẩu đã lưu và đặt con trỏ
        vào ô captcha. Người dùng tự nhập captcha và tự bấm Đăng nhập."""
        def run() -> bool:
            page = self._slot_page(slot, create=True)
            try:
                page.bring_to_front()
            except Exception:  # noqa: BLE001
                pass
            if not page.locator("#username").count():
                page.evaluate("""() => {
                  document.querySelectorAll('.ant-modal-close').forEach(b => b.click());
                  const el = [...document.querySelectorAll('*')].find(e => !e.children.length && (e.innerText || '').trim() === 'Đăng nhập');
                  if (el) el.click();
                }""")
                page.wait_for_selector("#username", timeout=15000)
            page.fill("#username", username)
            page.fill("#password", password)
            page.wait_for_timeout(500)            # form tự đặt con trỏ vào ô đầu -> chuyển sang ô captcha
            page.click("#cvalue")
            return True
        return bool(self.call(run, 60))

    def logout(self) -> None:
        """Đăng xuất phiên đang dùng (để đăng nhập công ty khác)."""
        def run() -> None:
            ctx = self._context()
            try:
                ctx.clear_cookies(name="jwt")
            except TypeError:                      # Playwright cũ: xóa hết cookie của cổng
                ctx.clear_cookies()
            page = self._portal_page()
            page.goto(PORTAL + "/", wait_until="domcontentloaded", timeout=45000)
        self.call(run, 90)

    # ---------- token theo từng công ty ----------
    def use_token(self, token: str | None) -> None:
        """Luồng hiện tại gọi cổng bằng token này (của 1 công ty) thay vì phiên đang đăng nhập trên Chrome.
        Cổng chỉ xét header Authorization nên nhiều công ty dùng chung 1 tab cổng vẫn đúng dữ liệu từng công ty."""
        self._tl.token = token or None

    def thread_token(self) -> str | None:
        return getattr(self._tl, "token", None)

    def use_slot(self, slot: int) -> None:
        """Luồng hiện tại gọi cổng trên trang HĐĐT số `slot` (0/1 = cửa sổ cổng thuế chính)."""
        self._tl.slot = int(slot or 0)

    def thread_slot(self) -> int:
        return getattr(self._tl, "slot", 0)

    # ---------- nhiều trang HĐĐT: mỗi trang 1 cửa sổ Chrome riêng (cookie riêng) để đăng nhập 1 công ty ----------
    def _slot_ctx(self, n: int):
        if n <= 1:
            return self._context()
        s = self._slots.get(n)
        if not s:
            raise BrowserAuthError(f"Chưa mở trang HĐĐT số {n}")
        return s["ctx"]

    def _slot_page(self, n: int, create: bool = False):
        """Tab cổng thuế của trang HĐĐT số n. Chỉ gọi bên trong call()."""
        if n <= 1:
            return self._portal_page() if create else self._logged_page()
        s = self._slots.get(n)
        if s:
            try:
                if not s["page"].is_closed():
                    return s["page"]
                for p in s["ctx"].pages:
                    if not p.is_closed() and p.url.startswith(PORTAL):
                        s["page"] = p
                        return p
                if create:
                    p = s["ctx"].new_page()
                    p.goto(PORTAL + "/", wait_until="domcontentloaded", timeout=45000)
                    s["page"] = p
                    return p
            except Exception:  # noqa: BLE001 - cửa sổ / context đã bị đóng
                self._slots.pop(n, None)
                if create:
                    return self._slot_page(n, True)
        elif create:
            ctx = self._browser.new_context(no_viewport=True)
            p = ctx.new_page()
            self._slots[n] = {"ctx": ctx, "page": p}
            for i in range(3):
                try:
                    p.goto(PORTAL + "/", wait_until="domcontentloaded", timeout=45000)
                    break
                except Exception as e:  # noqa: BLE001
                    log.warning("Mở cổng thuế (trang %d) lỗi lần %d: %s", n, i + 1, str(e).splitlines()[0])
                    time.sleep(2 + 3 * i)
            return p
        raise BrowserAuthError(f"Cửa sổ trang HĐĐT số {n} đã đóng – bấm 'Mở trang HĐĐT {n}' và đăng nhập lại.")

    def _slot_jwt(self, n: int) -> str | None:
        if n <= 1:
            return self._jwt()
        for c in self._slot_ctx(n).cookies(PORTAL):
            if c.get("name") == "jwt" and c.get("value"):
                return c["value"]
        return None

    def _slot_status(self, n: int) -> dict[str, Any]:
        if n <= 1:
            st = self._status()
            if st.get("open") and self._portal_page(create=False) is None:
                st = {"open": False, "logged_in": False}
            return st
        if not self._connected() or n not in self._slots:
            return {"open": False, "logged_in": False}
        try:
            self._slot_page(n)
            token = self._slot_jwt(n)
        except Exception:  # noqa: BLE001
            return {"open": False, "logged_in": False}
        if not token:
            return {"open": True, "logged_in": False}
        info = self.token_info(token)
        if info["exp"] and info["exp"] < time.time():
            return {"open": True, "logged_in": False, "expired": True}
        return {"open": True, "logged_in": True, "mst": info["mst"], "expires": info["exp"]}

    def open_slot(self, n: int) -> dict[str, Any]:
        """Mở (hoặc đưa lên trước) cửa sổ Chrome của trang HĐĐT số n để đăng nhập 1 công ty."""
        if n <= 1:
            return self.open()

        def run() -> dict[str, Any]:
            if not self._connected():
                self._open(False)
            page = self._slot_page(n, create=True)
            try:
                page.bring_to_front()
            except Exception:  # noqa: BLE001
                pass
            return self._slot_status(n)
        try:
            r = self.call(run, 240)
            self.last_error = ""
            return r
        except Exception as e:
            self.last_error = str(e)
            raise

    def close_slot(self, n: int) -> None:
        def run() -> None:
            s = self._slots.pop(n, None)
            if s:
                try:
                    s["ctx"].close()
                except Exception:  # noqa: BLE001
                    pass
        if n > 1 and self._thread and self._thread.is_alive():
            self.call(run, 60)

    def slots_status(self, ns: list[int]) -> dict[int, dict[str, Any]]:
        if not (self._thread and self._thread.is_alive()):
            return {n: {"open": False, "logged_in": False} for n in ns}
        try:
            return self.call(lambda: {n: self._slot_status(n) for n in ns}, 30)
        except Exception as e:  # noqa: BLE001
            return {n: {"open": False, "logged_in": False, "error": str(e)} for n in ns}

    def slot_token(self, n: int) -> str | None:
        return self.call(lambda: self._slot_jwt(n), 30)

    def current_token(self) -> str | None:
        """Token của công ty đang đăng nhập trên Chrome."""
        return self.call(self._jwt, 20)

    @staticmethod
    def token_info(token: str | None) -> dict[str, Any]:
        p = _jwt_payload(token or "")
        return {"mst": str(p.get("sub") or p.get("username") or ""), "exp": p.get("exp") or 0}

    def _run_async(self, reqs: list[dict[str, Any]], timeout: float) -> list[dict[str, Any]]:
        """Chạy các fetch trong tab cổng mà không chiếm luồng trình duyệt suốt lúc chờ mạng."""
        rid = uuid.uuid4().hex
        args = {"id": rid, "reqs": reqs, "timeoutMs": int(timeout * 1000)}
        slot = self.thread_slot()
        lost = [{"status": 0, "ctype": "", "b64": "", "error": "Trang cổng vừa tải lại"}] * len(reqs)
        try:
            self.call(lambda: self._slot_page(slot).evaluate(_START_JS, args), 60)
        except BrowserAuthError:
            raise
        except BrowserError as e:
            if "không phản hồi" in str(e):
                raise
            return [{**x, "error": str(e)} for x in lost]
        except Exception as e:  # noqa: BLE001 - trang đang chuyển trang
            return [{**x, "error": str(e)} for x in lost]
        end, delay = time.time() + timeout + 20, 0.04
        while True:
            time.sleep(delay)
            delay = min(0.3, delay * 1.6)
            try:
                r = self.call(lambda: self._slot_page(slot).evaluate(_POLL_JS, rid), 60)
            except BrowserAuthError:
                raise
            except Exception as e:  # noqa: BLE001
                return [{**x, "error": str(e)} for x in lost]
            if r is None:
                return lost
            if r.get("done"):
                return r["out"]
            if time.time() > end:
                return [{**x, "error": "Quá thời gian chờ cổng trả lời"} for x in lost]

    def _take(self, k: int) -> None:
        with self._conc:
            while self._inflight and self._inflight + k > PORTAL_CONC:
                self._conc.wait(1)
            self._inflight += k

    def _give(self, k: int) -> None:
        with self._conc:
            self._inflight -= k
            self._conc.notify_all()

    def _send(self, reqs: list[dict[str, Any]], timeout: float) -> list[dict[str, Any]]:
        """Gửi tối đa PORTAL_CONC lệnh (giữ chỗ trong giới hạn chung của cả phần mềm)."""
        self._throttle()
        self._take(len(reqs))
        try:
            return self._run_async(reqs, timeout)
        finally:
            self._give(len(reqs))

    def _throttle(self) -> None:
        with self._rate_lock:
            now = time.time()
            at = max(now, self._next_slot, self._cool_until)
            self._next_slot = at + self._gap
        if at > now:
            time.sleep(at - now)

    def _rate_result(self, statuses: list[int]) -> bool:
        """Ghi nhận kết quả; trả True nếu cổng báo quá tải."""
        busy = any(x == 429 for x in statuses)
        with self._rate_lock:
            if busy:
                self._gap = min(RATE_GAP_MAX, self._gap * 2)
                self._cool_until = max(self._cool_until, time.time() + 2 + self._gap * 2)
            else:
                self._gap = max(RATE_GAP, self._gap * 0.9)
        return busy

    def _headers(self, action: str | None, token: str | None) -> dict[str, str]:
        h = {"Accept": "application/json, text/plain, */*", "Accept-Language": "vi", "request-id": str(uuid.uuid4()),
             "Action": encode_action(action), "End-Point": END_POINT}
        if token:
            h["Authorization"] = f"Bearer {token}"
        return h

    def _auth_token(self) -> str:
        slot = self.thread_slot()
        token = self.thread_token() or self.call(lambda: self._slot_jwt(slot), 30)
        if not token:
            raise BrowserAuthError("Chưa đăng nhập cổng thuế trên trình duyệt (hoặc phiên đã hết hạn).")
        return token

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
            return {"open": False, "logged_in": False, "error": self.last_error}
        try:
            st = self.call(self._status, 20)
            if not st.get("open") and self.last_error:
                st["error"] = self.last_error
            return st
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

        headers = self._headers(action, self._auth_token() if auth else None)
        headers.update(extra_headers or {})
        res: dict[str, Any] = {}
        attempt = busy_n = 0
        while True:
            res = self._send([{"url": url, "headers": headers}], timeout)[0]
            if not res.get("error") and self._rate_result([res["status"]]) and busy_n < BUSY_RETRIES:
                busy_n += 1
                time.sleep(2 * busy_n)
                headers["request-id"] = str(uuid.uuid4())
                continue
            if not res.get("error"):
                break
            attempt += 1
            if attempt >= retries:
                break
            # "Failed to fetch" thoáng qua (trang đang tải / mạng chập chờn / cổng ngắt kết nối) -> chờ rồi thử lại
            time.sleep(4 * attempt)
            try:
                slot = self.thread_slot()
                self.call(lambda: self._slot_page(slot).wait_for_load_state("domcontentloaded", timeout=15000), 30)
            except BrowserAuthError:
                raise
            except Exception:  # noqa: BLE001
                pass
            headers["request-id"] = str(uuid.uuid4())
        if res.get("error"):
            raise BrowserError(_net_error(res["error"]))
        r = FetchResult(res["status"], res["ctype"], base64.b64decode(res["b64"] or ""))
        if r.status == 401 and raise_401:
            raise BrowserAuthError("Phiên đăng nhập cổng thuế đã hết hạn, hãy đăng nhập lại trên trình duyệt.")
        return r

    def fetch_many(self, items: list[tuple[str, dict[str, Any] | str | None, str | None]],
                   timeout: float = 120) -> list[FetchResult]:
        """Gọi song song nhiều lệnh GET (path, query, action) trong trang cổng thuế."""
        token = self._auth_token()
        reqs = []
        for path, query, action in items:
            q = urlencode(query) if isinstance(query, dict) else query
            reqs.append({"url": API + path + (("?" + q) if q else ""), "headers": self._headers(action, token)})
        res: list[dict[str, Any]] = [{}] * len(reqs)
        for start in range(0, len(reqs), PORTAL_CONC):      # từng nhóm nhỏ, không gửi dồn cả lô
            todo = list(range(start, min(len(reqs), start + PORTAL_CONC)))
            for busy_n in range(BUSY_RETRIES + 1):
                got = self._send([reqs[i] for i in todo], timeout)
                for i, x in zip(todo, got):
                    res[i] = x
                if not self._rate_result([x.get("status", 0) for x in got]) or busy_n == BUSY_RETRIES:
                    break
                todo = [i for i in todo if res[i].get("status") == 429]   # chỉ gửi lại các lệnh bị báo quá tải
                for i in todo:
                    reqs[i]["headers"]["request-id"] = str(uuid.uuid4())
                time.sleep(2 * (busy_n + 1))
        out = [FetchResult(x["status"], x["ctype"], base64.b64decode(x["b64"] or "")) for x in res]
        if any(r.status == 401 for r in out):
            raise BrowserAuthError("Phiên đăng nhập cổng thuế đã hết hạn, hãy đăng nhập lại trên trình duyệt.")
        return out


BROWSER = BrowserWorker()
