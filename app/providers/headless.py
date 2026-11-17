"""Lấy PDF hóa đơn "ngầm" từ BẤT KỲ trang tra cứu nào của nhà cung cấp HĐĐT – không cần đăng nhập cổng thuế.

Nhiều trang tra cứu là ứng dụng JavaScript (gõ mã tra cứu -> bấm Tra cứu -> trang gọi API rồi mới hiện nút
"Tải PDF"), nên tải bằng HTTP thường không thấy PDF. Ở đây phần mềm mở một Chrome/Edge CHẠY ẨN (headless,
hồ sơ trống, không đụng Chrome của người dùng) và làm như người thật:

  1. Mở link tra cứu (in trong XML) hoặc trang tra cứu của NCC.
  2. Tìm ô nhập mã tra cứu (và ô MST người bán nếu có), điền mã, bấm "Tra cứu".
  3. Bắt mọi PDF đi qua mạng: file PDF, file ZIP chứa PDF, JSON chứa PDF dạng base64, lượt tải (download),
     tab mới mở PDF / blob PDF.
  4. Chưa có thì bấm các nút "Tải PDF / Tải hóa đơn / Download / Xem hóa đơn / In"...
  5. Trang hiện hóa đơn dạng HTML (không có file PDF) thì in chính trang đó ra PDF.

Mọi PDF đều phải chứa đúng số hóa đơn + ký hiệu/MST người bán mới được nhận (tránh nhận nhầm file hướng dẫn).
Trang nào bắt buộc nhập captcha thì phần mềm DỪNG và báo NeedsCaptcha (không vượt captcha) – trừ khi được
gọi kèm hàm `ask`: khi đó ảnh captcha được chụp lại để NGƯỜI DÙNG nhập (phần mềm không tự đọc captcha).

Tăng tốc / thông minh:
  * Mỗi trang tra cứu (theo tên miền) dùng chung 1 phiên trình duyệt -> nhập captcha 1 lần, các hóa đơn sau
    của cùng trang dùng lại phiên đó (trang nào cho phép).
  * Tự học "công thức" tải: PDF lấy được qua link có chứa mã tra cứu -> lưu mẫu link (data/pdf_recipes.json);
    các hóa đơn sau của trang đó tải thẳng bằng HTTP, không cần mở trình duyệt.
  * Trang bố cục lạ không nhận ra ô mã tra cứu -> AI (nếu đã cấu hình) chỉ ra ô cần điền / nút cần bấm,
    kết quả nhớ theo tên miền nên mỗi trang chỉ hỏi AI 1 lần.
"""
from __future__ import annotations

import asyncio
import base64
import io
import re
import sys
import tempfile
import threading
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable
from urllib.parse import urlparse

CONCURRENCY = 8          # số trang tra cứu mở song song
HOST_LIMIT = 3           # tối đa số trang mở cùng lúc trên 1 trang NCC (tránh bị chặn vì quá tải)
ASK_TIMEOUT = 900        # giây – chờ người dùng nhập captcha
PAGE_TIMEOUT = 45_000    # ms – mở trang
STEP_WAIT = 4_000        # ms – chờ sau mỗi thao tác
TOTAL_TIMEOUT = 120      # giây cho 1 hóa đơn (mọi link)

# base64 của "%PDF-" bắt đầu bằng JVBERi0
B64_PDF = re.compile(r"JVBERi0[A-Za-z0-9+/=\\]{200,}")

CODE_WORDS = ("tra cứu", "tra cuu", "tracuu", "matracuu", "mã tra", "ma tra", "mã hóa đơn", "mã bảo mật", "ma bao mat",
              "fkey", "key", "code", "search", "lookup", "transaction", "secret", "mã nhận", "mã số bí mật",
              "mã", "keyword", "invoice", "reservation")
MST_WORDS = ("mst", "mã số thuế", "ma so thue", "taxcode", "tax code", "tax_code", "sellertax", "mstnban")
BAD_WORDS = ("captcha", "capcha", "email", "phone", "điện thoại", "password", "mật khẩu", "date", "ngày", "username",
             "tài khoản", "otp")
CAPTCHA_HINT = re.compile(r"captcha|capcha|recaptcha|hcaptcha|mã xác (nhận|thực)|mã kiểm tra|turnstile", re.I)
SUBMIT_TEXT = re.compile(r"^\s*(tra\s*c[ứu]u|tìm\s*kiếm|tìm|search|xem|kiểm tra|tra\s*cứu\s*hóa\s*đơn|lookup|ok|"
                         r"xác nhận)\s*$", re.I)
DOWNLOAD_TEXT = re.compile(r"t[ảa]i\s*(v[ềe]\s*)?(file\s*)?pdf|pdf|download|t[ảa]i\s*h[óo]a\s*[đd][ơo]n|t[ảa]i\s*v[ềe]|"
                           r"xem\s*h[óo]a\s*[đd][ơo]n|in\s*h[óo]a\s*[đd][ơo]n|b[ảa]n\s*th[ểe]\s*hi[ệe]n|"
                           r"xem\s*chi\s*ti[ếe]t|chuy[ểe]n\s*đ[ổo]i", re.I)
SKIP_TEXT = re.compile(r"xml|excel|hướng dẫn|huong dan|đăng nhập|login|đăng ký|zalo|facebook", re.I)


class NeedsCaptcha(Exception):
    pass


# ------------------------------------------------------------ công thức đã học theo từng trang (tên miền)
_recipe_lock = threading.Lock()
_recipes: dict[str, dict[str, Any]] | None = None


def _recipe_file() -> Path:
    from ..config import ROOT
    return ROOT / "data" / "pdf_recipes.json"


def recipes() -> dict[str, dict[str, Any]]:
    global _recipes
    with _recipe_lock:
        if _recipes is None:
            try:
                import json
                _recipes = json.loads(_recipe_file().read_text(encoding="utf-8"))
            except Exception:  # noqa: BLE001
                _recipes = {}
        return _recipes


def _save_recipes() -> None:
    import json
    with _recipe_lock:
        try:
            f = _recipe_file()
            f.parent.mkdir(parents=True, exist_ok=True)
            f.write_text(json.dumps(_recipes or {}, ensure_ascii=False, indent=1), encoding="utf-8")
        except Exception:  # noqa: BLE001
            pass


def _host_rec(host: str) -> dict[str, Any]:
    r = recipes()
    with _recipe_lock:
        return r.setdefault(host, {})


def learn_pdf_url(host: str, src: str, code: str | None, mst: str | None) -> None:
    """PDF tải được bằng link GET có chứa mã tra cứu -> ghi nhớ mẫu link cho cả trang."""
    from urllib.parse import quote
    if not code or len(code) < 6 or not src.startswith("http"):
        return
    tpl = None
    for c in (code, quote(code, safe=""), quote(code)):
        if c and c in src:
            tpl = src.replace(c, "{code}")
            break
    if not tpl:
        return
    if mst and len(mst) >= 10 and mst in tpl:
        tpl = tpl.replace(mst, "{mst}")
    rec = _host_rec(host)
    urls = rec.setdefault("pdf_urls", [])
    if tpl not in urls:
        urls.insert(0, tpl)
        del urls[3:]
        _save_recipes()


def recipe_fetch(urls: list[str], code: str | None, mst: str | None, accept: Callable[[bytes], bool]) -> "Found | None":
    """Tải thẳng bằng HTTP theo mẫu link đã học (nhanh, không mở trình duyệt)."""
    from urllib.parse import quote
    if not code:
        return None
    hosts = list(dict.fromkeys(host_of(u) for u in urls if u.startswith("http")))
    tpls = [t for h in hosts for t in (recipes().get(h, {}).get("pdf_urls") or [])]
    if not tpls:
        return None
    import httpx

    from .plugins import BROWSER_HEADERS
    with httpx.Client(timeout=httpx.Timeout(25.0, connect=10.0), follow_redirects=True, headers=BROWSER_HEADERS,
                      verify=False) as http:
        for t in tpls:
            if "{mst}" in t and not mst:
                continue
            u = t.replace("{code}", quote(code, safe="")).replace("{mst}", mst or "")
            try:
                pdf = _pdf_in(http.get(u).content)
            except Exception:  # noqa: BLE001
                continue
            if pdf and accept(pdf):
                return Found(pdf, u)
    return None


def _ai_pick(inputs: list[dict[str, Any]], btns: list[dict[str, Any]]) -> dict[str, Any] | None:
    """Hỏi AI: ô nào là ô mã tra cứu, ô nào là MST, nút nào là nút Tra cứu. Không dùng cho captcha."""
    import json
    try:
        from .. import assistant
        p = "gemini" if assistant.key_for("gemini") else assistant.provider()
        if not p or (p == "gemini" and not assistant.key_for("gemini")):
            return None
        desc = {"inputs": [{"i": i["i"], "mo_ta": i["text"][:140]} for i in inputs[:25]],
                "buttons": [{"i": b["i"], "chu": (b.get("text") or "")[:60], "class": (b.get("cls") or "")[:60]}
                            for b in btns[:60]]}
        prompt = ("Đây là danh sách ô nhập và nút bấm trên trang TRA CỨU HÓA ĐƠN ĐIỆN TỬ của một nhà cung cấp ở Việt Nam. "
                  "Chọn: ô để nhập MÃ TRA CỨU hóa đơn (code_input), ô nhập MST người bán nếu có (mst_input), "
                  "nút bấm để TRA CỨU (submit_button). KHÔNG chọn ô captcha. Chỉ trả về JSON dạng "
                  '{"code_input": số hoặc null, "mst_input": số hoặc null, "submit_button": số hoặc null}.\n'
                  + json.dumps(desc, ensure_ascii=False))
        model = "gemini-flash-lite-latest" if p == "gemini" else assistant.model_for(p)
        out = assistant._post_plain(p, model, [{"role": "user", "content": prompt}])
        m = re.search(r"\{.*\}", out or "", re.S)
        return json.loads(m.group(0)) if m else None
    except Exception:  # noqa: BLE001
        return None


# JS: tìm ảnh captcha (ảnh/canvas có tên captcha, hoặc ảnh nhỏ nằm sát ô nhập captcha) và ô nhập captcha.
_CAPTCHA_JS = r"""() => {
  const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
    return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const re = /captcha|capcha|kaptcha|securimage|verif|xacnhan|xac-nhan|validatecode|checkcode|imgcode|mã xác|ma xac|mã kiểm tra/i;
  const txt = el => [el.id, el.name, el.placeholder, el.getAttribute('aria-label'), el.title, el.alt, el.src,
                     typeof el.className === 'string' ? el.className : ''].join(' ');
  const inputs = [...document.querySelectorAll('input')].filter(el => vis(el) &&
      !['hidden','checkbox','radio','submit','button','file','image','password'].includes((el.type||'').toLowerCase()));
  let inp = inputs.find(el => re.test(txt(el) + ' ' + ((el.closest('div,td,li,p') || {}).innerText || '').slice(0, 60)));
  const imgs = [...document.querySelectorAll('img,canvas')].filter(el => { const r = el.getBoundingClientRect();
      return vis(el) && r.width >= 40 && r.width <= 420 && r.height >= 15 && r.height <= 160; });
  // chấm điểm ảnh: id/src chứa "captcha" (ảnh sinh động) > class > chỉ có alt (ảnh trang trí hay ghi alt="captcha")
  const dist = (p, q) => { const a = p.getBoundingClientRect(), b = q.getBoundingClientRect();
    return Math.hypot((a.x + a.width / 2) - (b.x + b.width / 2), (a.y + a.height / 2) - (b.y + b.height / 2)); };
  const score = el => (re.test(el.id || '') ? 4 : 0) + (re.test(el.src || '') ? 4 : 0) +
    (re.test(typeof el.className === 'string' ? el.className : '') ? 2 : 0) + (re.test(el.alt || '') ? 1 : 0);
  let img = null, bs = 0, bd = 1e9;
  for (const el of imgs) { const s = score(el); if (!s) continue; const d = inp ? dist(inp, el) : 0;
    if (s > bs || (s === bs && d < bd)) { img = el; bs = s; bd = d; } }
  if (!img && inp) { let best = null; bd = 1e9;
    for (const el of imgs) { const d = dist(inp, el); if (d < bd && d < 350) { bd = d; best = el; } } img = best; }
  if (img && !inp) { const b = img.getBoundingClientRect(); let bd = 1e9;
    for (const el of inputs) { if (el.value) continue; const a = el.getBoundingClientRect();
      const d = Math.hypot((a.x + a.width / 2) - (b.x + b.width / 2), (a.y + a.height / 2) - (b.y + b.height / 2));
      if (d < bd && d < 350) { bd = d; inp = el; } } }
  if (img) img.setAttribute('data-aw-capimg', '1');
  if (inp) inp.setAttribute('data-aw-capin', '1');
  return {img: !!img, input: !!inp};
}"""


@dataclass
class Found:
    content: bytes
    url: str
    printed: bool = False  # True = in từ trang hiển thị hóa đơn (trang NCC không có file PDF)


def _pdf_in(data: bytes) -> bytes | None:
    """PDF trong dữ liệu tải về: PDF trực tiếp, ZIP chứa PDF, hoặc JSON/text chứa PDF base64."""
    if not data:
        return None
    if data[:5] == b"%PDF-":
        return data
    if data[:2] == b"PK":
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                for n in zf.namelist():
                    if n.lower().endswith(".pdf"):
                        return zf.read(n)
        except zipfile.BadZipFile:
            return None
    if len(data) < 30_000_000:
        m = B64_PDF.search(data.decode("latin-1", "ignore"))
        if m:
            try:
                raw = base64.b64decode(m.group(0).replace("\\/", "/").replace("\\", ""), validate=False)
                if raw[:5] == b"%PDF-":
                    return raw
            except Exception:  # noqa: BLE001
                return None
    return None


# JS chạy trong trang: liệt kê ô nhập / nút bấm nhìn thấy được kèm "chữ mô tả" để chấm điểm.
_SCAN_JS = r"""() => {
  const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
    return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const lab = el => { let t = [el.id, el.name, el.placeholder, el.getAttribute('aria-label'), el.title,
      el.getAttribute('formcontrolname'), el.getAttribute('ng-model'), el.className].join(' ');
    if (el.id) { const l = document.querySelector('label[for="' + CSS.escape(el.id) + '"]'); if (l) t += ' ' + l.innerText; }
    const p = el.closest('div,td,li,p,form'); if (p) t += ' ' + (p.innerText || '').slice(0, 80);
    return t.toLowerCase(); };
  const inputs = [...document.querySelectorAll('input,textarea')].filter(el => vis(el) &&
      !['hidden','checkbox','radio','submit','button','file','image','password'].includes((el.type||'').toLowerCase()))
    .map((el, i) => { el.setAttribute('data-aw-in', i); const r = el.getBoundingClientRect();
      return {i, text: lab(el), value: el.value || '', x: r.x + r.width / 2, y: r.y + r.height / 2}; });
  const caps = [...document.querySelectorAll('img,canvas,iframe,div')].filter(el =>
      /captcha|capcha|recaptcha|hcaptcha|turnstile|challenges\.cloudflare/i.test((el.src||'') + ' ' + el.id + ' ' +
        (typeof el.className === 'string' ? el.className : '') + ' ' + (el.alt||'')) && (el.tagName === 'DIV' ? el.children.length > 0 && vis(el) : vis(el))).length;
  const btns = [...document.querySelectorAll('button,a,input[type=submit],input[type=button],[role=button],span.btn,div.btn,i.fa-search')]
    .filter(vis).map((el, i) => { el.setAttribute('data-aw-bt', i); const r = el.getBoundingClientRect();
      return {i, tag: el.tagName.toLowerCase(), x: r.x + r.width / 2, y: r.y + r.height / 2,
              text: ((el.innerText || el.value || '') + ' ' + (el.title || '') + ' ' + (el.getAttribute('aria-label') || '')).trim().slice(0, 80),
              href: el.getAttribute('href') || '', cls: (typeof el.className === 'string' ? el.className : '') + ' ' + (el.id || '')}; });
  const frames = [...document.querySelectorAll('iframe,embed,object')].map(el => el.src || el.data || '').filter(Boolean);
  return {inputs, caps, btns, frames, text: (document.body && document.body.innerText || '').slice(0, 20000)};
}"""


def _score_code_input(text: str) -> int:
    if any(b in text for b in BAD_WORDS) or any(w in text for w in MST_WORDS):
        return -1  # ô captcha / email / MST... không phải ô mã tra cứu
    s = 0
    for w in CODE_WORDS:
        if w in text:
            s += 3 if w in ("tra cứu", "tracuu", "matracuu", "mã tra", "fkey", "mã bảo mật", "transaction",
                            "reservation", "mã số bí mật") else 1
    return s


class HeadlessPdf:
    def __init__(self) -> None:
        self._loop: asyncio.AbstractEventLoop | None = None
        self._thread: threading.Thread | None = None
        self._pw = None
        self._browser = None
        self._sem: asyncio.Semaphore | None = None
        self._lock = threading.Lock()
        self._block_lock: asyncio.Lock | None = None
        self._ctxs: dict[str, Any] = {}             # tên miền -> phiên trình duyệt dùng chung (giữ cookie/captcha)
        self._host_sems: dict[str, asyncio.Semaphore] = {}

    # ------------------------------------------------------------ vòng lặp riêng
    def _ensure(self) -> asyncio.AbstractEventLoop:
        with self._lock:
            if self._loop and self._thread and self._thread.is_alive():
                return self._loop
            loop = asyncio.ProactorEventLoop() if sys.platform == "win32" else asyncio.new_event_loop()
            ready = threading.Event()

            def run() -> None:
                asyncio.set_event_loop(loop)
                self._sem = asyncio.Semaphore(CONCURRENCY)
                self._block_lock = asyncio.Lock()
                ready.set()
                loop.run_forever()

            self._thread = threading.Thread(target=run, daemon=True, name="headless-pdf")
            self._thread.start()
            ready.wait(10)
            self._loop = loop
            return loop

    def run(self, coro_fn: Callable[[], Any], timeout: float = TOTAL_TIMEOUT) -> Any:
        loop = self._ensure()
        fut = asyncio.run_coroutine_threadsafe(coro_fn(), loop)
        try:
            return fut.result(timeout)
        except TimeoutError:
            fut.cancel()
            raise RuntimeError("Trang tra cứu phản hồi quá lâu") from None

    async def _get_browser(self):
        async with self._block_lock:
            if self._browser and self._browser.is_connected():
                return self._browser
            from playwright.async_api import async_playwright

            from ..gdt.browser import find_browser
            if not self._pw:
                self._pw = await async_playwright().start()
            self._browser = await self._pw.chromium.launch(
                executable_path=find_browser(), headless=True,
                args=["--no-first-run", "--mute-audio", "--disable-extensions"])
            self._ctxs.clear()  # trình duyệt mới -> các phiên cũ không còn dùng được
            return self._browser

    async def _ctx_for(self, browser, host: str):
        async with self._block_lock:
            ctx = self._ctxs.get(host)
            if ctx is None:
                ctx = await browser.new_context(
                    accept_downloads=True, locale="vi-VN", ignore_https_errors=True,
                    viewport={"width": 1280, "height": 900},
                    user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                               "(KHTML, like Gecko) Chrome/130.0 Safari/537.36")
                self._ctxs[host] = ctx
            return ctx

    def _host_sem(self, host: str) -> asyncio.Semaphore:
        s = self._host_sems.get(host)
        if s is None:
            s = self._host_sems[host] = asyncio.Semaphore(HOST_LIMIT)
        return s

    # ------------------------------------------------------------ 1 hóa đơn
    def fetch(self, urls: list[str], code: str | None, seller_mst: str | None,
              accept: Callable[[bytes], bool], log: Callable[[str], None] | None = None,
              ask: Callable[[str, str], str] | None = None) -> Found | None:
        """Thử lần lượt các trang tra cứu. Trả về Found / None; raise NeedsCaptcha nếu trang bắt captcha.
        ask(ảnh_captcha_dataurl, lý_do) -> chữ người dùng nhập: có thì hiện captcha cho người dùng nhập."""
        if not ask:  # đường tắt: mẫu link đã học -> tải thẳng, không mở trình duyệt
            try:
                got = recipe_fetch(urls, code, seller_mst, accept)
                if got:
                    return got
            except Exception:  # noqa: BLE001
                pass
        # thời gian chờ tổng dài: gồm cả lúc xếp hàng chờ tới lượt (mỗi trang đã có giới hạn riêng)
        return self.run(lambda: self._fetch_all(urls, code, seller_mst, accept, log or (lambda _m: None), ask),
                        1800 + (ASK_TIMEOUT if ask else 0))

    async def _fetch_all(self, urls, code, seller_mst, accept, log, ask=None) -> Found | None:
        captcha = False
        async with self._sem:
            browser = await self._get_browser()
            for url in urls:
                try:
                    limit = ASK_TIMEOUT + TOTAL_TIMEOUT if ask else TOTAL_TIMEOUT / max(1, min(len(urls), 3))
                    async with self._host_sem(host_of(url)):
                        got = await asyncio.wait_for(self._one(browser, url, code, seller_mst, accept, log, ask),
                                                     limit)
                except NeedsCaptcha:
                    captcha = True
                    log(f"{url}: trang yêu cầu captcha")
                    continue
                except asyncio.TimeoutError:
                    log(f"{url}: quá thời gian")
                    continue
                except Exception as e:  # noqa: BLE001 - trang lỗi -> thử link khác
                    log(f"{url}: {str(e)[:160]}")
                    continue
                if got:
                    return got
        if captcha:
            raise NeedsCaptcha("Trang tra cứu yêu cầu nhập captcha")
        return None

    async def _one(self, browser, url: str, code: str | None, seller_mst: str | None,
                   accept: Callable[[bytes], bool], log, ask=None) -> Found | None:
        host = host_of(url)
        ctx = await self._ctx_for(browser, host)
        found: list[Found] = []
        tasks: set[asyncio.Task] = set()
        pages: list[Any] = []

        def offer(data: bytes | None, src: str, method: str = "") -> None:
            pdf = _pdf_in(data or b"")
            if pdf and not found and accept(pdf):
                found.append(Found(pdf, src))
                if method == "GET":
                    learn_pdf_url(host_of(src) or host, src, code, seller_mst)

        challenge: list[str] = []

        async def on_response(resp) -> None:
            try:
                ct = (resp.headers.get("content-type") or "").lower()
                u = resp.url.lower()
                if re.search(r"challenges\.cloudflare\.com/.*turnstile|recaptcha/api2/anchor|hcaptcha\.com/captcha", u):
                    challenge.append(u)  # trang có kiểm tra chống robot / captcha
                    return
                if resp.request.resource_type in ("image", "stylesheet", "font", "media"):
                    return
                if not ("pdf" in ct or "octet-stream" in ct or "zip" in ct or "json" in ct or "text/plain" in ct
                        or u.endswith(".pdf") or "pdf" in u or "download" in u):
                    return
                offer(await resp.body(), resp.url, resp.request.method)
            except Exception:  # noqa: BLE001
                pass

        async def on_download(dl) -> None:
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    dest = Path(tmp) / "f.bin"
                    await dl.save_as(dest)
                    offer(dest.read_bytes(), dl.url)
            except Exception:  # noqa: BLE001
                pass

        async def on_page(p) -> None:  # tab mới (PDF mở ở tab khác, blob:...)
            pages.append(p)
            try:
                p.on("response", lambda r: _spawn(on_response(r)))
                p.on("download", lambda d: _spawn(on_download(d)))
                await p.wait_for_load_state("domcontentloaded", timeout=15000)
                if p.url.startswith("blob:"):
                    b64 = await p.evaluate("""async (u) => { const b = new Uint8Array(await (await fetch(u)).arrayBuffer());
                        let s = ''; for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
                        return btoa(s); }""", p.url)
                    offer(base64.b64decode(b64), p.url)
                elif p.url.startswith("http") and not found:
                    r = await ctx.request.get(p.url)
                    offer(await r.body(), p.url, "GET")
            except Exception:  # noqa: BLE001
                pass

        def _spawn(c) -> None:
            t = asyncio.ensure_future(c)
            tasks.add(t)
            t.add_done_callback(tasks.discard)

        async def settle(ms: int = STEP_WAIT) -> None:
            for _ in range(ms // 250):
                if found:
                    return
                await asyncio.sleep(0.25)
            if tasks:
                await asyncio.wait(list(tasks), timeout=5)

        page = await ctx.new_page()
        pages.append(page)
        page.on("popup", lambda p: _spawn(on_page(p)))  # chỉ tab bật ra từ trang của hóa đơn này (phiên dùng chung)
        page.on("response", lambda r: _spawn(on_response(r)))
        page.on("download", lambda d: _spawn(on_download(d)))
        try:
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT)
            except Exception as e:  # noqa: BLE001 - link tải thẳng file thì goto báo "Download is starting"
                if "download" not in str(e).lower():
                    raise
            try:
                await page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:  # noqa: BLE001
                pass
            await settle(1500)
            if found:
                return found[0]

            solved = False
            # Trang chạy trên nền EasyInvoice (tên miền riêng): ô mã (#iFkey) ẩn tới khi chọn "tra cứu theo mã"
            if code and await page.evaluate("() => !!document.querySelector('#iFkey') && "
                                            "!!document.querySelector('img[src*=\"Captcha\"], img#captcha')"):
                if not ask:
                    raise NeedsCaptcha("trang EasyInvoice – captcha mỗi lần tra cứu")
                await page.evaluate("""(c) => { const r = document.querySelector('#fKeyCheck'); if (r) r.click();
                    const k = document.querySelector('#iFkey'); k.value = c;
                    k.dispatchEvent(new Event('input', {bubbles: true})); k.dispatchEvent(new Event('change', {bubbles: true})); }""",
                                    code)
                btn = await page.evaluate("() => { const b = document.querySelector('#btnSubmit'); "
                                          "if (b) b.setAttribute('data-aw-bt', 'sub'); return !!b; }")
                if not await self._solve_captcha(page, ask, {"i": "sub"} if btn else None, f"Mã tra cứu {code}"):
                    raise NeedsCaptcha("không lấy được ảnh captcha")
                solved = True
                try:
                    await page.wait_for_load_state("networkidle", timeout=15000)
                except Exception:  # noqa: BLE001
                    pass
                await settle()
                if found:
                    return found[0]
                code = None  # đã tra cứu xong -> khỏi điền lại mã
            info = await self._scan_all(page)
            # Có ô mã tra cứu đang trống -> điền mã + MST rồi bấm Tra cứu (có captcha + ask -> người dùng nhập)
            if code:
                filled = await self._fill_and_submit(page, info, code, seller_mst, host, ask)
                solved = filled == "captcha"
                if filled:
                    try:
                        await page.wait_for_load_state("networkidle", timeout=12000)
                    except Exception:  # noqa: BLE001
                        pass
                    await settle()
                    if found:
                        return found[0]
                    info = await self._scan_all(page)
            if self._captcha_needed(info) and not solved:
                if not ask:
                    raise NeedsCaptcha("captcha")
                # trang hiện captcha ngay (không có ô mã hoặc mã nằm sẵn trong link) -> người dùng nhập rồi gửi
                if not await self._solve_captcha(page, ask, None, "Nhập captcha của trang tra cứu"):
                    raise NeedsCaptcha("không lấy được ảnh captcha")
                try:
                    await page.wait_for_load_state("networkidle", timeout=12000)
                except Exception:  # noqa: BLE001
                    pass
                await settle()
                if found:
                    return found[0]
                info = await self._scan_all(page)

            # Khung iframe/embed đang hiển thị PDF
            for fr in info["frames"][:4]:
                if fr.startswith("http"):
                    try:
                        r = await ctx.request.get(fr)
                        offer(await r.body(), fr)
                    except Exception:  # noqa: BLE001
                        pass
                    if found:
                        return found[0]

            # Bấm các nút "Tải PDF / Tải hóa đơn / Xem hóa đơn..."
            clicked = 0
            for b in sorted(info["btns"], key=lambda b: -self._dl_score(b)):
                if self._dl_score(b) <= 0 or clicked >= 5:
                    break
                clicked += 1
                try:
                    await page.locator(f"[data-aw-bt='{b['i']}']").first.click(timeout=5000)
                except Exception:  # noqa: BLE001
                    continue
                await settle()
                if found:
                    return found[0]

            # Trang hiện hóa đơn dạng HTML -> in ra PDF
            pdf = await self._print_invoice_view(page)
            if pdf and accept(pdf):
                return Found(pdf, page.url, printed=True)
            if found:
                return found[0]
            if challenge:
                raise NeedsCaptcha("captcha / kiểm tra chống robot")
            return None
        finally:
            for t in list(tasks):
                t.cancel()
            for p in pages:  # đóng tab của hóa đơn này, GIỮ phiên của trang (cookie, captcha đã nhập) cho hóa đơn sau
                try:
                    await p.close()
                except Exception:  # noqa: BLE001
                    pass

    # ------------------------------------------------------------ trợ giúp
    async def _scan_all(self, page) -> dict[str, Any]:
        info = await page.evaluate(_SCAN_JS)
        info["frame_objs"] = []
        for fr in page.frames[1:]:  # ô tra cứu có thể nằm trong iframe
            try:
                sub = await fr.evaluate(_SCAN_JS)
            except Exception:  # noqa: BLE001
                continue
            if sub["inputs"]:
                info["frame_objs"].append((fr, sub))
            info["caps"] += sub["caps"]
            info["text"] += "\n" + sub["text"]
        return info

    @staticmethod
    def _captcha_needed(info: dict[str, Any]) -> bool:
        if info["caps"]:
            return True
        return any(CAPTCHA_HINT.search(i["text"]) for i in info["inputs"])

    @staticmethod
    def _dl_score(b: dict[str, Any]) -> int:
        t = (b.get("text") or "") + " " + (b.get("href") or "") + " " + (b.get("cls") or "")
        if not t.strip() or SKIP_TEXT.search(b.get("text") or ""):
            return 0
        s = 0
        if re.search(r"pdf", t, re.I):
            s += 5
        if DOWNLOAD_TEXT.search(t):
            s += 3
        if re.search(r"download|tai-?ve|taive", b.get("href") or "", re.I):
            s += 2
        return s

    async def _fill_and_submit(self, page, info: dict[str, Any], code: str, seller_mst: str | None,
                               host: str = "", ask=None) -> bool | str:
        """Điền mã tra cứu (+ MST), gặp captcha thì (nếu có ask) cho người dùng nhập, rồi bấm Tra cứu.
        Trả về False (không thấy ô mã), True (đã gửi), "captcha" (đã gửi kèm captcha người dùng nhập)."""
        hint = recipes().get(host, {}).get("fields") if host else None
        targets = [(page, info)] + info.get("frame_objs", [])
        for frame, inf in targets:
            if not inf["inputs"]:
                continue
            best, mst_in, btn = None, None, None
            if hint:  # cách điền đã học cho trang này (do AI chỉ ra lần trước)
                best = next((i for i in inf["inputs"] if hint.get("code") and hint["code"] in i["text"]), None)
                mst_in = next((i for i in inf["inputs"] if hint.get("mst") and hint["mst"] in i["text"]), None)
                btn = next((b for b in inf["btns"] if hint.get("btn") and hint["btn"] == (b.get("text") or "")[:60]), None)
            if best is None:
                cands = sorted(((_score_code_input(i["text"]), i) for i in inf["inputs"]), key=lambda x: -x[0])
                cands = [c for c in cands if c[0] > 0] or (
                    [(1, i) for i in inf["inputs"] if _score_code_input(i["text"]) >= 0][:1]
                    if len(inf["inputs"]) <= 3 else [])
                best = cands[0][1] if cands else None
            if best is None and frame is page:  # bố cục lạ -> nhờ AI chỉ ra ô mã / nút tra cứu (1 lần cho mỗi trang)
                pick = await asyncio.get_running_loop().run_in_executor(None, _ai_pick, inf["inputs"], inf["btns"])
                by_i = {i["i"]: i for i in inf["inputs"]}
                bt_i = {b["i"]: b for b in inf["btns"]}
                if pick and pick.get("code_input") in by_i:
                    best = by_i[pick["code_input"]]
                    mst_in = by_i.get(pick.get("mst_input"))
                    btn = bt_i.get(pick.get("submit_button"))
                    if host:
                        _host_rec(host)["fields"] = {"code": best["text"][:60],
                                                     "mst": mst_in["text"][:60] if mst_in else "",
                                                     "btn": (btn.get("text") or "")[:60] if btn else ""}
                        _save_recipes()
            if best is None:
                continue
            if best["value"].strip() != code.strip():
                loc = frame.locator(f"[data-aw-in='{best['i']}']").first
                await loc.fill("")
                await loc.type(code, delay=15)
            if seller_mst:
                for i in ([mst_in] if mst_in else inf["inputs"]):
                    if i is not best and (mst_in or any(w in i["text"] for w in MST_WORDS)) and not i["value"]:
                        try:
                            await frame.locator(f"[data-aw-in='{i['i']}']").first.fill(seller_mst)
                        except Exception:  # noqa: BLE001
                            pass
            result: bool | str = True
            if any(CAPTCHA_HINT.search(i["text"]) for i in inf["inputs"]) or inf["caps"]:
                if not ask:
                    raise NeedsCaptcha("captcha")
                btn = btn or self._submit_button(best, inf["btns"])
                if not await self._solve_captcha(frame, ask, btn, f"Mã tra cứu {code}"):
                    raise NeedsCaptcha("không lấy được ảnh captcha")
                return "captcha"
            # nhấn Enter trong ô, rồi bấm nút "Tra cứu" gần ô nhất (không bấm link menu sang trang khác)
            try:
                await frame.locator(f"[data-aw-in='{best['i']}']").first.press("Enter")
            except Exception:  # noqa: BLE001
                pass
            await asyncio.sleep(1.5)
            btn = btn or self._submit_button(best, inf["btns"])
            if btn:
                try:
                    await frame.locator(f"[data-aw-bt='{btn['i']}']").first.click(timeout=5000)
                except Exception:  # noqa: BLE001
                    pass
            return result
        return False

    async def _solve_captcha(self, frame, ask, btn: dict[str, Any] | None, reason: str) -> bool:
        """Chụp ảnh captcha -> người dùng nhập (ask chạy ngoài vòng lặp, không chặn các hóa đơn khác) -> điền, gửi."""
        st = await frame.evaluate(_CAPTCHA_JS)
        if not st.get("img") or not st.get("input"):
            return False
        png = await frame.locator("[data-aw-capimg='1']").first.screenshot(timeout=10000)
        img = "data:image/png;base64," + base64.b64encode(png).decode()
        answer = await asyncio.get_running_loop().run_in_executor(None, ask, img, reason)
        if not answer:
            return False
        cap = frame.locator("[data-aw-capin='1']").first
        await cap.fill("")
        await cap.type(answer.strip(), delay=20)
        # gửi đúng 1 lần (gửi 2 lần thì captcha đã dùng -> trang báo sai)
        if btn:
            try:
                await frame.locator(f"[data-aw-bt='{btn['i']}']").first.click(timeout=5000)
                return True
            except Exception:  # noqa: BLE001
                pass
        try:
            await cap.press("Enter")
        except Exception:  # noqa: BLE001
            pass
        return True

    @staticmethod
    def _submit_button(inp: dict[str, Any], btns: list[dict[str, Any]]) -> dict[str, Any] | None:
        def ok(b: dict[str, Any]) -> bool:
            href = (b.get("href") or "").strip()
            if b.get("tag") == "a" and href and not href.startswith(("#", "javascript")):
                return False  # link điều hướng (menu) – không phải nút gửi
            t = (b.get("text") or "").strip()
            return bool(SUBMIT_TEXT.search(t)) or (not t and re.search(r"search|submit|tracuu|btn", b.get("cls") or "", re.I))

        cands = []
        for b in btns:
            if not ok(b):
                continue
            d = ((b["x"] - inp["x"]) ** 2 + (b["y"] - inp["y"]) ** 2) ** 0.5
            if d < 450:
                cands.append((d - (200 if SUBMIT_TEXT.search(b.get("text") or "") else 0), b))
        return min(cands, key=lambda c: c[0])[1] if cands else None

    async def _print_invoice_view(self, page) -> bytes | None:
        """In trang đang hiện hóa đơn (chỉ khi trang có dấu hiệu là bản thể hiện hóa đơn)."""
        try:
            txt = (await page.evaluate("() => document.body ? document.body.innerText : ''")) or ""
        except Exception:  # noqa: BLE001
            return None
        low = txt.lower()
        if not (("hóa đơn" in low or "hoá đơn" in low or "invoice" in low)
                and re.search(r"t[ổo]ng\s*(c[ộo]ng\s*)?(ti[ềe]n\s*)?thanh\s*to[áa]n|total", low)):
            return None
        try:
            return await page.pdf(format="A4", print_background=True,
                                  margin={"top": "6mm", "bottom": "6mm", "left": "6mm", "right": "6mm"})
        except Exception:  # noqa: BLE001
            return None

    # ------------------------------------------------------------ HTML -> PDF (in tờ khai, bảng kê)
    def html_to_pdf(self, html: str, timeout: float = 120) -> bytes:
        async def job() -> bytes:
            browser = await self._get_browser()
            ctx = await browser.new_context()
            try:
                p = await ctx.new_page()
                await p.set_content(html, wait_until="load", timeout=60000)
                return await p.pdf(prefer_css_page_size=True, print_background=True)
            finally:
                await ctx.close()
        return self.run(job, timeout)


HEADLESS = HeadlessPdf()


def host_of(url: str) -> str:
    return (urlparse(url or "").hostname or "").lower()
