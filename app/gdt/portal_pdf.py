"""In "bản thể hiện" hóa đơn của trang hoadondientu.gdt.gov.vn ra PDF.

Cổng không có lệnh trả PDF: trang web tự vẽ hóa đơn (nút "Xem hóa đơn") rồi in bằng trình duyệt.
Ở đây gọi đúng chức năng "Xem hóa đơn" NGAY TRÊN TAB CỔNG THUẾ NGƯỜI DÙNG ĐÃ ĐĂNG NHẬP:
  - không mở tab mới, không tải lại trang (tải lại / mở trang lần 2 có thể làm mất phiên);
  - nếu tab đang ở trang khác thì chuyển sang Tra cứu bằng bộ định tuyến của chính trang (không reload);
  - in PDF bằng Chrome trên chính tab đó, chỉ lấy khung hóa đơn (ẩn phần còn lại khi in).
Mọi hàm chạy trên luồng trình duyệt (gọi qua BROWSER.call).
"""
from __future__ import annotations

import base64
from typing import Any

from .browser import BROWSER, PORTAL, BrowserError

LOOKUP_PATH = "/tra-cuu/tra-cuu-hoa-don"

_FIND = """() => { for (const el of document.querySelectorAll('*')) {
  const k = Object.keys(el).find(k => k.startsWith('__reactFiber')); if (!k) continue;
  for (let f = el[k], d = 0; f && d < 60; d++, f = f.return) { const s = f.stateNode;
    if (s && s.toggleViewInvoice && s.printInvoice) return s; } } return null; }"""

# Mở khung "Xem hóa đơn", đánh dấu khung đó và chèn CSS in chỉ hiện khung hóa đơn.
_OPEN_JS = """async ({row, tab}) => {
  const s = (""" + _FIND + """)();
  if (!s) return {err: 'nocomp'};
  window.__hddtOldTab = s.state.activeTab;
  await new Promise(r => s.setState({selectedRow: row, activeTab: tab}, r));
  s.toggleViewInvoice(true)();
  let m = null;
  for (let i = 0; i < 80; i++) {
    await new Promise(r => setTimeout(r, 200));
    const bodies = [...document.querySelectorAll('.ant-modal-body')];
    m = bodies[bodies.length - 1];
    if (m && /Tổng tiền thanh toán|Tổng cộng tiền thanh toán/.test(m.innerText)) break;
  }
  const ok = !!(m && /Tổng tiền thanh toán|Tổng cộng tiền thanh toán/.test(m.innerText));
  if (ok) {
    document.querySelectorAll('.__hddt_target').forEach(e => e.classList.remove('__hddt_target'));
    m.classList.add('__hddt_target');
    const root = m.closest('.ant-modal-root') || m.closest('.ant-modal-wrap') || m.parentElement;
    root.classList.add('__hddt_target_root');
    let st = document.getElementById('__hddt_print');
    if (!st) { st = document.createElement('style'); st.id = '__hddt_print'; document.head.appendChild(st); }
    st.textContent = `@media print {
      @page { size: A4; margin: 8mm; }
      html, body { overflow: visible !important; height: auto !important; background: #fff !important; }
      body > *:not(.__hddt_target_root):not(:has(.__hddt_target)) { display: none !important; }
      .ant-modal-mask, .ant-modal-header, .ant-modal-footer, .ant-modal-close, button { display: none !important; }
      .ant-modal-wrap { position: static !important; overflow: visible !important; }
      .ant-modal { position: static !important; top: 0 !important; margin: 0 !important; padding: 0 !important;
                   width: 100% !important; max-width: 100% !important; transform: none !important; }
      .ant-modal-content { box-shadow: none !important; }
      .__hddt_target { padding: 0 !important; overflow: visible !important; height: auto !important;
                       max-height: none !important; }
      .__hddt_target_root * { opacity: 1 !important; animation: none !important; transition: none !important; }
      /* CSS in của web cổng ẩn "body *" (chỉ hiện .printSection) -> cho khung hóa đơn hiện lại */
      .__hddt_target_root, .__hddt_target_root .ant-modal-wrap, .__hddt_target_root .ant-modal,
      .__hddt_target_root .ant-modal-content, .__hddt_target, .__hddt_target * {
        visibility: visible !important; print-color-adjust: exact !important; }
      .__hddt_target .printSection { position: static !important; }
    }`;
  }
  return {ok};
}"""

_CLOSE_JS = """async () => {
  const st = document.getElementById('__hddt_print'); if (st) st.remove();
  document.querySelectorAll('.__hddt_target').forEach(e => e.classList.remove('__hddt_target'));
  document.querySelectorAll('.__hddt_target_root').forEach(e => e.classList.remove('__hddt_target_root'));
  const s = (""" + _FIND + """)();
  if (!s) return;
  try { s.toggleViewInvoice(false)(); } catch (e) {}
  if (window.__hddtOldTab !== undefined) await new Promise(r => s.setState({activeTab: window.__hddtOldTab}, r));
}"""

# Chỉ lấy HTML hóa đơn đã vẽ (kèm CSS của trang) để in bằng Chrome chạy ẩn – không in trên tab cổng.
_RENDER_JS = """async ({row, tab, needCss}) => {
  const s = (""" + _FIND + """)();
  if (!s) return {err: 'nocomp'};
  const oldTab = s.state.activeTab;
  await new Promise(r => s.setState({selectedRow: row, activeTab: tab}, r));
  s.toggleViewInvoice(true)();
  let m = null;
  for (let i = 0; i < 80; i++) {
    await new Promise(r => setTimeout(r, 200));
    const bodies = [...document.querySelectorAll('.ant-modal-body')];
    m = bodies[bodies.length - 1];
    if (m && /Tổng tiền thanh toán|Tổng cộng tiền thanh toán/.test(m.innerText)) break;
  }
  const html = m ? m.innerHTML : '';
  let css = '';
  if (needCss) css = [...document.styleSheets].map(ss => { try { return [...ss.cssRules].map(r => r.cssText).join('\\n'); }
                       catch (e) { return ss.href ? '@import url(' + ss.href + ');' : ''; } }).join('\\n');
  s.toggleViewInvoice(false)();
  await new Promise(r => s.setState({activeTab: oldTab}, r));
  return {html, css, ok: !!(m && /Tổng tiền thanh toán|Tổng cộng tiền thanh toán/.test(m.innerText))};
}"""

_state: dict[str, Any] = {"css": None}
_unavailable: dict[int, str] = {}      # theo trang HĐĐT (tải nhiều công ty: mỗi công ty 1 cửa sổ)
_cdp: dict[int, Any] = {}              # id(tab) -> phiên CDP để in


def reset(slot: int = 0) -> None:
    """Gọi đầu mỗi lần tải: thử lại chức năng in từ cổng."""
    _unavailable.pop(slot, None)


def _lookup_page(slot: int = 0):
    """Tab cổng đã đăng nhập, đang có chức năng Xem hóa đơn. Không mở tab mới, không reload."""
    p = BROWSER._slot_page(slot)
    if p.evaluate("(" + _FIND + ")() === null"):
        # Chuyển trang bằng bộ định tuyến của web cổng (Next.js) -> giữ nguyên phiên đăng nhập.
        moved = p.evaluate("""(path) => { const r = window.next && window.next.router;
            if (!r || !r.push) return false; r.push(path); return true; }""", LOOKUP_PATH)
        if not moved:
            raise BrowserError("Tab cổng thuế không ở trang Tra cứu hóa đơn")
        try:
            p.wait_for_function("(" + _FIND + ")() !== null", timeout=20000)
        except Exception as e:  # noqa: BLE001
            raise BrowserError("Không mở được trang Tra cứu hóa đơn trên tab cổng thuế") from e
    return p


def render(row: dict[str, Any], sco: bool) -> bytes:
    """row: dữ liệu hóa đơn (nên là JSON chi tiết từ /invoices/detail). Trả về PDF."""
    slot = BROWSER.thread_slot()
    if _unavailable.get(slot):
        raise BrowserError(_unavailable[slot])

    def run() -> bytes:
        p = _lookup_page(slot)
        try:
            res = p.evaluate(_OPEN_JS, {"row": row, "tab": 2 if sco else 1})
            if res.get("err"):
                raise BrowserError("Không tìm thấy chức năng xem hóa đơn trên trang cổng")
            if not res.get("ok"):
                raise BrowserError("Cổng không hiển thị được hóa đơn")
            cdp = _cdp.get(id(p))
            if not cdp:
                cdp = p.context.new_cdp_session(p)
                _cdp[id(p)] = cdp
            out = cdp.send("Page.printToPDF", {"printBackground": True, "preferCSSPageSize": True})
            return base64.b64decode(out["data"])
        finally:
            try:
                p.evaluate(_CLOSE_JS)
            except Exception:  # noqa: BLE001
                pass

    try:
        for _attempt in range(2):  # PDF in ra phải có chữ (tránh file trắng khi trang chưa vẽ xong)
            pdf = BROWSER.call(run, 120)
            if _has_text(pdf):
                return pdf
        raise BrowserError("Bản in từ cổng bị trắng (không có nội dung)")
    except BrowserError as e:
        from .browser import BrowserAuthError
        if isinstance(e, BrowserAuthError):
            raise
        if "trang" in str(e).lower() and "Tra cứu" in str(e):
            _unavailable[slot] = str(e)  # không mở được trang -> các hóa đơn sau khỏi thử lại
        raise


def capture(row: dict[str, Any], sco: bool) -> str:
    """Chỉ lấy HTML hóa đơn đã vẽ trên trang cổng (nhanh); việc in ra PDF làm song song ở Chrome chạy ẩn."""

    slot = BROWSER.thread_slot()

    def run() -> str:
        p = _lookup_page(slot)
        res = p.evaluate(_RENDER_JS, {"row": row, "tab": 2 if sco else 1, "needCss": _state["css"] is None})
        if res.get("err"):
            raise BrowserError("Không tìm thấy chức năng xem hóa đơn trên trang cổng")
        if res.get("css"):
            _state["css"] = res["css"]
        if not res.get("ok"):
            raise BrowserError("Cổng không hiển thị được hóa đơn")
        return ("<html><head><meta charset='utf-8'><base href='" + PORTAL + "/'><style>" + (_state["css"] or "")
                + "</style><style>body{background:#fff;margin:0} button{display:none!important}"
                  " @page{size:A4;margin:8mm}</style></head><body>" + res["html"] + "</body></html>")
    return BROWSER.call(run, 120)


def print_html(html: str) -> bytes:
    """In HTML đã lấy từ cổng ra PDF bằng Chrome chạy ẩn (gọi được từ nhiều luồng cùng lúc)."""
    from ..providers.headless import HEADLESS
    pdf = HEADLESS.html_to_pdf(html, 180)
    if not _has_text(pdf):
        raise BrowserError("Bản in từ cổng bị trắng (không có nội dung)")
    return pdf


def _has_text(pdf: bytes) -> bool:
    import io

    from pypdf import PdfReader
    try:
        return len((PdfReader(io.BytesIO(pdf)).pages[0].extract_text() or "").strip()) > 50
    except Exception:  # noqa: BLE001
        return False
