"""In "bản thể hiện" hóa đơn của trang hoadondientu.gdt.gov.vn ra PDF.

Cổng không có lệnh trả PDF: trang web tự vẽ hóa đơn (nút "Xem hóa đơn") rồi in bằng trình duyệt.
Ở đây gọi đúng chức năng "Xem hóa đơn" của trang Tra cứu hóa đơn trên Chrome đã đăng nhập, lấy phần
hóa đơn đã vẽ cùng CSS của trang, rồi để Chrome in ra PDF (khổ A4).
Mọi hàm chạy trên luồng trình duyệt (gọi qua BROWSER.call).
"""
from __future__ import annotations

import base64
from typing import Any

from .browser import BROWSER, PORTAL, BrowserError

LOOKUP_PAGE = PORTAL + "/tra-cuu/tra-cuu-hoa-don"

_FIND = """() => { for (const el of document.querySelectorAll('*')) {
  const k = Object.keys(el).find(k => k.startsWith('__reactFiber')); if (!k) continue;
  for (let f = el[k], d = 0; f && d < 60; d++, f = f.return) { const s = f.stateNode;
    if (s && s.toggleViewInvoice && s.printInvoice) return s; } } return null; }"""

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

_state: dict[str, Any] = {"css": None, "cdp": None}


def _lookup_page():
    p = BROWSER._portal_page()
    if p.evaluate("(" + _FIND + ")() === null"):
        p.goto(LOOKUP_PAGE, wait_until="domcontentloaded", timeout=60000)
        p.wait_for_function("(" + _FIND + ")() !== null", timeout=30000)
    return p


def render(row: dict[str, Any], sco: bool) -> bytes:
    """row: dữ liệu hóa đơn (nên là JSON chi tiết từ /invoices/detail). Trả về PDF."""
    def run() -> bytes:
        p = _lookup_page()
        res = p.evaluate(_RENDER_JS, {"row": row, "tab": 2 if sco else 1, "needCss": _state["css"] is None})
        if res.get("err"):
            raise BrowserError("Không tìm thấy chức năng xem hóa đơn trên trang cổng")
        if res.get("css"):
            _state["css"] = res["css"]
        if not res.get("ok"):
            raise BrowserError("Cổng không hiển thị được hóa đơn")
        q = BROWSER.aux_page("hddt-print")
        q.set_content("<html><head><meta charset='utf-8'><base href='" + PORTAL + "/'><style>" + (_state["css"] or "")
                      + "</style><style>body{background:#fff;margin:0} button{display:none!important}"
                        " @page{size:A4;margin:8mm}</style></head><body>" + res["html"] + "</body></html>",
                      wait_until="load", timeout=60000)
        cdp = _state.get("cdp")
        if not cdp or _state.get("cdp_page") is not q:
            cdp = q.context.new_cdp_session(q)
            _state["cdp"], _state["cdp_page"] = cdp, q
        out = cdp.send("Page.printToPDF", {"printBackground": True, "preferCSSPageSize": True})
        return base64.b64decode(out["data"])

    for _attempt in range(2):  # PDF in ra phải có chữ (tránh file trắng khi trang chưa vẽ xong)
        pdf = BROWSER.call(run, 120)
        if _has_text(pdf):
            return pdf
        _state["css"] = None  # nạp lại CSS rồi in lại
    raise BrowserError("Bản in từ cổng bị trắng (không có nội dung)")


def _has_text(pdf: bytes) -> bool:
    import io

    from pypdf import PdfReader
    try:
        return len((PdfReader(io.BytesIO(pdf)).pages[0].extract_text() or "").strip()) > 50
    except Exception:  # noqa: BLE001
        return False
