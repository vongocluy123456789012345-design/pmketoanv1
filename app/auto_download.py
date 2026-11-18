"""Tự động tải hóa đơn từ cổng thuế qua Chrome mà người dùng đã đăng nhập.

Với mỗi tháng và từng loại hóa đơn:
1. Tra cứu danh sách (tự lật trang).
2. Tải bảng kê Excel của cổng ("Xuất hóa đơn").
3. Tải XML từng hóa đơn. Không có XML thì lưu dữ liệu chi tiết JSON.
4. Đọc mã/link tra cứu trong XML rồi tìm và tải PDF gốc từ trang của nhà cung cấp.
Cuối cùng lập file Excel tổng hợp có link tới từng file XML/PDF.

Thư mục lưu:  OUTPUT_DIR/<MST>/<YYYY-MM>/<loại>/{BangKe_*.xlsx, XML/, PDF/}
Chạy lại cùng khoảng thời gian thì file đã có sẽ được bỏ qua.
"""
from __future__ import annotations

import base64
import io
import json
import re
import threading
import time
import traceback
import uuid
import zipfile
from calendar import monthrange
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlencode

from .config import settings
from .gdt.browser import BROWSER, BrowserAuthError, FetchResult
from .dates import vn_date
from .gdt.xml_parser import parse_invoice_xml

# Khởi tạo engine OCR tự động giải captcha
try:
    import ddddocr
    _OCR_ENGINE = ddddocr.DdddOcr(show_ad=False)
except Exception:  # noqa: BLE001
    _OCR_ENGINE = None


def _solve_captcha_ocr(image_data_uri: str) -> str | None:
    """Giải mã captcha từ chuỗi data-uri (SVG hoặc PNG/JPG), trả về chữ thường + số."""
    if _OCR_ENGINE is None or not image_data_uri:
        return None
    try:
        header, b64_data = image_data_uri.split(",", 1) if "," in image_data_uri else ("", image_data_uri)
        raw_bytes = base64.b64decode(b64_data)
        if "svg" in header.lower() or raw_bytes.lstrip()[:4] in (b"<svg", b"<?xm"):
            import cairosvg
            raw_bytes = cairosvg.svg2png(bytestring=raw_bytes, background_color="white")
        res = _OCR_ENGINE.classification(raw_bytes)
        clean = "".join(c for c in str(res) if c.isalnum()).lower()
        return clean if len(clean) >= 3 else None
    except Exception:  # noqa: BLE001
        return None


# Mỗi loại = 1 tab / 1 điều kiện lọc trên trang "Tra cứu hóa đơn" của cổng.
#   ep: query (HĐ điện tử) | sco-query (HĐ có mã từ máy tính tiền – cổng bắt nhập captcha)
#   tab: phần trong ngoặc của header Action mà web cổng gửi
KINDS: dict[str, dict[str, Any]] = {
    "mua-vao": {"label": "Mua vào – đã cấp mã", "folder": "Mua vao - Co ma", "ep": "query",
                "direction": "purchase", "filter": "ttxly==5", "tab": "hóa đơn mua vào"},
    "mua-vao-khong-ma": {"label": "Mua vào – CQT nhận không mã", "folder": "Mua vao - Khong ma",
                         "ep": "query", "direction": "purchase", "filter": "ttxly==6",
                         "tab": "hóa đơn mua vào"},
    "mua-vao-mtt": {"label": "Mua vào – máy tính tiền", "folder": "Mua vao - May tinh tien",
                    "ep": "sco-query", "direction": "purchase", "filter": "ttxly==8",
                    "tab": "hóa đơn máy tính tiền mua vào", "captcha": True},
    "ban-ra": {"label": "Bán ra", "folder": "Ban ra", "ep": "query", "direction": "sold",
               "filter": None, "tab": "hóa đơn bán ra"},
    "ban-ra-mtt": {"label": "Bán ra – máy tính tiền", "folder": "Ban ra - May tinh tien",
                   "ep": "sco-query", "direction": "sold", "filter": None,
                   "tab": "hóa đơn máy tính tiền bán ra", "captcha": True},
}

TTHAI = {1: "Hóa đơn mới", 2: "Thay thế", 3: "Điều chỉnh", 4: "Đã bị thay thế",
         5: "Đã bị điều chỉnh", 6: "Đã bị hủy"}
PAGE_SIZE = 50
CAPTCHA_WAIT = 300  # giây chờ người dùng nhập captcha


class Skip(Exception):
    """Bỏ qua loại hóa đơn hiện tại (vd người dùng không nhập captcha)."""


class Cancelled(Exception):
    pass


def month_list(m_from: str, m_to: str) -> list[tuple[date, date]]:
    """'2025-01', '2025-03' -> [(1/1, 31/1), (1/2, 28/2), (1/3, 31/3)]."""
    y, m = map(int, m_from.split("-"))
    y2, m2 = map(int, m_to.split("-"))
    out = []
    while (y, m) <= (y2, m2):
        out.append((date(y, m, 1), date(y, m, monthrange(y, m)[1])))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def date_chunks(d_from: date, d_to: date) -> list[tuple[date, date]]:
    """Chia khoảng ngày thành từng tháng (cổng chỉ cho tra tối đa 1 tháng/lần).
    01/10–15/12 -> [(01/10, 31/10), (01/11, 30/11), (01/12, 15/12)]."""
    out, cur = [], d_from
    while cur <= d_to:
        end = min(date(cur.year, cur.month, monthrange(cur.year, cur.month)[1]), d_to)
        out.append((cur, end))
        cur = date(end.year + (end.month == 12), end.month % 12 + 1, 1)
    return out


def full_month(d1: date, d2: date) -> bool:
    return d1.day == 1 and d2.day == monthrange(d2.year, d2.month)[1]


def period(d1: date, d2: date) -> str:
    """Nhãn kỳ: '10/2026' nếu trọn tháng, ngược lại '05/10–20/10/2026'."""
    return f"{d1:%m/%Y}" if full_month(d1, d2) else f"{d1:%d/%m}–{d2:%d/%m/%Y}"


def _safe(s: Any) -> str:
    return re.sub(r"[^0-9A-Za-z]+", "-", str(s or "")).strip("-")


def _s(v: Any) -> str:
    return "" if v is None else str(v)


@dataclass
class AutoJob:
    id: str
    months: list[tuple[date, date]]
    kinds: list[str]
    get_excel: bool
    get_xml: bool
    get_pdf: bool
    easy_captcha: bool = False  # có hỏi captcha để lấy PDF gốc EasyInvoice của người bán khác không
    mst: str = ""
    out_dir: str = ""
    status: str = "running"  # running | done | error | auth_error | cancelled
    total: int = 0
    done: int = 0
    xml_ok: int = 0
    pdf_ok: int = 0
    pdf_missing: int = 0
    pdf_hddt: int = 0
    excel_ok: int = 0
    errors: list[str] = field(default_factory=list)
    log: list[str] = field(default_factory=list)
    captcha: dict[str, Any] | None = None  # {"image": data-uri, "reason": str} khi đang chờ nhập
    summary_file: str | None = None
    cancel: bool = False
    _captcha_ev: threading.Event = field(default_factory=threading.Event)
    _captcha_value: str | None = None
    _rows: list[dict[str, Any]] = field(default_factory=list)

    def say(self, msg: str) -> None:
        self.log.append(f"{datetime.now():%H:%M:%S} {msg}")
        self.log = self.log[-400:]

    def err(self, msg: str) -> None:
        self.errors.append(msg)
        self.say("LỖI " + msg)

    def public(self) -> dict[str, Any]:
        return {
            "id": self.id, "status": self.status, "mst": self.mst, "out_dir": self.out_dir,
            "total": self.total, "done": self.done, "xml_ok": self.xml_ok, "pdf_ok": self.pdf_ok,
            "pdf_missing": self.pdf_missing, "pdf_hddt": self.pdf_hddt, "excel_ok": self.excel_ok,
            "errors": self.errors[-30:], "log": self.log[-80:], "captcha": self.captcha,
            "summary_file": self.summary_file,
            "months": [period(a, b) for a, b in self.months], "kinds": self.kinds,
        }

    def prompt_captcha(self, image: str, reason: str, allow_ocr: bool = True) -> str:
        """Tự động đọc captcha bằng OCR trước; nếu không đọc được thì hiện ảnh chờ người dùng nhập."""
        if allow_ocr:
            auto_val = _solve_captcha_ocr(image)
            if auto_val:
                self.say(f"OCR tự giải captcha ({reason}): {auto_val}")
                return auto_val

        self._captcha_ev.clear()
        self._captcha_value = None
        self.captcha = {"image": image, "reason": reason}
        self.say(f"Chờ bạn nhập captcha: {reason}")
        ok = self._captcha_ev.wait(CAPTCHA_WAIT)
        self.captcha = None
        if self.cancel:
            raise Cancelled()
        if not ok or not self._captcha_value:
            raise Skip("không nhập captcha")
        return self._captcha_value

    def submit_captcha(self, value: str | None) -> None:
        self._captcha_value = (value or "").strip() or None
        self._captcha_ev.set()


JOBS: dict[str, AutoJob] = {}


# ---------------------------------------------------------------- cổng thuế
def _check(r: FetchResult, what: str) -> FetchResult:
    if not r.ok:
        raise RuntimeError(f"{what}: cổng trả lỗi {r.status} – {r.message()}")
    return r


class Runner:
    def __init__(self, job: AutoJob):
        self.job = job
        self.captcha: dict[str, str] | None = None  # captcha đang dùng cho sco-query
        self.excel_done = False
        self._page_state: tuple[str | None, int, int | None] = (None, 1, None)

    # ----- captcha (tự động OCR, dự phòng nhập tay) -----
    def ask_captcha(self, reason: str, allow_ocr: bool = True) -> dict[str, str]:
        job = self.job
        r = _check(BROWSER.fetch("/captcha", auth=False), "Lấy captcha")
        data = r.json()
        img = "data:image/svg+xml;base64," + base64.b64encode(data["content"].encode()).decode()
        self.captcha = {"ckey": data["key"], "cvalue": job.prompt_captcha(img, reason, allow_ocr=allow_ocr)}
        return self.captcha

    def call(self, kind: dict[str, Any], path: str, query: str | dict | None, action: str,
             what: str) -> FetchResult:
        """Gọi cổng; với loại cần captcha thì kèm ckey/cvalue (ở cả tham số và header như web cổng)."""
        if not kind.get("captcha") or "export-xml" in path:
            r = BROWSER.fetch(path, query, action, timeout=getattr(self, "timeout", 120),
                              retries=getattr(self, "retries", 3))
            time.sleep(settings.gdt_delay)
            return _check(r, what)
        last = None
        for attempt in range(4):
            # 2 lần đầu dùng OCR tự giải; từ lần thứ 3 nếu vẫn sai thì mở ô cho người dùng nhập tay
            use_ocr = attempt < 2
            cap = self.captcha or self.ask_captcha(f"{what} ({kind['label']})", allow_ocr=use_ocr)
            q = query if isinstance(query, str) else urlencode(query or {})
            q = (q + "&" if q else "") + urlencode(cap)
            r = BROWSER.fetch(path, q, action, extra_headers=cap, raise_401=False)
            time.sleep(settings.gdt_delay)
            self.captcha = None  # captcha của cổng chỉ dùng được 1 lần (mỗi trang / mỗi lần xuất cần mã mới)
            if r.ok:
                return r
            if r.status == 401:
                # Cổng trả 401 cả khi captcha sai -> kiểm tra phiên bằng 1 lệnh không cần captcha.
                BROWSER.fetch("/security-taxpayer/profile")  # phiên hết hạn thật -> BrowserAuthError
            last = r
            self.job.say(f"{what}: cổng từ chối captcha (HTTP {r.status} {r.message()}), thử lại mã mới")
        return _check(last, what)

    # ----- các bước -----
    def search_string(self, kind: dict[str, Any], d1: date, d2: date) -> str:
        s = f"tdlap=ge={d1:%d/%m/%Y}T00:00:00;tdlap=le={d2:%d/%m/%Y}T23:59:59"
        return s + (";" + kind["filter"] if kind["filter"] else "")

    def list_invoices(self, kind: dict[str, Any], d1: date, d2: date, rows: list | None = None,
                      state: str | None = None, page_no: int = 1, total: int | None = None,
                      one_page: bool = False) -> list[dict[str, Any]]:
        path = f"/{kind['ep']}/invoices/{kind['direction']}"
        search = self.search_string(kind, d1, d2)
        rows = list(rows or [])
        while True:
            q = f"sort=tdlap:desc&size={PAGE_SIZE}" + (f"&state={state}" if state else "") + f"&search={search}"
            what = f"Tra cứu danh sách {period(d1, d2)} – trang {page_no}"
            if total:
                what += f"/{-(-total // PAGE_SIZE)} (tổng {total} hóa đơn)"
            data = self.call(kind, path, q, f"Tìm kiếm ({kind['tab']})", what).json()
            page = data.get("datas") or []
            rows += page
            state = data.get("state")
            total = data.get("total") or total
            page_no += 1
            if self.job.cancel:
                raise Cancelled()
            self.last_total = total  # tổng số hóa đơn cổng báo – để đối chiếu không sót
            if one_page:
                self._page_state = (state if page else None, page_no, total)
                return rows
            if not state or not page:
                return rows

    def list_via_excel(self, kind: dict[str, Any], d1: date, d2: date, dest: Path) -> list[dict[str, Any]]:
        """Loại cần captcha: lấy danh sách từ bảng kê Excel (1 captcha cho cả tháng) thay vì
        lật từng trang (mỗi trang 1 captcha). Trả về danh sách theo đúng tên trường API."""
        from .sources.file_import import read_excel, row_from_excel

        first = self.list_invoices(kind, d1, d2, one_page=True)
        state, page_no, total = self._page_state
        if not state:  # chỉ có 1 trang -> đã đủ
            return first
        def parse() -> list[dict[str, Any]]:
            out = []
            for r in read_excel(dest.read_bytes(), dest.name):
                row = row_from_excel(self.job.mst, r, kind["direction"])
                row["tthai"] = r.get("tthai_text")
                out.append(row)
            return out

        if dest.exists():  # bảng kê đã tải lần trước và còn đủ -> khỏi hỏi thêm captcha
            try:
                rows = parse()
                if total and len(rows) >= total:
                    self.excel_done = True
                    self.job.say(f"{period(d1, d2)} {kind['label']}: dùng bảng kê Excel đã có ({len(rows)} hóa đơn)")
                    return rows
            except Exception:  # noqa: BLE001
                pass
        self.job.say(f"{period(d1, d2)} {kind['label']}: tổng {total} hóa đơn – lấy danh sách từ bảng kê Excel "
                     "(chỉ cần 1 captcha thay vì lật từng trang)")
        try:
            self.export_excel(kind, d1, d2, dest)
            self.excel_done = True
            rows = parse()
        except (Skip, Cancelled, BrowserAuthError):
            raise
        except Exception as e:  # noqa: BLE001
            self.job.err(f"{period(d1, d2)} {kind['label']}: không đọc được bảng kê Excel ({e}), chuyển sang lật trang")
            rows = []
        if total and len(rows) >= total:
            return rows
        if rows:
            self.job.say(f"Bảng kê Excel chỉ có {len(rows)}/{total} hóa đơn, lật trang lấy tiếp")
        return self.list_invoices(kind, d1, d2, rows=first, state=state, page_no=page_no, total=total)

    def export_excel(self, kind: dict[str, Any], d1: date, d2: date, dest: Path) -> None:
        search = self.search_string(kind, d1, d2)
        if kind["direction"] == "purchase":
            path, q = f"/{kind['ep']}/invoices/export-excel", f"sort=tdlap:desc&search={search}&type=purchase"
        else:
            path, q = f"/{kind['ep']}/invoices/export-excel-sold", f"sort=tdlap:desc&search={search}"
        r = self.call(kind, path, q, f"Xuất hóa đơn ({kind['tab']})", "Xuất bảng kê Excel")
        if r.body[:2] != b"PK":
            raise RuntimeError(f"Bảng kê Excel: cổng không trả file Excel ({r.message()})")
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(r.body)

    def export_xml(self, kind: dict[str, Any], inv: dict[str, Any]) -> bytes:
        params = {k: _s(inv.get(k)) for k in ("nbmst", "khhdon", "shdon", "khmshdon")}
        r = self.call(kind, f"/{kind['ep']}/invoices/export-xml", params,
                      f"Xuất xml ({kind['tab']})", "Tải XML")
        content = r.body
        if content[:2] == b"PK":
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                names = [n for n in zf.namelist() if n.lower().endswith(".xml")]
                if not names:
                    raise RuntimeError("file zip của cổng không có XML")
                names.sort(key=lambda n: (n.lower() != "invoice.xml", n))
                return zf.read(names[0])
        if content.lstrip()[:1] == b"<":
            return content
        raise RuntimeError(f"cổng không trả XML ({r.message()})")

    def detail(self, kind: dict[str, Any], inv: dict[str, Any]) -> dict[str, Any]:
        params = {k: _s(inv.get(k)) for k in ("nbmst", "khhdon", "shdon", "khmshdon")}
        r = BROWSER.fetch(f"/{kind['ep']}/invoices/detail", params, "")
        time.sleep(settings.gdt_delay)
        return _check(r, "Lấy chi tiết").json()


# ---------------------------------------------------------------- PDF gốc
def make_ref(inv: dict[str, Any], parsed: dict[str, Any]):
    from .providers import InvoiceRef
    return InvoiceRef(
        seller_mst=_s(inv.get("nbmst")), buyer_mst=_s(inv.get("nmmst")) or None,
        khmshdon=_s(inv.get("khmshdon")), khhdon=_s(inv.get("khhdon")), shdon=_s(inv.get("shdon")),
        msttcgp=_s(inv.get("msttcgp")) or parsed.get("msttcgp"),
        mccqt=parsed.get("mccqt") or _s(inv.get("mhdon")) or None,
        lookup_code=parsed.get("lookup_code"), lookup_urls=parsed.get("lookup_urls") or [],
    )


def fetch_pdf(inv: dict[str, Any], parsed: dict[str, Any]) -> tuple[bytes | None, dict[str, Any]]:
    from .providers import fetch_original_pdf

    ref = make_ref(inv, parsed)
    out = fetch_original_pdf(ref)
    info = {"provider": out.provider.name if out.provider else "",
            "lookup_code": ref.lookup_code or "", "lookup_url": (ref.lookup_urls or [""])[0]}
    if out.result:
        info["pdf_source"] = out.result.source_url
        if out.result.printed:
            info["pdf_note"] = "PDF in từ trang tra cứu của NCC (NCC không cho tải file)"
        return out.result.content, info
    reason = "; ".join(a.get("error") or a.get("result") or "" for a in out.attempts)[:500]
    if out.needs_captcha:
        reason = "Trang tra cứu cần captcha. " + reason
        info["needs_captcha"] = True
    info["pdf_error"] = reason or "Không tìm thấy PDF gốc"
    return None, info


# ---------------------------------------------------------------- điều phối
def run(job: AutoJob) -> None:
    rn = Runner(job)
    try:
        st = BROWSER.status()
        if not st.get("logged_in"):
            raise BrowserAuthError("Chưa đăng nhập cổng thuế trên trình duyệt.")
        job.mst = _s(st.get("mst")) or "chua-ro-MST"
        from .gdt import portal_pdf
        portal_pdf.reset()
        root = Path(settings.output_dir) / _safe(job.mst)
        job.out_dir = str(root)
        job.say(f"MST {job.mst} – lưu vào {root}")

        # Pha 1: giải/nhập captcha ngay từ đầu (lấy danh sách các loại cần captcha cho mọi tháng).
        listed: dict[tuple[date, str], Any] = {}
        cap_keys = [k for k in job.kinds if KINDS[k].get("captcha")]
        if cap_keys:
            job.say(f"Bước 1: lấy danh sách hóa đơn máy tính tiền cho {len(job.months)} kỳ – "
                    "tự động giải captcha, sau đó phần mềm tự tải hết")
            for d1, d2 in job.months:
                for key in cap_keys:
                    kind = KINDS[key]
                    if job.cancel:
                        raise Cancelled()
                    try:
                        listed[(d1, key)] = list_kind(rn, job, kind, d1, d2, root / f"{d1:%Y-%m}" / kind["folder"])
                    except Skip as e:
                        listed[(d1, key)] = "skip"
                        job.err(f"{period(d1, d2)} {kind['label']}: bỏ qua – {e}")
                    except (BrowserAuthError, Cancelled):
                        raise
                    except Exception as e:  # noqa: BLE001
                        listed[(d1, key)] = "skip"
                        job.err(f"{period(d1, d2)} {kind['label']}: {e}")
            job.say("ĐÃ XONG PHẦN XỬ LÝ CAPTCHA – phần còn lại chạy tự động")

        # Pha 2: tải bảng kê, XML, PDF cho từng tháng.
        for d1, d2 in job.months:
            for key in job.kinds:
                kind = KINDS[key]
                if job.cancel:
                    raise Cancelled()
                if listed.get((d1, key)) == "skip":
                    continue
                try:
                    run_kind(rn, job, kind, d1, d2, root / f"{d1:%Y-%m}" / kind["folder"], listed.get((d1, key)))
                except Skip as e:
                    job.err(f"{period(d1, d2)} {kind['label']}: bỏ qua – {e}")
                except (BrowserAuthError, Cancelled):
                    raise
                except Exception as e:  # noqa: BLE001 - lỗi 1 loại không dừng cả job
                    job.err(f"{period(d1, d2)} {kind['label']}: {e}")
            write_month_summary(job, root, d1, d2)   # file tổng hợp riêng của tháng
            write_summary(job, root)                 # file tổng hợp cả kỳ (cập nhật sau mỗi tháng)
            job.say(f"Xong {period(d1, d2)} – đã lưu file tổng hợp tháng")
        job.status = "done"
        job.say("Hoàn tất")
    except Cancelled:
        job.status = "cancelled"
        job.say("Đã dừng theo yêu cầu")
    except BrowserAuthError as e:
        job.status = "auth_error"
        job.err(str(e))
    except Exception as e:  # noqa: BLE001
        job.status = "error"
        job.err(f"{e}\n{traceback.format_exc()[-1200:]}")
    finally:
        job.captcha = None
        if job._rows and job.out_dir:
            try:
                write_summary(job, Path(job.out_dir))
            except Exception as e:  # noqa: BLE001
                job.err(f"Không ghi được file tổng hợp: {e}")


def list_kind(rn: Runner, job: AutoJob, kind: dict[str, Any], d1: date, d2: date,
              folder: Path) -> tuple[list[dict[str, Any]], Path, bool]:
    """Lấy danh sách hóa đơn 1 loại / 1 tháng. Trả về (rows, đường dẫn bảng kê, đã tải bảng kê chưa)."""
    suffix = f"{d1:%Y-%m}" if full_month(d1, d2) else f"{d1:%Y-%m-%d}_den_{d2:%Y-%m-%d}"
    dest = folder / f"BangKe_{_safe(kind['folder'])}_{suffix}.xlsx"
    rn.excel_done = False
    cache = folder / f".danhsach_{_safe(kind['folder'])}_{suffix}.json"
    if kind.get("captcha") and cache.exists():
        try:
            c = json.loads(cache.read_text(encoding="utf-8"))
            if c.get("saved", "") > d2.isoformat():
                rows = c["rows"]
                rn.excel_done = dest.exists()
                if rn.excel_done:
                    job.excel_ok += 1
                job.say(f"{period(d1, d2)} {kind['label']}: dùng danh sách đã lưu lần trước "
                        f"({len(rows)} hóa đơn) – không cần captcha")
                return rows, dest, rn.excel_done
        except Exception:  # noqa: BLE001 - file hỏng -> lấy lại
            pass
    if kind.get("captcha"):
        rows = rn.list_via_excel(kind, d1, d2, dest)
        if rows and job.get_excel and not rn.excel_done:
            try:
                rn.export_excel(kind, d1, d2, dest)
                rn.excel_done = True
            except Skip:
                raise
            except Exception as e:  # noqa: BLE001
                job.err(f"{period(d1, d2)} {kind['label']}: {e}")
                rn.excel_done = True
        if rn.excel_done and dest.exists():
            job.excel_ok += 1
        try:
            cache.parent.mkdir(parents=True, exist_ok=True)
            cache.write_text(json.dumps({"saved": date.today().isoformat(), "rows": rows},
                                        ensure_ascii=False, default=str), encoding="utf-8")
        except Exception as e:  # noqa: BLE001
            job.say(f"Không lưu được danh sách để chạy tiếp lần sau: {e}")
    else:
        rows = rn.list_invoices(kind, d1, d2)
    return rows, dest, rn.excel_done


def run_kind(rn: Runner, job: AutoJob, kind: dict[str, Any], d1: date, d2: date, folder: Path,
             listed: tuple[list[dict[str, Any]], Path, bool] | None = None) -> None:
    tag = f"{period(d1, d2)} {kind['label']}"
    rows, dest, excel_done = listed or list_kind(rn, job, kind, d1, d2, folder)
    if full_month(d1, d2):
        from .quick_read import mark_checked
        mark_checked(folder.parent, d1, kind, len(rows), getattr(rn, "last_total", None) if not kind.get("captcha") else None)
    job.total += len(rows)
    job.say(f"{tag}: {len(rows)} hóa đơn")
    if not rows:
        return

    if job.get_excel and not excel_done:
        try:
            rn.export_excel(kind, d1, d2, dest)
            job.excel_ok += 1
            job.say(f"{tag}: đã tải bảng kê {dest.name}")
        except Skip:
            raise
        except Exception as e:  # noqa: BLE001
            job.err(f"{tag}: {e}")

    items = []
    for inv in rows:
        name = f"{_safe(inv.get('khmshdon'))}{_safe(inv.get('khhdon'))}_{_s(inv.get('shdon')).zfill(8)}_{_safe(inv.get('nbmst'))}"
        items.append({"inv": inv, "rec": summary_row(kind, d1, inv), "parsed": None,
                      "xml_path": folder / "XML" / f"{name}.xml", "pdf_path": folder / "PDF" / f"{name}.pdf"})
    from concurrent.futures import ThreadPoolExecutor

    from .providers import easyinvoice

    ex = ThreadPoolExecutor(PDF_WORKERS)
    futs: dict[Any, dict[str, Any]] = {}

    def on_xml(it: dict[str, Any]) -> None:
        p = it.get("parsed")
        if not job.get_pdf or it["pdf_path"].exists() or not p or not has_lookup(p):
            return
        if easyinvoice.is_easyinvoice(it["rec"].get("link_tra_cuu")):
            return
        futs[ex.submit(fetch_pdf, it["inv"], p)] = it

    try:
        if job.get_xml or job.get_pdf:
            download_xmls(rn, job, kind, folder, items, tag, on_xml)
        else:
            job.done += len(items)
        if job.get_pdf:
            download_pdfs(rn, job, kind, items, tag, ex, futs, d1, d2)
    finally:
        ex.shutdown(wait=False, cancel_futures=True)
        job._rows += [it["rec"] for it in items]


def has_lookup(parsed: dict[str, Any] | None) -> bool:
    return bool(parsed and (parsed.get("lookup_code") or parsed.get("lookup_urls")))


XML_BATCH = 20
XML_BATCH_MAX = 30
PDF_WORKERS = 16
DETAIL_BATCH = 20


def xml_from_response(r: FetchResult) -> bytes:
    if not r.ok:
        raise RuntimeError(f"cổng trả lỗi {r.status} – {r.message()}")
    content = r.body
    if content[:2] == b"PK":
        with zipfile.ZipFile(io.BytesIO(content)) as zf:
            names = [n for n in zf.namelist() if n.lower().endswith(".xml")]
            if not names:
                raise RuntimeError("file zip của cổng không có XML")
            names.sort(key=lambda n: (n.lower() != "invoice.xml", n))
            return zf.read(names[0])
    if content.lstrip()[:1] == b"<":
        return content
    raise RuntimeError(f"cổng không trả XML ({r.message()})")


def _accept_xml(job: AutoJob, it: dict[str, Any], xml: bytes, on_xml) -> None:
    it["xml"] = xml
    job.xml_ok += 1
    it["rec"]["file_xml"] = str(it["xml_path"])
    try:
        p = parse_invoice_xml(xml)
        it["parsed"] = p
        it["rec"]["ma_tra_cuu"] = p.get("lookup_code") or ""
        it["rec"]["link_tra_cuu"] = (p.get("lookup_urls") or [""])[0]
        it["rec"]["_rates"] = p.get("theo_thue_suat") or {}
        it["rec"]["_items"] = p.get("hang_hoa") or []
        it["rec"]["_dchi"] = ((p.get("nguoi_ban") if it["rec"].get("huong") == "purchase"
                               else p.get("nguoi_mua")) or {}).get("dia_chi") or ""
    except Exception as e:  # noqa: BLE001
        it["rec"]["ghi_chu"] = f"XML đọc lỗi: {e}"[:300]
    job.done += 1
    if on_xml:
        on_xml(it)


def download_xmls(rn: Runner, job: AutoJob, kind: dict[str, Any], folder: Path,
                  items: list[dict[str, Any]], tag: str, on_xml=None) -> None:
    need = []
    for it in items:
        inv = it["inv"]
        if it["xml_path"].exists():
            _accept_xml(job, it, it["xml_path"].read_bytes(), on_xml)
        elif kind["direction"] == "sold" and "hsgoc" in inv and not inv.get("hsgoc"):
            it["rec"]["ghi_chu"] = "Cổng báo không tồn tại hồ sơ gốc (không có XML)"
            job.done += 1
        else:
            need.append(it)
    if need:
        job.say(f"{tag}: tải {len(need)} XML (song song {XML_BATCH} file/lần)")
    path, action = f"/{kind['ep']}/invoices/export-xml", f"Xuất xml ({kind['tab']})"
    keys = ("nbmst", "khhdon", "shdon", "khmshdon")
    failed: list[dict[str, Any]] = []
    later: list[dict[str, Any]] = []

    def get_batch(batch: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], bool]:
        res = BROWSER.fetch_many([(path, {k: _s(it["inv"].get(k)) for k in keys}, action) for it in batch])
        retry, busy = [], False
        for it, r in zip(batch, res):
            try:
                xml = xml_from_response(r)
                it["xml_path"].parent.mkdir(parents=True, exist_ok=True)
                it["xml_path"].write_bytes(xml)
                _accept_xml(job, it, xml, on_xml)
            except Exception as e:  # noqa: BLE001
                it["xml_err"] = str(e)
                if r.status in (0, 429) or r.status >= 500:
                    retry.append(it)
                    busy = busy or r.status == 429
                else:
                    failed.append(it)
        return retry, busy

    size, i, last_say = XML_BATCH, 0, 0
    while i < len(need):
        if job.cancel:
            raise Cancelled()
        batch = need[i:i + size]
        i += len(batch)
        busy = False
        for attempt in range(3):
            batch, b = get_batch(batch)
            busy = busy or b
            if not batch:
                break
            time.sleep((5 if b else 2) * (attempt + 1))
        later += batch
        size = max(5, size // 2) if busy else min(XML_BATCH_MAX, size + 5)
        time.sleep(settings.gdt_delay if not busy else 3)
        if i - last_say >= 60 or i == len(need):
            last_say = i
            job.say(f"{tag}: đã tải {i}/{len(need)} XML")

    for rnd, pause in enumerate((15, 30, 60, 90), 1):
        if not later:
            break
        job.say(f"{tag}: cổng báo quá tải – nghỉ {pause} giây rồi tải lại {len(later)} XML (lượt {rnd}/4)")
        for _ in range(pause):
            if job.cancel:
                raise Cancelled()
            time.sleep(1)
        todo, later = later, []
        for k in range(0, len(todo), 5):
            if job.cancel:
                raise Cancelled()
            retry, busy = get_batch(todo[k:k + 5])
            later += retry
            time.sleep(3 if busy else settings.gdt_delay)
        if not later:
            job.say(f"{tag}: đã tải lại đủ XML")
    failed += later

    for it in failed:
        it["rec"]["ghi_chu"] = f"Không tải được XML: {it.get('xml_err')}"[:300]
        try:
            det = rn.detail(kind, it["inv"])
            jp = it["xml_path"].with_suffix(".json")
            jp.parent.mkdir(parents=True, exist_ok=True)
            jp.write_text(json.dumps(det, ensure_ascii=False, indent=1), encoding="utf-8")
            it["rec"]["file_xml"] = str(jp)
        except (BrowserAuthError, Cancelled):
            raise
        except Exception:  # noqa: BLE001
            pass
        job.err(f"{tag} {it['inv'].get('khhdon')}-{it['inv'].get('shdon')}: không tải được XML ({it.get('xml_err')})")
        job.done += 1


def _collect(job: AutoJob, futs: dict[Any, dict[str, Any]], tag: str, label: str) -> list[dict[str, Any]]:
    from concurrent.futures import as_completed

    failed, n, ok = [], 0, 0
    for f in as_completed(list(futs)):
        if job.cancel:
            raise Cancelled()
        it, rec = futs[f], futs[f]["rec"]
        try:
            pdf, info = f.result()
        except Exception as e:  # noqa: BLE001
            pdf, info = None, {"pdf_error": str(e)[:300]}
        rec["ncc"] = info.get("provider", "") or rec.get("ncc", "")
        if pdf:
            it["pdf_path"].parent.mkdir(parents=True, exist_ok=True)
            it["pdf_path"].write_bytes(pdf)
            rec["file_pdf"], rec["ghi_chu"] = str(it["pdf_path"]), info.get("pdf_note", "")
            job.pdf_ok += 1
            ok += 1
        else:
            it["pdf_err"] = info.get("pdf_error", "")
            it["need_captcha"] = bool(info.get("needs_captcha"))
            failed.append(it)
        n += 1
        if n % 25 == 0 or n == len(futs):
            job.say(f"{tag}: {label} {n}/{len(futs)} – lấy được {ok}")
    return failed


def download_easyinvoice_mgmt(job: AutoJob, items: list[dict[str, Any]], tag: str, d1: date, d2: date) -> None:
    from .providers import easyinvoice

    mst = job.mst
    if not easyinvoice.mgmt_logged_in_quiet(mst):
        job.say(f"{tag}: {len(items)} hóa đơn bán ra EasyInvoice – dùng bản in từ cổng (muốn PDF gốc của EasyInvoice: "
                f"đăng nhập {easyinvoice.mgmt_origin(mst)} trên Chrome của phần mềm trước khi tải)")
        return
    groups: dict[str, list[dict[str, Any]]] = {}
    for it in items:
        groups.setdefault(f"{_s(it['inv'].get('khmshdon'))}{_s(it['inv'].get('khhdon'))}", []).append(it)
    ok = 0
    for pattern, its in groups.items():
        try:
            idx = easyinvoice.mgmt_index(mst, pattern, d1, d2)
        except Exception as e:  # noqa: BLE001
            job.err(f"{tag}: không lấy được danh sách EasyInvoice {pattern}: {e}")
            continue
        found = [it for it in its if str(int(_s(it["inv"].get("shdon")) or 0)) in idx]
        job.say(f"{tag}: EasyInvoice {pattern} – khớp {len(found)}/{len(its)} hóa đơn, đang tải PDF gốc hàng loạt")
        for k in range(0, len(found), 60):
            if job.cancel:
                raise Cancelled()
            chunk = found[k:k + 60]
            pdfs = easyinvoice.mgmt_download_many(
                mst, [(idx[str(int(_s(it["inv"].get("shdon"))))], pattern) for it in chunk])
            for it, pdf in zip(chunk, pdfs):
                if not pdf:
                    continue
                it["pdf_path"].parent.mkdir(parents=True, exist_ok=True)
                it["pdf_path"].write_bytes(pdf)
                rec = it["rec"]
                rec["file_pdf"], rec["ncc"], rec["ghi_chu"] = str(it["pdf_path"]), "EasyInvoice (SoftDreams)", ""
                job.pdf_ok += 1
                ok += 1
            job.say(f"{tag}: EasyInvoice {pattern} – đã tải {min(k + 60, len(found))}/{len(found)}")
    job.say(f"{tag}: lấy được {ok}/{len(items)} PDF gốc bán ra qua trang quản lý EasyInvoice")


def download_captcha_sites(job: AutoJob, items: list[dict[str, Any]], tag: str, ex=None) -> None:
    """Hóa đơn có trang tra cứu bắt captcha: tự động giải bằng OCR (dự phòng nhập tay)."""
    from concurrent.futures import ThreadPoolExecutor

    from .providers import fetch_with_captcha, lookup_host

    own_ex = ex is None
    ex = ex or ThreadPoolExecutor(PDF_WORKERS)
    groups: dict[str, list[dict[str, Any]]] = {}
    for it in items:
        it["_ref"] = make_ref(it["inv"], it.get("parsed") or {})
        groups.setdefault(lookup_host(it["_ref"]) or "?", []).append(it)
    job.say(f"{tag}: {len(items)} hóa đơn ở {len(groups)} trang tra cứu có captcha – đang tự giải OCR")

    def save(it: dict[str, Any], out) -> None:
        it["pdf_path"].parent.mkdir(parents=True, exist_ok=True)
        it["pdf_path"].write_bytes(out.result.content)
        rec = it["rec"]
        rec["file_pdf"] = str(it["pdf_path"])
        rec["ghi_chu"] = "PDF in từ trang tra cứu của NCC" if out.result.printed else ""
        job.pdf_ok += 1

    def try_captcha(it: dict[str, Any], host: str, reason: str) -> bool:
        state = {"skip": False, "attempt": 0}

        def ask(img: str, why: str) -> str:
            state["attempt"] += 1
            try:
                return job.prompt_captcha(img, f"{reason} – {why}", allow_ocr=(state["attempt"] <= 2))
            except (Skip, Cancelled):
                state["skip"] = True
                return ""
        for _attempt in range(3):
            try:
                out = fetch_with_captcha(it["_ref"], ask)
            except Exception:  # noqa: BLE001
                out = None
            if job.cancel:
                raise Cancelled()
            if out and out.result:
                save(it, out)
                return True
            if state["skip"]:
                return False
        return False

    try:
        for host, its in groups.items():
            if job.cancel:
                raise Cancelled()
            first = its[0]
            label = f"{first['inv'].get('khhdon')}-{first['inv'].get('shdon')}"
            if not try_captcha(first, host, f"{tag}: trang {host} (HĐ {label})"):
                for it in its:
                    it["pdf_err"] = f"Trang tra cứu {host} cần captcha – chưa giải được"
                job.say(f"{tag}: bỏ qua trang {host} ({len(its)} hóa đơn) – dùng bản in từ cổng")
                continue
            rest = its[1:]
            if not rest:
                continue
            job.say(f"{tag}: trang {host} – đã xác nhận captcha, đang tự tải {len(rest)} hóa đơn còn lại")
            fails = _collect(job, {ex.submit(fetch_pdf, it["inv"], it.get("parsed") or {}): it for it in rest}, tag,
                             f"PDF gốc trang {host}")
            still = [it for it in fails if it.get("need_captcha")]
            if still and job.easy_captcha:
                job.say(f"{tag}: trang {host} bắt captcha cho mỗi lần tra cứu – tự giải từng hóa đơn ({len(still)})")
                for k, it in enumerate(still, 1):
                    lb = f"{it['inv'].get('khhdon')}-{it['inv'].get('shdon')}"
                    if not try_captcha(it, host, f"{tag}: trang {host} – HĐ {lb} ({k}/{len(still)})"):
                        break
            elif still:
                for it in still:
                    it["pdf_err"] = (f"Trang {host} bắt captcha cho MỖI lần tra cứu – bật tùy chọn 'Hỏi captcha từng "
                                     "hóa đơn' để lấy PDF gốc")
                job.say(f"{tag}: trang {host} bắt captcha cho mỗi lần tra cứu – {len(still)} hóa đơn dùng bản in từ cổng "
                        "(bật 'Hỏi captcha từng hóa đơn' nếu cần PDF gốc)")
    finally:
        if own_ex:
            ex.shutdown(wait=False)


def download_pdfs(rn: Runner, job: AutoJob, kind: dict[str, Any], items: list[dict[str, Any]], tag: str,
                  ex=None, futs: dict | None = None, d1: date | None = None, d2: date | None = None) -> None:
    from .providers import easyinvoice

    futs = futs if futs is not None else {}
    for it in items:
        if it["pdf_path"].exists():
            it["rec"]["file_pdf"] = str(it["pdf_path"])
            job.pdf_ok += 1
    if futs:
        job.say(f"{tag}: đang tìm PDF gốc cho {len(futs)} hóa đơn có mã/link tra cứu (song song {PDF_WORKERS} luồng)")
    failed = _collect(job, futs, tag, "PDF gốc")
    if failed:
        job.say(f"{tag}: thử lại {len(failed)} hóa đơn chưa lấy được PDF gốc")
        time.sleep(3)
        failed = _collect(job, {ex.submit(fetch_pdf, it["inv"], it["parsed"]): it for it in failed}, tag,
                          "thử lại PDF gốc")
    cap = [it for it in failed if it.get("need_captcha") and not easyinvoice.is_easyinvoice(it["rec"].get("link_tra_cuu"))]
    if cap:
        download_captcha_sites(job, cap, tag, ex)
    for it in failed:
        if not it["rec"].get("file_pdf"):
            it["rec"]["ghi_chu"] = (it["rec"].get("ghi_chu", "") + " " + it.get("pdf_err", "")).strip()[:500]

    easy = [it for it in items if not it["rec"].get("file_pdf") and has_lookup(it.get("parsed"))
            and easyinvoice.is_easyinvoice(it["rec"].get("link_tra_cuu"))]
    own = [it for it in easy if _s(it["inv"].get("nbmst")) == job.mst]
    if own:
        download_easyinvoice_mgmt(job, own, tag, d1, d2)
    others = [it for it in easy if not it["rec"].get("file_pdf") and _s(it["inv"].get("nbmst")) != job.mst]
    if job.easy_captcha:
        download_easyinvoice(job, others, tag)
    elif others:
        for it in others:
            it["rec"]["ghi_chu"] = ("PDF gốc EasyInvoice của người bán cần captcha từng hóa đơn (tùy chọn đang tắt). "
                                    + it["rec"].get("ghi_chu", "")).strip()[:500]
        job.say(f"{tag}: {len(others)} hóa đơn EasyInvoice của người bán – dùng bản in từ cổng")
    job.pdf_missing += sum(1 for it in items if not it["rec"].get("file_pdf"))
    download_portal_pdfs(rn, job, kind, items, tag)


def download_portal_pdfs(rn: Runner, job: AutoJob, kind: dict[str, Any], items: list[dict[str, Any]],
                         tag: str) -> None:
    from .gdt import portal_pdf
    from .providers.fallback import render_pdf

    todo = [it for it in items if not it["rec"].get("file_pdf")]
    if not todo:
        return
    job.say(f"{tag}: lấy PDF từ trang hoadondientu đang mở cho {len(todo)} hóa đơn chưa có PDF gốc")
    keys = ("nbmst", "khhdon", "shdon", "khmshdon")
    need = [it for it in todo if "hdhhdvu" not in it["inv"]
            and not it["pdf_path"].with_name(it["pdf_path"].stem + "_hddt.pdf").exists()]
    path = f"/{kind['ep']}/invoices/detail"
    for k in range(0, len(need), DETAIL_BATCH):
        if job.cancel:
            raise Cancelled()
        batch = need[k:k + DETAIL_BATCH]
        for it, r in zip(batch, BROWSER.fetch_many([(path, {x: _s(it["inv"].get(x)) for x in keys}, "")
                                                     for it in batch])):
            try:
                det = r.json() if r.ok else None
                if isinstance(det, dict) and det.get("shdon"):
                    it["detail"] = det
            except Exception:  # noqa: BLE001
                pass
        time.sleep(settings.gdt_delay)
    from concurrent.futures import ThreadPoolExecutor
    sco = kind["ep"] == "sco-query"

    def printed(it: dict[str, Any], html: str | None) -> bytes:
        if html:
            try:
                return portal_pdf.print_html(html)
            except Exception:  # noqa: BLE001
                pass
        return portal_pdf.render(it.get("detail") or it["inv"], sco=sco)

    pool = ThreadPoolExecutor(PRINT_WORKERS)
    jobs: dict[int, Any] = {}
    try:
        for i, it in enumerate(todo):
            if job.cancel:
                raise Cancelled()
            hp = it["pdf_path"].with_name(it["pdf_path"].stem + "_hddt.pdf")
            if hp.exists():
                continue
            try:
                html = portal_pdf.capture(it.get("detail") or it["inv"], sco)
            except (Cancelled, BrowserAuthError):
                raise
            except Exception:  # noqa: BLE001
                html = None
            jobs[i] = pool.submit(printed, it, html)
        _finish_portal_pdfs(job, todo, jobs, tag)
    finally:
        pool.shutdown(wait=False, cancel_futures=True)


PRINT_WORKERS = 4


def _finish_portal_pdfs(job: AutoJob, todo: list[dict[str, Any]], jobs: dict[int, Any], tag: str) -> None:
    n = 0
    for i, it in enumerate(todo):
        if job.cancel:
            raise Cancelled()
        rec = it["rec"]
        hp = it["pdf_path"].with_name(it["pdf_path"].stem + "_hddt.pdf")
        try:
            note = "PDF thể hiện từ trang hoadondientu"
            if not hp.exists():
                pdf, portal_err = None, ""
                try:
                    pdf = jobs[i].result(timeout=300)
                except (Cancelled, BrowserAuthError):
                    raise
                except Exception as e:  # noqa: BLE001
                    portal_err = str(e)[:200]
                if pdf is None:
                    if not it.get("parsed"):
                        rec["ghi_chu"] = (f"Chưa có PDF (in từ cổng lỗi: {portal_err}; không có XML để dựng). "
                                          + rec.get("ghi_chu", "")).strip()[:500]
                        continue
                    pdf = render_pdf(it["parsed"])
                    note = f"PDF dựng từ XML (in từ cổng lỗi: {portal_err})"
                hp.parent.mkdir(parents=True, exist_ok=True)
                hp.write_bytes(pdf)
            rec["file_pdf"] = str(hp)
            if it.get("parsed") and (it["parsed"].get("lookup_code") or it["parsed"].get("lookup_urls")):
                note += " (chưa lấy được bản gốc của nhà cung cấp)"
            rec["ghi_chu"] = (note + ". " + rec.get("ghi_chu", "")).strip()[:500]
            job.pdf_hddt += 1
            job.pdf_missing -= 1
        except (Cancelled, BrowserAuthError):
            raise
        except Exception as e:  # noqa: BLE001
            rec["ghi_chu"] = (rec.get("ghi_chu", "") + f" | Lấy PDF lỗi: {e}").strip()[:500]
        finally:
            n += 1
            if n % 25 == 0 or n == len(todo):
                job.say(f"{tag}: PDF từ trang hoadondientu {n}/{len(todo)}")


def download_easyinvoice(job: AutoJob, items: list[dict[str, Any]], tag: str) -> None:
    """Hóa đơn EasyInvoice còn thiếu PDF: tự động giải captcha OCR (dự phòng nhập tay)."""
    from .providers import easyinvoice

    easy = [it for it in items if not it["rec"].get("file_pdf") and it.get("parsed")
            and it["parsed"].get("lookup_code") and easyinvoice.is_easyinvoice(it["rec"].get("link_tra_cuu"))]
    if not easy:
        return
    job.say(f"{tag}: {len(easy)} hóa đơn EasyInvoice – đang tự giải captcha lấy PDF gốc")
    for i, it in enumerate(easy, 1):
        inv, rec = it["inv"], it["rec"]
        label = f"HĐ {inv.get('khhdon')}-{inv.get('shdon')} ({i}/{len(easy)})"
        state = {"attempt": 0}

        def ask_easy(img: str, why: str) -> str:
            state["attempt"] += 1
            return job.prompt_captcha(img, f"{label} – {why}", allow_ocr=(state["attempt"] <= 2))

        try:
            pdf = easyinvoice.fetch(rec["link_tra_cuu"], it["parsed"]["lookup_code"], ask_easy, label)
        except Skip:
            job.say(f"{tag}: bỏ qua {len(easy) - i + 1} hóa đơn EasyInvoice còn lại")
            return
        except (Cancelled, BrowserAuthError):
            raise
        except Exception as e:  # noqa: BLE001
            rec["ghi_chu"] = f"EasyInvoice: {e}"[:500]
            job.say(f"{label}: chưa lấy được PDF ({e})")
            continue
        it["pdf_path"].parent.mkdir(parents=True, exist_ok=True)
        it["pdf_path"].write_bytes(pdf)
        rec["file_pdf"], rec["ncc"], rec["ghi_chu"] = str(it["pdf_path"]), "EasyInvoice (SoftDreams)", ""
        job.pdf_ok += 1
        job.say(f"{label}: có PDF gốc EasyInvoice")


def summary_row(kind: dict[str, Any], d1: date, inv: dict[str, Any]) -> dict[str, Any]:
    return {
        "huong": kind["direction"], "tthai_code": inv.get("tthai"),
        "loai": kind["label"], "thang": f"{d1:%m/%Y}", "ngay_lap": vn_date(inv.get("tdlap")),
        "khmshdon": _s(inv.get("khmshdon")), "khhdon": _s(inv.get("khhdon")), "shdon": _s(inv.get("shdon")),
        "nbmst": _s(inv.get("nbmst")), "nbten": _s(inv.get("nbten")),
        "nmmst": _s(inv.get("nmmst")), "nmten": _s(inv.get("nmten")),
        "tgtcthue": inv.get("tgtcthue"), "tgtthue": inv.get("tgtthue"), "tgtttbso": inv.get("tgtttbso"),
        "tthai": TTHAI.get(inv.get("tthai"), _s(inv.get("tthai"))),
        # hóa đơn gốc bị thay thế / bị điều chỉnh (để ghép cặp khi lên tờ khai)
        "khmshdgoc": _s(inv.get("khmshdgoc")), "khhdgoc": _s(inv.get("khhdgoc")), "shdgoc": _s(inv.get("shdgoc")),
        "tdlhdgoc": vn_date(inv.get("tdlhdgoc")) if inv.get("tdlhdgoc") else "",
        "file_xml": "", "file_pdf": "", "ncc": "", "ma_tra_cuu": "", "link_tra_cuu": "", "ghi_chu": "",
    }


COLS = [("loai", "Loại", 26), ("thang", "Tháng", 9), ("ngay_lap", "Ngày lập", 11),
        ("khmshdon", "Mẫu số", 7), ("khhdon", "Ký hiệu", 10), ("shdon", "Số HĐ", 10),
        ("nbmst", "MST người bán", 14), ("nbten", "Tên người bán", 34),
        ("nmmst", "MST người mua", 14), ("nmten", "Tên người mua", 34),
        ("tgtcthue", "Tiền chưa thuế", 15), ("tgtthue", "Tiền thuế", 13), ("tgtttbso", "Tổng thanh toán", 15),
        ("tthai", "Trạng thái HĐ", 14), ("file_xml", "XML", 8), ("file_pdf", "PDF", 13),
        ("ncc", "Nhà cung cấp HĐĐT", 18), ("ma_tra_cuu", "Mã tra cứu", 18),
        ("link_tra_cuu", "Link tra cứu", 30), ("ghi_chu", "Ghi chú / lỗi", 50)]


def write_month_summary(job: AutoJob, root: Path, d1: date, d2: date) -> None:
    """File tổng hợp riêng của 1 tháng, đặt trong thư mục tháng đó."""
    rows = [r for r in job._rows if r.get("thang") == f"{d1:%m/%Y}"]
    if not rows:
        return
    from . import reports

    folder = root / f"{d1:%Y-%m}"
    if full_month(d1, d2):
        labels = {KINDS[k]["label"] for k in job.kinds if k in KINDS}
        try:
            old = reports.load_period_data(root, [(d1.year, d1.month)])
        except Exception:  # noqa: BLE001
            old = []
        keep = [r for r in old if r.get("loai") and r.get("loai") not in labels]
        if keep:
            rows = keep + rows
            job._rows.extend(keep)
            job.say(f"{d1:%m/%Y}: giữ lại {len(keep)} hóa đơn các loại đã tải trước đó")
    tag = f"{d1:%Y-%m}" if full_month(d1, d2) else f"{d1:%Y-%m-%d}_den_{d2:%Y-%m-%d}"
    write_summary(job, folder, rows, f"TongHop_{tag}", by_month=False)
    try:
        reports.write_accounting_export(rows, folder / f"XuatPMKT_{tag}.xlsx")
        if full_month(d1, d2):
            reports.save_month_data(rows, folder / f".dulieu_{d1:%Y-%m}.json")
    except Exception as e:  # noqa: BLE001
        job.err(f"Không ghi được file xuất phần mềm kế toán tháng {d1:%m/%Y}: {e}")


def write_summary(job: AutoJob, root: Path, rows: list[dict[str, Any]] | None = None,
                  name: str | None = None, by_month: bool = True) -> Path | None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    main = rows is None
    rows = job._rows if rows is None else rows
    if not rows:
        return None
    wb = Workbook()
    ws = wb.active
    ws.title = "Tong hop"
    bold, fill = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
    for i, (_, title, width) in enumerate(COLS, 1):
        c = ws.cell(1, i, title)
        c.font, c.fill = bold, fill
        ws.column_dimensions[c.column_letter].width = width
    link_font = Font(color="0563C1", underline="single")
    if by_month:
        _month_sheet(wb, rows, bold, fill)
    from . import reports
    reports.add_issues_sheet(wb, rows)
    for r, rec in enumerate(rows, 2):
        for i, (k, _, _) in enumerate(COLS, 1):
            v = rec.get(k)
            if k in ("file_xml", "file_pdf"):
                label = "Mở" if v else "Không"
                if k == "file_pdf" and v and str(v).endswith("_hddt.pdf"):
                    label = "Bản cổng thuế"
                c = ws.cell(r, i, label)
                if v:
                    c.hyperlink = Path(v).relative_to(root).as_posix() if Path(v).is_relative_to(root) else v
                    c.font = link_font
            elif k == "link_tra_cuu" and v:
                c = ws.cell(r, i, v)
                c.hyperlink, c.font = v, link_font
            else:
                c = ws.cell(r, i, v)
                if k in ("tgtcthue", "tgtthue", "tgtttbso"):
                    c.number_format = "#,##0"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    if name is None:
        m1, m2 = job.months[0][0], job.months[-1][1]
        name = f"TongHop_{m1:%Y-%m-%d}_den_{m2:%Y-%m-%d}"
    root.mkdir(parents=True, exist_ok=True)
    dest = root / f"{name}.xlsx"
    try:
        wb.save(dest)
    except PermissionError:
        dest = root / f"{name}_{datetime.now():%H%M%S}.xlsx"
        wb.save(dest)
    if main:
        job.summary_file = str(dest)
        try:
            reports.write_accounting_export(rows, root / f"{dest.stem.replace('TongHop', 'XuatPMKT')}.xlsx")
        except Exception as e:  # noqa: BLE001
            job.err(f"Không ghi được file xuất phần mềm kế toán: {e}")
    return dest


def _month_sheet(wb, rows: list[dict[str, Any]], bold, fill) -> None:
    """Trang 'Theo thang': số hóa đơn, tiền, tình trạng PDF theo từng tháng và loại."""
    ws = wb.create_sheet("Theo thang", 0)
    head = ["Tháng", "Loại", "Số HĐ", "Tiền chưa thuế", "Tiền thuế", "Tổng thanh toán",
            "Có XML", "PDF gốc", "PDF bản cổng", "Thiếu PDF"]
    for i, h in enumerate(head, 1):
        c = ws.cell(1, i, h)
        c.font, c.fill = bold, fill
        ws.column_dimensions[c.column_letter].width = [10, 28, 8, 16, 14, 16, 8, 9, 12, 10][i - 1]
    agg: dict[tuple[str, str], list[float]] = {}
    for r in rows:
        k = (r.get("thang") or "", r.get("loai") or "")
        a = agg.setdefault(k, [0, 0, 0, 0, 0, 0, 0, 0])
        pdf = str(r.get("file_pdf") or "")
        a[0] += 1
        a[1] += float(r.get("tgtcthue") or 0)
        a[2] += float(r.get("tgtthue") or 0)
        a[3] += float(r.get("tgtttbso") or 0)
        a[4] += 1 if r.get("file_xml") else 0
        a[5] += 1 if pdf and not pdf.endswith("_hddt.pdf") else 0
        a[6] += 1 if pdf.endswith("_hddt.pdf") else 0
        a[7] += 0 if pdf else 1
    keyf = lambda k: (k[0][3:], k[0][:2], k[1])  # noqa: E731
    total = [0.0] * 8
    row = 2
    for k in sorted(agg, key=keyf):
        vals = agg[k]
        total = [x + y for x, y in zip(total, vals)]
        for i, v in enumerate([k[0], k[1], *vals], 1):
            c = ws.cell(row, i, v)
            if i in (4, 5, 6):
                c.number_format = "#,##0"
        row += 1
    for i, v in enumerate(["TỔNG", "", *total], 1):
        c = ws.cell(row, i, v)
        c.font = bold
        if i in (4, 5, 6):
            c.number_format = "#,##0"
    ws.freeze_panes = "A2"


def running_job() -> AutoJob | None:
    return next((j for j in JOBS.values() if j.status == "running"), None)


def start(d_from: date, d_to: date, kinds: list[str], get_excel: bool, get_xml: bool,
          get_pdf: bool, easy_captcha: bool = False) -> AutoJob:
    job = AutoJob(uuid.uuid4().hex, date_chunks(d_from, d_to), kinds, get_excel, get_xml, get_pdf, easy_captcha)
    JOBS[job.id] = job
    threading.Thread(target=run, args=(job,), daemon=True, name=f"auto-{job.id[:6]}").start()
    return job
