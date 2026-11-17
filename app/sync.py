"""Điều phối đồng bộ: cổng thuế -> XML -> PDF gốc -> Supabase."""
from __future__ import annotations

import threading
import traceback
import uuid
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from typing import Any

from . import db
from .config import settings
from .dates import vn_date
from .gdt.client import GDTAuthError, GDTClient
from .gdt.xml_parser import parse_invoice_xml
from .providers import InvoiceRef, fetch_original_pdf, load_providers
from .providers.fallback import render_pdf

DONE_PDF = ("original", "manual")


@dataclass
class Job:
    id: str
    owner_mst: str
    direction: str
    date_from: date
    date_to: date
    fetch_pdf: bool
    status: str = "running"
    total: int = 0
    done: int = 0
    pdf_ok: int = 0
    pdf_rendered: int = 0
    pdf_failed: int = 0
    errors: list[str] = field(default_factory=list)
    log: list[str] = field(default_factory=list)
    started_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    finished_at: str | None = None

    def say(self, msg: str) -> None:
        self.log.append(f"{datetime.now():%H:%M:%S} {msg}")
        self.log = self.log[-300:]

    def public(self) -> dict[str, Any]:
        d = {k: v for k, v in self.__dict__.items()}
        d["date_from"] = self.date_from.isoformat()
        d["date_to"] = self.date_to.isoformat()
        d["log"] = self.log[-60:]
        return d


JOBS: dict[str, Job] = {}


def _str(v: Any) -> str:
    return "" if v is None else str(v)


def _storage_base(row: dict[str, Any]) -> str:
    month = vn_date(row.get("tdlap"))[:7] or "unknown"
    safe = lambda s: "".join(c if c.isalnum() else "-" for c in _str(s))  # noqa: E731
    return (f"{row['owner_mst']}/{row['direction']}/{month}/"
            f"{safe(row['nbmst'])}_{safe(row['khmshdon'])}{safe(row['khhdon'])}_{safe(row['shdon'])}")


def _ref_from(row: dict[str, Any], parsed: dict[str, Any]) -> InvoiceRef:
    return InvoiceRef(
        seller_mst=row["nbmst"],
        buyer_mst=row.get("nmmst") or (parsed.get("nguoi_mua") or {}).get("mst"),
        khmshdon=row["khmshdon"], khhdon=row["khhdon"], shdon=row["shdon"],
        msttcgp=row.get("msttcgp") or parsed.get("msttcgp"),
        mccqt=parsed.get("mccqt") or row.get("mhdon"),
        lookup_code=parsed.get("lookup_code"),
        lookup_urls=parsed.get("lookup_urls") or [],
    )


def process_pdf(row: dict[str, Any], parsed: dict[str, Any], providers=None) -> dict[str, Any]:
    """Lấy PDF gốc cho 1 hóa đơn (đã có XML). Trả về các trường cần cập nhật."""
    ref = _ref_from(row, parsed)
    outcome = fetch_original_pdf(ref, providers)
    upd: dict[str, Any] = {
        "provider_code": outcome.provider.code if outcome.provider else None,
        "lookup_code": ref.lookup_code,
        "lookup_url": ref.lookup_urls[0] if ref.lookup_urls else None,
    }
    base = _storage_base(row)
    if outcome.result:
        db.upload(base + ".pdf", outcome.result.content, "application/pdf")
        upd.update(pdf_path=base + ".pdf", pdf_status="original", pdf_is_original=True,
                   pdf_source_url=outcome.result.source_url, pdf_error=None)
        return upd

    reason = "; ".join(a.get("error") or a.get("result") or "" for a in outcome.attempts)[:900]
    if outcome.needs_captcha:
        reason = "Trang tra cứu cần captcha – hãy mở link tra cứu, tải PDF và upload. " + reason
    upd["pdf_error"] = reason or "Không tìm thấy PDF gốc"
    if settings.render_fallback:
        try:
            pdf = render_pdf(parsed)
            db.upload(base + ".view.pdf", pdf, "application/pdf")
            upd.update(pdf_path=base + ".view.pdf", pdf_status="rendered", pdf_is_original=False)
            return upd
        except Exception as e:  # noqa: BLE001
            upd["pdf_error"] += f" | Dựng PDF dự phòng lỗi: {e}"
    upd["pdf_status"] = "failed"
    return upd


def _row_from_listing(owner_mst: str, direction: str, inv: dict[str, Any]) -> dict[str, Any]:
    return {
        "owner_mst": owner_mst,
        "direction": direction,
        "nbmst": _str(inv.get("nbmst")),
        "khmshdon": _str(inv.get("khmshdon")),
        "khhdon": _str(inv.get("khhdon")),
        "shdon": _str(inv.get("shdon")),
        "tdlap": inv.get("tdlap"),
        "nbten": inv.get("nbten"),
        "nmmst": inv.get("nmmst"),
        "nmten": inv.get("nmten"),
        "tgtcthue": inv.get("tgtcthue"),
        "tgtthue": inv.get("tgtthue"),
        "tgtttbso": inv.get("tgtttbso"),
        "tthai": inv.get("tthai"),
        "ttxly": inv.get("ttxly"),
        "mhdon": inv.get("mhdon"),
        "msttcgp": inv.get("msttcgp"),
        "source_query": inv.get("_query"),
        "raw": inv,
    }


def run_job(job: Job, gdt: GDTClient) -> None:
    providers = load_providers()
    try:
        db.upsert_company(job.owner_mst)
        job.say("Đang lấy danh sách hóa đơn từ cổng thuế...")
        listing = list(gdt.iter_invoices(job.direction, job.date_from, job.date_to))
        job.total = len(listing)
        job.say(f"Tìm thấy {job.total} hóa đơn")

        for inv in listing:
            label = f"{inv.get('khhdon')}-{inv.get('shdon')} ({inv.get('nbten') or inv.get('nbmst')})"
            try:
                row = db.upsert_invoice(_row_from_listing(job.owner_mst, job.direction, inv))
                parsed = row.get("parsed")
                if not row.get("xml_path") or not parsed:
                    xml = gdt.export_xml(inv)
                    parsed = parse_invoice_xml(xml)
                    path = _storage_base(row) + ".xml"
                    db.upload(path, xml, "application/xml")
                    row = db.update_invoice(row["id"], {"xml_path": path, "parsed": parsed,
                                                        "lookup_code": parsed.get("lookup_code")})
                if job.fetch_pdf and row.get("pdf_status") not in DONE_PDF:
                    upd = process_pdf(row, parsed, providers)
                    db.update_invoice(row["id"], upd)
                    st = upd.get("pdf_status")
                    if st == "original":
                        job.pdf_ok += 1
                    elif st == "rendered":
                        job.pdf_rendered += 1
                    else:
                        job.pdf_failed += 1
                    job.say(f"{label}: PDF {st}")
                elif row.get("pdf_status") in DONE_PDF:
                    job.pdf_ok += 1
            except GDTAuthError:
                raise
            except Exception as e:  # noqa: BLE001 - 1 hóa đơn lỗi không dừng cả job
                job.errors.append(f"{label}: {e}")
                job.say(f"LỖI {label}: {e}")
            job.done += 1
        job.status = "done"
        job.say("Hoàn tất")
    except GDTAuthError as e:
        job.status = "auth_error"
        job.errors.append(str(e))
        job.say(str(e))
    except Exception as e:  # noqa: BLE001
        job.status = "error"
        job.errors.append(f"{e}\n{traceback.format_exc()[-1500:]}")
        job.say(f"LỖI: {e}")
    finally:
        _finish(job)


def _finish(job: Job) -> None:
    job.finished_at = datetime.now(timezone.utc).isoformat()
    try:
        db.insert_job({
            "owner_mst": job.owner_mst, "direction": job.direction,
            "date_from": job.date_from.isoformat(), "date_to": job.date_to.isoformat(),
            "status": job.status, "total": job.total, "pdf_ok": job.pdf_ok,
            "pdf_failed": job.pdf_failed,
            "log": {"errors": job.errors[-50:], "rendered": job.pdf_rendered},
            "started_at": job.started_at, "finished_at": job.finished_at,
        })
    except Exception:  # noqa: BLE001
        pass


def start_job(gdt: GDTClient, owner_mst: str, direction: str, d_from: date, d_to: date,
              fetch_pdf: bool) -> Job:
    job = Job(uuid.uuid4().hex, owner_mst, direction, d_from, d_to, fetch_pdf)
    JOBS[job.id] = job
    threading.Thread(target=run_job, args=(job, gdt), daemon=True).start()
    return job


# ======================= Nhập file từ cổng thuế =======================
def run_import_job(job: Job, files: list[tuple[str, bytes]]) -> None:
    from .sources.file_import import iter_items, read_excel, row_from_excel, row_from_xml

    providers = load_providers()
    try:
        db.upsert_company(job.owner_mst)
        xmls, excel_rows = [], []
        for name, data in files:
            for kind, item_name, payload in iter_items(name, data):
                if kind == "xml":
                    xmls.append((item_name, payload))
                elif kind == "excel":
                    try:
                        rows = read_excel(payload, item_name)
                        excel_rows += rows
                        job.say(f"{item_name}: đọc được {len(rows)} dòng hóa đơn")
                        if not rows:
                            job.errors.append(f"{item_name}: không nhận ra cột 'Số hóa đơn' trong file Excel")
                    except Exception as e:  # noqa: BLE001
                        job.errors.append(f"{item_name}: lỗi đọc Excel: {e}")
                else:
                    job.say(f"Bỏ qua {item_name}: {payload.decode('utf-8', 'ignore')}")
        job.total = len(xmls) + len(excel_rows)
        job.say(f"Tổng cộng {len(xmls)} XML, {len(excel_rows)} dòng Excel")

        # Excel trước (chỉ danh sách), XML sau để bổ sung dữ liệu đầy đủ cho cùng hóa đơn.
        for r in excel_rows:
            label = f"{r.get('khhdon')}-{r.get('shdon')}"
            try:
                db.upsert_invoice(row_from_excel(job.owner_mst, r, job.direction))
            except Exception as e:  # noqa: BLE001
                job.errors.append(f"Excel {label}: {e}")
            job.done += 1

        for item_name, xml in xmls:
            try:
                parsed = parse_invoice_xml(xml)
                row = db.upsert_invoice(row_from_xml(job.owner_mst, parsed, job.direction))
                label = f"{row['khhdon']}-{row['shdon']} ({row.get('nbten') or row['nbmst']})"
                path = _storage_base(row) + ".xml"
                db.upload(path, xml, "application/xml")
                row = db.update_invoice(row["id"], {"xml_path": path})
                if job.fetch_pdf and row.get("pdf_status") not in DONE_PDF:
                    upd = process_pdf(row, parsed, providers)
                    db.update_invoice(row["id"], upd)
                    st = upd.get("pdf_status")
                    if st == "original":
                        job.pdf_ok += 1
                    elif st == "rendered":
                        job.pdf_rendered += 1
                    else:
                        job.pdf_failed += 1
                    job.say(f"{label}: PDF {st}")
                else:
                    job.say(f"{label}: đã lưu")
            except Exception as e:  # noqa: BLE001
                job.errors.append(f"{item_name}: {e}")
                job.say(f"LỖI {item_name}: {e}")
            job.done += 1
        job.status = "done"
        job.say("Hoàn tất")
    except Exception as e:  # noqa: BLE001
        job.status = "error"
        job.errors.append(f"{e}\n{traceback.format_exc()[-1500:]}")
        job.say(f"LỖI: {e}")
    finally:
        _finish(job)


def start_import_job(owner_mst: str, direction_hint: str, files: list[tuple[str, bytes]],
                     fetch_pdf: bool) -> Job:
    today = date.today()
    job = Job(uuid.uuid4().hex, owner_mst, direction_hint, today, today, fetch_pdf)
    JOBS[job.id] = job
    threading.Thread(target=run_import_job, args=(job, files), daemon=True).start()
    return job


# ======================= Nguồn TConnect =======================
def tconnect_client():
    from .sources.tconnect import TConnectClient, TConnectConfig
    s = settings
    return TConnectClient(TConnectConfig(
        s.tconnect_auth_base, s.tconnect_partner_code, s.tconnect_client_id,
        s.tconnect_client_secret, s.tconnect_aes_key, s.tconnect_username, s.tconnect_password))


def run_tconnect_job(job: Job) -> None:
    from .sources.tconnect import TConnectAuthError, to_parsed, to_row

    providers = load_providers()
    tc = None
    try:
        company = db.get_company(job.owner_mst) or {}
        account_id = company.get("tconnect_account_id")
        if not account_id:
            raise TConnectAuthError("MST này chưa được kết nối với TConnect")
        tc = tconnect_client()
        directions = ["purchase", "sold"] if job.direction == "both" else [job.direction]
        d_from, d_to = job.date_from.isoformat(), job.date_to.isoformat()

        job.say(f"Yêu cầu TConnect đồng bộ {', '.join(directions)} {d_from} → {d_to}...")
        tj = tc.start_sync(account_id, d_from, d_to, directions, fetch_details=True)
        last = {}

        def progress(j):
            key = (j.get("status"), j.get("totalFound"), j.get("totalSaved"))
            if key != last.get("k"):
                last["k"] = key
                job.say(f"TConnect: {j.get('status')} – tìm thấy {j.get('totalFound') or 0}, "
                        f"đã lưu {j.get('totalSaved') or 0}")

        tc.wait_job(tj["id"], progress)

        listing = []
        for d in directions:
            for inv in tc.iter_invoices(account_id, d, d_from, d_to):
                inv.setdefault("direction", d)
                listing.append(inv)
        job.total = len(listing)
        job.say(f"Nhận {job.total} hóa đơn từ TConnect")

        for inv in listing:
            label = f"{inv.get('invoiceSymbol')}-{inv.get('invoiceNumber')} ({inv.get('sellerName')})"
            try:
                row = db.upsert_invoice(to_row(job.owner_mst, inv))
                parsed = row.get("parsed")
                if not parsed or not parsed.get("hang_hoa"):
                    detail = tc.get_invoice(inv["id"])  # tính phí 10%/lần -> chỉ gọi khi chưa có
                    detail.setdefault("direction", inv.get("direction"))
                    parsed = to_parsed(detail)
                    row = db.update_invoice(row["id"], {
                        "parsed": parsed, "lookup_code": parsed.get("lookup_code"),
                        "nmten": detail.get("buyerName"), "raw": detail,
                    })
                if job.fetch_pdf and row.get("pdf_status") not in DONE_PDF:
                    upd = process_pdf(row, parsed, providers)
                    db.update_invoice(row["id"], upd)
                    st = upd.get("pdf_status")
                    if st == "original":
                        job.pdf_ok += 1
                    elif st == "rendered":
                        job.pdf_rendered += 1
                    else:
                        job.pdf_failed += 1
                    job.say(f"{label}: PDF {st}")
                elif row.get("pdf_status") in DONE_PDF:
                    job.pdf_ok += 1
            except Exception as e:  # noqa: BLE001
                job.errors.append(f"{label}: {e}")
                job.say(f"LỖI {label}: {e}")
            job.done += 1
        job.status = "done"
        job.say("Hoàn tất")
    except TConnectAuthError as e:
        job.status = "auth_error"
        job.errors.append(str(e))
        job.say(str(e))
    except Exception as e:  # noqa: BLE001
        job.status = "error"
        job.errors.append(f"{e}\n{traceback.format_exc()[-1500:]}")
        job.say(f"LỖI: {e}")
    finally:
        if tc:
            tc.close()
        _finish(job)


def start_tconnect_job(owner_mst: str, direction: str, d_from: date, d_to: date,
                       fetch_pdf: bool) -> Job:
    job = Job(uuid.uuid4().hex, owner_mst, direction, d_from, d_to, fetch_pdf)
    JOBS[job.id] = job
    threading.Thread(target=run_tconnect_job, args=(job,), daemon=True).start()
    return job
