"""Lấy PDF gốc từ trang tra cứu của NCC mà KHÔNG cần đăng nhập cổng thuế.

* lookup_one(): người dùng dán link tra cứu + mã tra cứu (in trên hóa đơn / email) -> tải PDF về máy.
* Lấy bù hàng loạt: quét các XML đã tải trong thư mục lưu, hóa đơn nào chưa có PDF gốc (chưa có hoặc chỉ có
  bản in từ cổng _hddt.pdf) thì đọc mã/link tra cứu trong XML và lấy PDF gốc chạy ngầm.
"""
from __future__ import annotations

import re
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from .config import settings
from .gdt.xml_parser import parse_invoice_xml
from .providers import InvoiceRef, fetch_original_pdf


def _safe(s: Any) -> str:
    return re.sub(r"[^0-9A-Za-z]+", "-", str(s or "")).strip("-")


def ref_from_parsed(p: dict[str, Any]) -> InvoiceRef:
    nb = p.get("nguoi_ban") or {}
    return InvoiceRef(seller_mst=nb.get("mst") or "", buyer_mst=(p.get("nguoi_mua") or {}).get("mst"),
                      khmshdon=p.get("khmshdon") or "", khhdon=p.get("khhdon") or "", shdon=p.get("shdon") or "",
                      msttcgp=p.get("msttcgp"), mccqt=p.get("mccqt"), lookup_code=p.get("lookup_code"),
                      lookup_urls=p.get("lookup_urls") or [])


def lookup_one(url: str, code: str, seller_mst: str = "", khhdon: str = "", shdon: str = "") -> dict[str, Any]:
    """Tra cứu 1 hóa đơn theo link + mã. Lưu vào <nơi lưu>/TraCuuPDF/."""
    url = (url or "").strip()
    if url and not re.match(r"https?://", url, re.I):
        url = "http://" + url
    ref = InvoiceRef(seller_mst=seller_mst.strip(), buyer_mst=None, khmshdon="", khhdon=khhdon.strip().upper(),
                     shdon=shdon.strip(), msttcgp=None, mccqt=None, lookup_code=(code or "").strip() or None,
                     lookup_urls=[url] if url else [])
    out = fetch_original_pdf(ref)
    attempts = [a.get("error") or a.get("result") or "" for a in out.attempts]
    if not out.result:
        return {"ok": False, "needs_captcha": out.needs_captcha, "provider": out.provider.name if out.provider else "",
                "attempts": [a for a in attempts if a]}
    folder = Path(settings.output_dir) / "TraCuuPDF"
    folder.mkdir(parents=True, exist_ok=True)
    name = "_".join(x for x in (_safe(seller_mst), _safe(khhdon), _safe(shdon), _safe(code)) if x) or "hoadon"
    dest = folder / f"{name}_{datetime.now():%Y%m%d_%H%M%S}.pdf"
    dest.write_bytes(out.result.content)
    return {"ok": True, "file": str(dest), "source": out.result.source_url, "printed": out.result.printed,
            "provider": out.provider.name if out.provider else ""}


# ---------------------------------------------------------------- lấy bù hàng loạt
@dataclass
class FillJob:
    id: str
    mst: str
    months: list[str]
    status: str = "running"
    total: int = 0
    done: int = 0
    ok: int = 0
    captcha: int = 0
    log: list[str] = field(default_factory=list)
    cancel: bool = False

    def say(self, m: str) -> None:
        self.log.append(f"{datetime.now():%H:%M:%S} {m}")
        self.log = self.log[-300:]

    def public(self) -> dict[str, Any]:
        return {k: getattr(self, k) for k in ("id", "mst", "months", "status", "total", "done", "ok", "captcha")} | {
            "log": self.log[-80:]}


FILL_JOBS: dict[str, FillJob] = {}


def missing_items(root: Path, months: list[str]) -> list[tuple[Path, Path]]:
    """[(xml, pdf đích)] của hóa đơn có XML nhưng chưa có PDF gốc."""
    out = []
    for m in months:
        for xml in sorted((root / m).glob("*/XML/*.xml")):
            pdf = xml.parent.parent / "PDF" / (xml.stem + ".pdf")
            if not pdf.exists():
                out.append((xml, pdf))
    return out


def _run_fill(job: FillJob) -> None:
    from .providers.easyinvoice import is_easyinvoice

    root = Path(settings.output_dir) / job.mst
    try:
        items = missing_items(root, job.months)
        todo = []
        for xml, pdf in items:
            try:
                p = parse_invoice_xml(xml.read_bytes())
            except Exception:  # noqa: BLE001
                continue
            if not (p.get("lookup_code") or p.get("lookup_urls")):
                continue
            if all(is_easyinvoice(u) for u in p.get("lookup_urls") or [""]) and p.get("lookup_urls"):
                continue  # EasyInvoice: trang tra cứu luôn bắt captcha
            todo.append((xml, pdf, p))
        job.total = len(todo)
        job.say(f"{len(items)} hóa đơn chưa có PDF gốc, {len(todo)} hóa đơn có mã/link tra cứu để lấy ngầm "
                f"(bỏ qua EasyInvoice vì luôn bắt captcha)")

        def one(xml: Path, pdf: Path, p: dict[str, Any]):
            return fetch_original_pdf(ref_from_parsed(p))

        with ThreadPoolExecutor(6) as ex:
            futs = {ex.submit(one, *t): t for t in todo}
            for f in as_completed(futs):
                if job.cancel:
                    ex.shutdown(wait=False, cancel_futures=True)
                    job.status = "cancelled"
                    return
                xml, pdf, _p = futs[f]
                job.done += 1
                try:
                    o = f.result()
                except Exception as e:  # noqa: BLE001
                    job.say(f"{xml.name}: lỗi {e}")
                    continue
                if o.result:
                    pdf.parent.mkdir(parents=True, exist_ok=True)
                    pdf.write_bytes(o.result.content)
                    job.ok += 1
                    job.say(f"{xml.stem}: có PDF gốc" + (" (in từ trang NCC)" if o.result.printed else ""))
                elif o.needs_captcha:
                    job.captcha += 1
                    job.say(f"{xml.stem}: trang tra cứu bắt captcha – cần tải tay")
                if job.done % 20 == 0:
                    job.say(f"Đã xử lý {job.done}/{job.total}, lấy được {job.ok}")
        job.status = "done"
        job.say(f"Xong: lấy được {job.ok}/{job.total} PDF gốc, {job.captcha} hóa đơn cần captcha")
    except Exception as e:  # noqa: BLE001
        job.status = "error"
        job.say(f"LỖI {e}")


def start_fill(mst: str, months: list[str]) -> FillJob:
    job = FillJob(uuid.uuid4().hex, mst, months)
    FILL_JOBS[job.id] = job
    threading.Thread(target=_run_fill, args=(job,), daemon=True, name="pdf-fill").start()
    return job
