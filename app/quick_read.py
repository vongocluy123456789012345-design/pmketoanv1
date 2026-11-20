"""Đọc nhanh hóa đơn trên cổng thuế để dự báo VAT ngay – không tải XML, PDF.

Đọc ĐỦ 5 loại để không sót: mua vào có mã, mua vào không mã (CQT nhận không mã), mua vào máy tính tiền,
bán ra, bán ra máy tính tiền.
* Loại thường: lấy DANH SÁCH qua API tra cứu (vài giây). Mỗi dòng đã có tổng tiền theo từng thuế suất (thttltsuat).
* Máy tính tiền: cổng bắt captcha -> dùng danh sách đã lưu (nếu tháng đã kết thúc và đã đọc trước đó), nếu không thì hỏi
  khách 1 captcha / tháng / loại (lấy qua bảng kê Excel của cổng). Dòng lấy từ Excel chưa có thuế suất -> lấy "chi tiết"
  ngay để có thuế suất chính xác.
* Đối chiếu số hóa đơn đọc được với TỔNG cổng báo; ghi nhận loại nào đã đọc đủ (.kiemtra_YYYY-MM.json) để dự báo VAT
  cảnh báo nếu còn loại chưa đọc / đọc thiếu.
Bước 2 (chạy nền): lấy chi tiết từng mặt hàng song song -> lập lại dự báo VAT + bảng kê đầy đủ, báo ở chuông.
"""
from __future__ import annotations

import json
import threading
import time
from calendar import monthrange
from datetime import date, datetime
from pathlib import Path
from typing import Any

from .config import settings

ALL_KINDS = ["mua-vao", "mua-vao-khong-ma", "mua-vao-mtt", "ban-ra", "ban-ra-mtt"]
QUICK_KINDS = ALL_KINDS
DETAIL_BATCH = 25
STATE: dict[str, Any] = {}  # tiến độ bước 2 theo MST


def _rates(inv: dict[str, Any]) -> dict[str, list[float]]:
    out: dict[str, list[float]] = {}
    for e in inv.get("thttltsuat") or []:
        k = str(e.get("tsuat") or e.get("ltsuat") or "").upper()
        a = out.setdefault(k, [0.0, 0.0])
        a[0] += float(e.get("thtien") or 0)
        a[1] += float(e.get("tthue") or 0)
    return out


def kind_key(kind: dict[str, Any]) -> str:
    from .auto_download import KINDS
    return next((k for k, v in KINDS.items() if v is kind or v["label"] == kind.get("label")), "")


def mark_checked(month_dir: Path, d1: date, kind: dict[str, Any], count: int, total: int | None) -> None:
    """Ghi nhận đã đọc đủ 1 loại hóa đơn của 1 tháng (số đọc được, tổng cổng báo)."""
    f = Path(month_dir) / f".kiemtra_{d1:%Y-%m}.json"
    try:
        d = json.loads(f.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        d = {}
    d[kind_key(kind)] = {"label": kind["label"], "so_hd": count, "tong_cong": total,
                         "luc": datetime.now().isoformat(timespec="seconds")}
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


def coverage(root: Path, months: list[tuple[int, int]], rows: list[dict[str, Any]] | None = None) -> list[str]:
    """Báo rõ từng tháng x từng loại hóa đơn (mua vào có mã / không mã / máy tính tiền, bán ra / bán ra máy tính
    tiền): đã đọc bao nhiêu HĐ, đủ so với cổng chưa, loại nào CHƯA đọc, loại nào đọc THIẾU bao nhiêu HĐ."""
    from collections import Counter

    from .auto_download import KINDS
    have = Counter((r.get("thang"), r.get("loai")) for r in (rows or []))
    out = []
    for y, m in months:
        f = Path(root) / f"{y:04d}-{m:02d}" / f".kiemtra_{y:04d}-{m:02d}.json"
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
        except (FileNotFoundError, ValueError):
            d = {}
        parts, miss, short = [], [], []
        for k in ALL_KINDS:
            lab = KINDS[k]["label"]
            n = have.get((f"{m:02d}/{y}", lab), 0)
            v = d.get(k)
            # loại đã tải đầy đủ từ trước (có thư mục loại trong tháng) coi như đã đọc
            folder = (Path(root) / f"{y:04d}-{m:02d}" / KINDS[k]["folder"]).is_dir()
            if v is None and not folder:
                miss.append(lab)
                parts.append(f"{lab}: CHƯA đọc")
            elif v and v.get("tong_cong") and v["so_hd"] < v["tong_cong"]:
                short.append(f"{lab}: đọc {v['so_hd']}/{v['tong_cong']} – THIẾU {v['tong_cong'] - v['so_hd']} HĐ")
                parts.append(f"{lab}: {n} HĐ (thiếu {v['tong_cong'] - v['so_hd']})")
            else:
                tc = f" – đủ theo cổng ({v['tong_cong']})" if v and v.get("tong_cong") is not None else ""
                parts.append(f"{lab}: {n} HĐ{tc}")
        if miss:
            out.append(f"Tháng {m:02d}/{y} CHƯA đọc: {', '.join(miss)} – dự báo có thể thiếu hóa đơn loại này "
                       "(bấm ⚡ Đọc nhanh từ cổng).")
        if short:
            out.append(f"Tháng {m:02d}/{y} đọc THIẾU so với cổng: {'; '.join(short)} – bấm ⚡ Đọc nhanh từ cổng để đọc lại.")
        out.append(f"Tháng {m:02d}/{y} – hóa đơn đã đọc theo từng loại: {'; '.join(parts)}.")
    return out


def _name(ad, inv: dict[str, Any]) -> str:
    return f"{ad._safe(inv.get('khmshdon'))}{ad._safe(inv.get('khhdon'))}_{ad._s(inv.get('shdon')).zfill(8)}_{ad._safe(inv.get('nbmst'))}"


def read_period(mst: str, kieu: str, ky: int, nam: int, details: bool = True,
                kinds: list[str] | None = None, inline: bool = False) -> dict[str, Any]:
    """Đọc nhanh hóa đơn cả kỳ trên Chrome đang đăng nhập cổng (đúng MST). Dữ liệu lưu vào .dulieu_YYYY-MM.json để dự báo
    VAT dùng ngay. Loại máy tính tiền có thể hiện ô captcha cho khách nhập (mục Tải hóa đơn, có tiếng báo)."""
    from . import auto_download as ad
    from . import reports, vat
    from .gdt.browser import BROWSER

    token = BROWSER.thread_token()      # dự báo VAT nhiều công ty: luồng này dùng phiên của 1 trang HĐĐT
    if token:
        if BROWSER.token_info(token)["mst"] != mst:
            return {"loi": f"Trang HĐĐT đang đăng nhập MST {BROWSER.token_info(token)['mst']}, không phải {mst}."}
        if any(j.status == "running" and j.mst == mst for j in ad.JOBS.values()):
            return {"loi": f"MST {mst} đang có lượt tải / đọc khác chạy – chờ xong."}
    else:
        st = BROWSER.status()
        if not st.get("logged_in"):
            return {"loi": "Chưa đăng nhập cổng thuế – gọi open_portal và nhờ khách tự đăng nhập."}
        if str(st.get("mst")) != mst:
            return {"loi": f"Chrome đang đăng nhập MST {st.get('mst')}, không phải {mst}. Nhờ khách đăng xuất và đăng "
                           f"nhập bằng MST {mst}."}
        if ad.busy():
            return {"loi": "Đang có lượt tải hóa đơn khác chạy – chờ xong rồi đọc nhanh."}
    t0 = time.time()
    kinds = [k for k in (kinds or ALL_KINDS) if k in ad.KINDS]
    months = [(y, m) for y, m in vat.period_months(kieu, ky, nam) if date(y, m, 1) <= date.today()]
    job = ad.AutoJob("quick-" + mst, [(date(y, m, 1), date(y, m, monthrange(y, m)[1])) for y, m in months], kinds,
                     False, False, False)
    job.mst = mst
    job.token, job.slot = token or "", BROWSER.thread_slot() if token else 0
    ad.JOBS[job.id] = job  # để giao diện hiện ô captcha (máy tính tiền) như khi tải thường
    job.say(f"Đọc nhanh hóa đơn {vat.period_label(kieu, ky, nam)} – MST {mst}")
    rn = ad.Runner(job)
    rn.timeout, rn.retries = 25, 1  # cổng bình thường trả trong 1–3 giây; quá tải thì báo sớm thay vì chờ lâu
    root = Path(settings.output_dir) / mst
    today = date.today()
    stat: dict[str, Any] = {"mua_vao": 0, "ban_ra": 0, "thang": [], "theo_loai": {}, "canh_bao": []}
    pending: list[tuple[dict[str, Any], dict[str, Any], Path]] = []  # (kind, inv, đường dẫn json chi tiết)
    try:
        for y, m in months:
            d1 = date(y, m, 1)
            d2 = min(date(y, m, monthrange(y, m)[1]), today)
            full = d2.day == monthrange(y, m)[1]
            rows, need_now = [], []
            for key in kinds:
                kind = ad.KINDS[key]
                folder = root / f"{d1:%Y-%m}" / kind["folder"]
                try:
                    if kind.get("captcha"):
                        job.say(f"{m:02d}/{y} {kind['label']}: cần captcha của cổng (nếu chưa đọc tháng này)")
                        listed, _dest, _x = ad.list_kind(rn, job, kind, d1, d2, folder)
                        total = None
                    else:
                        listed = rn.list_invoices(kind, d1, d2)
                        total = getattr(rn, "last_total", None)
                except ad.Skip as e:
                    stat["canh_bao"].append(f"{m:02d}/{y} {kind['label']}: chưa đọc ({e})")
                    continue
                except Exception as e:  # noqa: BLE001
                    msg = str(e)
                    if any(c in msg for c in ("504", "502", "503", "Gateway", "AbortError", "aborted", "Không gọi được", "quá lâu")):
                        return {"loi": "Cổng hoadondientu.gdt.gov.vn đang quá tải (lỗi 504 phía cổng thuế) nên chưa đọc "
                                       "được hóa đơn. Báo khách thử lại sau 5–10 phút. KHÔNG gọi lại ngay và KHÔNG chuyển "
                                       "sang start_download (tải đầy đủ còn lâu hơn).", "cong_qua_tai": True}
                    if isinstance(e, ad.BrowserAuthError):
                        return {"loi": "Phiên đăng nhập cổng thuế đã hết hạn – nhờ khách đăng nhập lại."}
                    return {"loi": f"Không đọc được danh sách {kind['label']} {m:02d}/{y}: {msg[:200]}"}
                if total is not None and len(listed) < total:
                    stat["canh_bao"].append(f"{m:02d}/{y} {kind['label']}: đọc {len(listed)}/{total} hóa đơn")
                for inv in listed:
                    rec = ad.summary_row(kind, d1, inv)
                    rec["_rates"] = _rates(inv)
                    name = _name(ad, inv)
                    xml, js = folder / "XML" / f"{name}.xml", folder / "XML" / f"{name}.json"
                    if xml.exists():
                        rec["file_xml"] = str(xml)  # đã tải đầy đủ trước đó
                    elif js.exists():
                        rec["file_xml"] = str(js)
                    elif not rec["_rates"]:
                        need_now.append((kind, inv, js, rec))  # dòng từ bảng kê Excel: lấy chi tiết ngay để có thuế suất
                    else:
                        pending.append((kind, inv, js))
                    rows.append(rec)
                    stat["mua_vao" if kind["direction"] == "purchase" else "ban_ra"] += 1
                stat["theo_loai"][f"{m:02d}/{y} {kind['label']}"] = len(listed)
                if full:
                    mark_checked(root / f"{d1:%Y-%m}", d1, kind, len(listed), total)
            if need_now:
                job.say(f"{m:02d}/{y}: lấy thuế suất của {len(need_now)} hóa đơn máy tính tiền")
                _fetch_details([(k, inv, js) for k, inv, js, _ in need_now])
                for k, inv, js, rec in need_now:
                    if js.exists():
                        rec["file_xml"] = str(js)
                        rec["_rates"] = _rates(json.loads(js.read_text(encoding="utf-8")))
            # dữ liệu cho dự báo ngay (ngày lưu = hôm nay -> trọn tháng nếu tháng đã qua)
            dest = root / f"{d1:%Y-%m}" / f".dulieu_{d1:%Y-%m}.json"
            try:  # giữ lại link PDF / XML của hóa đơn đã tải đầy đủ trước đó
                old = {reports._row_key(r): r for r in json.loads(dest.read_text(encoding="utf-8")).get("rows", [])}
            except (FileNotFoundError, ValueError):
                old = {}
            for r in rows:
                o = old.get(reports._row_key(r)) or {}
                r["file_pdf"] = r.get("file_pdf") or o.get("file_pdf") or ""
                if str(o.get("file_xml") or "").endswith(".xml"):
                    r["file_xml"] = o["file_xml"]
            reports.save_month_data(rows, dest)
            stat["thang"].append(f"{m:02d}/{y}" + ("" if full else f" (đến {d2:%d/%m})"))
    finally:
        job.status = "done"
        job.captcha = None
    stat["giay"] = round(time.time() - t0, 1)
    stat["can_lay_chi_tiet"] = len(pending)
    if inline:   # bên gọi tự lập dự báo: lấy chi tiết ngay trên luồng này (giữ đúng phiên của trang HĐĐT)
        STATE[mst] = {"total": len(pending), "done": 0, "status": "running"}
        job.say(f"Lấy chi tiết từng mặt hàng của {len(pending)} hóa đơn")
        try:
            if pending:
                _fetch_details(pending, lambda n: STATE[mst].__setitem__("done", n))
                _relink(mst, kieu, ky, nam)
            STATE[mst]["status"] = "done"
        except Exception as e:  # noqa: BLE001 - vẫn lập dự báo bằng dữ liệu đã có
            STATE[mst]["status"] = f"error: {e}"
            stat["canh_bao"].append(f"Chưa lấy đủ chi tiết mặt hàng: {str(e)[:150]}")
        return stat
    if details and pending:
        STATE[mst] = {"total": len(pending), "done": 0, "status": "running"}
        threading.Thread(target=_details, args=(mst, kieu, ky, nam, pending), daemon=True, name="quick-detail").start()
    elif details:
        STATE[mst] = {"total": 0, "done": 0, "status": "done"}
        _export_async(mst, kieu, ky, nam)
    return stat


def _fetch_details(items: list[tuple[dict[str, Any], dict[str, Any], Path]], progress=None) -> None:
    """Lấy 'chi tiết hóa đơn' song song theo lô, lưu JSON (có từng mặt hàng + thuế suất)."""
    from . import auto_download as ad
    from .gdt.browser import BROWSER
    keys = ("nbmst", "khhdon", "shdon", "khmshdon")
    for i in range(0, len(items), DETAIL_BATCH):
        batch = items[i:i + DETAIL_BATCH]
        res = BROWSER.fetch_many([(f"/{k['ep']}/invoices/detail", {x: ad._s(inv.get(x)) for x in keys}, "")
                                  for k, inv, _ in batch])
        for (_k, _inv, js), r in zip(batch, res):
            try:
                d = r.json() if r.ok else None
            except Exception:  # noqa: BLE001
                d = None
            if isinstance(d, dict) and d.get("shdon"):
                js.parent.mkdir(parents=True, exist_ok=True)
                js.write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")
        if progress:
            progress(min(i + DETAIL_BATCH, len(items)))
        time.sleep(settings.gdt_delay)


def _details(mst: str, kieu: str, ky: int, nam: int, pending: list) -> None:
    """Bước 2: lấy chi tiết từng hóa đơn song song (từng mặt hàng) rồi lập lại dự báo + bảng kê đầy đủ."""
    s = STATE[mst]
    try:
        _fetch_details(pending, lambda n: s.__setitem__("done", n))
        _relink(mst, kieu, ky, nam)
        s["status"] = "done"
    except Exception as e:  # noqa: BLE001 - phiên hết hạn, mất mạng... vẫn lập dự báo bằng dữ liệu đã có
        s["status"] = f"error: {e}"
    _export_async(mst, kieu, ky, nam, wait=True)


def _relink(mst: str, kieu: str, ky: int, nam: int) -> None:
    """Gắn file chi tiết vừa lấy vào dữ liệu tháng (để bảng kê có từng mặt hàng)."""
    from . import auto_download as ad
    from . import vat
    root = Path(settings.output_dir) / mst
    for y, m in vat.period_months(kieu, ky, nam):
        f = root / f"{y:04d}-{m:02d}" / f".dulieu_{y:04d}-{m:02d}.json"
        if not f.exists():
            continue
        d = json.loads(f.read_text(encoding="utf-8"))
        for r in d.get("rows", []):
            if r.get("file_xml"):
                continue
            kind = next((k for k in ad.KINDS.values() if k["label"] == r.get("loai")), None)
            if not kind:
                continue
            name = f"{ad._safe(r.get('khmshdon'))}{ad._safe(r.get('khhdon'))}_{str(r.get('shdon')).zfill(8)}_{ad._safe(r.get('nbmst'))}"
            js = root / f"{y:04d}-{m:02d}" / kind["folder"] / "XML" / f"{name}.json"
            if js.exists():
                r["file_xml"] = str(js)
        f.write_text(json.dumps(d, ensure_ascii=False, default=str), encoding="utf-8")


def _export_async(mst: str, kieu: str, ky: int, nam: int, wait: bool = False) -> None:
    from . import vat

    def run() -> None:
        try:
            vat.forecast(mst, kieu, ky, nam, source="doc-nhanh")
        except Exception as e:  # noqa: BLE001
            vat.add_notification(f"Không lập được bảng kê {vat.period_label(kieu, ky, nam)} – MST {mst}", str(e), "warn")
    if wait:
        run()
    else:
        threading.Thread(target=run, daemon=True, name="vat-export").start()
