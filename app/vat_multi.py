"""Dự báo VAT NHIỀU CÔNG TY CÙNG LÚC trên các "trang HĐĐT" (mỗi trang 1 cửa sổ Chrome đăng nhập 1 công ty).

Mỗi trang chạy 1 luồng riêng: đọc nhanh hóa đơn kỳ cần dự báo trên cổng bằng phiên của trang đó (giữ token lúc bấm),
lấy chi tiết mặt hàng, rồi lập dự báo tờ khai 01/GTGT như mục Dự báo VAT (dùng số đã nhập / [22] kỳ trước đã lưu).
"""
from __future__ import annotations

import threading
import traceback
from datetime import date, datetime
from typing import Any

from . import auto_download as ad
from . import quick_read, vat
from .gdt.browser import BROWSER

RUNS: dict[int, dict[str, Any]] = {}


def period_of(thang: str, kieu: str) -> tuple[str, int, int]:
    """thang 'YYYY-MM' -> (kiểu, kỳ, năm); kê khai quý thì lấy quý chứa tháng đó."""
    y, m = (int(x) for x in thang.split("-")[:2])
    return (kieu, m, y) if kieu == "M" else (kieu, (m - 1) // 3 + 1, y)


def _ten(c: Any) -> str:
    return str(c.get("ten") or "") if isinstance(c, dict) else str(c or "")


def _lite(s: dict[str, Any]) -> dict[str, Any]:
    ct = s.get("ct") or {}
    return {"label": s.get("label_den") or s.get("label"), "phai_nop": s.get("phai_nop", 0),
            "chuyen_ky_sau": s.get("chuyen_ky_sau", 0), "ct22": (s.get("inputs") or {}).get("ct22", 0),
            "thue_ra": ct.get("35", 0), "thue_vao": ct.get("25", 0), "doanh_thu": ct.get("34", 0),
            "n_items": s.get("n_items"), "warnings": (s.get("warnings") or [])[:6], "missing": (s.get("missing") or [])[:6],
            "files": s.get("files") or {}, "company": _ten(s.get("company")), "tom_tat": s.get("tom_tat", ""),
            "chua_het_ky": s.get("chua_het_ky", False), "key": s.get("key"), "kieu": s.get("kieu"), "ky": s.get("ky"),
            "nam": s.get("nam")}


def _run(n: int, run: dict[str, Any], token: str, read_portal: bool, kinds: list[str] | None = None) -> None:
    mst, kieu, ky, nam = run["mst"], run["kieu"], run["ky"], run["nam"]
    BROWSER.use_slot(n)
    BROWSER.use_token(token)
    try:
        if read_portal:
            run["buoc"] = "Đọc hóa đơn trên cổng"
            st = quick_read.read_period(mst, kieu, ky, nam, inline=True, kinds=kinds)
            if st.get("loi"):
                raise RuntimeError(st["loi"])
            run["doc"] = {"mua_vao": st.get("mua_vao", 0), "ban_ra": st.get("ban_ra", 0), "giay": st.get("giay"),
                          "canh_bao": st.get("canh_bao", [])[:6], "thang": st.get("thang", [])}
        run["buoc"] = "Lập dự báo tờ khai"
        res = vat.forecast(mst, kieu, ky, nam, source="nhieu-cty", notify=False)
        run["result"] = _lite(res)
        run["status"] = "done"
        run["buoc"] = "Xong"
    except Exception as e:  # noqa: BLE001
        run["status"] = "error"
        run["error"] = str(e).splitlines()[0][:400] if str(e) else repr(e)
        run["trace"] = traceback.format_exc()[-800:]
        run["buoc"] = "Lỗi"
    finally:
        run["xong_luc"] = datetime.now().strftime("%H:%M:%S")
        BROWSER.use_token(None)
        BROWSER.use_slot(0)


def start(n: int, thang: str, kieu_chon: str = "auto", read_portal: bool = True, mtt: bool = True) -> dict[str, Any]:
    cur = RUNS.get(n)
    if cur and cur.get("status") == "running":
        return {"n": n, "ok": False, "msg": "Đang dự báo"}
    st = BROWSER.slots_status([n]).get(n, {})
    if not st.get("logged_in"):
        return {"n": n, "ok": False, "msg": "Chưa đăng nhập"}
    token = BROWSER.slot_token(n)
    mst = BROWSER.token_info(token)["mst"] if token else ""
    if not mst:
        return {"n": n, "ok": False, "msg": "Chưa đăng nhập"}
    if any(r.get("status") == "running" and r.get("mst") == mst for r in RUNS.values()):
        return {"n": n, "ok": False, "msg": f"MST {mst} đang dự báo ở trang khác"}
    if read_portal and any(j.status == "running" and j.mst == mst for j in ad.JOBS.values()):
        return {"n": n, "ok": False, "msg": f"MST {mst} đang tải hóa đơn – chờ xong"}
    kieu = kieu_chon if kieu_chon in ("M", "Q") else vat.get_cfg(mst)["kieu"]
    kieu, ky, nam = period_of(thang, kieu)
    today = date.today()
    if vat.period_range(kieu, ky, nam)[0] > today:
        return {"n": n, "ok": False, "msg": "Kỳ chưa bắt đầu"}
    run = {"n": n, "mst": mst, "kieu": kieu, "ky": ky, "nam": nam, "label": vat.period_label(kieu, ky, nam),
           "status": "running", "buoc": "Bắt đầu", "bat_dau": datetime.now().strftime("%H:%M:%S"), "result": None,
           "error": "", "doc": None, "doc_cong": read_portal}
    RUNS[n] = run
    kinds = None if mtt else [k for k in quick_read.ALL_KINDS if not ad.KINDS[k].get("captcha")]
    threading.Thread(target=_run, args=(n, run, token, read_portal, kinds), daemon=True, name=f"vat-multi-{n}").start()
    return {"n": n, "ok": True, "mst": mst, "label": run["label"]}


def public(n: int) -> dict[str, Any] | None:
    run = RUNS.get(n)
    if not run:
        return None
    out = {k: v for k, v in run.items() if k != "trace"}
    job = ad.JOBS.get("quick-" + run["mst"])
    if job:
        out["log"] = job.log[-80:]
        out["captcha"] = job.captcha
        out["job_id"] = job.id
    d = quick_read.STATE.get(run["mst"]) if run["status"] == "running" else None
    if d and d.get("status") == "running":
        out["chi_tiet"] = {"done": d.get("done", 0), "total": d.get("total", 0)}
    return out
