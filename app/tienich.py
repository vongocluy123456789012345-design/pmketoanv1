"""Tiện ích kế toán (menu ☰ Tiện ích) – dùng chung cho giao diện và Trợ lý AI.

* Bảng điều khiển tất cả khách hàng (MST): đã tải HĐ, dự báo VAT, đối chiếu sao kê, hạn nộp sắp tới, rủi ro HĐ.
* Lịch hạn nộp thuế / BHXH / quyết toán / môn bài cho từng khách hàng, đánh dấu đã nộp, nhắc trước 3 ngày.
* Kiểm tra rủi ro hóa đơn đầu vào trước khi kê khai.
* Lãi lỗ & thuế tạm tính (TNDN doanh nghiệp / thuế hộ kinh doanh).
* Biên bản đối chiếu công nợ từng khách hàng / nhà cung cấp (Excel) + lời nhắn gửi kèm.
* Định khoản hóa đơn mua vào theo mặt hàng (quy tắc + học theo cách kế toán đã hạch toán).
* Xuất XML tờ khai 01/GTGT (TT80) đúng cấu trúc HTKK kết xuất, từ kết quả dự báo VAT.
"""
from __future__ import annotations

import io
import json
import re
from calendar import monthrange
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from .config import ROOT, settings

DATA = ROOT / "data"
HTKK_DATA = Path(r"C:\Program Files (x86)\HTKK\DataFiles")
# NQ 204/2025/QH15: giảm 2% thuế GTGT (thuế suất 8%) đến hết 31/12/2026 – kiểm tra lại khi có văn bản mới
GIAM_THUE_DEN = date(2026, 12, 31)


def _fold(s: Any) -> str:
    from .bank_post import fold
    return fold(s)


def _jload(p: Path, default: Any) -> Any:
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return default


def _jsave(p: Path, d: Any) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(d, ensure_ascii=False, indent=1, default=str), encoding="utf-8")


# ------------------------------------------------------------------ khách hàng (MST) và dữ liệu
def msts() -> list[str]:
    base = Path(settings.output_dir)
    if not base.is_dir():
        return []
    return sorted(d.name for d in base.iterdir() if d.is_dir() and re.fullmatch(r"\d{10}(-?\d{3})?|\d{12}", d.name))


def months_of(mst: str) -> list[str]:
    root = Path(settings.output_dir) / mst
    return sorted(d.name for d in root.iterdir() if d.is_dir() and re.fullmatch(r"\d{4}-\d{2}", d.name)) if root.is_dir() else []


def htkk_header(mst: str) -> dict[str, str]:
    """Thông tin NNT trong HTKK (DataFiles/MST/HEADER.xml): tên, địa chỉ, tỉnh, cơ quan thuế."""
    f = HTKK_DATA / mst / "HEADER.xml"
    out: dict[str, str] = {}
    if not f.exists():
        return out
    try:
        txt = f.read_text(encoding="utf-8-sig")
        for prop, val in re.findall(r'PropertyInfo="([^"]+)"[^>]*?Value="([^"]*)"', txt):
            out.setdefault(prop, val)
    except Exception:  # noqa: BLE001
        pass
    return out


def invoices(mst: str, d1: date, d2: date) -> list[dict[str, Any]]:
    from .reports import load_period_data
    root = Path(settings.output_dir) / mst
    months, cur = [], d1.replace(day=1)
    while cur <= d2:
        months.append((cur.year, cur.month))
        cur = date(cur.year + (cur.month == 12), cur.month % 12 + 1, 1)
    out = []
    for r in load_period_data(root, months):
        try:
            d = date.fromisoformat(str(r.get("ngay_lap"))[:10])
        except ValueError:
            d = None
        if d is None or d1 <= d <= d2:
            out.append(r)
    return out


def company(mst: str) -> dict[str, str]:
    """Tên, địa chỉ công ty: HTKK -> kết quả dự báo VAT -> hóa đơn bán ra / mua vào."""
    h = htkk_header(mst)
    ten, dchi = h.get("tenNNT", ""), h.get("dchiNNT", "")
    if not ten:
        from . import vat
        for hh in vat.history(mst)[:3]:
            r = vat.saved_result(mst, hh["key"]) or {}
            ten = ten or (r.get("company") or {}).get("ten", "")
            dchi = dchi or (r.get("company") or {}).get("dchi", "")
    if not ten:
        ms = months_of(mst)
        if ms:
            y, m = map(int, ms[-1].split("-"))
            for r in invoices(mst, date(y, m, 1), date(y, m, monthrange(y, m)[1])):
                if r.get("huong") == "sold" and str(r.get("nbmst")) == mst and r.get("nbten"):
                    ten = r["nbten"]
                    break
                if r.get("huong") == "purchase" and str(r.get("nmmst")) == mst and r.get("nmten"):
                    ten = ten or r["nmten"]
    hkd = bool(re.search(r"\b(ho kinh doanh|hkd)\b", _fold(ten)))
    return {"mst": mst, "ten": ten, "dchi": dchi, "hkd": hkd, "ma_cqt": h.get("maCQTNoiNop", ""),
            "ten_cqt": h.get("tenCQTNoiNop", ""), "ma_tinh": h.get("maTinhNNT", ""), "ten_tinh": h.get("tenTinhNNT", "")}


def _kieu(mst: str) -> str:
    from . import vat
    try:
        return vat.get_cfg(mst)["kieu"]
    except Exception:  # noqa: BLE001
        return "M"


# ------------------------------------------------------------------ lịch hạn nộp
HAN_FILE = DATA / "han_nop.json"


def _workday(d: date) -> date:
    """Hạn trùng thứ 7 / chủ nhật -> ngày làm việc tiếp theo (chưa tính ngày lễ)."""
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d


def _eom(y: int, m: int) -> date:
    return date(y, m, monthrange(y, m)[1])


def deadlines(mst: str, today: date | None = None, ahead: int = 75, back: int = 40) -> list[dict[str, Any]]:
    """Các hạn nộp của 1 khách hàng trong khoảng [hôm nay - back, hôm nay + ahead]."""
    today = today or date.today()
    info = company(mst)
    kieu = _kieu(mst)
    done = _jload(HAN_FILE, {}).get(mst, {})
    items: list[dict[str, Any]] = []

    def add(loai: str, ky: str, han: date, ghi_chu: str = "") -> None:
        han = _workday(han)
        if today - timedelta(days=back) <= han <= today + timedelta(days=ahead):
            key = f"{loai}|{ky}"
            items.append({"mst": mst, "ten": info["ten"], "loai": loai, "ky": ky, "han": han.isoformat(),
                          "con_ngay": (han - today).days, "da_nop": bool(done.get(key)), "key": key, "ghi_chu": ghi_chu})

    for y in (today.year - 1, today.year, today.year + 1):
        for m in range(1, 13):
            ny, nm = (y + (m == 12), m % 12 + 1)
            if kieu == "M":
                add("Tờ khai thuế GTGT (tháng)", f"{m:02d}/{y}", date(ny, nm, 20))
                add("Tờ khai thuế TNCN (tháng, nếu có khấu trừ)", f"{m:02d}/{y}", date(ny, nm, 20))
            add("Đóng BHXH / BHYT / BHTN", f"{m:02d}/{y}", _eom(y, m), "Đóng hằng tháng chậm nhất ngày cuối tháng")
        for q in range(1, 5):
            qy, qm = (y + (q == 4), (q * 3) % 12 + 1)
            if kieu == "Q":
                add("Tờ khai thuế GTGT (quý)", f"Q{q}/{y}", _eom(qy, qm))
                add("Tờ khai thuế TNCN (quý, nếu có khấu trừ)", f"Q{q}/{y}", _eom(qy, qm))
            if not info["hkd"]:
                add("Tạm nộp thuế TNDN quý", f"Q{q}/{y}", date(qy, qm, 30), "Tạm nộp, không phải nộp tờ khai")
        add("Lệ phí môn bài", str(y), date(y, 1, 30))
        if not info["hkd"]:
            add("Quyết toán thuế TNDN + Báo cáo tài chính", str(y), date(y + 1, 3, 31))
            add("Quyết toán thuế TNCN", str(y), date(y + 1, 3, 31))
    for it in items:
        if it["loai"].startswith("Tờ khai thuế GTGT"):
            from . import vat
            ky = it["ky"]
            key = (f"{ky[-4:]}-Q{ky[1]}" if ky.startswith("Q") else f"{ky[-4:]}-{ky[:2]}")
            r = vat.saved_result(mst, key)
            if r:
                it["ghi_chu"] = (f"Đã dự báo: phải nộp {r.get('phai_nop', 0):,.0f} đ" if r.get("phai_nop") else
                                 f"Đã dự báo: còn khấu trừ {r.get('chuyen_ky_sau', 0):,.0f} đ").replace(",", ".")
    items.sort(key=lambda x: (x["han"], x["loai"]))
    return items


def all_deadlines(today: date | None = None, ahead: int = 45) -> list[dict[str, Any]]:
    out = []
    for m in msts():
        out += [x for x in deadlines(m, today, ahead, back=10) if not x["da_nop"]]
    out.sort(key=lambda x: (x["han"], x["mst"]))
    return out


def mark_done(mst: str, key: str, done: bool = True) -> dict[str, Any]:
    d = _jload(HAN_FILE, {})
    d.setdefault(mst, {})[key] = bool(done)
    _jsave(HAN_FILE, d)
    return {"ok": True, "mst": mst, "key": key, "da_nop": done}


def tick_deadlines(now: datetime | None = None) -> int:
    """Gọi định kỳ: hạn còn <= 3 ngày mà chưa đánh dấu đã nộp -> thông báo (mỗi hạn 1 lần / ngày)."""
    from . import vat
    now = now or datetime.now()
    st = _jload(DATA / "han_nop_nhac.json", {})
    if st.get("ngay") == now.date().isoformat():
        return 0
    st["ngay"] = now.date().isoformat()
    _jsave(DATA / "han_nop_nhac.json", st)
    due = [x for x in all_deadlines(now.date(), ahead=3) if x["con_ngay"] >= -3]
    if due:
        lines = [f"{x['han'][8:]}/{x['han'][5:7]} – {x['ten'] or x['mst']}: {x['loai']} {x['ky']}"
                 + (" (QUÁ HẠN)" if x["con_ngay"] < 0 else "") for x in due[:12]]
        vat.add_notification(f"{len(due)} hạn nộp trong 3 ngày tới / quá hạn chưa đánh dấu đã nộp",
                             "\n".join(lines), "warn", kind="han-nop")
    return len(due)


# ------------------------------------------------------------------ bảng điều khiển
def dashboard(today: date | None = None) -> list[dict[str, Any]]:
    from . import vat
    today = today or date.today()
    rows = []
    for m in msts():
        info = company(m)
        kieu = _kieu(m)
        ms = months_of(m)
        k1, ky, nam = vat.last_completed(today, kieu)
        key = vat.period_key(k1, ky, nam)
        cur = vat.period_key(k1, *((today.month if kieu == "M" else (today.month - 1) // 3 + 1), today.year))
        res = vat.saved_result(m, key) or {}
        res_cur = vat.saved_result(m, cur) or {}
        need = [f"{y:04d}-{mm:02d}" for y, mm in vat.period_months(k1, ky, nam)]
        dl = [x for x in deadlines(m, today, ahead=60, back=10) if not x["da_nop"]]
        bank = sorted((Path(settings.output_dir) / m / "DocChungTu").glob("KetQua_GhepThuChi_*.json")) \
            if (Path(settings.output_dir) / m / "DocChungTu").is_dir() else []
        bk = _jload(bank[-1], {}) if bank else {}
        rr = _jload(DATA / "rui_ro" / f"{m}.json", {})
        rows.append({
            "mst": m, "ten": info["ten"], "hkd": info["hkd"], "kieu": kieu, "thang_moi_nhat": ms[-1] if ms else "",
            "ky_vua_xong": vat.period_label(k1, ky, nam), "da_tai_ky_vua_xong": all(x in ms for x in need),
            "vat_ky_vua_xong": ({"phai_nop": res.get("phai_nop", 0), "chuyen_ky_sau": res.get("chuyen_ky_sau", 0),
                                 "luc": res.get("created", "")} if res else None),
            "vat_ky_nay": ({"phai_nop": res_cur.get("phai_nop", 0), "chuyen_ky_sau": res_cur.get("chuyen_ky_sau", 0),
                            "luc": res_cur.get("created", "")} if res_cur else None),
            "han_gan_nhat": dl[0] if dl else None, "so_han_qua_han": sum(1 for x in dl if x["con_ngay"] < 0),
            "sao_ke": bk or None, "rui_ro": rr.get("tong") if rr else None, "rui_ro_luc": rr.get("luc", "") if rr else ""})
    return rows


# ------------------------------------------------------------------ kiểm tra rủi ro hóa đơn đầu vào
RISKY_FILE = DATA / "mst_rui_ro.json"


def mst_valid(mst: Any) -> bool:
    """Kiểm tra số kiểm tra của MST 10 số (MST 12 số / CCCD và MST chi nhánh -xxx: chỉ xét 10 số đầu)."""
    s = re.sub(r"\D", "", str(mst or ""))
    if len(s) == 12:
        return True
    if len(s) not in (10, 13):
        return False
    w = [31, 29, 23, 19, 17, 13, 7, 5, 3]
    chk = 10 - sum(int(a) * b for a, b in zip(s[:9], w)) % 11
    return chk == int(s[9]) if chk < 10 else False


def risky_list() -> list[str]:
    return _jload(RISKY_FILE, [])


def set_risky_list(text: str) -> list[str]:
    lst = sorted(set(re.findall(r"\d{10}(?:-\d{3})?", text or "")))
    _jsave(RISKY_FILE, lst)
    return lst


def risk_check(mst: str, d1: date, d2: date) -> dict[str, Any]:
    from .bank_post import NON_CASH_LIMIT, name_sim
    from .reports import tthai_code
    info = company(mst)
    rows = [r for r in invoices(mst, d1, d2) if r.get("huong") == "purchase"]
    risky = set(risky_list())
    issues: list[dict[str, Any]] = []
    seen: dict[tuple, dict] = {}

    def add(r, muc, loai, chi_tiet):
        issues.append({"muc": muc, "loai": loai, "chi_tiet": chi_tiet, "ngay": str(r.get("ngay_lap"))[:10],
                       "hd": f"{r.get('khhdon') or ''}-{str(r.get('shdon') or '').lstrip('0')}",
                       "mst_ban": r.get("nbmst"), "ten_ban": r.get("nbten"), "tong": float(r.get("tgtttbso") or 0),
                       "loai_hd": r.get("loai", "")})

    for r in rows:
        tc = tthai_code(r)
        key = (str(r.get("nbmst")), str(r.get("khhdon")), str(r.get("shdon")).lstrip("0"))
        if key in seen:
            add(r, "LỖI", "Hóa đơn trùng", "Cùng người bán, ký hiệu, số HĐ xuất hiện 2 lần – chỉ kê khai 1 lần")
        seen[key] = r
        if tc == 6:
            add(r, "LỖI", "Hóa đơn đã hủy", "Không được kê khai khấu trừ")
        elif tc == 4:
            add(r, "LỖI", "Hóa đơn đã bị thay thế", "Kê khai theo hóa đơn thay thế, không kê khai hóa đơn này")
        elif tc == 5:
            add(r, "CẢNH BÁO", "Hóa đơn bị điều chỉnh", "Kiểm tra hóa đơn điều chỉnh tăng / giảm đi kèm")
        if r.get("nmmst") and re.sub(r"\D", "", str(r["nmmst"]))[:10] != re.sub(r"\D", "", mst)[:10]:
            add(r, "LỖI", "Sai MST người mua", f"MST người mua trên HĐ: {r.get('nmmst')} ≠ {mst}")
        elif info["ten"] and r.get("nmten") and name_sim(r["nmten"], info["ten"]) < 0.6:
            add(r, "CẢNH BÁO", "Tên người mua khác tên công ty", f"Trên HĐ: {r.get('nmten')}")
        if r.get("nbmst") and not mst_valid(r["nbmst"]):
            add(r, "CẢNH BÁO", "MST người bán sai số kiểm tra", f"MST {r.get('nbmst')} có thể ghi sai")
        if str(r.get("nbmst") or "")[:14] in risky or str(r.get("nbmst") or "")[:10] in risky:
            add(r, "LỖI", "Người bán trong danh sách rủi ro",
                "MST nằm trong danh sách doanh nghiệp rủi ro / bỏ địa chỉ kinh doanh bạn đã nhập")
        for ts, (gt, thue) in (r.get("_rates") or {}).items():
            pct = re.sub(r"[^\d.]", "", ts)
            if not pct or gt <= 0:
                continue
            rate = float(pct) / 100
            if abs(thue - gt * rate) > max(1000, gt * rate * 0.005):
                add(r, "CẢNH BÁO", "Tiền thuế không khớp thuế suất",
                    f"Thuế suất {ts}: tiền hàng {gt:,.0f} × {ts} = {gt * rate:,.0f} nhưng thuế ghi {thue:,.0f}".replace(",", "."))
            if pct == "8":
                try:
                    dd = date.fromisoformat(str(r.get("ngay_lap"))[:10])
                    if dd > GIAM_THUE_DEN:
                        add(r, "LỖI", "Thuế suất 8% hết thời hạn giảm",
                            f"HĐ lập {dd:%d/%m/%Y} sau {GIAM_THUE_DEN:%d/%m/%Y} vẫn dùng 8%")
                except ValueError:
                    pass
        if float(r.get("tgtttbso") or 0) >= NON_CASH_LIMIT and tc not in (4, 6):
            add(r, "LƯU Ý", "Cần chứng từ thanh toán không dùng tiền mặt",
                f"Tổng thanh toán từ {NON_CASH_LIMIT:,.0f} đ – giữ ủy nhiệm chi / sao kê để được khấu trừ".replace(",", "."))
    order = {"LỖI": 0, "CẢNH BÁO": 1, "LƯU Ý": 2}
    issues.sort(key=lambda x: (order[x["muc"]], x["ngay"]))
    tong = {k: sum(1 for x in issues if x["muc"] == k) for k in order}
    res = {"mst": mst, "ten": info["ten"], "tu": d1.isoformat(), "den": d2.isoformat(), "so_hd": len(rows),
           "tong": tong, "issues": issues, "luc": datetime.now().isoformat(timespec="seconds"),
           "tra_cuu_mst": "https://tracuunnt.gdt.gov.vn/tcnnt/mstdn.jsp"}
    (DATA / "rui_ro").mkdir(parents=True, exist_ok=True)
    _jsave(DATA / "rui_ro" / f"{mst}.json", {k: v for k, v in res.items() if k != "issues"})
    return res


def risk_excel(res: dict[str, Any]) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "Rui ro HD dau vao"
    ws.append([f"KIỂM TRA RỦI RO HÓA ĐƠN ĐẦU VÀO – {res['ten']} ({res['mst']}) – từ {res['tu']} đến {res['den']}"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([f"{res['so_hd']} hóa đơn mua vào · LỖI {res['tong']['LỖI']} · CẢNH BÁO {res['tong']['CẢNH BÁO']} · "
               f"LƯU Ý {res['tong']['LƯU Ý']} · Tra cứu tình trạng MST người bán: {res['tra_cuu_mst']}"])
    head = ["Mức", "Loại", "Chi tiết", "Ngày HĐ", "Ký hiệu - Số", "MST người bán", "Tên người bán", "Tổng thanh toán", "Loại HĐ"]
    ws.append(head)
    for c in ws[3]:
        c.font = Font(bold=True)
    fills = {"LỖI": "F8CBAD", "CẢNH BÁO": "FFF2CC", "LƯU Ý": "DDEBF7"}
    for x in res["issues"]:
        ws.append([x["muc"], x["loai"], x["chi_tiet"], x["ngay"], x["hd"], x["mst_ban"], x["ten_ban"], x["tong"], x["loai_hd"]])
        for c in ws[ws.max_row]:
            c.fill = PatternFill("solid", fgColor=fills[x["muc"]])
        ws.cell(ws.max_row, 8).number_format = "#,##0"
    for col, w in zip("ABCDEFGHI", (10, 34, 60, 11, 20, 15, 40, 15, 26)):
        ws.column_dimensions[col].width = w
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


# ------------------------------------------------------------------ lãi lỗ & thuế tạm tính
HKD_RATES = {  # nhóm ngành hộ kinh doanh: (% GTGT, % TNCN) trên doanh thu (TT 40/2021) – kiểm tra lại khi luật thay đổi
    "phan_phoi": ("Phân phối, cung cấp hàng hóa", 1.0, 0.5),
    "dich_vu": ("Dịch vụ, xây dựng không bao thầu nguyên vật liệu", 5.0, 2.0),
    "san_xuat": ("Sản xuất, vận tải, dịch vụ gắn với hàng hóa, xây dựng có bao thầu NVL", 3.0, 1.5),
    "khac": ("Hoạt động kinh doanh khác", 2.0, 1.0),
    "cho_thue_ts": ("Cho thuê tài sản", 5.0, 5.0),
}


def profit(mst: str, d1: date, d2: date, chi_phi_khac: float = 0, luong: float = 0, nhom_hkd: str = "phan_phoi",
           thue_suat_tndn: float | None = None) -> dict[str, Any]:
    from .reports import tthai_code
    info = company(mst)
    rows = [r for r in invoices(mst, d1, d2) if tthai_code(r) not in (4, 6)]
    sold = [r for r in rows if r.get("huong") == "sold"]
    buy = [r for r in rows if r.get("huong") == "purchase"]
    dt = sum(float(r.get("tgtcthue") or 0) for r in sold)
    vat_ra = sum(float(r.get("tgtthue") or 0) for r in sold)
    mua = sum(float(r.get("tgtcthue") or 0) for r in buy)
    vat_vao = sum(float(r.get("tgtthue") or 0) for r in buy)
    ln = dt - mua - luong - chi_phi_khac
    months = (d2.year - d1.year) * 12 + d2.month - d1.month + 1
    res = {"mst": mst, "ten": info["ten"], "hkd": info["hkd"], "tu": d1.isoformat(), "den": d2.isoformat(),
           "so_hd_ban": len(sold), "so_hd_mua": len(buy), "doanh_thu": dt, "thue_gtgt_ban_ra": vat_ra,
           "mua_vao": mua, "thue_gtgt_mua_vao": vat_vao, "luong": luong, "chi_phi_khac": chi_phi_khac,
           "loi_nhuan_truoc_thue": ln, "ghi_chu": []}
    if info["hkd"]:
        ten, g, t = HKD_RATES.get(nhom_hkd, HKD_RATES["phan_phoi"])
        res.update({"nhom_hkd": ten, "thue_gtgt_hkd": round(dt * g / 100), "thue_tncn_hkd": round(dt * t / 100),
                    "tong_thue": round(dt * (g + t) / 100)})
        res["ghi_chu"].append(f"Hộ kinh doanh – nhóm '{ten}': GTGT {g}% + TNCN {t}% trên doanh thu (tỷ lệ theo TT 40/2021,"
                              " kiểm tra lại quy định mới nhất)")
    else:
        dt_nam = dt * 12 / max(1, months)
        rate = thue_suat_tndn if thue_suat_tndn else (15 if dt_nam <= 3e9 else 17 if dt_nam <= 50e9 else 20)
        res.update({"thue_suat_tndn": rate, "doanh_thu_uoc_nam": round(dt_nam),
                    "thue_tndn_tam_tinh": round(max(0, ln) * rate / 100),
                    "thue_gtgt_uoc_phai_nop": round(max(0, vat_ra - vat_vao))})
        res["ghi_chu"].append(f"Thuế TNDN tạm tính {rate}% trên lợi nhuận (Luật TNDN 2025: 15% doanh thu năm ≤ 3 tỷ, "
                              "17% đến 50 tỷ, 20% còn lại) – chưa trừ chi phí không có hóa đơn ngoài số bạn nhập")
    res["ghi_chu"].append("Chỉ tính từ hóa đơn điện tử đã tải + lương / chi phí khác bạn nhập; chưa gồm khấu hao, giá vốn "
                          "tồn kho, chi phí không hóa đơn")
    return res


# ------------------------------------------------------------------ biên bản đối chiếu công nợ
def debt_letters(own: dict[str, str], ledger: dict[str, Any], loai: str, ngay: date, chon: list[str] | None = None,
                 ) -> tuple[bytes, list[dict[str, Any]]]:
    """Từ chi tiết công nợ (bank_smart.read_ledger) -> 1 file Excel, mỗi đối tượng 1 sheet biên bản đối chiếu công nợ."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, Side
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = Workbook()
    wb.remove(wb.active)
    msgs = []
    by: dict[str, list[dict[str, Any]]] = {}
    for e in ledger["entries"]:
        by.setdefault(e["ten"], []).append(e)
    for p in ledger["partners"]:
        if chon and p["ten"] not in chon:
            continue
        es = sorted(by.get(p["ten"], []), key=lambda e: e["ngay"] or date.min)
        ps = sum(e["phat_sinh"] for e in es)
        da = sum(e["da"] for e in es)
        cuoi = float(p.get("du") or 0)
        dau = cuoi - ps + da
        title = re.sub(r"[\\/*?:\[\]]", " ", p["ten"])[:28] or "Doi tuong"
        name = title
        k = 2
        while name in wb.sheetnames:
            name = f"{title[:25]}_{k}"
            k += 1
        ws = wb.create_sheet(name)
        rows = [["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"], ["Độc lập – Tự do – Hạnh phúc"], [],
                ["BIÊN BẢN ĐỐI CHIẾU CÔNG NỢ"], [f"Ngày {ngay:%d} tháng {ngay:%m} năm {ngay:%Y}"], [],
                [f"Bên A: {own.get('ten') or ''}" + (f" – MST {own.get('mst')}" if own.get("mst") else "")],
                [f"Địa chỉ: {own.get('dchi') or ''}"],
                [f"Bên B: {p['ten']}" + (f" – MST {p['mst']}" if p.get("mst") else "")], [],
                ["Hai bên cùng đối chiếu công nợ " + ("phải thu (Bên B nợ Bên A)" if loai == "KH" else
                                                     "phải trả (Bên A nợ Bên B)") + " như sau:"], [],
                ["Ngày", "Số chứng từ", "Diễn giải", "Phát sinh tăng", "Phát sinh giảm (đã thanh toán)", "Số còn lại"]]
        for r in rows:
            ws.append(r)
        hdr = ws.max_row
        ws.append(["", "", "Số dư đầu kỳ", None, None, dau])
        run = dau
        for e in es:
            run += e["phat_sinh"] - e["da"]
            ws.append([e["ngay"], e["so_ct"], e["dien_giai"], e["phat_sinh"] or None, e["da"] or None, run])
        ws.append(["", "", "Cộng phát sinh", ps, da, None])
        ws.append(["", "", "SỐ DƯ CUỐI KỲ", None, None, cuoi])
        for r in ws.iter_rows(min_row=hdr, max_row=ws.max_row):
            for c in r:
                c.border = box
                if c.column >= 4:
                    c.number_format = "#,##0"
                if c.column == 1 and isinstance(c.value, date):
                    c.number_format = "dd/mm/yyyy"
        for c in ws[hdr]:
            c.font = Font(bold=True)
        ws.append([])
        ws.append([f"Kết luận: tính đến ngày {ngay:%d/%m/%Y}, " + ("Bên B còn nợ Bên A" if (loai == "KH") == (cuoi >= 0)
                                                                 else "Bên A còn nợ Bên B") +
                   f" số tiền {abs(cuoi):,.0f} đồng.".replace(",", ".")])
        ws.append(["Biên bản lập thành 02 bản, mỗi bên giữ 01 bản có giá trị như nhau. Nếu có sai lệch, đề nghị Bên B phản "
                   "hồi trong 05 ngày làm việc."])
        ws.append([])
        ws.append(["ĐẠI DIỆN BÊN A", "", "", "ĐẠI DIỆN BÊN B"])
        ws.append(["(Ký, ghi rõ họ tên, đóng dấu)", "", "", "(Ký, ghi rõ họ tên, đóng dấu)"])
        for r in (1, 2, 4, 5):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
            ws.cell(r, 1).alignment = Alignment(horizontal="center")
            ws.cell(r, 1).font = Font(bold=True, size=14 if r == 4 else 11)
        for col, w in zip("ABCDEF", (12, 16, 46, 17, 20, 17)):
            ws.column_dimensions[col].width = w
        msgs.append({"ten": p["ten"], "so_du": cuoi, "loi_nhan":
                     f"Dạ kính gửi {p['ten']}, {own.get('ten') or 'bên em'} xin gửi biên bản đối chiếu công nợ tính đến ngày "
                     f"{ngay:%d/%m/%Y}: số dư {'quý công ty còn nợ' if (loai == 'KH') == (cuoi >= 0) else 'bên em còn nợ'} "
                     f"{abs(cuoi):,.0f} đồng. Nhờ anh/chị kiểm tra, ký xác nhận và gửi lại giúp bên em. Em cảm ơn ạ!"
                     .replace(",", ".")})
    if not wb.sheetnames:
        wb.create_sheet("Trong")
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue(), msgs


# ------------------------------------------------------------------ định khoản hóa đơn mua vào theo mặt hàng
TK_FILE = DATA / "tk_hoc.json"
_TK_RULES = [  # (từ khóa tên hàng / người bán đã bỏ dấu, khóa, diễn giải)
    (r"van chuyen|giao hang|cuoc (xe|tau|van)|ship|logistic|chuyen phat|buu chinh|boc xep", "cp_dv", "Chi phí vận chuyển"),
    (r"xang (ron|e5|a95|ron95)|\bxang\b dau|dau diesel|dau do\b|nhien lieu", "cp_dv", "Xăng dầu"),
    (r"tien dien|dien nang|nuoc sach|tien nuoc|internet|cap quang|vien thong|cuoc (dien thoai|di dong)", "cp_dv",
     "Điện, nước, viễn thông"),
    (r"quang cao|marketing|facebook|google|tiktok ads|truyen thong|in an to roi|banner", "cp_bh", "Chi phí quảng cáo"),
    (r"thue (nha|van phong|kho|mat bang)|tien thue|cho thue", "cp_dv", "Thuê nhà / mặt bằng"),
    (r"phan mem|ban quyen|hosting|ten mien|chu ky so|token|license|email doanh nghiep|cloud", "cp_tra_truoc",
     "Phần mềm / dịch vụ trả trước"),
    (r"van phong pham|giay (in|a4|a3)|muc in|but|so sach|kep|ghim|bia ho so", "cp_dv", "Văn phòng phẩm"),
    (r"an uong|tiep khach|nha hang|khach san|ve may bay|cong tac", "cp_dv", "Tiếp khách / công tác"),
    (r"bao hiem (xe|tai san|chay no)", "cp_tra_truoc", "Bảo hiểm tài sản"),
    (r"sua chua|bao duong|bao tri", "cp_dv", "Sửa chữa, bảo dưỡng"),
    (r"phi (dich vu|tu van|kiem toan|ke toan)|tu van|kiem toan|dich vu ke toan", "cp_dv", "Phí dịch vụ / tư vấn"),
    (r"nguyen lieu|vat lieu|phu tung|linh kien|bao bi|thung carton|tem nhan", "nvl", "Nguyên vật liệu / bao bì"),
]
_TK = {"TT200": {"cp_dv": "6427", "cp_bh": "6418", "cp_tra_truoc": "242", "nvl": "152", "ccdc": "153", "tscd": "211",
                 "hang": "1561"},
       "TT133": {"cp_dv": "6422", "cp_bh": "6421", "cp_tra_truoc": "242", "nvl": "152", "ccdc": "153", "tscd": "211",
                 "hang": "1561"}}
_ASSET = r"\b(may|thiet bi|xe|o to|laptop|may tinh|dieu hoa|tu lanh|camera|may in|may anh|may chup)\b"


def _item_key(name: str) -> str:
    w = [x for x in _fold(name).split() if not any(c.isdigit() for c in x)]
    return " ".join(w[:4])


def suggest_tk(ten: str, don_gia: float = 0, nguoi_ban: str = "", che_do: str = "TT200", mac_dinh: str = "",
               tai_san: bool = True) -> tuple[str, str]:
    """TK nợ cho 1 dòng hàng mua vào: theo cách kế toán đã hạch toán (đã học) -> quy tắc từ khóa -> mặc định hàng hóa."""
    acc = _TK.get((che_do or "TT200").upper(), _TK["TT200"])
    mem = _jload(TK_FILE, {})
    k = _item_key(ten)
    if k and k in mem:
        tk = max(mem[k], key=mem[k].get)
        return tk, f"Đã học: '{k}' → {tk} ({mem[k][tk]} lần)"
    f = _fold(f"{ten} {nguoi_ban}")
    for pat, key, txt in _TK_RULES:
        if re.search(pat, f):
            return acc[key], txt
    if tai_san and re.search(_ASSET, _fold(ten)):
        if don_gia >= 30_000_000:
            return acc["tscd"], "Tài sản cố định (đơn giá từ 30 triệu)"
        return acc["ccdc"], "Công cụ dụng cụ (đơn giá dưới 30 triệu – phân bổ qua 242)"
    if mac_dinh:
        return mac_dinh, f"TK mặc định bạn chọn ({mac_dinh})"
    return acc["hang"], "Hàng hóa (mặc định)"


def learn_tk(pairs: list[tuple[str, str]]) -> int:
    """Lưu cách hạch toán: [(tên hàng, TK)] -> các lần sau gợi ý đúng TK đó."""
    mem = _jload(TK_FILE, {})
    n = 0
    for ten, tk in pairs:
        k, tk = _item_key(ten), str(tk or "").strip()
        if k and re.fullmatch(r"\d{3,5}", tk):
            mem.setdefault(k, {})
            mem[k][tk] = mem[k].get(tk, 0) + 1
            n += 1
    _jsave(TK_FILE, mem)
    return n


def learn_tk_from_file(name: str, data: bytes) -> int:
    """Học từ file mua hàng đã hạch toán (mẫu MISA mua hàng có cột Tên hàng + TK kho / TK chi phí / TK Nợ)."""
    from . import misa_import as mi
    wb, ws, hr, cols, _ = mi.read_template(data, values=True)
    c_ten = next((c for c, k in cols.items() if k == "ten_hang"), None)
    c_tk = next((c for c, k in cols.items() if k in ("tk_kho", "tk_no")), None)
    if not c_ten or not c_tk:
        raise ValueError("File cần có cột Tên hàng và TK kho / TK chi phí / TK Nợ")
    pairs = []
    for r in range(hr + 1, ws.max_row + 1):
        ten, tk = ws.cell(r, c_ten).value, ws.cell(r, c_tk).value
        if ten and tk:
            pairs.append((str(ten), str(tk).split(".")[0]))
    return learn_tk(pairs)


def purchase_lines(mst: str, d1: date, d2: date, che_do: str = "TT200", mac_dinh: str = "") -> list[dict[str, Any]]:
    """Từng dòng hàng của HĐ mua vào (đọc XML đã tải) + TK gợi ý."""
    from .misa_import import from_xml
    from .reports import tthai_code
    out = []
    for r in invoices(mst, d1, d2):
        if r.get("huong") != "purchase" or tthai_code(r) in (4, 6):
            continue
        items = []
        if r.get("file_xml") and Path(str(r["file_xml"])).exists():
            try:
                items = from_xml(Path(str(r["file_xml"])).read_bytes()).get("items") or []
            except Exception:  # noqa: BLE001
                items = []
        if not items:
            items = [{"ten": f"(HĐ {r.get('shdon')} – chưa có XML chi tiết)", "dvt": "", "sl": None, "dg": 0,
                      "tt": float(r.get("tgtcthue") or 0), "thue": float(r.get("tgtthue") or 0)}]
        for it in items:
            tk, why = suggest_tk(it.get("ten") or "", float(it.get("dg") or 0), r.get("nbten") or "", che_do, mac_dinh)
            out.append({"ngay": str(r.get("ngay_lap"))[:10], "hd": f"{r.get('khhdon') or ''}-{str(r.get('shdon') or '').lstrip('0')}",
                        "nguoi_ban": r.get("nbten"), "mst_ban": r.get("nbmst"), "ten": it.get("ten") or "",
                        "dvt": it.get("dvt") or "", "sl": it.get("sl"), "don_gia": it.get("dg"),
                        "thanh_tien": float(it.get("tt") or 0), "thue": float(it.get("thue") or 0), "tk": tk, "ly_do": why})
    return out


# ------------------------------------------------------------------ XML tờ khai 01/GTGT (cấu trúc HTKK kết xuất)
def vat_xml(mst: str, key: str, nguoi_ky: str = "") -> tuple[Path, str]:
    """Xuất XML tờ khai 01/GTGT (TT80, mã 842) từ kết quả dự báo VAT đã lập – cùng cấu trúc file HTKK kết xuất để nộp
    qua thuedientu. Trả (đường dẫn file, ghi chú)."""
    from xml.sax.saxutils import escape as e

    from . import vat
    r = vat.saved_result(mst, key)
    if not r:
        raise ValueError(f"Chưa có kết quả dự báo VAT kỳ {key} của {mst} – bấm Dự báo VAT trước")
    info = company(mst)
    ct = r.get("ct") or {}
    v = lambda k: str(int(round(float(ct.get(k) or 0))))  # noqa: E731
    kieu, ky, nam = r["kieu"], r["ky"], r["nam"]
    d1, d2 = date.fromisoformat(r["from"]), date.fromisoformat(r["to"])
    ky_txt = f"{ky}/{nam}" if kieu == "Q" else f"{ky:02d}/{nam}"
    today = date.today().isoformat()
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue id="ID_1">
    <TTinChung>
      <TTinDVu>
        <maDVu>HTKK</maDVu>
        <tenDVu>HỖ TRỢ KÊ KHAI THUẾ</tenDVu>
        <pbanDVu>5.7.9</pbanDVu>
        <ttinNhaCCapDVu>TỔNG CỤC THUẾ HÀ NỘI</ttinNhaCCapDVu>
      </TTinDVu>
      <TTinTKhaiThue>
        <TKhaiThue>
          <maTKhai>842</maTKhai>
          <tenTKhai>TỜ KHAI THUẾ GIÁ TRỊ GIA TĂNG (Mẫu số 01/GTGT)</tenTKhai>
          <moTaBMau>(Ban hành kèm theo Thông tư số 80/2021/TT-BTC ngày 29 tháng 9 năm 2021 của Bộ trưởng Bộ Tài chính)</moTaBMau>
          <pbanTKhaiXML>2.8.3</pbanTKhaiXML>
          <loaiTKhai>C</loaiTKhai>
          <soLan>0</soLan>
          <KyKKhaiThue>
            <kieuKy>{kieu}</kieuKy>
            <kyKKhai>{ky_txt}</kyKKhai>
            <kyKKhaiTuNgay>{d1:%d/%m/%Y}</kyKKhaiTuNgay>
            <kyKKhaiDenNgay>{d2:%d/%m/%Y}</kyKKhaiDenNgay>
            <kyKKhaiTuThang />
            <kyKKhaiDenThang />
          </KyKKhaiThue>
          <maCQTNoiNop>{e(info['ma_cqt'])}</maCQTNoiNop>
          <tenCQTNoiNop>{e(info['ten_cqt'])}</tenCQTNoiNop>
          <ngayLapTKhai>{today}</ngayLapTKhai>
          <GiaHan>
            <maLyDoGiaHan />
            <lyDoGiaHan />
          </GiaHan>
          <nguoiKy>{e(nguoi_ky)}</nguoiKy>
          <ngayKy>{today}</ngayKy>
          <nganhNgheKD />
        </TKhaiThue>
        <NNT>
          <mst>{e(mst)}</mst>
          <tenNNT>{e(info['ten'])}</tenNNT>
          <dchiNNT>{e(info['dchi'])}</dchiNNT>
          <phuongXa />
          <maHuyenNNT />
          <tenHuyenNNT />
          <maTinhNNT>{e(info['ma_tinh'])}</maTinhNNT>
          <tenTinhNNT>{e(info['ten_tinh'])}</tenTinhNNT>
          <dthoaiNNT />
          <faxNNT />
          <emailNNT />
        </NNT>
      </TTinTKhaiThue>
    </TTinChung>
    <CTieuTKhaiChinh>
      <ma_NganhNghe>00</ma_NganhNghe>
      <ten_NganhNghe>Hoạt động sản xuất kinh doanh thông thường</ten_NganhNghe>
      <tieuMucHachToan>1701</tieuMucHachToan>
      <Header>
        <ct09 />
        <ct10 />
        <DiaChiHDSXKDKhacTinhNDTSC>
          <ct11a_phuongXa_ma />
          <ct11a_phuongXa_ten />
          <ct11b_quanHuyen_ma />
          <ct11b_quanHuyen_ten />
          <ct11c_tinhTP_ma />
          <ct11c_tinhTP_ten />
        </DiaChiHDSXKDKhacTinhNDTSC>
      </Header>
      <ct21>{1 if r.get('ct21') else 0}</ct21>
      <ct22>{v('22')}</ct22>
      <GiaTriVaThueGTGTHHDVMuaVao>
        <ct23>{v('23')}</ct23>
        <ct24>{v('24')}</ct24>
      </GiaTriVaThueGTGTHHDVMuaVao>
      <HangHoaDichVuNhapKhau>
        <ct23a>{v('23a')}</ct23a>
        <ct24a>{v('24a')}</ct24a>
      </HangHoaDichVuNhapKhau>
      <ct25>{v('25')}</ct25>
      <ct26>{v('26')}</ct26>
      <HHDVBRaChiuThueGTGT>
        <ct27>{v('27')}</ct27>
        <ct28>{v('28')}</ct28>
      </HHDVBRaChiuThueGTGT>
      <ct29>{v('29')}</ct29>
      <HHDVBRaChiuTSuat5>
        <ct30>{v('30')}</ct30>
        <ct31>{v('31')}</ct31>
      </HHDVBRaChiuTSuat5>
      <HHDVBRaChiuTSuat10>
        <ct32>{v('32')}</ct32>
        <ct33>{v('33')}</ct33>
      </HHDVBRaChiuTSuat10>
      <ct32a>{v('32a')}</ct32a>
      <TongDThuVaThueGTGTHHDVBRa>
        <ct34>{v('34')}</ct34>
        <ct35>{v('35')}</ct35>
      </TongDThuVaThueGTGTHHDVBRa>
      <ct36>{v('36')}</ct36>
      <ct37>{v('37')}</ct37>
      <ct38>{v('38')}</ct38>
      <ct39a>{v('39a')}</ct39a>
      <ct40a>{v('40a')}</ct40a>
      <ct40b>{v('40b')}</ct40b>
      <ct40>{v('40')}</ct40>
      <ct41>{v('41')}</ct41>
      <ct42>{v('42')}</ct42>
      <ct43>{v('43')}</ct43>
    </CTieuTKhaiChinh>
  </HSoKhaiThue>
</HSoThueDTu>
"""
    out = vat.out_folder(mst, key) / f"ToKhai_01GTGT_{key}_HTKK.xml"
    out.write_text(xml, encoding="utf-8")
    notes = []
    if not info["ma_cqt"]:
        notes.append("Chưa có mã cơ quan thuế (lấy từ HTKK) – mở file bằng HTKK để chọn cơ quan thuế trước khi nộp")
    if float(ct.get("33") or 0) and (r.get("pl8") or {}).get("giam"):
        notes.append("Kỳ này có giảm thuế 8% – lập phụ lục 01-10/MGT trên HTKK đi kèm")
    notes.append("Kiểm tra lại trên HTKK / thuedientu trước khi ký nộp")
    return out, " · ".join(notes)
