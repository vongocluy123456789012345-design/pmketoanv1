"""Dự báo thuế GTGT (tờ khai 01/GTGT – TT80/2021, phương pháp khấu trừ) từ hóa đơn mua vào / bán ra đã tải.

* compute(): cộng hóa đơn trong kỳ (tháng / quý) theo từng chỉ tiêu [21]…[43] của tờ khai 01/GTGT.
  - Bỏ hóa đơn "đã bị thay thế" (4) và "đã bị hủy" (6).
  - Mua vào: chỉ hóa đơn GTGT (mẫu 1) có thuế suất 0/5/8/10%… vào [23]/[24]; hóa đơn bán hàng (mẫu 2),
    HHDV không chịu thuế (KCT), không kê khai tính nộp (KKKNT) liệt kê riêng, không khấu trừ.
  - Bán ra: KCT -> [26]; 0% -> [29]; 5% -> [30]/[31]; 10% và 8% (được giảm) -> [32]/[33]; KKKNT -> [32a].
  - Bán ra 8%: lập Phụ lục giảm thuế GTGT (Mẫu 01, kèm tờ khai 01/GTGT); mua vào 8% có bảng kê tham khảo.
* export(): Excel (tờ khai, phụ lục, bảng kê hóa đơn, bảng kê chi tiết từng mặt hàng, cảnh báo) + HTML/PDF để in.
* Thông báo trong app (data/notifications.json) + tự dự báo lúc 00:00 ngày đầu kỳ sau (xem auto_tick()).

ĐÂY LÀ BẢN DỰ BÁO để kiểm tra số liệu, không thay thế tờ khai nộp qua eTax / HTKK.
"""
from __future__ import annotations

import json
import re
import threading
import uuid
from calendar import monthrange
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from .config import ROOT, settings

STATE_FILE = ROOT / "data" / "vat.json"
NOTI_FILE = ROOT / "data" / "notifications.json"
CASH_LIMIT = 5_000_000  # Luật Thuế GTGT 2024 (hiệu lực 01/07/2025): mua vào từ 5 triệu phải thanh toán không dùng tiền mặt
_lock = threading.RLock()

INPUT_KEYS = ("ct22", "ct23a", "ct24a", "khong_kt", "ct32b", "ct34a", "ct37", "ct38", "ct39a", "ct40b", "ct42")

# Căn cứ pháp lý (đã tra 11/2026 trên thuvienphapluat.vn và các trang tư vấn; luôn đối chiếu văn bản gốc trước khi nộp).
TT89_FROM = date(2026, 7, 1)  # kỳ tháng 7/2026, quý III/2026 trở đi dùng mẫu 01/GTGT theo TT 89/2026/TT-BTC
LEGAL = {
    "TT80": ("Thông tư số 80/2021/TT-BTC ngày 29/9/2021 của Bộ trưởng Bộ Tài chính",
             "https://thuvienphapluat.vn/van-ban/Thue-Phi-Le-Phi/Thong-tu-80-2021-TT-BTC-huong-dan-Luat-Quan-ly-thue-490140.aspx"),
    "TT89": ("Thông tư số 89/2026/TT-BTC ngày 30/6/2026 của Bộ trưởng Bộ Tài chính (hiệu lực 01/7/2026, thay TT 80/2021)",
             "https://thuvienphapluat.vn/hoi-dap-phap-luat/cach-dien-chi-tieu-to-khai-01gtgt-theo-thong-tu-89-day-du-moi-nhat-2026-138098760.html"),
    "LUAT_GTGT": ("Luật Thuế GTGT số 48/2024/QH15, Điều 14 (khấu trừ: mua vào từ 5 triệu đồng phải có chứng từ thanh toán "
                  "không dùng tiền mặt); Nghị định 181/2025/NĐ-CP, Điều 26",
                  "https://luatvietnam.vn/tin-van-ban-moi/tra-gop-tu-5-trieu-dong-phai-thanh-toan-khong-dung-tien-mat-moi-duoc-khau-tru-thue-gtgt-186-105579-article.html"),
    "GIAM_8": ("Nghị quyết 204/2025/QH15 và Nghị định 174/2025/NĐ-CP: giảm 2% thuế GTGT (10% → 8%) từ 01/7/2025 đến 31/12/2026, "
               "kê khai theo Mẫu số 01 Phụ lục III kèm Nghị định 174/2025",
               "https://thuvienphapluat.vn/ma-so-thue/phap-luat-thue/chi-tiet-mat-hang-giam-thue-gtgt-tu-10-xuong-8-den-het-nam-2026-892633-222484.html"),
    "HAN_NOP": ("Hạn nộp: tháng – ngày 20 tháng sau; quý – ngày cuối tháng đầu quý sau (Luật Quản lý thuế; từ 01/7/2026 "
                "đối chiếu thêm Luật Quản lý thuế 2025, Nghị định 252/2026/NĐ-CP Điều 10)",
                "https://thuvienphapluat.vn/hoi-dap-phap-luat/han-nop-to-khai-theo-thang-2026-la-khi-nao-138098496.html"),
}


def uses_tt89(kieu: str, ky: int, nam: int) -> bool:
    return period_range(kieu, ky, nam)[0] >= TT89_FROM


# ------------------------------------------------------------------ kỳ tính thuế
def period_months(kieu: str, ky: int, nam: int) -> list[tuple[int, int]]:
    return [(nam, m) for m in range(3 * ky - 2, 3 * ky + 1)] if kieu == "Q" else [(nam, ky)]


def period_key(kieu: str, ky: int, nam: int) -> str:
    return f"{nam}-Q{ky}" if kieu == "Q" else f"{nam}-{ky:02d}"


def parse_key(key: str) -> tuple[str, int, int]:
    m = re.fullmatch(r"(\d{4})-Q([1-4])", key)
    if m:
        return "Q", int(m.group(2)), int(m.group(1))
    y, mo = key.split("-")
    return "M", int(mo), int(y)


def period_label(kieu: str, ky: int, nam: int) -> str:
    return f"Quý {ky}/{nam}" if kieu == "Q" else f"Tháng {ky:02d}/{nam}"


def period_range(kieu: str, ky: int, nam: int) -> tuple[date, date]:
    ms = period_months(kieu, ky, nam)
    (y1, m1), (y2, m2) = ms[0], ms[-1]
    return date(y1, m1, 1), date(y2, m2, monthrange(y2, m2)[1])


def prev_period(kieu: str, ky: int, nam: int) -> tuple[str, int, int]:
    if kieu == "Q":
        return ("Q", 4, nam - 1) if ky == 1 else ("Q", ky - 1, nam)
    return ("M", 12, nam - 1) if ky == 1 else ("M", ky - 1, nam)


def last_completed(today: date, kieu: str) -> tuple[str, int, int]:
    """Kỳ vừa kết thúc gần nhất (ngày 01/12 -> tháng 11; ngày 01/10 -> quý 3)."""
    if kieu == "Q":
        q = (today.month - 1) // 3 + 1
        return prev_period("Q", q, today.year)
    return prev_period("M", today.month, today.year)


def deadline(kieu: str, ky: int, nam: int) -> date:
    """Hạn nộp tờ khai: ngày 20 tháng sau (tháng) / ngày cuối tháng đầu quý sau (quý)."""
    end = period_range(kieu, ky, nam)[1] + timedelta(days=1)
    if kieu == "Q":
        return date(end.year, end.month, monthrange(end.year, end.month)[1])
    return date(end.year, end.month, 20)


def reduction_basis(d: date) -> str:
    """Văn bản giảm thuế GTGT 2% (thuế suất 10% -> 8%) áp dụng cho kỳ."""
    if date(2025, 7, 1) <= d <= date(2026, 12, 31):
        return "Nghị quyết số 204/2025/QH15 và Nghị định số 174/2025/NĐ-CP"
    if date(2025, 1, 1) <= d <= date(2025, 6, 30):
        return "Nghị định số 180/2024/NĐ-CP"
    if date(2024, 7, 1) <= d <= date(2024, 12, 31):
        return "Nghị định số 72/2024/NĐ-CP"
    if date(2024, 1, 1) <= d <= date(2024, 6, 30):
        return "Nghị định số 94/2023/NĐ-CP"
    if date(2023, 7, 1) <= d <= date(2023, 12, 31):
        return "Nghị định số 44/2023/NĐ-CP"
    return "Nghị định giảm thuế GTGT áp dụng cho kỳ"


# ------------------------------------------------------------------ số / thuế suất
def _num(v: Any) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def rate_class(r: str | None) -> str:
    """'8%' -> '8' ; 'KCT' ; 'KKKNT' ; 'KHAC:5.26%' -> 'K5.26' ; '' -> '?'"""
    u = (r or "").upper().replace(" ", "")
    if u in ("KCT", "KKKNT"):
        return u
    m = re.search(r"(\d+(?:[.,]\d+)?)", u)
    if not m:
        return "?"
    pct = float(m.group(1).replace(",", "."))
    if pct < 1 and "%" not in u and pct != 0:  # 0.08 (dữ liệu chi tiết của cổng)
        pct *= 100
    pct = round(pct, 2)
    if pct in (0, 5, 8, 10):
        return str(int(pct))
    return f"K{pct:g}"


RATE_LABEL = {"KCT": "Không chịu thuế (KCT)", "KKKNT": "Không kê khai, tính nộp thuế (KKKNT)", "0": "Thuế suất 0%",
              "5": "Thuế suất 5%", "8": "Thuế suất 8% (giảm 2%)", "10": "Thuế suất 10%", "?": "Không rõ thuế suất"}
RATE_ORDER = ["KCT", "0", "5", "10", "8", "KKKNT", "?"]


def rate_label(c: str) -> str:
    return RATE_LABEL.get(c) or (f"Thuế suất khác {c[1:]}%" if c.startswith("K") else c)


def rate_pct(c: str) -> float | None:
    if c in ("KCT", "KKKNT", "?"):
        return None
    return float(c[1:]) if c.startswith("K") else float(c)


# ------------------------------------------------------------------ đọc hóa đơn
def _key(r: dict[str, Any]) -> tuple:
    return (r.get("huong"), str(r.get("nbmst") or ""), str(r.get("khmshdon") or ""), str(r.get("khhdon") or ""),
            str(r.get("shdon") or "").lstrip("0"))


def load_rows(root: Path, kieu: str, ky: int, nam: int) -> tuple[list[dict[str, Any]], list[str], list[str]]:
    """(hóa đơn của kỳ – đã bỏ trùng, tháng chưa có dữ liệu, tháng dữ liệu chưa trọn tháng)."""
    rows: dict[tuple, dict[str, Any]] = {}
    missing, partial = [], []
    for y, m in period_months(kieu, ky, nam):
        f = root / f"{y:04d}-{m:02d}" / f".dulieu_{y:04d}-{m:02d}.json"
        if not f.exists():
            missing.append(f"{m:02d}/{y}")
            continue
        d = json.loads(f.read_text(encoding="utf-8"))
        if str(d.get("saved", "")) <= date(y, m, monthrange(y, m)[1]).isoformat():
            partial.append(f"{m:02d}/{y} (tải ngày {d.get('saved')})")
        for r in d.get("rows", []):
            rows.setdefault(_key(r), r)
    return list(rows.values()), missing, partial


def _detail(rec: dict[str, Any], mst: str, info: dict[str, Any]) -> tuple[list[dict[str, Any]], dict[str, list[float]]]:
    """Dòng hàng (đã chuẩn hóa) + tổng theo thuế suất của 1 hóa đơn, đọc từ file XML/JSON đã tải."""
    path = Path(str(rec.get("file_xml") or ""))
    items: list[dict[str, Any]] = []
    rates: dict[str, list[float]] = {}
    if not path.is_file():
        return items, rates
    try:
        if path.suffix.lower() == ".json":  # dữ liệu chi tiết của cổng (khi không tải được XML)
            d = json.loads(path.read_text(encoding="utf-8"))
            for h in d.get("hdhhdvu") or []:
                tc = h.get("tchat")
                if tc == 4:
                    continue
                tt = _num(h.get("thtien")) * (-1 if tc == 3 else 1)
                items.append({"ten": h.get("ten") or "", "dvt": h.get("dvtinh") or "", "sl": h.get("sluong"),
                              "dg": h.get("dgia"), "tt": tt, "ts": rate_class(h.get("ltsuat") or str(h.get("tsuat"))),
                              "thue": _num(h.get("tthue")) if h.get("tthue") is not None else None})
            for e in d.get("thttltsuat") or []:
                c = rate_class(e.get("tsuat") or e.get("ltsuat"))
                a = rates.setdefault(c, [0.0, 0.0])
                a[0] += _num(e.get("thtien"))
                a[1] += _num(e.get("tthue"))
            if mst and not info.get("ten"):
                if str(d.get("nmmst")) == mst:
                    info.update(ten=d.get("nmten"), dchi=d.get("nmdchi"))
                elif str(d.get("nbmst")) == mst:
                    info.update(ten=d.get("nbten"), dchi=d.get("nbdchi"))
        else:
            from .gdt.xml_parser import parse_invoice_xml
            p = parse_invoice_xml(path.read_bytes())
            for h in p.get("hang_hoa") or []:
                tc = str(h.get("tinh_chat") or "1")
                if tc == "4":
                    continue
                tt = _num(h.get("thanh_tien")) * (-1 if tc == "3" else 1)
                items.append({"ten": h.get("ten") or "", "dvt": h.get("dvt") or "", "sl": h.get("so_luong"),
                              "dg": h.get("don_gia"), "tt": tt, "ts": rate_class(h.get("thue_suat")), "thue": None})
            for r, (v, t) in (p.get("theo_thue_suat") or {}).items():
                a = rates.setdefault(rate_class(r), [0.0, 0.0])
                a[0] += v
                a[1] += t
            if mst and not info.get("ten"):
                for side in ("nguoi_mua", "nguoi_ban"):
                    party = p.get(side) or {}
                    if (party.get("mst") or "") == mst:
                        info.update(ten=party.get("ten"), dchi=party.get("dia_chi"))
    except Exception:  # noqa: BLE001 - file hỏng -> dùng tổng tiền của cổng
        return [], {}
    if not rates and items:  # không có bảng tổng hợp theo thuế suất -> cộng từ dòng hàng
        for it in items:
            a = rates.setdefault(it["ts"], [0.0, 0.0])
            a[0] += it["tt"]
            pct = rate_pct(it["ts"])
            a[1] += it["thue"] if it["thue"] is not None else (round(it["tt"] * pct / 100) if pct else 0.0)
    return items, rates


def _spread_tax(items: list[dict[str, Any]], rates: dict[str, list[float]]) -> None:
    """Tiền thuế từng dòng hàng = thành tiền x thuế suất; chênh lệch làm tròn dồn vào dòng lớn nhất
    để tổng thuế từng dòng khớp đúng tổng thuế theo thuế suất trên hóa đơn."""
    for c, (_, tax) in rates.items():
        group = [it for it in items if it["ts"] == c]
        pct = rate_pct(c)
        for it in group:
            if it["thue"] is None:
                it["thue"] = round(it["tt"] * pct / 100) if pct is not None else 0.0
        if group and pct:
            diff = round(tax - sum(it["thue"] for it in group))
            if diff and abs(diff) <= max(10, 0.02 * abs(tax)):
                max(group, key=lambda it: abs(it["tt"]))["thue"] += diff
    for it in items:
        if it["thue"] is None:
            pct = rate_pct(it["ts"])
            it["thue"] = round(it["tt"] * pct / 100) if pct is not None else 0.0


# ------------------------------------------------------------------ tính tờ khai
CT_LABELS = {
    "21": "Không phát sinh hoạt động mua, bán trong kỳ",
    "22": "Thuế GTGT còn được khấu trừ kỳ trước chuyển sang",
    "23": "Giá trị HHDV mua vào", "24": "Thuế GTGT của HHDV mua vào",
    "23a": "Giá trị HHDV nhập khẩu", "24a": "Thuế GTGT HHDV nhập khẩu",
    "25": "Thuế GTGT của HHDV mua vào được khấu trừ kỳ này",
    "26": "HHDV bán ra không chịu thuế GTGT",
    "27": "HHDV bán ra chịu thuế GTGT – giá trị", "28": "HHDV bán ra chịu thuế GTGT – thuế",
    "29": "Bán ra thuế suất 0%", "30": "Bán ra thuế suất 5% – giá trị", "31": "Bán ra thuế suất 5% – thuế",
    "32": "Bán ra thuế suất 10% (gồm 8% được giảm) – giá trị", "33": "Bán ra thuế suất 10% (gồm 8%) – thuế",
    "32a": "HHDV bán ra không tính thuế",
    "32b": "HHDV bán ra không tính vào giá tính thuế GTGT (TT89)",
    "34a": "HHDV bán ra không thuộc phạm vi điều chỉnh của pháp luật thuế GTGT (TT89)",
    "34": "Tổng doanh thu HHDV bán ra", "35": "Tổng thuế GTGT HHDV bán ra",
    "36": "Thuế GTGT phát sinh trong kỳ ([35]-[25])",
    "37": "Điều chỉnh giảm thuế GTGT còn được khấu trừ các kỳ trước",
    "38": "Điều chỉnh tăng thuế GTGT còn được khấu trừ các kỳ trước",
    "39a": "Thuế GTGT nhận bàn giao được khấu trừ trong kỳ",
    "40a": "Thuế GTGT phải nộp của hoạt động SXKD trong kỳ",
    "40b": "Thuế GTGT mua vào của dự án đầu tư bù trừ",
    "40": "Thuế GTGT còn phải nộp trong kỳ",
    "41": "Thuế GTGT chưa khấu trừ hết kỳ này",
    "42": "Thuế GTGT đề nghị hoàn",
    "43": "Thuế GTGT còn được khấu trừ chuyển kỳ sau",
}


def _orig_ref(r: dict[str, Any]) -> dict[str, Any] | None:
    """Hóa đơn gốc mà HĐ này thay thế / điều chỉnh: {kh, so, ngay, loai(1 thay thế, 2 điều chỉnh)}.
    Lấy từ danh sách cổng (khhdgoc/shdgoc), nếu không có thì đọc XML (TTHDLQuan) hoặc JSON chi tiết."""
    kh, so, ngay, loai = r.get("khhdgoc"), r.get("shdgoc"), r.get("tdlhdgoc"), None
    f = str(r.get("file_xml") or "")
    if not (kh and so) and f:
        try:
            p = Path(f)
            if p.suffix.lower() == ".json" and p.exists():
                d = json.loads(p.read_text(encoding="utf-8"))
                kh, so, ngay = d.get("khhdgoc"), d.get("shdgoc"), d.get("tdlhdgoc")
            elif p.suffix.lower() == ".xml" and p.exists():
                t = p.read_text(encoding="utf-8", errors="ignore")
                m = re.search(r"<TTHDLQuan>(.*?)</TTHDLQuan>", t, re.S)
                if m:
                    g = lambda tag: (re.search(rf"<{tag}>([^<]*)</{tag}>", m.group(1)) or [None, None])[1]  # noqa: E731
                    kh, so, ngay, loai = g("KHHDCLQuan"), g("SHDCLQuan"), g("NLHDCLQuan"), g("TCHDon")
        except (OSError, ValueError):
            pass
    if not (kh and str(so or "").strip()):
        return None
    try:
        so_i = int(str(so).strip())
    except ValueError:
        return None
    return {"kh": str(kh).strip().upper(), "so": so_i, "ngay": str(ngay or "")[:10], "loai": str(loai or "")}


def _inv_name(r: dict[str, Any]) -> str:
    return f"HĐ {r.get('khhdon') or ''}-{r.get('shdon') or ''} ({'mua vào' if r.get('huong') == 'purchase' else 'bán ra'}, " \
           f"{str(r.get('ngay_lap') or '')[8:10]}/{str(r.get('ngay_lap') or '')[5:7]})"


def _sum_rates(rates: dict[str, list[float]]) -> tuple[float, float]:
    return sum(v[0] for v in rates.values()), sum(v[1] for v in rates.values())


def _pair_invoices(prepared: list[dict[str, Any]]) -> list[str]:
    """Ghép hóa đơn thay thế / điều chỉnh với hóa đơn gốc cùng kỳ:
      * HĐ bị thay thế: loại, chỉ kê khai HĐ thay thế.
      * HĐ điều chỉnh: KHÔNG kê khai riêng – số tăng/giảm cộng/trừ vào HĐ bị điều chỉnh;
        điều chỉnh giảm hết (tổng về 0) -> loại cả 2.
      * HĐ gốc ở kỳ khác: HĐ thay thế kê khai đủ, HĐ điều chỉnh kê khai số tăng/giảm của chính nó.
    Trả về các dòng giải thích để hiện cho người dùng."""
    def key(d: str, nbmst: Any, kh: Any, so: Any) -> tuple:
        try:
            so = int(str(so).strip())
        except ValueError:
            pass
        return d, str(nbmst or "").replace("-", ""), str(kh or "").strip().upper(), so

    idx = {key(p["dir"], p["r"].get("nbmst"), p["r"].get("khhdon"), p["r"].get("shdon")): p for p in prepared}
    notes: list[str] = []
    replaced_by: dict[int, dict[str, Any]] = {}
    # HĐ thay thế (2) / điều chỉnh (3): tìm HĐ gốc
    for p in prepared:
        ref = p["ref"]
        kind = 2 if p["code"] == 2 or (ref and ref["loai"] == "1") else 3 if p["code"] == 3 or (ref and ref["loai"] == "2") else 0
        if not kind:
            continue
        r = p["r"]
        orig = idx.get(key(p["dir"], r.get("nbmst"), ref["kh"], ref["so"])) if ref else None
        if not orig and not ref:
            # dữ liệu cũ chưa có số HĐ gốc: ghép với HĐ "đã bị thay thế"/"đã bị điều chỉnh" DUY NHẤT trong kỳ
            # cùng người bán, cùng ký hiệu, lập trước HĐ này
            want = 4 if kind == 2 else 5
            cands = [q for q in prepared if q is not p and q["code"] == want and not q["drop"] and q["dir"] == p["dir"]
                     and str(q["r"].get("nbmst") or "") == str(r.get("nbmst") or "")
                     and (q["r"].get("khhdon") or "") == (r.get("khhdon") or "")
                     and (q["r"].get("ngay_lap") or "") <= (r.get("ngay_lap") or "")]
            if len(cands) == 1:
                orig = cands[0]
                ref = {"kh": str(orig["r"].get("khhdon") or "").upper(), "so": orig["r"].get("shdon"),
                       "ngay": str(orig["r"].get("ngay_lap") or ""), "loai": ""}
                p["note"].append("ghép HĐ gốc theo cùng người bán & ký hiệu (dữ liệu chưa có số HĐ gốc)")
        if ref and not p.get("link"):
            p["link"] = f"{ref['kh']}-{ref['so']} (kỳ khác)"
        goc_txt = f"HĐ gốc {ref['kh']}-{ref['so']}" + (f" ngày {ref['ngay'][8:10]}/{ref['ngay'][5:7]}/{ref['ngay'][:4]}"
                                                        if len(ref["ngay"]) >= 10 else "") if ref else "HĐ gốc (không đọc được số)"
        if kind == 2:
            if orig and orig is not p:
                orig["drop"] = f"Bị thay thế bởi HĐ {r.get('khhdon')}-{r.get('shdon')} – loại, không kê khai"
                replaced_by[id(orig)] = p
                orig["role"], orig["link"] = "bi_thay_the", f"{r.get('khhdon')}-{r.get('shdon')}"
                p["link"] = f"{ref['kh']}-{ref['so']}"
                p["note"].append(f"Thay thế HĐ {ref['kh']}-{ref['so']}")
                notes.append(f"{_inv_name(orig['r'])} bị thay thế bởi {_inv_name(r)}: loại HĐ {ref['kh']}-{ref['so']}, "
                             f"chỉ kê khai HĐ {r.get('khhdon')}-{r.get('shdon')}.")
            elif ref:
                p["note"].append(f"Thay thế {goc_txt} (không thuộc kỳ này)")
                notes.append(f"{_inv_name(r)} thay thế {goc_txt} không thuộc kỳ này: kê khai HĐ thay thế; nếu HĐ gốc đã "
                             "kê khai ở kỳ trước thì kiểm tra lại kỳ đó.")
            else:
                p["note"].append("Thay thế HĐ chưa xác định được số (dữ liệu chưa có số HĐ gốc)")
                notes.append(f"{_inv_name(r)} là HĐ thay thế nhưng dữ liệu chưa có số HĐ gốc: đã kê khai HĐ thay thế; "
                             "các HĐ 'đã bị thay thế' trong kỳ đã loại. Tải XML/chi tiết hóa đơn để ghép chính xác.")
        else:
            val, tax = _sum_rates(p["rates"])
            sign = "tăng" if val + tax > 0 else "giảm"
            amt = f"{fmt(abs(val))} đ tiền hàng, {fmt(abs(tax))} đ thuế"
            if orig and orig is not p:
                for c, (v, t) in p["rates"].items():
                    cur = orig["rates"].setdefault(c, [0.0, 0.0])
                    cur[0] += v
                    cur[1] += t
                orig["its"] = orig["its"] + p["its"]
                orig["adj"].append(p)
                orig["role"] = "bi_dieu_chinh"
                orig["link"] = ", ".join(x for x in (orig.get("link"), f"{r.get('khhdon')}-{r.get('shdon')}") if x)
                p["link"] = f"{ref['kh']}-{ref['so']}"
                p["drop"] = (f"HĐ điều chỉnh {sign} cho HĐ {ref['kh']}-{ref['so']} – đã cộng/trừ vào HĐ bị điều chỉnh, "
                             "không kê khai riêng")
                orig["note"].append(f"Đã {'cộng' if sign == 'tăng' else 'trừ'} HĐ điều chỉnh {r.get('khhdon')}-{r.get('shdon')} "
                                    f"({amt})")
                p["pair_note"] = len(notes)
                notes.append(f"{_inv_name(orig['r'])} bị điều chỉnh {sign} bởi {_inv_name(r)} ({amt}): "
                             f"{'cộng vào' if sign == 'tăng' else 'trừ vào'} HĐ {ref['kh']}-{ref['so']}, "
                             f"không kê khai HĐ {r.get('khhdon')}-{r.get('shdon')}.")
            elif not ref:
                p["note"].append(f"Điều chỉnh {sign} HĐ chưa xác định được số – kê khai số {sign}")
                notes.append(f"{_inv_name(r)} là HĐ điều chỉnh {sign} ({amt}) nhưng dữ liệu chưa có số HĐ gốc: tạm kê khai "
                             f"số điều chỉnh {sign}. Tải XML/chi tiết hóa đơn để ghép đúng HĐ bị điều chỉnh.")
            else:
                p["note"].append(f"Điều chỉnh {sign} {goc_txt} (không thuộc kỳ này) – kê khai số {sign}")
                notes.append(f"{_inv_name(r)} điều chỉnh {sign} {goc_txt} không thuộc kỳ này ({amt}): kê khai số "
                             f"điều chỉnh {sign} ở kỳ này.")
    # điều chỉnh giảm hết -> triệt tiêu, loại cả HĐ gốc và HĐ điều chỉnh
    for p in prepared:
        if p["adj"] and not p["drop"]:
            val, tax = _sum_rates(p["rates"])
            if abs(val) < 1 and abs(tax) < 1:
                adj = ", ".join(f"{a['r'].get('khhdon')}-{a['r'].get('shdon')}" for a in p["adj"])
                p["drop"] = f"Bị điều chỉnh giảm hết bởi HĐ {adj} – 2 HĐ triệt tiêu, không kê khai"
                for a in p["adj"]:
                    a["drop"] = f"Điều chỉnh giảm hết HĐ {p['r'].get('khhdon')}-{p['r'].get('shdon')} – triệt tiêu, không kê khai"
                    if a.get("pair_note") is not None:
                        notes[a["pair_note"]] = ""   # thay bằng dòng "triệt tiêu" bên dưới
                notes.append(f"{_inv_name(p['r'])} và HĐ điều chỉnh {adj} triệt tiêu về 0: loại cả 2, không kê khai.")
    # HĐ bị thay thế mà HĐ thay thế không có trong kỳ: vẫn loại (cổng đã ghi nhận bị thay thế)
    for p in prepared:
        if p["code"] == 4 and not p["drop"]:
            p["drop"] = "Đã bị thay thế – loại, không kê khai"
            notes.append(f"{_inv_name(p['r'])} đã bị thay thế (HĐ thay thế ở kỳ khác hoặc chưa ghép được số): loại, "
                         "không kê khai.")
        elif p["code"] == 5 and not p["adj"] and not p["drop"]:
            p["note"].append("Đã bị điều chỉnh (HĐ điều chỉnh ở kỳ khác)")
            notes.append(f"{_inv_name(p['r'])} đã bị điều chỉnh bởi HĐ ở kỳ khác: kê khai theo số gốc; số điều chỉnh "
                         "kê khai ở kỳ có HĐ điều chỉnh.")
    return [n for n in notes if n]


def compute(mst: str, kieu: str, ky: int, nam: int, inputs: dict[str, Any] | None = None) -> dict[str, Any]:
    inputs = {k: round(_num((inputs or {}).get(k))) for k in INPUT_KEYS}
    root = Path(settings.output_dir) / mst
    rows, missing, partial = load_rows(root, kieu, ky, nam)
    # ---- kỳ chưa kết thúc (vd dự báo quý 4 vào ngày 19/11): tính bằng dữ liệu đến hôm nay, ghi rõ "tính đến ngày"
    p_from, p_to = period_range(kieu, ky, nam)
    today = date.today()
    chua_het_ky = today < p_to
    future = [f"{m:02d}/{y}" for y, m in period_months(kieu, ky, nam) if date(y, m, 1) > today]  # tháng chưa tới
    current = f"{today:%m/%Y}" if p_from <= today <= p_to else ""
    missing = [x for x in missing if x not in future]
    partial = [x for x in partial if not x.startswith(current or "-")]   # tháng đang diễn ra: tải giữa tháng là bình thường
    den_ngay, doc_ngay = (min(today, p_to) if chua_het_ky else p_to), None
    if chua_het_ky:
        # ngày đọc dữ liệu của tháng gần nhất có dữ liệu (đọc hôm 15/11 thì số liệu chỉ đến 15/11)
        for y, m in reversed(period_months(kieu, ky, nam)):
            f = root / f"{y:04d}-{m:02d}" / f".dulieu_{y:04d}-{m:02d}.json"
            if f.exists():
                try:
                    sv = date.fromisoformat(str(json.loads(f.read_text(encoding="utf-8")).get("saved", ""))[:10])
                    month_end = date(y, m, monthrange(y, m)[1])
                    if sv < min(month_end, today):
                        doc_ngay = sv
                        den_ngay = max(p_from, sv - timedelta(days=0))
                except (OSError, ValueError):
                    pass
                break
    info: dict[str, Any] = {}
    ct = {k: 0.0 for k in CT_LABELS if k != "21"}
    bk = {"purchase": [], "sold": []}       # dòng bảng kê: 1 hóa đơn x 1 thuế suất
    items = {"purchase": [], "sold": []}    # dòng chi tiết từng mặt hàng
    excluded, warn_rows = [], []
    stat = {"mua_vao": 0, "ban_ra": 0, "loai_tru": 0, "khong_khau_tru": 0, "khong_xml": 0}
    no_deduct_val = no_deduct_tax = 0.0
    kct_val = 0.0   # mua vào KCT / KKKNT: giá trị đưa vào [23]

    from .bank_filter import BANK_NOTE, is_bank_purchase
    from .reports import tthai_code
    # ---- bước 1: đọc từng hóa đơn (thuế suất, dòng hàng, hóa đơn gốc nếu là HĐ thay thế / điều chỉnh)
    prepared: list[dict[str, Any]] = []
    for r in sorted(rows, key=lambda x: (x.get("ngay_lap") or "", x.get("khhdon") or "", str(x.get("shdon") or ""))):
        code = tthai_code(r)
        direction = "purchase" if r.get("huong") == "purchase" else "sold"
        if code == 6:
            excluded.append((r, "Hóa đơn đã bị hủy – không kê khai"))
            continue
        if is_bank_purchase(r):
            excluded.append((r, BANK_NOTE))
            continue
        its, rates = _detail(r, mst, info)
        if not rates:
            rates = {rate_class(k): list(v) for k, v in (r.get("_rates") or {}).items()} if r.get("_rates") else {}
        no_detail = not rates
        if no_detail:
            tax = _num(r.get("tgtthue"))
            val = _num(r.get("tgtcthue"))
            guess = "?"
            if val and tax:
                p = round(tax / val * 100)
                guess = str(p) if p in (5, 8, 10) else "?"
            rates = {guess if tax else ("?" if direction == "purchase" else "KCT"): [val, tax]}
        _spread_tax(its, rates)
        prepared.append({"r": r, "code": code, "dir": direction, "its": its, "rates": rates, "no_detail": no_detail,
                         "ref": _orig_ref(r), "drop": None, "adj": [], "note": []})
    pair_notes = _pair_invoices(prepared)   # ghép HĐ thay thế / điều chỉnh với HĐ gốc trong kỳ
    for p in prepared:
        if p["drop"]:
            excluded.append((p["r"], p["drop"]))
    stat["loai_tru"] = len(excluded)

    # ---- bước 2: kê khai các hóa đơn còn lại
    for p in prepared:
        if p["drop"]:
            continue
        r, code, direction, its, rates = p["r"], p["code"], p["dir"], p["its"], p["rates"]
        if p["no_detail"]:
            stat["khong_xml"] += 1
            warn_rows.append((r, "Không có XML/chi tiết – dùng tổng tiền trên cổng, thuế suất suy ra"))
        partner_mst = r.get("nbmst") if direction == "purchase" else r.get("nmmst")
        partner = r.get("nbten") if direction == "purchase" else r.get("nmten")
        mau = str(r.get("khmshdon") or "")
        deductible = direction == "purchase" and mau == "1"
        if direction == "purchase":
            stat["mua_vao"] += 1
            if not deductible:
                stat["khong_khau_tru"] += 1
            if _num(r.get("tgtttbso")) >= CASH_LIMIT:
                warn_rows.append((r, f"Mua vào từ {CASH_LIMIT:,.0f}đ: cần chứng từ thanh toán không dùng tiền mặt "
                                     "để được khấu trừ".replace(",", ".")))
        else:
            stat["ban_ra"] += 1
        for c, (val, tax) in rates.items():
            line = {"ngay": r.get("ngay_lap"), "mau": mau, "kh": r.get("khhdon"), "so": r.get("shdon"),
                    "mst": partner_mst, "ten": partner, "ts": c, "gt": round(val), "thue": round(tax),
                    "tthai": r.get("tthai"), "loai": r.get("loai"), "ghi_chu": "; ".join(p["note"])}
            if direction == "purchase":
                if deductible and rate_pct(c) is not None or (deductible and c == "?" and tax):
                    line["kt"] = True
                    ct["23"] += val
                    ct["24"] += tax
                elif c in ("KCT", "KKKNT"):
                    # HHDV không chịu thuế / không kê khai tính nộp thuế: giá trị chưa thuế vẫn đưa vào [23], không có thuế
                    line["kt"] = False
                    line["ghi_chu"] = "; ".join(x for x in (line["ghi_chu"], f"{c}: giá trị đưa vào [23], không có thuế "
                                                           "GTGT khấu trừ") if x)
                    ct["23"] += val
                    kct_val += val
                else:
                    # HĐ bán hàng (mẫu 2) / HĐ không phải HĐ GTGT: giá trị vẫn đưa vào [23], thuế không khấu trừ
                    line["kt"] = False
                    line["ghi_chu"] = "; ".join(x for x in (line["ghi_chu"], (
                        "Hóa đơn bán hàng: giá trị đưa vào [23], không có thuế GTGT khấu trừ" if mau == "2"
                        else "Không phải HĐ GTGT: giá trị đưa vào [23], thuế không được khấu trừ")) if x)
                    ct["23"] += val
                    no_deduct_val += val
                    no_deduct_tax += tax
            else:
                if c == "KCT":
                    ct["26"] += val
                elif c == "KKKNT":
                    ct["32a"] += val
                elif c == "0":
                    ct["29"] += val
                elif c == "5":
                    ct["30"] += val
                    ct["31"] += tax
                elif c == "?":
                    ct["32"] += val
                    ct["33"] += tax
                    line["ghi_chu"] = "Không rõ thuế suất – tạm cộng vào [32]/[33]"
                else:  # 8%, 10% và mức khác
                    ct["32"] += val
                    ct["33"] += tax
                    if c.startswith("K"):
                        line["ghi_chu"] = f"Thuế suất {c[1:]}% – tạm cộng vào [32]/[33], cần xem lại"
            bk[direction].append(line)
        for it in its:
            items[direction].append({"ngay": r.get("ngay_lap"), "kh": r.get("khhdon"), "so": r.get("shdon"),
                                     "mst": partner_mst, "ten_dt": partner, **it, "tong": it["tt"] + (it["thue"] or 0)})

    # ----- danh sách riêng: HĐ ngân hàng, HĐ bị thay thế / thay thế / bị điều chỉnh / điều chỉnh (sheet trong Excel)
    def _lq(r: dict[str, Any], lien_quan: str, xu_ly: str, sau: tuple[float, float] | None = None) -> dict[str, Any]:
        mua = r.get("huong") == "purchase"
        d = {"huong": "Mua vào" if mua else "Bán ra", "loai": r.get("loai"), "ngay": r.get("ngay_lap"),
             "mau": r.get("khmshdon"), "kh": r.get("khhdon"), "so": r.get("shdon"),
             "mst": r.get("nbmst") if mua else r.get("nmmst"), "ten": r.get("nbten") if mua else r.get("nmten"),
             "gt": round(_num(r.get("tgtcthue"))), "thue": round(_num(r.get("tgtthue"))), "tong": round(_num(r.get("tgtttbso"))),
             "tthai": r.get("tthai"), "lien_quan": lien_quan, "xu_ly": xu_ly}
        if sau is not None:
            d["gt_sau"], d["thue_sau"] = round(sau[0]), round(sau[1])
        return d

    lien_quan: dict[str, list[dict[str, Any]]] = {"ngan_hang": [], "bi_thay_the": [], "thay_the": [],
                                                  "bi_dieu_chinh": [], "dieu_chinh": []}
    for r, why in excluded:
        if why == BANK_NOTE:
            lien_quan["ngan_hang"].append(_lq(r, "", why))
    role_of = {2: "thay_the", 3: "dieu_chinh", 4: "bi_thay_the", 5: "bi_dieu_chinh"}
    for p in prepared:
        role = p.get("role") or role_of.get(p["code"])
        if not role:
            continue
        xu_ly = p["drop"] or ("Kê khai" + (": " + "; ".join(p["note"]) if p["note"] else ""))
        sau = _sum_rates(p["rates"]) if role == "bi_dieu_chinh" and not p["drop"] else None
        lien_quan[role].append(_lq(p["r"], p.get("link") or "", xu_ly, sau))

    # ----- chỉ tiêu nhập tay + công thức của tờ khai
    ct["23a"], ct["24a"] = inputs["ct23a"], inputs["ct24a"]
    ct["23"] += inputs["ct23a"]
    ct["24"] += inputs["ct24a"]
    ct["25"] = max(0.0, ct["24"] - inputs["khong_kt"])
    ct["22"], ct["37"], ct["38"], ct["39a"] = inputs["ct22"], inputs["ct37"], inputs["ct38"], inputs["ct39a"]
    ct = {k: round(v) for k, v in ct.items()}
    tt89 = uses_tt89(kieu, ky, nam)
    ct["32b"] = inputs["ct32b"] if tt89 else 0  # TT89: HHDV bán ra không tính vào giá tính thuế GTGT
    ct["34a"] = inputs["ct34a"] if tt89 else 0  # TT89: HHDV bán ra không thuộc phạm vi điều chỉnh của Luật GTGT
    ct["27"] = ct["29"] + ct["30"] + ct["32"] + ct["32a"] - ct["32b"]
    ct["28"] = ct["31"] + ct["33"]
    ct["34"] = ct["26"] + ct["27"] + ct["34a"]
    ct["35"] = ct["28"]
    ct["36"] = ct["35"] - ct["25"]
    x = ct["36"] - ct["22"] + ct["37"] - ct["38"] - ct["39a"]
    ct["40a"] = x if x >= 0 else 0
    ct["40b"] = min(inputs["ct40b"], ct["40a"])
    ct["40"] = ct["40a"] - ct["40b"]
    ct["41"] = -x if x < 0 else 0
    ct["42"] = min(inputs["ct42"], ct["41"])
    ct["43"] = ct["41"] - ct["42"]
    ct21 = not rows and not any(inputs.values())

    # ----- phụ lục giảm thuế 8%
    d1, d2 = period_range(kieu, ky, nam)
    pl8 = _pl8(items["sold"], bk["sold"])
    pl8_mua = _pl8(items["purchase"], [b for b in bk["purchase"] if b.get("kt")])

    warnings = []
    if missing:
        warnings.append(f"Chưa có dữ liệu hóa đơn tháng {', '.join(missing)} – hãy tải trọn tháng ở mục Tải hóa đơn.")
    if partial:
        warnings.append(f"Dữ liệu tải trước khi hết tháng: {', '.join(partial)} – có thể còn thiếu hóa đơn cuối tháng.")
    if rows:
        from .quick_read import coverage
        warnings[:0] = coverage(root, [(y, m) for y, m in period_months(kieu, ky, nam)
                                       if f"{m:02d}/{y}" not in missing and f"{m:02d}/{y}" not in future], rows)
    if chua_het_ky:
        warnings.insert(0, f"{period_label(kieu, ky, nam)} CHƯA KẾT THÚC: số liệu chỉ tính đến ngày {den_ngay:%d/%m/%Y}"
                        + (f" (tháng {', '.join(future)} chưa tới)" if future else "")
                        + f" – đây là số TẠM TÍNH, sẽ thay đổi khi phát sinh thêm hóa đơn đến hết ngày {p_to:%d/%m/%Y}.")
        if doc_ngay:
            warnings.insert(1, f"Dữ liệu tháng {current or den_ngay.strftime('%m/%Y')} đọc từ ngày {doc_ngay:%d/%m/%Y} – bấm "
                               "⚡ Đọc nhanh từ cổng để cập nhật đến hôm nay.")
    if stat["loai_tru"]:
        warnings.append(f"Đã loại {stat['loai_tru']} hóa đơn không kê khai: "
                        + "; ".join(f"{_inv_name(r)} – {why}" for r, why in excluded[:30])
                        + (f"; … và {len(excluded) - 30} HĐ khác (xem sheet Lưu ý)" if len(excluded) > 30 else "") + ".")
    warnings.extend(pair_notes)
    if stat["khong_khau_tru"]:
        warnings.append(f"{stat['khong_khau_tru']} hóa đơn mua vào không phải hóa đơn GTGT (mẫu 2…) – không khấu trừ.")
    if no_deduct_val:
        warnings.append(f"HHDV mua vào theo hóa đơn bán hàng (mẫu 2) / không phải HĐ GTGT: giá trị {fmt(no_deduct_val)} đ "
                        "đã đưa vào chỉ tiêu [23], không có thuế đầu vào [24].")
    if kct_val:
        warnings.append(f"HHDV mua vào không chịu thuế (KCT) / không kê khai tính nộp (KKKNT): giá trị {fmt(kct_val)} đ "
                        "đã đưa vào chỉ tiêu [23], không có thuế đầu vào [24].")
    cash = [w for w in warn_rows if "tiền mặt" in w[1]]
    if cash:
        warnings.append(f"{len(cash)} hóa đơn mua vào từ 5 triệu đồng: cần chứng từ thanh toán không dùng tiền mặt "
                        "mới được khấu trừ (Luật Thuế GTGT 2024). Nếu chưa thanh toán qua ngân hàng, nhập số thuế không "
                        "được khấu trừ vào ô tương ứng.")
    if stat["khong_xml"]:
        nx = [r for r, why in warn_rows if why.startswith("Không có XML")]
        warnings.append(f"{stat['khong_xml']} hóa đơn chưa có XML/chi tiết – thuế suất suy ra từ tổng tiền: "
                        + ", ".join(f"{r.get('khhdon')}-{r.get('shdon')}" for r in nx[:25])
                        + (f" … (+{len(nx) - 25})" if len(nx) > 25 else "") + ".")
    if not inputs["ct22"]:
        warnings.append("Chỉ tiêu [22] (thuế còn được khấu trừ kỳ trước chuyển sang) đang là 0 – nhập nếu kỳ trước còn dư.")
    if pl8["rows"]:
        warnings.append(f"Có bán ra 8%: kèm Phụ lục giảm thuế GTGT theo {reduction_basis(d2)} – "
                        f"số thuế được giảm {fmt(pl8['giam'])} đ.")

    mau = "TT89" if tt89 else "TT80"
    legal = [LEGAL[mau], LEGAL["LUAT_GTGT"], LEGAL["HAN_NOP"]] + ([LEGAL["GIAM_8"]] if pl8["rows"] or pl8_mua["rows"] else [])
    return {
        "mau": mau, "mau_text": ("Ban hành kèm theo Thông tư số 89/2026/TT-BTC ngày 30 tháng 6 năm 2026 của Bộ trưởng Bộ Tài chính"
                                 if tt89 else "Ban hành kèm theo Thông tư số 80/2021/TT-BTC ngày 29 tháng 9 năm 2021 "
                                              "của Bộ trưởng Bộ Tài chính"),
        "legal": [{"text": t, "url": u} for t, u in legal],
        "mst": mst, "kieu": kieu, "ky": ky, "nam": nam, "key": period_key(kieu, ky, nam),
        "label": period_label(kieu, ky, nam), "from": d1.isoformat(), "to": d2.isoformat(),
        "deadline": deadline(kieu, ky, nam).isoformat(), "basis": reduction_basis(d2),
        "company": {"ten": info.get("ten") or "", "dchi": info.get("dchi") or ""},
        "inputs": inputs, "ct": ct, "ct21": ct21, "stat": stat, "missing": missing, "partial": partial,
        "warnings": warnings, "bk": bk, "items": items, "pl8": pl8, "pl8_mua": pl8_mua,
        "warn_rows": [{"ngay": r.get("ngay_lap"), "kh": r.get("khhdon"), "so": r.get("shdon"),
                       "mst": r.get("nbmst") if r.get("huong") == "purchase" else r.get("nmmst"),
                       "ten": r.get("nbten") if r.get("huong") == "purchase" else r.get("nmten"),
                       "huong": "Mua vào" if r.get("huong") == "purchase" else "Bán ra",
                       "tong": r.get("tgtttbso"), "why": why} for r, why in warn_rows],
        "excluded": [{"ngay": r.get("ngay_lap"), "kh": r.get("khhdon"), "so": r.get("shdon"), "tthai": r.get("tthai"),
                      "huong": "Mua vào" if r.get("huong") == "purchase" else "Bán ra", "tong": r.get("tgtttbso"),
                      "loai": r.get("loai"), "why": why}
                     for r, why in excluded],
        "rows": len(rows),
        "lien_quan": lien_quan,
        "chua_het_ky": chua_het_ky, "den_ngay": den_ngay.isoformat(), "future": future,
        "label_den": (f"{period_label(kieu, ky, nam)} (tính đến ngày {den_ngay:%d/%m/%Y})" if chua_het_ky
                      else period_label(kieu, ky, nam)),
    }


def _pl8(items: list[dict[str, Any]], lines: list[dict[str, Any]]) -> dict[str, Any]:
    """Gộp HHDV thuế suất 8% theo tên hàng. Thiếu chi tiết dòng hàng thì lấy theo hóa đơn."""
    by: dict[str, list] = {}
    for it in items:
        if it["ts"] != "8":
            continue
        name = re.sub(r"\s+", " ", (it.get("ten") or "").strip()) or "(không tên)"
        a = by.setdefault(name.upper(), [name, 0.0])
        a[1] += it["tt"]
    total_items = sum(v for _, v in by.values())
    total_inv = sum(b["gt"] for b in lines if b["ts"] == "8")
    rows = [{"ten": n, "gt": round(v)} for n, v in sorted(by.values(), key=lambda x: -x[1])]
    diff = round(total_inv - total_items)
    if abs(diff) >= 1:
        rows.append({"ten": "Chênh lệch theo tổng hóa đơn (chiết khấu, làm tròn, hóa đơn không có chi tiết)", "gt": diff})
    for x in rows:
        x["giam"] = round(x["gt"] * 0.02)
    return {"rows": [x for x in rows if x["gt"]], "gt": round(total_inv), "giam": round(total_inv * 0.02)}


def fmt(v: Any) -> str:
    try:
        n = round(float(v))
    except (TypeError, ValueError):
        return ""
    s = f"{abs(n):,}".replace(",", ".")
    return ("-" if n < 0 else "") + s


# ------------------------------------------------------------------ xuất file
FORM = [  # (STT, nội dung, mã cột giá trị, mã cột thuế, kiểu dòng)
    ("A", "Không phát sinh hoạt động mua, bán trong kỳ (đánh dấu \"X\")", "21", None, "chk"),
    ("B", "Thuế giá trị gia tăng (GTGT) còn được khấu trừ kỳ trước chuyển sang", None, "22", "b"),
    ("C", "Kê khai thuế GTGT phải nộp Ngân sách nhà nước", None, None, "h"),
    ("I", "Hàng hóa, dịch vụ (HHDV) mua vào trong kỳ", None, None, "h"),
    ("1", "Giá trị và thuế GTGT của hàng hóa, dịch vụ mua vào", "23", "24", ""),
    ("", "Trong đó: hàng hóa, dịch vụ nhập khẩu", "23a", "24a", "i"),
    ("2", "Thuế GTGT của hàng hóa, dịch vụ mua vào được khấu trừ kỳ này", None, "25", ""),
    ("II", "Hàng hóa, dịch vụ bán ra trong kỳ", None, None, "h"),
    ("1", "Hàng hóa, dịch vụ bán ra không chịu thuế GTGT", "26", None, ""),
    ("2", "Hàng hóa, dịch vụ bán ra chịu thuế GTGT ([27]=[29]+[30]+[32]+[32a]; [28]=[31]+[33])", "27", "28", ""),
    ("a", "Hàng hoá, dịch vụ bán ra chịu thuế suất 0%", "29", None, ""),
    ("b", "Hàng hoá, dịch vụ bán ra chịu thuế suất 5%", "30", "31", ""),
    ("c", "Hàng hoá, dịch vụ bán ra chịu thuế suất 10%", "32", "33", ""),
    ("d", "Hàng hoá, dịch vụ bán ra không tính thuế", "32a", None, ""),
    ("3", "Tổng doanh thu và thuế GTGT của HHDV bán ra ([34]=[26]+[27]; [35]=[28])", "34", "35", "b"),
    ("III", "Thuế GTGT phát sinh trong kỳ ([36]=[35]-[25])", None, "36", "b"),
    ("IV", "Điều chỉnh tăng, giảm thuế GTGT còn được khấu trừ của các kỳ trước", None, None, "h"),
    ("1", "Điều chỉnh giảm", None, "37", ""),
    ("2", "Điều chỉnh tăng", None, "38", ""),
    ("V", "Thuế GTGT nhận bàn giao được khấu trừ trong kỳ", None, "39a", "b"),
    ("VI", "Xác định nghĩa vụ thuế GTGT phải nộp trong kỳ:", None, None, "h"),
    ("1", "Thuế GTGT phải nộp của hoạt động sản xuất kinh doanh trong kỳ (nếu [40a]=([36]-[22]+[37]-[38]-[39a]) ≥ 0)",
     None, "40a", ""),
    ("2", "Thuế GTGT mua vào của dự án đầu tư được bù trừ với thuế GTGT còn phải nộp của hoạt động sản xuất kinh "
          "doanh cùng kỳ tính thuế", None, "40b", ""),
    ("3", "Thuế GTGT còn phải nộp trong kỳ ([40]=[40a]-[40b])", None, "40", "b"),
    ("4", "Thuế GTGT chưa khấu trừ hết kỳ này (nếu ([41]=[36]-[22]+[37]-[38]-[39a]) < 0)", None, "41", "b"),
    ("4.1", "Thuế GTGT đề nghị hoàn", None, "42", ""),
    ("4.2", "Thuế GTGT còn được khấu trừ chuyển kỳ sau ([43]=[41]-[42])", None, "43", "b"),
]


def _tt89_form() -> list[tuple]:
    """Mẫu 01/GTGT theo TT 89/2026: thêm [32b], [34a]; [27]=[29]+[30]+[32]+[32a]-[32b]; [34]=[26]+[27]+[34a];
    [39a] và [40b] đổi nội dung (theo hướng dẫn trên thuvienphapluat.vn)."""
    out = []
    for row in FORM:
        stt, text, cv, cx, kind = row
        if cv == "27":
            row = (stt, "Hàng hóa, dịch vụ bán ra chịu thuế GTGT ([27]=[29]+[30]+[32]+[32a]-[32b]; [28]=[31]+[33])",
                   cv, cx, kind)
        elif cv == "34":
            out.append(("e", "Hàng hóa, dịch vụ bán ra không thuộc phạm vi điều chỉnh của pháp luật thuế GTGT", "34a",
                        None, ""))
            row = (stt, "Tổng doanh thu và thuế GTGT của HHDV bán ra ([34]=[26]+[27]+[34a]; [35]=[28])", cv, cx, kind)
        elif cx == "39a":
            row = (stt, "Thuế GTGT còn được khấu trừ chưa đề nghị hoàn của dự án đầu tư chuyển cho người nộp thuế "
                        "tiếp tục khấu trừ", cv, cx, kind)
        elif cx == "40b":
            row = (stt, "Tổng số thuế GTGT đã khai tại chỉ tiêu [28a] và [28b] của Tờ khai mẫu số 02/GTGT", cv, cx, kind)
        out.append(row)
        if cv == "32a":
            out.append(("đ", "Hàng hóa, dịch vụ bán ra không tính vào giá tính thuế GTGT", "32b", None, ""))
    return out


FORM_TT89 = _tt89_form()


def form_for(mau: str) -> list[tuple]:
    return FORM_TT89 if mau == "TT89" else FORM


def out_folder(mst: str, key: str) -> Path:
    return Path(settings.output_dir) / mst / "DuBaoVAT" / key


def export(res: dict[str, Any]) -> dict[str, str]:
    folder = out_folder(res["mst"], res["key"])
    folder.mkdir(parents=True, exist_ok=True)
    tag = res["key"].replace("-", "_")
    files = {"xlsx": str(_write_excel(res, folder / f"ToKhai_01GTGT_{tag}.xlsx"))}
    html = render_html(res)
    hp = folder / f"ToKhai_01GTGT_{tag}.html"
    hp.write_text(html, encoding="utf-8")
    files["html"] = str(hp)
    try:
        from .providers.headless import HEADLESS
        pdf = HEADLESS.html_to_pdf(html, 300)
        pp = folder / f"ToKhai_01GTGT_{tag}.pdf"
        try:
            pp.write_bytes(pdf)
        except PermissionError:  # đang mở trong trình đọc PDF
            pp = folder / f"ToKhai_01GTGT_{tag}_{datetime.now():%H%M%S}.pdf"
            pp.write_bytes(pdf)
        files["pdf"] = str(pp)
    except Exception as e:  # noqa: BLE001 - không có Chrome/Edge: vẫn còn bản HTML để in
        files["pdf_error"] = str(e)[:200]
    return files


def render_html(res: dict[str, Any]) -> str:
    from jinja2 import Environment, FileSystemLoader, select_autoescape
    env = Environment(loader=FileSystemLoader(Path(__file__).resolve().parent / "templates"),
                      autoescape=select_autoescape(["html"]))
    env.filters["m"] = fmt
    env.filters["rl"] = rate_label
    groups = {d: _groups(res["bk"][d]) for d in ("sold", "purchase")}
    return env.get_template("vat_print.html").render(r=res, form=form_for(res.get("mau", "TT80")), groups=groups, now=datetime.now())


def _groups(lines: list[dict[str, Any]]) -> list[tuple[str, list[dict[str, Any]], int, int]]:
    out = []
    for c in RATE_ORDER + sorted({b["ts"] for b in lines} - set(RATE_ORDER)):
        g = [b for b in lines if b["ts"] == c]
        if g:
            out.append((c, g, sum(b["gt"] for b in g), sum(b["thue"] for b in g)))
    return out


def _write_excel(res: dict[str, Any], dest: Path) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    from .reports import _save
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    bold, head = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
    wrap = Alignment(wrap_text=True, vertical="top")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    money = "#,##0;-#,##0;0"
    ct = res["ct"]
    wb = Workbook()

    # ----- Tờ khai 01/GTGT
    ws = wb.active
    ws.title = "To khai 01GTGT"
    for col, w in zip("ABCDEF", (7, 62, 7, 20, 7, 20)):
        ws.column_dimensions[col].width = w
    lines = [
        (f"Mẫu số: 01/GTGT ({res['mau_text']})", False),
        ("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", True), ("Độc lập - Tự do - Hạnh phúc", True), ("", False),
        ("TỜ KHAI THUẾ GIÁ TRỊ GIA TĂNG (Mẫu số 01/GTGT)", True),
        ("(Áp dụng đối với người nộp thuế tính thuế theo phương pháp khấu trừ có hoạt động sản xuất kinh doanh)", False),
        ("BẢN DỰ BÁO do AWBook lập từ hóa đơn điện tử – để kiểm tra, không thay thế tờ khai nộp qua eTax/HTKK", True),
        ("[01a] Tên hoạt động sản xuất kinh doanh: Hoạt động sản xuất kinh doanh thông thường", False),
        (f"[01b] Kỳ tính thuế: {res['label']}" + (f"  –  SỐ TẠM TÍNH ĐẾN NGÀY "
                                                 f"{date.fromisoformat(res['den_ngay']):%d/%m/%Y} (kỳ chưa kết thúc)"
                                                 if res.get("chua_het_ky") else ""), False), ("[02] Lần đầu: [X]          [03] Bổ sung lần thứ: ", False),
        (f"[04] Tên người nộp thuế: {res['company']['ten']}", False), (f"[05] Mã số thuế: {res['mst']}", False),
        ("[06] Tên đại lý thuế (nếu có):          [07] Mã số thuế:", False),
        ("[08] Hợp đồng đại lý thuế: Số: ...... Ngày: ......", False),
        ("[09] Tên đơn vị phụ thuộc/địa điểm kinh doanh khác tỉnh: ", False),
        ("[10] Mã số thuế đơn vị phụ thuộc/Mã số địa điểm kinh doanh: ", False),
        ("[11] Địa chỉ nơi có hoạt động kinh doanh khác tỉnh nơi đóng trụ sở chính: ", False),
    ]
    r = 1
    for text, b in lines:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        c = ws.cell(r, 1, text)
        c.font = Font(bold=b, size=13 if "TỜ KHAI" in text else 11, color="C00000" if "DỰ BÁO" in text else None)
        if b or r <= 6:
            c.alignment = Alignment(horizontal="center", wrap_text=True)
        r += 1
    ws.cell(r, 6, "Đơn vị tiền: Đồng Việt Nam").font = Font(italic=True)
    r += 1
    for i, h in enumerate(["STT", "CHỈ TIÊU", "", "GIÁ TRỊ HHDV (CHƯA CÓ THUẾ GTGT)", "", "THUẾ GTGT"], 1):
        c = ws.cell(r, i, h)
        c.font, c.fill, c.alignment, c.border = bold, head, center, box
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
    r += 1
    for stt, text, cv, cx, kind in form_for(res.get("mau", "TT80")):
        vals = [stt, text, f"[{cv}]" if cv else "", None, f"[{cx}]" if cx else "", None]
        if kind == "chk":
            vals[3] = "X" if res["ct21"] else ""
        elif cv:
            vals[3] = ct.get(cv, 0)
        if cx:
            vals[5] = ct.get(cx, 0)
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            c.border = box
            c.alignment = wrap if i == 2 else Alignment(horizontal="center" if i in (1, 3, 5) else "right")
            if i in (4, 6) and isinstance(v, (int, float)):
                c.number_format = money
            if kind in ("h", "b"):
                c.font = bold
            if kind == "i" and i == 2:
                c.font = Font(italic=True)
        r += 1
    r += 1
    for text in ("Tôi cam đoan số liệu khai trên là đúng và chịu trách nhiệm trước pháp luật về số liệu đã khai./.",
                 f"Hạn nộp tờ khai kỳ này: {date.fromisoformat(res['deadline']):%d/%m/%Y}",
                 "NGƯỜI NỘP THUẾ hoặc ĐẠI DIỆN HỢP PHÁP CỦA NGƯỜI NỘP THUẾ (Chữ ký, ghi rõ họ tên; chức vụ và đóng dấu "
                 "(nếu có)/Ký điện tử)"):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.cell(r, 1, text).alignment = Alignment(wrap_text=True)
        r += 1
    ws.print_options.horizontalCentered = True
    ws.page_setup.paperSize, ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 9, 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ----- Phụ lục 8%
    def pl_sheet(title: str, heading: list[str], data: dict[str, Any]) -> None:
        w = wb.create_sheet(title)
        for col, wd in zip("ABCDEF", (6, 52, 22, 14, 14, 18)):
            w.column_dimensions[col].width = wd
        rr = 1
        for t in heading:
            w.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=6)
            c = w.cell(rr, 1, t)
            c.alignment = Alignment(horizontal="center", wrap_text=True)
            c.font = Font(bold=rr <= 3)
            rr += 1
        w.cell(rr, 6, "Đơn vị tiền: Đồng Việt Nam").font = Font(italic=True)
        rr += 1
        hd = ["STT", "Tên hàng hóa, dịch vụ", "Giá trị HHDV chưa có thuế GTGT / Doanh thu HHDV chịu thuế",
              "Thuế suất theo quy định", "Thuế suất sau giảm", "Thuế GTGT được giảm"]
        for i, h in enumerate(hd, 1):
            c = w.cell(rr, i, h)
            c.font, c.fill, c.alignment, c.border = bold, head, center, box
        rr += 1
        for i, h in enumerate(["(1)", "(2)", "(3)", "(4)", "(5)", "(6)=(3)x[(4)-(5)]"], 1):
            c = w.cell(rr, i, h)
            c.alignment, c.border = center, box
        rr += 1
        for n, x in enumerate(data["rows"], 1):
            for i, v in enumerate([n, x["ten"], x["gt"], "10%", "8%", x["giam"]], 1):
                c = w.cell(rr, i, v)
                c.border = box
                if i in (3, 6):
                    c.number_format = money
                if i in (1, 4, 5):
                    c.alignment = Alignment(horizontal="center")
            rr += 1
        for i, v in enumerate(["", "Tổng cộng", data["gt"], "", "", data["giam"]], 1):
            c = w.cell(rr, i, v)
            c.font, c.border = bold, box
            if i in (3, 6):
                c.number_format = money

    if res["pl8"]["rows"]:
        pl_sheet("PL giam thue 8%", [
            "Mẫu số 01 – PHỤ LỤC", f"GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO {res['basis'].upper()}",
            f"(Kèm theo Tờ khai 01/GTGT kỳ tính thuế {res['label']})",
            f"[01] Tên người nộp thuế: {res['company']['ten']}      [02] Mã số thuế: {res['mst']}"], res["pl8"])
    if res["pl8_mua"]["rows"]:
        pl_sheet("BK mua vao 8%", [
            "BẢNG KÊ HÀNG HÓA, DỊCH VỤ MUA VÀO ÁP DỤNG THUẾ SUẤT 8%", "(tham khảo – không phải phụ lục bắt buộc)",
            f"Kỳ tính thuế {res['label']} – MST {res['mst']}", ""], res["pl8_mua"])

    # ----- Bảng kê hóa đơn
    def bk_sheet(title: str, heading: str, lines: list[dict[str, Any]], who: str) -> None:
        w = wb.create_sheet(title)
        cols = [("STT", 6), ("Mẫu số", 7), ("Ký hiệu", 10), ("Số hóa đơn", 11), ("Ngày lập", 11),
                (f"Tên {who}", 40), (f"MST {who}", 15), ("Giá trị chưa thuế", 17), ("Thuế suất", 9),
                ("Thuế GTGT", 15), ("Trạng thái HĐ", 14), ("Ghi chú", 40)]
        w.merge_cells("A1:L1")
        w.cell(1, 1, heading).font = Font(bold=True, size=13)
        w.cell(2, 1, f"Kỳ tính thuế: {res['label']} – Người nộp thuế: {res['company']['ten']} – MST {res['mst']}")
        for i, (h, wd) in enumerate(cols, 1):
            c = w.cell(4, i, h)
            c.font, c.fill, c.alignment, c.border = bold, head, center, box
            w.column_dimensions[c.column_letter].width = wd
        rr, n = 5, 0
        tot_v = tot_t = 0
        for c_rate, g, sv, st in _groups(lines):
            w.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=12)
            w.cell(rr, 1, rate_label(c_rate)).font = bold
            rr += 1
            for b in g:
                n += 1
                vals = [n, b["mau"], b["kh"], b["so"], b["ngay"], b["ten"], b["mst"], b["gt"],
                        rate_label(b["ts"]).replace("Thuế suất ", "").replace(" (giảm 2%)", ""), b["thue"], b["tthai"],
                        b["ghi_chu"]]
                for i, v in enumerate(vals, 1):
                    c = w.cell(rr, i, v)
                    c.border = box
                    if i in (8, 10):
                        c.number_format = money
                rr += 1
            for i, v in ((6, f"Cộng {rate_label(c_rate)}"), (8, sv), (10, st)):
                c = w.cell(rr, i, v)
                c.font = bold
                if i != 6:
                    c.number_format = money
            tot_v += sv
            tot_t += st
            rr += 1
        for i, v in ((6, "TỔNG CỘNG"), (8, tot_v), (10, tot_t)):
            c = w.cell(rr, i, v)
            c.font = Font(bold=True, color="C00000")
            if i != 6:
                c.number_format = money
        w.freeze_panes = "A5"

    bk_sheet("BK ban ra", "BẢNG KÊ HÓA ĐƠN HÀNG HÓA, DỊCH VỤ BÁN RA", res["bk"]["sold"], "người mua")
    bk_sheet("BK mua vao", "BẢNG KÊ HÓA ĐƠN HÀNG HÓA, DỊCH VỤ MUA VÀO", res["bk"]["purchase"], "người bán")

    # ----- Chi tiết từng mặt hàng
    def item_sheet(title: str, data: list[dict[str, Any]], who: str) -> None:
        w = wb.create_sheet(title)
        cols = [("STT", 6), ("Ngày lập", 11), ("Ký hiệu", 10), ("Số HĐ", 10), (f"MST {who}", 14), (f"Tên {who}", 32),
                ("Tên hàng hóa, dịch vụ", 44), ("ĐVT", 8), ("Số lượng", 11), ("Đơn giá", 14), ("Thành tiền", 16),
                ("Thuế suất", 9), ("Tiền thuế GTGT", 15), ("Tổng cộng", 16)]
        for i, (h, wd) in enumerate(cols, 1):
            c = w.cell(1, i, h)
            c.font, c.fill, c.alignment, c.border = bold, head, center, box
            w.column_dimensions[c.column_letter].width = wd
        for n, it in enumerate(data, 1):
            vals = [n, it["ngay"], it["kh"], it["so"], it["mst"], it["ten_dt"], it["ten"], it["dvt"], it["sl"], it["dg"],
                    it["tt"], rate_label(it["ts"]).replace("Thuế suất ", "").replace(" (giảm 2%)", ""), it["thue"],
                    it["tong"]]
            for i, v in enumerate(vals, 1):
                c = w.cell(n + 1, i, v)
                if i in (9, 10, 11, 13, 14):
                    c.number_format = "#,##0.##" if i in (9, 10) else money
        rr = len(data) + 2
        for i, v in ((7, "TỔNG CỘNG"), (11, sum(x["tt"] for x in data)), (13, sum(x["thue"] or 0 for x in data)),
                     (14, sum(x["tong"] for x in data))):
            c = w.cell(rr, i, v)
            c.font = bold
            if i != 7:
                c.number_format = money
        w.freeze_panes = "A2"
        w.auto_filter.ref = f"A1:N{max(2, len(data) + 1)}"

    item_sheet("CT HHDV ban ra", res["items"]["sold"], "người mua")
    item_sheet("CT HHDV mua vao", res["items"]["purchase"], "người bán")

    # ----- HĐ ngân hàng + HĐ thay thế / điều chỉnh (luôn có sheet, không có HĐ thì để trống)
    from openpyxl.utils import get_column_letter

    def lq_sheet(title: str, heading: str, data: list[dict[str, Any]], lien_quan: str, sau: bool = False) -> None:
        w = wb.create_sheet(title)
        cols = [("STT", 6), ("Hướng", 9), ("Loại hóa đơn", 24), ("Ngày lập", 11), ("Mẫu số", 7), ("Ký hiệu", 10),
                ("Số HĐ", 10), ("MST đối tác", 15), ("Tên đối tác", 36), ("Giá trị chưa thuế", 17), ("Thuế GTGT", 15),
                ("Tổng thanh toán", 17), ("Trạng thái HĐ", 16)]
        if lien_quan:
            cols.append((lien_quan, 18))
        if sau:
            cols += [("Giá trị chưa thuế sau điều chỉnh", 20), ("Thuế GTGT sau điều chỉnh", 17)]
        cols.append(("Xử lý trên tờ khai", 60))
        money_cols = {"Giá trị chưa thuế", "Thuế GTGT", "Tổng thanh toán", "Giá trị chưa thuế sau điều chỉnh",
                      "Thuế GTGT sau điều chỉnh"}
        w.merge_cells(f"A1:{get_column_letter(len(cols))}1")
        w.cell(1, 1, heading).font = Font(bold=True, size=13)
        w.cell(2, 1, f"Kỳ tính thuế: {res['label']} – Người nộp thuế: {res['company']['ten']} – MST {res['mst']}")
        for i, (h, wd) in enumerate(cols, 1):
            c = w.cell(4, i, h)
            c.font, c.fill, c.alignment, c.border = bold, head, center, box
            w.column_dimensions[c.column_letter].width = wd
        if not data:
            w.cell(5, 1, "Không có hóa đơn nào trong kỳ.").font = Font(italic=True, color="808080")
        for n, x in enumerate(data, 1):
            vals = [n, x["huong"], x["loai"], x["ngay"], x["mau"], x["kh"], x["so"], x["mst"], x["ten"], x["gt"], x["thue"],
                    x["tong"], x["tthai"]]
            if lien_quan:
                vals.append(x["lien_quan"])
            if sau:
                vals += [x.get("gt_sau"), x.get("thue_sau")]
            vals.append(x["xu_ly"])
            for i, v in enumerate(vals, 1):
                c = w.cell(n + 4, i, v)
                c.border = box
                if cols[i - 1][0] in money_cols:
                    c.number_format = money
                if i == len(vals):
                    c.alignment = Alignment(wrap_text=True, vertical="top")
        if data:
            rr = len(data) + 5
            w.cell(rr, 9, "TỔNG CỘNG").font = bold
            for i, k in ((10, "gt"), (11, "thue"), (12, "tong")):
                c = w.cell(rr, i, sum(x[k] for x in data))
                c.font, c.number_format = bold, money
        w.freeze_panes = "A5"

    lq = res.get("lien_quan") or {}
    lq_sheet("HD ngan hang", "HÓA ĐƠN CỦA NGÂN HÀNG (ĐÃ TÁCH RA – KHÔNG ĐƯA VÀO BẢNG KÊ MUA VÀO VÀ TỜ KHAI)",
             lq.get("ngan_hang", []), "")
    lq_sheet("HD bi thay the", "HÓA ĐƠN BỊ THAY THẾ", lq.get("bi_thay_the", []), "Bị thay thế bởi HĐ")
    lq_sheet("HD thay the", "HÓA ĐƠN THAY THẾ", lq.get("thay_the", []), "Thay thế cho HĐ")
    lq_sheet("HD bi dieu chinh", "HÓA ĐƠN BỊ ĐIỀU CHỈNH", lq.get("bi_dieu_chinh", []), "Bị điều chỉnh bởi HĐ", sau=True)
    lq_sheet("HD dieu chinh", "HÓA ĐƠN ĐIỀU CHỈNH", lq.get("dieu_chinh", []), "Điều chỉnh cho HĐ")

    # ----- Cảnh báo
    w = wb.create_sheet("Canh bao", 1)
    w.column_dimensions["A"].width = 120
    w.cell(1, 1, f"LƯU Ý KHI DỰ BÁO {res['label'].upper()}").font = Font(bold=True, size=13)
    rr = 2
    for t in res["warnings"]:
        w.cell(rr, 1, "• " + t).alignment = Alignment(wrap_text=True)
        rr += 1
    rr += 1
    hd = ["Hướng", "Ngày", "Ký hiệu", "Số HĐ", "MST đối tác", "Tên đối tác", "Tổng thanh toán", "Lưu ý"]
    for i, h in enumerate(hd, 1):
        c = w.cell(rr, i, h)
        c.font, c.fill = bold, head
    for x in res["warn_rows"]:
        rr += 1
        for i, v in enumerate([x["huong"], x["ngay"], x["kh"], x["so"], x["mst"], x["ten"], x["tong"], x["why"]], 1):
            c = w.cell(rr, i, v)
            if i == 7:
                c.number_format = money
    for x in res["excluded"]:
        rr += 1
        for i, v in enumerate([x["huong"], x["ngay"], x["kh"], x["so"], "", "", x["tong"],
                               f"Đã loại khỏi tờ khai ({x['tthai']}): {x.get('why') or ''}"], 1):
            w.cell(rr, i, v)
    return _save(wb, dest)


# ------------------------------------------------------------------ cài đặt + lịch sử
def _load_state() -> dict[str, Any]:
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return {}


def _save_state(s: dict[str, Any]) -> None:
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(s, ensure_ascii=False, indent=1), encoding="utf-8")


def get_cfg(mst: str) -> dict[str, Any]:
    with _lock:
        c = _load_state().get(mst) or {}
    return {"kieu": c.get("kieu", "M"), "auto": c.get("auto", True), "inputs": c.get("inputs", {}),
            "auto_state": c.get("auto_state", {})}


def set_cfg(mst: str, **kw: Any) -> dict[str, Any]:
    with _lock:
        s = _load_state()
        c = s.setdefault(mst, {})
        for k, v in kw.items():
            if k == "inputs":
                c.setdefault("inputs", {}).update(v)
            else:
                c[k] = v
        _save_state(s)
    return get_cfg(mst)


def saved_result(mst: str, key: str) -> dict[str, Any] | None:
    f = out_folder(mst, key) / "ketqua.json"
    try:
        return json.loads(f.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return None


def suggest_ct22(mst: str, kieu: str, ky: int, nam: int) -> dict[str, Any]:
    """[22] kỳ này: số đã nhập cho kỳ này > [43] của kỳ trước (đã dự báo) > 0."""
    key = period_key(kieu, ky, nam)
    cfg = get_cfg(mst)
    if key in cfg["inputs"] and "ct22" in cfg["inputs"][key]:
        return {"value": cfg["inputs"][key]["ct22"], "source": "đã nhập"}
    pk = period_key(*prev_period(kieu, ky, nam))
    prev = saved_result(mst, pk)
    if prev:
        return {"value": prev["ct"].get("43", 0), "source": f"[43] kỳ {prev['label']}"}
    return {"value": 0, "source": ""}


def summary(res: dict[str, Any]) -> dict[str, Any]:
    ct = res["ct"]
    return {k: res[k] for k in ("mau", "mau_text", "legal", "mst", "kieu", "ky", "nam", "key", "label", "from", "to", "deadline", "company",
                                "inputs", "ct", "ct21", "stat", "missing", "partial", "warnings", "rows", "basis")} | {
        "chua_het_ky": res.get("chua_het_ky", False), "den_ngay": res.get("den_ngay"),
        "label_den": res.get("label_den") or res["label"], "tom_tat": res.get("tom_tat", ""),
        "pl8": {"gt": res["pl8"]["gt"], "giam": res["pl8"]["giam"], "n": len(res["pl8"]["rows"])},
        "pl8_mua": {"gt": res["pl8_mua"]["gt"], "n": len(res["pl8_mua"]["rows"])},
        "phai_nop": ct["40"], "chuyen_ky_sau": ct["43"], "files": res.get("files", {}),
        "created": res.get("created"), "source": res.get("source"),
        "n_items": {"sold": len(res["items"]["sold"]), "purchase": len(res["items"]["purchase"])},
    }


def history(mst: str) -> list[dict[str, Any]]:
    base = Path(settings.output_dir) / mst / "DuBaoVAT"
    out = []
    if base.is_dir():
        for f in base.glob("*/ketqua.json"):
            try:
                d = json.loads(f.read_text(encoding="utf-8"))
                out.append({k: d.get(k) for k in ("key", "label", "phai_nop", "chuyen_ky_sau", "created", "source", "files")})
            except ValueError:
                continue
    return sorted(out, key=lambda x: x["key"], reverse=True)


def forecast(mst: str, kieu: str, ky: int, nam: int, inputs: dict[str, Any] | None = None,
             source: str = "manual", notify: bool = True) -> dict[str, Any]:
    if inputs is None:
        inputs = dict(get_cfg(mst)["inputs"].get(period_key(kieu, ky, nam), {}))
        if "ct22" not in inputs:
            inputs["ct22"] = suggest_ct22(mst, kieu, ky, nam)["value"]
    else:
        set_cfg(mst, inputs={period_key(kieu, ky, nam): {k: inputs.get(k, 0) for k in INPUT_KEYS}})
    res = compute(mst, kieu, ky, nam, inputs)
    res["files"] = export(res)
    res["created"] = datetime.now().isoformat(timespec="seconds")
    res["source"] = source
    try:  # tóm tắt theo văn mẫu gửi khách ("Dạ anh/chị, em xin gửi... tính tới thời điểm...")
        from . import vat_client
        res["tom_tat"] = vat_client.message(vat_client.build(res))
    except Exception:  # noqa: BLE001
        res["tom_tat"] = ""
    s = summary(res)
    out_folder(mst, res["key"]).mkdir(parents=True, exist_ok=True)
    (out_folder(mst, res["key"]) / "ketqua.json").write_text(json.dumps(s, ensure_ascii=False, indent=1),
                                                             encoding="utf-8")
    if notify:
        notify_result(s)
    return s


# ------------------------------------------------------------------ thông báo trong app
def _load_noti() -> list[dict[str, Any]]:
    try:
        return json.loads(NOTI_FILE.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return []


def add_notification(title: str, body: str, level: str = "info", **extra: Any) -> dict[str, Any]:
    n = {"id": uuid.uuid4().hex[:12], "time": datetime.now().isoformat(timespec="seconds"), "title": title,
         "body": body, "level": level, "read": False, **extra}
    with _lock:
        items = [n] + _load_noti()
        NOTI_FILE.parent.mkdir(parents=True, exist_ok=True)
        NOTI_FILE.write_text(json.dumps(items[:200], ensure_ascii=False, indent=1), encoding="utf-8")
    return n


def notifications() -> list[dict[str, Any]]:
    with _lock:
        return _load_noti()


def mark_read(ids: list[str] | None = None) -> None:
    with _lock:
        items = _load_noti()
        for n in items:
            if ids is None or n["id"] in ids:
                n["read"] = True
        NOTI_FILE.write_text(json.dumps(items, ensure_ascii=False, indent=1), encoding="utf-8")


def notify_result(s: dict[str, Any]) -> None:
    ct = s["ct"]
    head = (f"Phải nộp {fmt(ct['40'])} đ" if ct["40"] > 0 else
            f"Còn được khấu trừ chuyển kỳ sau {fmt(ct['43'])} đ" if ct["43"] > 0 else "Không phát sinh số phải nộp")
    title = f"Dự báo VAT {s.get('label_den') or s['label']} – MST {s['mst']}: {head}"
    body = ((s.get("tom_tat") + chr(10)) if s.get("tom_tat") else "") + (f"Bán ra: doanh thu [34] {fmt(ct['34'])} đ, thuế [35] {fmt(ct['35'])} đ ({s['stat']['ban_ra']} HĐ). "
            f"Mua vào: [23] {fmt(ct['23'])} đ, thuế được khấu trừ [25] {fmt(ct['25'])} đ ({s['stat']['mua_vao']} HĐ). "
            f"[22] kỳ trước chuyển sang {fmt(ct['22'])} đ. Thuế phát sinh [36] {fmt(ct['36'])} đ. "
            + (f"Phải nộp [40] {fmt(ct['40'])} đ, hạn nộp {date.fromisoformat(s['deadline']):%d/%m/%Y}."
               if ct["40"] > 0 else f"Chuyển kỳ sau [43] {fmt(ct['43'])} đ.")
            + (f" Bán ra 8%: giảm {fmt(s['pl8']['giam'])} đ (có phụ lục)." if s["pl8"]["n"] else ""))
    level = "warn" if s["missing"] or s["partial"] else "ok"
    add_notification(title, body, level, kind="vat", mst=s["mst"], key=s["key"], warnings=s["warnings"][:8],
                     files=s.get("files", {}))
    try:
        from .scheduler import toast
        toast(title, body[:250])
    except Exception:  # noqa: BLE001
        pass


# ------------------------------------------------------------------ tự dự báo đầu kỳ
DOWNLOAD_KINDS = ["mua-vao", "mua-vao-khong-ma", "ban-ra"]
GIVE_UP_DAYS = 10


def _msts() -> list[str]:
    base = Path(settings.output_dir)
    if not base.is_dir():
        return []
    return [d.name for d in base.iterdir() if d.is_dir() and re.fullmatch(r"\d{10}(-?\d{3})?", d.name)]


def _months_to_download(root: Path, kieu: str, ky: int, nam: int) -> list[tuple[int, int]]:
    _, missing, partial = load_rows(root, kieu, ky, nam)
    bad = {s.split(" ")[0] for s in missing + partial}
    return [(y, m) for y, m in period_months(kieu, ky, nam) if f"{m:02d}/{y}" in bad]


def auto_tick(now: datetime | None = None) -> None:
    """Gọi định kỳ (30 giây). Từ 00:00 ngày đầu kỳ sau: tải hóa đơn kỳ vừa hết (nếu Chrome đang đăng nhập đúng MST)
    rồi dự báo và báo lên app. Chưa đăng nhập: dự báo ngay bằng dữ liệu đang có, báo cần đăng nhập để cập nhật."""
    from .auto_download import JOBS, running_job, start
    from .gdt.browser import BROWSER

    now = now or datetime.now()
    for mst in _msts():
        cfg = get_cfg(mst)
        if not cfg["auto"]:
            continue
        kieu, ky, nam = last_completed(now.date(), cfg["kieu"])
        key = period_key(kieu, ky, nam)
        st = dict(cfg["auto_state"])
        if st.get("done") == key:
            continue
        root = Path(settings.output_dir) / mst
        # MST không có dữ liệu ở kỳ này lẫn kỳ trước -> khách hàng không còn theo dõi, bỏ qua lặng lẽ
        pk = prev_period(kieu, ky, nam)
        has_now = any((root / f"{y:04d}-{m:02d}").is_dir() for y, m in period_months(kieu, ky, nam))
        has_prev = any((root / f"{y:04d}-{m:02d}").is_dir() for y, m in period_months(*pk))
        if not (has_now or has_prev):
            set_cfg(mst, auto_state={**st, "done": key})
            continue
        if st.get("job"):
            job = JOBS.get(st["job"])
            if job and job.status == "running":
                continue
            st["downloaded"] = key
            st.pop("job", None)
            set_cfg(mst, auto_state=st)
        need = _months_to_download(root, kieu, ky, nam)
        if need and st.get("downloaded") != key:
            bst = BROWSER.status()
            if bst.get("logged_in") and str(bst.get("mst")) == mst and not running_job():
                (y1, m1), (y2, m2) = need[0], need[-1]
                job = start(date(y1, m1, 1), date(y2, m2, monthrange(y2, m2)[1]), DOWNLOAD_KINDS, True, True, True)
                st["job"] = job.id
                set_cfg(mst, auto_state=st)
                add_notification(f"Đang tải hóa đơn {period_label(kieu, ky, nam)} – MST {mst}",
                                 "Kỳ đã kết thúc. Phần mềm đang tải hóa đơn mua vào / bán ra để dự báo thuế GTGT, "
                                 "xong sẽ báo kết quả.", "info", kind="vat-info", mst=mst, key=key)
                continue
            if st.get("partial") != key:
                st["partial"] = key
                set_cfg(mst, auto_state=st)
                if has_now:
                    forecast(mst, kieu, ky, nam, source="auto-partial")
                add_notification(f"Cần đăng nhập cổng thuế để chốt dự báo VAT {period_label(kieu, ky, nam)} – MST {mst}",
                                 ("Đã dự báo tạm bằng dữ liệu đang có. " if has_now else "Chưa có hóa đơn của kỳ này. ")
                                 + "Mở mục Tải hóa đơn, đăng nhập hoadondientu bằng MST này – phần mềm sẽ tự tải "
                                   "đủ hóa đơn rồi dự báo lại.", "warn", kind="vat-info", mst=mst, key=key)
            if now.date() > period_range(kieu, ky, nam)[1] + timedelta(days=GIVE_UP_DAYS):
                set_cfg(mst, auto_state={**st, "done": key})
            continue
        forecast(mst, kieu, ky, nam, source="auto")
        st["done"] = key
        set_cfg(mst, auto_state=st)
