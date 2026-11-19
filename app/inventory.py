"""Tổng hợp tồn kho theo mẫu MISA AMIS (báo cáo "TỔNG HỢP TỒN KHO").

Đầu vào:
  * File tồn kho đầu kỳ: báo cáo Tổng hợp tồn kho kỳ trước của MISA (lấy cột CUỐI KỲ làm đầu kỳ này) hoặc file
    Danh mục vật tư hàng hóa (chỉ có mã / tên / ĐVT, số lượng đầu kỳ = 0).
  * File mua vào (mẫu "Mua hàng trong nước" của MISA), file bán ra (mẫu "Bán hàng"), file hàng bán trả lại (tùy chọn)
    – đã có dữ liệu từ hóa đơn.
  * File mẫu Tổng hợp tồn kho để xuất (không có thì dùng chính file đầu kỳ).

Cách tính (như MISA – giá xuất kho bình quân cuối kỳ):
  * Ghép MÃ HÀNG cho từng dòng mua vào / bán ra theo tên hàng (mã có sẵn trong danh mục -> tên trùng -> tên gần đúng);
    hàng chưa có trong danh mục -> tạo mã mới (kèm file danh mục VTHH để khai vào MISA).
  * Nhập kho = mua vào (thành tiền - chiết khấu) + hàng bán trả lại (theo đơn giá bình quân).
  * Đơn giá bình quân = (giá trị đầu kỳ + giá trị mua vào) / (SL đầu kỳ + SL mua vào); Xuất kho = SL bán x đơn giá bình quân.
  * Cuối kỳ = Đầu kỳ + Nhập - Xuất.
"""
from __future__ import annotations

import difflib
import io
import re
import uuid
from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from .config import settings
from .docreader import fold

DEFAULT_KHO = "Hàng hóa"


# ------------------------------------------------------------------ tiện ích
def _norm(s: Any) -> str:
    t = fold(str(s or "")).lower()
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _num(v: Any) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    from .docreader import vn_num
    x = vn_num(str(v).replace("(", "-").replace(")", ""))
    return float(x or 0)


def _date(v: Any) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    m = re.match(r"\s*(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})", str(v or ""))
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    m = re.match(r"\s*(\d{4})-(\d{2})-(\d{2})", str(v or ""))
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


# ------------------------------------------------------------------ đọc tồn kho đầu kỳ / danh mục
def read_opening(data: bytes) -> dict[str, Any]:
    """-> {"items": [{kho, ma, ten, dvt, sl, gt}], "loai": "tonkho"|"danhmuc"}"""
    from .misa_import import load_any_workbook
    wb = load_any_workbook(data, values=True)   # mã hàng có thể là công thức (VLOOKUP): lấy giá trị
    for ws in wb.worksheets:
        for r in range(1, min(ws.max_row, 30) + 1):
            heads = {c.column: _norm(c.value) for c in ws[r] if isinstance(c.value, str)}
            col = lambda *pats: next((c for c, h in heads.items() if any(h == p or h.startswith(p) for p in pats)), None)  # noqa: E731
            c_ma, c_ten = col("ma hang"), col("ten hang")
            if not (c_ma and c_ten):
                continue
            if col("dau ky") and col("cuoi ky"):          # báo cáo Tổng hợp tồn kho
                return {"loai": "tonkho", "items": _read_tonkho(ws, r, heads)}
            c_dvt = col("don vi tinh chinh", "don vi tinh", "dvt")
            c_kho = col("ma kho ngam dinh", "ma kho")
            items = []
            for rr in range(r + 1, ws.max_row + 1):
                ma = str(ws.cell(rr, c_ma).value or "").strip()
                ten = str(ws.cell(rr, c_ten).value or "").strip()
                if not ma or not ten or any(i["ma"] == ma for i in items):
                    continue
                items.append({"kho": DEFAULT_KHO, "ma_kho": str(ws.cell(rr, c_kho).value or "").strip() if c_kho else "",
                              "ma": ma, "ten": ten, "dvt": str(ws.cell(rr, c_dvt).value or "").strip() if c_dvt else "",
                              "sl": 0.0, "gt": 0.0})
            return {"loai": "danhmuc", "items": items}
    # Bảng kê 01-2/BK-HTK (hàng tồn kho, máy móc, thiết bị của hộ kinh doanh) hoặc bảng tồn kho tự lập bất kỳ:
    # cột tên hàng + số lượng + giá trị (+ mã hàng, ĐVT), tiêu đề có thể nằm trên 2-3 dòng
    for ws in wb.worksheets:
        items = _read_generic(ws)
        if items:
            return {"loai": "bangke", "items": items}
    raise ValueError("Không nhận ra file tồn kho đầu kỳ (cần báo cáo Tổng hợp tồn kho, Danh mục VTHH của MISA, bảng kê "
                     "01-2/BK-HTK hàng tồn kho, hoặc bảng có các cột Tên hàng, ĐVT, Số lượng, Giá trị, Mã hàng)")


def _read_generic(ws) -> list[dict[str, Any]]:
    """Bảng tồn kho dạng danh sách: tìm cột theo tiêu đề (trải trên vài dòng), đọc từng dòng hàng; dòng mục
    "I. HÀNG HÓA...", "II. MÁY MÓC, THIẾT BỊ" dùng làm Tên kho."""
    top = min(ws.max_row, 40)
    for hr in range(1, top + 1):
        cols: dict[str, int] = {}
        for rr in range(hr, min(hr + 3, ws.max_row + 1)):
            for c in range(1, min(ws.max_column, 40) + 1):
                h = _norm(ws.cell(rr, c).value)
                if not h:
                    continue
                if "ten" not in cols and (h.startswith("ten hang") or h.startswith("hang ton kho") or h.startswith("ten vat tu")
                                          or h == "hang hoa"):
                    cols["ten"] = c
                elif "ma" not in cols and (h.startswith("ma hang") or h.startswith("ma vat tu")):
                    cols["ma"] = c
                elif "dvt" not in cols and (h in ("dvt", "dv tinh") or h.startswith("don vi tinh")):
                    cols["dvt"] = c
                elif "sl" not in cols and h.startswith("so luong"):
                    cols["sl"] = c
                elif "gt" not in cols and (h.startswith("gia tri") or h == "thanh tien"):
                    cols["gt"] = c
                elif "kho" not in cols and h.startswith("ten kho"):
                    cols["kho"] = c
        if not ("ten" in cols and "sl" in cols):
            continue
        last_head = max(r for r in range(hr, min(hr + 3, ws.max_row + 1))
                        if any(_norm(ws.cell(r, c).value) for c in cols.values()))
        items: list[dict[str, Any]] = []
        kho = DEFAULT_KHO
        used: set[str] = set()
        for r in range(last_head + 1, ws.max_row + 1):
            ten = str(ws.cell(r, cols["ten"]).value or "").strip()
            first = str(ws.cell(r, 1).value or ws.cell(r, 2).value or "").strip()
            # dòng mục: "I. HÀNG HÓA, SẢN PHẨM TỒN KHO" / "II. MÁY MÓC, THIẾT BỊ"
            sec = first if re.match(r"^[IVX]+\s*[.)]", first) else (ten if re.match(r"^[IVX]+\s*[.)]", ten) else "")
            if sec:
                t = _norm(sec)
                kho = ("Máy móc, thiết bị" if "may moc" in t else "Công cụ dụng cụ" if "cong cu" in t
                       else "Nguyên vật liệu" if "nguyen" in t else "Thành phẩm" if "thanh pham" in t else DEFAULT_KHO)
                continue
            if not ten or re.match(r"^\[\d+\]$", ten) or _norm(ten).startswith(("tong cong", "cong")):
                continue
            ma = str(ws.cell(r, cols["ma"]).value or "").strip() if "ma" in cols else ""
            if not ma:
                ma = new_code(ten, used)
            used.add(ma.upper())
            items.append({"kho": str(ws.cell(r, cols["kho"]).value or "").strip() if "kho" in cols else kho, "ma_kho": "",
                          "ma": ma, "ten": ten, "dvt": str(ws.cell(r, cols["dvt"]).value or "").strip() if "dvt" in cols else "",
                          "sl": _num(ws.cell(r, cols["sl"]).value), "gt": _num(ws.cell(r, cols["gt"]).value) if "gt" in cols else 0.0})
        if items:
            return items
    return []


def _read_tonkho(ws, hr: int, heads: dict[int, str]) -> list[dict[str, Any]]:
    """Báo cáo Tổng hợp tồn kho MISA: tiêu đề 2 dòng (Đầu kỳ / Nhập kho / Xuất kho / Cuối kỳ x Số lượng / Giá trị),
    có thể lặp lại tiêu đề + chân trang mỗi trang in. Đầu kỳ kỳ này = CUỐI KỲ của báo cáo."""
    groups = sorted((c, h) for c, h in heads.items() if h in ("dau ky", "nhap kho", "xuat kho", "cuoi ky"))
    sub = {c: _norm(ws.cell(hr + 1, c).value) for c in range(1, ws.max_column + 1)}
    col_of = {}
    for i, (c, h) in enumerate(groups):
        end = groups[i + 1][0] if i + 1 < len(groups) else ws.max_column + 1
        sl = next((x for x in range(c, end) if sub.get(x) == "so luong"), None)
        gt = next((x for x in range(c, end) if sub.get(x) == "gia tri"), None)
        col_of[h] = (sl, gt)
    find = lambda p: next((c for c, h in heads.items() if h.startswith(p)), None)  # noqa: E731
    c_kho, c_ma, c_ten, c_dvt = find("ten kho"), find("ma hang"), find("ten hang"), find("dvt") or find("don vi tinh")
    sl_c, gt_c = col_of.get("cuoi ky", (None, None))
    items: list[dict[str, Any]] = []
    for r in range(hr + 2, ws.max_row + 1):
        ma = str(ws.cell(r, c_ma).value or "").strip()
        ten = str(ws.cell(r, c_ten).value or "").strip()
        if not ma or not ten or _norm(ma) == "ma hang":   # dòng tiêu đề lặp lại / tổng cộng / chân trang
            continue
        items.append({"kho": str(ws.cell(r, c_kho).value or "").strip() if c_kho else DEFAULT_KHO, "ma_kho": "",
                      "ma": ma, "ten": ten, "dvt": str(ws.cell(r, c_dvt).value or "").strip() if c_dvt else "",
                      "sl": _num(ws.cell(r, sl_c).value) if sl_c else 0.0, "gt": _num(ws.cell(r, gt_c).value) if gt_c else 0.0})
    return items


# ------------------------------------------------------------------ đọc file mua vào / bán ra / trả lại (mẫu MISA)
def read_lines(data: bytes) -> dict[str, Any]:
    """Dòng hàng của file chứng từ mẫu MISA -> {"lines": [...], "hr": dòng tiêu đề, "cols": {khóa: cột}}"""
    from .misa_import import read_template
    wb, ws, hr, cols, _ = read_template(data, values=True)
    by = {k: c for c, k in cols.items()}
    if "ten_hang" not in by and "ma_hang" not in by:
        raise ValueError("File không có cột Mã hàng / Tên hàng")
    lines = []
    for r in range(hr + 1, ws.max_row + 1):
        g = lambda k: ws.cell(r, by[k]).value if k in by else None  # noqa: E731
        ten, ma = str(g("ten_hang") or "").strip(), str(g("ma_hang") or "").strip()
        sl = _num(g("so_luong"))
        if not (ten or ma):
            continue
        tt = _num(g("thanh_tien")) or round(sl * _num(g("don_gia")))
        lines.append({"row": r, "ma": ma, "ten": ten, "dvt": str(g("dvt") or "").strip(), "sl": sl, "tt": tt,
                      "ck": _num(g("tien_ck")), "ngay": _date(g("ngay_hach_toan") or g("ngay_chung_tu") or g("ngay_hoa_don")),
                      "so_ct": str(g("so_chung_tu") or "").strip(), "so_hd": str(g("so_hoa_don") or "").strip()})
    return {"lines": lines, "hr": hr, "cols": by, "sheet": ws.title, "sheet_index": wb.worksheets.index(ws) + 1}


# ------------------------------------------------------------------ ghép mã hàng
_GENERIC = set("hang hoa san pham thuc pham dinh duong y hoc loai cai chiec hop goi bo the tui cho dung bang va voi "
               "cua cac nhung mot moi size model hieu nsx loai ltd co new".split())


def _distinct(words: set[str]) -> set[str]:
    """Từ phân biệt sản phẩm (tên thương hiệu, dòng sản phẩm...): chữ cái, dài >= 3, không phải từ chung chung."""
    return {w for w in words if len(w) >= 3 and w.isalpha() and w not in _GENERIC}


def _score(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    ta, tb = set(a.split()), set(b.split())
    jac = len(ta & tb) / max(1, len(ta | tb))
    small = min(len(ta), len(tb))
    cover = len(ta & tb) / max(1, small)      # tên ngắn nằm gọn trong tên dài
    ratio = difflib.SequenceMatcher(None, a, b).ratio()
    pre = 0.9 if (small >= 4 and (a.startswith(b) or b.startswith(a))) else 0.0   # tên ngắn chung chung không tính
    sc = max(ratio, jac, pre, cover * 0.92 if small >= 4 and cover >= 0.9 else 0.0)
    # 2 tên khác nhau ở từ phân biệt (vd "Ensure" vs "Glucerna"): khác sản phẩm dù chung phần mô tả
    if _distinct(ta - tb) and _distinct(tb - ta):
        sc *= 0.7
    # tên trên hóa đơn chứa mã model của danh mục (vd "TRM2974ZZTW1DKGR..." chứa "TRM2974"): cùng hàng
    mb, ma_ = _models(tb), _models(ta)
    if mb and any(any(t.startswith(m) for t in ta) for m in mb):
        sc = max(sc, 0.9)
    elif mb and ma_ and not (mb & ma_):
        # cả 2 có mã SKU / model nhưng khác nhau (BSRY65YE vàng vs BSRY65RE đỏ, AYPU011 vs AYPT315): khác hàng
        sc *= 0.6
    # số model / số hiệu của danh mục KHÔNG có trên hóa đơn (1000 vs 4000, Ryzonic 62 vs 65, N65 vs NO1): khác hàng
    if _keys(tb) - ta:
        sc *= 0.6
    return sc


def _models(words: set[str]) -> set[str]:
    """Mã model dạng chữ + số, dài >= 5 (TRM2974, BSRY62PI, AYPR285...)."""
    return {w for w in words if len(w) >= 5 and any(c.isdigit() for c in w) and any(c.isalpha() for c in w)
            and not _UNIT.match(w)}


# số + đơn vị đo (200ml, 16mm, 850g, 200m, 4u5): không phải mã model
_UNIT = re.compile(r"^\d+(?:[.,]\d+)?(?:ml|mm|cm|m|g|gr|kg|l|lit|u|u\d|oz|inch|cai|chiec|hop|chai|cuon|vien|goi)$")


def _keys(words: set[str]) -> set[str]:
    """Số hiệu ngắn phân biệt model: số >= 2 chữ số (1000, 62) hoặc chữ+số ngắn (N65, A880, 3S, 5U)."""
    return {w for w in words if any(c.isdigit() for c in w)
            and ((w.isdigit() and len(w) >= 2) or (not w.isdigit() and 2 <= len(w) <= 4))}


AUTO_OK = 0.78      # ghép tự động
MAYBE = 0.55        # ghép nhưng cần kiểm tra


def new_code(ten: str, used: set[str]) -> str:
    """Mã dễ đọc: chữ cái đầu các từ + số hiệu model, vd "Áo cầu lông Taro TR025-A29" -> ACLT_TR025_A29."""
    words = [w for w in _norm(ten).upper().split() if w]
    model = [w for w in words if any(c.isdigit() for c in w)][:2]
    head = "".join(w[0] for w in words if w not in model and not w.isdigit())[:5]
    base = "_".join(x for x in [head] + model if x) or "HH"
    base = re.sub(r"[^A-Z0-9_]", "", base)[:20] or "HH"
    code, i = base, 1
    while code.upper() in used:
        i += 1
        code = f"{base}_{i}"
    used.add(code.upper())
    return code


def match_codes(items: list[dict[str, Any]], groups: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    """Gắn mã hàng cho mọi dòng (sửa trực tiếp line["ma_ghep"]). Trả về bảng ghép theo từng tên hàng."""
    base = list(items)
    by_code = {i["ma"].upper(): i for i in items}
    by_name = {_norm(i["ten"]): i for i in items}
    used = set(by_code)
    memo: dict[str, dict[str, Any]] = {}
    report = []
    for kind, lines in groups.items():
        for ln in lines:
            if ln["ma"] and ln["ma"].upper() in by_code:
                ln["ma_ghep"], ln["cach"] = by_code[ln["ma"].upper()]["ma"], "Mã có sẵn trong file"
                continue
            key = _norm(ln["ten"])
            if key in memo:
                m = memo[key]
            else:
                it = by_name.get(key)
                if it:
                    m = {"ma": it["ma"], "diem": 1.0, "cach": "Trùng tên trong danh mục"}
                else:
                    best, sc = None, 0.0
                    # tên gần đúng: so với danh mục đầu kỳ + hàng MỚI đã gặp ở mua vào (bán ra hàng vừa mua trong kỳ)
                    for i in (base if kind == "Mua vào" else items):
                        s = _score(key, _norm(i["ten"]))
                        if ln["dvt"] and i["dvt"] and _norm(ln["dvt"]) != _norm(i["dvt"]):
                            s -= 0.08
                        if s > sc:
                            best, sc = i, s
                    if best and sc >= AUTO_OK:
                        m = {"ma": best["ma"], "diem": round(sc, 2), "cach": "Tên gần đúng"}
                    elif best and sc >= MAYBE:
                        m = {"ma": best["ma"], "diem": round(sc, 2), "cach": "Tên gần đúng – CẦN KIỂM TRA"}
                    else:
                        code = ln["ma"] or new_code(ln["ten"], used)
                        used.add(code.upper())
                        it = {"kho": DEFAULT_KHO, "ma_kho": "", "ma": code, "ten": ln["ten"], "dvt": ln["dvt"],
                              "sl": 0.0, "gt": 0.0, "moi": True}
                        items.append(it)
                        by_code[code.upper()] = it
                        by_name[key] = it
                        m = {"ma": code, "diem": 0.0, "cach": "Hàng MỚI – chưa có trong danh mục" +
                             (", dùng mã trên hóa đơn" if ln["ma"] else ", đã tạo mã")}
                memo[key] = m
                report.append({"loai": kind, "ten": ln["ten"], "dvt": ln["dvt"], "ma": m["ma"], "diem": m["diem"],
                               "cach": m["cach"], "ten_dm": by_code[m["ma"].upper()]["ten"]})
            ln["ma_ghep"], ln["cach"] = m["ma"], m["cach"]
    return report


# ------------------------------------------------------------------ tính tồn kho
def compute(items: list[dict[str, Any]], buy: list[dict[str, Any]], sell: list[dict[str, Any]],
            ret: list[dict[str, Any]]) -> list[dict[str, Any]]:
    # 1 mã hàng có thể nằm ở nhiều kho (vd vừa "Hàng hóa" vừa "Công cụ dụng cụ"): mỗi (kho, mã) 1 dòng như MISA;
    # phát sinh mua / bán ghi vào kho chính của mã đó (kho đang tồn nhiều nhất)
    rows = [{**i, "sl_dau": i["sl"], "gt_dau": i["gt"], "sl_nhap": 0.0, "gt_nhap": 0.0, "sl_xuat": 0.0, "gt_xuat": 0.0,
             "sl_tra": 0.0} for i in items]
    main: dict[str, dict[str, Any]] = {}
    for r in rows:
        k = r["ma"].upper()
        if k not in main or r["sl_dau"] > main[k]["sl_dau"]:
            main[k] = r
    for ln in buy:
        if ln["sl"] > 0:                       # dòng dịch vụ / ghi chú không có số lượng: không qua kho
            r = main[ln["ma_ghep"].upper()]
            r["sl_nhap"] += ln["sl"]
            r["gt_nhap"] += ln["tt"] - ln["ck"]
    for ln in sell:
        if ln["sl"] > 0:
            main[ln["ma_ghep"].upper()]["sl_xuat"] += ln["sl"]
    for ln in ret:
        if ln["sl"] > 0:
            main[ln["ma_ghep"].upper()]["sl_tra"] += ln["sl"]
    out = []
    for r in rows:
        sl_co, gt_co = r["sl_dau"] + r["sl_nhap"], r["gt_dau"] + r["gt_nhap"]
        if sl_co > 0:
            dg = gt_co / sl_co
        elif r["sl_dau"]:
            dg = r["gt_dau"] / r["sl_dau"]
        else:
            dg = 0.0
        r["dg_bq"] = dg
        gt_tra = round(r["sl_tra"] * dg)
        r["sl_nhap"] += r["sl_tra"]
        r["gt_nhap"] = round(r["gt_nhap"] + gt_tra)
        r["gt_xuat"] = round(r["sl_xuat"] * dg)
        r["sl_cuoi"] = r["sl_dau"] + r["sl_nhap"] - r["sl_xuat"]
        r["gt_cuoi"] = round(r["gt_dau"] + r["gt_nhap"] - r["gt_xuat"])
        if abs(r["sl_cuoi"]) < 1e-9 and abs(r["gt_cuoi"]) <= 2:   # hết hàng: dồn chênh lệch làm tròn vào giá trị xuất
            r["gt_xuat"] += r["gt_cuoi"]
            r["gt_cuoi"] = 0
        out.append(r)
    order = {k: i for i, k in enumerate(dict.fromkeys(_norm(x["kho"]) for x in items))}   # thứ tự kho như file đầu kỳ
    out.sort(key=lambda x: (order.get(_norm(x["kho"]), 99), x["ma"].upper()))
    return out


# ------------------------------------------------------------------ ghi báo cáo theo mẫu Tổng hợp tồn kho
def period_text(d1: date | None, d2: date | None) -> str:
    if not d1 or not d2:
        return ""
    if (d1.year, d1.month) == (d2.year, d2.month):
        return f"Tháng {d2.month} năm {d2.year}"
    if d1.year == d2.year and d1.month in (1, 4, 7, 10) and d2.month == d1.month + 2:
        return f"Quý {(d1.month - 1) // 3 + 1} năm {d2.year}"
    if d1.year == d2.year and (d1.month, d2.month) == (1, 12):
        return f"Năm {d2.year}"
    return f"Từ ngày {d1:%d/%m/%Y} đến ngày {d2:%d/%m/%Y}"


def write_report(tpl: bytes, rows: list[dict[str, Any]], ky_text: str, ngay_ky: date | None) -> bytes:
    """Ghi vào CHÍNH file mẫu Tổng hợp tồn kho (.xlsx của MISA): giữ phần đầu (tên đơn vị, tiêu đề), dòng tiêu đề
    cột, định dạng ô, chữ ký; thay dữ liệu, dòng Tổng cộng, ngày ký. Tiêu đề cột in lặp lại ở mỗi trang."""
    from openpyxl import load_workbook
    from openpyxl.worksheet.pagebreak import RowBreak

    if tpl[:2] != b"PK":
        from .misa_import import load_any_workbook
        buf = io.BytesIO()
        load_any_workbook(tpl).save(buf)
        tpl = buf.getvalue()
    wb = load_workbook(io.BytesIO(tpl))
    ws = wb.worksheets[0]
    # dòng tiêu đề cột (Tên kho / Mã hàng ...) + dòng Số lượng/Giá trị ngay dưới
    hr = next(r for r in range(1, 40) if _norm(ws.cell(r, 2).value) == "ma hang" or _norm(ws.cell(r, 1).value) == "ten kho")
    first = hr + 2
    # dòng mẫu: dòng dữ liệu đầu, dòng Tổng cộng, khối ngày ký + chữ ký
    tot_r = next((r for r in range(first, ws.max_row + 1) if _norm(ws.cell(r, 1).value).startswith("tong cong")), None)
    sign_r = next((r for r in range(first, ws.max_row + 1)
                   if any(_norm(ws.cell(r, c).value).startswith("nguoi lap bieu") for c in range(1, ws.max_column + 1))), None)
    date_r = sign_r - 1 if sign_r else None
    while date_r and date_r > (tot_r or first) and not any(
            _norm(ws.cell(date_r, c).value).startswith("ngay") for c in range(1, ws.max_column + 1)):
        date_r -= 1
    ncol = max(ws.max_column, 16)

    def snap(r: int | None, nrows: int = 1) -> list[dict[str, Any]]:
        out = []
        for rr in range(r, r + nrows) if r else []:
            out.append({"h": ws.row_dimensions[rr].height,
                        "cells": [(copy(ws.cell(rr, c)._style), ws.cell(rr, c).value, ws.cell(rr, c).number_format)
                                  for c in range(1, ncol + 1)]})
        return out

    # dòng mẫu: dòng dữ liệu có số lượng KHÔNG âm (dòng âm của MISA tô chữ đỏ)
    sample = first
    for rr in range(first, (tot_r or ws.max_row)):
        v = ws.cell(rr, 7).value
        if isinstance(v, (int, float)) and v >= 0 and ws.cell(rr, 2).value:
            sample = rr
            break
    data_style = snap(sample)[0]
    tot_style = snap(tot_r)[0] if tot_r else data_style
    tail = snap(date_r, (sign_r + 2 - date_r)) if sign_r and date_r else []
    data_merges = [(m.min_col, m.max_col) for m in ws.merged_cells.ranges if m.min_row == m.max_row == sample]
    tot_merges = [(m.min_col, m.max_col) for m in ws.merged_cells.ranges if tot_r and m.min_row == m.max_row == tot_r]
    tail_merges = [(m.min_row - date_r, m.min_col, m.max_row - date_r, m.max_col) for m in ws.merged_cells.ranges
                   if date_r and m.min_row >= date_r and m.max_row <= sign_r + 1]
    # cột dữ liệu theo dòng tiêu đề của mẫu
    heads = {c: _norm(ws.cell(hr, c).value) for c in range(1, ncol + 1)}
    subs = {c: _norm(ws.cell(hr + 1, c).value) for c in range(1, ncol + 1)}
    grp = sorted((c, h) for c, h in heads.items() if h in ("dau ky", "nhap kho", "xuat kho", "cuoi ky"))
    cols: dict[str, int] = {}
    for i, (c, h) in enumerate(grp):
        end = grp[i + 1][0] if i + 1 < len(grp) else ncol + 1
        key = {"dau ky": "dau", "nhap kho": "nhap", "xuat kho": "xuat", "cuoi ky": "cuoi"}[h]
        for x in range(c, end):
            if subs.get(x) == "so luong":
                cols[f"sl_{key}"] = x
            elif subs.get(x) == "gia tri":
                cols[f"gt_{key}"] = x
    for c, h in heads.items():
        if h.startswith("ten kho"):
            cols["kho"] = c
        elif h.startswith("ma hang"):
            cols["ma"] = c
        elif h.startswith("ten hang"):
            cols["ten"] = c
        elif h in ("dvt", "don vi tinh"):
            cols["dvt"] = c
    # xóa toàn bộ phần dưới tiêu đề (dữ liệu cũ, tiêu đề lặp lại của trang 2, chân trang, chữ ký)
    for m in [m for m in ws.merged_cells.ranges if m.min_row >= first]:
        ws.unmerge_cells(str(m))
    ws.delete_rows(first, ws.max_row - first + 1)
    ws.row_breaks = type(ws.row_breaks)()

    def put_row(r: int, style: dict[str, Any], merges: list[tuple[int, int]]) -> None:
        if style["h"]:
            ws.row_dimensions[r].height = style["h"]
        for c, (st, _v, nf) in enumerate(style["cells"], 1):
            cell = ws.cell(r, c)
            cell._style = copy(st)
            cell.number_format = nf
        for a, b in merges:
            ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)

    r = first
    tot = {k: 0.0 for k in ("sl_dau", "gt_dau", "sl_nhap", "gt_nhap", "sl_xuat", "gt_xuat", "sl_cuoi", "gt_cuoi")}
    for x in rows:
        put_row(r, data_style, data_merges)
        for k, c in cols.items():
            v = x.get(k)
            if k.startswith("sl_"):
                v = round(v, 4)
                v = int(v) if float(v).is_integer() else v
            cell = ws.cell(r, c)
            cell.value = v
            if isinstance(v, (int, float)) and v < 0:      # số âm chữ đỏ như MISA
                f = copy(cell.font)
                f.color = "FFFF0000"
                cell.font = f
        for k in tot:
            tot[k] += x.get(k) or 0
        # chiều cao dòng theo độ dài tên hàng (cột tên hàng ~36 ký tự / dòng như mẫu MISA)
        ln = len(str(x.get("ten") or ""))
        ws.row_dimensions[r].height = 15.75 * max(1, -(-ln // 36))
        r += 1
    # Tổng cộng
    put_row(r, tot_style, tot_merges)
    ws.cell(r, 1).value = "Tổng cộng"
    for k, c in cols.items():
        if k in tot:
            v = round(tot[k], 4)
            ws.cell(r, c).value = int(v) if float(v).is_integer() else v
    r += 2
    # ngày ký + chữ ký
    for i, st in enumerate(tail):
        put_row(r + i, st, [])
        for c, (_s, val, _nf) in enumerate(st["cells"], 1):
            if val is not None:
                ws.cell(r + i, c).value = val
    for a_r, a_c, b_r, b_c in tail_merges:
        ws.merge_cells(start_row=r + a_r, start_column=a_c, end_row=r + b_r, end_column=b_c)
    if tail and ngay_ky:
        for c in range(1, ncol + 1):
            if _norm(ws.cell(r, c).value).startswith("ngay"):
                ws.cell(r, c).value = f"Ngày {ngay_ky.day:02d} tháng {ngay_ky.month:02d} năm {ngay_ky.year}"
    # kỳ báo cáo ("Tháng 11 năm 2026")
    if ky_text:
        for rr in range(1, hr):
            v = _norm(ws.cell(rr, 1).value)
            if re.match(r"^(thang|quy|nam|tu ngay) ", v):
                ws.cell(rr, 1).value = ky_text
                break
    ws.print_title_rows = f"{hr}:{hr + 1}"     # tiêu đề cột in lặp lại mỗi trang như MISA
    from openpyxl.utils import get_column_letter
    last_row = r + max(0, len(tail) - 1)
    ws.print_area = f"A1:{get_column_letter(ncol)}{last_row}"   # vùng in đúng tới dòng chữ ký
    _ = RowBreak
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ------------------------------------------------------------------ ghi mã hàng vào file mua vào / bán ra
def write_codes(data: bytes, name: str, parsed: dict[str, Any], lines: list[dict[str, Any]]) -> tuple[bytes, str]:
    """Điền cột Mã hàng (*) của file chứng từ MISA theo mã đã ghép. Giữ nguyên file (xls dùng Excel trên máy)."""
    from . import fill_template as ft
    c_ma = parsed["cols"].get("ma_hang")
    if not c_ma:
        return data, Path(name).suffix or ".xlsx"
    ext = Path(name).suffix.lower() or ".xlsx"
    if data[:2] == b"PK":
        from openpyxl import load_workbook
        keep_vba = ext in (".xlsm", ".xltm")
        wb = load_workbook(io.BytesIO(data), keep_vba=keep_vba)
        ws = wb.worksheets[parsed["sheet_index"] - 1]
        for ln in lines:
            ws.cell(ln["row"], c_ma).value = ln["ma_ghep"]
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue(), ext
    if ft.excel_available():
        first = min((ln["row"] for ln in lines), default=parsed["hr"] + 1)
        last = max((ln["row"] for ln in lines), default=first)
        by_row = {ln["row"]: ln["ma_ghep"] for ln in lines}
        rows = [{"ma_hang": by_row.get(rr)} for rr in range(first, last + 1)]
        return ft.fill_with_excel(data, ext, parsed["sheet_index"], first - 1, {c_ma: "ma_hang"}, rows, clear=False), ext
    raise ValueError("File .xls cần Microsoft Excel trên máy để điền mã hàng – hãy lưu file thành .xlsx")


def write_catalog(tpl: bytes | None, items: list[dict[str, Any]]) -> tuple[bytes, str]:
    """File Danh mục vật tư hàng hóa (mẫu MISA) cho các mã hàng MỚI để khai vào MISA trước khi nhập chứng từ."""
    new = [i for i in items if i.get("moi")]
    rows = [{"ma_hang": i["ma"], "ten_hang": i["ten"], "tinh_chat": "Hàng hóa", "dvt": i["dvt"],
             "nhom_nn": "101 - Hoạt động bán buôn, bán lẻ các loại hàng hóa (trừ giá trị hàng hóa đại lý bán đúng giá "
                        "hưởng hoa hồng);"} for i in new]
    from openpyxl import Workbook
    if tpl:
        from . import fill_template as ft
        from .misa_import import load_any_workbook
        wb = load_any_workbook(tpl)
        ws = wb.worksheets[0]
        hr = next((r for r in range(1, 30) if _norm(ws.cell(r, 1).value).startswith("ma hang")), None)
        if hr:
            heads = {c: _norm(ws.cell(hr, c).value) for c in range(1, ws.max_column + 1)}
            cmap = {}
            for c, h in heads.items():
                if h.startswith("ma hang") and "ma_hang" not in cmap.values():
                    cmap[c] = "ma_hang"
                elif h.startswith("ten hang") and "ten_hang" not in cmap.values():
                    cmap[c] = "ten_hang"
                elif h == "tinh chat":
                    cmap[c] = "tinh_chat"
                elif h.startswith("don vi tinh chinh"):
                    cmap[c] = "dvt"
                elif h.startswith("ma nhom nganh nghe"):
                    cmap[c] = "nhom_nn"
            if ft.kind_of(tpl) == "ooxml":
                return ft.fill_ooxml(tpl, ws.title, hr, cmap, rows), ".xlsx"
            if ft.excel_available():
                return ft.fill_with_excel(tpl, ".xls", 1, hr, cmap, rows), ".xls"
    wb = Workbook()
    ws = wb.active
    ws.title = "Danh muc VTHH moi"
    for i, h in enumerate(["Mã hàng (*)", "Tên hàng (*)", "Tính chất", "Đơn vị tính chính", "Mã nhóm ngành nghề (*)"], 1):
        ws.cell(1, i, h)
    for r, x in enumerate(rows, 2):
        for i, k in enumerate(["ma_hang", "ten_hang", "tinh_chat", "dvt", "nhom_nn"], 1):
            ws.cell(r, i, x[k])
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), ".xlsx"


def write_match_report(report: list[dict[str, Any]], rows: list[dict[str, Any]]) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "Ghep ma hang"
    hd = ["Loại", "Tên hàng trên hóa đơn", "ĐVT", "Mã hàng ghép", "Tên hàng trong danh mục", "Độ khớp", "Cách ghép"]
    for i, h in enumerate(hd, 1):
        c = ws.cell(1, i, h)
        c.font, c.fill = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
    warn = PatternFill("solid", fgColor="FFF2CC")
    for r, x in enumerate(report, 2):
        for i, v in enumerate([x["loai"], x["ten"], x["dvt"], x["ma"], x["ten_dm"], x["diem"], x["cach"]], 1):
            c = ws.cell(r, i, v)
            if "KIỂM TRA" in x["cach"] or "MỚI" in x["cach"]:
                c.fill = warn
    for col, w in zip("ABCDEFG", (10, 60, 8, 16, 60, 9, 40)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    w2 = wb.create_sheet("Don gia binh quan")
    for i, h in enumerate(["Mã hàng", "Tên hàng", "ĐVT", "SL đầu kỳ + mua vào", "Đơn giá bình quân", "SL xuất",
                           "Giá trị xuất", "SL cuối kỳ"], 1):
        w2.cell(1, i, h).font = Font(bold=True)
    for r, x in enumerate(rows, 2):
        for i, v in enumerate([x["ma"], x["ten"], x["dvt"], x["sl_dau"] + x["sl_nhap"] - x["sl_tra"], round(x["dg_bq"], 2),
                               x["sl_xuat"], x["gt_xuat"], x["sl_cuoi"]], 1):
            w2.cell(r, i, v)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ------------------------------------------------------------------ chạy cả quy trình
@dataclass
class Job:
    id: str
    status: str = "running"
    log: list[str] = field(default_factory=list)
    files: dict[str, str] = field(default_factory=dict)
    stat: dict[str, Any] = field(default_factory=dict)
    report: list[dict[str, Any]] = field(default_factory=list)

    def say(self, m: str) -> None:
        self.log.append(f"{datetime.now():%H:%M:%S} {m}")

    def public(self) -> dict[str, Any]:
        return {"id": self.id, "status": self.status, "log": self.log[-80:], "files": self.files, "stat": self.stat,
                "report": self.report[:400]}


JOBS: dict[str, Job] = {}


def run(job: Job, inp: dict[str, Any]) -> None:
    """inp: opening(name,bytes), tpl(name,bytes)|None, buy/sell/ret: {"file": (name, bytes)|None, "invoices": [...],
    "mau": (name, bytes)|None}, catalog_tpl, mst, ky_text, ngay"""
    from . import misa_import as mi
    try:
        out_dir = Path(settings.output_dir) / "TonKho" / datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir.mkdir(parents=True, exist_ok=True)
        op_name, op_data = inp["opening"]
        opening = read_opening(op_data)
        items = opening["items"]
        kieu = {"tonkho": "báo cáo Tổng hợp tồn kho – lấy số CUỐI KỲ", "danhmuc": "danh mục VTHH – SL đầu kỳ 0",
                "bangke": "bảng kê hàng tồn kho – tên hàng, ĐVT, số lượng, giá trị"}[opening["loai"]]
        job.say(f"Tồn kho đầu kỳ: {len(items)} mã hàng ({kieu}); tổng SL {sum(i['sl'] for i in items):,.0f}, "
                f"giá trị {sum(i['gt'] for i in items):,.0f} đ".replace(",", "."))
        parsed: dict[str, Any] = {}
        sources: dict[str, tuple[str, bytes]] = {}
        labels = {"buy": "Mua vào", "sell": "Bán ra", "ret": "Hàng bán trả lại"}
        for key, direction in (("buy", "purchase"), ("sell", "sold"), ("ret", "sold")):
            part = inp.get(key) or {}
            src = part.get("file")
            if part.get("invoices"):
                tpl_name, tpl = part.get("mau") or src or (None, b"")
                if not tpl:
                    tpl = mi.basic_template(direction)
                    tpl_name = f"MISA_{key}.xlsx"
                job.say(f"{labels[key]}: chuyển {len(part['invoices'])} hóa đơn vào mẫu MISA {tpl_name}…")
                mj = mi.Job(id=uuid.uuid4().hex)
                mi.run(mj, part["invoices"], tpl, direction, inp.get("mst", ""), inp.get("acc") or {}, tpl_name)
                if mj.status != "done":
                    raise RuntimeError(f"Chuyển hóa đơn {labels[key].lower()} lỗi: " + " | ".join(mj.log[-3:]))
                src = (Path(mj.file).name, Path(mj.file).read_bytes())
                job.say(f"{labels[key]}: đã điền {mj.done} hóa đơn vào mẫu ({Path(mj.file).name})")
            if src:
                sources[key] = src
                parsed[key] = read_lines(src[1])
                job.say(f"{labels[key]}: {len(parsed[key]['lines'])} dòng hàng")
        if "buy" not in parsed and "sell" not in parsed:
            raise ValueError("Cần ít nhất file mua vào hoặc bán ra (hoặc hóa đơn để chuyển vào mẫu)")
        groups = {labels[k]: parsed[k]["lines"] for k in ("buy", "sell", "ret") if k in parsed}
        report = match_codes(items, groups)
        rows = compute(items, parsed.get("buy", {}).get("lines", []), parsed.get("sell", {}).get("lines", []),
                       parsed.get("ret", {}).get("lines", []))
        dates = [ln["ngay"] for k in parsed for ln in parsed[k]["lines"] if ln["ngay"]]
        d1, d2 = (min(dates), max(dates)) if dates else (None, None)
        if d2:
            from calendar import monthrange
            d1 = d1.replace(day=1)
            d2 = d2.replace(day=monthrange(d2.year, d2.month)[1])
        ky_text = inp.get("ky_text") or period_text(d1, d2)
        ngay = _date(inp.get("ngay")) or d2
        new = sum(1 for i in items if i.get("moi"))
        chk = sum(1 for x in report if "KIỂM TRA" in x["cach"])
        job.stat = {"ma_hang": len(rows), "moi": new, "can_kiem_tra": chk, "ky": ky_text,
                    "dong_mua": len(parsed.get("buy", {}).get("lines", [])), "dong_ban": len(parsed.get("sell", {}).get("lines", [])),
                    "am": sum(1 for r in rows if r["sl_cuoi"] < 0)}
        job.report = report
        job.say(f"Ghép mã: {len(report)} tên hàng – {new} mã MỚI, {chk} cần kiểm tra")
        # 1) Tổng hợp tồn kho theo mẫu
        tpl_name, tpl = inp.get("tpl") or (op_name, op_data)
        f1 = out_dir / f"TONG_HOP_TON_KHO_{re.sub(r'[^0-9A-Za-z]+', '_', fold(ky_text))}.xlsx"
        f1.write_bytes(write_report(tpl, rows, ky_text, ngay))
        job.files["tonkho"] = str(f1)
        job.say(f"Đã lập Tổng hợp tồn kho {ky_text}: {len(rows)} mã hàng" +
                (f" – {job.stat['am']} dòng tồn cuối kỳ ÂM (kiểm tra lại số đầu kỳ / chứng từ bán)" if job.stat["am"] else ""))
        # 2) file mua vào / bán ra đã điền mã hàng
        for key in ("buy", "sell", "ret"):
            if key in sources:
                name, data = sources[key]
                try:
                    b, ext = write_codes(data, name, parsed[key], parsed[key]["lines"])
                    f = out_dir / f"{Path(name).stem}_co_ma_hang{ext}"
                    f.write_bytes(b)
                    job.files[key] = str(f)
                except Exception as e:  # noqa: BLE001
                    job.say(f"{labels[key]}: không điền được mã hàng vào file ({e})")
        # 3) danh mục hàng mới + bảng ghép mã
        if new:
            b, ext = write_catalog((inp.get("catalog_tpl") or (None, None))[1], items)
            f = out_dir / f"Danh_muc_VTHH_moi{ext}"
            f.write_bytes(b)
            job.files["danhmuc"] = str(f)
        f = out_dir / "Ghep_ma_hang.xlsx"
        f.write_bytes(write_match_report(report, rows))
        job.files["ghepma"] = str(f)
        job.files["folder"] = str(out_dir)
        job.status = "done"
        job.say("Xong.")
    except Exception as e:  # noqa: BLE001
        job.status = "error"
        job.say(f"LỖI: {e}")


def start(inp: dict[str, Any]) -> Job:
    import threading
    job = Job(id=uuid.uuid4().hex[:12])
    JOBS[job.id] = job
    threading.Thread(target=run, args=(job, inp), daemon=True, name="inventory").start()
    return job
