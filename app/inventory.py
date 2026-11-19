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
    t = re.sub(r"(?<=[a-z])\.(?=\d)", "", t)          # "S.70" = "S70", "N.38" = "N38"
    t = re.sub(r"[^a-z0-9]+", " ", t)
    t = re.sub(r"\b([a-z]{5,})(\d{1,3})\b", r"\1 \2", t)     # "ARCSABER11" = "Arcsaber 11", "Calibar200"
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
    # file mẫu MISA còn nguyên dòng VÍ DỤ minh họa (ngày 31/01/2020, BPHP "Bàn phím máy tính"...): không phải dữ liệu
    is_tpl = any("file mau" in _norm(ws.cell(r, 1).value) for r in range(1, min(hr, 6)))
    skipped = 0
    for r in range(hr + 1, ws.max_row + 1):
        g = lambda k: ws.cell(r, by[k]).value if k in by else None  # noqa: E731
        ten, ma = str(g("ten_hang") or "").strip(), str(g("ma_hang") or "").strip()
        sl = _num(g("so_luong"))
        if not (ten or ma):
            continue
        ngay = _date(g("ngay_hach_toan") or g("ngay_chung_tu") or g("ngay_hoa_don"))
        if is_tpl and ngay == date(2020, 1, 31):
            skipped += 1
            continue
        tt = _num(g("thanh_tien")) or round(sl * _num(g("don_gia")))
        lines.append({"row": r, "ma": ma, "ten": ten, "dvt": str(g("dvt") or "").strip(), "sl": sl, "tt": tt,
                      "ck": _num(g("tien_ck")), "ngay": _date(g("ngay_hach_toan") or g("ngay_chung_tu") or g("ngay_hoa_don")),
                      "so_ct": str(g("so_chung_tu") or "").strip(), "so_hd": str(g("so_hoa_don") or "").strip()})
    return {"lines": lines, "hr": hr, "cols": by, "sheet": ws.title, "sheet_index": wb.worksheets.index(ws) + 1,
            "vi_du": skipped}


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


# Từ chỉ LOẠI hàng / màu / size / chữ chung: bỏ khi so từ khóa nhận dạng (2 nơi đặt tên khác nhau: "Cước cầu lông Gosen
# Ryzonic 69- Trắng" và "Dây cước căng vợt GOSEN Ryzonic 69" cùng là 1 hàng: từ khóa {gosen, ryzonic, 69})
# (giữ lại từ PHÂN BIỆT sản phẩm: giày, áo, quần, túi, balo, bình, xịt, bóng, hộp, cuộn... vd "xịt vệ sinh GIÀY" khác
# "xịt vệ sinh mặt vợt", cước "Ryzonic 65 CUỘN" khác cước lẻ)
_KW_STOP = set("""day cuoc cang vot cau long khung bo
cai chiec doi soi cay hang hoa san pham dung cho the thao chinh hang ma made in japan china viet nam mau size sz loai moi
nam nu tre em unisex co tron ngan dai chai goi ong qua pickleball tennis paddle paddles
trang den do xanh la cam hong vang tim xam bac nau be kem ngoc trai dat duong navy blue black white red pink orange
yellow green purple gray grey silver gold pastel light dark cyan mint chim tuyen dao troi thep
s m l xl xxl xxxl 2xl 3xl 4xl 5xl 3u 4u 5u 6u 4u5 3u5 5u5 3ug5 4ug5 5ug5 cl va voi""".split())


def _kw(name: str) -> frozenset[str]:
    """Từ khóa nhận dạng hàng: hãng, dòng, số model (bỏ từ loại hàng, màu, size, mã SKU dài, số size giày 34-46)."""
    out = set()
    for w in _norm(_NCC_PREFIX.sub("", str(name or ""))).split():
        if w in _KW_STOP or _UNIT.match(w):
            continue
        if w.isdigit() and 34 <= int(w) <= 46 and len(w) == 2:   # size giày / áo
            continue
        if len(w) >= 8 and any(c.isdigit() for c in w):             # mã SKU dài (BSRY69WH, BFRAX100ZVAZ...)
            continue
        out.add(w)
    return frozenset(out)


def _pick(name: str, dvt: str, cands: list[dict[str, Any]]) -> dict[str, Any]:
    """Chọn hàng gần nhất: điểm tên -> cùng ĐVT (bán "BG 65" theo Cái thì không lấy "BG 65 cuộn") -> tên sát hơn."""
    k, d = _norm(name), _norm(dvt)
    return max(cands, key=lambda it: (round(_score(k, _norm(it["ten"])), 2), bool(d) and _norm(it.get("dvt")) == d,
                                      difflib.SequenceMatcher(None, k, _norm(it["ten"])).ratio()))


def _kw_match(inv_name: str, base: list[dict[str, Any]], base_kw: list[frozenset[str]],
              dvt: str = "") -> tuple[dict | None, float, str]:
    """Hàng danh mục có ĐỦ từ khóa nhận dạng nằm trong tên hóa đơn (chọn hàng có nhiều từ khóa nhất – cụ thể nhất).
    Thiếu đúng 1 từ chữ (thường là tên hãng: "Lining Halbertec 4000" vs "Halbertec 4000") -> cần kiểm tra."""
    ki = _kw(inv_name)
    if not ki:
        return None, 0.0, ""
    full, near = [], []
    for it, kc in zip(base, base_kw):
        # cần ít nhất 1 từ nhận dạng thật (số hiệu model / tên dòng >= 4 chữ), không tính tên hãng ("đan … Yonex"
        # không phải "dán … Yonex")
        if len(kc) < 2 or not any((any(c.isdigit() for c in t) or len(t) >= 4) and t not in _FAM_WEAK for t in kc):
            continue
        miss = kc - ki
        if not miss:
            full.append((len(kc), it))
        elif len(miss) == 1 and len(kc) >= 3 and not any(c.isdigit() for c in next(iter(miss))):
            near.append((len(kc), it))      # thiếu 1 từ chữ (thường là tên hãng), còn >= 2 từ khóa khớp
    for cand, sc, why in ((full, 0.93, "Khớp từ khóa tên chuẩn (hãng, dòng, model)"),
                          (near, 0.8, "Khớp từ khóa tên chuẩn (thiếu 1 từ) – CẦN KIỂM TRA")):
        if cand:
            top = max(n for n, _ in cand)
            best = [it for n, it in cand if n == top]
            if len(best) == 1:
                return best[0], sc, why
            # nhiều hàng cùng mức: chọn hàng có tên gần nhất, cùng ĐVT
            return _pick(inv_name, dvt, best), sc - 0.05, why
    return None, 0.0, ""


_NCC_PREFIX = re.compile(r"^\s*p\s*-\s*[a-z0-9]+\s*-\s*\d+\s*-\s*", re.I)   # mã NCC đầu tên: "P-AYPT395-4-"


def _fam_kw(name: str) -> frozenset[str]:
    return _kw(name)


def _sku_models(name: str) -> set[str]:
    """Mã model đầu SKU dài: "BAG524B0112ZJTBKZZ" -> bag524, "TRM2974ZZTW1DKGR" -> trm2974."""
    out = set()
    for w in _norm(name).split():
        m = re.match(r"[a-z]{2,4}\d{3,}", w) if len(w) >= 8 else None
        if m:
            out.add(m.group(0))
    return out


def _num_toks(kw: frozenset[str]) -> set[str]:
    """Số hiệu model (>= 2 chữ số): 300, 300b, 6000d, 72, s70, trm2988 (bỏ cc8, 4u5... chỉ 1 chữ số)."""
    return {w for w in kw if sum(c.isdigit() for c in w) >= 2}


def _same_model(t: str, x: str) -> bool:
    """Số hiệu danh mục t ~ số hiệu hóa đơn x: 300 ~ 300b, 6000 ~ 6000d, trm2988 ~ 2988, trm2988 ~ trm2988zzcc8."""
    if t == x:
        return True
    if x.startswith(t) and x[len(t):].isalpha() and len(x) - len(t) <= 2:
        return True
    if len(t) >= 5 and t[0].isalpha() and any(c.isdigit() for c in t) and x.startswith(t):
        return True                                    # bag524 ~ bag524b0612z (cùng model, khác size / màu)
    d = re.sub(r"^[a-z]{1,4}", "", t)
    return len(d) >= 4 and d.isdigit() and (x == d or (t != d and x.startswith(t)))


def _family(inv_name: str, it_name: str, inv_kw: frozenset[str] | None = None) -> bool:
    """CÙNG DÒNG / MODEL, chỉ khác phiên bản – màu – size (Calibar 300B ~ Calibar 300, Aeronaut 6000 Max ~ Aeronaut
    6000 Speed, BG 65 cuộn ~ BG 65 Titanium, Yonex-2988-CC8 ~ Áo Yonex TRM2988): mọi số hiệu model 2 bên trùng nhau,
    cùng loại hàng, chung tên dòng / hãng. Khác số hiệu (Halbertec 6000 vs 1000, Axforce 9 vs 90) thì KHÔNG."""
    ki = inv_kw if inv_kw is not None else _fam_kw(inv_name)
    kc = _fam_kw(it_name)
    nc = _num_toks(kc) | _sku_models(it_name)
    if not nc or _cat(inv_name) != _cat(it_name):
        return False
    raw = set(_norm(_NCC_PREFIX.sub("", inv_name)).split())          # cả mã SKU dài (trm2998zzcc8...)
    if not all(any(_same_model(t, x) for x in ki | raw) for t in nc):
        return False
    if not all(any(_same_model(t, x) for t in nc) for x in _num_toks(ki)):
        return False                                                   # hóa đơn có số hiệu khác không có trong danh mục
    if any(len(t) >= 5 and sum(c.isdigit() for c in t) >= 3 and any(c.isalpha() for c in t) for t in nc):
        return True                                                    # mã model chữ + số (TRM2988, AYZU001) đủ nhận dạng
    shared = {w for w in ki & kc if w.isalpha()}
    if any(t.isdigit() and len(t) >= 4 for t in nc) and shared and \
            not {w for w in kc - ki if w.isalpha() and w not in _FAM_WEAK}:
        return True                                    # áo Yonex-2998 size L ~ áo Yonex-2998 TRM2998ZZCC8 size XS
    # phải chung TÊN DÒNG hàng (Aeronaut, Calibar, BG...), không tính hãng / loại hàng / chữ phiên bản (cuộn, pro, max)
    return bool({w for w in ki & kc if w.isalpha() and w not in _FAM_WEAK})


_FAM_WEAK = set("""yonex lining li ning victor gosen mizuno kawasaki felet apacs joola kaiwin facolos taro asics ligpro
arronax spikopoll diadem franklin selkirk head wilson babolat adidas nike kumpoo vnb hai yen
cuon pro max plus light power speed control tour lite new ii iii iv game edition limited
giay ao quan vay tui balo bao bong binh xit bang ong hop vo tat dep mieng gen chi boc thun""".split())


_CAT_MAP = {"khung": "vot", "vot": "vot", "vet": "vot", "day": "cuoc", "cuoc": "cuoc", "cuon": "cuoc", "ao": "ao",
            "quan": "quan", "vay": "quan", "giay": "giay", "dep": "giay", "tat": "tat", "vo": "tat", "tui": "tui",
            "balo": "tui", "bao": "tui", "bong": "bong", "qua": "cau", "ong": "cau", "cau": "cau", "binh": "xit",
            "xit": "xit", "bang": "bang", "mieng": "phukien", "gen": "phukien", "chi": "phukien", "boc": "phukien"}


def _cat(name: str) -> str:
    """Loại hàng theo chữ đầu tên (bỏ mã NCC kiểu "P-AYPU011-4-"): vợt / cước / áo / quần / giày / tất / túi..."""
    ws = [w for w in _norm(name).split() if w.isalpha() and len(w) <= 6 and w not in ("p",)]
    for k, w in enumerate(ws[:3]):
        nxt = ws[k + 1] if k + 1 < len(ws) else ""
        if w in ("quan", "boc") and nxt == "can":      # "Quấn cán" / "Bọc cán": phụ kiện, không phải quần / vợt
            return "phukien"
        if w in _CAT_MAP:
            c = _CAT_MAP[w]
            if c == "cau" and len(ws) > 1 and ws[ws.index(w) + 1:ws.index(w) + 2] == ["long"]:
                continue     # "cầu lông" là môn, không phải loại hàng
            return c
    # tên OCR cụt không có chữ loại hàng: nhận theo dấu hiệu (chưa căng lưới / 4U5 / mã vợt Yonex BFR...)
    t = " " + _norm(name) + " "
    if re.search(r" (chua cang luoi|cang luoi|[3-6]u[4-7]?|bfr[0-9a-z]*) ", t):
        return "vot"
    return ""


_OCR_DIGIT = str.maketrans("iol", "101")
_YONEX_LINE = {"ax": "Astrox", "arc": "Arcsaber", "nf": "Nanoflare", "vc": "Voltric", "dr": "Duora", "nr": "Nanoray"}


def _expand_sku(name: str) -> str:
    """Giải nghĩa mã vợt Yonex trong tên (OCR hay chỉ còn mã): "BFR3AX88DZTZ" -> thêm "Yonex Astrox 88D"."""
    out = []
    for w in _norm(name).split():
        m = re.match(r"bfr\d?(ax|arc|nf|vc|dr|nr)([0-9iol]{1,3})([a-y]?)z", w)   # OCR: II = 11, IOO = 100
        if m:
            out.append(f"Vợt Yonex {_YONEX_LINE[m.group(1)]} {m.group(2).translate(_OCR_DIGIT)}{m.group(3).upper()}")
    return f"{name} {' '.join(out)}" if out else name


def _available(items: list[dict[str, Any]], groups: dict[str, list[dict[str, Any]]]) -> dict[str, float]:
    """Số lượng đang có của từng mã: tồn đầu kỳ + mua vào đã ghép mã."""
    av: dict[str, float] = {}
    for i in items:
        av[i["ma"].upper()] = av.get(i["ma"].upper(), 0.0) + float(i.get("sl") or 0)
    for ln in groups.get("Mua vào", []):
        if ln.get("ma_ghep"):
            av[ln["ma_ghep"].upper()] = av.get(ln["ma_ghep"].upper(), 0.0) + float(ln.get("sl") or 0)
    return av


def match_codes(items: list[dict[str, Any]], groups: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    """Gắn mã hàng cho mọi dòng (sửa trực tiếp line["ma_ghep"]). Trả về bảng ghép theo từng tên hàng.
    Ưu tiên MÃ CHUẨN của file đầu kỳ: mã có sẵn -> trùng tên -> khớp từ khóa tên chuẩn -> tên gần đúng -> hàng mới."""
    base = list(items)
    by_code = {i["ma"].upper(): i for i in items}
    by_name = {_norm(i["ten"]): i for i in items}
    used = set(by_code)
    memo: dict[str, dict[str, Any]] = {}
    kw_cache: dict[int, frozenset[str]] = {}
    avail: dict[str, float] = {}
    report = []
    # mua vào ghép TRƯỚC (có thể tạo mã mới), bán ra / trả lại ghép SAU vào hàng đã có
    order = sorted(groups.items(), key=lambda kv: kv[0] != "Mua vào")
    for kind, lines in order:
        for ln in lines:
            if ln["ma"] and ln["ma"].upper() in by_code:
                ln["ma_ghep"], ln["cach"] = by_code[ln["ma"].upper()]["ma"], "Mã có sẵn trong file"
                continue
            nm = _expand_sku(ln["ten"])      # tên + mã SKU giải nghĩa (BFR3AX88D -> Astrox 88D)
            key = _norm(nm)
            if key in memo:
                m = memo[key]
                if kind != "Mua vào" and "MỚI" in m["cach"]:     # bán đúng hàng vừa mua (mã tạo lúc ghép mua vào)
                    m = {**m, "cach": "Hàng mua mới trong kỳ (mã vừa tạo khi ghép mua vào)"}
            else:
                buying = kind == "Mua vào"
                # mua vào: so với danh mục đầu kỳ; bán ra / trả lại: so với MỌI hàng đang có (đầu kỳ + hàng mới mua)
                # hàng mới tạo khi mua (vd áo Yonex-2993 size XS) cũng làm mã chuẩn để size / màu sau gom vào
                pool = items
                pool_kw = [kw_cache.setdefault(id(i), _kw(i["ten"])) for i in items]
                it = by_name.get(key)
                kw_it, kw_sc, kw_why = (None, 0.0, "") if it else _kw_match(nm, pool, pool_kw, ln["dvt"])
                if it:
                    m = {"ma": it["ma"], "diem": 1.0, "cach": "Trùng tên trong danh mục"}
                elif kw_it is not None:
                    m = {"ma": kw_it["ma"], "diem": round(kw_sc, 2), "cach": kw_why}
                elif fam := [i for i in pool if _family(nm, i["ten"], fk)] if (fk := _fam_kw(nm)) else []:
                    if not buying:                     # bán ra: ưu tiên hàng cùng dòng còn tồn
                        if not avail:
                            avail.update(_available(items, groups))
                        fam = [i for i in fam if avail.get(i["ma"].upper(), 0) > 0] or fam
                    best = _pick(nm, ln["dvt"], fam)
                    m = {"ma": best["ma"], "diem": 0.85,
                         "cach": "Cùng dòng / model, khác phiên bản – màu – size: ghép mã chuẩn – CẦN KIỂM TRA"}
                else:
                    scored = []
                    for i in (base if buying else pool):
                        s = _score(key, _norm(i["ten"]))
                        if ln["dvt"] and i["dvt"] and _norm(ln["dvt"]) != _norm(i["dvt"]):
                            s -= 0.08
                        scored.append((s, i))
                    scored.sort(key=lambda x: -x[0])
                    sc, best = scored[0] if scored else (0.0, None)
                    if not buying and scored:
                        # bán ra phải là hàng ĐANG CÓ: trong các hàng gần nhất, ưu tiên hàng còn tồn (đầu kỳ + đã mua)
                        if not avail:
                            avail.update(_available(items, groups))
                        cat = _cat(nm)
                        same = [(s, i) for s, i in scored if cat and _cat(i["ten"]) == cat]
                        if same:                     # cùng loại hàng (vợt với vợt, áo với áo...)
                            scored = same
                            sc, best = scored[0]
                        near = [i for s, i in scored if s >= sc - 0.05]
                        stocked = [i for i in near if avail.get(i["ma"].upper(), 0) > 0]
                        if stocked:
                            best = stocked[0]
                            sc = next(s for s, i in scored if i is best)
                    if best and sc >= AUTO_OK:
                        m = {"ma": best["ma"], "diem": round(sc, 2), "cach": "Tên gần đúng"}
                    elif best and (sc >= MAYBE or not buying):
                        # bán ra KHÔNG tạo mã mới: luôn ghép vào hàng có sẵn gần nhất, đánh dấu để kiểm tra
                        m = {"ma": best["ma"], "diem": round(sc, 2),
                             "cach": "Tên gần đúng – CẦN KIỂM TRA" if sc >= MAYBE
                             else "Bán ra: ghép vào hàng có sẵn gần nhất (không tạo mã mới) – CẦN KIỂM TRA"}
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
    _rebalance_sales(items, groups, report)
    return report


def _line_words(name: str) -> frozenset[str]:
    """Tên hàng bỏ hết số hiệu / mã phía sau: "Ống cầu lông Hải Yến S90" -> {hai, yen}."""
    return frozenset(w for w in _kw(name) if w.isalpha())


def _near_tier(a: str, b: str) -> int:
    """Mức gần giống để bù hàng thiếu tồn: 1 = cùng model khác phiên bản, 2 = cùng tên dòng khác số hiệu
    (Halbertec 6000 ~ 1000), 3 = cùng loại + cùng tên chỉ khác mã phía sau (Hải Yến S90 ~ S70); 0 = khác hàng."""
    if _cat(a) != _cat(b):
        return 0
    if _family(a, b):
        return 1
    la, lb = _line_words(a), _line_words(b)
    if not la or not lb or not (la <= lb or lb <= la):
        return 0
    return 2 if (la & lb) - _FAM_WEAK else 3


def _rebalance_sales(items: list[dict[str, Any]], groups: dict[str, list[dict[str, Any]]],
                     report: list[dict[str, Any]]) -> None:
    """Bán ra VƯỢT số đang có của mã đã ghép (đầu kỳ + mua + trả lại): phần thiếu chuyển sang hàng GẦN GIỐNG còn dư
    (ưu tiên cùng model khác phiên bản -> cùng dòng khác số hiệu -> cùng tên chỉ khác mã phía sau), vd thiếu
    "BG 65 cuộn" -> "BG 65 Titanium", thiếu "Hải Yến S90" -> "Hải Yến S70"."""
    sell = [ln for ln in groups.get("Bán ra") or [] if ln["sl"] > 0 and ln.get("ma_ghep")]
    if not sell:
        return
    have: dict[str, float] = {}
    for i in items:
        have[i["ma"].upper()] = have.get(i["ma"].upper(), 0.0) + float(i.get("sl") or 0)
    for kind in ("Mua vào", "Hàng bán trả lại"):
        for ln in groups.get(kind) or []:
            if ln.get("ma_ghep") and ln["sl"] > 0:
                have[ln["ma_ghep"].upper()] = have.get(ln["ma_ghep"].upper(), 0.0) + ln["sl"]
    need: dict[str, float] = {}
    for ln in sell:
        need[ln["ma_ghep"].upper()] = need.get(ln["ma_ghep"].upper(), 0.0) + ln["sl"]
    spare = {k: have.get(k, 0.0) - need.get(k, 0.0) for k in set(have) | set(need)}   # dư (+) / thiếu (-)
    by_code = {i["ma"].upper(): i for i in items}
    tiers: dict[tuple[str, str], list[list[dict[str, Any]]]] = {}
    moved: dict[tuple[str, str], str] = {}
    for k in sorted(c for c, v in spare.items() if v < 0):
        cur = by_code.get(k)
        lines = [ln for ln in sell if ln["ma_ghep"].upper() == k and ln.get("cach") != "Mã có sẵn trong file"]
        # dòng bán muộn nhất chuyển trước (hàng đầu kỳ xuất trước)
        lines.sort(key=lambda x: (x.get("ngay") or date.min, x.get("row", 0)), reverse=True)
        while spare[k] < 0 and lines:
            # chuyển ÍT dòng nhất: dòng nhỏ nhất vừa đủ phần thiếu, không có thì dòng lớn nhất
            gap = -spare[k]
            enough = [x for x in lines if x["sl"] >= gap]
            ln = min(enough, key=lambda x: x["sl"]) if enough else max(lines, key=lambda x: x["sl"])
            lines.remove(ln)
            tk = (ln["ten"], k)
            if tk not in tiers:
                lv: list[list[dict[str, Any]]] = [[], [], []]
                for i in items:
                    if i["ma"].upper() == k:
                        continue
                    t = max(_near_tier(_expand_sku(ln["ten"]), i["ten"]), 0)
                    if cur is not None:
                        t2 = _near_tier(cur["ten"], i["ten"])
                        t = min(x for x in (t, t2) if x) if (t or t2) else 0
                    if t:
                        lv[t - 1].append(i)
                tiers[tk] = lv
            best = None
            for grp in tiers[tk]:
                alt = [i for i in grp if spare.get(i["ma"].upper(), 0.0) >= ln["sl"]]
                if alt:
                    best = _pick(_expand_sku(ln["ten"]), ln["dvt"], alt)
                    break
            if best is None and ln.get("cach", "").startswith("Bán ra: ghép vào hàng có sẵn gần nhất"):
                # tên đọc không rõ (OCR cụt) vốn đã ghép tạm: lấy hàng cùng loại tên gần nhất còn đủ tồn
                cat = _cat(_expand_sku(ln["ten"])) or (_cat(cur["ten"]) if cur else "")
                alt = [i for i in items if i["ma"].upper() != k and spare.get(i["ma"].upper(), 0.0) >= ln["sl"]
                       and (not cat or _cat(i["ten"]) == cat)]
                if alt:
                    best = _pick(_expand_sku(ln["ten"]), ln["dvt"], alt)
            if best is None:
                continue
            spare[best["ma"].upper()] -= ln["sl"]
            spare[k] += ln["sl"]
            moved[(ln["ten"], best["ma"])] = ln["ma_ghep"]
            ln["cach"] = f"Mã {ln['ma_ghep']} không đủ tồn – xuất hàng gần giống còn tồn {best['ma']} – CẦN KIỂM TRA"
            ln["ma_ghep"] = best["ma"]
    for (ten, ma), old in moved.items():
        report.append({"loai": "Bán ra", "ten": ten, "dvt": by_code[ma.upper()]["dvt"], "ma": ma, "diem": 0.8,
                       "cach": f"Chuyển từ mã {old} (không đủ tồn) sang hàng gần giống còn tồn – CẦN KIỂM TRA",
                       "ten_dm": by_code[ma.upper()]["ten"]})


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


DEFAULT_TPL = Path(__file__).resolve().parent / "assets" / "mau_tong_hop_ton_kho.xlsx"


def write_report(tpl: bytes, rows: list[dict[str, Any]], ky_text: str, ngay_ky: date | None, ten_dn: str = "") -> bytes:
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
    hr = next((r for r in range(1, 40) if _norm(ws.cell(r, 2).value) == "ma hang" or _norm(ws.cell(r, 1).value) == "ten kho"), None)
    if not hr:
        raise ValueError("File mẫu Tổng hợp tồn kho không có dòng tiêu đề 'Tên kho / Mã hàng' như báo cáo MISA")
    if ten_dn:
        ws.cell(1, 1).value = ten_dn
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
        # chiều cao dòng theo độ dài tên hàng (cột tên hàng ~30 ký tự / dòng như mẫu MISA)
        ln = len(str(x.get("ten") or ""))
        ws.row_dimensions[r].height = 15.75 * max(1, -(-ln // 30))
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
                import threading
                th = threading.Thread(target=mi.run, args=(mj, part["invoices"], tpl, direction, inp.get("mst", ""),
                                                            inp.get("acc") or {}, tpl_name), daemon=True)
                th.start()
                seen = 0
                while th.is_alive() or seen < len(mj.log):     # báo tiến độ đọc hóa đơn ra màn hình (không "đứng")
                    th.join(2)
                    for m in mj.log[seen:]:
                        job.log.append(f"{m[:9]}{labels[key]}: {m[9:]}")
                    seen = len(mj.log)
                if mj.status != "done":
                    raise RuntimeError(f"Chuyển hóa đơn {labels[key].lower()} lỗi: " + " | ".join(mj.log[-3:]))
                src = (Path(mj.file).name, Path(mj.file).read_bytes())
                job.say(f"{labels[key]}: đã điền {mj.done} hóa đơn vào mẫu ({Path(mj.file).name})")
            if src:
                pk = read_lines(src[1])
                if pk["vi_du"]:
                    job.say(f"{labels[key]}: bỏ qua {pk['vi_du']} dòng ví dụ có sẵn của file mẫu MISA")
                if not pk["lines"]:
                    job.say(f"{labels[key]}: file chưa có dữ liệu – bỏ qua")
                    continue
                sources[key] = src
                parsed[key] = pk
                job.say(f"{labels[key]}: {len(pk['lines'])} dòng hàng")
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
        if inp.get("tpl"):
            tpl_name, tpl = inp["tpl"]
        elif opening["loai"] == "tonkho":          # đầu kỳ là báo cáo Tổng hợp tồn kho MISA: dùng luôn định dạng đó
            tpl_name, tpl = op_name, op_data
        else:                                    # mẫu Tổng hợp tồn kho chuẩn MISA có sẵn trong phần mềm
            tpl_name, tpl = DEFAULT_TPL.name, DEFAULT_TPL.read_bytes()
        f1 = out_dir / f"TONG_HOP_TON_KHO_{re.sub(r'[^0-9A-Za-z]+', '_', fold(ky_text))}.xlsx"
        f1.write_bytes(write_report(tpl, rows, ky_text, ngay, inp.get("ten_dn", "")))
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
        # 2 bản so sánh tên hàng ghép: mua vào ↔ đầu kỳ, bán ra ↔ mua vào ↔ đầu kỳ
        from .inventory_compare import compare_buy, compare_sell
        buy_l, sell_l = parsed.get("buy", {}).get("lines", []), parsed.get("sell", {}).get("lines", [])
        if buy_l:
            f = out_dir / "So_sanh_Mua_vao_Dau_ky.xlsx"
            f.write_bytes(compare_buy(buy_l, items))
            job.files["ss_mua"] = str(f)
        if sell_l:
            f = out_dir / "So_sanh_Ban_ra_Mua_vao_Dau_ky.xlsx"
            f.write_bytes(compare_sell(sell_l, buy_l, items, rows))
            job.files["ss_ban"] = str(f)
        job.files["folder"] = str(out_dir)
        job.status = "done"
        job.say("Xong.")
    except Exception as e:  # noqa: BLE001
        job.status = "error"
        job.say(f"LỖI: {e or type(e).__name__}")


def start(inp: dict[str, Any]) -> Job:
    import threading
    job = Job(id=uuid.uuid4().hex[:12])
    JOBS[job.id] = job
    threading.Thread(target=run, args=(job, inp), daemon=True, name="inventory").start()
    return job
