"""Đọc hàng loạt file hóa đơn: XML (TT78), ZIP (chứa XML/PDF) và PDF hóa đơn điện tử có lớp chữ.

PDF được đọc theo nhãn chuẩn của mẫu hóa đơn TT78 (tiếng Việt và song ngữ):
ký hiệu, số, ngày, MST người bán/mua, tên, cộng tiền hàng, tiền thuế, tổng thanh toán, mã CQT, mã tra cứu.
PDF dạng ảnh scan (không có lớp chữ) sẽ được đánh dấu "cần OCR".
"""
from __future__ import annotations

import io
import re
import zipfile
from typing import Any, Iterator

from .gdt.xml_parser import parse_invoice_xml


def vn_num(s: Any) -> float | None:
    """'826.999.992' -> 826999992 ; '58.981,5' -> 58981.5 ; '1,200.00' -> 1200.0"""
    if s is None or s == "":
        return None
    if isinstance(s, (int, float)):
        return float(s)
    s = str(s).strip().replace(" ", "").replace(" ", "")
    neg = s.startswith("-") or (s.startswith("(") and s.endswith(")"))
    s = s.strip("-()")
    if not re.fullmatch(r"[\d.,]+", s):
        return None
    if "." in s and "," in s:
        dec = "," if s.rfind(",") > s.rfind(".") else "."
        s = s.replace("." if dec == "," else ",", "").replace(dec, ".")
    elif re.fullmatch(r"\d{1,3}([.,]\d{3})+", s):
        s = re.sub(r"[.,]", "", s)
    else:
        s = s.replace(",", ".")
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def iter_files(name: str, data: bytes, depth: int = 0) -> Iterator[tuple[str, bytes]]:
    low = name.lower()
    if data[:2] == b"PK" and low.endswith(".zip") and depth < 3:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            for n in zf.namelist():
                if not n.endswith("/"):
                    yield from iter_files(f"{name}/{n}", zf.read(n), depth + 1)
        return
    yield name, data


# ------------------------------------------------------------------ PDF
_L = r"(?:\s*\([^)]{1,40}\))?"      # nhãn tiếng Anh trong ngoặc, vd "(Serial)"
_MONEY = r"(-?[\d][\d.,]*)"

# Mẫu tìm trên chữ ĐÃ BỎ DẤU (fold) -> chạy được cả với chữ OCR bị mất/sai dấu ("Mä thué", "Töng tién").
# Giá trị lấy từ chữ gốc theo cùng vị trí (fold giữ nguyên độ dài chuỗi).
_C = r"\s*[:;.•]?\s*"
RX = {
    "ky_hieu": re.compile(r"Ky\s*hieu" + _L + r"\s*[:;.•]*\s*([1-6Il|]?\s?[CK]\s?\d{2}\s?[A-Z]{3})\b", re.I),
    "so": re.compile(r"(?<![A-Za-z])So" + _L + r"\s*[:;]\s*(\d{1,8})\b"),
    "ngay": re.compile(r"Ngay" + _L + r"\s*(\d{1,2})\s*thang" + _L + r"\s*(\d{1,2})\s*nam" + _L + r"\s*(\d{4})", re.I),
    # MST 10 số, 13 số (chi nhánh) hoặc 12 số (cá nhân dùng số CCCD). OCR hay mất chữ "số": "Mä thué"
    "mst": re.compile(r"(?:Ma\s*(?:[a-z0-9]{1,3}\s*)?thu[eu]|MST\s*/\s*CCCD[^:\n]{0,15}|T[ae][xr]\s*code\s*\)?)" + _L
                      + r"[^\d\n]{0,25}?((?:\d[ \t.]?){9}\d(?:[ \t]?\d){0,2}(?:[ \t]?-?[ \t]?\d{3})?)", re.I),
    "chua_thue": re.compile(r"(?:Cong\s*tien\s*hang|Tong\s*tien\s*chua\s*(?:co\s*)?thue|Tong\s*cong\s*tien\s*hang)"
                            r"(?:\s*\([^)]*\))?[^\d\n]{0,40}?" + _MONEY, re.I),
    "tong_thue": re.compile(r"Tong\s*(?:cong\s*)?tien\s*thue(?:\s*G\s*T\s*G\s*T?)?(?:\s*\([^)]*\))?[^\d\n]{0,30}?" + _MONEY, re.I),
    "thue": re.compile(r"Tien\s*thu[eu]\s*G\s*T\s*G\s*T?(?:\s*\([^)]*\))?[^\d\n]{0,30}?" + _MONEY, re.I),
    "tong_tt": re.compile(r"Tong\s*(?:cong\s*)?(?:so\s*)?tien\s*thanh\s*toan(?:\s*\([^)]*\))?[^\d\n]{0,30}?" + _MONEY, re.I),
    "thue_suat": re.compile(r"Thue\s*suat(?:\s*GTGT)?" + _L + r"\s*[:;]\s*(\d{1,2}\s*(?:%|0/0)|KCT|KKKNT)", re.I),
    "mccqt": re.compile(r"(?:Ma\s*(?:cua\s*)?(?:CQT|co\s*quan\s*thue)|MCCQT)" + _L + r"\s*[:;.•]*\s*([0-9A-Z][0-9A-Z-]{15,})", re.I),
    "ma_tra_cuu": re.compile(r"Ma\s*tra\s*c[uc]u(?:\s*h[oa]a\s*d[oa]n)?" + _L + r"\s*[:;]\s*([^\s]+)", re.I),
    "nguoi_ban": re.compile(r"(?:Don\s*(?:v[il]\s*)?ban\s*hang|Ten\s*nguoi\s*ban)" + _L + r"\s*[:;]\s*(.+)", re.I),
    "nguoi_mua": re.compile(r"(?:Ten\s*don\s*vi|Ten\s*nguoi\s*mua)" + _L + r"\s*[:;]\s*(.+)", re.I),
    "tong_cong3": re.compile(r"Tong\s*cong" + _L + r"\s*:\s*([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", re.I),
}


def fold(s: str) -> str:
    """Bỏ dấu tiếng Việt, GIỮ NGUYÊN độ dài (mỗi ký tự -> 1 ký tự)."""
    import unicodedata
    out = []
    for ch in s:
        if ch in "đĐ":
            out.append("d" if ch == "đ" else "D")
            continue
        b = unicodedata.normalize("NFD", ch)[0]
        out.append(b if len(b) == 1 else ch)
    return "".join(out)


_MONEY_TOKEN = re.compile(r"(?<![\d.,])\d{1,3}(?:[.,]\d{3})+(?:,\d{1,2})?(?![\d])")


def pdf_text(data: bytes) -> str:
    import pdfplumber
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        return "\n".join((p.extract_text() or "") for p in pdf.pages[:4])


def _ocr_clean(t: str) -> str:
    """Sửa lỗi OCR hay gặp trong số tiền: 'o'/'O' thành '0' ('41.400.ooo', '2.ooo.ooo', '1.200,oo')
    và bỏ khoảng trắng chen giữa số ('17.850 .ooo'). Lặp tới khi ổn định."""
    prev = None
    while prev != t:
        prev = t
        t = re.sub(r"(?<=\d)[ \t]+(?=[.,](?:\d{3}|[oO0]{3})(?![A-Za-z]))", "", t)
        t = re.sub(r"(?<=[\d][.,])[oO]{1,3}(?=$|[^A-Za-z])|(?<=[0-9oO][.,])[oO]{3}(?=$|[^A-Za-z])|(?<=\d)[oO](?=[.,]\d)",
                   lambda m: "0" * len(m.group(0)), t, flags=re.M)
    return t


def parse_pdf_text(t: str, source: str = "PDF") -> dict[str, Any]:
    if source == "OCR":
        t = _ocr_clean(t)
    f = fold(t)

    def g(k: str, grp: int = 1) -> str | None:
        m = RX[k].search(f)
        return t[m.start(grp):m.end(grp)].strip() if m else None

    out: dict[str, Any] = {"nguon": source}
    kh = g("ky_hieu")
    if not kh:  # nhãn "Ký hiệu" bị OCR đọc sai ("K! hieu", "KY hi$u"): tìm thẳng mẫu ký hiệu 7 ký tự
        m = re.search(r"(?<![A-Za-z0-9])([1-6Il|][CK]\d{2}[A-Z]{3})(?![A-Za-z0-9])", f)
        kh = m.group(1) if m else None
    if kh:
        kh = re.sub(r"\s", "", kh)
        kh = re.sub(r"^[Il|]", "1", kh).upper()  # OCR: "IC26TDD" -> "1C26TDD"
        out["khmshdon"], out["khhdon"] = (kh[0], kh[1:]) if kh[0].isdigit() else ("", kh)
    so = g("so")
    if not so:  # nhãn "Số" bị OCR đọc sai nhưng còn "(No.)": "S6 (No.): 587", "s6(N0.): 2105"
        m = re.search(r"\(\s*N[o0]\.?\s*\)\s*[:;.]?\s*(\d{1,8})\b", f, re.I)
        so = m.group(1) if m else None
    if not so and kh:  # nhãn "Số" mất hẳn: số đứng cuối dòng, trong vài dòng ngay sau ký hiệu
        km = RX["ky_hieu"].search(f)
        m = re.search(r"(?m)[^\d.,:/](\d{1,8})[ \t]*$", f[km.end():km.end() + 160]) if km else None
        so = m.group(1) if m else None
    if not so:  # OCR tách nhãn "Số:" và giá trị: lấy số có đệm 0 kiểu "00001384"
        m = re.search(r"(?<![\d.,])0{2,}\d{1,6}(?![\d.,])", f)
        so = m.group(0) if m else None
    out["shdon"] = str(int(so)) if so else None
    m = RX["ngay"].search(f)
    out["ngay_lap"] = None
    if m and 1 <= int(m.group(1)) <= 31 and 1 <= int(m.group(2)) <= 12:
        out["ngay_lap"] = f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    else:  # ngày đọc lỗi: dùng ngày dạng số đầu tiên (thường là "Ký ngày: 01-10-2026")
        # OCR hay đọc số thành chữ trong ngày: "02- IO- 2026" -> 02-10-2026
        for m in re.finditer(r"(?<![\w])([\dIlO|]{1,2})\s*[-/.]\s*([\dIlO|]{1,2})\s*[-/.]\s*(20\d{2})(?!\d)", f):
            dd, mm = (x.translate(_DIGIT_FIX) for x in (m.group(1), m.group(2)))
            if dd.isdigit() and mm.isdigit() and 1 <= int(dd) <= 31 and 1 <= int(mm) <= 12:
                out["ngay_lap"] = f"{m.group(3)}-{int(mm):02d}-{int(dd):02d}"
                break
    out["nbmst"], out["nmmst"] = _pick_msts(f)
    nb = g("nguoi_ban")
    if not nb:  # mẫu không có nhãn: tên người bán là dòng ngay trên "Mã số thuế" đầu tiên
        lines, flines = t.splitlines(), f.splitlines()
        for i, ln in enumerate(flines):
            if re.search(r"Ma\s*(so\s*)?thue", ln, re.I) and i > 0:
                nb = lines[i - 1].strip()
                break
    out["nbten"] = nb
    out["nmten"] = g("nguoi_mua")
    out["tgtcthue"] = vn_num(g("chua_thue"))
    tt = vn_num(g("tong_thue"))
    if tt is None:
        vals = [vn_num(t[m.start(1):m.end(1)]) for m in RX["thue"].finditer(f)]
        vals = [v for v in vals if v is not None]
        tt = sum(vals) if vals else None
    out["tgtthue"] = tt
    out["tgtttbso"] = vn_num(g("tong_tt"))
    m = RX["tong_cong3"].search(f)  # bảng tổng hợp nhiều thuế suất: "Tổng cộng: <trước thuế> <thuế> <thanh toán>"
    if m and out["tgtttbso"] is None:
        out["tgtcthue"], out["tgtthue"], out["tgtttbso"] = (vn_num(m.group(k)) for k in (1, 2, 3))
    _fix_amounts(out, f)
    ts = g("thue_suat")
    out["thue_suat"] = re.sub(r"\s|0/0", lambda x: "" if x.group(0).isspace() else "%", ts) if ts else None
    mc = g("mccqt")
    if mc and source == "OCR" and re.fullmatch(r"[0-9A-FO]{30,}", mc):  # mã CQT dạng hex: chữ O là số 0
        mc = mc.replace("O", "0")
    out["mccqt"] = mc
    out["ma_tra_cuu"] = g("ma_tra_cuu")
    return out


_DIGIT_FIX = str.maketrans("IlO|", "1101")


def _rate_ok(tax: float, base: float) -> bool:
    """Thuế / tiền hàng đúng 1 mức thuế suất (0, 5, 8, 10%) – sai lệch do làm tròn."""
    if base <= 0 or tax < 0:
        return False
    return any(abs(tax - base * r) <= max(2.0, base * 0.0005) for r in (0, 0.05, 0.08, 0.1))


_BUYER = re.compile(r"nguoi\s*mua|mua\s*h[ai]ng|\(\s*Buyer\s*\)|Ten\s*don\s*v|Ho\s*ten", re.I)
# số đứng sau các nhãn này không phải MST: điện thoại, số tài khoản, MST đơn vị cung cấp phần mềm (chân trang)
_NOT_MST = re.compile(r"(thoai|tel|fax|\bdt|hotline|khoan|stk|a/c|giai\s*phap|phat\s*hanh|cung\s*cap|misa|"
                      r"easyinvoice|viettel|vnpt|bkav|fpt)[^\n]{0,60}$", re.I)


def _pick_msts(f: str) -> tuple[str | None, str | None]:
    """MST người bán / người mua theo VỊ TRÍ: người bán đứng trước phần "người mua", người mua đứng sau.
    Nhãn MST bị OCR đọc hỏng (chỉ còn dãy số) -> lấy dãy 10/13 số đứng riêng, trừ điện thoại / số tài khoản /
    MST nhà cung cấp phần mềm ở chân trang."""
    bm = _BUYER.search(f)
    cut = bm.start() if bm else len(f)
    lab = [(m.start(), re.sub(r"\D", "", f[m.start(1):m.end(1)])) for m in RX["mst"].finditer(f)]
    lab = [(p, x) for p, x in lab if len(x) in (10, 12, 13)]
    loose = [(m.start(), re.sub(r"\D", "", m.group(1)))
             for m in re.finditer(r"(?<![\d.,])(\d{10}(?:[ \t]?-[ \t]?\d{3})?)(?![\d.,])", f)
             if not _NOT_MST.search(f[max(0, m.start() - 70):m.start()].split("\n")[-1])]
    seller = next((x for p, x in lab if p < cut), None) or next((x for p, x in loose if p < cut), None)
    buyer = next((x for p, x in lab if p >= cut and x != seller), None)
    if not seller and lab:
        seller = lab[0][1]
        buyer = next((x for p, x in lab if x != seller), None)
    if not buyer:
        buyer = next((x for p, x in lab if x != seller), None)
    if not buyer:  # nhãn MST người mua bị hỏng: dãy số đứng riêng trong vài dòng sau "người mua"
        buyer = next((x for p, x in loose if cut <= p < cut + 400 and x != seller), None)
    return _fmt_mst(seller), _fmt_mst(buyer)


def _fmt_mst(x: str | None) -> str | None:
    """MST chi nhánh 13 số viết theo chuẩn có gạch: 3702177878-001."""
    return f"{x[:10]}-{x[10:]}" if x and len(x) == 13 else x


def _fix_amounts(out: dict[str, Any], f: str) -> None:
    """Khi nhãn tiền và số bị tách (bố cục cột / OCR): tổng thanh toán = số tiền lớn nhất,
    tìm cặp (tiền hàng, thuế) cộng lại đúng bằng tổng thanh toán."""
    a, b, c = out.get("tgtcthue"), out.get("tgtthue"), out.get("tgtttbso")
    if a is not None and b is not None and c is not None and abs(a + b - c) <= 2:
        return  # đã khớp
    if a and b is not None and c is not None:  # đủ 3 số nhưng lệch: OCR đọc sai 1 chữ số
        if _rate_ok(c - a, a):    # tiền hàng + tổng hợp lý -> sửa tiền thuế
            out["tgtthue"] = round(c - a, 2)
            return
        if _rate_ok(b, c - b):    # thuế + tổng hợp lý -> sửa tiền hàng
            out["tgtcthue"] = round(c - b, 2)
            return
        return  # tổng có thể gồm phí/lệ phí ngoài thuế GTGT -> giữ nguyên số đọc được
    # OCR làm mất đúng 1 con số: tính từ 2 số còn lại (tổng = tiền hàng + thuế)
    if a is not None and b is not None and c is None and b < a:
        out["tgtttbso"] = a + b
        return
    if a is not None and c is not None and b is None and 0 <= c - a < a:
        out["tgtthue"] = round(c - a, 2)
        return
    nums = sorted({v for v in (vn_num(x) for x in _MONEY_TOKEN.findall(f)) if v and v >= 1000}, reverse=True)
    if not nums:
        return
    total = c if c is not None else nums[0]
    if c is not None and a is None and b is not None:
        out["tgtcthue"] = total - b
        return
    for x in nums:
        if x >= total:
            continue
        y = round(total - x, 2)
        if y in nums or abs(y) < 1:
            out["tgtttbso"] = total
            out["tgtcthue"], out["tgtthue"] = (x, y) if x >= y else (y, x)
            return
    if c is None:
        out["tgtttbso"] = total


# ------------------------------------------------------------------ OCR 2 lượt
_OCR_KEY = ("khhdon", "shdon", "ngay_lap", "nbmst", "tgtttbso")


def ocr_read(data: bytes, max_pages: int = 3) -> tuple[dict[str, Any], str]:
    """OCR hóa đơn scan / ảnh chụp. Lượt 1: ảnh xám. Còn thiếu trường chính (thường do số hóa đơn, ký hiệu in
    màu đỏ/xanh bị nhạt) -> lượt 2 đọc theo kênh màu đậm nhất, chỉ lấy bù các trường còn trống.
    Trả về (trường đọc được, toàn bộ chữ OCR – để AI đọc tiếp nếu cần)."""
    from . import ocr
    imgs = ocr.page_images(data, max_pages)
    text = "\n".join(ocr.ocr_pil(i) for i in imgs)
    rec = parse_pdf_text(text, "OCR")
    if any(not rec.get(k) for k in _OCR_KEY) and any(i.mode not in ("L", "1", "LA") for i in imgs):
        t2 = "\n".join(ocr.ocr_pil(i, ink=True) for i in imgs)
        r2 = parse_pdf_text(t2, "OCR")
        for k, v in r2.items():
            if v and not rec.get(k):
                rec[k] = v
        _fix_amounts(rec, fold(text + "\n" + t2))
        text += "\n" + t2
    return rec, text


# ------------------------------------------------------------------ đọc 1 file
def read_invoice(name: str, data: bytes) -> dict[str, Any]:
    low = name.lower()
    rec: dict[str, Any] = {"file": name}
    try:
        if low.endswith(".xml") or data.lstrip()[:5] == b"<?xml" or data.lstrip()[:5] == b"<HDon":
            p = parse_invoice_xml(data)
            nb, nm = p.get("nguoi_ban") or {}, p.get("nguoi_mua") or {}
            rec.update(nguon="XML", khmshdon=p.get("khmshdon"), khhdon=p.get("khhdon"),
                       shdon=str(int(p["shdon"])) if (p.get("shdon") or "").isdigit() else p.get("shdon"),
                       ngay_lap=(p.get("ngay_lap") or "")[:10], nbmst=nb.get("mst"), nbten=nb.get("ten"),
                       nmmst=nm.get("mst"), nmten=nm.get("ten"), tgtcthue=p.get("tong_chua_thue"),
                       tgtthue=p.get("tong_thue"), tgtttbso=p.get("tong_thanh_toan"), mccqt=p.get("mccqt"),
                       ma_tra_cuu=p.get("lookup_code"),
                       thue_suat=",".join(k for k in (p.get("theo_thue_suat") or {}) if k))
        elif low.endswith(".pdf") or data[:5] == b"%PDF-":
            t = pdf_text(data)
            if len(t.strip()) < 40:  # PDF scan dạng ảnh -> OCR
                rec.update(ocr_read(data)[0])
            else:
                rec.update(parse_pdf_text(t))
        else:
            from . import ocr
            if ocr.is_image(name, data):  # ảnh chụp / scan hóa đơn
                rec.update(ocr_read(data)[0])
            else:
                rec["loi"] = "Định dạng chưa hỗ trợ"
        if rec.get("nguon") in ("PDF", "OCR"):
            miss = [k for k in ("khhdon", "shdon", "nbmst", "tgtttbso") if not rec.get(k)]
            if miss:
                rec["loi"] = "Không đọc được: " + ", ".join(miss) + (" (bản OCR – nên kiểm tra lại)" if rec["nguon"] == "OCR" else "")
    except Exception as e:  # noqa: BLE001
        rec["loi"] = f"Lỗi đọc file: {e}"
    return rec


def read_many(files: list[tuple[str, bytes]], progress=None) -> list[dict[str, Any]]:
    """Đọc nhiều file. PDF đọc song song bằng nhiều tiến trình (đọc PDF tốn CPU)."""
    import os
    from concurrent.futures import ProcessPoolExecutor

    items = [(n, d) for name, data in files for n, d in iter_files(name, data)
             if n.lower().endswith((".xml", ".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"))]
    if len(items) < 8:
        out = []
        for i, (n, d) in enumerate(items, 1):
            out.append(read_invoice(n, d))
            if progress:
                progress(i, len(items))
        return out
    workers = max(2, min(8, (os.cpu_count() or 4) - 1))
    out = [None] * len(items)
    with ProcessPoolExecutor(workers) as ex:
        for i, rec in enumerate(ex.map(_read_tuple, items, chunksize=4)):
            out[i] = rec
            if progress:
                progress(i + 1, len(items))
    return out  # type: ignore[return-value]


def _read_tuple(x: tuple[str, bytes]) -> dict[str, Any]:
    return read_invoice(*x)


# ------------------------------------------------------------------ kiểm tra với dữ liệu cổng thuế
def load_portal_index(root) -> dict[tuple[str, str, str], dict[str, Any]]:
    """Toàn bộ hóa đơn đã tải từ cổng (mọi tháng) -> khóa (MST bán, ký hiệu, số)."""
    import json
    from pathlib import Path
    idx: dict = {"_months": set()}
    for f in Path(root).glob("*/.dulieu_*.json"):
        idx["_months"].add(f.stem.replace(".dulieu_", ""))  # 'YYYY-MM' đã có dữ liệu
        for r in json.loads(f.read_text(encoding="utf-8")).get("rows", []):
            idx[(str(r.get("nbmst") or ""), str(r.get("khhdon") or "").upper(), str(r.get("shdon") or "").lstrip("0"))] = r
    return idx


def check_portal(rec: dict[str, Any], idx: dict) -> str:
    if not (rec.get("nbmst") and rec.get("khhdon") and rec.get("shdon")):
        return "Thiếu MST/ký hiệu/số – không so được"
    p = idx.get((rec["nbmst"], str(rec["khhdon"]).upper(), str(rec["shdon"]).lstrip("0")))
    if not p:
        ym = str(rec.get("ngay_lap") or "")[:7]
        if ym and ym not in idx.get("_months", set()):
            return f"Chưa có dữ liệu cổng tháng {ym[5:]}/{ym[:4]} – hãy tải hóa đơn tháng này trước"
        return "KHÔNG thấy trên cổng (trong dữ liệu đã tải)"
    from .reports import tthai_code
    code = tthai_code(p)
    if code in (4, 6):
        return f"Trên cổng: {p.get('tthai')}"
    a, b = rec.get("tgtttbso"), p.get("tgtttbso")
    if a is not None and b is not None and abs(float(a) - float(b)) > 1:
        return f"LỆCH tổng tiền (cổng: {float(b):,.0f})"
    return "Khớp với cổng thuế"


def write_invoice_report(dest, recs: list[dict[str, Any]], idx: dict):
    from openpyxl import Workbook

    from .reports import WARN, _header, _save
    wb = Workbook()
    ws = wb.active
    ws.title = "Hoa don doc duoc"
    cols = [("File", 34), ("Nguồn", 6), ("Mẫu số", 6), ("Ký hiệu", 9), ("Số HĐ", 9), ("Ngày", 11), ("MST bán", 14),
            ("Tên người bán", 34), ("MST mua", 14), ("Tên người mua", 34), ("Chưa thuế", 15), ("Thuế", 13),
            ("Tổng thanh toán", 15), ("Thuế suất", 9), ("Mã CQT", 22), ("Mã tra cứu", 18), ("Kiểm tra cổng thuế", 34),
            ("Lỗi đọc", 34)]
    _header(ws, cols)
    keys = ["file", "nguon", "khmshdon", "khhdon", "shdon", "ngay_lap", "nbmst", "nbten", "nmmst", "nmten",
            "tgtcthue", "tgtthue", "tgtttbso", "thue_suat", "mccqt", "ma_tra_cuu", "_check", "loi"]
    for r, rec in enumerate(recs, 2):
        rec["_check"] = check_portal(rec, idx)
        for c, k in enumerate(keys, 1):
            cell = ws.cell(r, c, rec.get(k))
            if k in ("tgtcthue", "tgtthue", "tgtttbso"):
                cell.number_format = "#,##0"
            if (k == "_check" and not str(rec["_check"]).startswith("Khớp")) or (k == "loi" and rec.get("loi")):
                cell.fill = WARN
    ws.auto_filter.ref = ws.dimensions
    return _save(wb, dest)
