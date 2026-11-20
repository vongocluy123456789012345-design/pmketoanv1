"""Chuyển hàng loạt hóa đơn (PDF / XML / ZIP / ảnh) vào FILE MẪU NHẬP KHẨU của MISA (mua vào hoặc bán ra).

Mỗi hóa đơn lấy dữ liệu theo thứ tự chính xác -> nhanh:
  1. XML hóa đơn (file tải lên, hoặc XML/JSON chi tiết đã tải từ cổng có trên máy) – chính xác 100%.
  2. PDF có lớp chữ: đọc phần đầu hóa đơn + bảng hàng hóa; chỉ nhận khi tổng các dòng KHỚP tổng tiền hàng.
  3. Còn lại (PDF bố cục lạ, scan/ảnh): AI đọc thông minh (Gemini, chạy song song) -> kiểm tra lại tổng.
Rồi ghép từng dòng hàng vào đúng cột của file mẫu (tự nhận cột theo tên tiêu đề; cột lạ nhờ AI đoán 1 lần).
"""
from __future__ import annotations

import io
import json
import re
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from .config import settings
from .docreader import fold, iter_files, ocr_read, parse_pdf_text, pdf_text, vn_num

# ------------------------------------------------------------------ nhận cột của file mẫu
# khóa -> các cụm từ (đã bỏ dấu, chữ thường) có thể xuất hiện trong tiêu đề cột của mẫu MISA
FIELDS: dict[str, list[str]] = {
    "ngay_hach_toan": ["ngay hach toan"],
    "ngay_chung_tu": ["ngay chung tu", "ngay phieu", "ngay nhap", "ngay xuat"],
    "so_chung_tu": ["so chung tu", "so phieu nhap", "so phieu xuat", "so phieu"],
    "ngay_hoa_don": ["ngay hoa don", "ngay hd"],
    "mau_so": ["mau so hoa don", "mau so hd", "ky hieu mau", "mau so"],
    "ky_hieu": ["ky hieu hoa don", "ky hieu hd", "ky hieu"],
    "so_hoa_don": ["so hoa don", "so hd"],
    "ma_doi_tuong": ["ma nha cung cap", "ma ncc", "ma khach hang", "ma kh", "ma doi tuong"],
    "ten_doi_tuong": ["ten nha cung cap", "ten ncc", "ten khach hang", "ten kh", "ten doi tuong", "nha cung cap",
                      "khach hang", "don vi ban", "don vi mua"],
    "dia_chi": ["dia chi"],
    "mst": ["ma so thue", "mst"],
    "nguoi_giao_nhan": ["nguoi giao hang", "nguoi nhan hang", "nguoi mua hang", "nguoi giao", "nguoi nhan"],
    "dien_giai": ["dien giai"],
    "ly_do_chi": ["ly do chi", "ly do thu", "noi dung thanh toan", "ly do"],
    "ma_hang": ["ma hang", "ma vat tu", "ma hhdv", "ma san pham"],
    "ten_hang": ["ten hang", "ten vat tu", "ten hhdv", "ten san pham", "dien giai hang", "ten hang hoa"],
    "dvt": ["don vi tinh", "dvt"],
    "so_luong": ["so luong"],
    "don_gia": ["don gia"],
    "thanh_tien": ["thanh tien", "tien hang"],
    "ty_le_ck": ["ty le ck", "ty le chiet khau", "% ck", "% chiet khau", "tl ck"],
    "tien_ck": ["tien chiet khau", "tien ck"],
    "thue_suat": ["% thue gtgt", "thue suat", "% thue", "ty le thue"],
    "tien_thue": ["tien thue gtgt", "tien thue", "thue gtgt"],
    "tong_tt": ["tong tien thanh toan", "tong thanh toan", "tong cong"],
    "tk_no": ["tk no"],
    "tk_co": ["tk co"],
    "tk_kho": ["tk kho", "tk chi phi"],
    "tk_cong_no": ["tk cong no", "tk phai tra", "tk phai thu"],
    "tk_doanh_thu": ["tk doanh thu"],
    "tk_thue": ["tk thue gtgt", "tk thue"],
    "ma_kho": ["ma kho", "kho"],
    "loai_tien": ["loai tien"],
    "ty_gia": ["ty gia"],
    "hinh_thuc_tt": ["hinh thuc thanh toan"],
    "pttt": ["phuong thuc thanh toan"],
    "hinh_thuc_mua": ["hinh thuc mua hang", "hinh thuc ban hang"],
    "nhan_kem_hd": ["nhan kem hoa don", "kem hoa don"],
    "ma_cqt": ["ma cqt", "ma co quan thue"],
    "ma_tra_cuu": ["ma tra cuu"],
    "nhom_hhdv": ["nhom hhdv", "nhom hang hoa dich vu mua vao"],
}
# Cột dễ bị nhận nhầm (chứa cụm từ của cột khác) -> KHÔNG điền gì, để trống như mẫu
_SKIP_HEADER = re.compile(r"kem theo|thue suat khac|ghi no|so chung tu thanh toan|so tai khoan|ngan hang|nhan vien|"
                          r"han thanh toan|giu ho|ban ho|chi phi san xuat|khoan muc|ma don vi|thanh pham|"
                          r"don dat hang|don mua hang|ma thong ke|duong dan|quy doi|dong ghi chu|kiem phieu|gia von|"
                          r"tien von|cach lay don gia|khong cap nhat gia|don hang online|gian hang|san kenh|thue gtgt tu san|"
                          r"thue tncn tu san|hang khuyen mai|cp khong hop ly|giam 20%|so phieu xuat|ngay phieu nhap")
ACCOUNTS = {"purchase": {"tk_kho": "1561", "tk_cong_no": "331", "tk_thue": "1331"},
            "sold": {"tk_doanh_thu": "5111", "tk_cong_no": "131", "tk_thue": "33311", "tk_gia_von": "632", "tk_kho": "1561"}}


def _norm(s: Any) -> str:
    s = fold(str(s or "")).lower().replace("(*)", "").replace("*", "")
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9%/ ]", " ", s)).strip()


def map_header(title: str) -> str | None:
    t = _norm(title)
    if not t or _SKIP_HEADER.search(t):
        return None
    best, blen = None, 0
    for key, pats in FIELDS.items():
        for p in pats:
            if (t == p or re.search(rf"(^| ){re.escape(p)}( |$)", t)) and len(p) > blen:
                # tránh nhầm: "thanh tien" trong "thanh tien quy doi", "don gia" trong "don gia sau thue"
                if key in ("thanh_tien", "don_gia") and re.search(r"quy doi|sau thue|thue", t.replace(p, "")):
                    continue
                best, blen = key, len(p)
    return best


# ------------------------------------------------------------------ mở file mẫu ở mọi định dạng
def _sheet_title(name: Any, used: set[str]) -> str:
    t = re.sub(r"[\[\]\*\?/\\:]", " ", str(name or "Sheet")).strip()[:31] or "Sheet"
    base, i = t, 2
    while t in used:
        t = f"{base[:27]} ({i})"
        i += 1
    used.add(t)
    return t


def _book_from_tables(tables: list[tuple[str, list[list[Any]]]]):
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    for name, rows in tables or [("Sheet1", [])]:
        ws = wb.create_sheet(_sheet_title(name, used))
        for r, row in enumerate(rows, 1):
            for c, v in enumerate(row, 1):
                if v not in (None, ""):
                    ws.cell(r, c, v)
    return wb


def _xls_tables(data: bytes):
    import xlrd
    book = xlrd.open_workbook(file_contents=data)
    out = []
    for sh in book.sheets():
        rows = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                v = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        v = xlrd.xldate.xldate_as_datetime(v, book.datemode)
                    except Exception:  # noqa: BLE001
                        pass
                elif cell.ctype == xlrd.XL_CELL_NUMBER and float(v).is_integer():
                    v = int(v)
                row.append(v)
            rows.append(row)
        out.append((sh.name, rows))
    return out


def _ods_tables(data: bytes):
    import zipfile
    import xml.etree.ElementTree as ET
    ns = {"table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
          "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
          "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0"}
    q = lambda p, n: f"{{{ns[p]}}}{n}"  # noqa: E731
    root = ET.fromstring(zipfile.ZipFile(io.BytesIO(data)).read("content.xml"))
    out = []
    for t in root.iter(q("table", "table")):
        rows = []
        for tr in t.iter(q("table", "table-row")):
            row: list[Any] = []
            for tc in tr:
                if tc.tag not in (q("table", "table-cell"), q("table", "covered-table-cell")):
                    continue
                typ = tc.get(q("office", "value-type"))
                if typ in ("float", "percentage", "currency"):
                    v: Any = float(tc.get(q("office", "value")) or 0)
                    v = int(v) if v.is_integer() else v
                elif typ == "date":
                    v = tc.get(q("office", "date-value"))
                else:
                    v = "\n".join("".join(p.itertext()) for p in tc.iter(q("text", "p"))) or None
                n = int(tc.get(q("table", "number-columns-repeated")) or 1)
                row += [v] * (n if v not in (None, "") else min(n, 1))
            n = int(tr.get(q("table", "number-rows-repeated")) or 1)
            rows += [row] * (n if any(x not in (None, "") for x in row) else min(n, 1))
        while rows and not any(x not in (None, "") for x in rows[-1]):
            rows.pop()
        out.append((t.get(q("table", "name")) or "Sheet", rows))
    return out


def _spreadsheetml_tables(text: str):
    """File "XML Spreadsheet 2003" (nhiều phần mềm xuất ra với đuôi .xls)."""
    import xml.etree.ElementTree as ET
    ss = "urn:schemas-microsoft-com:office:spreadsheet"
    root = ET.fromstring(re.sub(r"^\s*<\?xml[^>]*\?>", "", text))  # đã giải mã sẵn -> bỏ khai báo encoding
    out = []
    for w in root.iter(f"{{{ss}}}Worksheet"):
        rows = []
        for r in w.iter(f"{{{ss}}}Row"):
            ri = r.get(f"{{{ss}}}Index")
            while ri and len(rows) < int(ri) - 1:
                rows.append([])
            row: list[Any] = []
            for c in r.iter(f"{{{ss}}}Cell"):
                ci = c.get(f"{{{ss}}}Index")
                while ci and len(row) < int(ci) - 1:
                    row.append(None)
                d = c.find(f"{{{ss}}}Data")
                v: Any = "".join(d.itertext()) if d is not None else None
                if d is not None and d.get(f"{{{ss}}}Type") == "Number":
                    try:
                        v = float(v)
                        v = int(v) if v.is_integer() else v
                    except (TypeError, ValueError):
                        pass
                row.append(v)
            rows.append(row)
        out.append((w.get(f"{{{ss}}}Name") or "Sheet", rows))
    return out


def _html_tables(text: str):
    """File "xls" thật ra là bảng HTML (MISA, phần mềm khác hay xuất kiểu này)."""
    import lxml.html
    doc = lxml.html.fromstring(text)
    out = []
    for i, t in enumerate(doc.iter("table"), 1):
        rows = []
        for tr in t.iter("tr"):
            row: list[Any] = []
            for td in tr:
                if td.tag not in ("td", "th"):
                    continue
                v = " ".join(td.text_content().split()) or None
                row += [v] * max(1, int(td.get("colspan") or 1) if str(td.get("colspan") or "1").isdigit() else 1)
            rows.append(row)
        if rows:
            out.append((f"Sheet{i}", rows))
    return out


def _csv_tables(text: str):
    import csv
    sample = text[:20000]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel_tab if sample.count("\t") > sample.count(",") else csv.excel
    rows = [list(r) for r in csv.reader(io.StringIO(text), dialect)]
    return [("Sheet1", rows)]


def _decode(data: bytes) -> str:
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return data.decode("utf-16")
    for enc in ("utf-8-sig", "cp1258", "cp1252"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1")


def load_any_workbook(data: bytes, values: bool = False):
    """Mở file mẫu ở BẤT KỲ định dạng bảng tính thường gặp -> workbook openpyxl:
    .xlsx/.xlsm/.xltx/.xltm, .xls (Excel 97-2003), .ods (LibreOffice/OpenOffice), .csv/.txt,
    và file "xls" thực chất là HTML hoặc XML Spreadsheet 2003."""
    from openpyxl import load_workbook
    if not data:
        raise ValueError("File mẫu rỗng")
    if data[:2] == b"PK":
        import zipfile
        try:
            names = zipfile.ZipFile(io.BytesIO(data)).namelist()
        except zipfile.BadZipFile:
            names = []
        if "content.xml" in names and not any(n.startswith("xl/") for n in names):
            return _book_from_tables(_ods_tables(data))
        wb = load_workbook(io.BytesIO(data), data_only=values)   # values=True: đọc kết quả công thức
        if wb.template:  # .xltx/.xltm -> lưu ra dạng sổ thường
            wb.template = False
        return wb
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        try:
            return _book_from_tables(_xls_tables(data))
        except Exception as e:  # noqa: BLE001
            raise ValueError(f"Không đọc được file Excel .xls: {e}") from e
    text = _decode(data)
    head = text[:3000].lower()
    if "urn:schemas-microsoft-com:office:spreadsheet" in head:
        return _book_from_tables(_spreadsheetml_tables(text))
    if "<table" in text[:200000].lower() or "<html" in head:
        tables = _html_tables(text)
        if tables:
            return _book_from_tables(tables)
    if "\x00" in text[:1000]:
        raise ValueError("Không nhận ra định dạng file mẫu (hỗ trợ .xlsx, .xlsm, .xls, .ods, .csv, .txt, .xml, .html)")
    return _book_from_tables(_csv_tables(text))


def read_template(data: bytes, values: bool = False):
    """-> (workbook, sheet, dòng tiêu đề, {cột: khóa}, [tiêu đề chưa nhận được])"""
    wb = load_any_workbook(data, values)
    best = None
    for ws in wb.worksheets:
        for r in range(1, min(ws.max_row, 20) + 1):
            cols = {c.column: map_header(c.value) for c in ws[r] if isinstance(c.value, str) and c.value.strip()}
            hits = sum(1 for v in cols.values() if v)
            if hits >= 3 and (best is None or hits > best[0]):
                best = (hits, ws, r, cols)
    if not best:
        raise ValueError("Không nhận ra dòng tiêu đề cột trong file mẫu (cần các cột như Ngày hóa đơn, Số hóa đơn, "
                         "Tên hàng, Số lượng, Đơn giá, Thành tiền…)")
    _, ws, hr, cols = best
    unknown = [ws.cell(hr, c).value for c, k in cols.items() if not k]
    # 2 cột cùng nghĩa (vd "Số chứng từ (*)" và "Số phiếu nhập"): giữ cột bắt buộc (*) – không thì cột đầu tiên
    keep: dict[str, int] = {}
    for c, k in sorted(cols.items()):
        if not k:
            continue
        if k not in keep or ("(*)" in str(ws.cell(hr, c).value) and "(*)" not in str(ws.cell(hr, keep[k]).value)):
            keep[k] = c
    return wb, ws, hr, {c: k for k, c in sorted(keep.items(), key=lambda kv: kv[1])}, unknown


# ------------------------------------------------------------------ đọc hóa đơn
def _num(v: Any) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def _rate_pct(ts: Any) -> float | None:
    m = re.search(r"(\d+(?:[.,]\d+)?)", str(ts or ""))
    if not m:
        return None
    v = float(m.group(1).replace(",", "."))
    return v * 100 if v < 1 and "%" not in str(ts) and v else v


def from_xml(data: bytes) -> dict[str, Any]:
    from .gdt.xml_parser import parse_invoice_xml
    p = parse_invoice_xml(data)
    nb, nm = p.get("nguoi_ban") or {}, p.get("nguoi_mua") or {}
    items = []
    for h in p.get("hang_hoa") or []:
        tc = str(h.get("tinh_chat") or "1")
        if tc == "4":
            continue
        tt = _num(h.get("thanh_tien")) * (-1 if tc == "3" else 1)
        pct = _rate_pct(h.get("thue_suat"))
        items.append({"ma": h.get("ma"), "ten": h.get("ten"), "dvt": h.get("dvt"), "sl": h.get("so_luong"),
                      "dg": h.get("don_gia"), "tt": tt, "ck": h.get("chiet_khau"), "ts": h.get("thue_suat"),
                      "thue": round(tt * pct / 100) if pct is not None else 0})
    return {"khmshdon": p.get("khmshdon"), "khhdon": p.get("khhdon"), "shdon": str(int(p["shdon"])) if str(p.get("shdon") or "").isdigit() else p.get("shdon"),
            "ngay_lap": (p.get("ngay_lap") or "")[:10], "nbmst": nb.get("mst"), "nbten": nb.get("ten"), "nbdchi": nb.get("dia_chi"),
            "nmmst": nm.get("mst"), "nmten": nm.get("ten") or nm.get("ho_ten_nguoi_mua"), "nmdchi": nm.get("dia_chi"),
            "tgtcthue": p.get("tong_chua_thue"), "tgtthue": p.get("tong_thue"), "tgtttbso": p.get("tong_thanh_toan"),
            "htttoan": p.get("htttoan"), "mccqt": p.get("mccqt"), "ma_tra_cuu": p.get("lookup_code"), "items": items,
            "nguon": "XML"}


def from_portal_json(d: dict[str, Any]) -> dict[str, Any]:
    items = []
    for h in d.get("hdhhdvu") or []:
        if h.get("tchat") == 4:
            continue
        tt = _num(h.get("thtien")) * (-1 if h.get("tchat") == 3 else 1)
        items.append({"ma": h.get("mhhdvu"), "ten": h.get("ten"), "dvt": h.get("dvtinh"), "sl": h.get("sluong"),
                      "dg": h.get("dgia"), "tt": tt, "ck": h.get("stckhau"), "ts": h.get("ltsuat") or h.get("tsuat"),
                      "thue": _num(h.get("tthue")) if h.get("tthue") is not None else 0})
    return {"items": items, "nbdchi": d.get("nbdchi"), "nmdchi": d.get("nmdchi"), "htttoan": d.get("thtttoan")}


_ITEM_COLS = {"ten": ["ten hang", "ten hhdv", "ten hang hoa", "hang hoa", "dien giai", "description", "ten san pham",
                      "noi dung"],
              "dvt": ["don vi tinh", "dvt", "unit"], "sl": ["so luong", "quantity", "sl"],
              "dg": ["don gia", "unit price", "price"], "tt": ["thanh tien", "amount", "tien hang"],
              "ts": ["thue suat", "% thue", "vat rate", "ts"], "thue": ["tien thue", "vat amount"],
              "ma": ["ma hang", "ma hhdv", "ma sp", "code"], "ck": ["chiet khau", "discount"]}


def items_from_pdf_tables(data: bytes) -> list[dict[str, Any]]:
    """Đọc bảng hàng hóa trong PDF có lớp chữ (theo dòng tiêu đề bảng: Tên hàng, ĐVT, Số lượng, Đơn giá, Thành tiền…)."""
    import pdfplumber
    items: list[dict[str, Any]] = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages[:4]:
            for tb in page.extract_tables() or []:
                hdr_i, cmap = None, {}
                for i, row in enumerate(tb[:6]):
                    m = {}
                    for j, cell in enumerate(row):
                        t = _norm(cell)
                        for k, pats in _ITEM_COLS.items():
                            if k not in m.values() and any(p == t or t.startswith(p) or f" {p}" in f" {t}" for p in pats):
                                m[j] = k
                                break
                    if "ten" in m.values() and "tt" in m.values():
                        hdr_i, cmap = i, m
                        break
                if hdr_i is None:
                    continue
                for row in tb[hdr_i + 1:]:
                    rec = {k: (row[j] if j < len(row) else None) for j, k in cmap.items()}
                    name = re.sub(r"\s+", " ", str(rec.get("ten") or "")).strip()
                    tt = vn_num(rec.get("tt"))
                    if not name or tt is None or re.fullmatch(r"[\d\s()=x+\-*/.,ABC]*", name):
                        continue  # dòng số thứ tự cột (1)(2)(3)… hoặc dòng trống
                    if re.match(r"(?i)(cong tien|tong cong|tong tien|thue suat|tien thue)", fold(name)):
                        continue
                    ts = rec.get("ts")
                    items.append({"ma": rec.get("ma"), "ten": name, "dvt": rec.get("dvt"), "sl": vn_num(rec.get("sl")),
                                  "dg": vn_num(rec.get("dg")), "tt": tt, "ck": vn_num(rec.get("ck")),
                                  "ts": (str(ts).strip() if ts else None), "thue": vn_num(rec.get("thue"))})
    return items


_NUM = r"(-?\d[\d.,]*)"
_ROW_FULL = re.compile(r"^\s*(\d{1,3})\s+(.+?)\s+(\S+)\s+" + _NUM + r"\s+" + _NUM + r"\s+" + _NUM + r"\s*$")
_ROW_NUMS = re.compile(r"^\s*(\d{1,3})\s+(\S+(?:\s+\S+)?)\s+" + _NUM + r"\s+" + _NUM + r"\s+" + _NUM + r"\s*$")
_ROW_NODVT = re.compile(r"^\s*(\d{1,3})\s+" + _NUM + r"\s+" + _NUM + r"\s+" + _NUM + r"\s*$")   # không có ĐVT


_PAGE_NOISE = re.compile(r"trang\s*:|tiep theo|tra cuu|giai phap|cung cap|website|dia chi|ho kinh doanh|so tai khoan|"
                         r"hoa don ban hang|ky hieu|\(day\)|ma co quan|ho ten nguoi|ten don vi|ma so thue|can cuoc|"
                         r"hinh thuc|^stt|\(no\.\)|goods, services|mst ")


def _name_lines(raw: list[str], folded: list[str]) -> list[str]:
    """Dòng chữ thuộc tên hàng trong khoảng giữa 2 dòng số: bỏ chân trang / đầu trang sau khi sang trang
    (từ "Trang: 1/3" tới dòng số cột "1 2 3 4 5 6=4x5" của trang sau)."""
    out, skip = [], False
    for ln, fl in zip(raw, folded):
        if not ln:
            continue
        if "=4x5" in fl.replace(" ", ""):
            skip = False
            continue
        if re.search(r"trang\s*:|tiep theo", fl):
            skip = True
        if skip or _PAGE_NOISE.search(fl):
            continue
        out.append(ln)
    return out


def items_from_ocr_text(text: str) -> list[dict[str, Any]]:
    """Đọc bảng hàng hóa từ chữ OCR (hóa đơn scan / máy tính tiền) KHÔNG cần AI.
    Dòng hàng: "STT [tên hàng] ĐVT SL Đơn giá Thành tiền" (SL x Đơn giá = Thành tiền). Tên hàng dài bị tách lên / xuống
    dòng số (căn giữa ô) -> ghép các dòng chữ quanh dòng số; khoảng giữa 2 dòng hàng chia đôi cho 2 hàng."""
    from .docreader import _ocr_clean
    lines = [ln.strip() for ln in _ocr_clean(text).splitlines()]
    f = [fold(ln).lower() for ln in lines]
    start = next((i for i, x in enumerate(f) if "=4x5" in x.replace(" ", "")), None)      # dòng số cột (1)(2)..6=4x5
    if start is None:
        start = next((i for i, x in enumerate(f) if "(amount)" in x or "thanh tien" in x), None)
    if start is None:
        start = next((i for i, x in enumerate(f) if x.startswith("stt")), None)
    if start is None:
        return []
    end = next((i for i in range(start + 1, len(f)) if re.search(r"tong cong|cong tien|total amount|tong tien", f[i])), len(f))
    rows: list[tuple[int, dict[str, Any]]] = []
    for i in range(start + 1, end):
        ln = lines[i]
        nd = _ROW_NODVT.match(ln)
        m = None if nd else _ROW_NUMS.match(ln)
        full = None if (nd or m) else _ROW_FULL.match(ln)
        if nd:
            stt, sl, dg, tt = nd.group(1), vn_num(nd.group(2)), vn_num(nd.group(3)), vn_num(nd.group(4))
        elif m or full:
            mm = m or full
            stt, sl, dg, tt = mm.group(1), vn_num(mm.group(m and 3 or 4)), vn_num(mm.group(m and 4 or 5)), vn_num(mm.group(m and 5 or 6))
        else:
            continue
        if not (sl and dg and tt) or abs(sl * dg - tt) > max(2, abs(tt) * 0.01):
            continue
        if nd:
            name, dvt = "", ""
        elif m:
            name, dvt = "", m.group(2)
        else:
            name, dvt = full.group(2), full.group(3)
        rows.append((i, {"stt": int(stt), "ten": name, "dvt": dvt, "sl": sl, "dg": dg, "tt": tt}))
    if not rows:
        return []
    # ghép tên hàng cho dòng chỉ có số: chữ ở khoảng trống trên / dưới dòng số
    idxs = [i for i, _ in rows]
    inline = [bool(r["ten"]) for _, r in rows]     # dòng có sẵn tên ngay trên dòng số (không chia khoảng trống)
    for k, (i, r) in enumerate(rows):
        if r["ten"]:
            continue
        prev_i = idxs[k - 1] if k else start
        next_i = idxs[k + 1] if k + 1 < len(rows) else end
        above = _name_lines(lines[prev_i + 1:i], f[prev_i + 1:i])
        below = _name_lines(lines[i + 1:next_i], f[i + 1:next_i])
        if k and not inline[k - 1]:             # khoảng giữa với hàng trước (cũng tách tên): mỗi hàng nửa khoảng
            above = above[len(above) // 2:]
        if k + 1 < len(rows) and not inline[k + 1]:
            below = below[: (len(below) + 1) // 2]
        r["ten"] = re.sub(r"\s+", " ", " ".join(above + below)).strip(" -")
    return [{"ma": None, "ten": r["ten"], "dvt": r["dvt"], "sl": r["sl"], "dg": r["dg"], "tt": r["tt"], "ck": None,
             "ts": None, "thue": None} for _, r in rows if r["ten"]]


def items_ok(inv: dict[str, Any]) -> bool:
    """Tổng thành tiền các dòng phải khớp tổng tiền hàng của hóa đơn (sai số 0,5% hoặc 1.000đ)."""
    items, total = inv.get("items") or [], inv.get("tgtcthue")
    if not items or total is None:
        return False
    s = sum(_num(i.get("tt")) for i in items)
    return abs(s - float(total)) <= max(1000, abs(float(total)) * 0.005)


AI_PROMPT = """Bạn trích xuất dữ liệu hóa đơn GTGT Việt Nam từ văn bản (đọc từ PDF/ảnh). Trả về DUY NHẤT 1 JSON:
{"khmshdon":"1","khhdon":"C26TAA","shdon":"123","ngay_lap":"YYYY-MM-DD","nbmst":"","nbten":"","nbdchi":"","nmmst":"","nmten":"",
"nmdchi":"","htttoan":"","tgtcthue":0,"tgtthue":0,"tgtttbso":0,"mccqt":"","ma_tra_cuu":"",
"items":[{"ma":"","ten":"","dvt":"","sl":0,"dg":0,"tt":0,"ck":0,"ts":"10%","thue":0}]}
Quy tắc: số tiền là số (không dấu chấm phân cách nghìn); tt = thành tiền chưa thuế từng dòng; ts dạng "8%", "10%",
"5%", "0%", "KCT" hoặc "KKKNT"; dòng chiết khấu thương mại ghi tt âm; bỏ dòng tổng cộng; không bịa – thiếu thì để rỗng/0."""


def ai_extract(text: str) -> dict[str, Any]:
    from . import assistant
    if not (assistant.key_for("gemini") or assistant.provider() != "gemini"):
        raise RuntimeError("Chưa có khóa AI")
    p = "gemini" if assistant.key_for("gemini") else assistant.provider()
    models = ["gemini-flash-lite-latest", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-flash-latest"] \
        if p == "gemini" else [assistant.model_for(p)]
    last = None
    for m in models:
        try:
            out = assistant._post_plain(p, m, [{"role": "system", "content": AI_PROMPT},
                                               {"role": "user", "content": text[:14000]}])
            js = re.search(r"\{.*\}", out, re.S)
            if js:
                d = json.loads(js.group(0))
                d["nguon"] = f"AI ({m})"
                return d
        except Exception as e:  # noqa: BLE001 - hết lượt / quá tải -> model kế tiếp
            last = e
    raise RuntimeError(f"AI không đọc được: {last}")


def local_index() -> dict[tuple[str, str, str], Path]:
    """Hóa đơn đã tải từ cổng trên máy: (MST bán, ký hiệu, số) -> file XML/JSON chi tiết."""
    idx: dict[tuple[str, str, str], Path] = {}
    base = Path(settings.output_dir)
    for f in list(base.glob("*/*/*/XML/*.xml")) + list(base.glob("*/*/*/XML/*.json")):
        m = re.match(r"(\d)?([CK]\d{2}[A-Z0-9]{3})_0*(\d+)_(.+)$", f.stem)
        if m:
            k = (m.group(4).replace("-", ""), m.group(2), m.group(3))
            if k not in idx or f.suffix == ".xml":
                idx[k] = f
    return idx


def read_one(name: str, data: bytes, idx: dict) -> dict[str, Any]:
    """Đọc 1 hóa đơn KHÔNG dùng AI. Trả về dict có 'items' (có thể rỗng nếu cần AI) + '_text' để AI đọc tiếp."""
    low = name.lower()
    inv: dict[str, Any] = {"file": name}
    if low.endswith(".xml") or data.lstrip()[:5] in (b"<?xml", b"<HDon"):
        return inv | from_xml(data)
    if low.endswith(".pdf") or data[:5] == b"%PDF-":
        text = pdf_text(data)
        fields = None
        if len(text.strip()) < 40:
            fields, text = ocr_read(data)
    else:
        from . import ocr
        if not ocr.is_image(name, data):
            return inv | {"loi": "Định dạng chưa hỗ trợ", "items": []}
        fields, text = ocr_read(data)
    inv.update(fields if fields is not None else parse_pdf_text(text, "PDF"))
    inv["_text"] = text
    # đã có XML/chi tiết của hóa đơn này trên máy (tải từ cổng) -> dùng dòng hàng chính xác
    k = (str(inv.get("nbmst") or "").replace("-", ""), str(inv.get("khhdon") or "").upper(),
         str(inv.get("shdon") or "").lstrip("0"))
    f = idx.get(k)
    if f:
        if f.suffix == ".xml":
            x = from_xml(f.read_bytes())
            return inv | {k2: v for k2, v in x.items() if v not in (None, "", [])} | {"nguon": "XML trên máy"}
        d = from_portal_json(json.loads(f.read_text(encoding="utf-8")))
        if d["items"]:
            return inv | {k2: v for k2, v in d.items() if v} | {"nguon": "Dữ liệu cổng trên máy"}
    if fields is None:  # PDF có lớp chữ -> đọc bảng hàng hóa
        try:
            inv["items"] = items_from_pdf_tables(data)
        except Exception:  # noqa: BLE001
            inv["items"] = []
        if items_ok(inv):
            inv["nguon"] = "PDF (bảng hàng hóa)"
            return inv
    else:  # bản scan / ảnh: đọc bảng hàng từ chữ OCR (nhanh, không cần AI) – chỉ nhận khi tổng khớp
        try:
            inv["items"] = items_from_ocr_text(text)
        except Exception:  # noqa: BLE001
            inv["items"] = []
        if items_ok(inv):
            inv["nguon"] = "OCR (bảng hàng hóa)"
            return inv
    inv["items"] = []
    return inv


# ------------------------------------------------------------------ ghi vào mẫu
def build_rows(inv: dict[str, Any], direction: str, seq: int, acc: dict[str, str]) -> list[dict[str, Any]]:
    from datetime import date
    purchase = direction == "purchase"
    partner = {"mst": inv.get("nbmst"), "ten": inv.get("nbten"), "dchi": inv.get("nbdchi")} if purchase else \
        {"mst": inv.get("nmmst"), "ten": inv.get("nmten"), "dchi": inv.get("nmdchi")}
    ngay = inv.get("ngay_lap") or ""
    try:
        d = date.fromisoformat(ngay)
        ngay_vn = d
    except ValueError:
        ngay_vn = ngay
    so_ct = f"{'NK' if purchase else 'BH'}{seq:05d}"
    dien_giai = (f"Mua hàng của {partner['ten'] or ''} theo HĐ {inv.get('khhdon') or ''}-{inv.get('shdon') or ''}"
                 if purchase else f"Bán hàng cho {partner['ten'] or ''} theo HĐ {inv.get('khhdon') or ''}-{inv.get('shdon') or ''}")
    items = inv.get("items") or [{"ten": dien_giai, "tt": inv.get("tgtcthue"), "thue": inv.get("tgtthue"),
                                  "ts": inv.get("thue_suat")}]
    a = ACCOUNTS[direction] | {k: v for k, v in (acc or {}).items() if v}
    # thuế suất chung của hóa đơn (PDF 1 thuế suất thường chỉ ghi ở cuối): theo nhãn, hoặc suy ra thuế / tiền hàng
    inv_pct = _rate_pct(inv.get("thue_suat")) if re.fullmatch(r"\s*\d+(?:[.,]\d+)?\s*%?\s*", str(inv.get("thue_suat") or "x")) else None
    if inv_pct is None and _num(inv.get("tgtcthue")):
        r = round(_num(inv.get("tgtthue")) / _num(inv.get("tgtcthue")) * 100, 1)
        inv_pct = next((p for p in (0, 5, 8, 10) if abs(r - p) <= 0.15), None)
    taxes = []
    for it in items:
        pct = _rate_pct(it.get("ts"))
        if pct is None and not str(it.get("ts") or "").strip():
            pct = inv_pct
        tt = _num(it.get("tt"))
        thue = it.get("thue")
        if thue in (None, "") or (not thue and pct):
            thue = round(tt * pct / 100) if pct is not None else 0
        taxes.append([it, pct, tt, thue])
    # làm tròn: tổng thuế các dòng phải bằng tổng thuế trên hóa đơn
    diff = round(_num(inv.get("tgtthue")) - sum(_num(x[3]) for x in taxes)) if inv.get("tgtthue") is not None else 0
    if taxes and diff and abs(diff) <= max(10, len(taxes) * 2):
        max(taxes, key=lambda x: abs(x[2]))[3] += diff
    # Thông tin chứng từ theo danh sách chọn của mẫu MISA (ghi chú từng cột):
    #   Hình thức mua hàng: nhập kho (cần Mã kho) / không qua kho;  Phương thức thanh toán: Chưa thanh toán / Tiền mặt /
    #   Ủy nhiệm chi;  Nhận kèm hóa đơn: GTGT (mẫu số 1) / bán hàng (mẫu số 2).
    ma_kho = str(acc.get("ma_kho") or "").strip() if acc else ""
    ht = str((acc or {}).get("hinh_thuc") or "").strip()
    nhap_kho = ht == "nhap_kho" or (not ht and bool(ma_kho))
    if purchase:
        hinh_thuc = "Mua hàng trong nước nhập kho" if nhap_kho else "Mua hàng trong nước không qua kho"
    else:
        hinh_thuc = None  # mẫu bán hàng: lấy theo danh sách chọn của mẫu
    pttt = {"tm": "Tiền mặt", "unc": "Ủy nhiệm chi"}.get(str((acc or {}).get("pttt") or ""), "Chưa thanh toán")
    if not purchase:
        pttt = {"Tiền mặt": "Thu tiền ngay - Tiền mặt", "Ủy nhiệm chi": "Thu tiền ngay - Chuyển khoản"}.get(pttt, "Chưa thu tiền")
    paid = not pttt.startswith("Chưa")
    mau = str(inv.get("khmshdon") or "1").strip()
    nhan_kem = ("Nhận kèm hóa đơn bán hàng" if mau == "2" else "Nhận kèm hóa đơn GTGT") if purchase else "Có"  # bán ra: "Lập kèm hóa đơn" Có/Không
    mtc = str(inv.get("ma_tra_cuu") or "").strip()
    mtc = mtc if re.fullmatch(r"[A-Za-z0-9_\-]{5,60}", mtc) else None  # mã đọc từ PDF lỗi -> để trống
    mst_doi_tac = str(partner["mst"] or "").strip() or None
    rows = []
    for it, pct, tt, thue in taxes:
        ts = it.get("ts")
        if pct is not None:
            ts_txt = f"{pct:g}"   # mẫu MISA: 10 / 8 / 5 / 0 (không có dấu %)
        else:
            f = fold(str(ts or "")).upper().replace(" ", "")
            ts_txt = "KCT" if "KCT" in f or "KHONGCHIU" in f else "KKKNT" if "KKKNT" in f or "KHONGKEKHAI" in f else \
                ("KHAC" if f else None)
        tk_mua = a.get("tk_kho")
        if purchase:  # định khoản theo mặt hàng: đã học từ cách kế toán hạch toán -> từ khóa (vận chuyển, điện nước...) -> mặc định
            try:
                from .tienich import suggest_tk
                tk_mua = suggest_tk(it.get("ten") or "", float(it.get("dg") or 0), partner["ten"] or "",
                                    mac_dinh=a.get("tk_kho") or "1561", tai_san=False)[0]
            except Exception:  # noqa: BLE001
                pass
        rows.append({
            "hinh_thuc_mua": hinh_thuc, "pttt": pttt, "nhan_kem_hd": nhan_kem,
            "ly_do_chi": ("Chi tiền " + dien_giai[0].lower() + dien_giai[1:]) if paid and purchase else (dien_giai if paid else None),
            "ma_kho": (ma_kho or None) if (nhap_kho or not purchase) else None,
            "ngay_hach_toan": ngay_vn, "ngay_chung_tu": ngay_vn, "so_chung_tu": so_ct, "ngay_hoa_don": ngay_vn,
            "mau_so": mau, "ky_hieu": inv.get("khhdon"),
            "so_hoa_don": str(inv.get("shdon")) if inv.get("shdon") not in (None, "") else None,
            "ma_doi_tuong": mst_doi_tac, "ten_doi_tuong": partner["ten"], "dia_chi": partner["dchi"],
            "mst": mst_doi_tac, "dien_giai": dien_giai, "ma_hang": it.get("ma"), "ten_hang": it.get("ten"),
            "dvt": it.get("dvt"), "so_luong": it.get("sl"), "don_gia": it.get("dg"), "thanh_tien": tt,
            "tien_ck": it.get("ck") or None, "thue_suat": ts_txt,
            "tien_thue": thue, "tong_tt": inv.get("tgtttbso"),
            "tk_no": tk_mua if purchase else a["tk_cong_no"], "tk_co": a["tk_cong_no"] if purchase else a["tk_doanh_thu"],
            "tk_kho": tk_mua if purchase else a.get("tk_kho"), "tk_cong_no": a["tk_cong_no"], "tk_doanh_thu": a.get("tk_doanh_thu"),
            "tk_thue": a["tk_thue"], "loai_tien": "VND", "ty_gia": 1, "hinh_thuc_tt": inv.get("htttoan"),
            "ma_cqt": inv.get("mccqt"), "ma_tra_cuu": mtc, "nhom_hhdv": "1" if purchase else None,
        })
    return rows


# ------------------------------------------------------------------ danh sách chọn (ghi chú cột) của mẫu
def template_lists(ws, hr: int, cols: dict[int, str]) -> dict[int, list[str]]:
    """Đọc danh sách giá trị hợp lệ (Data Validation – 'ghi chú' chọn giá trị) của từng cột ở dòng dữ liệu đầu tiên."""
    from openpyxl.utils import range_boundaries
    out: dict[int, list[str]] = {}
    try:
        dvs = list(ws.data_validations.dataValidation)
    except Exception:  # noqa: BLE001
        return out
    for dv in dvs:
        if dv.type != "list" or not dv.formula1:
            continue
        f1 = str(dv.formula1).strip()
        if f1.startswith('"') and f1.endswith('"'):
            opts = [x.strip() for x in re.split(r"[,;]", f1[1:-1]) if x.strip()]
        else:  # danh sách nằm ở vùng ô khác (=Sheet!$A$1:$A$5)
            opts = []
            try:
                ref = f1.lstrip("=")
                sh, rng = (ref.rsplit("!", 1) + [""])[:2] if "!" in ref else (ws.title, ref)
                src = ws.parent[sh.strip("'")]
                c1, r1, c2, r2 = range_boundaries(rng.replace("$", ""))
                opts = [str(src.cell(r, c).value).strip() for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)
                        if src.cell(r, c).value not in (None, "")]
            except Exception:  # noqa: BLE001
                continue
        for rg in str(dv.sqref).split():
            try:
                c1, r1, c2, r2 = range_boundaries(rg)
            except ValueError:
                continue
            for c in cols:
                if c1 <= c <= c2 and r1 <= hr + 1 <= (r2 or r1):
                    out[c] = opts
    return out


def _fit_option(v: Any, opts: list[str]) -> Any:
    """Đưa giá trị về đúng 1 lựa chọn trong danh sách của mẫu (so không dấu, không phân biệt hoa thường)."""
    if v is None or v == "" or not opts:
        return v
    fv = _norm(v)
    for o in opts:
        if _norm(o) == fv:
            return o
    if re.fullmatch(r"\d+(\.\d+)?", fv):  # thuế suất 8 / 8% / 0.08
        for o in opts:
            if re.fullmatch(r"\d+(\.\d+)?%?", _norm(o).replace(" ", "")) and float(_norm(o).rstrip("%")) == float(fv):
                return o
    words = set(fv.split())
    best, score = None, 0.0
    for o in opts:
        ow = set(_norm(o).split())
        s = len(words & ow) / max(1, len(words | ow))
        if s > score:
            best, score = o, s
    return best if score >= 0.34 else v


def apply_lists(rows: list[dict[str, Any]], cols: dict[int, str], lists: dict[int, list[str]]) -> list[str]:
    """Sửa giá trị các cột có danh sách chọn cho đúng mẫu. Cột chọn bắt buộc (hình thức, phương thức, kèm hóa đơn)
    mà chưa có giá trị -> lấy lựa chọn đầu tiên của mẫu. Trả về cảnh báo giá trị không khớp danh sách."""
    warn = []
    for c, key in cols.items():
        opts = lists.get(c)
        if not opts:
            continue
        for r in rows:
            v = r.get(key)
            if v in (None, ""):
                if key in ("hinh_thuc_mua", "pttt", "nhan_kem_hd"):
                    r[key] = opts[0]
                continue
            fit = _fit_option(v, opts)
            if fit not in opts:
                warn.append(f"{v} (cột {key})")
            r[key] = fit
    return sorted(set(warn))


def write_template(tpl: bytes, rows: list[dict[str, Any]], colmap: dict[int, str] | None = None):
    wb, ws, hr, cols, _unknown = read_template(tpl)
    cols = colmap or cols
    # xóa dòng ví dụ có sẵn dưới tiêu đề (giữ định dạng)
    for r in range(hr + 1, ws.max_row + 1):
        for c in cols:
            ws.cell(r, c).value = None
    for i, row in enumerate(rows, hr + 1):
        for c, key in cols.items():
            v = row.get(key)
            cell = ws.cell(i, c, v)
            if key in ("thanh_tien", "tien_thue", "don_gia", "tong_tt", "tien_ck"):
                cell.number_format = "#,##0.##"
            elif key.startswith("ngay"):
                cell.number_format = "dd/mm/yyyy"
    return wb


BASIC_COLS = {
    "purchase": ["Ngày hạch toán (*)", "Ngày chứng từ (*)", "Số chứng từ (*)", "Mã nhà cung cấp", "Tên nhà cung cấp",
                 "Địa chỉ", "Mã số thuế", "Diễn giải", "Ngày hóa đơn", "Ký hiệu HĐ", "Số hóa đơn", "Mã hàng (*)",
                 "Tên hàng", "Kho", "TK kho (Nợ)", "TK công nợ (Có)", "ĐVT", "Số lượng", "Đơn giá", "Thành tiền",
                 "Tiền chiết khấu", "% thuế GTGT", "Tiền thuế GTGT", "TK thuế GTGT", "Nhóm HHDV mua vào"],
    "sold": ["Ngày hạch toán (*)", "Ngày chứng từ (*)", "Số chứng từ (*)", "Mã khách hàng", "Tên khách hàng", "Địa chỉ",
             "Mã số thuế", "Diễn giải", "Ngày hóa đơn", "Ký hiệu HĐ", "Số hóa đơn", "Mã hàng (*)", "Tên hàng", "ĐVT",
             "Số lượng", "Đơn giá", "Thành tiền", "Tiền chiết khấu", "TK nợ", "TK có", "% thuế GTGT", "Tiền thuế GTGT",
             "TK thuế GTGT"],
}


def basic_template(direction: str) -> bytes:
    """Mẫu cơ bản theo các cột thường gặp của mẫu nhập khẩu MISA (dùng khi khách chưa có file mẫu riêng)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "Mua hang" if direction == "purchase" else "Ban hang"
    for i, c in enumerate(BASIC_COLS[direction], 1):
        cell = ws.cell(1, i, c)
        cell.font, cell.fill = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
        ws.column_dimensions[cell.column_letter].width = max(12, len(c) + 2)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------------ chạy hàng loạt
@dataclass
class Job:
    id: str
    status: str = "running"
    total: int = 0
    done: int = 0
    ai: int = 0
    log: list[str] = field(default_factory=list)
    results: list[dict[str, Any]] = field(default_factory=list)
    file: str = ""
    cancel: bool = False

    def say(self, m: str) -> None:
        self.log.append(f"{datetime.now():%H:%M:%S} {m}")
        self.log = self.log[-200:]

    def public(self) -> dict[str, Any]:
        return {"id": self.id, "status": self.status, "total": self.total, "done": self.done, "ai": self.ai,
                "log": self.log[-60:], "file": self.file, "results": self.results[:500]}


JOBS: dict[str, Job] = {}


def _read_tuple(args):
    name, data, idx = args
    try:
        return read_one(name, data, idx)
    except Exception as e:  # noqa: BLE001
        return {"file": name, "loi": f"Lỗi đọc: {e}", "items": []}


def _read_path(args):
    """Tiến trình phụ: đọc file từ ổ đĩa (không gửi nội dung file qua lại giữa tiến trình -> đỡ tốn RAM)."""
    name, path, idx = args
    try:
        return read_one(name, Path(path).read_bytes(), idx)
    except MemoryError:
        return {"file": name, "loi": "Không đủ bộ nhớ để đọc file này", "items": []}
    except Exception as e:  # noqa: BLE001
        return {"file": name, "loi": f"Lỗi đọc: {e}", "items": []}


def _read_all(items: list[tuple[str, bytes]], idx: dict, job: "Job") -> list[dict[str, Any]]:
    """Đọc nhiều hóa đơn song song bằng ÍT tiến trình (OCR bản scan rất tốn RAM): ghi file tạm ra ổ đĩa, mỗi tiến
    trình đọc 1 file. Tiến trình phụ hỏng (thiếu RAM...) thì đọc nốt phần còn lại ngay trong tiến trình chính."""
    import os
    import tempfile
    from concurrent.futures import as_completed

    invs: list[dict[str, Any] | None] = [None] * len(items)
    # Dùng LUỒNG trong tiến trình chính, không dùng tiến trình phụ: bộ OCR của Windows (WinRT) chạy trong tiến trình phụ
    # báo "Not enough memory resources" với mọi file -> hóa đơn scan không đọc được. OCR WinRT nhả GIL nên vẫn song song.
    workers = max(2, min(4, os.cpu_count() or 2))
    with tempfile.TemporaryDirectory(prefix="awb_hd_") as td:
        paths = []
        for i, (n, d) in enumerate(items):
            p = Path(td) / f"{i:05d}{Path(n).suffix.lower() or '.bin'}"
            p.write_bytes(d)
            paths.append(str(p))
        done = 0

        def took(i: int, inv: dict[str, Any]) -> None:
            nonlocal done
            invs[i] = inv
            done += 1
            if inv.get("items") or inv.get("loi"):
                job.done += 1
            if done % 10 == 0 or done == len(items):
                job.say(f"Đã đọc {done}/{len(items)} hóa đơn")

        if len(items) >= 2:
            try:
                with ThreadPoolExecutor(workers) as ex:
                    futs = {ex.submit(_read_path, (n, paths[i], idx)): i for i, (n, _d) in enumerate(items)}
                    for f in as_completed(futs):
                        took(futs[f], f.result())
            except (MemoryError, OSError) as e:
                job.say(f"Đọc song song bị dừng ({type(e).__name__}) – đọc tiếp lần lượt từng hóa đơn")
        # file lỗi vì thiếu tài nguyên khi đọc song song: đọc lại lần lượt
        for i, inv in enumerate(invs):
            if inv and "memory" in str(inv.get("loi") or "").lower() + str(inv.get("loi") or "").lower().replace("bộ nhớ", "memory"):
                invs[i] = None
                job.done -= 1 if (inv.get("items") or inv.get("loi")) else 0
                done -= 1
        for i, (n, _d) in enumerate(items):
            if invs[i] is None:
                took(i, _read_path((n, paths[i], idx)))
    return invs  # type: ignore[return-value]


_CACHE_VER = "1"


def _cache_dir() -> Path:
    return Path(settings.output_dir) / ".cache" / "hoadon"


def _cache_key(data: bytes) -> str:
    import hashlib
    return hashlib.sha1(data).hexdigest() + _CACHE_VER


def _cache_get(key: str, name: str) -> dict[str, Any] | None:
    try:
        v = json.loads((_cache_dir() / f"{key}.json").read_text(encoding="utf-8"))
        return v | {"file": name}
    except Exception:  # noqa: BLE001
        return None


def _cache_put(key: str, inv: dict[str, Any] | None) -> None:
    """Chỉ lưu hóa đơn đọc được dòng hàng (không lưu file lỗi để lần sau còn đọc lại)."""
    if not inv or not inv.get("items") or inv.get("loi"):
        return
    try:
        _cache_dir().mkdir(parents=True, exist_ok=True)
        keep = {k: v for k, v in inv.items() if k != "_text"}
        (_cache_dir() / f"{key}.json").write_text(json.dumps(keep, ensure_ascii=False, default=str), encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass


def run(job: Job, files: list[tuple[str, bytes]], tpl: bytes, direction: str, own_mst: str, acc: dict[str, str],
        tpl_name: str = "") -> None:
    from . import fill_template as ft
    try:
        kind, ext = ft.kind_of(tpl), ft.out_ext(tpl, tpl_name)
        use_excel = kind in ("xls", "ods", "html", "xml2003") and ft.excel_available()
        if use_excel:  # đọc tiêu đề đúng như Excel thấy, rồi Excel tự điền vào chính file mẫu
            job.say(f"File mẫu {ext}: dùng Excel trên máy để điền dữ liệu, giữ nguyên định dạng")
            wb, ws, hr, cols, unknown = read_template(ft.probe_with_excel(tpl, ext))
        else:
            wb, ws, hr, cols, unknown = read_template(tpl)
        sheet_index = wb.worksheets.index(ws) + 1
        job.say(f"File mẫu: sheet '{ws.title}', tiêu đề ở dòng {hr}, nhận được {len(cols)} cột"
                + (f"; chưa nhận: {', '.join(map(str, unknown[:8]))}" if unknown else ""))
        items = [(n, d) for name, data in files for n, d in iter_files(name, data)
                 if n.lower().endswith((".xml", ".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"))]
        job.total = len(items)
        idx = local_index()
        job.say(f"Đọc {len(items)} hóa đơn (dùng XML/dữ liệu cổng trên máy nếu có – {len(idx)} hóa đơn đã tải)")
        # hóa đơn đã đọc ở lần chạy trước (cùng nội dung file): lấy lại kết quả, không OCR / AI lại
        keys = [_cache_key(d) for _n, d in items]
        invs: list[dict[str, Any] | None] = [_cache_get(k, n) for k, (n, _d) in zip(keys, items)]
        todo = [i for i, v in enumerate(invs) if v is None]
        if len(todo) < len(items):
            job.done += len(items) - len(todo)
            job.say(f"{len(items) - len(todo)} hóa đơn đã đọc ở lần trước – dùng lại kết quả, đọc mới {len(todo)}")
        # bước 1: đọc nhanh không cần AI – vài tiến trình song song (đọc PDF / OCR tốn CPU, RAM)
        for i, v in zip(todo, _read_all([items[i] for i in todo], idx, job) if todo else []):
            invs[i] = v
        need_ai = [i for i, v in enumerate(invs) if not v.get("items") and not v.get("loi")]
        if need_ai:
            job.say(f"{len(need_ai)} hóa đơn bố cục lạ / bản scan – AI đang đọc (song song)")

            def ai_one(i: int) -> None:
                v = invs[i]
                try:
                    d = ai_extract(v.get("_text") or "")
                    keep = {k: v.get(k) for k in ("khhdon", "shdon", "nbmst", "nmmst", "ngay_lap") if v.get(k)}
                    invs[i] = v | {k: x for k, x in d.items() if x not in (None, "", [])} | keep
                    if not items_ok(invs[i]):
                        invs[i]["canh_bao"] = "Tổng dòng hàng AI đọc chưa khớp tổng tiền hàng – kiểm tra lại"
                    job.ai += 1
                except Exception as e:  # noqa: BLE001
                    invs[i]["loi"] = f"{e}"[:200]
                job.done += 1
            with ThreadPoolExecutor(4) as ex:
                for _ in ex.map(ai_one, need_ai):
                    if job.ai and job.ai % 5 == 0:
                        job.say(f"AI đã đọc {job.ai}/{len(need_ai)} hóa đơn")
        for i in todo + need_ai:
            _cache_put(keys[i], invs[i])
        rows: list[dict[str, Any]] = []
        seq = 0
        for v in invs:
            d = direction
            if direction == "auto":
                d = "sold" if own_mst and str(v.get("nbmst")) == own_mst else "purchase"
            if v.get("loi") and not v.get("shdon"):
                job.results.append({"file": v["file"], "loi": v["loi"]})
                continue
            seq += 1
            r = build_rows(v, d, seq, acc)
            rows += r
            job.results.append({"file": v["file"], "hd": f"{v.get('khhdon') or ''}-{v.get('shdon') or ''}",
                                "ngay": v.get("ngay_lap"), "doi_tac": v.get("nbten") if d == "purchase" else v.get("nmten"),
                                "dong": len(r), "tong": v.get("tgtttbso"), "nguon": v.get("nguon"),
                                "huong": "Mua vào" if d == "purchase" else "Bán ra",
                                "loi": v.get("loi") or v.get("canh_bao") or ""})
        # làm theo ghi chú (danh sách chọn) từng cột của mẫu
        lists = template_lists(ws, hr, cols)
        if lists:
            job.say("Theo danh sách chọn của mẫu: " + "; ".join(
                f"{ws.cell(hr, c).value}: {' / '.join(o[:4])}" for c, o in lists.items() if c in cols)[:600])
            bad = apply_lists(rows, cols, lists)
            if bad:
                job.say("Giá trị chưa có trong danh sách chọn của mẫu (kiểm tra lại): " + ", ".join(bad[:10]))
        # Điền vào CHÍNH file mẫu, giữ nguyên định dạng, cùng loại file với mẫu
        if kind == "ooxml":
            data = ft.fill_ooxml(tpl, ws.title, hr, cols, rows)
        elif kind == "text":
            data = ft.fill_text(tpl, hr, cols, rows)
        elif use_excel:
            data = ft.fill_with_excel(tpl, ext, sheet_index, hr, cols, rows)
        else:
            job.say(f"Máy không có Microsoft Excel nên không ghi được thẳng vào file mẫu {ext} – "
                    "xuất .xlsx theo đúng cột của mẫu (định dạng ô có thể khác mẫu)")
            buf = io.BytesIO()
            write_template(tpl, rows).save(buf)
            data, ext = buf.getvalue(), ".xlsx"
        out = Path(settings.output_dir) / "ChuyenMISA"
        out.mkdir(parents=True, exist_ok=True)
        stem = Path(tpl_name).stem if tpl_name else ("MISA_" + ("MuaVao" if direction == "purchase" else
                                                               "BanRa" if direction == "sold" else "HoaDon"))
        dest = out / f"{stem}_{datetime.now():%Y%m%d_%H%M%S}{ext}"
        dest.write_bytes(data)
        job.file = str(dest)
        job.status = "done"
        job.say(f"Xong: {seq} hóa đơn → {len(rows)} dòng trong file mẫu ({job.ai} hóa đơn AI đọc)")
    except Exception as e:  # noqa: BLE001
        job.status = "error"
        job.say(f"LỖI {e}")


def start(files: list[tuple[str, bytes]], tpl: bytes, direction: str, own_mst: str = "",
          acc: dict[str, str] | None = None, tpl_name: str = "") -> Job:
    job = Job(uuid.uuid4().hex)
    JOBS[job.id] = job
    threading.Thread(target=run, args=(job, files, tpl, direction, own_mst, acc or {}, tpl_name), daemon=True,
                     name="misa").start()
    return job
