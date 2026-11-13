"""Nhập hóa đơn từ file tải trên cổng hoadondientu.gdt.gov.vn.

Hỗ trợ:
  * .xml  – XML hóa đơn (TT78), từng file
  * .zip  – file zip cổng thuế xuất (chứa invoice.xml), kể cả zip lồng zip / nhiều hóa đơn
  * .xlsx / .xls – file Excel "Xuất Excel" danh sách hóa đơn (chỉ có thông tin tổng, không có dòng hàng)

XML là nguồn đầy đủ nhất (có mã tra cứu -> lấy được PDF gốc). Excel chỉ dùng để có danh sách nhanh;
nhập XML sau cho cùng hóa đơn sẽ tự bổ sung vào bản ghi đã có.
"""
from __future__ import annotations

import io
import re
import zipfile
from datetime import date, datetime
from typing import Any, Iterator

from ..gdt.xml_parser import norm, parse_invoice_xml

MAX_DEPTH = 3


# ---------------- tách file ----------------
def iter_items(name: str, data: bytes, depth: int = 0) -> Iterator[tuple[str, str, bytes]]:
    """Trả về (loại, tên, dữ liệu) với loại ∈ {'xml', 'excel', 'skip'}."""
    low = name.lower()
    if data[:2] == b"PK" and (low.endswith(".zip") or not low.endswith((".xlsx", ".xlsm"))):
        if depth >= MAX_DEPTH:
            yield "skip", name, "zip lồng quá sâu".encode()
            return
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                # file .xlsx cũng là zip -> nhận diện qua [Content_Types].xml + xl/
                names = zf.namelist()
                if any(n.startswith("xl/") for n in names) and "[Content_Types].xml" in names:
                    yield "excel", name, data
                    return
                for n in names:
                    if n.endswith("/"):
                        continue
                    yield from iter_items(f"{name}/{n}", zf.read(n), depth + 1)
        except zipfile.BadZipFile:
            yield "skip", name, "zip hỏng".encode()
        return
    if low.endswith((".xlsx", ".xlsm", ".xls")):
        yield "excel", name, data
        return
    if low.endswith(".xml") or data.lstrip()[:5] == b"<?xml":
        head = data[:4000]
        if b"HDon" in head or b"DLHDon" in data[:20000]:
            yield "xml", name, data
        else:
            yield "skip", name, "XML không phải hóa đơn".encode()
        return
    yield "skip", name, "định dạng không hỗ trợ".encode()


# ---------------- XML -> bản ghi ----------------
def _date_iso(v: Any) -> str | None:
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return f"{m[1]}-{m[2]}-{m[3]}"
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        return f"{m[3]}-{int(m[2]):02d}-{int(m[1]):02d}"
    return None


def decide_direction(owner_mst: str, seller_mst: str | None, buyer_mst: str | None,
                     hint: str) -> str:
    o = (owner_mst or "").strip()
    if o and seller_mst and o == seller_mst.strip():
        return "sold"
    if o and buyer_mst and o == buyer_mst.strip():
        return "purchase"
    return hint if hint in ("purchase", "sold") else "purchase"


def row_from_xml(owner_mst: str, parsed: dict[str, Any], hint: str) -> dict[str, Any]:
    nb, nm = parsed.get("nguoi_ban") or {}, parsed.get("nguoi_mua") or {}
    d = _date_iso(parsed.get("ngay_lap"))
    return {
        "owner_mst": owner_mst,
        "direction": decide_direction(owner_mst, nb.get("mst"), nm.get("mst"), hint),
        "nbmst": nb.get("mst") or "",
        "khmshdon": parsed.get("khmshdon") or "",
        "khhdon": parsed.get("khhdon") or "",
        "shdon": str(int(parsed["shdon"])) if (parsed.get("shdon") or "").isdigit() else (parsed.get("shdon") or ""),
        "tdlap": f"{d}T00:00:00+07:00" if d else None,
        "nbten": nb.get("ten"), "nmmst": nm.get("mst"), "nmten": nm.get("ten"),
        "tgtcthue": parsed.get("tong_chua_thue"), "tgtthue": parsed.get("tong_thue"),
        "tgtttbso": parsed.get("tong_thanh_toan"),
        "mhdon": parsed.get("mccqt"), "msttcgp": parsed.get("msttcgp"),
        "source_query": "import-xml",
        "parsed": parsed, "lookup_code": parsed.get("lookup_code"),
        "lookup_url": (parsed.get("lookup_urls") or [None])[0],
    }


# ---------------- Excel -> bản ghi ----------------
# Khóa: tên cột đã chuẩn hóa (bỏ dấu, viết thường, bỏ khoảng trắng). Thứ tự quan trọng: khớp cụ thể trước.
COLUMN_RULES: list[tuple[str, tuple[str, ...]]] = [
    ("khmshdon", ("kyhieumauso", "mauso")),
    ("khhdon", ("kyhieuhoadon", "kyhieu")),
    ("shdon", ("sohoadon",)),
    ("tdlap", ("ngaylap", "ngayhoadon")),
    ("nbmst", ("mstnguoiban", "masothuenguoiban", "mstnban", "mstbenban", "mstnguoibanhang")),
    ("nbten", ("tennguoiban", "tennban", "tenbenban", "tennguoibanhang")),
    ("nmmst", ("mstnguoimua", "masothuenguoimua", "mstnmua", "mstbenmua", "mstnguoimuahang")),
    ("nmten", ("tennguoimua", "tennmua", "tenbenmua", "tennguoimuahang")),
    ("tgtcthue", ("tongtienchuathue", "tienchuathue")),
    ("tgtthue", ("tongtienthue", "tienthue")),
    ("tgtttbso", ("tongtienthanhtoan", "tongthanhtoan")),
    ("tthai_text", ("trangthaihoadon",)),
]


def _map_header(cells: list[Any]) -> dict[str, int]:
    mapping: dict[str, int] = {}
    normed = [norm(str(c or "")) for c in cells]
    for field, keys in COLUMN_RULES:
        for i, h in enumerate(normed):
            if i in mapping.values() or not h:
                continue
            if any(h == k or h.startswith(k) for k in keys):
                mapping[field] = i
                break
    return mapping


def _num(v: Any) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace(" ", "")
    if re.fullmatch(r"-?\d{1,3}([.,]\d{3})+", s):
        s = re.sub(r"[.,]", "", s)
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _rows_from_sheet(rows: list[list[Any]]) -> list[dict[str, Any]]:
    header_idx, mapping = None, {}
    for i, r in enumerate(rows[:30]):
        m = _map_header(r)
        if "shdon" in m and ("khhdon" in m or "nbmst" in m):
            header_idx, mapping = i, m
            break
    if header_idx is None:
        return []
    out = []
    for r in rows[header_idx + 1:]:
        get = lambda f: r[mapping[f]] if f in mapping and mapping[f] < len(r) else None  # noqa: E731
        shdon = get("shdon")
        if shdon in (None, ""):
            continue
        shdon_s = str(int(shdon)) if isinstance(shdon, float) and shdon.is_integer() else str(shdon).strip()
        if not shdon_s.isdigit():
            continue  # dòng tổng cộng / ghi chú
        out.append({k: get(k) for k in mapping} | {"shdon": str(int(shdon_s))})
    return out


def read_excel(data: bytes, name: str) -> list[dict[str, Any]]:
    sheets: list[list[list[Any]]] = []
    if name.lower().endswith(".xls") and data[:2] != b"PK":
        import xlrd
        book = xlrd.open_workbook(file_contents=data)
        for sh in book.sheets():
            rows = []
            for i in range(sh.nrows):
                row = []
                for c in sh.row(i):
                    if c.ctype == xlrd.XL_CELL_DATE:
                        row.append(xlrd.xldate.xldate_as_datetime(c.value, book.datemode))
                    else:
                        row.append(c.value)
                rows.append(row)
            sheets.append(rows)
    else:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        for ws in wb.worksheets:
            # File Excel của cổng thuế ghi sai kích thước sheet (A1:A1) -> phải bỏ qua, nếu không
            # chế độ read_only chỉ đọc được 1 ô.
            ws.reset_dimensions()
            sheets.append([list(r) for r in ws.iter_rows(values_only=True)])
    result: list[dict[str, Any]] = []
    head: list[str] = []
    for rows in sheets:
        result += _rows_from_sheet(rows)
        for r in rows[:12]:
            head += [str(c) for c in r if c not in (None, "")]
    read_excel.last_head = " ".join(head)  # type: ignore[attr-defined]
    return result


MST_RE = re.compile(r"(?<!\d)(\d{10}(?:-?\d{3})?)(?!\d)")


def detect_owner(kind: str, data: bytes, name: str, managed: set[str]) -> tuple[str | None, str]:
    """Đoán file thuộc MST khách hàng nào (trong danh sách managed). Trả về (mst, lý do)."""
    if not managed:
        return None, "chưa khai báo MST khách hàng nào"
    if kind == "xml":
        p = parse_invoice_xml(data)
        buyer = (p.get("nguoi_mua") or {}).get("mst")
        seller = (p.get("nguoi_ban") or {}).get("mst")
        if buyer in managed:
            return buyer, "MST người mua trong XML"
        if seller in managed:
            return seller, "MST người bán trong XML"
        return None, f"XML của {seller} → {buyer}, không thuộc MST nào đã khai báo"
    if kind == "excel":
        rows = read_excel(data, name)
        head = getattr(read_excel, "last_head", "")
        found = {m.replace("-", "") for m in MST_RE.findall(head)} & managed
        if len(found) == 1:
            return found.pop(), "MST ở tiêu đề file Excel"
        for col in ("nmmst", "nbmst"):
            vals = {str(r.get(col) or "").strip() for r in rows} - {""}
            if len(vals) == 1 and vals <= managed:
                return vals.pop(), f"cột {col} trong Excel"
        if len(managed) == 1:
            return next(iter(managed)), "chỉ có 1 MST được khai báo"
        return None, "Excel không ghi rõ thuộc MST nào"
    return None, "không hỗ trợ"


def row_from_excel(owner_mst: str, r: dict[str, Any], hint: str) -> dict[str, Any]:
    from .tconnect import split_symbol

    khms = str(r.get("khmshdon") or "").strip()
    khhd = str(r.get("khhdon") or "").strip()
    if khms.endswith(".0"):
        khms = khms[:-2]
    if not khms:
        khms, khhd = split_symbol(khhd)
    nbmst = str(r.get("nbmst") or "").strip()
    nmmst = str(r.get("nmmst") or "").strip() or None
    if hint == "auto":  # tự đoán: file mua vào chỉ có cột người bán, file bán ra chỉ có cột người mua
        if nbmst and not nmmst:
            hint = "sold" if nbmst == owner_mst else "purchase"
        elif nmmst and not nbmst:
            hint = "purchase" if nmmst == owner_mst else "sold"
        else:
            hint = "sold" if nbmst == owner_mst else "purchase"
    # File "bán ra" thường không có cột MST người bán (chính là mình) và ngược lại.
    if not nbmst and hint == "sold":
        nbmst = owner_mst
    if not nmmst and hint == "purchase":
        nmmst = owner_mst
    d = _date_iso(r.get("tdlap"))
    return {
        "owner_mst": owner_mst,
        "direction": decide_direction(owner_mst, nbmst, nmmst, hint),
        "nbmst": nbmst, "khmshdon": khms, "khhdon": khhd, "shdon": r["shdon"],
        "tdlap": f"{d}T00:00:00+07:00" if d else None,
        "nbten": r.get("nbten"), "nmmst": nmmst, "nmten": r.get("nmten"),
        "tgtcthue": _num(r.get("tgtcthue")), "tgtthue": _num(r.get("tgtthue")),
        "tgtttbso": _num(r.get("tgtttbso")),
        "source_query": "import-excel",
        "raw": {k: (v.isoformat() if isinstance(v, (date, datetime)) else v) for k, v in r.items()},
    }
