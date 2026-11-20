"""Đọc sao kê ngân hàng (Excel / CSV / PDF có chữ) và đối chiếu với hóa đơn.

Không làm mẫu cứng cho từng ngân hàng: tự tìm dòng tiêu đề và nhận cột theo các cách gọi phổ biến của
Vietcombank, Techcombank, MB, ACB, VietinBank, BIDV, Agribank, VPBank... (Ngày giao dịch, Nội dung/Diễn giải,
Ghi nợ/Debit/Số tiền rút, Ghi có/Credit/Số tiền gửi, Số dư/Balance, Số tham chiếu...).
"""
from __future__ import annotations

import io
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from .docreader import vn_num
from .gdt.xml_parser import norm

# Khóa: tên cột đã chuẩn hóa (bỏ dấu, viết thường, bỏ khoảng trắng). Khớp "chứa" chuỗi.
COLS: list[tuple[str, tuple[str, ...]]] = [
    ("ngay", ("ngaygiaodich", "ngaygd", "ngayhieuluc", "transactiondate", "txndate", "ngayhachtoan",
              "ngaychungtu", "postingdate", "valuedate", "ngay", "date")),
    ("no", ("ghino", "sotienghino", "phatsinhno", "sotienrut", "tienra", "debit", "withdrawal", "chi", "rut")),
    ("co", ("ghico", "sotienghico", "phatsinhco", "sotiengui", "tienvao", "credit", "deposit", "thu", "gui")),
    ("so_du", ("sodu", "balance")),
    ("mo_ta", ("noidunggiaodichngan", "noidunggiaodich", "noidung", "diengiai", "mota", "chitietgiaodich", "chitiet", "description", "transactiondetail",
               "details", "remark", "narrative")),
    ("tham_chieu", ("sothamchieu", "magiaodich", "sogiaodich", "sochungtu", "sobuttoan", "reference", "refno",
                    "transactionno", "seqno", "sogd", "sott")),
    ("so_tien", ("sotien", "amount")),  # 1 cột số tiền có dấu +/- (một số ngân hàng)
    ("loai", ("loaigiaodich", "nohayco", "drcr", "type")),
    ("doi_ung", ("tentaikhoandoiung", "tendoiung", "tennguoichuyen", "tennguoinhan", "beneficiary",
                 "counterparty", "remittersaccountname")),
    ("tk_doi_ung", ("sotaikhoandoiung", "taikhoandoiung", "tkdoiung", "remittersaccountnumber")),
    ("nh_doi_ung", ("nganhangdoiung", "remittersbank", "nganhangnhan", "nganhangchuyen")),
]


def _map_header(cells: list[Any]) -> dict[str, int]:
    names = [norm(str(c or "")) for c in cells]
    m: dict[str, int] = {}
    for field, keys in COLS:
        for k in keys:  # ưu tiên khóa cụ thể đứng trước
            for i, h in enumerate(names):
                # khóa ngắn (thu, chi, rut, gui...) phải khớp nguyên tên cột, tránh nhận nhầm "Thứ tự"
                # cột ghi chú / hạch toán của KẾ TOÁN (vd "Nội dung bên kế toán") không phải dữ liệu ngân hàng
                if "ketoan" in h or "hachtoan" in h and field != "ngay":
                    continue
                if h and i not in m.values() and (h == k or (len(k) > 3 and k in h)):
                    m[field] = i
                    break
            if field in m:
                break
    return m


def _parse_date(v: Any) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v or "").strip()
    m = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})", s)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def _rows_to_txns(rows: list[list[Any]]) -> list[dict[str, Any]]:
    head_i, mp = None, {}
    for i, r in enumerate(rows[:60]):
        m = _map_header(r)
        if "ngay" in m and ("mo_ta" in m) and (("no" in m and "co" in m) or "so_tien" in m):
            head_i, mp = i, m
            break
    if head_i is None:
        return []
    out, last = [], None
    for r in rows[head_i + 1:]:
        get = lambda f: r[mp[f]] if f in mp and mp[f] < len(r) else None  # noqa: E731
        d = _parse_date(get("ngay"))
        desc = str(get("mo_ta") or "").strip()
        if not d:  # dòng nối tiếp nội dung của giao dịch trước (PDF hay bị xuống dòng)
            if last is not None and desc and not any(vn_num(get(f)) for f in ("no", "co", "so_tien")):
                last["mo_ta"] = (last["mo_ta"] + " " + desc).strip()
            continue
        no, co = vn_num(get("no")) or 0.0, vn_num(get("co")) or 0.0
        if "so_tien" in mp and not ("no" in mp and "co" in mp):
            amt = vn_num(get("so_tien")) or 0.0
            loai = norm(str(get("loai") or ""))
            if amt < 0 or loai in ("d", "dr", "no", "ghino", "debit"):
                no, co = abs(amt), 0.0
            else:
                no, co = 0.0, amt
        if not no and not co:
            continue
        txt = lambda f: re.sub(r"\s+", " ", str(get(f) or "")).strip()  # noqa: E731
        last = {"ngay": d, "mo_ta": desc, "no": abs(no), "co": abs(co), "so_du": vn_num(get("so_du")),
                # số bút toán / số tài khoản bị xuống dòng trong ô PDF -> nối liền lại
                "tham_chieu": txt("tham_chieu").replace(" ", ""), "doi_ung": txt("doi_ung"),
                "tk_doi_ung": txt("tk_doi_ung").replace(" ", ""), "nh_doi_ung": txt("nh_doi_ung")}
        out.append(last)
    return out


def _fix_direction(txns: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Kiểm tra Nợ/Có bằng số dư chạy: số dư sau = số dư trước + có - nợ.
    Tự nhận thứ tự sao kê (cũ -> mới hoặc mới -> cũ); giao dịch nào chỉ khớp khi đảo chiều thì đảo lại."""
    bal = [t.get("so_du") for t in txns]
    if sum(b is not None for b in bal) < 3:
        return txns

    def ok(prev: float, t: dict[str, Any], swap: bool) -> bool:
        no, co = (t["co"], t["no"]) if swap else (t["no"], t["co"])
        return abs(prev + co - no - t["so_du"]) <= 1

    def score(seq: list[dict[str, Any]]) -> int:
        return sum(1 for a, b in zip(seq, seq[1:]) if a.get("so_du") is not None and b.get("so_du") is not None
                   and ok(a["so_du"], b, False))

    seq = txns if score(txns) >= score(txns[::-1]) else txns[::-1]  # seq: cũ -> mới
    for a, b in zip(seq, seq[1:]):
        if a.get("so_du") is None or b.get("so_du") is None:
            continue
        if not ok(a["so_du"], b, False) and ok(a["so_du"], b, True):
            b["no"], b["co"] = b["co"], b["no"]
    return txns


def _excel_rows(data: bytes, name: str) -> list[list[list[Any]]]:
    low = name.lower()
    if low.endswith(".csv"):
        import csv
        txt = data.decode("utf-8-sig", "ignore")
        return [list(csv.reader(io.StringIO(txt)))]
    if low.endswith(".xls") and data[:2] != b"PK":
        import xlrd
        book = xlrd.open_workbook(file_contents=data)
        sheets = []
        for sh in book.sheets():
            rows = []
            for i in range(sh.nrows):
                row = []
                for c in sh.row(i):
                    row.append(xlrd.xldate.xldate_as_datetime(c.value, book.datemode)
                               if c.ctype == xlrd.XL_CELL_DATE else c.value)
                rows.append(row)
            sheets.append(rows)
        return sheets
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        ws.reset_dimensions()
        sheets.append([list(r) for r in ws.iter_rows(values_only=True)])
    return sheets


def _pdf_rows(data: bytes) -> list[list[Any]]:
    import pdfplumber
    rows: list[list[Any]] = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for p in pdf.pages:
            for t in p.extract_tables() or []:
                rows += [[(c or "").replace("\n", " ").strip() for c in r] for r in t]
    return rows


def _pdf_grid_rows(data: bytes) -> list[list[Any]]:
    """Sao kê PDF mà chỉ dòng tiêu đề có khung đủ cột, các dòng giao dịch chỉ có đường kẻ ngang
    (vd Techcombank): lấy ranh giới cột từ ô tiêu đề, ranh giới dòng từ đường kẻ ngang,
    rồi xếp từng chữ vào đúng ô -> nội dung nhiều dòng không bị trộn sang cột khác."""
    import pdfplumber
    out: list[list[Any]] = []
    header_done = False
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for p in pdf.pages:
            head = None
            for t in p.find_tables():
                for row in t.rows:
                    cells = [c for c in row.cells if c]
                    if len(cells) < 4:
                        continue
                    texts = [(p.crop(c).extract_text() or "").replace("\n", " ") for c in cells]
                    m = _map_header(texts)
                    if "ngay" in m and "mo_ta" in m and (("no" in m and "co" in m) or "so_tien" in m):
                        head = (cells, texts, row.bbox[3])
                        break
                if head:
                    break
            if not head:
                continue
            cells, texts, top = head
            if not header_done:
                out.append(texts)
                header_done = True
            xs = [(c[0], c[2]) for c in cells]
            x_min, x_max = xs[0][0], xs[-1][1]
            # Chỉ lấy đường kẻ phủ gần hết bề ngang bảng (đường chia dòng giao dịch),
            # bỏ khung chân trang / ghi chú để chữ ở đó không lẫn vào giao dịch cuối trang.
            cover: dict[float, float] = {}
            for e in p.edges:
                if e["orientation"] == "h" and e["top"] > top + 2:
                    a, b = max(e["x0"], x_min), min(e["x1"], x_max)
                    if b > a:
                        y = round(e["top"])
                        cover[y] = cover.get(y, 0.0) + (b - a)
            bounds = [top] + sorted(y for y, w in cover.items() if w >= 0.9 * (x_max - x_min))
            if len(bounds) < 2:
                continue
            words = [w for w in p.extract_words() if w["top"] > top]
            grid: dict[tuple[int, int], list[dict]] = {}
            for w in words:
                cy, cx = (w["top"] + w["bottom"]) / 2, (w["x0"] + w["x1"]) / 2
                ri = next((i for i in range(len(bounds) - 1) if bounds[i] <= cy < bounds[i + 1]), None)
                ci = next((i for i, (a, b) in enumerate(xs) if a <= cx < b), None)
                if ri is not None and ci is not None:
                    grid.setdefault((ri, ci), []).append(w)
            for ri in range(len(bounds) - 1):
                row = []
                for ci in range(len(xs)):
                    ws = sorted(grid.get((ri, ci), []), key=lambda w: (round(w["top"]), w["x0"]))
                    row.append(" ".join(w["text"] for w in ws))
                if any(row):
                    out.append(row)
    return out


_DATE = re.compile(r"(?<!\d)(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(20\d{2})(?!\d)")
_AMT = re.compile(r"(?<![\d.,])[+-]?\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{1,2})?(?![\d])")


def _text_to_txns(text: str) -> list[dict[str, Any]]:
    """Đọc sao kê từ chữ thuần (PDF không có bảng / ảnh scan qua OCR), mỗi giao dịch 1 dòng:
    ngày ... nội dung ... số tiền ... số dư. Chiều tiền xác định bằng biến động số dư so với dòng trước
    (giảm đúng bằng số tiền -> ghi nợ, tăng -> ghi có); không có số dư thì theo dấu +/- và từ khóa."""
    from .docreader import fold, _ocr_clean
    text = _ocr_clean(text)
    f = fold(text)
    prev_bal = None
    m = re.search(r"(?:so\s*du\s*dau(?:\s*ky)?|opening\s*balance|so\s*du\s*dau\s*ngay)\D{0,40}?(\d{1,3}(?:[.,]\d{3})+)", f, re.I)
    if m:
        prev_bal = vn_num(m.group(1))
    out: list[dict[str, Any]] = []
    for raw, ln in zip(text.splitlines(), f.splitlines()):
        dm = _DATE.search(ln)
        amts = [(a.start(), a.group(0)) for a in _AMT.finditer(ln)]
        if not dm or not amts:
            if out and ln.strip() and not amts and not dm:  # dòng nối tiếp nội dung
                out[-1]["mo_ta"] = (out[-1]["mo_ta"] + " " + raw.strip())[:500]
            continue
        try:
            d = date(int(dm.group(3)), int(dm.group(2)), int(dm.group(1)))
        except ValueError:
            continue
        vals = [vn_num(a) for _, a in amts]
        bal = vals[-1] if len(vals) >= 2 else None
        amt = vals[-2] if len(vals) >= 2 else vals[0]
        if amt is None:
            continue
        sign_txt = amts[-2][1] if len(vals) >= 2 else amts[0][1]
        no = co = 0.0
        if bal is not None and prev_bal is not None and abs(prev_bal - abs(amt) - bal) <= 1:
            no = abs(amt)
        elif bal is not None and prev_bal is not None and abs(prev_bal + abs(amt) - bal) <= 1:
            co = abs(amt)
        elif sign_txt.startswith("-"):
            no = abs(amt)
        elif sign_txt.startswith("+"):
            co = abs(amt)
        elif re.search(r"\b(nhan|nop tien|credit|tien vao|hoan tien|lai)\b", ln, re.I):
            co = abs(amt)
        else:
            no = abs(amt)
        prev_bal = bal if bal is not None else prev_bal
        first_amt = amts[0][0]
        desc = raw[dm.end():first_amt].strip(" |-:") or raw[first_amt:].strip()
        out.append({"ngay": d, "mo_ta": desc[:500], "no": no, "co": co, "so_du": bal,
                    "tham_chieu": "", "doi_ung": ""})
    return out


def read_statement(name: str, data: bytes) -> tuple[list[dict[str, Any]], str | None]:
    """Trả về (giao dịch, lỗi nếu có). PDF có bảng -> đọc bảng; PDF/ảnh scan -> OCR rồi đọc theo dòng."""
    from . import ocr
    try:
        if name.lower().endswith(".pdf") or data[:5] == b"%PDF-":
            rows = _pdf_rows(data)
            tx = _rows_to_txns(rows) if rows else []
            grid = _pdf_grid_rows(data)
            tx_grid = _rows_to_txns(grid) if grid else []
            if len(tx_grid) > len(tx):  # bảng chỉ có khung tiêu đề -> đọc theo lưới đường kẻ
                tx = tx_grid
            if not tx:  # không có bảng: thử chữ thuần, rồi OCR nếu là bản scan
                from .docreader import pdf_text
                t = pdf_text(data)
                if len(t.strip()) < 40:
                    t = ocr.ocr_pdf(data)
                tx = _text_to_txns(t)
        elif ocr.is_image(name, data):
            tx = _text_to_txns(ocr.ocr_image(data))
        else:
            tx = []
            for rows in _excel_rows(data, name):
                tx += _rows_to_txns(rows)
        if not tx:
            return [], "Không nhận ra bảng giao dịch (không tìm thấy dòng tiêu đề Ngày / Nội dung / Ghi nợ / Ghi có)"
        tx = _fix_direction(tx)
        for t in tx:
            t["file"] = name
        return tx, None
    except Exception as e:  # noqa: BLE001
        return [], f"Lỗi đọc file: {e}"


# ------------------------------------------------------------------ đối chiếu
_STOP = {"cong", "ty", "tnhh", "co", "phan", "mtv", "mot", "thanh", "vien", "thuong", "mai", "dich", "vu", "va",
         "xuat", "nhap", "khau", "san", "xuat", "tm", "dv", "cp", "ctcp", "cty", "chi", "nhanh", "tp", "hcm"}


def _tokens(s: str) -> set[str]:
    import unicodedata
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").replace("đ", "d").replace("Đ", "D").lower()
    return {w for w in re.findall(r"[a-z0-9]+", s) if len(w) > 2 and w not in _STOP}


def _score(inv: dict[str, Any], tx: dict[str, Any]) -> tuple[int, list[str]]:
    amt = float(inv.get("tgtttbso") or 0)
    val = tx["no"] if inv.get("huong") == "purchase" else tx["co"]
    if not amt or not val:
        return 0, []
    why, s = [], 0
    if abs(val - amt) <= 1:
        s += 50
        why.append("đúng số tiền")
    elif abs(val - amt) <= max(50000, amt * 0.003):
        s += 25
        why.append("lệch phí nhỏ")
    else:
        return 0, []
    desc = " ".join([tx.get("mo_ta", ""), tx.get("doi_ung", ""), tx.get("tham_chieu", "")])
    digits = set(re.findall(r"\d+", desc))
    shdon = str(inv.get("shdon") or "").lstrip("0")
    if shdon and (shdon in {d.lstrip("0") for d in digits}):
        s += 30
        why.append(f"có số HĐ {shdon}")
    mst = inv.get("nbmst") if inv.get("huong") == "purchase" else inv.get("nmmst")
    if mst and mst in re.sub(r"\D", " ", desc).split():
        s += 30
        why.append("có MST đối tác")
    name = inv.get("nbten") if inv.get("huong") == "purchase" else inv.get("nmten")
    common = _tokens(name or "") & _tokens(desc)
    if len(common) >= 2:
        s += 15
        why.append("trùng tên đối tác")
    try:
        d_inv = date.fromisoformat(str(inv.get("ngay_lap"))[:10])
        gap = (tx["ngay"] - d_inv).days
        if -5 <= gap <= 90:
            s += 5
    except ValueError:
        pass
    return s, why


def match(invoices: list[dict[str, Any]], txns: list[dict[str, Any]], threshold: int = 55):
    """Ghép 1–1 giao dịch với hóa đơn theo điểm (số tiền, số HĐ / MST / tên đối tác trong nội dung, ngày)."""
    cands = []
    for i, inv in enumerate(invoices):
        for j, tx in enumerate(txns):
            s, why = _score(inv, tx)
            if s >= threshold:
                cands.append((s, i, j, why))
    cands.sort(key=lambda x: -x[0])
    used_i, used_j, pairs = set(), set(), []
    for s, i, j, why in cands:
        if i in used_i or j in used_j:
            continue
        used_i.add(i)
        used_j.add(j)
        pairs.append((invoices[i], txns[j], s, why))
    unpaid = [inv for k, inv in enumerate(invoices) if k not in used_i]
    unmatched = [tx for k, tx in enumerate(txns) if k not in used_j]
    return pairs, unpaid, unmatched


def write_report(dest: Path, pairs, unpaid, unmatched, txns, errors: list[str]) -> Path:
    from openpyxl import Workbook

    from .reports import BOLD, FILL, WARN, _header, _save
    wb = Workbook()
    ws = wb.active
    ws.title = "Tom tat"
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 18
    tin = sum(t["co"] for t in txns)
    tout = sum(t["no"] for t in txns)
    lines = [("ĐỐI CHIẾU SAO KÊ – HÓA ĐƠN", None), ("Số giao dịch", len(txns)), ("Tổng tiền vào (ghi có)", tin),
             ("Tổng tiền ra (ghi nợ)", tout), ("Giao dịch đã khớp hóa đơn", len(pairs)),
             ("Hóa đơn chưa thấy thanh toán", len(unpaid)), ("Giao dịch chưa có hóa đơn", len(unmatched))]
    for i, (a, b) in enumerate(lines, 1):
        c = ws.cell(i, 1, a)
        if i == 1:
            c.font = BOLD
        if b is not None:
            ws.cell(i, 2, b).number_format = "#,##0"
    for k, e in enumerate(errors, len(lines) + 2):
        ws.cell(k, 1, "Lỗi đọc: " + e)

    # Cột giao dịch dùng chung cho các sheet: diễn giải tách riêng khỏi người/đơn vị đối ứng.
    TX_COLS = [("Ngày GD", 11), ("Ghi nợ (tiền ra)", 15), ("Ghi có (tiền vào)", 15), ("Diễn giải", 45),
               ("Tên đối ứng (người chuyển / nhận)", 34), ("Số TK đối ứng", 18), ("Ngân hàng đối ứng", 24),
               ("Số bút toán", 18), ("File", 18)]
    N_TX = len(TX_COLS)

    def tx_vals(t):
        return [t["ngay"], t["no"] or None, t["co"] or None, t.get("mo_ta"), t.get("doi_ung"),
                t.get("tk_doi_ung"), t.get("nh_doi_ung"), t.get("tham_chieu"), t.get("file")]

    ws2 = wb.create_sheet("Da khop")
    _header(ws2, TX_COLS + [("Loại HĐ", 22), ("Ký hiệu", 10), ("Số HĐ", 9), ("Ngày HĐ", 11),
                            ("Đối tác trên HĐ", 32), ("Tổng TT HĐ", 15), ("Điểm", 7), ("Căn cứ", 40)])
    for r, (inv, t, s, why) in enumerate(pairs, 2):
        partner = inv.get("nbten") if inv.get("huong") == "purchase" else inv.get("nmten")
        for c, v in enumerate(tx_vals(t) + [inv.get("loai"), inv.get("khhdon"), inv.get("shdon"), inv.get("ngay_lap"),
                                             partner, inv.get("tgtttbso"), s, ", ".join(why)], 1):
            cell = ws2.cell(r, c, v)
            if c in (2, 3, N_TX + 6):
                cell.number_format = "#,##0"
            if c == 1:
                cell.number_format = "dd/mm/yyyy"

    ws3 = wb.create_sheet("HD chua thanh toan")
    _header(ws3, [("Loại", 24), ("Ngày HĐ", 11), ("Ký hiệu", 10), ("Số HĐ", 9), ("MST đối tác", 14), ("Đối tác", 36),
                  ("Tổng thanh toán", 15), ("Trạng thái", 14)])
    for r, inv in enumerate(unpaid, 2):
        purchase = inv.get("huong") == "purchase"
        for c, v in enumerate([inv.get("loai"), inv.get("ngay_lap"), inv.get("khhdon"), inv.get("shdon"),
                               inv.get("nbmst") if purchase else inv.get("nmmst"),
                               inv.get("nbten") if purchase else inv.get("nmten"), inv.get("tgtttbso"), inv.get("tthai")], 1):
            cell = ws3.cell(r, c, v)
            if c == 7:
                cell.number_format = "#,##0"

    ws4 = wb.create_sheet("GD chua co HD")
    _header(ws4, TX_COLS)
    for r, t in enumerate(unmatched, 2):
        for c, v in enumerate(tx_vals(t), 1):
            cell = ws4.cell(r, c, v)
            if c in (2, 3):
                cell.number_format = "#,##0"
            if c == 1:
                cell.number_format = "dd/mm/yyyy"
            if t["no"] >= 20_000_000:  # chi từ 20 triệu không có hóa đơn: lưu ý khấu trừ thuế
                cell.fill = WARN

    ws5 = wb.create_sheet("Sao ke chuan hoa")
    _header(ws5, TX_COLS + [("Số dư", 16)])
    for r, t in enumerate(txns, 2):
        for c, v in enumerate(tx_vals(t) + [t.get("so_du")], 1):
            cell = ws5.cell(r, c, v)
            if c in (2, 3, N_TX + 1):
                cell.number_format = "#,##0"
            if c == 1:
                cell.number_format = "dd/mm/yyyy"
    _ = (FILL,)
    return _save(wb, dest)


def invoices_for_range(root: Path, d1: date, d2: date) -> list[dict[str, Any]]:
    """Hóa đơn đã tải (file .dulieu_YYYY-MM.json) trong khoảng ngày sao kê, lùi thêm 2 tháng
    (hóa đơn thường được thanh toán sau ngày lập). Bỏ hóa đơn bị thay thế / bị hủy."""
    from .reports import load_period_data, tthai_code
    start = (d1.replace(day=1) - timedelta(days=62)).replace(day=1)
    months, cur = [], start
    while cur <= d2:
        months.append((cur.year, cur.month))
        cur = date(cur.year + (cur.month == 12), cur.month % 12 + 1, 1)
    rows = load_period_data(root, months)
    return [r for r in rows if tthai_code(r) not in (4, 6) and r.get("tgtttbso")]
