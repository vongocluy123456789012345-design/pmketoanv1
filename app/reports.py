"""Báo cáo nâng cao từ dữ liệu hóa đơn đã tải:

  * Xuất Excel chi tiết từng dòng hàng hóa để nhập vào phần mềm kế toán (MISA, Fast, EasyBooks...).
  * Báo cáo hóa đơn cần chú ý: thay thế / điều chỉnh / bị thay thế / bị điều chỉnh / bị hủy, thiếu XML, chỉ có bản in.
  * Lưu dữ liệu từng tháng (.dulieu_YYYY-MM.json) để đối chiếu tờ khai về sau.
  * Đối chiếu tờ khai thuế GTGT 01/GTGT (file XML nộp qua eTax hoặc file Excel HTKK) với hóa đơn.
"""
from __future__ import annotations

import io
import json
import re
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

BOLD = Font(bold=True)
FILL = PatternFill("solid", fgColor="DDEBF7")
WARN = PatternFill("solid", fgColor="FFF2CC")

TTHAI_TEXT = {1: "Hóa đơn mới", 2: "Thay thế", 3: "Điều chỉnh", 4: "Đã bị thay thế",
              5: "Đã bị điều chỉnh", 6: "Đã bị hủy"}


def tthai_code(rec: dict[str, Any]) -> int:
    v = rec.get("tthai_code")
    if isinstance(v, int):
        return v
    try:
        return int(v)
    except (TypeError, ValueError):
        pass
    t = str(rec.get("tthai") or v or "").lower()
    # cụm dài trước: "đã bị thay thế" chứa "thay thế", "đã bị điều chỉnh" chứa "điều chỉnh"
    for k, s in sorted(TTHAI_TEXT.items(), key=lambda kv: -len(kv[1])):
        if s.lower() in t:
            return k
    return 1


def _num(v: Any) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def _rate_pct(r: str) -> float | None:
    r = (r or "").strip().rstrip("%")
    try:
        return float(r.replace(",", "."))
    except ValueError:
        return None


def _header(ws, cols: list[tuple[str, int]]) -> None:
    for i, (title, w) in enumerate(cols, 1):
        c = ws.cell(1, i, title)
        c.font, c.fill = BOLD, FILL
        ws.column_dimensions[c.column_letter].width = w
    ws.freeze_panes = "A2"


def _save(wb: Workbook, dest: Path) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(dest)
    except PermissionError:  # file đang mở trong Excel
        from datetime import datetime
        dest = dest.with_name(f"{dest.stem}_{datetime.now():%H%M%S}{dest.suffix}")
        wb.save(dest)
    return dest


# ------------------------------------------------------------------ xuất phần mềm kế toán
PMKT_COLS = [("Ngày hạch toán", 12), ("Ngày hóa đơn", 12), ("Mẫu số", 7), ("Ký hiệu HĐ", 10), ("Số hóa đơn", 10),
             ("Mã số thuế đối tác", 14), ("Tên đối tác", 36), ("Địa chỉ", 40), ("Mã hàng", 12), ("Tên hàng hóa, dịch vụ", 40),
             ("ĐVT", 8), ("Số lượng", 10), ("Đơn giá", 14), ("Thành tiền", 15), ("Chiết khấu", 12),
             ("% thuế GTGT", 10), ("Tiền thuế GTGT", 14), ("Tổng thanh toán dòng", 16), ("Trạng thái HĐ", 14),
             ("Mã tra cứu", 16), ("Loại", 24)]


def _pmkt_sheet(wb: Workbook, title: str, rows: list[dict[str, Any]]) -> int:
    ws = wb.create_sheet(title)
    _header(ws, PMKT_COLS)
    r = 2
    for rec in rows:
        purchase = rec.get("huong") == "purchase"
        mst = rec.get("nbmst") if purchase else rec.get("nmmst")
        ten = rec.get("nbten") if purchase else rec.get("nmten")
        items = rec.get("_items") or [{"ten": "(không có XML – chỉ có tổng tiền)", "thanh_tien": rec.get("tgtcthue"),
                                       "thue_suat": "", "_thue": rec.get("tgtthue")}]
        for it in items:
            tt = _num(it.get("thanh_tien"))
            pct = _rate_pct(it.get("thue_suat") or "")
            thue = _num(it.get("_thue")) if "_thue" in it else (round(tt * pct / 100) if pct is not None else 0)
            vals = [rec.get("ngay_lap"), rec.get("ngay_lap"), rec.get("khmshdon"), rec.get("khhdon"), rec.get("shdon"),
                    mst, ten, rec.get("_dchi"), it.get("ma"), it.get("ten"), it.get("dvt"), it.get("so_luong"),
                    it.get("don_gia"), tt, it.get("chiet_khau"), it.get("thue_suat"), thue, tt + thue,
                    rec.get("tthai"), rec.get("ma_tra_cuu"), rec.get("loai")]
            for i, v in enumerate(vals, 1):
                c = ws.cell(r, i, v)
                if i in (12, 13, 14, 15, 17, 18):
                    c.number_format = "#,##0.##"
            r += 1
    ws.auto_filter.ref = ws.dimensions
    return r - 2


def write_accounting_export(rows: list[dict[str, Any]], dest: Path) -> Path | None:
    """Excel chi tiết từng dòng hàng (sheet Mua vao / Ban ra) – cột theo mẫu import thông dụng."""
    if not rows:
        return None
    wb = Workbook()
    wb.remove(wb.active)
    n1 = _pmkt_sheet(wb, "Mua vao", [r for r in rows if r.get("huong") == "purchase"])
    n2 = _pmkt_sheet(wb, "Ban ra", [r for r in rows if r.get("huong") != "purchase"])
    if not (n1 or n2):
        return None
    return _save(wb, dest)


# ------------------------------------------------------------------ hóa đơn cần chú ý
def issues(rows: list[dict[str, Any]]) -> list[tuple[dict[str, Any], str]]:
    out = []
    for rec in rows:
        why = []
        code = tthai_code(rec)
        if code != 1:
            why.append(TTHAI_TEXT.get(code, str(rec.get("tthai"))))
        if not rec.get("file_xml"):
            why.append("Không có XML")
        pdf = str(rec.get("file_pdf") or "")
        if not pdf:
            why.append("Không có PDF")
        elif pdf.endswith("_hddt.pdf") and rec.get("ma_tra_cuu"):
            why.append("Có mã tra cứu nhưng chưa lấy được PDF gốc")
        if why:
            out.append((rec, "; ".join(why)))
    return out


def add_issues_sheet(wb: Workbook, rows: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet("Can chu y", 1 if "Theo thang" in wb.sheetnames else 0)
    _header(ws, [("Vấn đề", 40), ("Loại", 26), ("Ngày lập", 11), ("Ký hiệu", 10), ("Số HĐ", 10),
                 ("MST người bán", 14), ("Tên người bán", 32), ("MST người mua", 14), ("Tên người mua", 32),
                 ("Tổng thanh toán", 15), ("Trạng thái", 16), ("Ghi chú", 50)])
    r = 2
    for rec, why in issues(rows):
        vals = [why, rec.get("loai"), rec.get("ngay_lap"), rec.get("khhdon"), rec.get("shdon"), rec.get("nbmst"),
                rec.get("nbten"), rec.get("nmmst"), rec.get("nmten"), rec.get("tgtttbso"), rec.get("tthai"),
                rec.get("ghi_chu")]
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            if tthai_code(rec) in (4, 5, 6):
                c.fill = WARN
            if i == 10:
                c.number_format = "#,##0"
        r += 1
    if r == 2:
        ws.cell(2, 1, "Không có hóa đơn nào cần chú ý")


# ------------------------------------------------------------------ dữ liệu tháng (để đối chiếu)
DATA_KEYS = ("huong", "loai", "thang", "ngay_lap", "khmshdon", "khhdon", "shdon", "nbmst", "nbten", "nmmst", "nmten",
             "tgtcthue", "tgtthue", "tgtttbso", "tthai", "tthai_code", "file_xml", "file_pdf", "_rates")


def _row_key(r: dict[str, Any]) -> tuple:
    return (r.get("huong"), str(r.get("loai") or ""), str(r.get("nbmst") or ""), str(r.get("khmshdon") or ""),
            str(r.get("khhdon") or ""), str(r.get("shdon") or "").lstrip("0"))


def save_month_data(rows: list[dict[str, Any]], dest: Path) -> None:
    """Lưu dữ liệu tháng. GỘP với dữ liệu đã lưu (lượt chạy trước, loại hóa đơn khác) – không ghi đè,
    để chạy lại một phần (vd chỉ 'mua vào') không làm mất dữ liệu 'bán ra' đã có."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    merged: dict[tuple, dict[str, Any]] = {}
    try:
        for r in json.loads(dest.read_text(encoding="utf-8")).get("rows", []):
            merged[_row_key(r)] = r
    except (FileNotFoundError, ValueError):
        pass
    for r in rows:
        merged[_row_key(r)] = {k: r.get(k) for k in DATA_KEYS}
    dest.write_text(json.dumps({"saved": date.today().isoformat(), "rows": list(merged.values())},
                               ensure_ascii=False, default=str), encoding="utf-8")


def load_period_data(root: Path, months: list[tuple[int, int]]) -> list[dict[str, Any]]:
    out = []
    for y, m in months:
        f = root / f"{y:04d}-{m:02d}" / f".dulieu_{y:04d}-{m:02d}.json"
        if f.exists():
            out += json.loads(f.read_text(encoding="utf-8")).get("rows", [])
    return out


# ------------------------------------------------------------------ đối chiếu tờ khai 01/GTGT
# Chỉ tiêu tờ khai 01/GTGT (TT80/2021):
#   [23] giá trị HHDV mua vào   [24] thuế GTGT HHDV mua vào
#   [26] HHDV bán ra không chịu thuế   [29] bán ra 0%
#   [30] bán ra 5% – giá trị   [31] thuế   [32] bán ra 10% (gồm cả 8% được giảm) – giá trị   [33] thuế
#   [32a] HHDV bán ra không phải kê khai, tính nộp thuế (KKKNT)   [34] tổng doanh thu bán ra   [35] tổng thuế bán ra
CHI_TIEU = [("23", "Giá trị HHDV mua vào"), ("24", "Thuế GTGT HHDV mua vào"),
            ("26", "Bán ra không chịu thuế"), ("29", "Bán ra thuế suất 0%"),
            ("30", "Bán ra 5% – doanh thu"), ("31", "Bán ra 5% – thuế"),
            ("32", "Bán ra 8%/10% – doanh thu"), ("33", "Bán ra 8%/10% – thuế"),
            ("32a", "Bán ra không kê khai, tính nộp (KKKNT)"),
            ("34", "Tổng doanh thu bán ra"), ("35", "Tổng thuế GTGT bán ra")]


def parse_tax_return(data: bytes, name: str) -> tuple[dict[str, float], tuple[str, int, int] | None]:
    """Đọc tờ khai 01/GTGT. Trả về ({chỉ tiêu: số}, (kiểu kỳ 'Q'/'M', kỳ, năm) hoặc None)."""
    vals: dict[str, float] = {}
    text_all = ""
    low = name.lower()
    if data.lstrip()[:1] == b"<":  # XML eTax: <ct23>...</ct23>
        txt = data.decode("utf-8", "ignore")
        text_all = txt
        for k, v in re.findall(r"<ct(\d+[a-z]?)>\s*(-?[\d.,]+)\s*</ct\1>", txt, re.I):
            vals[k.lower()] = _vn_num(v)
        kieu = re.search(r"<kieuKy>\s*([QM])\s*</kieuKy>", txt, re.I)
        ky = re.search(r"<kyKKhai>\s*(\d{1,2})/(\d{4})\s*</kyKKhai>", txt, re.I)
        if kieu and ky:
            return vals, (kieu.group(1).upper(), int(ky.group(1)), int(ky.group(2)))
    else:  # Excel HTKK / iHTKK: tìm ô "[23]" rồi lấy số ở ô kế tiếp bên phải
        for row in _sheet_rows(data, low):
            text_all += " ".join(str(c) for c in row if c not in (None, "")) + "\n"
            for i, c in enumerate(row):
                m = re.fullmatch(r"\s*\[(\d+[a-z]?)\]\s*", str(c or ""), re.I)
                if not m:
                    continue
                for nxt in row[i + 1:]:
                    if isinstance(nxt, (int, float)):
                        vals[m.group(1).lower()] = float(nxt)
                        break
                    if isinstance(nxt, str) and re.fullmatch(r"\s*-?[\d.,]+\s*", nxt):
                        vals[m.group(1).lower()] = _vn_num(nxt)
                        break
                    if isinstance(nxt, str) and re.fullmatch(r"\s*\[\d+[a-z]?\]\s*", nxt):
                        break
    return vals, _period_from_text(text_all + " " + name)


def _vn_num(s: str) -> float:
    s = s.strip()
    if re.fullmatch(r"-?\d{1,3}([.,]\d{3})+", s):
        s = re.sub(r"[.,]", "", s)
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _sheet_rows(data: bytes, low: str):
    if low.endswith(".xls") and data[:2] != b"PK":
        import xlrd
        book = xlrd.open_workbook(file_contents=data)
        for sh in book.sheets():
            for i in range(sh.nrows):
                yield [c.value for c in sh.row(i)]
    else:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        for ws in wb.worksheets:
            ws.reset_dimensions()
            yield from (list(r) for r in ws.iter_rows(values_only=True))


def _period_from_text(t: str) -> tuple[str, int, int] | None:
    n = re.sub(r"[\s_\-]+", " ", t.lower())
    m = re.search(r"qu[ýy] (\d) (?:n[aă]m )?(\d{4})", n) or re.search(r"quy (\d) nam (\d{4})", n)
    if m:
        return "Q", int(m.group(1)), int(m.group(2))
    m = re.search(r"th[áa]ng (\d{1,2}) (?:n[aă]m )?(\d{4})", n) or re.search(r"thang (\d{1,2}) nam (\d{4})", n)
    if m:
        return "M", int(m.group(1)), int(m.group(2))
    return None


def period_months(kieu: str, ky: int, nam: int) -> list[tuple[int, int]]:
    return [(nam, m) for m in range(3 * ky - 2, 3 * ky + 1)] if kieu == "Q" else [(nam, ky)]


def invoice_totals(rows: list[dict[str, Any]]) -> dict[str, float]:
    """Cộng hóa đơn theo chỉ tiêu tờ khai. Bỏ hóa đơn đã bị thay thế / bị hủy."""
    t = {k: 0.0 for k, _ in CHI_TIEU}
    for r in rows:
        if tthai_code(r) in (4, 6):
            continue
        rates = r.get("_rates") or {}
        if r.get("huong") == "purchase":
            if rates:
                for v in rates.values():
                    t["23"] += _num(v[0])
                    t["24"] += _num(v[1])
            else:
                t["23"] += _num(r.get("tgtcthue"))
                t["24"] += _num(r.get("tgtthue"))
            continue
        if not rates:
            rates = {"10%": [r.get("tgtcthue"), r.get("tgtthue")]}
        for rate, v in rates.items():
            val, tax = _num(v[0]), _num(v[1])
            ru = (rate or "").upper().replace(" ", "")
            pct = _rate_pct(ru)
            if ru in ("KCT",):
                t["26"] += val
            elif ru in ("KKKNT",):
                t["32a"] += val
            elif pct == 0:
                t["29"] += val
            elif pct == 5:
                t["30"] += val
                t["31"] += tax
            elif pct is not None:  # 8%, 10% và các mức khác
                t["32"] += val
                t["33"] += tax
    t["34"] = t["26"] + t["29"] + t["30"] + t["32"] + t["32a"]
    t["35"] = t["31"] + t["33"]
    return t


def write_reconcile(root: Path, tk_vals: dict[str, float], period: tuple[str, int, int],
                    rows: list[dict[str, Any]], src_name: str) -> tuple[Path, list[dict[str, Any]]]:
    kieu, ky, nam = period
    inv = invoice_totals(rows)
    wb = Workbook()
    ws = wb.active
    ws.title = "Doi chieu"
    label = f"Quý {ky}/{nam}" if kieu == "Q" else f"Tháng {ky}/{nam}"
    ws.cell(1, 1, f"ĐỐI CHIẾU TỜ KHAI 01/GTGT {label} – file {src_name}").font = Font(bold=True, size=13)
    ws.cell(2, 1, f"Hóa đơn: {len(rows)} dòng đã tải (bỏ hóa đơn đã bị thay thế / bị hủy). "
                  "Bán ra 8% được cộng vào chỉ tiêu [32]/[33].")
    heads = [("Chỉ tiêu", 9), ("Nội dung", 40), ("Tờ khai", 18), ("Theo hóa đơn", 18), ("Chênh lệch", 18), ("Kết quả", 14)]
    for i, (h, w) in enumerate(heads, 1):
        c = ws.cell(4, i, h)
        c.font, c.fill = BOLD, FILL
        ws.column_dimensions[c.column_letter].width = w
    result = []
    r = 5
    for k, name in CHI_TIEU:
        a, b = tk_vals.get(k), inv.get(k, 0.0)
        diff = None if a is None else round(b - a)
        ok = "Không có trên TK" if a is None else ("Khớp" if abs(diff) < 1000 else "LỆCH")
        for i, v in enumerate([f"[{k}]", name, a, b, diff, ok], 1):
            c = ws.cell(r, i, v)
            if i in (3, 4, 5):
                c.number_format = "#,##0"
            if ok == "LỆCH":
                c.fill = WARN
        result.append({"ct": k, "name": name, "to_khai": a, "hoa_don": b, "chenh_lech": diff, "ket_qua": ok})
        r += 1
    ws2 = wb.create_sheet("Hoa don trong ky")
    _header(ws2, [("Loại", 26), ("Tháng", 9), ("Ngày lập", 11), ("Ký hiệu", 10), ("Số HĐ", 10), ("MST bán", 14),
                  ("MST mua", 14), ("Chưa thuế", 15), ("Thuế", 13), ("Trạng thái", 16), ("Theo thuế suất", 40)])
    for i, x in enumerate(rows, 2):
        vals = [x.get("loai"), x.get("thang"), x.get("ngay_lap"), x.get("khhdon"), x.get("shdon"), x.get("nbmst"),
                x.get("nmmst"), x.get("tgtcthue"), x.get("tgtthue"), x.get("tthai"),
                json.dumps(x.get("_rates") or {}, ensure_ascii=False)]
        for j, v in enumerate(vals, 1):
            c = ws2.cell(i, j, v)
            if j in (8, 9):
                c.number_format = "#,##0"
    suffix = f"Q{ky}-{nam}" if kieu == "Q" else f"{nam}-{ky:02d}"
    return _save(wb, root / f"DoiChieu_ToKhai_{suffix}.xlsx"), result
