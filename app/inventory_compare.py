"""2 bản so sánh tên hàng ghép mã (kèm báo cáo Tổng hợp tồn kho):
  * Mua vào ↔ Đầu kỳ: tên hàng nào ghép đúng mã chuẩn đầu kỳ, tên nào phải tạo mã mới.
  * Bán ra ↔ Mua vào ↔ Đầu kỳ: hàng bán lấy từ đâu (tồn đầu kỳ / vừa mua), còn đủ số lượng không.
"""
from __future__ import annotations

import io
from typing import Any

_GREEN, _YELLOW, _ORANGE, _RED = "E2EFDA", "FFF2CC", "FCE4D6", "F8CBAD"


def _sheet(ws, heads: list[tuple[str, int]], data: list[list[Any]], fills: list[str | None], title: str, note: str) -> None:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.cell(1, 1, title).font = Font(bold=True, size=13)
    ws.cell(2, 1, note).font = Font(italic=True)
    for c, (h, w) in enumerate(heads, 1):
        cell = ws.cell(3, c, h)
        cell.font, cell.fill, cell.border = Font(bold=True), PatternFill("solid", fgColor="DDEBF7"), box
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[3].height = 32
    for r, (vals, fill) in enumerate(zip(data, fills), 4):
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = box
            cell.alignment = Alignment(vertical="top", wrap_text=isinstance(v, str) and len(v) > 30)
            if isinstance(v, (int, float)) and not isinstance(v, bool) and c > 1:
                cell.number_format = "#,##0.##"
            if fill:
                cell.fill = PatternFill("solid", fgColor=fill)
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(heads))}{max(4, len(data) + 3)}"


def _save(wb) -> bytes:
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def compare_buy(buy: list[dict[str, Any]], items: list[dict[str, Any]]) -> bytes:
    from openpyxl import Workbook
    opening = {i["ma"].upper(): i for i in items if not i.get("moi")}
    grp: dict[tuple[str, str], dict[str, Any]] = {}
    for ln in buy:
        g = grp.setdefault((ln["ten"], ln["ma_ghep"]), {"dvt": ln["dvt"], "sl": 0.0, "tt": 0.0, "cach": ln.get("cach", "")})
        g["sl"] += ln["sl"]
        g["tt"] += ln["tt"] - ln["ck"]
    data, fills = [], []
    stat = {"dung": 0, "kt": 0, "moi": 0}
    for (ten, ma), g in sorted(grp.items(), key=lambda kv: (kv[0][1].upper() in opening, kv[0][1].upper(), kv[0][0])):
        op = opening.get(ma.upper())
        if not op:
            kq, fill, k = "TẠO MỚI (chưa có trong đầu kỳ)", _ORANGE, "moi"
        elif "KIỂM TRA" in g["cach"]:
            kq, fill, k = "Ghép mã đầu kỳ – CẦN KIỂM TRA", _YELLOW, "kt"
        else:
            kq, fill, k = "GHÉP ĐÚNG mã đầu kỳ", _GREEN, "dung"
        stat[k] += 1
        data.append([len(data) + 1, ten, g["dvt"], g["sl"], round(g["tt"]), ma, op["ten"] if op else "",
                     op["dvt"] if op else "", op["sl"] if op else None, kq, g["cach"]])
        fills.append(fill)
    wb = Workbook()
    ws = wb.active
    ws.title = "Mua vao - Dau ky"
    _sheet(ws, [("STT", 6), ("Tên hàng trên hóa đơn MUA VÀO", 52), ("ĐVT", 8), ("SL mua", 10), ("Giá trị mua", 15),
                ("Mã ghép", 18), ("Tên hàng ĐẦU KỲ (mã chuẩn)", 46), ("ĐVT đầu kỳ", 9), ("SL tồn đầu kỳ", 11),
                ("Kết quả", 30), ("Cách ghép", 42)], data, fills, "SO SÁNH TÊN HÀNG GHÉP: MUA VÀO ↔ ĐẦU KỲ",
           f"Ghép đúng mã đầu kỳ: {stat['dung']} tên hàng | Ghép mã đầu kỳ cần kiểm tra: {stat['kt']} | Tạo mới: "
           f"{stat['moi']}   (xanh = đúng, vàng = kiểm tra, cam = tạo mới)")
    # mỗi mã gom những tên hóa đơn nào
    by: dict[str, list[str]] = {}
    for (ten, ma), _g in grp.items():
        by.setdefault(ma.upper(), []).append(ten)
    d2, f2 = [], []
    for ma, names in sorted(by.items()):
        op = opening.get(ma)
        d2.append([len(d2) + 1, op["ma"] if op else ma, op["ten"] if op else "(mã MỚI)", len(names), " | ".join(names)])
        f2.append(None if op else _ORANGE)
    _sheet(wb.create_sheet("Theo ma"), [("STT", 6), ("Mã", 18), ("Tên hàng ĐẦU KỲ", 46), ("Số tên HĐ gom vào", 10),
                                        ("Các tên trên hóa đơn mua vào", 90)], d2, f2,
           "MỖI MÃ GOM NHỮNG TÊN HÀNG MUA VÀO NÀO", "Cam = mã mới tạo từ hóa đơn mua vào (chưa có trong đầu kỳ)")
    return _save(wb)


def compare_sell(sell: list[dict[str, Any]], buy: list[dict[str, Any]], items: list[dict[str, Any]],
                 rows: list[dict[str, Any]]) -> bytes:
    from openpyxl import Workbook
    opening = {i["ma"].upper(): i for i in items if not i.get("moi")}
    names_buy: dict[str, list[str]] = {}
    for ln in buy:
        lst = names_buy.setdefault(ln["ma_ghep"].upper(), [])
        if ln["ten"] not in lst:
            lst.append(ln["ten"])
    per: dict[str, dict[str, float]] = {}
    for r in rows:                                   # cộng theo mã (gộp các kho)
        p = per.setdefault(r["ma"].upper(), {"dau": 0.0, "mua": 0.0, "ban": 0.0, "cuoi": 0.0})
        p["dau"] += r["sl_dau"]
        p["mua"] += r["sl_nhap"] - r.get("sl_tra", 0)
        p["ban"] += r["sl_xuat"]
        p["cuoi"] += r["sl_cuoi"]
    grp: dict[tuple[str, str], dict[str, Any]] = {}
    for ln in sell:
        g = grp.setdefault((ln["ten"], ln["ma_ghep"]), {"dvt": ln["dvt"], "sl": 0.0, "tt": 0.0, "cach": ln.get("cach", "")})
        g["sl"] += ln["sl"]
        g["tt"] += ln["tt"]
    data, fills = [], []
    stat = {"chac": 0, "kt": 0, "vuot": 0}
    for (ten, ma), g in sorted(grp.items(), key=lambda kv: (kv[0][1].upper(), kv[0][0])):
        k = ma.upper()
        op, p, nb = opening.get(k), per.get(k, {}), names_buy.get(k, [])
        nguon = "Đầu kỳ + Mua trong kỳ" if op and nb else "Đầu kỳ" if op else "Mua trong kỳ (mã mới)" if nb else ""
        if p.get("cuoi", 0) < 0:
            kq, fill, s = "BÁN VƯỢT số đang có (đầu kỳ + mua)", _RED, "vuot"
        elif "KIỂM TRA" in g["cach"]:
            kq, fill, s = "CẦN KIỂM TRA mã ghép", _YELLOW, "kt"
        else:
            kq, fill, s = "Khớp chắc chắn", _GREEN, "chac"
        stat[s] += 1
        data.append([len(data) + 1, ten, g["dvt"], g["sl"], round(g["tt"]), ma, op["ten"] if op else "",
                     " | ".join(nb[:4]) + (" …" if len(nb) > 4 else ""), nguon, p.get("dau", 0), p.get("mua", 0),
                     p.get("ban", 0), p.get("cuoi", 0), kq, g["cach"]])
        fills.append(fill)
    wb = Workbook()
    ws = wb.active
    ws.title = "Ban ra - Mua vao - Dau ky"
    _sheet(ws, [("STT", 6), ("Tên hàng trên hóa đơn BÁN RA", 48), ("ĐVT", 8), ("SL bán", 9), ("Doanh thu", 14),
                ("Mã ghép", 18), ("Tên hàng ĐẦU KỲ", 40), ("Tên hàng MUA VÀO cùng mã", 48), ("Nguồn hàng", 18),
                ("SL đầu kỳ (mã)", 10), ("SL mua (mã)", 10), ("SL bán (mã)", 10), ("SL còn (mã)", 10),
                ("Kết quả", 28), ("Cách ghép", 42)], data, fills, "SO SÁNH TÊN HÀNG GHÉP: BÁN RA ↔ MUA VÀO ↔ ĐẦU KỲ",
           f"Khớp chắc chắn: {stat['chac']} tên hàng | Cần kiểm tra: {stat['kt']} | Bán vượt số đang có: {stat['vuot']}   "
           "(xanh = chắc chắn, vàng = kiểm tra, đỏ = bán vượt; số lượng (mã) cộng theo cả mã hàng)")
    return _save(wb)
