"""Xuất "Dự báo thuế GTGT" gửi khách hàng (file Excel 1 trang theo mẫu của đơn vị kế toán dịch vụ) + lời nhắn gửi kèm.

Số liệu lấy từ kết quả dự báo VAT đã lập (tờ khai 01/GTGT):
  Kỳ trước chuyển sang = -[22] · Doanh thu HHDV đầu vào = [23] · Thuế GTGT đầu vào = [25] · Doanh thu bán ra = [34]
  Thuế GTGT đầu ra = [35] · Phát sinh thuế GTGT trong kỳ = [36] · Lũy kế = [36] - [22] + [37] - [38] - [39a]
  (Lũy kế > 0: thuế phải nộp; < 0: thuế còn được khấu trừ, hiện trong ngoặc).
"""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

NUM_FMT = '#,##0;(#,##0);"-"'


def _d(s: str | None) -> date | None:
    try:
        return date.fromisoformat(str(s)[:10])
    except (TypeError, ValueError):
        return None


def short_name(ten: str) -> str:
    """'CÔNG TY TNHH ISTEEL' -> 'ISTEEL' (tên gọi ngắn trong lời nhắn)."""
    s = re.sub(r"\s+", " ", (ten or "").strip())
    s = re.sub(r"^(CÔNG TY|CTY|Công ty|công ty)\s+", "", s, flags=re.I)
    s = re.sub(r"^(TNHH\s+(MTV|MỘT THÀNH VIÊN|HAI THÀNH VIÊN TRỞ LÊN)?|CỔ PHẦN|CP|HỢP DANH)\s*", "", s, flags=re.I)
    return s.strip() or ten


def vnd(v: float) -> str:
    return f"{abs(round(v)):,}".replace(",", ".")


def build(res: dict[str, Any], sender: str = "ATAX", ten: str | None = None, ten_ngan: str | None = None,
          ngay: date | None = None) -> dict[str, Any]:
    ct = res["ct"]
    monthly = res["kieu"] == "M"
    ky, nam = int(res["ky"]), int(res["nam"])
    d_to = _d(res.get("to")) or date.today()
    if ngay is None:  # ngày dự báo = ngày lập dự báo (không quá cuối kỳ)
        ngay = min(_d(res.get("created")) or date.today(), d_to)
    ten = (ten or company_name(res)).strip()
    ten_ngan = (ten_ngan or short_name(ten)).strip()
    sender = (sender or "ATAX").strip()
    ky_txt = f"tháng {ky}" if monthly else f"quý {ky}"           # tháng 10
    ky_cap = f"Tháng {ky}" if monthly else f"Quý {ky}"           # Tháng 10
    ky_so = f"{ky:02d}/{nam}" if monthly else f"{ky}/{nam}"      # 10/2026
    ngay_txt = f"{ngay:%d/%m/%Y}"
    ngay_ngan = f"{ngay:%d/%m}"
    vals = {
        "truoc": -float(ct.get("22", 0)), "dt_vao": float(ct.get("23", 0)), "vat_vao": float(ct.get("25", 0)),
        "dt_ra": float(ct.get("34", 0)), "vat_ra": float(ct.get("35", 0)), "ps": float(ct.get("36", 0)),
    }
    vals["luy_ke"] = (float(ct.get("36", 0)) - float(ct.get("22", 0)) + float(ct.get("37", 0))
                      - float(ct.get("38", 0)) - float(ct.get("39a", 0)))
    phai_nop = vals["luy_ke"] > 0
    return {"ten": ten, "ten_ngan": ten_ngan, "sender": sender, "ngay": ngay, "ngay_txt": ngay_txt,
            "ngay_ngan": ngay_ngan, "ky_txt": ky_txt, "ky_cap": ky_cap, "ky_so": ky_so, "nam": nam, "vals": vals,
            "phai_nop": phai_nop, "den_cuoi_ky": ngay < d_to, "cuoi_ky": f"{d_to:%d/%m}"}


def message(b: dict[str, Any]) -> str:
    """Lời nhắn gửi khách – luôn đủ 3 phần theo mẫu: kỳ + thời điểm, số thuế, nhắc bổ sung đến cuối kỳ."""
    v = b["vals"]["luy_ke"]
    loai = "phải nộp" if b["phai_nop"] else "được khấu trừ"
    return "\n".join([
        f"Dạ anh/chị, em xin gửi Doanh nghiệp dự báo thuế GTGT {b['ky_txt']}/{b['nam']} "
        f"(tính tới thời điểm {b['ngay_txt']}) như sau:",
        f"- Đến ngày {b['ngay_txt']} dự kiến Công ty {b['ten_ngan']} có số thuế GTGT {loai} trong "
        f"{b['ky_txt']}/{b['nam']} là: {vnd(v)} đ",
        f"Công ty dự tính phát sinh doanh thu và chi phí đến ngày {b['cuoi_ky']} để kịp bổ sung sớm ạ. Em xin cảm ơn ạ",
    ])


def company_name(res: dict[str, Any]) -> str:
    """Tên doanh nghiệp: từ kết quả dự báo; chưa có thì lấy từ hóa đơn đã tải (người bán của HĐ bán ra /
    người mua của HĐ mua vào trùng MST)."""
    ten = (res.get("company") or {}).get("ten") or ""
    if ten.strip():
        return ten.strip()
    try:
        from .config import settings
        from .vat import load_rows
        mst = str(res["mst"])
        import json
        root = Path(settings.output_dir) / mst
        rows, _, _ = load_rows(root, res["kieu"], int(res["ky"]), int(res["nam"]))
        if not rows:  # kỳ này chưa có hóa đơn: tìm tên ở các tháng khác của MST này (mới nhất trước)
            for f in sorted(root.glob("*/.dulieu_*.json"), reverse=True)[:6]:
                try:
                    rows += json.loads(f.read_text(encoding="utf-8")).get("rows", [])
                except (OSError, ValueError):
                    continue
        from collections import Counter
        names = Counter()
        for r in rows:
            for m, t in (("nbmst", "nbten"), ("nmmst", "nmten")):
                if str(r.get(m) or "").replace("-", "") == mst and str(r.get(t) or "").strip():
                    names[str(r[t]).strip()] += 1
        if names:
            return names.most_common(1)[0][0]
    except Exception:  # noqa: BLE001
        pass
    return f"MST {res['mst']}"


def write_excel(b: dict[str, Any], dest: Path) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "Du bao GTGT"
    F = "Times New Roman"
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    yellow = PatternFill("solid", fgColor="FFFF00")
    for col, w in zip("ABCDEFGH", (15.5, 15.5, 20, 30, 32, 17, 17, 14.5)):
        ws.column_dimensions[col].width = w

    def put(ref: str, v: Any, size: int = 11, bold: bool = False, h: str | None = None, fmt: str | None = None,
            wrap: bool = False) -> Any:
        c = ws[ref]
        c.value = v
        c.font = Font(name=F, size=size, bold=bold)
        c.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
        if fmt:
            c.number_format = fmt
        return c

    v = b["vals"]
    # ----- phần đầu
    put("A1", "DỰ BÁO THUẾ GTGT", 12, True)
    for col in "AB":
        ws[f"{col}1"].border = Border(bottom=thin)
    put("A2", f"Kính gửi: Quý {b['ten']}")
    put("A3", f"{b['sender']} xin gửi Quý Công ty dự báo thuế GTGT {b['ky_txt']} năm {b['nam']}, như sau:")
    put("D4", "DỰ BÁO THUẾ GTGT HÀNG THÁNG" if b["ky_cap"].startswith("Tháng") else "DỰ BÁO THUẾ GTGT HÀNG QUÝ")
    put("D5", "Ngày dự báo:")
    put("E5", b["ngay_txt"], h="right")
    # ----- bảng số liệu
    heads = ["", "Kỳ trước chuyển sang", "Doanh thu hàng hóa đầu vào", "Thuế GTGT đầu vào", "Doanh thu bán ra",
             "Thuế GTGT đầu ra", "Phát sinh thuế GTGT trong kỳ", "Lũy kế"]
    for i, t in enumerate(heads):
        c = put(f"{'ABCDEFGH'[i]}7", t, bold=True, h="center", wrap=True)
        c.border = box
    ws.row_dimensions[7].height = 45
    put("A8", b["nam"], bold=True, h="center")
    row9 = [b["ky_cap"], v["truoc"], v["dt_vao"], v["vat_vao"], v["dt_ra"], v["vat_ra"], v["ps"], v["luy_ke"]]
    for i, x in enumerate(row9):
        col = "ABCDEFGH"[i]
        ws[f"{col}8"].border = box
        c = put(f"{col}9", round(x) if i else x, size=10 if col == "G" else 11, h=None if i == 0 else "right",
                fmt=None if i == 0 else NUM_FMT)
        c.border = box
    ws["H9"].fill = yellow
    # ----- khung lời gửi
    put("B11", "Kính gửi Quý Công Ty", 13)
    tg = f" (tới thời điểm {b['ngay_ngan']})"
    put("B12", f"{b['sender']} xin gửi Quý Công ty dự báo thuế GTGT {b['ky_txt']}{tg} năm {b['nam']}, như sau", 13)
    put("B13", "DT mua vào", 13, True)
    put("C13", round(v["dt_vao"]), 13, True, "right", NUM_FMT)
    put("D13", "VAT mua vào", 13, True)
    put("E13", round(v["vat_vao"]), 13, True, "right", NUM_FMT)
    put("B14", "DT bán ra:", 13, True)
    put("C14", round(v["dt_ra"]), 13, True, "right", NUM_FMT)
    put("D14", "VAT bán ra", 13, True)
    put("E14", round(v["vat_ra"]), 13, True, "right", NUM_FMT)
    loai = "Tổng số thuế GTGT phải nộp" if b["phai_nop"] else "Tổng số thuế còn được khấu trừ"
    put("B15", f"{loai} {b['ky_txt']} tạm tính tới thời điểm ({b['ngay_txt']}):", 13)
    c = put("F15", round(v["luy_ke"]), 14, True, "right", NUM_FMT)
    c.fill = yellow
    put("B16", "Trân trọng cảm ơn!", 13, True)
    for r in range(11, 17):
        ws.row_dimensions[r].height = 21
        for col in "BCDEF":
            c = ws[f"{col}{r}"]
            c.border = Border(left=thin if col == "B" else None, right=thin if col == "F" else None,
                              top=thin if r == 11 else None, bottom=thin if r == 16 else None)
    # in vừa 1 trang ngang
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(dest)
    except PermissionError:  # file cũ đang mở trong Excel
        dest = dest.with_name(f"{dest.stem}_{datetime.now():%H%M%S}{dest.suffix}")
        wb.save(dest)
    return dest


def export(res: dict[str, Any], write: bool = True, **kw: Any) -> dict[str, Any]:
    """write=False: chỉ trả về thông tin tự điền + lời nhắn (để hiện sẵn trên màn hình), không ghi file."""
    from .vat import out_folder
    b = build(res, **kw)
    out = {"file": "", "message": message(b), "sender": b["sender"], "ten": b["ten"], "ten_ngan": b["ten_ngan"],
           "ngay": b["ngay"].isoformat(), "luy_ke": round(b["vals"]["luy_ke"]), "phai_nop": b["phai_nop"]}
    if write:
        tag = res["key"].replace("-", "_")
        out["file"] = str(write_excel(b, out_folder(res["mst"], res["key"]) / f"DuBao_GuiKhachHang_{tag}.xlsx"))
    return out
