"""Hóa đơn mua vào do NGÂN HÀNG xuất (phí dịch vụ ngân hàng, phí chuyển tiền...).

Quy ước của văn phòng: hóa đơn đầu vào của ngân hàng KHÔNG đưa vào bảng kê mua vào và KHÔNG
tính vào tờ khai 01/GTGT / dự báo VAT, áp dụng cho mọi công ty.

- is_bank_seller(): nhận diện người bán là ngân hàng theo tên (NGÂN HÀNG ..., VIETCOMBANK, VPBANK...).
- strip_bank_rows(): tách các dòng ngân hàng khỏi sheet chính của bảng kê Excel tải từ cổng thuế,
  chuyển sang sheet riêng "HĐ ngân hàng (đã loại)" (giữ nguyên dòng tiêu đề cột để phần mềm vẫn đọc
  đủ danh sách hóa đơn khi đối chiếu số lượng với cổng).
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Any

BANK_SHEET = "HĐ ngân hàng (đã loại)"
BANK_NOTE = "Hóa đơn của ngân hàng – không đưa vào bảng kê mua vào và tờ khai"

_BANK_RE = re.compile(r"NGÂN\s*HÀNG|NGAN\s*HANG|\b[A-Z]*BANK\b")


def _norm(s: Any) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(s or "")).upper()).strip()


def is_bank_seller(name: Any) -> bool:
    """True nếu tên người bán là ngân hàng."""
    return bool(_BANK_RE.search(_norm(name)))


def is_bank_purchase(rec: dict[str, Any]) -> bool:
    """Bản ghi hóa đơn (theo tên trường API cổng thuế) là hóa đơn mua vào của ngân hàng."""
    return rec.get("huong") == "purchase" and is_bank_seller(rec.get("nbten"))


def strip_bank_rows(path: Path) -> int:
    """Chuyển các dòng người bán là ngân hàng sang sheet riêng. Trả về số dòng đã chuyển (0 nếu không có).
    Gọi nhiều lần vẫn an toàn (lần sau không còn dòng ngân hàng trong sheet chính)."""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    ws = wb.worksheets[0]
    if ws.title == BANK_SHEET:
        return 0
    head_row = name_col = None
    for r in range(1, min(ws.max_row, 20) + 1):
        for c in range(1, ws.max_column + 1):
            v = _norm(ws.cell(r, c).value)
            if v.startswith("TÊN NGƯỜI BÁN"):
                head_row, name_col = r, c
                break
        if head_row:
            break
    if not head_row:
        return 0
    bank_rows = [r for r in range(head_row + 1, ws.max_row + 1) if is_bank_seller(ws.cell(r, name_col).value)]
    if not bank_rows:
        return 0
    ncol = ws.max_column
    out = wb[BANK_SHEET] if BANK_SHEET in wb.sheetnames else wb.create_sheet(BANK_SHEET)
    if out.max_row <= 1 and out.cell(1, 1).value is None:
        out.cell(1, 1, "HÓA ĐƠN MUA VÀO CỦA NGÂN HÀNG – ĐÃ LOẠI KHỎI BẢNG KÊ MUA VÀO, KHÔNG TÍNH VÀO TỜ KHAI")
        for c in range(1, ncol + 1):
            out.cell(3, c, ws.cell(head_row, c).value)
            out.column_dimensions[out.cell(3, c).column_letter].width = \
                ws.column_dimensions[ws.cell(head_row, c).column_letter].width or 14
    for r in bank_rows:
        out.append([ws.cell(r, c).value for c in range(1, ncol + 1)])
    for r in reversed(bank_rows):
        ws.delete_rows(r)
    # đánh lại số thứ tự cột STT của sheet chính
    if _norm(ws.cell(head_row, 1).value) == "STT":
        n = 0
        for r in range(head_row + 1, ws.max_row + 1):
            if any(ws.cell(r, c).value not in (None, "") for c in range(2, ncol + 1)):
                n += 1
                ws.cell(r, 1, n)
    wb.save(path)
    return len(bank_rows)
