"""Test offline cho chức năng nhập file (XML / ZIP / Excel)."""
import io
import sys
import zipfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import Workbook  # noqa: E402

from app.gdt.xml_parser import parse_invoice_xml  # noqa: E402
from app.sources.file_import import (  # noqa: E402
    iter_items, read_excel, row_from_excel, row_from_xml,
)
from tests.test_offline import SAMPLE  # noqa: E402

OWNER = "0109876543"  # = MST người mua trong SAMPLE


def _zip(entries):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for n, d in entries.items():
            zf.writestr(n, d)
    return buf.getvalue()


def make_excel() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append(["DANH SÁCH HÓA ĐƠN"])
    ws.append(["Từ ngày 01/10/2026 đến ngày 31/10/2026"])
    ws.append([])
    ws.append(["STT", "Ký hiệu mẫu số", "Ký hiệu hóa đơn", "Số hóa đơn", "Ngày lập",
               "MST người bán/MST người xuất hàng", "Tên người bán/Tên người xuất hàng",
               "Tổng tiền chưa thuế", "Tổng tiền thuế", "Tổng tiền chiết khấu thương mại",
               "Tổng tiền thanh toán", "Trạng thái hóa đơn"])
    ws.append([1, 1, "C24TAA", 123, "05/10/2026", "0312345678", "Công ty Bán",
               "700.000", "56.000", 0, "756.000", "Hóa đơn mới"])
    ws.append([2, 1, "C24TBB", "00000045", datetime(2026, 10, 7), "0311111111", "Công ty B",
               1000000, 100000, 0, 1100000, "Hóa đơn mới"])
    ws.append(["Tổng cộng", None, None, None, None, None, None, 1700000, 156000, 0, 1856000])
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def test_iter_items_nested_zip():
    inner = _zip({"invoice.xml": SAMPLE, "details.js": b"x"})
    outer = _zip({"HD_1.zip": inner, "readme.txt": b"hi", "b/invoice.xml": SAMPLE})
    kinds = [(k, n) for k, n, _ in iter_items("tai_ve.zip", outer)]
    assert [k for k, _ in kinds].count("xml") == 2
    assert ("skip", "tai_ve.zip/readme.txt") in kinds


def test_xml_row_direction():
    p = parse_invoice_xml(SAMPLE)
    r = row_from_xml(OWNER, p, "sold")  # owner là người mua -> purchase dù gợi ý sai
    assert r["direction"] == "purchase" and r["nbmst"] == "0312345678"
    assert r["shdon"] == "123" and r["tdlap"] == "2024-10-05T00:00:00+07:00"
    assert row_from_xml("0312345678", p, "purchase")["direction"] == "sold"


def test_excel():
    data = make_excel()
    kinds = [k for k, _, _ in iter_items("ds.xlsx", data)]
    assert kinds == ["excel"]
    rows = read_excel(data, "ds.xlsx")
    assert len(rows) == 2, rows
    r1 = row_from_excel(OWNER, rows[0], "purchase")
    assert (r1["khmshdon"], r1["khhdon"], r1["shdon"]) == ("1", "C24TAA", "123")
    assert r1["tgtttbso"] == 756000 and r1["tdlap"] == "2026-10-05T00:00:00+07:00"
    assert r1["nmmst"] == OWNER and r1["direction"] == "purchase"
    r2 = row_from_excel(OWNER, rows[1], "purchase")
    assert r2["shdon"] == "45" and r2["tdlap"] == "2026-10-07T00:00:00+07:00"


def test_pdf_validation():
    from app.providers import InvoiceRef, pdf_matches_invoice
    from app.providers.fallback import render_pdf
    pdf = render_pdf(parse_invoice_xml(SAMPLE))
    ok = InvoiceRef("0312345678", OWNER, "1", "C24TAA", "123", None, None, None)
    wrong_num = InvoiceRef("0312345678", OWNER, "1", "C24TAA", "999", None, None, None)
    wrong_seller = InvoiceRef("0399999999", OWNER, "1", "K99XYZ", "123", None, None, None)
    assert pdf_matches_invoice(pdf, ok)
    assert not pdf_matches_invoice(pdf, wrong_num)
    assert not pdf_matches_invoice(pdf, wrong_seller)
    assert not pdf_matches_invoice(b"%PDF-1.4 rac", ok)


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
