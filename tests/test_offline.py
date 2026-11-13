"""Test không cần mạng: python -m pytest tests  (hoặc python tests/test_offline.py)."""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.gdt.client import month_chunks  # noqa: E402
from app.gdt.xml_parser import parse_invoice_xml  # noqa: E402
from app.providers import InvoiceRef, detect_provider, load_providers, _fill  # noqa: E402
from app.providers.fallback import render_pdf  # noqa: E402

SAMPLE = """<?xml version="1.0" encoding="UTF-8"?>
<HDon><DLHDon Id="data">
 <TTChung><PBan>2.0.0</PBan><THDon>HÓA ĐƠN GIÁ TRỊ GIA TĂNG</THDon><KHMSHDon>1</KHMSHDon>
  <KHHDon>C24TAA</KHHDon><SHDon>123</SHDon><NLap>2024-10-05</NLap><DVTTe>VND</DVTTe>
  <HTTToan>TM/CK</HTTToan><MSTTCGP>0101243150</MSTTCGP>
  <TTKhac><TTin><TTruong>Mã tra cứu</TTruong><KDLieu>string</KDLieu><DLieu>ABC123XYZ</DLieu></TTin>
   <TTin><TTruong>Link tra cứu</TTruong><KDLieu>string</KDLieu><DLieu>https://www.meinvoice.vn/tra-cuu</DLieu></TTin></TTKhac>
 </TTChung>
 <NDHDon>
  <NBan><Ten>Công ty Bán</Ten><MST>0312345678</MST><DChi>Q1, TP.HCM</DChi></NBan>
  <NMua><Ten>Công ty Mua</Ten><MST>0109876543</MST><DChi>Hà Nội</DChi></NMua>
  <DSHHDVu><HHDVu><STT>1</STT><THHDVu>Giấy A4</THHDVu><DVTinh>Ram</DVTinh><SLuong>10</SLuong>
   <DGia>70000</DGia><ThTien>700000</ThTien><TSuat>8%</TSuat></HHDVu></DSHHDVu>
  <TToan><TgTCThue>700000</TgTCThue><TgTThue>56000</TgTThue><TgTTTBSo>756000</TgTTTBSo>
   <TgTTTBChu>Bảy trăm năm mươi sáu nghìn đồng</TgTTTBChu></TToan>
 </NDHDon>
</DLHDon><MCCQT>00ABCDEF1234567890</MCCQT>
<DSCKS><NBan><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><X>http://www.w3.org/x</X></Signature></NBan></DSCKS>
</HDon>""".encode("utf-8")


def test_parse():
    p = parse_invoice_xml(SAMPLE)
    assert p["khhdon"] == "C24TAA" and p["shdon"] == "123"
    assert p["nguoi_ban"]["mst"] == "0312345678"
    assert p["tong_thanh_toan"] == 756000
    assert p["lookup_code"] == "ABC123XYZ"
    assert p["lookup_urls"] == ["https://www.meinvoice.vn/tra-cuu"]
    assert p["mccqt"] == "00ABCDEF1234567890"
    assert p["hang_hoa"][0]["so_luong"] == 10


def test_detect_and_fill():
    p = parse_invoice_xml(SAMPLE)
    ref = InvoiceRef("0312345678", "0109876543", "1", "C24TAA", "123", p["msttcgp"], p["mccqt"],
                     p["lookup_code"], p["lookup_urls"])
    prov = detect_provider(ref, load_providers())
    assert prov and prov.code == "misa"
    assert _fill("x?c={lookup_code}", ref.vars()) == "x?c=ABC123XYZ"
    assert _fill("x?c={lookup_code}", {"lookup_code": None}) is None


def test_month_chunks():
    ch = list(month_chunks(date(2024, 11, 15), date(2025, 2, 3)))
    assert ch[0] == (date(2024, 11, 15), date(2024, 11, 30))
    assert ch[1] == (date(2024, 12, 1), date(2024, 12, 31))
    assert ch[-1] == (date(2025, 2, 1), date(2025, 2, 3))
    assert len(ch) == 4


def test_render_pdf():
    pdf = render_pdf(parse_invoice_xml(SAMPLE))
    assert pdf[:5] == b"%PDF-"
    Path(__file__).with_name("_sample_fallback.pdf").write_bytes(pdf)


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
