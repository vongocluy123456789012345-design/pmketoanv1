"""Test xuất dự báo thuế GTGT gửi khách hàng: python -m pytest tests/test_vat_client.py"""
import sys
import tempfile
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import vat_client as vc  # noqa: E402

RES = {"mst": "0100000000", "kieu": "M", "ky": 10, "nam": 2026, "key": "2026-10", "to": "2026-10-31",
       "created": "2026-10-24T10:00:00", "company": {"ten": "CÔNG TY TNHH ISTEEL"},
       "ct": {"22": 573998483, "23": 20471224273, "25": 2040737410, "34": 23961703621, "35": 2396170366,
              "36": 355432956, "37": 0, "38": 0, "39a": 0}}


def test_numbers_and_message():
    b = vc.build(RES)
    assert b["ngay"] == date(2026, 10, 24) and b["ten_ngan"] == "ISTEEL"
    assert round(b["vals"]["luy_ke"]) == -218565527 and not b["phai_nop"]
    m = vc.message(b)
    assert "tháng 10/2026 (tính tới thời điểm 24/10/2026)" in m
    assert "Công ty ISTEEL có số thuế GTGT được khấu trừ trong tháng 10/2026 là: 218.565.527 đ" in m
    assert "đến ngày 31/10" in m


def test_excel_layout():
    from openpyxl import load_workbook
    with tempfile.TemporaryDirectory() as tmp:
        p = vc.write_excel(vc.build(RES), Path(tmp) / "x.xlsx")
        ws = load_workbook(p).active
        assert ws["A1"].value == "DỰ BÁO THUẾ GTGT" and ws["A2"].value == "Kính gửi: Quý CÔNG TY TNHH ISTEEL"
        assert [ws[f"{c}9"].value for c in "BCDEFGH"] == [-573998483, 20471224273, 2040737410, 23961703621,
                                                            2396170366, 355432956, -218565527]
        assert ws["H9"].fill.fgColor.rgb.endswith("FFFF00") and ws["F15"].value == -218565527
        assert ws["B15"].value.startswith("Tổng số thuế còn được khấu trừ tháng 10 tạm tính tới thời điểm (24/10/2026)")


def test_payable():
    res = RES | {"ct": RES["ct"] | {"22": 0}}
    b = vc.build(res, sender="ATAX", ngay=date(2026, 10, 31))
    assert b["phai_nop"] and "phải nộp" in vc.message(b) and not b["den_cuoi_ky"]
