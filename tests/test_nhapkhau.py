from app import nhapkhau as nk
from app import vat

SAMPLE = """TỜ KHAI HÀNG HÓA NHẬP KHẨU (THÔNG QUAN)
Số tờ khai 107123456789 Số tờ khai đầu tiên Mã phân loại kiểm tra 1
Ngày đăng ký 05/10/2026 09:15:22
Người nhập khẩu Mã 0313749848 Tên CÔNG TY TNHH ABC
Người xuất khẩu
Tên: SHENZHEN XYZ TRADING CO., LTD
Mã sắc thuế Tên sắc thuế Tổng tiền thuế
N Thuế NK 3.000.000 VND
V Thuế GTGT 11.300.000 VND
Thuế GTGT Trị giá tính thuế 113.000.000 Thuế suất 10% Số tiền thuế 11.300.000
"""


def test_parse_vnaccs_printout_without_ai():
    rows = nk.parse_text(SAMPLE)
    assert len(rows) == 1
    r = rows[0]
    assert r["so_tk"] == "107123456789" and r["ngay"] == "05/10/2026"
    assert r["ncc"].startswith("SHENZHEN XYZ")
    assert r["thue_gtgt"] == 11300000 and r["thue_suat"] == "10%"
    assert r["tri_gia_tinh_thue"] == 113000000      # suy ra từ thuế ÷ thuế suất


def test_number_formats():
    assert nk._num("12.345.678") == 12345678
    assert nk._num("12,345,678") == 12345678
    assert nk._num("1.234") == 1234
    assert nk._num("1234567.5") == 1234567.5
    assert nk.clean({"tri_gia_tinh_thue": "50.000.000", "thue_gtgt": "4.000.000"})["thue_suat"] == "8%"


def test_imports_go_into_23_24_and_purchase_list(tmp_path, monkeypatch):
    monkeypatch.setattr(vat, "STATE_FILE", tmp_path / "vat.json")
    old = vat.settings.output_dir
    object.__setattr__(vat.settings, "output_dir", str(tmp_path))
    try:
        _check(tmp_path)
    finally:
        object.__setattr__(vat.settings, "output_dir", old)


def _check(tmp_path):
    key = vat.period_key("M", 10, 2026)
    vat.save_imports("0313749848", key, [{"so_tk": "107123456789", "ngay": "05/10/2026", "ncc": "SHENZHEN XYZ",
                                         "tri_gia_tinh_thue": "113.000.000", "thue_suat": "10%", "thue_gtgt": 11300000}])
    res = vat.compute("0313749848", "M", 10, 2026, {"ct23a": 5, "ct24a": 5})
    assert res["ct"]["23a"] == 113000000 and res["ct"]["24a"] == 11300000
    assert res["ct"]["23"] == 113000000 and res["ct"]["24"] == 11300000 and res["ct"]["25"] == 11300000
    line = [b for b in res["bk"]["purchase"] if b.get("nk")][0]
    assert line["so"] == "107123456789" and line["ten"] == "SHENZHEN XYZ" and line["gt"] == 113000000
    vat.save_imports("0313749848", key, [])
    assert vat.imports("0313749848", key) == []
