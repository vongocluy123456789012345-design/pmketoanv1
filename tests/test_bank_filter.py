from pathlib import Path

from openpyxl import Workbook, load_workbook

from app.bank_filter import BANK_SHEET, is_bank_purchase, is_bank_seller, strip_bank_rows
from app.sources.file_import import read_excel


def test_is_bank_seller():
    assert is_bank_seller("NGÂN HÀNG TMCP SÀI GÒN THƯƠNG TÍN - CHI NHÁNH QUẬN 1")
    assert is_bank_seller("Ngân hàng Nông nghiệp và Phát triển Nông thôn Việt Nam")
    assert is_bank_seller("VIETCOMBANK")
    assert is_bank_seller("Ngan hang TMCP A Chau")
    assert not is_bank_seller("CHI NHÁNH TỔNG CÔNG TY ĐIỆN LỰC")
    assert not is_bank_seller("CÔNG TY TNHH SẢN XUẤT THƯƠNG MẠI BIOVET")
    assert not is_bank_seller(None)


def test_is_bank_purchase_only_for_purchase():
    assert is_bank_purchase({"huong": "purchase", "nbten": "NGÂN HÀNG TMCP Á CHÂU"})
    assert not is_bank_purchase({"huong": "sold", "nbten": "NGÂN HÀNG TMCP Á CHÂU"})


def _portal_like(path: Path) -> None:
    """Bảng kê giống file cổng thuế: tiêu đề ở dòng 3-4, dòng cột ở dòng 6, dữ liệu từ dòng 7."""
    wb = Workbook()
    ws = wb.active
    ws["A3"] = "DANH SÁCH HÓA ĐƠN"
    ws["A4"] = "Từ ngày 01/07/2026 đến ngày 31/07/2026"
    ws.append([])
    ws.append(["STT", "Ký hiệu mẫu số", "Ký hiệu hóa đơn", "Số hóa đơn", "Ngày lập", "MST người bán/MST người xuất hàng",
               "Tên người bán/Tên người xuất hàng", "Tổng tiền chưa thuế", "Tổng tiền thuế"])
    ws.append([1, 1, "K26TTU", 2853619, "02/07/2026", "0300951119-011", "CHI NHÁNH TỔNG CÔNG TY ĐIỆN LỰC", 4271473, 341718])
    ws.append([2, 1, "K26TAB", 153033, "28/07/2026", "0301103908-214", "NGÂN HÀNG TMCP SÀI GÒN THƯƠNG TÍN", 110000, 11000])
    ws.append([3, 1, "C26TAA", 77, "29/07/2026", "0312172156", "CÔNG TY TNHH BIOVET", 500000, 40000])
    wb.save(path)


def test_strip_bank_rows(tmp_path):
    f = tmp_path / "BangKe_Mua-vao_2026-07.xlsx"
    _portal_like(f)
    assert strip_bank_rows(f) == 1
    wb = load_workbook(f)
    main = wb.worksheets[0]
    names = [main.cell(r, 7).value for r in range(7, main.max_row + 1)]
    assert names == ["CHI NHÁNH TỔNG CÔNG TY ĐIỆN LỰC", "CÔNG TY TNHH BIOVET"]
    assert [main.cell(r, 1).value for r in range(7, main.max_row + 1)] == [1, 2]  # STT đánh lại
    bank = wb[BANK_SHEET]
    assert bank.cell(4, 7).value == "NGÂN HÀNG TMCP SÀI GÒN THƯƠNG TÍN"
    # gọi lại không đổi gì
    assert strip_bank_rows(f) == 0
    # phần mềm vẫn đọc đủ 3 hóa đơn (để đối chiếu số lượng với cổng, không phải tải lại bảng kê)
    rows = read_excel(f.read_bytes(), f.name)
    assert sorted(r["shdon"] for r in rows) == ["153033", "2853619", "77"]


def test_vat_forecast_excludes_bank_invoices():
    import json
    import tempfile

    from app import vat
    from app.config import settings

    mst = "0312345678"

    def buy(so, seller, val):
        return {"huong": "purchase", "loai": "Mua vào", "thang": "10/2026", "ngay_lap": "2026-10-15", "khmshdon": "1",
                "khhdon": "C26TAA", "shdon": str(so), "nbmst": f"03000000{so:02d}", "nbten": seller, "nmmst": mst,
                "nmten": "Cty A", "tgtcthue": val, "tgtthue": round(val * 0.1), "tgtttbso": val + round(val * 0.1),
                "tthai": "Hóa đơn mới", "tthai_code": 1, "file_xml": "", "file_pdf": ""}

    rows = [buy(1, "CÔNG TY TNHH VẬT TƯ X", 1_000_000), buy(2, "NGÂN HÀNG TMCP Á CHÂU - CN SÀI GÒN", 100_000)]
    old = settings.output_dir
    with tempfile.TemporaryDirectory() as tmp:
        object.__setattr__(settings, "output_dir", tmp)
        try:
            d = Path(tmp) / mst / "2026-10"
            d.mkdir(parents=True)
            (d / ".dulieu_2026-10.json").write_text(json.dumps({"saved": "2026-11-05", "rows": rows}, ensure_ascii=False),
                                                    encoding="utf-8")
            r = vat.compute(mst, "M", 10, 2026, {})
        finally:
            object.__setattr__(settings, "output_dir", old)
    assert r["ct"]["23"] == 1_000_000 and r["ct"]["24"] == 100_000, r["ct"]
    assert [b["so"] for b in r["bk"]["purchase"]] == ["1"]
    assert r["stat"]["mua_vao"] == 1 and r["stat"]["loai_tru"] == 1
