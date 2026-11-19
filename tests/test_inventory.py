"""Tổng hợp tồn kho mẫu MISA: python -m pytest tests/test_inventory.py"""
import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import Workbook, load_workbook  # noqa: E402

from app import inventory as inv  # noqa: E402


def _tonkho_file() -> bytes:
    """Báo cáo Tổng hợp tồn kho kiểu MISA: tiêu đề 2 dòng, 1 mã nằm ở 2 kho, dòng Tổng cộng + chữ ký."""
    wb = Workbook()
    ws = wb.active
    ws["A1"], ws["A2"], ws["A3"] = "HỘ KINH DOANH A", "TỔNG HỢP TỒN KHO", "Tháng 10 năm 2026"
    for c, h in zip("ABCEGILN", ["Tên kho", "Mã hàng", "Tên hàng", "ĐVT", "Đầu kỳ", "Nhập kho", "Xuất kho", "Cuối kỳ"]):
        ws[f"{c}5"] = h
    for c, h in zip("GHIJLMNO", ["Số lượng", "Giá trị"] * 4):
        ws[f"{c}6"] = h
    data = [("Hàng hóa", "AO04", "Áo nịt bụng toàn thân dùng cho phụ nữ model A04", "Chiếc", 100, 2_000_000),
            ("Công cụ dụng cụ", "AO04", "Áo nịt bụng toàn thân dùng cho phụ nữ model A04", "Chiếc", -2, 0),
            ("Hàng hóa", "ENS", "Thực phẩm dinh dưỡng y học Ensure Powder 850g", "Hộp", 10, 5_000_000)]
    for r, (kho, ma, ten, dvt, sl, gt) in enumerate(data, 7):
        ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"], ws[f"E{r}"] = kho, ma, ten, dvt
        ws[f"G{r}"], ws[f"H{r}"], ws[f"N{r}"], ws[f"O{r}"] = sl, gt, sl, gt
    ws["A10"] = "Tổng cộng"
    ws["K12"] = "Ngày 31 tháng 10 năm 2026"
    ws["A13"], ws["F13"], ws["K13"] = "Người lập biểu", "Kế toán", "Người đại diện"
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


def test_opening_match_compute_report():
    op = inv.read_opening(_tonkho_file())
    assert op["loai"] == "tonkho" and len(op["items"]) == 3
    items = op["items"]
    buy = [{"row": 2, "ma": "", "ten": "Áo nịt bụng toàn thân, dùng cho phụ nữ, model A04", "dvt": "Chiếc", "sl": 100,
            "tt": 3_000_000, "ck": 0, "ngay": None},
           {"row": 3, "ma": "GLU", "ten": "Thực phẩm dinh dưỡng y học Glucerna Vanilla 850g", "dvt": "Hộp", "sl": 5,
            "tt": 2_500_000, "ck": 0, "ngay": None}]
    sell = [{"row": 2, "ma": "", "ten": "Áo nịt bụng toàn thân dùng cho phụ nữ model A04", "dvt": "Chiếc", "sl": 150,
             "tt": 0, "ck": 0, "ngay": None}]
    rep = inv.match_codes(items, {"Mua vào": buy, "Bán ra": sell})
    assert buy[0]["ma_ghep"] == "AO04" and sell[0]["ma_ghep"] == "AO04"
    assert buy[1]["ma_ghep"] == "GLU" and "MỚI" in buy[1]["cach"]          # Glucerna KHÔNG bị ghép nhầm vào Ensure
    assert any(x["ma"] == "GLU" for x in rep)
    rows = inv.compute(items, buy, sell, [])
    ao = next(r for r in rows if r["ma"] == "AO04" and r["kho"] == "Hàng hóa")
    # bình quân = (2.000.000 + 3.000.000) / (100 + 100) = 25.000 -> xuất 150 x 25.000
    assert ao["sl_nhap"] == 100 and ao["gt_nhap"] == 3_000_000 and ao["gt_xuat"] == 3_750_000
    assert ao["sl_cuoi"] == 50 and ao["gt_cuoi"] == 1_250_000
    ccdc = next(r for r in rows if r["ma"] == "AO04" and r["kho"] == "Công cụ dụng cụ")
    assert ccdc["sl_cuoi"] == -2                                             # kho khác giữ nguyên, không gộp
    out = load_workbook(io.BytesIO(inv.write_report(_tonkho_file(), rows, "Tháng 11 năm 2026", None))).active
    assert out["A3"].value == "Tháng 11 năm 2026"
    vals = [[out.cell(r, c).value for c in (1, 2)] for r in range(7, 12)]
    assert ["Hàng hóa", "AO04"] in vals and ["Công cụ dụng cụ", "AO04"] in vals and ["Hàng hóa", "GLU"] in vals
    tot = next(r for r in range(7, 20) if out.cell(r, 1).value == "Tổng cộng")
    assert out.cell(tot, 10).value == 5_500_000            # tổng giá trị nhập


def test_match_by_standard_keywords():
    """Ghép theo MÃ CHUẨN đầu kỳ bằng từ khóa (hãng, dòng, model), bỏ qua cách gọi loại hàng / màu / size."""
    items = [{"kho": "Hàng hóa", "ma": "DC_G_65", "ten": "Dây cước căng vợt GOSEN Ryzonic 65", "dvt": "Sợi", "sl": 1, "gt": 1},
             {"kho": "Hàng hóa", "ma": "DC_G_65C", "ten": "Dây cước căng vợt GOSEN Ryzonic 65 cuộn", "dvt": "Cuộn", "sl": 1, "gt": 1},
             {"kho": "Hàng hóa", "ma": "BXG_LIG", "ten": "Bình xịt vệ sinh giày Ligpro (420ml)", "dvt": "Chai", "sl": 1, "gt": 1},
             {"kho": "Hàng hóa", "ma": "V_LI_H4000", "ten": "Vợt cầu lông Lining Halbertec 4000", "dvt": "Cái", "sl": 1, "gt": 1}]
    mk = lambda t: {"row": 1, "ma": "", "ten": t, "dvt": "", "sl": 1, "tt": 1, "ck": 0, "ngay": None}  # noqa: E731
    lines = [mk("Dây căng vợt cầu lông Gosen Ryzonic 65- Đỏ- mã hàng: BSRY65RE"),
             mk("Dây căng vợt cầu lông Gosen Ryzonic 65 cuộn 200m- Trắng"),
             mk("Xịt vệ sinh mặt vợt pickleball Ligpro (chai 200ml)"),
             mk("Khung vợt cầu lông Halbertec 4000 - 4U Pink")]
    inv.match_codes(items, {"Mua vào": lines})
    assert [ln["ma_ghep"] for ln in lines[:2]] == ["DC_G_65", "DC_G_65C"]
    assert lines[2]["ma_ghep"] != "BXG_LIG"                          # xịt mặt vợt KHÁC bình xịt vệ sinh giày
    assert lines[3]["ma_ghep"] == "V_LI_H4000" and "KIỂM TRA" in lines[3]["cach"]   # thiếu tên hãng -> cần kiểm tra
