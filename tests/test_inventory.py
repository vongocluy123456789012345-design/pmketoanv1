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


def test_same_line_other_variant():
    """Cùng dòng / model khác phiên bản – màu – size vẫn ghép mã chuẩn; khác số model hay khác dòng thì KHÔNG."""
    mk_i = lambda m, t, d="Cái", sl=5: {"kho": "Hàng hóa", "ma": m, "ten": t, "dvt": d, "sl": sl, "gt": 1}  # noqa: E731
    items = [mk_i("V_LI_C300", "Vợt cầu lông Lining Calibar 300"),
             mk_i("V_LI_AE6000", "P-AYPR285-3- Khung vợt cầu lông Aeronaut 6000 Speed"),
             mk_i("ACL_Y_TRM2988", "Áo cầu lông Yonex TRM2988"),
             mk_i("OCL_S70", "Ống cầu lông Hải Yến S70", "Ống"),
             mk_i("BALO_Y_524", "Balo Yonex BAG524B0112ZJTBKZZZZZZ"),
             mk_i("V_Y_A11PRO", "Vợt cầu lông Yonex Arcsaber 11 Pro  - 4U5"),
             mk_i("DC_G_P17C", "Cước căng vợt Tennis Gosen Polylon 17 cuộn 200m", "Cuộn")]
    mk = lambda t: {"row": 1, "ma": "", "ten": t, "dvt": "", "sl": 1, "tt": 1, "ck": 0, "ngay": None}  # noqa: E731
    buy = [mk("P-AYPT395-4- Vợt cầu lông 3D Calibar 300B"), mk("Khung vợt cầu lông Aeronaut 6000 Max"),
           mk("Áo thun cổ tròn Yonex-2988-CC8- L-Xanh da trời"), mk("Cầu lông Hải Yến (S.70)"),
           mk("Túi/balo Yonex-All Day-Trắng thép BAG524B0612Z-STARWT- ZZZZ"),
           mk("Vợt cầu lông Yonex- ARCSABER11 PRO-4U5- Xám ngọc trai BFRARC11ZPZZ"),
           mk("Khung vợt cầu lông Halbertec 6000 - 4U"),                 # khác dòng (Halbertec ≠ Aeronaut)
           mk("Comfort 17 cuộn 220m - Trắng"),                            # khác dòng (Comfort ≠ Polylon)
           mk("Áo thun cổ tròn Yonex-2993-CC8- XS-Đen/vàng"), mk("Áo thun cổ tròn Yonex-2993-CC8- 2XL-Xanh dương chìm")]
    inv.match_codes(items, {"Mua vào": buy})
    got = [ln["ma_ghep"] for ln in buy]
    assert got[:6] == ["V_LI_C300", "V_LI_AE6000", "ACL_Y_TRM2988", "OCL_S70", "BALO_Y_524", "V_Y_A11PRO"]
    assert "MỚI" in buy[6]["cach"] and "MỚI" in buy[7]["cach"]
    assert got[8] == got[9] and "MỚI" in buy[8]["cach"]                 # cùng model áo mới, khác size / màu: 1 mã


def test_sales_spill_to_same_line_in_stock():
    """Bán "BG 65" vượt tồn mã đã ghép (BG 65 cuộn) -> xuất hàng cùng dòng còn tồn (BG 65 Titanium)."""
    items = [{"kho": "Hàng hóa", "ma": "DC_Y_65C", "ten": "Dây cước căng vợt Yonex BG 65 cuộn", "dvt": "Cuộn", "sl": 2, "gt": 1},
             {"kho": "Hàng hóa", "ma": "DC_Y_65T", "ten": "Dây cước căng vợt Yonex BG 65 TITANIUM", "dvt": "Cái", "sl": 133, "gt": 1},
             {"kho": "Hàng hóa", "ma": "DC_Y_80", "ten": "Dây cước căng vợt Yonex BG 80", "dvt": "Cái", "sl": 100, "gt": 1}]
    sell = [{"row": r, "ma": "", "ten": "Dây cước căng vợt Yonex BG 65 cuộn", "dvt": "Cuộn", "sl": 2, "tt": 1, "ck": 0,
             "ngay": None} for r in range(5)]
    inv.match_codes(items, {"Bán ra": sell})
    assert sell[0]["ma_ghep"] == "DC_Y_65C"
    assert all(ln["ma_ghep"] == "DC_Y_65T" and "KIỂM TRA" in ln["cach"] for ln in sell[1:])
    rows = inv.compute(items, [], sell, [])
    assert not any(r["sl_cuoi"] < 0 for r in rows)


def test_sales_spill_to_similar_name_other_code():
    """Thiếu "Hải Yến S90" -> xuất "Hải Yến S70" (chỉ khác mã phía sau); không lấy sang loại hàng khác."""
    items = [{"kho": "Hàng hóa", "ma": "OCL_S90", "ten": "Ống cầu lông Hải Yến S90", "dvt": "Ống", "sl": 1, "gt": 1},
             {"kho": "Hàng hóa", "ma": "OCL_S70", "ten": "Ống cầu lông Hải Yến S70", "dvt": "Ống", "sl": 10, "gt": 1},
             {"kho": "Hàng hóa", "ma": "BXG", "ten": "Bình xịt vệ sinh giày Ligpro", "dvt": "Ống", "sl": 50, "gt": 1},
             {"kho": "Hàng hóa", "ma": "V_Y_A88D", "ten": "Vợt cầu lông Yonex Astrox 88D Pro", "dvt": "Cái", "sl": 3, "gt": 1}]
    mk = lambda t, sl: {"row": 1, "ma": "", "ten": t, "dvt": "", "sl": sl, "tt": 1, "ck": 0, "ngay": None}  # noqa: E731
    sell = [mk("Ống cầu lông Hải Yến S90", 1), mk("Ống cầu lông Hải Yến S90", 3), mk("2024-4U5-Den bac BFR3AX88DZTZ-", 1)]
    inv.match_codes(items, {"Bán ra": sell})
    assert [ln["ma_ghep"] for ln in sell] == ["OCL_S90", "OCL_S70", "V_Y_A88D"]
    assert not any(r["sl_cuoi"] < 0 for r in inv.compute(items, [], sell, []))


def test_sales_never_create_codes():
    """Bán ra chỉ xuất hàng ĐANG CÓ: không khớp chắc chắn thì ghép vào hàng cùng loại gần nhất, không tạo mã mới."""
    items = [{"kho": "Hàng hóa", "ma": "V_Y_A77", "ten": "Vợt cầu lông Yonex Astrox 77 Pro", "dvt": "Cái", "sl": 2, "gt": 1},
             {"kho": "Hàng hóa", "ma": "ACL_Y_1", "ten": "Áo cầu lông Yonex TRM2861", "dvt": "Cái", "sl": 5, "gt": 1}]
    mk = lambda t, sl=1: {"row": 1, "ma": "", "ten": t, "dvt": "", "sl": sl, "tt": 1, "ck": 0, "ngay": None}  # noqa: E731
    buy = [mk("Xịt lạnh Ligpro Fast Cold 200ml", 10)]                # hàng mới hoàn toàn khi mua -> tạo mã mới
    sell = [mk("Vot cau long Yonex Nanoray 72 Light"), mk("Xit lanh Ligpro Fast Cold (200ml)")]
    rep = inv.match_codes(items, {"Mua vào": buy, "Bán ra": sell})
    assert "MỚI" in buy[0]["cach"]
    assert not any("MỚI" in x["cach"] for x in rep if x["loai"] == "Bán ra")
    assert sell[0]["ma_ghep"] == "V_Y_A77"                             # vợt ghép vào vợt (không vào áo)
    assert sell[1]["ma_ghep"] == buy[0]["ma_ghep"]                     # bán đúng hàng vừa mua
