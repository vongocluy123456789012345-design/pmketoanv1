from app import kiemxuat as k


def test_text_message():
    items = k.items_from_text("Chị Trang ơi em muốn xuất hóa đơn\n1. 10 thùng sữa Ensure 850g\nBàn làm việc x 2 cái\n"
                              "Ghế xoay: 5\n3 cuộn dây cước Yonex BG 65 giá 120k")
    assert [(i["ten"], i["sl"], i["dvt"]) for i in items] == [
        ("sữa Ensure 850g", 10, "thùng"), ("Bàn làm việc", 2, "cái"), ("Ghế xoay", 5, ""), ("dây cước Yonex BG 65", 3, "cuộn")]
    assert items[3]["dg"] == 120000


def test_contract_table_two_row_header():
    rows = [["STT", "HẠNG MỤC", "MÔ TẢ CHI TIẾT", "ĐVT", "KÍCH THƯỚC(Mét)", None, None, None, None, "ĐƠN GIÁ", "THÀNH TIỀN"],
            [None, None, None, None, "Dài", "Rộng", "Cao", "Tổng", "SL", None, None],
            [1, "Bàn làm việc", "Melamine chống ẩm", "Cái", 1.4, 0.93, None, None, 1, 1200000, 1200000],
            [6, "Tủ quần áo 1", "Melamine chống ẩm 17", "m2", 2.15, None, 2.99, 6.428, 1, 2300000, 14784400],
            [18, "Trần thạch cao 5", "Khung và tấm Vĩnh Tường", "m2", 6.7, 6.58, None, "(44.086 - 0.57) = 43.516", None,
             180000, 7832880],
            [None, "Tổng Cộng", None, None, None, None, None, None, None, None, 126870200]]
    items = k._rows_to_items(rows)
    assert [(i["ten"], i["sl"]) for i in items] == [("Bàn làm việc", 1), ("Tủ quần áo 1", 6.428), ("Trần thạch cao 5", 43.516)]


def test_match_group_and_dvt():
    cands = [{"ten": "Vải giả da PU, đã được ngâm tấm tráng phủ polyurethane, khổ 135cm", "dvt": "Mét", "ton": 100.0},
             {"ten": "Vải giả da được tráng phủ nhựa polyurethane PU BONG NANO, khổ 1m4", "dvt": "Mét", "ton": 50.0},
             {"ten": "Gót PC", "dvt": "Đôi", "ton": 160.0}]
    kws = [k._kw(c["ten"]) for c in cands]
    g, _, kind = k._best("Vải giả da PU", "m", cands, kws)
    assert kind == "gan" and g["ton"] == 150
    c, _, kind = k._best("GÓT PC", "đôi", cands, kws)
    assert kind == "khop" and c["ten"] == "Gót PC"
    assert not k._dvt_ok("cái", "Đôi") and k._dvt_ok("Cái", "chiếc")
