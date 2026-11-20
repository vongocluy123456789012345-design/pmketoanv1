from datetime import date

from app import tienich as ti


def test_mst_checksum():
    assert ti.mst_valid("0313749848")
    assert ti.mst_valid("0101243150")
    assert not ti.mst_valid("0313749849")
    assert ti.mst_valid("0101243150-001")


def test_workday_shift():
    assert ti._workday(date(2026, 12, 19)) == date(2026, 12, 21)  # thứ 7 -> thứ 2
    assert ti._workday(date(2026, 11, 20)) == date(2026, 11, 20)


def test_suggest_tk_rules_and_learning(tmp_path, monkeypatch):
    monkeypatch.setattr(ti, "TK_FILE", tmp_path / "tk.json")
    assert ti.suggest_tk("Cước vận chuyển hàng hóa")[0] == "6427"
    assert ti.suggest_tk("Cước vận chuyển", che_do="TT133")[0] == "6422"
    assert ti.suggest_tk("Vải đỏ cuộn 50m")[0] == "1561"
    assert ti.suggest_tk("Máy in Canon", 5_000_000)[0] == "153"
    assert ti.suggest_tk("Máy in Canon", 5_000_000, tai_san=False)[0] == "1561"
    assert ti.suggest_tk("Đế giày cao su", mac_dinh="152")[0] == "152"
    ti.learn_tk([("Đế giày cao su 25", "152"), ("Đế giày cao su 27", "152")])
    tk, why = ti.suggest_tk("Đế giày cao su 30")
    assert tk == "152" and "Đã học" in why


def test_debt_letters():
    led = {"partners": [{"ten": "Công ty A", "mst": "0313749848", "du": 1_500_000}],
           "entries": [{"ngay": date(2026, 10, 1), "so_ct": "BH01", "dien_giai": "Bán hàng", "ten": "Công ty A",
                        "phat_sinh": 2_000_000, "da": 500_000, "loai": "KH"}]}
    data, msgs = ti.debt_letters({"ten": "Công ty B", "mst": "0101243150"}, led, "KH", date(2026, 10, 31))
    assert data[:2] == b"PK" and len(msgs) == 1
    assert "1.500.000" in msgs[0]["loi_nhan"] and "quý công ty còn nợ" in msgs[0]["loi_nhan"]
