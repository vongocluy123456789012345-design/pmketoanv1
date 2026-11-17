"""Test đọc nhanh hóa đơn trên cổng (giả lập trình duyệt): python -m tests.test_quick_read"""
import json
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import auto_download as ad  # noqa: E402
from app import quick_read, vat  # noqa: E402
from app.config import settings  # noqa: E402
from app.gdt import browser as br  # noqa: E402


def _inv(direction, so, rates, items):
    return {"nbmst": "0100000001" if direction == "purchase" else "0312345678", "nbten": "Bán A",
            "nmmst": "0312345678" if direction == "purchase" else "0100000002", "nmten": "Mua B",
            "khmshdon": 1, "khhdon": "C26TAA", "shdon": so, "tdlap": "2026-09-15T00:00:00", "tthai": 1,
            "tgtcthue": sum(r[1] for r in rates), "tgtthue": sum(r[2] for r in rates),
            "tgtttbso": sum(r[1] + r[2] for r in rates),
            "thttltsuat": [{"tsuat": r[0], "thtien": r[1], "tthue": r[2]} for r in rates], "_items": items}


def test_read_period():
    sold = [_inv("sold", 11, [("8%", 100_000_000, 8_000_000)], [("Xe tải", 100_000_000, "8%", 0.08)])]
    bought = [_inv("purchase", 22, [("10%", 30_000_000, 3_000_000), ("KCT", 1_000_000, 0)],
                   [("Sắt thép", 30_000_000, "10%", 0.1), ("Phí ngân hàng", 1_000_000, "KCT", 0)])]

    def fake_list(self, kind, d1, d2, *a, **k):
        if kind.get("captcha"):  # máy tính tiền: kỳ này không có hóa đơn
            return []
        if kind["direction"] == "sold":
            return [dict(x) for x in sold]
        return [dict(x) for x in bought] if kind["filter"] == "ttxly==5" else []

    class R:
        def __init__(self, d):
            self.ok, self.d = True, d

        def json(self):
            return self.d

    def fake_many(reqs, timeout=120):
        out = []
        for path, q, _ in reqs:
            inv = next(x for x in sold + bought if str(x["shdon"]) == q["shdon"])
            out.append(R({"shdon": inv["shdon"], "nbmst": inv["nbmst"], "nmmst": inv["nmmst"], "thttltsuat": inv["thttltsuat"],
                          "hdhhdvu": [{"ten": n, "thtien": t, "ltsuat": ls, "tsuat": ts, "dvtinh": "cái", "sluong": 1,
                                       "dgia": t, "tthue": round(t * ts)} for n, t, ls, ts in inv["_items"]]}))
        return out

    old_out = settings.output_dir
    saved = (ad.Runner.list_invoices, br.BROWSER.status, br.BROWSER.fetch_many, ad.running_job, vat.export, vat.notify_result)
    with tempfile.TemporaryDirectory() as tmp:
        object.__setattr__(settings, "output_dir", tmp)
        ad.Runner.list_invoices = fake_list
        br.BROWSER.status = lambda: {"logged_in": True, "mst": "0312345678"}
        br.BROWSER.fetch_many = fake_many
        ad.running_job = lambda: None
        vat.export = lambda res: {"xlsx": "x.xlsx"}
        vat.notify_result = lambda s: None
        try:
            st = quick_read.read_period("0312345678", "M", 9, 2026)
            assert st["mua_vao"] == 1 and st["ban_ra"] == 1 and st["can_lay_chi_tiet"] == 2, st
            r = vat.compute("0312345678", "M", 9, 2026, {})  # dự báo ngay từ danh sách (chưa có chi tiết)
            assert r["ct"]["32"] == 100_000_000 and r["ct"]["33"] == 8_000_000 and r["ct"]["24"] == 3_000_000
            for _ in range(50):  # bước 2 chạy nền: lấy chi tiết + lập bảng kê
                if quick_read.STATE["0312345678"]["status"] != "running":
                    break
                time.sleep(0.1)
            time.sleep(0.5)
            assert quick_read.STATE["0312345678"]["status"] == "done"
            r = vat.compute("0312345678", "M", 9, 2026, {})
            names = sorted(i["ten"] for d in ("sold", "purchase") for i in r["items"][d])
            assert names == ["Phí ngân hàng", "Sắt thép", "Xe tải"], names
            assert r["pl8"]["rows"][0]["ten"] == "Xe tải" and r["ct"]["40a"] == 5_000_000
            assert vat.saved_result("0312345678", "2026-09")["source"] == "doc-nhanh"
            js = list(Path(tmp).rglob("XML/*.json"))
            assert len(js) == 2 and json.loads(js[0].read_text(encoding="utf-8"))["hdhhdvu"]
        finally:
            (ad.Runner.list_invoices, br.BROWSER.status, br.BROWSER.fetch_many, ad.running_job, vat.export,
             vat.notify_result) = saved
            object.__setattr__(settings, "output_dir", old_out)


if __name__ == "__main__":
    test_read_period()
    print("OK test_read_period")
