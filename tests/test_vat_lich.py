"""Lịch nhắc dự báo VAT: python -m pytest tests/test_vat_lich.py"""
import sys
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pytest  # noqa: E402

from app import vat_lich as vl  # noqa: E402


def test_parse_days_and_time():
    assert vl.parse_days("1, 15;25 15 40 0") == [1, 15, 25]
    assert vl.parse_days([20, 1]) == [1, 20]
    assert vl.parse_time("8") == "08:00" and vl.parse_time("8h30") == "08:30" and vl.parse_time("17:05") == "17:05"
    with pytest.raises(ValueError):
        vl.parse_time("25:00")


def test_due_once_per_day():
    d = {"bat": True, "ngay": [1, 15], "gio": "08:00"}
    assert vl.due(datetime(2026, 11, 15, 8, 0), d)
    assert not vl.due(datetime(2026, 11, 15, 7, 59), d)
    assert not vl.due(datetime(2026, 11, 16, 9, 0), d)
    assert not vl.due(datetime(2026, 11, 15, 9, 0), {**d, "last": "2026-11-15"})
    assert not vl.due(datetime(2026, 11, 15, 9, 0), {**d, "bat": False})


def test_task_xml():
    x = vl.task_xml({"bat": True, "ngay": [1, 15, 25], "gio": "08:30"})
    ns = {"t": "http://schemas.microsoft.com/windows/2004/02/mit/task"}
    root = ET.fromstring(x.split("\n", 1)[1])
    assert [e.text for e in root.findall(".//t:DaysOfMonth/t:Day", ns)] == ["1", "15", "25"]
    assert root.find(".//t:StartBoundary", ns).text.endswith("T08:30:00")
    assert root.find(".//t:StartWhenAvailable", ns).text == "true"         # máy tắt đúng giờ -> nhắc khi bật lại
    assert "--nhac-vat" in root.find(".//t:Arguments", ns).text
    assert "tải hóa đơn" in vl.reminder_text() and "dự báo" in vl.reminder_text()


def test_schedule_due_and_period():
    from datetime import date
    from types import SimpleNamespace

    from app import scheduler as sc
    s = sc.Scheduler.__new__(sc.Scheduler)
    s.enabled, s.day, s.hour, s.minute, s.ky, s.last_run = True, 20, 9, 35, "nay", ""
    assert not s._due(datetime(2026, 11, 20, 9, 34))
    assert s._due(datetime(2026, 11, 20, 9, 35))
    assert s._due(datetime(2026, 11, 22, 8, 0))                 # máy tắt đúng giờ -> chạy bù khi mở
    s.last_run = "2026-11-20"
    assert not s._due(datetime(2026, 11, 25, 8, 0))             # mỗi tháng 1 lần
    assert sc.period_for(date(2026, 11, 20), "nay") == (date(2026, 11, 1), date(2026, 11, 20))
    assert sc.period_for(date(2026, 11, 20), "truoc") == (date(2026, 10, 1), date(2026, 10, 31))
    job = SimpleNamespace(status="done", get_xml=True, get_pdf=True, total=10, done=10, xml_ok=10, pdf_ok=8,
                          pdf_hddt=2, pdf_missing=0, excel_ok=3, errors=[], out_dir="x", summary_file=None, mst="1")
    assert sc.judge(job)["ok"] and "ĐẦY ĐỦ" in sc.judge(job)["title"]
    job.pdf_missing, job.errors = 2, ["11/2026 HĐ máy tính tiền bán ra: bỏ qua – không nhập captcha"]
    r = sc.judge(job)
    assert not r["ok"] and any("PDF" in m for m in r["missing"]) and any("captcha" in m for m in r["missing"])
    job.pdf_missing, job.xml_ok = 0, 8
    job.errors = ["10/2026 Mua vào – CQT nhận không mã K26TAB-1: không tải được XML (cổng trả lỗi 500 – Không tồn tại hồ sơ gốc của hóa đơn.)",
                  "10/2026 Mua vào – CQT nhận không mã K26TTU-2: không tải được XML (cổng trả lỗi 500 – Không tồn tại hồ sơ gốc của hóa đơn.)"]
    r = sc.judge(job)
    assert r["ok"] and r["notes"] and not r["missing"]           # cổng không có XML gốc: lưu ý, không phải thiếu


def test_schedule_period_choices():
    from datetime import date

    from app import scheduler as sc
    t = date(2026, 11, 20)
    assert sc.period_for(t, "quy_truoc") == (date(2026, 7, 1), date(2026, 9, 30))
    assert sc.period_for(t, "quy_nay") == (date(2026, 10, 1), date(2026, 11, 20))
    assert sc.period_for(t, "quy", "2026-07-01", "2026-09-30") == (date(2026, 7, 1), date(2026, 9, 30))
    assert sc.period_for(t, "thang", "2026-02-01", "2026-02-28") == (date(2026, 2, 1), date(2026, 2, 28))
    assert sc.period_for(t, "khoang", "2026-11-05", "2026-12-31") == (date(2026, 11, 5), date(2026, 11, 20))
    assert sc.describe_ky("quy", "2026-07-01", "2026-09-30") == "quý 3/2026"
    assert sc.describe_ky("khoang", "2026-01-05", "2026-02-10") == "từ 05/01/2026 đến 10/02/2026"
