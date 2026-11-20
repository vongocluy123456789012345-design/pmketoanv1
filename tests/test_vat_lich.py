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
