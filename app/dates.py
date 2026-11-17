"""Ngày giờ hóa đơn: cổng thuế trả giờ UTC (vd "2026-10-06T17:00:00Z" = 07/10/2026 giờ Việt Nam)."""
from __future__ import annotations

from datetime import datetime, timedelta, timezone

VN_TZ = timezone(timedelta(hours=7))


def vn_date(value: object) -> str:
    """Trả "YYYY-MM-DD" theo giờ Việt Nam. Chuỗi không có múi giờ thì giữ nguyên phần ngày."""
    s = str(value or "").strip()
    if not s:
        return ""
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        return s[:10]
    if dt.tzinfo is None:
        return dt.date().isoformat()
    return dt.astimezone(VN_TZ).date().isoformat()
