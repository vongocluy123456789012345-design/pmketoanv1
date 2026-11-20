"""Lịch nhắc / tự dự báo thuế GTGT theo NGÀY tự chọn trong tháng (vd ngày 1, 15, 25 lúc 08:00).

* Phần mềm đang mở: đúng ngày giờ đã đặt, tự dự báo kỳ đang chạy ("tính đến ngày") cho các MST bật tự dự báo và
  báo lên chuông thông báo.
* Phần mềm CHƯA mở: lịch được đăng ký vào Windows Task Scheduler (tác vụ "AWBook\\NhacDuBaoVAT"); đến giờ Windows chạy
  "AWBook.exe --nhac-vat" -> hiện bảng nhắc "Hãy vào phần mềm để tải hóa đơn và dự báo VAT" (bấm Có để mở phần mềm).
  Máy tắt đúng giờ đó thì bảng nhắc hiện ngay lần bật máy / đăng nhập sau.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Any

from .config import ROOT, settings

LICH_FILE = ROOT / "data" / "vat_lich.json"
TASK_NAME = r"AWBook\NhacDuBaoVAT"
DEFAULT = {"bat": False, "ngay": [1, 20], "gio": "08:00"}   # người dùng bấm Lưu lịch nhắc mới bật


# ------------------------------------------------------------------ cài đặt
def parse_days(v: Any) -> list[int]:
    """"1, 15; 25" / [1, 15, 25] -> [1, 15, 25] (bỏ trùng, chỉ nhận 1..31)."""
    items = v if isinstance(v, list) else re.split(r"[^0-9]+", str(v or ""))
    out = sorted({int(x) for x in items if str(x).strip().isdigit() and 1 <= int(x) <= 31})
    return out


def parse_time(v: Any) -> str:
    m = re.fullmatch(r"\s*(\d{1,2})(?:[:hg.](\d{1,2}))?\s*", str(v or ""))
    if not m or int(m.group(1)) > 23 or int(m.group(2) or 0) > 59:
        raise ValueError("Giờ nhắc không hợp lệ (vd 08:00)")
    return f"{int(m.group(1)):02d}:{int(m.group(2) or 0):02d}"


def get() -> dict[str, Any]:
    try:
        d = json.loads(LICH_FILE.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        d = {}
    return {**DEFAULT, **d}


def _write(d: dict[str, Any]) -> None:
    LICH_FILE.parent.mkdir(parents=True, exist_ok=True)
    LICH_FILE.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


def save(bat: bool, ngay: Any, gio: Any) -> dict[str, Any]:
    days = parse_days(ngay)
    if bat and not days:
        raise ValueError("Chọn ít nhất 1 ngày trong tháng (1–31), vd: 1, 15, 25")
    d = {**get(), "bat": bool(bat), "ngay": days, "gio": parse_time(gio)}
    d["windows"] = register_task(d)
    _write(d)
    return d


def describe(d: dict[str, Any] | None = None) -> str:
    d = d or get()
    if not d["bat"] or not d["ngay"]:
        return "Đang tắt lịch nhắc dự báo VAT"
    return f"Nhắc dự báo VAT ngày {', '.join(map(str, d['ngay']))} hằng tháng lúc {d['gio']}"


# ------------------------------------------------------------------ Windows Task Scheduler
def _command(arg: str = "--nhac-vat") -> tuple[str, str, str]:
    """(chương trình, tham số, thư mục) để Windows chạy phần mềm với tham số arg."""
    if getattr(sys, "frozen", False):
        exe = Path(sys.executable)
        return str(exe), arg, str(exe.parent)
    py = Path(sys.executable)
    pyw = py.with_name("pythonw.exe")
    launcher = Path(__file__).resolve().parent.parent / "desktop.py"
    return str(pyw if pyw.exists() else py), f'"{launcher}" {arg}', str(launcher.parent)


def task_xml(d: dict[str, Any], arg: str = "--nhac-vat",
             desc: str = "AWBook: nhắc tải hóa đơn và dự báo thuế GTGT") -> str:
    from xml.sax.saxutils import escape
    cmd, args, wd = _command(arg)
    days = "".join(f"<Day>{x}</Day>" for x in d["ngay"])
    months = "".join(f"<{m}/>" for m in ("January", "February", "March", "April", "May", "June", "July", "August",
                                          "September", "October", "November", "December"))
    return f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>{escape(desc)}</Description></RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>{date.today().year}-01-01T{d['gio']}:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByMonth><DaysOfMonth>{days}</DaysOfMonth><Months>{months}</Months></ScheduleByMonth>
    </CalendarTrigger>
  </Triggers>
  <Principals><Principal id="Author"><LogonType>InteractiveToken</LogonType><RunLevel>LeastPrivilege</RunLevel></Principal></Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Enabled>true</Enabled>
  </Settings>
  <Actions Context="Author">
    <Exec><Command>{escape(cmd)}</Command><Arguments>{escape(args)}</Arguments><WorkingDirectory>{escape(wd)}</WorkingDirectory></Exec>
  </Actions>
</Task>
"""


def _schtasks(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(["schtasks", *args], capture_output=True, text=True, encoding="mbcs", errors="replace",
                          creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), timeout=30)


def register_task(d: dict[str, Any], name: str = "", arg: str = "--nhac-vat",
                  desc: str = "AWBook: nhắc tải hóa đơn và dự báo thuế GTGT",
                  ok_text: str = "Đã đặt lịch trên Windows: đến giờ máy tự hiện bảng nhắc, kể cả khi chưa mở phần mềm") -> str:
    """Đăng ký / gỡ lịch trong Windows Task Scheduler (d: bat, ngay, gio). Trả về mô tả trạng thái."""
    name = name or TASK_NAME
    if sys.platform != "win32":
        return "Không phải Windows – chỉ nhắc khi phần mềm đang mở"
    try:
        if not d["bat"] or not d["ngay"]:
            _schtasks("/Delete", "/TN", name, "/F")
            return "Đã gỡ lịch nhắc khỏi Windows"
        with tempfile.TemporaryDirectory() as td:
            f = Path(td) / "awbook_nhac_vat.xml"
            f.write_text(task_xml(d, arg, desc), encoding="utf-16")
            r = _schtasks("/Create", "/TN", name, "/XML", str(f), "/F")
        if r.returncode != 0:
            return f"Chưa đặt được lịch trên Windows: {(r.stderr or r.stdout).strip()[:200]}"
        return ok_text
    except Exception as e:  # noqa: BLE001
        return f"Chưa đặt được lịch trên Windows: {e}"


def task_exists(name: str = "") -> bool:
    return sys.platform == "win32" and _schtasks("/Query", "/TN", name or TASK_NAME).returncode == 0


# ------------------------------------------------------------------ chạy khi phần mềm đang mở
def due(now: datetime, d: dict[str, Any]) -> bool:
    return bool(d["bat"] and now.day in d["ngay"] and now.strftime("%H:%M") >= d["gio"]
                and d.get("last") != now.date().isoformat())


def tick(now: datetime | None = None) -> None:
    """Gọi mỗi 30 giây từ bộ chạy tự động: đúng ngày giờ đã đặt -> dự báo kỳ đang chạy + báo lên chuông."""
    from . import vat
    now = now or datetime.now()
    d = get()
    if not due(now, d):
        return
    d["last"] = now.date().isoformat()
    _write(d)
    done, ten = [], []
    for mst in vat._msts():
        cfg = vat.get_cfg(mst)
        if not cfg["auto"]:
            continue
        kieu = cfg["kieu"]
        ky, nam = (now.month if kieu == "M" else (now.month - 1) // 3 + 1), now.year
        if not any((Path(settings.output_dir) / mst / f"{y:04d}-{m:02d}").is_dir() for y, m in vat.period_months(kieu, ky, nam)):
            ten.append(mst)
            continue
        try:
            vat.forecast(mst, kieu, ky, nam, source="lich")
            done.append(f"{mst} ({vat.period_label(kieu, ky, nam)})")
        except Exception as e:  # noqa: BLE001
            ten.append(f"{mst} (lỗi: {str(e)[:80]})")
    body = "Hãy vào mục Tải hóa đơn đăng nhập cổng thuế để tải hóa đơn mới nhất, rồi bấm Dự báo VAT / ⚡ Đọc nhanh từ cổng."
    if done:
        body = "Đã dự báo tạm (tính đến hôm nay) cho: " + "; ".join(done) + ". " + body
    if ten:
        body += " Chưa có hóa đơn kỳ này: " + ", ".join(ten) + "."
    vat.add_notification(f"Hôm nay {now:%d/%m/%Y}: đến lịch tải hóa đơn & dự báo VAT", body, "warn", kind="vat-info")


# ------------------------------------------------------------------ bảng nhắc (Windows chạy "AWBook.exe --nhac-vat")
def reminder_text(now: datetime | None = None) -> str:
    now = now or datetime.now()
    base = Path(settings.output_dir)
    msts = sorted(p.name for p in base.iterdir() if p.is_dir() and re.fullmatch(r"\d{10}(-?\d{3})?", p.name)) \
        if base.is_dir() else []
    lines = [f"Hôm nay {now:%d/%m/%Y} là ngày tải hóa đơn và dự báo thuế GTGT.", "",
             "Hãy vào phần mềm AWBook để:",
             "   1. Tải hóa đơn: mục Tải hóa đơn → đăng nhập cổng hóa đơn điện tử → Bắt đầu tải",
             "   2. Dự báo VAT: mục Dự báo VAT → Dự báo VAT (hoặc ⚡ Đọc nhanh từ cổng)"]
    if msts:
        lines += ["", "Mã số thuế đang theo dõi: " + ", ".join(msts[:12]) + (" …" if len(msts) > 12 else "")]
    lines += ["", "Mở phần mềm AWBook ngay bây giờ?"]
    return "\n".join(lines)


def show_reminder(app_running: bool) -> bool:
    """Hiện bảng nhắc nổi trên cùng màn hình. Trả về True nếu người dùng muốn mở phần mềm."""
    import ctypes
    MB_YESNO, MB_OK, MB_ICONINFO, MB_TOPMOST, MB_SETFG = 0x4, 0x0, 0x40, 0x40000, 0x10000
    title = "AWBook – Nhắc dự báo thuế GTGT"
    if app_running:
        txt = reminder_text().rsplit("\n\n", 1)[0] + "\n\nPhần mềm đang mở: đã tự dự báo tạm, xem chuông Thông báo."
        ctypes.windll.user32.MessageBoxW(None, txt, title, MB_OK | MB_ICONINFO | MB_TOPMOST | MB_SETFG)
        return True
    r = ctypes.windll.user32.MessageBoxW(None, reminder_text(), title, MB_YESNO | MB_ICONINFO | MB_TOPMOST | MB_SETFG)
    return r == 6   # IDYES
