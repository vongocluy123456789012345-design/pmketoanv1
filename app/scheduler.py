"""Chạy tự động theo lịch: vào ngày/giờ đã đặt mỗi tháng, tải hóa đơn của THÁNG TRƯỚC.

* Chỉ chạy được khi phần mềm đang mở và Chrome của phần mềm đã đăng nhập cổng thuế.
  Chưa đăng nhập -> mở Chrome tới trang cổng, hiện thông báo Windows, thử lại mỗi 15 phút (tối đa 24 giờ).
* Mặc định chỉ chọn các loại KHÔNG cần captcha (mua vào, bán ra). Nếu chọn cả loại máy tính tiền thì
  phần mềm sẽ kêu "ting-ting" và chờ bạn nhập captcha như khi chạy tay.
* Xong thì hiện thông báo Windows kèm kết quả.
"""
from __future__ import annotations

import json
import subprocess
import threading
import time
from calendar import monthrange
from datetime import date, datetime, timedelta
from typing import Any

from .config import ROOT

STATE = ROOT / "data" / "schedule.json"
RETRY_MIN = 15
GIVE_UP_HOURS = 24


def toast(title: str, body: str) -> None:
    """Thông báo Windows (không cần cài thêm gì)."""
    esc = lambda s: s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("'", "''")  # noqa: E731
    ps = f"""
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] > $null
$x = New-Object Windows.Data.Xml.Dom.XmlDocument
$x.LoadXml('<toast><visual><binding template="ToastGeneric"><text>{esc(title)}</text><text>{esc(body)}</text></binding></visual><audio src="ms-winsoundevent:Notification.IM"/></toast>')
$app = '{{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}}\\WindowsPowerShell\\v1.0\\powershell.exe'
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app).Show([Windows.UI.Notifications.ToastNotification]::new($x))
"""
    try:
        subprocess.Popen(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                         creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception:  # noqa: BLE001
        pass


def prev_month(today: date) -> tuple[date, date]:
    last = today.replace(day=1) - timedelta(days=1)
    return last.replace(day=1), date(last.year, last.month, monthrange(last.year, last.month)[1])


class Scheduler:
    def __init__(self) -> None:
        self.enabled = False
        self.day = 5
        self.hour = 8
        self.kinds = ["mua-vao", "mua-vao-khong-ma", "ban-ra"]
        self.last_run = ""        # YYYY-MM của tháng đã tải tự động gần nhất
        self.pending_since = ""   # đang chờ đăng nhập từ lúc nào
        self.next_try = 0.0
        self.message = ""
        self.job_id: str | None = None
        self.vat_error = ""
        self._force = False
        self._thread: threading.Thread | None = None
        self._load()

    def _load(self) -> None:
        try:
            s = json.loads(STATE.read_text(encoding="utf-8"))
            for k in ("enabled", "day", "hour", "kinds", "last_run"):
                if k in s:
                    setattr(self, k, s[k])
        except (FileNotFoundError, ValueError):
            pass

    def _save(self) -> None:
        STATE.parent.mkdir(parents=True, exist_ok=True)
        STATE.write_text(json.dumps({"enabled": self.enabled, "day": self.day, "hour": self.hour,
                                     "kinds": self.kinds, "last_run": self.last_run}, ensure_ascii=False),
                         encoding="utf-8")

    def status(self) -> dict[str, Any]:
        d1, d2 = prev_month(date.today())
        return {"enabled": self.enabled, "day": self.day, "hour": self.hour, "kinds": self.kinds,
                "last_run": self.last_run, "message": self.message, "job_id": self.job_id,
                "waiting_login": bool(self.pending_since), "next_period": f"{d1:%m/%Y}"}

    def configure(self, enabled: bool, day: int, hour: int, kinds: list[str]) -> None:
        self.enabled, self.day, self.hour, self.kinds = enabled, day, hour, kinds
        self.message = (f"Đã bật: ngày {day} hằng tháng lúc {hour}h tự tải hóa đơn tháng trước"
                        if enabled else "Đã tắt chạy tự động")
        self._save()
        self.ensure_running()

    def trigger(self) -> None:
        """Chạy ngay (bỏ qua ngày/giờ đã đặt)."""
        self._force = True
        self.next_try = 0
        self.ensure_running()

    def ensure_running(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._loop, daemon=True, name="scheduler")
        self._thread.start()

    # ------------------------------------------------------------------
    def _due(self, now: datetime) -> bool:
        d1, _ = prev_month(now.date())
        return (self.enabled and self.last_run != f"{d1:%Y-%m}"
                and (now.day > self.day or (now.day == self.day and now.hour >= self.hour)))

    def _loop(self) -> None:
        from . import vat
        while True:
            try:
                self._tick()
            except Exception as e:  # noqa: BLE001
                self.message = f"Lỗi bộ chạy tự động: {e}"
            try:  # dự báo thuế GTGT ngay từ 00:00 ngày đầu kỳ sau
                vat.auto_tick()
                self.vat_error = ""
            except Exception as e:  # noqa: BLE001
                self.vat_error = f"Lỗi dự báo VAT tự động: {e}"
            time.sleep(30)

    def _tick(self) -> None:
        from .auto_download import JOBS, running_job, start
        from .gdt.browser import BROWSER

        now = datetime.now()
        if self.job_id:  # đang theo dõi lượt tự động
            job = JOBS.get(self.job_id)
            if job and job.status != "running":
                d1, _ = prev_month(now.date())
                if job.status == "done":
                    self.last_run = f"{d1:%Y-%m}"
                    self._save()
                self.message = (f"Lượt tự động {d1:%m/%Y}: {job.status} – {job.done}/{job.total} hóa đơn, "
                                f"PDF gốc {job.pdf_ok}, PDF từ cổng {job.pdf_hddt}")
                toast("Tải hóa đơn tự động " + ("xong" if job.status == "done" else "đã dừng"), self.message)
                self.job_id = None
            return
        if not (self._force or self._due(now)) or time.time() < self.next_try:
            return
        if running_job():
            self.next_try = time.time() + 60
            return
        st = BROWSER.status()
        if not st.get("open"):
            try:
                st = BROWSER.open()
            except Exception as e:  # noqa: BLE001
                self.message = f"Không mở được Chrome: {e}"
        if not st.get("logged_in"):
            if not self.pending_since:
                self.pending_since = now.isoformat(timespec="minutes")
                toast("Cần đăng nhập cổng thuế",
                      "Đến lịch tải hóa đơn tự động. Hãy đăng nhập hoadondientu trên cửa sổ Chrome của phần mềm.")
            if now - datetime.fromisoformat(self.pending_since) > timedelta(hours=GIVE_UP_HOURS):
                self.message = "Bỏ lượt tự động: chờ đăng nhập quá 24 giờ"
                self.pending_since, self._force = "", False
                d1, _ = prev_month(now.date())
                self.last_run = f"{d1:%Y-%m}"
                self._save()
                return
            self.message = f"Đang chờ đăng nhập cổng thuế (từ {self.pending_since}), thử lại mỗi {RETRY_MIN} phút"
            self.next_try = time.time() + RETRY_MIN * 60
            return
        d1, d2 = prev_month(now.date())
        job = start(d1, d2, self.kinds or ["mua-vao", "ban-ra"], True, True, True)
        self.job_id, self.pending_since, self._force = job.id, "", False
        self.message = f"Đang tự động tải hóa đơn {d1:%m/%Y}"
        toast("Bắt đầu tải hóa đơn tự động", f"Kỳ {d1:%m/%Y} – MST {st.get('mst')}")


SCHEDULER = Scheduler()
