import threading
import time

from app import auto_download as ad
from app.gdt import browser as b


def test_portal_never_more_than_4_in_flight(monkeypatch):
    w = b.BrowserWorker()
    state = {"now": 0, "max": 0, "calls": 0}
    lock = threading.Lock()

    def fake_run(reqs, timeout):
        with lock:
            state["now"] += len(reqs)
            state["max"] = max(state["max"], state["now"])
            state["calls"] += 1
            first = state["calls"] == 1
        time.sleep(0.02)
        with lock:
            state["now"] -= len(reqs)
        # lần gọi đầu tiên: cổng báo quá tải 1 lệnh -> phải gửi lại đúng lệnh đó
        return [{"status": 429 if first and i == 0 else 200, "ctype": "", "b64": ""} for i in range(len(reqs))]

    monkeypatch.setattr(w, "_run_async", fake_run)
    monkeypatch.setattr(w, "_auth_token", lambda: "tok")
    monkeypatch.setattr(b.time, "sleep", lambda s: None)
    monkeypatch.setattr(b, "RATE_GAP", 0)
    w._gap = 0
    out = {}

    def job(n):
        out[n] = w.fetch_many([("/x", {"i": i}, "") for i in range(13)])

    ths = [threading.Thread(target=job, args=(n,)) for n in range(3)]
    [t.start() for t in ths]
    [t.join() for t in ths]
    assert state["max"] <= b.PORTAL_CONC
    assert all(len(v) == 13 and all(r.status == 200 for r in v) for v in out.values())


def test_transient_pdf_errors():
    assert ad._transient("ReadTimeout: timed out")
    assert ad._transient("Server disconnected without sending a response")
    assert not ad._transient("không tìm thấy PDF khớp hóa đơn")
    assert not ad._transient("XML không có link/mã tra cứu")
