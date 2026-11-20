from app import vat_multi as vm


def test_period_of_month_and_quarter():
    assert vm.period_of("2026-10", "M") == ("M", 10, 2026)
    assert vm.period_of("2026-10", "Q") == ("Q", 4, 2026)
    assert vm.period_of("2026-03", "Q") == ("Q", 1, 2026)


def test_start_needs_login(monkeypatch):
    monkeypatch.setattr(vm.BROWSER, "slots_status", lambda ns: {n: {"open": True, "logged_in": False} for n in ns})
    r = vm.start(3, "2026-10")
    assert r == {"n": 3, "ok": False, "msg": "Chưa đăng nhập"}
    assert 3 not in vm.RUNS
