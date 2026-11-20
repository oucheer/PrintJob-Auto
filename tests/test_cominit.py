"""COM 单元模式管理测试：STA 探测/初始化幂等，自检在独立线程执行。"""
from __future__ import annotations

import sys
import threading
import types

import pytest

from printcase import cominit

KNOWN = {"STA", "MTA", "中性(NA)", "主 STA", "TE", "未初始化"}


def test_apartment_probe_returns_known_name():
    assert cominit.apartment_name() in KNOWN


def test_ensure_sta_is_idempotent():
    first = cominit.ensure_sta()
    assert cominit.ensure_sta() == first
    assert cominit.apartment_name() == first


def test_ensure_sta_does_not_leak_into_other_threads():
    """COM 单元是线程私有的：主线程初始化不会让新线程「顺带」拥有 COM。"""
    seen: dict[str, str] = {}

    def worker() -> None:
        seen["before"] = cominit.apartment_name()
        seen["after"] = cominit.ensure_sta()

    cominit.ensure_sta()
    thread = threading.Thread(target=worker)
    thread.start()
    thread.join()
    assert seen["before"] == "未初始化"
    assert seen["after"] in KNOWN - {"未初始化"}


def test_ui_ready_runs_in_probe_thread(monkeypatch):
    """自检必须在独立线程里跑：主线程被 Word COM 占成 STA 也不受影响。"""
    called: dict[str, str] = {}

    class FakeDesktop:
        def __init__(self, backend: str = "uia") -> None:
            assert backend == "uia"
            called["thread"] = threading.current_thread().name

    fake = types.ModuleType("pywinauto")
    fake.__version__ = "0.0.0-test"
    fake.Desktop = FakeDesktop  # type: ignore[attr-defined]
    monkeypatch.setitem(sys.modules, "pywinauto", fake)

    ok, detail = cominit.ui_ready(timeout=10)
    assert ok is True
    assert called["thread"] == "printcase-ui-probe"
    assert "就绪" in detail


def test_ui_ready_reports_conflict(monkeypatch):
    def boom(*_args, **_kwargs):
        raise cominit.ComApartmentError("线程已用另一种单元模式初始化")

    monkeypatch.setattr(cominit, "import_ui_engine", boom)
    ok, detail = cominit.ui_ready(timeout=10)
    assert ok is False
    assert "COM 单元模式冲突" in detail

    with pytest.raises(cominit.ComApartmentError):
        cominit.import_ui_engine()
