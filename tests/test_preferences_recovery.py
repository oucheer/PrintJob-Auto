"""驱动属性页窗口的恢复能力：printui 静默退出要立刻重试，而不是干等到超时。"""
from __future__ import annotations

from types import SimpleNamespace

import pytest

import printcase.driver_ui as du


class FakeButton:
    def __init__(self, text: str, on_click=None) -> None:
        self._text = text
        self._on_click = on_click

    def window_text(self) -> str:
        return self._text

    def click_input(self) -> None:
        if self._on_click is not None:
            self._on_click()


class FakeWindow:
    def __init__(self, title: str, pid: int = 4242, buttons=(), handle: int = 0) -> None:
        self.title = title
        self.closed = False
        self._buttons = list(buttons)
        self.handle = handle
        self.element_info = SimpleNamespace(process_id=pid)

    def window_text(self) -> str:
        return "" if self.closed else self.title

    def is_visible(self) -> bool:
        return not self.closed

    def class_name(self) -> str:
        return "#32770"

    def close(self) -> None:
        self.closed = True

    def descendants(self, **_kwargs):
        return list(self._buttons)


class FakeDesktop:
    def __init__(self, windows=()) -> None:
        self._windows = list(windows)

    def windows(self):
        return list(self._windows)

    def add(self, window: FakeWindow) -> FakeWindow:
        self._windows.append(window)
        return window


class FakeProc:
    """模拟 printui 宿主：exited=True 表示“静默退出，一个窗口都没建”。"""

    def __init__(self, exited: bool = True, pid: int = 9911) -> None:
        self.exited = exited
        self.pid = pid
        self.terminated = False

    def poll(self):
        return 0 if (self.exited or self.terminated) else None

    def terminate(self) -> None:
        self.terminated = True

    def wait(self, timeout=None):  # noqa: ARG002
        return 0

    def kill(self) -> None:
        self.terminated = True


@pytest.fixture(autouse=True)
def fast_retry(monkeypatch):
    """把重试里的等待缩到毫秒级，测试不陪着 sleep。"""
    monkeypatch.setattr(du, "PRINTUI_EXIT_GRACE", 0.01)
    monkeypatch.setattr(du, "PREFERENCE_RETRY_PAUSE", 0.01)
    yield
    du._live_launchers.clear()


def test_retries_when_printui_exits_without_window(monkeypatch):
    """复现故障：printui 进程秒退、不建窗口。旧实现会干等 25 秒，现在应立刻重试。"""
    desktop = FakeDesktop()
    launches = []

    def launcher(_printer):
        launches.append(1)
        return FakeProc(exited=True)

    with pytest.raises(du.DriverUiError) as excinfo:
        du.open_preferences("Fake Printer", timeout=3, desktop=desktop, launcher=launcher, attempts=3)

    assert len(launches) == 3
    assert "已尝试 3 次" in str(excinfo.value)
    assert "printui 进程已退出" in str(excinfo.value)


def test_recovers_on_second_attempt(monkeypatch):
    """第一次秒退、第二次正常弹出窗口：整条链路必须自愈，不能判用例失败。"""
    desktop = FakeDesktop()
    window = FakeWindow("Fake Printer 打印首选项")
    launches = []

    def launcher(_printer):
        launches.append(1)
        if len(launches) == 2:
            desktop.add(window)
        return FakeProc(exited=True)

    logs = []
    result = du.open_preferences(
        "Fake Printer", timeout=4, desktop=desktop, launcher=launcher, attempts=3, log=logs.append,
    )
    assert result is window
    assert len(launches) == 2
    assert any("重试" in line for line in logs)


def test_reuses_already_open_preferences_window():
    """上一轮留下的属性页仍可用时应直接复用，不再启动新的 printui。"""
    desktop = FakeDesktop([FakeWindow("Fake Printer 打印首选项")])
    launches = []

    def launcher(_printer):
        launches.append(1)
        return FakeProc()

    window = du.open_preferences("Fake Printer", timeout=3, desktop=desktop, launcher=launcher)
    assert window.window_text() == "Fake Printer 打印首选项"
    assert launches == []


def test_close_stale_preferences_closes_window_and_kills_host():
    """停止/失败后清理：窗口点“取消”关掉，仍挂着的 printui 宿主一并结束。"""
    window = FakeWindow("Fake Printer 打印首选项", buttons=[FakeButton("取消")])
    window.descendants = lambda **kwargs: [FakeButton("取消", on_click=lambda: setattr(window, "closed", True))]
    desktop = FakeDesktop([window])
    proc = FakeProc(exited=False)
    du._live_launchers.append(proc)

    cleared = du.close_stale_preferences("Fake Printer", desktop=desktop)

    assert window.closed
    assert proc.terminated
    assert cleared == 2


def test_close_preference_window_kills_hung_host(monkeypatch):
    """点不动、关不掉的属性页：结束它的 printui 宿主进程，避免挡住下一次打开。"""
    window = FakeWindow("Fake Printer 打印首选项", pid=7777)
    killed = []
    monkeypatch.setattr(du, "_process_exe", lambda pid: r"C:\Windows\System32\rundll32.exe")
    monkeypatch.setattr(du, "_terminate_pid", lambda pid: killed.append(pid) or True)
    monkeypatch.setattr(du, "_window_gone", lambda _w, _t=1.5: bool(killed))

    assert du._close_preference_window(window) is True
    assert killed == [7777]


class HandleDesktop(FakeDesktop):
    """按句柄取窗口的替身：模拟“原生枚举可用、UIA 全量遍历不可用”。"""

    def __init__(self, windows=()) -> None:
        super().__init__(windows)
        self.by_handle = {w.handle: w for w in windows}
        self.handle_requests: list[int] = []

    def window(self, handle=None, **_kwargs):
        self.handle_requests.append(handle)
        return self.by_handle[handle]

    def windows(self):
        raise AssertionError("不应回退到 UIA 全量遍历")


def test_preferences_window_found_by_native_enumeration(monkeypatch):
    """窗口发现优先走原生 EnumWindows，不再依赖 UIA 全量遍历。"""
    window = FakeWindow("Fake Printer 打印首选项", handle=131072)
    desktop = HandleDesktop([window])
    monkeypatch.setattr(du, "_native_preference_handles", lambda _p: [131072])

    assert du._preferences_windows(desktop, "Fake Printer") == [window]
    assert desktop.handle_requests == [131072]


def test_open_preferences_survives_uia_enumeration_failure(monkeypatch):
    """UIA 全量枚举整体失效时，靠原生枚举仍要能打开（用户现场的症状）。"""
    window = FakeWindow("Fake Printer 打印首选项", handle=131072)
    desktop = HandleDesktop([window])
    monkeypatch.setattr(du, "_native_preference_handles", lambda _p: [131072])
    # 假句柄上没有真实消息循环，这里只验证“窗口被发现并返回”
    monkeypatch.setattr(du, "_responsive", lambda *_a, **_k: True)

    result = du.open_preferences("Fake Printer", timeout=3, desktop=desktop, attempts=1)
    assert result is window
    assert desktop.handle_requests == [131072]


def test_uia_blind_window_reports_accurate_reason(monkeypatch):
    """窗口确实起来了但 UIA 读不到：原因要说清，并提示重启程序。"""

    class BlindDesktop(FakeDesktop):
        def windows(self):
            raise RuntimeError("UIA 全量枚举失败")

        def window(self, handle=None, **_kwargs):
            raise RuntimeError("UIA 单句柄查询也失败")

    monkeypatch.setattr(du, "_native_preference_handles", lambda _p: [131072])
    with pytest.raises(du.DriverUiError) as excinfo:
        du.open_preferences(
            "Fake Printer", timeout=2, attempts=1, desktop=BlindDesktop(),
            launcher=lambda _p: FakeProc(exited=False),
        )

    text = str(excinfo.value)
    assert "无法通过 UI 自动化读取" in text
    assert "重启程序" in text
