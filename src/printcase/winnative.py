"""原生窗口/进程小工具。

为什么需要这一层：pywinauto 的 UIA 桌面全量枚举（``Desktop.windows()``）在
Word 刚启动、系统繁忙或长时间运行后会整体失效——要么返回不完整列表，要么
单次调用就卡十几秒。文档窗口其实早就显示出来了，程序却“看不见”，表现为
“打开着文档但报未找到文档窗口”，并且把整个用例超时白白耗光。

``EnumWindows`` 只读句柄与窗口标题，毫秒级返回，且能看到 UIA 漏掉的窗口。
这里把它包装成几个无副作用的函数，UI 自动化只负责“拿句柄去包装成控件树”。
"""
from __future__ import annotations


def _win32gui():
    import win32gui  # noqa: WPS433 - 仅 Windows 环境可用，按需导入

    return win32gui


def hwnd_of(obj) -> int:
    """取窗口句柄；pywinauto 包装对象/原生整数都支持，取不到返回 0。"""
    for attribute in ("handle",):
        try:
            return int(getattr(obj, attribute, 0) or 0)
        except Exception:  # noqa: BLE001
            continue
    try:
        return int(obj or 0)
    except Exception:  # noqa: BLE001
        return 0


def window_title(hwnd: int) -> str:
    try:
        return str(_win32gui().GetWindowText(int(hwnd)) or "")
    except Exception:  # noqa: BLE001
        return ""


def is_window(hwnd: int) -> bool:
    try:
        return bool(_win32gui().IsWindow(int(hwnd)))
    except Exception:  # noqa: BLE001
        return False


def is_visible(hwnd: int) -> bool:
    try:
        return bool(_win32gui().IsWindowVisible(int(hwnd)))
    except Exception:  # noqa: BLE001
        return False


def window_pid(hwnd: int) -> int | None:
    try:
        import win32process  # noqa: WPS433

        return int(win32process.GetWindowThreadProcessId(int(hwnd))[1]) or None
    except Exception:  # noqa: BLE001
        return None


def enum_top_windows() -> list[tuple[int, str]]:
    """所有顶层窗口的 (hwnd, 标题)，含隐藏窗口；失败时返回空列表。"""
    try:
        gui = _win32gui()
    except Exception:  # noqa: BLE001
        return []
    found: list[tuple[int, str]] = []

    def callback(hwnd, _param):
        try:
            found.append((int(hwnd), str(gui.GetWindowText(hwnd) or "")))
        except Exception:  # noqa: BLE001
            pass
        return True

    try:
        gui.EnumWindows(callback, None)
    except Exception:  # noqa: BLE001
        return found
    return found


def process_name(pid: int) -> str:
    """进程映像文件名（小写，含扩展名）；取不到返回空串。"""
    try:
        import win32api  # noqa: WPS433
        import win32con  # noqa: WPS433
        import win32process  # noqa: WPS433

        handle = win32api.OpenProcess(
            win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, int(pid),
        )
        try:
            return str(win32process.GetModuleFileNameEx(handle, 0) or "").split("\\")[-1].lower()
        finally:
            win32api.CloseHandle(handle)
    except Exception:  # noqa: BLE001
        return ""


def process_alive(pid: int) -> bool:
    """进程是否还活着（无需管理员权限）。"""
    try:
        import win32api  # noqa: WPS433
        import win32con  # noqa: WPS433

        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    except Exception:  # noqa: BLE001
        return False
    try:
        import win32process  # noqa: WPS433

        code = win32process.GetExitCodeProcess(handle)
        return code == 259                   # STILL_ACTIVE
    except Exception:  # noqa: BLE001
        return False
    finally:
        try:
            win32api.CloseHandle(handle)
        except Exception:  # noqa: BLE001
            pass


def terminate_pid(pid: int) -> bool:
    """结束指定进程（仅用于本程序自己启动的 Word/WPS 实例）。"""
    try:
        import win32api  # noqa: WPS433
        import win32con  # noqa: WPS433
        import win32process  # noqa: WPS433

        handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, int(pid))
        try:
            win32process.TerminateProcess(handle, 1)
        finally:
            win32api.CloseHandle(handle)
        return True
    except Exception:  # noqa: BLE001
        return False
