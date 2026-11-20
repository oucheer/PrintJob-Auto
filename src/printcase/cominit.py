"""COM 单元（apartment）一致性管理。

背景：同一线程里两套 COM 初始化方式会互相打架——

- pywin32（Word/WPS 自动化，``pythoncom``）默认把线程初始化成 **STA**（单元线程）；
- pywinauto / comtypes 默认想用 **MTA**（多线程单元），且 comtypes 在 ``import`` 阶段
  就调用 ``CoInitializeEx(MTA)``。

线程一旦已经是 STA，再切 MTA 会被 Windows 拒绝，返回
``RPC_E_CHANGED_MODE``（-2147417850），中文提示正是
“无法在设置线程模式后对其加以更改”。comtypes 的调用没有兜底，于是
``import pywinauto`` 直接抛 OSError，表现为“UI 自动化不可用”。

本模块的做法：
1. 先探测当前线程真实公寓模式（``CoGetApartmentType``）；
2. 未初始化时显式初始化为 STA（Word COM 需要 STA，pywinauto 在 STA 下也能工作）；
3. 在导入 pywinauto 之前把 ``sys.coinit_flags`` 设成**匹配**的值，
   让 comtypes 不再“猜”；
4. 万一仍然冲突，清理半成品模块后按实际模式重试一次，并给出可读的原因。

自检（``ui_ready``）在一个临时线程里跑，绝不碰 Tk 主线程的 COM 状态。
"""
from __future__ import annotations

import ctypes
import sys
import threading
import warnings
from ctypes import wintypes

COINIT_MULTITHREADED = 0x0
COINIT_APARTMENTTHREADED = 0x2

RPC_E_CHANGED_MODE = -2147417850  # 0x80010106

APTTYPE_NAMES = {0: "STA", 1: "MTA", 2: "中性(NA)", 3: "主 STA", 4: "TE"}
_STA_LIKE = {"STA", "主 STA"}

_ole32 = ctypes.WinDLL("ole32")
_ole32.CoInitializeEx.argtypes = [wintypes.LPVOID, wintypes.DWORD]
_ole32.CoInitializeEx.restype = ctypes.c_long

_combase = ctypes.WinDLL("combase")
_combase.CoGetApartmentType.argtypes = [ctypes.POINTER(wintypes.DWORD), ctypes.POINTER(wintypes.DWORD)]
_combase.CoGetApartmentType.restype = ctypes.c_long

_lock = threading.Lock()


class ComApartmentError(RuntimeError):
    """COM 公寓模式无法满足 UI 自动化要求。"""


def apartment() -> str | None:
    """当前线程的 COM 公寓模式；未真正初始化返回 None。

    注意：从未调用过 CoInitializeEx 的线程，``CoGetApartmentType`` 会返回
    MTA + 限定符 ``IMPLICIT_MTA``（Windows 8 起的“隐式 MTA”），此时线程
    其实还没绑定单元，仍然可以初始化为 STA，所以按“未初始化”处理。
    """
    apt = wintypes.DWORD(0)
    qual = wintypes.DWORD(0)
    hr = _combase.CoGetApartmentType(ctypes.byref(apt), ctypes.byref(qual))
    if hr < 0:
        return None
    if qual.value == 1:  # IMPLICIT_MTA：线程未显式初始化
        return None
    return APTTYPE_NAMES.get(apt.value, f"未知({apt.value})")


def apartment_name() -> str:
    return apartment() or "未初始化"


def ensure_sta() -> str:
    """确保当前线程有 COM 初始化，并优先是 STA；返回实际模式名。

    已初始化成 MTA 的线程无法再改成 STA，此时原样返回 "MTA"，
    由调用方决定是否接受（pywinauto 在 MTA 下可用，Word COM 也能跑）。
    """
    state = apartment()
    if state is not None:
        return state
    with _lock:
        hr = _ole32.CoInitializeEx(None, COINIT_APARTMENTTHREADED)
    if hr < 0 and (hr & 0xFFFFFFFF) != 0x00000001:  # S_FALSE=1 表示已初始化
        raise ComApartmentError(
            f"本线程 COM 初始化失败（HRESULT 0x{hr & 0xFFFFFFFF:08X}）：{_hint(hr)}"
        )
    return apartment() or "STA"


def _hint(hr: int) -> str:
    code = hr & 0xFFFFFFFF
    if code == 0x80010106:
        return "线程已用另一种单元模式初始化，无法再改（RPC_E_CHANGED_MODE）"
    if code == 0x800401F0:
        return "COM 尚未初始化（CO_E_NOTINITIALIZED）"
    return "未知 COM 错误"


def _purge(*prefixes: str) -> None:
    for name in [n for n in sys.modules if n in prefixes or any(n.startswith(p + ".") for p in prefixes)]:
        sys.modules.pop(name, None)


def import_ui_engine():
    """导入并返回 pywinauto，保证与本线程公寓模式一致。

    幂等：已导入则直接返回。
    """
    cached = sys.modules.get("pywinauto")
    if cached is not None:
        return cached

    state = ensure_sta()
    flags = COINIT_APARTMENTTHREADED if state in _STA_LIKE else COINIT_MULTITHREADED
    sys.coinit_flags = flags  # comtypes 导入时按此调用 CoInitializeEx

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=".*coinit_flags.*")
        warnings.filterwarnings("ignore", message=".*Revert to STA COM threading mode.*")
        try:
            import pywinauto

            return pywinauto
        except Exception as exc:  # noqa: BLE001
            if not _is_changed_mode(exc):
                raise ComApartmentError(f"pywinauto 导入失败：{type(exc).__name__}: {exc}") from exc

        # 冲突兜底：清掉 comtypes 半成品，按当前真实模式再试一次
        _purge("comtypes", "pywinauto")
        actual = apartment()
        sys.coinit_flags = COINIT_APARTMENTTHREADED if actual in _STA_LIKE else COINIT_MULTITHREADED
        try:
            import pywinauto

            return pywinauto
        except Exception as exc:  # noqa: BLE001
            raise ComApartmentError(
                "pywinauto 与 Word/WPS 的 COM 单元模式冲突，无法在同一线程初始化"
                f"（当前：{apartment_name()}）。原始错误：{exc}"
            ) from exc


def _is_changed_mode(exc: BaseException) -> bool:
    if getattr(exc, "winerror", None) == RPC_E_CHANGED_MODE:
        return True
    return any(isinstance(arg, int) and arg == RPC_E_CHANGED_MODE for arg in getattr(exc, "args", ()))


def ui_ready(timeout: float = 30.0) -> tuple[bool, str]:
    """在独立线程里做 UI 引擎自检：导入 pywinauto 并真正创建一次 UIA 入口。

    放在临时线程里跑，可以彻底避开 Tk 主线程的 COM 状态（那里常年被
    Word/WPS 的 STA 初始化占用），也顺带验证了新线程能否正常工作。
    """
    box: dict[str, tuple[bool, str]] = {}

    def probe() -> None:
        try:
            state = ensure_sta()
            pywinauto = import_ui_engine()
            from pywinauto import Desktop  # noqa: WPS433

            Desktop(backend="uia")  # 触发 UIAutomation 接口创建
            box["result"] = (True, f"pywinauto {pywinauto.__version__} 就绪 · {state} 线程")
        except ComApartmentError as exc:
            box["result"] = (False, f"COM 单元模式冲突：{exc}")
        except Exception as exc:  # noqa: BLE001
            box["result"] = (False, f"UI 引擎自检失败：{type(exc).__name__}: {exc}")

    thread = threading.Thread(target=probe, name="printcase-ui-probe", daemon=True)
    thread.start()
    thread.join(timeout)
    if thread.is_alive():
        return False, f"UI 引擎自检 {timeout:.0f} 秒无响应（UIAutomation 可能被占用）"
    return box.get("result", (False, "UI 引擎自检未返回结果"))
