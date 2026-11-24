"""驱动属性页（打印首选项）自动化——用例属性的唯一设置通道。

设计原则（0.4 起替代原“直接驱动”里的 DEVMODE 直写）：

- 属性只通过**驱动自己的界面**设置，等于用户手工点“打印首选项”；
  驱动内部怎么把界面选项翻译成私有 DEVMODE 编码，由驱动自己负责。
- 设置完成后**只读**回读打印机默认 DEVMODE，逐项核对是否真的生效；
  核对不一致的用例判定失败且不下发（不出现“属性静默失效但作业照常打印”）。
- 用例结束后用执行前的 DEVMODE 快照还原，避免污染本机默认设置。

之所以不再直写 DEVMODE：多数驱动（含 Pantum）把克重介质、专有纸盒等
编码成私有不透明 DWORD，DC_MEDIATYPES 上报的值甚至每次调用都不同，
外部写死数值无法保证指向正确介质；而驱动界面里的下拉序号是稳定的。
"""
from __future__ import annotations

import ctypes
import platform
import subprocess
import threading
import time
from dataclasses import dataclass
from typing import Any, Callable

from .models import PrintCase
from .ui_print import _best_match, _descendants, _matches_labels, _norm, _text_of
from .validator import normalize_paper

# 常用纸张的 DMPAPER_* 代号（驱动未上报 DC_PAPERS 时使用）
FALLBACK_PAPER_CODES = {
    "letter": 1, "legal": 5, "statement": 6, "executive": 7,
    "a3": 8, "a4": 9, "a5": 11, "folio": 14, "ledger": 15,
    "jisb4": 12, "jisb5": 13, "jisb6": 88, "a6": 70,
}

# DEVMODE 里的双面/彩色取值
DUPLEX_CODES = {"simplex": 1, "duplex": 2, "duplexlong": 2, "duplexshort": 3}
DUPLEX_LABELS = {
    "simplex": ("单面打印", "单面", "1-sided", "simplex"),
    "duplex": ("双面打印-长边", "双面长边", "长边", "duplex"),
    "duplexlong": ("双面打印-长边", "双面长边", "长边", "duplex"),
    "duplexshort": ("双面打印-短边", "双面短边", "短边", "short edge"),
}
COLOR_LABELS = {"color": ("彩色打印", "彩色", "color"), "monochrome": ("黑白打印", "黑白", "monochrome")}
COLOR_CODES = {"monochrome": 1, "color": 2}

TAB_BASIC = "基本"
TAB_PAPER = "纸张"
TAB_LAYOUT = "版面"

# 纸张页内下拉框的排布顺序（驱动把标签错位挂到下一个控件上，只能按位置取）
PAPER_TAB_COMBOS = ("纸张尺寸", "纸张类型", "纸张来源")


class DriverUiError(RuntimeError):
    """驱动属性页无法打开/控件无法定位/回读校验不通过。"""


@dataclass(frozen=True)
class Action:
    """一条“在驱动界面上要做的事”。"""

    page: str
    label: str
    kind: str                 # combo_index / combo_text / radio / toggle / int_text
    value: Any
    expect_field: str = ""    # 回读校验对应的 DEVMODE 字段
    expect_value: Any = None
    target_text: str = ""     # 目标选项的显示文本（判断驱动界面是否已是目标项）

    def describe(self) -> str:
        if self.kind == "combo_index":
            return f"{self.label}=第 {int(self.value) + 1} 项"
        if self.kind == "toggle":
            return f"{self.label}={'开' if self.value else '关'}"
        return f"{self.label}={self.value}"


def _index_in_names(wanted: str, names: list[str] | None) -> int | None:
    """在驱动上报的名称列表里定位下标（先精确、再双向包含）。"""
    if not names:
        return None
    target = _norm(wanted)
    if not target:
        return None
    for index, name in enumerate(names):
        if _norm(name) == target:
            return index
    for index, name in enumerate(names):
        text = _norm(name)
        if text and (target in text or text in target):
            return index
    return None


def plan_actions(case: PrintCase, caps: dict | None = None) -> list[Action]:
    """把用例属性翻译成驱动界面操作计划（纯函数，便于单元测试）。"""
    caps = caps or {}
    actions: list[Action] = []

    paper_token = "mixed" if case.mixed_sizes.lower() == "yes" else normalize_paper(case.paper)
    if paper_token and paper_token not in {"", "auto", "mixed"}:
        paper_index = _index_in_names(case.paper, caps.get("papers")) if caps.get("papers") else None
        paper_ids = caps.get("paper_ids") or []
        expected = paper_ids[paper_index] if paper_index is not None and paper_index < len(paper_ids) else FALLBACK_PAPER_CODES.get(paper_token)
        if paper_index is not None:
            matched_paper = (caps.get("papers") or [])[paper_index]
            actions.append(Action(TAB_PAPER, "纸张尺寸", "combo_index", paper_index, "PaperSize", expected, matched_paper))
        else:
            actions.append(Action(TAB_PAPER, "纸张尺寸", "combo_text", case.paper, "PaperSize", expected, case.paper))

    media = str(case.media_type or "").strip()
    if media:
        # 介质代号由驱动私有编码决定，没有可预期的标准值：只校验“确实变了且非零”
        media_index = _index_in_names(media, caps.get("media_names"))
        if media_index is not None:
            matched_media = (caps.get("media_names") or [])[media_index]
            actions.append(Action(TAB_PAPER, "纸张类型", "combo_index", media_index, "MediaType", None, matched_media))
        else:
            actions.append(Action(TAB_PAPER, "纸张类型", "combo_text", media, "MediaType", None, media))

    bin_value = str(case.bin or "").strip()
    if bin_value:
        bin_index = _index_in_names(bin_value, caps.get("bin_names")) if not bin_value.isdigit() else None
        bin_ids = caps.get("bin_ids") or []
        expected_bin = bin_ids[bin_index] if bin_index is not None and bin_index < len(bin_ids) else (int(bin_value) if bin_value.isdigit() else None)
        if bin_index is not None:
            matched_bin = (caps.get("bin_names") or [])[bin_index]
            actions.append(Action(TAB_PAPER, "纸张来源", "combo_index", bin_index, "DefaultSource", expected_bin, matched_bin))
        else:
            actions.append(Action(TAB_PAPER, "纸张来源", "combo_text", bin_value, "DefaultSource", expected_bin, bin_value))

    # 份数不在驱动属性页里设置：Word 打印对话框按“份数”提交，PDF/图片由内置渲染器
    # 按页重复展开；两处都设会让驱动与上层各乘一次，导致份数翻倍。

    if case.duplex and case.duplex != "simplex":
        actions.append(Action(TAB_BASIC, "双面", "radio", DUPLEX_LABELS.get(case.duplex, (case.duplex,)), "Duplex", DUPLEX_CODES.get(case.duplex)))
    elif case.duplex == "simplex":
        actions.append(Action(TAB_BASIC, "双面", "radio", DUPLEX_LABELS["simplex"], "Duplex", 1))

    # 黑白设备本身只有一种颜色，驱动属性页通常没有“颜色”选项；此时不下发颜色动作，
    # 否则严格模式会因“找不到颜色选项”拒绝下发，且回读 COLOR 也永远匹配不上彩色。
    # 驱动未提供能力（caps 为 None / color=None）时保持原行为，按名称到界面里匹配。
    device_is_monochrome = caps.get("color") is False
    if case.color and not device_is_monochrome:
        actions.append(Action(TAB_BASIC, "颜色", "radio", COLOR_LABELS.get(case.color, (case.color,)), "Color", COLOR_CODES.get(case.color)))

    quality = str(case.print_quality or "").lower()
    if quality and quality != "auto":
        label = f"{quality} dpi" if quality.isdigit() else quality
        expected = int(quality) if quality.isdigit() else None
        actions.append(Action(TAB_BASIC, "打印质量", "radio", (label, quality), "PrintQuality", expected))

    if case.collate:
        actions.append(Action(TAB_BASIC, "逐份", "toggle", case.collate == "yes", "Collate", 1 if case.collate == "yes" else 0))

    if case.skip_blank == "yes":
        actions.append(Action(TAB_BASIC, "跳过空白页", "toggle", True))

    if case.borderless == "yes":
        actions.append(Action(TAB_LAYOUT, "无边距", "toggle", True))

    return actions


# ---------------------------------------------------------------------------
# 界面操作
# ---------------------------------------------------------------------------

def _desktop():
    from .cominit import import_ui_engine, ensure_sta

    ensure_sta()
    import_ui_engine()
    from pywinauto import Desktop  # noqa: WPS433

    return Desktop(backend="uia")


# ---------------------------------------------------------------------------
# 驱动属性页窗口：启动、复用、清理
# ---------------------------------------------------------------------------

# printui 宿主（rundll32.exe printui.dll）在驱动/队列还没收尾时会“静默退出”，
# 一个窗口都不建；旧实现只会傻等到超时再把用例判失败。现在盯住进程，
# 秒退就立即清理重试。
PRINTUI_EXIT_GRACE = 1.2     # 进程退出后再给窗口 1.2 秒出现的机会
PREFERENCE_OPEN_ATTEMPTS = 3
PREFERENCE_RETRY_PAUSE = 0.8
PREFERENCE_MIN_BUDGET = 4.0  # 单次尝试最少等多久（慢驱动冷启动）

_PREFERENCE_WORDS = ("首选项", "属性", "Preferences", "Properties")
_PRINTUI_EXE_NAMES = {"rundll32.exe", "printui.exe"}

_launcher_lock = threading.Lock()
_live_launchers: list = []


def _default_launcher(printer: str):
    """启动 printui（等价于手工打开“打印首选项”），返回进程句柄。"""
    return subprocess.Popen(
        ["rundll32.exe", "printui.dll,PrintUIEntry", "/e", "/n", printer],
        close_fds=True,
    )


def _launch(printer: str, launcher=None):
    handle = (launcher or _default_launcher)(printer)
    if handle is not None:
        with _launcher_lock:
            _live_launchers[:] = [item for item in _live_launchers if _process_alive(item) is not False]
            _live_launchers.append(handle)
    return handle


def _process_alive(handle) -> bool | None:
    """进程是否还在跑；拿不到状态时返回 None（例如测试替身/句柄已失效）。"""
    if handle is None:
        return None
    try:
        return handle.poll() is None
    except Exception:  # noqa: BLE001
        return None


def _forget_launcher(handle) -> None:
    with _launcher_lock:
        try:
            _live_launchers.remove(handle)
        except ValueError:
            pass


def _terminate_launcher(handle, log=None) -> bool:
    """结束我们自己启动、仍挂着的 printui 宿主进程。"""
    _forget_launcher(handle)
    if _process_alive(handle) is not True:
        return False
    try:
        handle.terminate()
        try:
            handle.wait(timeout=1.5)
        except Exception:  # noqa: WPS420
            handle.kill()
    except Exception:  # noqa: BLE001
        return False
    if log:
        log("已结束残留的驱动属性页进程（printui）")
    return True


def _desktop_windows(desktop) -> list:
    """桌面顶层窗口；UIA 偶发异常时当作没有窗口，不让整条链路崩掉。"""
    try:
        return list(desktop.windows())
    except Exception:  # noqa: BLE001
        return []


def _preferences_windows(desktop, printer: str) -> list:
    """所有属于该打印机的属性页窗口（含不可见窗口，标题里带“首选项/属性”）。"""
    wanted = _norm(printer)
    matches = []
    for window in _desktop_windows(desktop):
        text = _text_of(window)
        if not text:
            continue
        if wanted and wanted not in _norm(text):
            continue
        if any(word in text for word in _PREFERENCE_WORDS):
            matches.append(window)
    return matches


def _window_pid(window) -> int | None:
    try:
        return int(window.element_info.process_id)
    except Exception:  # noqa: BLE001
        return None


def _responsive(window, timeout_ms: int = 400) -> bool:
    """窗口是否还在处理消息；卡死的驱动页点不动任何控件。"""
    try:
        hwnd = int(getattr(window, "handle", 0) or 0)
    except Exception:  # noqa: BLE001
        return True
    if not hwnd:
        return True
    try:
        import win32con  # noqa: WPS433
        import win32gui  # noqa: WPS433

        flags = getattr(win32con, "SMTO_ABORTIFHUNG", 0x0002) | getattr(win32con, "SMTO_BLOCK", 0x0001)
        win32gui.SendMessageTimeout(hwnd, win32con.WM_NULL, 0, 0, flags, timeout_ms)
        return True
    except Exception:  # noqa: BLE001
        return False


def _process_exe(pid: int) -> str:
    """进程映像路径（同用户进程无需管理员权限）。"""
    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        handle = kernel32.OpenProcess(0x1000, False, int(pid))  # QUERY_LIMITED_INFORMATION
        if not handle:
            return ""
        try:
            size = ctypes.c_ulong(1024)
            buffer = ctypes.create_unicode_buffer(1024)
            if kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
                return buffer.value
            return ""
        finally:
            kernel32.CloseHandle(handle)
    except Exception:  # noqa: BLE001
        return ""


def _terminate_pid(pid: int) -> bool:
    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        handle = kernel32.OpenProcess(0x0001, False, int(pid))  # PROCESS_TERMINATE
        if not handle:
            return False
        try:
            return bool(kernel32.TerminateProcess(handle, 1))
        finally:
            kernel32.CloseHandle(handle)
    except Exception:  # noqa: BLE001
        return False


def _window_gone(window, timeout: float = 1.5) -> bool:
    """等待窗口消失；句柄失效/读不到标题都按“已消失”处理。"""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            window.is_visible()
        except Exception:  # noqa: BLE001
            return True
        if not _text_of(window):
            return True
        time.sleep(0.1)
    return False


def _close_preference_window(window, log=None) -> bool:
    """关掉驱动属性页：先点“取消/关闭”，关不掉就结束它的 printui 宿主进程。

    停在一半的驱动属性页既会挡住下一次打开（printui 可能直接静默退出），
    也会让回读到的 DEVMODE 和界面对不上，所以必须确认“真的没了”。
    """
    clicked = False
    for button in _descendants(window, control_type="Button"):
        if _norm(_text_of(button)) in {_norm("取消"), _norm("关闭"), _norm("cancel"), _norm("close")}:
            try:
                button.click_input()
                clicked = True
            except Exception:  # noqa: WPS420
                pass
            break
    if not clicked:
        try:
            window.close()
        except Exception:  # noqa: WPS420
            pass
    if _window_gone(window, 2.0):
        return True

    pid = _window_pid(window)
    if pid and _process_exe(pid).lower().endswith(tuple(_PRINTUI_EXE_NAMES)):
        if _terminate_pid(pid):
            if log:
                log("驱动属性页未响应关闭请求，已结束其 printui 进程")
            return _window_gone(window, 1.5)
    if log:
        log(f"驱动属性页未能关闭：{_text_of(window) or '(无标题)'}")
    return False


def _usable_preferences(desktop, printer: str, log=None):
    """复用已经打开、且还能响应的属性页窗口；卡死的先清理掉。"""
    windows = _preferences_windows(desktop, printer)
    for window in windows:
        if _responsive(window):
            return window
    for window in windows:
        if log:
            log(f"发现无响应的驱动属性页窗口，正在清理：{_text_of(window) or printer}")
        _close_preference_window(window, log=log)
    return None


def _wait_preferences(desktop, printer: str, handle, timeout: float, log=None):
    """等到属性页窗口出现；printui 秒退时立即放弃（交给外层重试）。"""
    started = time.monotonic()
    deadline = started + max(2.0, timeout)
    exited_at: float | None = None
    while True:
        window = _usable_preferences(desktop, printer, log=log)
        if window is not None:
            return window, "已出现"
        now = time.monotonic()
        if _process_alive(handle) is False:
            if exited_at is None:
                exited_at = now
            elif now - exited_at >= PRINTUI_EXIT_GRACE:
                return None, "启动的 printui 进程已退出，且没有创建任何窗口"
        if now >= deadline:
            if exited_at is not None:
                return None, "启动的 printui 进程已退出，且没有创建任何窗口"
            return None, f"{deadline - started:.0f} 秒内没有出现属性页窗口"
        time.sleep(0.25)


def open_preferences(
    printer: str,
    timeout: float = 25,
    desktop=None,
    log=None,
    launcher: Callable[[str], object] | None = None,
    attempts: int = PREFERENCE_OPEN_ATTEMPTS,
):
    """打开驱动“打印首选项”窗口（rundll32 printui /e，等价于手工点开首选项）。

    上一轮下发被停止/中断后，驱动往往还在收尾，此时 printui 可能**静默退出**、
    一个窗口都不建。旧实现会一直等到超时才判失败，于是“停止后再勾选一批直接失败”。
    这里改为：复用已打开的窗口 → 盯着 printui 进程，秒退立即重试 → 每次重试前
    清掉残留窗口和进程。默认最多 3 次，单次预算按总超时均分。
    """
    desktop = desktop or _desktop()
    attempts = max(1, int(attempts))
    total = max(3.0, float(timeout))
    # 第一次给一半预算（冷启动的驱动可能慢），剩余预算均分给重试
    first = max(PREFERENCE_MIN_BUDGET, total * 0.5) if attempts > 1 else total
    rest = max(PREFERENCE_MIN_BUDGET, total * 0.5 / max(1, attempts - 1))
    notes: list[str] = []
    for attempt in range(1, attempts + 1):
        budget = total if attempts == 1 else (first if attempt == 1 else rest)
        existing = _usable_preferences(desktop, printer, log=log)
        if existing is not None:
            if log:
                log(f"复用已打开的驱动属性页：{_text_of(existing) or printer}")
            return existing

        handle = _launch(printer, launcher)
        window, note = _wait_preferences(desktop, printer, handle, budget, log=log)
        if window is not None:
            if log:
                log(f"已打开驱动属性页：{_text_of(window) or printer}")
            return window

        notes.append(f"第 {attempt} 次：{note}")
        _terminate_launcher(handle, log=log)
        for leftover in _preferences_windows(desktop, printer):
            _close_preference_window(leftover, log=log)
        if attempt < attempts:
            if log:
                log(f"驱动属性页未出现（{note}），清理残留后重试（第 {attempt + 1}/{attempts} 次）")
            time.sleep(PREFERENCE_RETRY_PAUSE)

    raise DriverUiError(
        f"驱动“打印首选项”窗口未出现（{printer}）；已尝试 {attempts} 次。"
        + ("原因：" + "；".join(notes) + "。" if notes else "")
        + "请确认驱动已正确安装；若刚停止过上一轮打印，请等打印机队列清空后重试。"
    )


def close_stale_preferences(printer: str, desktop=None, log=None) -> int:
    """清理该打印机残留的属性页窗口与 printui 宿主进程（用例之间/运行之间调用）。"""
    closed = 0
    try:
        desktop = desktop or _desktop()
        for window in _preferences_windows(desktop, printer):
            if _close_preference_window(window, log=log):
                closed += 1
    except Exception as exc:  # noqa: BLE001
        if log:
            log(f"清理残留驱动属性页失败：{type(exc).__name__}: {exc}")
    with _launcher_lock:
        handles = list(_live_launchers)
    for handle in handles:
        if _terminate_launcher(handle, log=log):
            closed += 1
    return closed


def _tab_names(window) -> list[str]:
    return [_text_of(item).strip() for item in _descendants(window, control_type="TabItem")]


def _switch_tab(window, name: str, log=None) -> bool:
    for tab in _descendants(window, control_type="TabItem"):
        if _norm(_text_of(tab)) == _norm(name) or _norm(name) in _norm(_text_of(tab)):
            tab.click_input()
            time.sleep(0.35)
            if log:
                log(f"驱动属性页：切换到“{_text_of(tab)}”页")
            return True
    return False


def _visible_combos(window) -> list:
    """当前页可见的下拉框（按界面从上到下顺序）。"""
    combos = _descendants(window, control_type="ComboBox")
    visible = []
    for combo in combos:
        try:
            if combo.is_visible():
                visible.append(combo)
        except Exception:  # noqa: WPS420
            visible.append(combo)
    return visible


def _combos_ready(window, need: int, timeout: float = 5.0) -> list:
    """等待当前页至少渲染出 need 个可见下拉框。

    驱动切换页签、或改了“纸张尺寸”后会重建纸张页控件，立即枚举只能看到部分控件；
    固定等待时间在不同机器上不稳定，改为轮询直到数量就绪或超时。
    """
    deadline = time.monotonic() + timeout
    combos = _visible_combos(window)
    while len(combos) < need and time.monotonic() < deadline:
        time.sleep(0.2)
        combos = _visible_combos(window)
    return combos


def _select_combo_index(combo, index: int) -> None:
    """按序号选中下拉项。

    这类驱动的下拉项是自绘的，UIA 读不到条目文本，但键盘定位稳定有效：
    聚焦 → Home 回第一项 → 下移 index 次 → 回车。
    """
    combo.click_input()
    time.sleep(0.25)
    combo.type_keys("{HOME}")
    time.sleep(0.15)
    for _ in range(int(index)):
        combo.type_keys("{DOWN}")
        time.sleep(0.1)
    time.sleep(0.15)
    combo.type_keys("{ENTER}")
    time.sleep(0.25)


def _combo_options(combo) -> list[str]:
    try:
        return [str(item) for item in (getattr(combo, "texts", lambda: [])() or [])]
    except Exception:  # noqa: BLE001
        return []


def _combo_current_index(combo) -> int | None:
    """读取下拉框当前选中项下标；自绘/不可读时返回 None。"""
    try:
        return int(combo.selected_index())
    except Exception:  # noqa: BLE001
        return None


def _combo_current_text(combo) -> str:
    """读取下拉框当前选中项文本；不可读时返回空串。"""
    try:
        return str(combo.selected_text() or "")
    except Exception:  # noqa: BLE001
        return ""


def _combo_already_selected(combo, action: "Action") -> bool:
    """驱动界面当前是否已选中目标项。

    这类驱动的下拉是自绘的，``selected_index()`` 会因条目列表读不到而报错，
    但 ``selected_text()`` 能返回当前选中文本，因此以文本比对为准，索引比对兜底。
    """
    target = str(action.target_text or "").strip()
    current = _combo_current_text(combo).strip()
    if target and current:
        a, b = _norm(current), _norm(target)
        if a and b and (a == b or a in b or b in a):
            return True
    if action.kind == "combo_index":
        return _combo_current_index(combo) == int(action.value)
    return False


def _select_combo_text(combo, wanted: str) -> bool:
    options = _combo_options(combo)
    match = _best_match(options, wanted)
    if match is None:
        return False
    combo.select(match)
    return True


def _apply_action(window, action: Action, log=None) -> None:
    if action.kind in {"combo_index", "combo_text"}:
        position = PAPER_TAB_COMBOS.index(action.label) if action.label in PAPER_TAB_COMBOS else -1
        combo = None
        if position >= 0:
            # 纸张页的下拉框按视觉顺序排列（尺寸及纸型/纸张尺寸/来源…）。
            # 注意：驱动的 UIA 标签与视觉标签错位（视觉“纸张类型”那行的 ComboBox 名叫“纸张尺寸”），
            # 因此只能按位置取，不能拿 UIA 名称做校验。
            combos = _combos_ready(window, position + 1)
            if position < len(combos):
                combo = combos[position]
        if combo is None:
            for candidate in _visible_combos(window):
                if _matches_labels(_text_of(candidate), (action.label,)):
                    combo = candidate
                    break
            if combo is None:
                raise DriverUiError(f"驱动属性页里找不到“{action.label}”下拉框。")
        if _combo_already_selected(combo, action):
            # 界面里已经是目标项：DEVMODE 回读自然不变，属正常而非失败
            if log:
                log(f"驱动属性页：{action.label} 已是目标项（{action.target_text or action.value}），无需变更")
            return False
        if action.kind == "combo_index":
            _select_combo_index(combo, int(action.value))
        else:
            # 驱动能读到条目时按名称选；读不到（自绘下拉）时给出明确原因
            if not _select_combo_text(combo, str(action.value)):
                raise DriverUiError(
                    f"驱动未暴露“{action.label}”的可读条目，无法按名称选择“{action.value}”；"
                    "请改用驱动下拉里的完整名称，或填数字序号。"
                )
        if log:
            log(f"驱动属性页：{action.label} → {action.describe()}")
        return True

    if action.kind == "radio":
        words = action.value if isinstance(action.value, (tuple, list)) else (action.value,)
        for radio in _descendants(window, control_type="RadioButton"):
            if _matches_labels(_text_of(radio), words):
                radio.click_input()
                time.sleep(0.15)
                if log:
                    log(f"驱动属性页：{action.label} → {_text_of(radio)}")
                return True
        raise DriverUiError(f"驱动属性页里找不到“{action.label}”选项：{'/'.join(words)}")

    if action.kind == "toggle":
        words = (action.label,)
        for check in _descendants(window, control_type="CheckBox"):
            if _matches_labels(_text_of(check), words):
                try:
                    current = check.get_toggle_state()
                except Exception:  # noqa: WPS420
                    current = None
                if current != (1 if action.value else 0):
                    check.click_input()
                    time.sleep(0.15)
                if log:
                    log(f"驱动属性页：{action.describe()}")
                return True
        # 开关型下拉（如“跳过空白页”：关闭/开启）
        for combo in _visible_combos(window):
            if _matches_labels(_text_of(combo), words):
                if not _select_combo_text(combo, "开启" if action.value else "关闭"):
                    raise DriverUiError(f"“{action.label}”下拉里找不到开启/关闭选项。")
                if log:
                    log(f"驱动属性页：{action.describe()}")
                return True
        raise DriverUiError(f"驱动属性页里找不到“{action.label}”勾选框。")

    if action.kind == "int_text":
        for edit in _descendants(window, control_type="Edit"):
            if _matches_labels(_text_of(edit), (action.label,)):
                edit.set_edit_text(str(action.value))
                if log:
                    log(f"驱动属性页：{action.describe()}")
                return True
        # 驱动常把输入框的标签错位挂到前一个控件上，退化为按控件名前缀匹配
        for edit in _descendants(window, control_type="Edit"):
            if _norm(action.label) in _norm(_text_of(edit)):
                edit.set_edit_text(str(action.value))
                if log:
                    log(f"驱动属性页：{action.describe()}")
                return True
        raise DriverUiError(f"驱动属性页里找不到“{action.label}”输入框。")

    raise DriverUiError(f"不支持的操作类型：{action.kind}")


def _confirm(window, log=None) -> None:
    for button in _descendants(window, control_type="Button"):
        if _norm(_text_of(button)) in {_norm("确定"), _norm("ok")}:
            button.click_input()
            if log:
                log("驱动属性页：已点击“确定”，设置交由驱动保存")
            return
    raise DriverUiError("驱动属性页里找不到“确定”按钮。")


def _close(window, log=None) -> bool:
    """关掉属性页且确认真的关掉了（关不掉就结束 printui 宿主）。"""
    return _close_preference_window(window, log=log)


# ---------------------------------------------------------------------------
# 默认 DEVMODE 快照 / 还原 / 回读校验
# ---------------------------------------------------------------------------

def _win32print():
    if platform.system() != "Windows":
        raise DriverUiError("驱动属性页自动化仅支持 Windows。")
    try:
        import win32print  # noqa: WPS433
    except ImportError as exc:
        raise DriverUiError("缺少 pywin32，无法读取/还原打印机默认设置。") from exc
    return win32print


def snapshot(printer: str, w=None) -> dict:
    """只读读取驱动当前默认 DEVMODE（用于校验与还原）。

    只需要 PRINTER_ACCESS_USE：普通用户会话下请求 PRINTER_ALL_ACCESS /
    PRINTER_ACCESS_ADMINISTER 会被拒绝（Win32 错误 5），导致整个下发链路
    在“设置属性”之前就中断。只读快照与回读校验都属于只读操作。
    """
    w = w or _win32print()
    handle = w.OpenPrinter(printer, {"DesiredAccess": w.PRINTER_ACCESS_USE})
    try:
        info = w.GetPrinter(handle, 2)
        devmode = info.get("pDevMode")
        if devmode is None:
            raise DriverUiError(f"打印机“{printer}”的驱动未提供可读的 DEVMODE。")
        return {
            "info": info,
            "devmode": devmode,
            "PaperSize": int(devmode.PaperSize),
            "MediaType": int(devmode.MediaType),
            "DefaultSource": int(devmode.DefaultSource),
            "Duplex": int(devmode.Duplex),
            "Color": int(devmode.Color),
            "Copies": int(devmode.Copies),
            "PrintQuality": int(devmode.PrintQuality),
            "Collate": int(devmode.Collate),
        }
    finally:
        w.ClosePrinter(handle)


def restore(before: dict, printer: str, w=None, log=None) -> None:
    """把驱动默认 DEVMODE 还原成执行前快照（只写回原值，不参与属性设置）。

    写回打印机全局默认设置需要管理权限，普通用户会话下会被拒绝（错误 5）。
    此时降级为明确警告，绝不能把已经成功下发的作业改判为失败；调用方通过
    返回值/日志即可知道需要人工核对。
    """
    w = w or _win32print()
    failures: list[str] = []
    # 部分系统/驱动上 OpenPrinter(ADMINISTER) 会“假成功”返回句柄，但紧接着的
    # GetPrinter 被拒绝；因此必须把 打开→读取→写回 当作一个完整探测，逐档尝试。
    for access_name in ("PRINTER_ACCESS_ADMINISTER", "PRINTER_ACCESS_USE"):
        access = getattr(w, access_name, None)
        if access is None:
            continue
        handle = None
        try:
            try:
                handle = w.OpenPrinter(printer, {"DesiredAccess": access})
            except Exception as open_error:  # noqa: WPS420
                failures.append(f"{access_name} 打开失败：{open_error}")
                continue
            info = w.GetPrinter(handle, 2)
            info["pDevMode"] = before["devmode"]
            w.SetPrinter(handle, 2, info, 0)
        except Exception as step_error:  # noqa: WPS420
            failures.append(f"{access_name} 读写失败：{step_error}")
            continue
        finally:
            if handle is not None:
                w.ClosePrinter(handle)
        if log:
            log("驱动默认设置已还原为用例执行前的快照。")
        return True
    if log:
        log("当前账户无写回打印机默认设置的权限，未能自动还原驱动默认值（"
            + "；".join(failures) + "）；请在打印首选项中手工核对，本次打印结果不受影响。")
    return False


def verify(printer: str, actions: list[Action], before: dict, w=None, unchanged_labels: set[str] | None = None) -> list[str]:
    """回读默认 DEVMODE，核对每条可校验的操作是否真的生效。

    ``unchanged_labels``：设置前界面里就已是目标项的动作。这类动作不需要改动界面，
    DEVMODE 回读保持不变是正常结果（例如目标“普通纸”恰好等于驱动默认值），不能判失败。
    """
    unchanged_labels = unchanged_labels or set()
    after = snapshot(printer, w=w)
    problems: list[str] = []
    for action in actions:
        if not action.expect_field:
            continue
        actual = after.get(action.expect_field)
        expected = action.expect_value
        if expected is None:
            # 无固定期望值（如驱动私有介质编码）：只要求相对快照发生变化且为正值。
            # 例外：设置前已是目标项时，回读不变即正确。
            if action.label in unchanged_labels:
                if not actual:
                    problems.append(f"{action.label}：回读值无效（{action.expect_field}={actual}）")
                continue
            if actual == before.get(action.expect_field) or not actual:
                problems.append(f"{action.label}：驱动默认值未变化（{action.expect_field}={actual}）")
            continue
        if actual != expected:
            problems.append(f"{action.label}（期望 {action.expect_field}={expected}，实际 {actual}）")
    return problems


def _apply_extended(window, attrs, log) -> list[str]:
    """驱动扩展属性（水印/装订/作业类型等）同样在驱动属性页里设置。"""
    from .ui_print import apply_driver_attrs  # noqa: WPS433

    return apply_driver_attrs(window, list(attrs), log or (lambda _text: None))


def apply_case(
    printer: str,
    case: PrintCase,
    caps: dict | None = None,
    log: Callable[[str], None] | None = None,
    desktop=None,
    w=None,
) -> Callable[[], None]:
    """按用例在驱动属性页设置属性，回读确认后返回还原函数。"""
    actions = plan_actions(case, caps)
    before = snapshot(printer, w=w)
    if not actions:
        if log:
            log("用例未填写需要驱动界面设置的属性，保留驱动当前默认设置。")
        return lambda: None

    window = None
    unchanged: set[str] = set()
    try:
        window = open_preferences(printer, desktop=desktop, log=log)
        pages = {action.page for action in actions}
        for page in sorted(pages, key=lambda name: [TAB_BASIC, TAB_PAPER, TAB_LAYOUT].index(name) if name in {TAB_BASIC, TAB_PAPER, TAB_LAYOUT} else 9):
            if not _switch_tab(window, page, log):
                missing = [action.label for action in actions if action.page == page]
                raise DriverUiError(f"驱动属性页没有“{page}”页，无法设置：{'、'.join(missing)}")
            for action in [item for item in actions if item.page == page]:
                if _apply_action(window, action, log) is False:
                    unchanged.add(action.label)
        if case.extended_attrs:
            unresolved = _apply_extended(window, case.extended_attrs, log)
            if unresolved:
                raise DriverUiError("以下驱动扩展属性无法在驱动界面设置：" + "、".join(unresolved))
        _confirm(window, log)
        window = None
        time.sleep(0.6)
    except DriverUiError:
        if window is not None:
            _close(window, log=log)
        raise
    except Exception as exc:  # noqa: BLE001
        if window is not None:
            _close(window, log=log)
        raise DriverUiError(f"驱动属性页操作失败：{type(exc).__name__}: {exc}") from exc

    problems = verify(printer, actions, before, w=w, unchanged_labels=unchanged)
    if problems:
        restore(before, printer, w=w, log=log)
        raise DriverUiError("驱动未接受以下设置（写入后回读不一致）：" + "、".join(problems))
    if log:
        summary = "、".join(action.describe() for action in actions)
        log(f"驱动属性页设置完成并回读确认：{summary}")

    def _restore() -> None:
        restore(before, printer, w=w, log=log)

    return _restore


def available(printer: str, timeout: float = 20) -> tuple[bool, str]:
    """自检：能否打开该打印机的驱动属性页（不修改任何设置）。"""
    try:
        window = open_preferences(printer, timeout=timeout)
    except Exception as exc:  # noqa: BLE001
        return False, f"驱动属性页打不开：{exc}"
    try:
        tabs = [name for name in _tab_names(window) if name]
        _close(window)
    except Exception as exc:  # noqa: BLE001
        return False, f"驱动属性页控件不可读：{exc}"
    if not tabs:
        return False, "驱动属性页没有可识别的页签"
    return True, f"驱动属性页就绪（{'/'.join(tabs[:4])}…）"
