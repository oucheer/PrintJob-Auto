"""pywinauto UI 自动化打印通道。

按真实操作路径下发打印作业（不依赖任何第三方打印工具）：

    打开文档 → Ctrl+P 调出打印对话框 → 选择打印机驱动
    → 打开“首选项/属性” → 按用例属性设置纸张/纸盒/双面/质量/无边距
    → 确定 → 点击“打印”

设计约定：
- 每一步都通过 log 回调输出，前端“执行日志”页可完整看到操作过程。
- 用例显式填写、但在驱动界面无法定位或设置成功的属性会抛出 UiPrintError，
  用例判定失败且不下发（与 DEVMODE 严格模式一致，避免属性静默失效）。
- 仅适用于能在 Word/WPS 中打开的文档；PDF/图片没有可自动化的应用界面，
  仍由内置 GDI 直打通道处理。
- 所有 pywinauto 控件对象均通过参数注入，便于单元测试用替身验证流程。
"""
from __future__ import annotations

import time
from pathlib import Path
from typing import Callable

from .attr_catalog import ATTRIBUTE_BY_LABEL, page_keywords
from .cominit import ensure_sta, import_ui_engine, ui_ready
from .extended_attrs import (
    ExtendedAttr,
    base_label,
    bool_value,
    group_by_page,
    split_composite,
    split_pair,
)
from .models import PrintCase
from .office import WPS_PROG_IDS
from .winnative import (
    enum_top_windows,
    process_alive,
    process_name,
    terminate_pid,
    window_pid,
    window_title,
)

UI_PRINTABLE_EXTENSIONS = {".docx", ".doc", ".rtf", ".odt", ".wps", ".wpt"}

PRINT_DIALOG_TITLES = ("打印", "print")
PREFERENCE_BUTTON_WORDS = ("首选项", "属性", "preferences", "properties", "打印机属性", "高级")
CONFIRM_BUTTON_WORDS = ("打印", "print")  # 优先真正的“打印”按钮，避免误点“确定”
# 包含“打印”但绝不是提交按钮的控件文本（模糊匹配时排除）
PRINT_BUTTON_NOISE = ("打印预览", "打印到文件", "print preview", "print to file")

# 各属性的驱动界面标签关键字（中英文驱动常见写法）
FIELD_LABELS = {
    "paper": ("纸张尺寸", "纸张大小", "paper size", "size"),
    "media": ("纸张类型", "介质类型", "paper type", "media type"),
    "bin": ("纸张来源", "纸盒", "进纸", "paper source", "tray", "source"),
    "duplex": ("双面", "duplex", "打印方式", "print on both sides", "2-sided"),
    "quality": ("打印质量", "print quality", "分辨率", "quality"),
    "borderless": ("无边距", "borderless", "满版"),
}

DUPLEX_TOKENS = {
    "simplex": ("单面", "仅一面", "1-sided", "simplex", "none"),
    "duplex": ("双面长边", "双面", "长边", "duplex", "long edge"),
    "duplexlong": ("双面长边", "长边", "long edge", "flip on long edge", "duplex"),
    "duplexshort": ("双面短边", "短边", "short edge", "flip on short edge"),
}


class UiPrintError(RuntimeError):
    """UI 自动化打印失败（窗口/控件未找到、属性无法设置等）。"""


class UiPrintInterrupted(UiPrintError):
    """用户在执行过程中点击了“停止”，本次下发被主动中止。"""


def _norm(text: object) -> str:
    return "".join(str(text or "").casefold().split()).replace("(", "").replace(")", "")


def _short_token(value: str) -> str:
    """去掉尺寸/克重括号后缀：'A4 (210 x 297mm)' → 'A4'，'厚纸2(158-209gsm)' → '厚纸2'。"""
    return str(value or "").split("(")[0].strip()


def _best_match(options: list[str], wanted: str) -> str | None:
    """在驱动下拉项里匹配目标值：先精确，再双向包含。"""
    target = _norm(_short_token(wanted))
    if not target:
        return None
    normalized = [(option, _norm(option)) for option in options if str(option).strip()]
    for option, text in normalized:
        if text == target:
            return option
    for option, text in normalized:
        if text and (target in text or text in target):
            return option
    return None


def _text_of(control) -> str:
    try:
        return str(control.window_text() or "")
    except Exception:  # noqa: WPS420
        return ""


def _descendants(container, **kwargs):
    try:
        return list(container.descendants(**kwargs))
    except Exception:  # noqa: WPS420
        return []


def _label_before(control) -> str:
    """取控件前方最近的文本标签（UI 自动化里用相邻 Text 推断字段名）。"""
    try:
        parent = control.parent()
        raw = getattr(parent, "children", None)
        # pywinauto 的 children 是方法；替身控件里是列表属性，两种都兼容
        siblings = list(raw() if callable(raw) else (raw or []))
    except Exception:  # noqa: WPS420
        return ""
    label = ""
    for sibling in siblings:
        if sibling == control:
            return label
        text = _text_of(sibling)
        if text.strip():
            label = text
    return ""


def _matches_labels(text: str, labels) -> bool:
    normalized = _norm(text)
    return any(_norm(word) in normalized for word in labels if word)


def _wait(condition: Callable[[], object], timeout: float, message: str):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        result = condition()
        if result:
            return result
        time.sleep(0.4)
    raise UiPrintError(message)


def _default_desktop():
    # 通过 cominit 导入，保证 comtypes 与本线程 COM 单元模式一致
    import_ui_engine()
    from pywinauto import Desktop  # noqa: WPS433

    return Desktop(backend="uia")


def _pause(seconds: float, should_stop=None) -> bool:
    """等待一小段并响应停止请求；返回是否因“停止”提前结束。"""
    deadline = time.monotonic() + max(0.0, seconds)
    while time.monotonic() < deadline:
        if should_stop is not None and should_stop():
            return True
        time.sleep(min(0.1, max(0.02, deadline - time.monotonic())))
    return bool(should_stop is not None and should_stop())


def _desktop_windows(desktop) -> list:
    """UIA 桌面窗口列表；引擎异常时当作没有窗口，不让整条链路崩掉。"""
    try:
        return list(desktop.windows())
    except Exception:  # noqa: BLE001
        return []


def _wrap_handle(desktop, hwnd: int):
    """把原生句柄包装成 pywinauto 窗口对象；失败返回 None。"""
    if not hwnd:
        return None
    try:
        return desktop.window(handle=int(hwnd))
    except Exception:  # noqa: BLE001
        return None


def _native_document_handle(title: str, prefer_pid: int | None = None) -> int | None:
    """原生枚举找文档窗口句柄：标题包含文件名主干，优先本进程启动的那个实例。"""
    wanted = _norm(title)
    if not wanted:
        return None
    candidates = [
        hwnd for hwnd, text in enum_top_windows()
        if wanted in _norm(text)
    ]
    if not candidates:
        return None
    if prefer_pid:
        for hwnd in candidates:
            if window_pid(hwnd) == prefer_pid:
                return hwnd
    return candidates[0]


def _app_window_handle(app, title: str) -> int | None:
    """Word/WPS 通过 COM 直接给出的窗口句柄（Documents.Open 返回后立即可用）。"""
    if app is None:
        return None
    try:
        hwnd = int(app.ActiveWindow.Hwnd)
    except Exception:  # noqa: BLE001 - 没有活动文档/旧版 WPS 无该属性
        return None
    if not hwnd:
        return None
    wanted = _norm(title)
    if wanted and wanted not in _norm(window_title(hwnd)):
        return None
    return hwnd


DOC_WINDOW_POLL = 0.25         # 文档窗口轮询间隔
DOC_WINDOW_UIA_RETRY = 4.0     # 原生一直找不到时，才低频试一次 UIA 全量枚举


def _find_window(desktop, title: str, timeout: float, should_stop=None, app=None,
                 prefer_pid: int | None = None, log: Callable[[str], None] | None = None):
    """按标题找文档窗口（标题包含文件名主干）。

    旧实现每 0.4 秒调一次 ``Desktop.windows()``（UIA 桌面全量枚举）。UIA 在这个
    环境里会整体失效：文档明明开着、屏幕上看得见，枚举却返回不完整列表，于是
    一路空等，把用例超时（90 秒）耗光——用户看到的就是“点停止后卡住不动”。
    这里改成：COM 给的 HWND → 原生 EnumWindows（毫秒级）→ UIA 全量枚举兜底。
    """
    wanted = _norm(title)
    started = time.monotonic()
    deadline = started + max(1.0, float(timeout))
    uia_next = started
    wrap_error = ""
    seen_titles: dict[int, str] = {}
    while time.monotonic() < deadline:
        if should_stop is not None and should_stop():
            raise UiPrintInterrupted("收到停止请求，已中止本次下发（等待文档窗口阶段）")
        hwnd = _app_window_handle(app, title)
        source = "COM 句柄"
        if not hwnd:
            hwnd = _native_document_handle(title, prefer_pid=prefer_pid)
            source = "原生枚举"
        if hwnd:
            window = _wrap_handle(desktop, hwnd)
            if window is not None:
                if log is not None:
                    log(
                        f"文档窗口已就绪（{source}，用时 {time.monotonic() - started:.1f} 秒）："
                        f"{window_title(hwnd) or title}"
                    )
                return window
            wrap_error = f"窗口句柄 {hwnd} 已存在，但 UI 自动化读不到它"
            seen_titles[hwnd] = window_title(hwnd)
        elif time.monotonic() >= uia_next:
            uia_next = time.monotonic() + DOC_WINDOW_UIA_RETRY
            for window in _desktop_windows(desktop):
                if wanted and wanted in _norm(_text_of(window)):
                    if log is not None:
                        log(f"文档窗口已就绪（UIA 兜底，用时 {time.monotonic() - started:.1f} 秒）：{_text_of(window)}")
                    return window
        _pause(DOC_WINDOW_POLL, should_stop)
    detail = ""
    if wrap_error:
        detail = f"；{wrap_error}，请重启本程序后重试"
    elif seen_titles:
        detail = "；已检测到窗口但读不到控件树：" + "、".join(seen_titles.values())
    else:
        opened = [
            text for _hwnd, text in enum_top_windows()
            if "word" in text.casefold() or "wps" in text.casefold()
        ]
        if opened:
            detail = "；当前桌面上的 Word/WPS 窗口：" + "、".join(opened[:3])
    raise UiPrintError(f"未找到文档窗口：{title}（请确认文档已用 Word/WPS 打开）{detail}")


def _find_print_dialog(desktop, timeout: float):
    """找打印对话框（Windows 通用对话框 class #32770，标题含“打印/Print”）。"""
    def search():
        for window in desktop.windows():
            try:
                class_name = window.class_name()
            except Exception:  # noqa: WPS420
                class_name = ""
            text = _text_of(window)
            if class_name == "#32770" and _matches_labels(text, PRINT_DIALOG_TITLES):
                return window
        return None

    return _wait(search, timeout, "未能调出打印对话框（Ctrl+P 后未出现“打印”窗口）")


def _select_printer(dialog, printer: str, log: Callable[[str], None]) -> None:
    """在打印对话框里选择目标打印机驱动。"""
    combos = _descendants(dialog, control_type="ComboBox")
    for combo in combos:
        options = [str(item) for item in (getattr(combo, "texts", lambda: [])() or [])]
        match = _best_match(options, printer)
        if match is not None:
            combo.select(match)
            log(f"打印对话框：已选择打印机驱动 {match}")
            return
    # 下拉未展开时按名称直接在编辑框输入
    for combo in combos:
        try:
            combo.edit.set_edit_text(printer)
            log(f"打印对话框：已在打印机输入框填入 {printer}")
            return
        except Exception:  # noqa: WPS420
            continue
    raise UiPrintError(f"打印对话框里找不到打印机“{printer}”的下拉项；请确认该驱动已安装。")


def _set_combo_field(container, labels, wanted: str, field_name: str, log) -> bool:
    for combo in _descendants(container, control_type="ComboBox"):
        label = _label_before(combo)
        if not _matches_labels(label, labels):
            continue
        options = [str(item) for item in (getattr(combo, "texts", lambda: [])() or [])]
        match = _best_match(options, wanted)
        if match is None:
            raise UiPrintError(
                f"{field_name}：驱动下拉里找不到“{wanted}”（可选：{'、'.join(options[:6]) or '无'}）"
            )
        combo.select(match)
        log(f"驱动属性：{field_name} 设置为 {match}")
        return True
    return False


def _set_checkbox_field(container, labels, enabled: bool, field_name: str, log) -> bool:
    for check in _descendants(container, control_type="CheckBox"):
        if not _matches_labels(_text_of(check), labels):
            continue
        try:
            if check.get_toggle_state() != (1 if enabled else 0):
                check.toggle()
        except Exception:  # noqa: WPS420
            check.click_input()
        log(f"驱动属性：{field_name} = {'是' if enabled else '否'}")
        return True
    return False


def _find_print_dialog_or_backstage(desktop, window, timeout: float, should_stop=None):
    """返回 (容器, 是否后台视图)。

    Word 2013+ 的 Ctrl+P 打开的是“打印”后台视图（嵌在主窗口内），
    而不是独立对话框；WPS/旧版 Word 才是 #32770 对话框。两种都要支持。
    """
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if should_stop is not None and should_stop():
            raise UiPrintInterrupted("收到停止请求，已中止本次下发（等待打印界面阶段）")
        for candidate in _desktop_windows(desktop):
            try:
                class_name = candidate.class_name()
            except Exception:  # noqa: WPS420
                class_name = ""
            text = _text_of(candidate)
            if class_name == "#32770" and _matches_labels(text, PRINT_DIALOG_TITLES):
                return candidate, False
        # 后台视图：主窗口内出现“选择打印机/打印机”下拉
        try:
            combo = _backstage_printer_combo(window)
        except Exception:  # noqa: BLE001 - Word 忙时 UIA 会抛 COMError，继续等
            combo = None
        if combo is not None:
            return window, True
        _pause(0.35, should_stop)
    raise UiPrintError("未能调出打印界面（Ctrl+P 后未出现“打印”对话框或后台打印视图）。")


# 调出打印界面的重试策略：Word 冷启动/焦点被抢时首次 Ctrl+P 常被吞掉，
# 表现为“文档窗口已就绪”之后长时间没有任何反应，最后超时判失败。
PRINT_VIEW_WAIT_PER_TRY = 8.0
PRINT_VIEW_MAX_ATTEMPTS = 6
BLOCKING_BANNER_WORDS = ("启用编辑", "enable editing", "启用内容", "enable content")

# 点击“打印”后的确认策略。Word 的后台打印视图在收到打印指令后会关闭并回到文档；
# 如果点击没有被接受（按钮还在渲染、视图被遮挡），视图会一直停在原处。
# 旧实现点完就当作“已提交”，0.8 秒后直接关 Word——作业还没进队列就被结束，
# 表现为“日志说有 3 条，实际只出来 2 份”。
PRINT_CLICK_ATTEMPTS = 2        # 点击后视图不关闭时，再点一次的上限
PRINT_START_TIMEOUT = 6.0       # 单击后等打印视图关闭的时长
PRINT_BUTTON_READY_TIMEOUT = 10.0  # 等“打印”按钮从禁用变为可点击（预览还在渲染）
PRINT_CONFIRM_WAIT = 2.0        # 等 Word 报告后台打印任务的时长
PRINT_DRAIN_TIMEOUT = 45.0      # 等 Word 把数据交给打印后台的上限
PRINT_CLOSE_GRACE = 0.8         # 未确认时关闭前的固定等待
_HANDOFF_LABELS = {
    "handoff": "Word 后台打印任务已完成交接",
    "none": "仅确认打印视图已关闭，未读到 Word 后台任务",
    "dialog": "打印对话框已关闭",
    "timeout": "等待 Word 后台任务超时",
}


def _dismiss_blocking_banner(window, log: Callable[[str], None] | None = None) -> bool:
    """点掉“受保护的视图/启用编辑”一类横幅，否则键盘输入进不到正文。"""
    for button in _descendants(window, control_type="Button"):
        text = _text_of(button)
        if not _matches_labels(text, BLOCKING_BANNER_WORDS):
            continue
        try:
            button.click_input()
        except Exception:  # noqa: WPS420
            continue
        if log:
            log(f"文档窗口有“{text.strip()}”提示，已点击以恢复可编辑状态")
        time.sleep(0.6)
        return True
    return False


def _ensure_document_focus(window, log: Callable[[str], None] | None = None) -> bool:
    """把焦点真正放到文档正文上；返回是否点进了正文。

    Word 刚启动时焦点可能停在标题栏或加载中的窗格，此时 Ctrl+P 会被吞掉，
    打印界面永远不出现。这里先激活窗口、清掉拦截横幅，再点一下正文。
    """
    try:
        window.set_focus()
    except Exception:  # noqa: WPS420
        pass
    _dismiss_blocking_banner(window, log)
    try:
        window.set_focus()
    except Exception:  # noqa: WPS420
        pass
    for control_type in ("Document", "Edit"):
        for control in _descendants(window, control_type=control_type):
            try:
                if not _is_visible(control):
                    continue
                control.click_input()
                return True
            except Exception:  # noqa: WPS420
                continue
    return False


def _open_print_view(desktop, window, log: Callable[[str], None], timeout: float, should_stop=None):
    """反复尝试 Ctrl+P 直到打印界面出现；返回 (容器, 是否后台视图)。

    单次尝试只等一小段，失败就重新聚焦再发一次 Ctrl+P，避免把一次被吞掉的
    按键变成 90 秒空等 + 用例失败（批量下发时首条用例最容易踩到）。
    """
    deadline = time.monotonic() + max(1.0, timeout)
    attempts = 0
    last_error = ""
    while attempts < PRINT_VIEW_MAX_ATTEMPTS and time.monotonic() < deadline:
        if should_stop is not None and should_stop():
            raise UiPrintInterrupted("收到停止请求，已中止本次下发（调出打印界面阶段）")
        attempts += 1
        focused = _ensure_document_focus(window, log)
        try:
            window.type_keys("^p")
        except Exception as exc:  # noqa: BLE001
            last_error = f"发送 Ctrl+P 失败：{type(exc).__name__}: {exc}"
        wait_for = min(PRINT_VIEW_WAIT_PER_TRY, max(1.0, deadline - time.monotonic()))
        log(f"已发送 Ctrl+P（第 {attempts} 次{f'，已聚焦正文' if focused else ''}），等待打印界面…")
        try:
            return _find_print_dialog_or_backstage(desktop, window, wait_for, should_stop)
        except UiPrintInterrupted:
            raise
        except UiPrintError as exc:
            last_error = str(exc)
            if time.monotonic() >= deadline:
                break
            log(f"第 {attempts} 次未出现打印界面，重新聚焦后重试")
    raise UiPrintError(
        f"未能调出打印界面（已尝试 {attempts} 次 Ctrl+P，仍未出现打印对话框或后台打印视图）。{last_error}"
    )


def _backstage_printer_combo(window):
    """Word 后台打印视图里的打印机下拉（其子文本含“选择打印机/打印机/Printer”）。"""
    for combo in _descendants(window, control_type="ComboBox"):
        texts = [str(item) for item in (getattr(combo, "texts", lambda: [])() or [])]
        joined = " ".join(texts)
        if "打印机" in joined or "Printer" in joined or "printer" in joined:
            return combo
    return None


def _select_printer_in_backstage(window, printer: str, log: Callable[[str], None]) -> bool:
    """后台视图选打印机：展开下拉 → 在列表项里点选目标驱动。"""
    try:
        combo = _backstage_printer_combo(window)
    except Exception as exc:  # noqa: BLE001 - Word 忙时 UIA 会抛 COMError
        log(f"后台打印视图：读取打印机下拉失败（{type(exc).__name__}），稍后继续")
        return False
    if combo is None:
        return False
    combo.click_input()
    time.sleep(0.6)
    try:
        items = _descendants(window, control_type="ListItem")
    except Exception:  # noqa: BLE001
        items = []
    for item in items:
        text = _text_of(item)
        if _norm(printer) and _norm(printer) in _norm(text):
            try:
                item.click_input()
            except Exception:  # noqa: BLE001
                continue
            log(f"后台打印视图：已选择打印机驱动 {text.strip() or printer}")
            return True
    log(f"后台打印视图：下拉中未找到 {printer}，无法保证作业发往目标打印机")
    return False


def _open_preferences(dialog, desktop, log: Callable[[str], None], timeout: float):
    """点开驱动“首选项/属性”对话框，返回驱动属性窗口。"""
    candidates = (
        _descendants(dialog, control_type="Button")
        + _descendants(dialog, control_type="Hyperlink")
        + _descendants(dialog, control_type="MenuItem")
    )
    for button in candidates:
        if _matches_labels(_text_of(button), PREFERENCE_BUTTON_WORDS):
            log(f"打印对话框：打开驱动属性（{_text_of(button)}）")
            button.click_input()
            return _wait(
                lambda: next(
                    (
                        window
                        for window in _desktop_windows(desktop)
                        if window != dialog
                        and _matches_labels(_text_of(window), ("属性", "首选项", "preferences", "properties", "设置"))
                    ),
                    None,
                ),
                timeout,
                "驱动“首选项/属性”窗口未出现。",
            )
    raise UiPrintError("打印对话框里找不到“首选项/属性”按钮，无法设置驱动属性。")


def _apply_case_attributes(driver_window, case: PrintCase, log: Callable[[str], None]) -> list[str]:
    """在驱动属性窗口里按用例设置各项属性；返回未能定位的必需属性列表。"""
    unresolved: list[str] = []

    paper = str(case.paper or "")
    if case.mixed_sizes == "yes":
        log("驱动属性：混合尺寸=是，跳过纸张尺寸（由驱动按每页尺寸处理）")
    elif _norm(paper) not in {"", "auto", "mixed"}:
        if not _set_combo_field(driver_window, FIELD_LABELS["paper"], paper, "纸张尺寸", log):
            unresolved.append(f"纸张尺寸 {paper}")

    if str(case.media_type or "").strip():
        if not _set_combo_field(driver_window, FIELD_LABELS["media"], case.media_type, "纸张类型", log):
            unresolved.append(f"纸张类型 {case.media_type}")

    if str(case.bin or "").strip():
        if not _set_combo_field(driver_window, FIELD_LABELS["bin"], case.bin, "纸张来源", log):
            unresolved.append(f"纸张来源 {case.bin}")

    if case.duplex and case.duplex != "simplex":
        tokens = DUPLEX_TOKENS.get(case.duplex, (case.duplex,))
        matched = False
        for token in tokens:
            try:
                if _set_combo_field(driver_window, FIELD_LABELS["duplex"], token, "双面", log):
                    matched = True
                    break
            except UiPrintError:
                continue
        if not matched:
            unresolved.append(f"双面 {case.duplex}")

    quality = str(case.print_quality or "")
    if quality and quality.lower() != "auto":
        if not _set_combo_field(driver_window, FIELD_LABELS["quality"], quality, "打印质量", log):
            unresolved.append(f"打印质量 {quality}")

    if case.borderless == "yes":
        if not _set_checkbox_field(driver_window, FIELD_LABELS["borderless"], True, "无边距", log):
            unresolved.append("无边距")

    if case.skip_blank == "yes":
        if not _set_checkbox_field(driver_window, ("跳过空白页", "skip blank"), True, "跳过空白页", log):
            log("驱动属性：未找到“跳过空白页”选项，按驱动默认执行（该属性不参与强校验）")

    return unresolved


def _confirm(window, log: Callable[[str], None], words=("确定", "ok", "应用")) -> None:
    for button in _descendants(window, control_type="Button"):
        if _matches_labels(_text_of(button), words):
            button.click_input()
            return
    raise UiPrintError("窗口中找不到“确定”按钮。")


def _set_copies(dialog, copies: int, log) -> None:
    if copies <= 1:
        return
    for edit in _descendants(dialog, control_type="Edit"):
        if _matches_labels(_label_before(edit), ("份数", "copies", "number of copies")):
            edit.set_edit_text(str(copies))
            log(f"打印对话框：份数 = {copies}")
            return
    raise UiPrintError(f"打印对话框里找不到“份数”输入框，无法保证按 {copies} 份打印。")


def _set_page_range(dialog, pages: str, log) -> None:
    text = str(pages or "").strip()
    if not text:
        return
    for radio in _descendants(dialog, control_type="RadioButton"):
        if _matches_labels(_text_of(radio), ("页码范围", "pages", "指定页")):
            radio.click_input()
            break
    for edit in _descendants(dialog, control_type="Edit"):
        if _matches_labels(_label_before(edit), ("页码范围", "pages", "页")):
            edit.set_edit_text(text)
            log(f"打印对话框：页码 = {text}")
            return
    raise UiPrintError(f"打印对话框里找不到“页码范围”输入框，无法保证按 {text} 打印。")


def _button_enabled(button) -> bool:
    """控件是否可点击；拿不到状态时按可点击处理（替身/旧引擎）。"""
    checker = getattr(button, "is_enabled", None)
    if checker is None:
        return True
    try:
        return bool(checker())
    except Exception:  # noqa: BLE001
        return True


def _print_buttons(container) -> list[tuple[object, str]]:
    """按“完全匹配优先”排序后返回 (按钮, 文本)。"""
    exact = {_norm(word) for word in CONFIRM_BUTTON_WORDS}
    ranked: list[tuple[int, object, str]] = []
    try:
        buttons = _descendants(container, control_type="Button")
    except Exception:  # noqa: BLE001 - Word 忙时 UIA 会抛 COMError，交给调用方重试
        return []
    for button in buttons:
        text = _text_of(button)
        if _matches_labels(text, PRINT_BUTTON_NOISE):
            continue
        if _norm(text) in exact:
            ranked.append((0, button, text))
        elif _matches_labels(text, CONFIRM_BUTTON_WORDS):
            ranked.append((1, button, text))
    ranked.sort(key=lambda item: item[0])
    return [(button, text) for _, button, text in ranked]


def _click_print_button(container, log: Callable[[str], None],
                        wait: float = PRINT_BUTTON_READY_TIMEOUT, should_stop=None) -> str | None:
    """点击“打印”按钮，返回实际点到的文本；找不到可点击的按钮返回 None。

    后台打印视图刚切换过来时，“打印”按钮可能还是禁用状态（预览仍在渲染），
    此时 click_input 不报错但也不生效——旧实现直接把这种情况记成“已提交”，
    于是出现“日志 3 条、实际只出 2 份”。这里先等按钮真正可点击再点。
    """
    deadline = time.monotonic() + max(0.0, wait)
    while True:
        if should_stop is not None and should_stop():
            raise UiPrintInterrupted("收到停止请求，已中止本次下发（等待“打印”按钮就绪阶段）")
        for button, text in _print_buttons(container):
            if not _button_enabled(button):
                continue
            button.click_input()
            return text
        if time.monotonic() >= deadline:
            return None
        _pause(0.2, should_stop)


def _print_view_gone(window) -> bool:
    """后台打印视图是否已关闭（读 UIA 失败时按“还在视图里”处理，避免误判）。"""
    try:
        return _backstage_printer_combo(window) is None
    except Exception:  # noqa: BLE001 - Word 忙时 UIA 会抛 COMError
        return False


def _left_print_view(window, timeout: float, should_stop=None) -> bool:
    """等 Word 后台打印视图关闭（回到文档）。

    点击“打印”后视图一定会关闭；视图不动说明这次点击没有被接受，
    此时把作业记成“已提交”就会丢单。判定“已关闭”后会再确认一次，
    避免视图切换瞬间读不到控件被误判成“已关闭”。
    """
    deadline = time.monotonic() + max(0.5, float(timeout))
    while time.monotonic() < deadline:
        if should_stop is not None and should_stop():
            return True                     # 已点击过，停止请求不再回滚，按已点击处理
        if _print_view_gone(window):
            time.sleep(0.4)
            if _print_view_gone(window):
                return True
            continue                      # 只是瞬时读不到，继续等
        time.sleep(0.15)
    return _print_view_gone(window)


def _background_print_status(app) -> int | None:
    """Word 的 Application.BackgroundPrintingStatus（后台打印任务数）；读不到返回 None。"""
    if app is None:
        return None
    try:
        return int(app.BackgroundPrintingStatus)
    except Exception:  # noqa: BLE001 - WPS/旧版 Word 没有该属性，或 Word 忙时拒绝调用
        return None


def _confirm_print_handoff(app, log: Callable[[str], None],
                           confirm_wait: float = PRINT_CONFIRM_WAIT,
                           drain_timeout: float = PRINT_DRAIN_TIMEOUT,
                           should_stop=None) -> str:
    """确认 Word 真的把作业交给了打印后台；返回 handoff / none。

    点击“打印”后 Word 会开一个后台打印任务。这里等它出现再等它落盘完成，
    确认之后才允许关闭文档——否则 Word 被提前结束会把作业一起带走。
    """
    if app is None:
        time.sleep(PRINT_CLOSE_GRACE)
        return "none"

    deadline = time.monotonic() + max(0.5, confirm_wait)
    status = None
    while time.monotonic() < deadline:
        status = _background_print_status(app)
        if status is None:
            time.sleep(0.1)          # Word 忙（刚点完打印），稍后重试而不是立刻放弃
            continue
        if status > 0:
            log(f"Word 已接收打印指令（后台打印任务 {status} 个），等待其交给打印后台…")
            drain_deadline = time.monotonic() + max(1.0, drain_timeout)
            while time.monotonic() < drain_deadline:
                status = _background_print_status(app)
                if status is None or status == 0:
                    log("Word 后台打印任务已全部交给打印后台，可以关闭文档")
                    return "handoff"
                if should_stop is not None and should_stop():
                    # 作业已经进了 Word 的后台打印队列，停止请求不再回滚，先等它落盘
                    log("收到停止请求：作业已在 Word 后台打印队列中，先等它交给打印后台再关闭文档")
                    should_stop = None      # 只提示一次，继续等落盘
                time.sleep(0.2)
            log(f"Word 后台打印任务在 {drain_timeout:.0f} 秒内未结束，按已提交处理")
            return "handoff"
        time.sleep(0.1)

    if status is None:
        log("读取 Word 后台打印状态失败（文档可能已被驱动瞬间接手），按视图关闭确认结果")
    else:
        log("Word 未报告后台打印任务（作业可能已被驱动瞬间接手），关闭文档后按队列确认")
    return "none"


def _open_document_visible(doc_path: Path, log: Callable[[str], None]):
    """用 Word（优先）或 WPS 以可见方式打开文档，返回 (关闭函数, 窗口标题, 文档对象, 应用对象, 进程号)。"""
    import pythoncom  # noqa: WPS433

    pythoncom.CoInitialize()
    app = None
    backend = ""
    errors: list[str] = []
    for backend_name, prog_ids in (
        ("Microsoft Word", ["Word.Application"]),
        ("WPS 文字", list(WPS_PROG_IDS)),
    ):
        if doc_path.suffix.lower() in {".wps", ".wpt"} and backend_name.startswith("Microsoft"):
            continue
        try:
            import win32com.client  # noqa: WPS433

            for prog_id in prog_ids:
                try:
                    app = win32com.client.DispatchEx(prog_id)
                    break
                except Exception as exc:  # noqa: WPS421
                    errors.append(f"{prog_id}: {exc}")
        except ImportError as exc:
            errors.append(str(exc))
        if app is not None:
            backend = backend_name
            break
    if app is None:
        raise UiPrintError("本机没有可打开该文档的 Word/WPS：" + "；".join(errors))

    log(f"已启动 {backend}，正在打开文档：{doc_path.name}")
    try:
        app.Visible = True
    except Exception:  # noqa: WPS420
        pass
    try:
        app.DisplayAlerts = 0
    except Exception:  # noqa: WPS420
        pass
    document = app.Documents.Open(str(doc_path), ReadOnly=True, AddToRecentFiles=False)

    # 记录我们创建的实例进程号：退出失败时只能清理自己启动的这一个
    pid = None
    try:
        pid = window_pid(int(app.ActiveWindow.Hwnd))
    except Exception:  # noqa: BLE001
        pid = None
    if not pid:
        try:
            pid = window_pid(int(document.ActiveWindow.Hwnd))
        except Exception:  # noqa: BLE001
            pid = None

    def close() -> None:
        try:
            document.Close(False)
        except Exception:  # noqa: WPS420
            pass
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
        # Quit 是异步的：文档还在渲染/打印时 Word 可能不响应，进程会残留。
        # 残留的 Word 会让下一次下发“文档窗口已开但读不到”，必须收干净。
        if pid:
            deadline = time.monotonic() + 3.0
            while time.monotonic() < deadline:
                if not process_alive(pid):
                    break
                time.sleep(0.2)
            if process_alive(pid):
                name = process_name(pid)
                if name.startswith(("winword", "wps")):
                    if terminate_pid(pid):
                        log(f"文档进程 {name}({pid}) 未响应关闭请求，已结束该实例")
                    else:
                        log(f"文档进程 {name}({pid}) 未能关闭，请手工关闭残留的 Word/WPS 窗口")
        pythoncom.CoUninitialize()

    return close, doc_path.stem, document, app, pid


def _default_opener(doc_path: Path, log: Callable[[str], None]):
    return _open_document_visible(doc_path, log)


# 键为 _norm() 规范化后的文本（去括号/空格/小写）：'横向(S)' → '横向s'
WD_ORIENT = {"纵向p": 0, "纵向": 0, "横向s": 1, "横向": 1, "portrait": 0, "landscape": 1}
WD_PAPER = {
    "a4:210x297mm": 7, "a3:297x420mm": 6, "a5:148x210mm": 11,
    "letter:8.5x11inch": 2, "legal:8.5x14inch": 4,
    "jisb4:257x364mm": 12, "jisb5:182x257mm": 13,
}
MM_TO_POINTS = 2.834645669


def _page_setup_key(paper: str) -> str:
    text = str(paper or "").strip().lower().replace(" ", "")
    if "(" in text and ")" in text:
        name, rest = text.split("(", 1)
        return f"{name}:{rest.rstrip(')')}"
    return text


def apply_wps_page_attrs(document, attrs: list[ExtendedAttr], log: Callable[[str], None]) -> list[str]:
    """WPS/Word 页面布局属性：通过 COM PageSetup 直接设置（比点界面稳）。

    返回无法应用（且影响版面）的属性列表；纯预览类属性只记录不阻塞。
    """
    unresolved: list[str] = []
    for attr in attrs:
        label = attr.label
        if "预览" in label:
            log(f"WPS 页面布局：{label}={attr.value} 为预览范围，不影响打印输出，已记录")
            continue
        try:
            if "画像方向" in label:
                orientation = WD_ORIENT.get(_norm(attr.value))
                if orientation is None:
                    unresolved.append(f"{label}={attr.value}")
                    continue
                document.PageSetup.Orientation = orientation
                log(f"WPS 页面布局：纸张方向 = {attr.value}")
            elif "自定义纸大小" in label:
                pair = split_pair(attr.value)
                if pair is None:
                    unresolved.append(f"{label}={attr.value}")
                    continue
                width_mm, height_mm = float(pair[0]), float(pair[1])
                document.PageSetup.PageWidth = width_mm * MM_TO_POINTS
                document.PageSetup.PageHeight = height_mm * MM_TO_POINTS
                log(f"WPS 页面布局：自定义纸张 = {width_mm}×{height_mm} mm")
            elif "纸张大小" in label:
                code = WD_PAPER.get(_page_setup_key(attr.value))
                if code is None:
                    unresolved.append(f"{label}={attr.value}")
                    continue
                document.PageSetup.PaperSize = code
                log(f"WPS 页面布局：纸张大小 = {attr.value}")
            else:
                log(f"WPS 页面布局：{label} 不在自动设置范围内，已记录")
        except Exception as exc:  # noqa: WPS421
            unresolved.append(f"{label}（{type(exc).__name__}: {exc}）")
    return unresolved


def _set_bool_control(container, label: str, desired: bool, log: Callable[[str], None], description: str = "") -> bool:
    label_text = base_label(label)
    for check in _descendants(container, control_type="CheckBox"):
        text = _text_of(check)
        if not text:
            continue
        normalized = _norm(text)
        if _norm(label_text) in normalized or (normalized and normalized in _norm(label) and len(normalized) >= 2):
            try:
                current = check.get_toggle_state()
            except Exception:  # noqa: WPS420
                current = None
            if current != (1 if desired else 0):
                try:
                    check.toggle()
                except Exception:  # noqa: WPS420
                    check.click_input()
            log(f"{description or '驱动属性'}：{label_text} = {'是' if desired else '否'}")
            return True
    return False


def _visible_control(control) -> bool:
    try:
        return bool(control.is_visible())
    except Exception:  # noqa: WPS420
        return True


def _combo_enabled(combo) -> bool:
    try:
        return bool(combo.is_enabled())
    except Exception:  # noqa: WPS420
        return True


def _combo_current_text(combo) -> str:
    """读下拉当前选中文本；自绘/不可读时返回空串。"""
    try:
        return str(combo.selected_text() or "").strip()
    except Exception:  # noqa: BLE001
        return ""


def _combo_item_wrappers(combo) -> list:
    """展开下拉并返回条目控件（读完保持展开，调用方决定点选还是收起）。

    这类驱动的下拉是自绘的：收起时 ``texts()`` 直接抛 COMError，
    条目只有展开后才以 ListItem 形式出现在可访问性树里。
    """
    for _ in range(3):
        wrappers = [
            item for item in _descendants(combo, control_type="ListItem")
            if _text_of(item).strip()
        ]
        if wrappers:
            return wrappers
        try:
            combo.click_input()
        except Exception:  # noqa: WPS420
            return []
        time.sleep(0.35)
    return [
        item for item in _descendants(combo, control_type="ListItem")
        if _text_of(item).strip()
    ]


def _combo_items(combo) -> list[str]:
    """读取下拉条目文本。

    注意：这类驱动收起状态下 ``texts()`` 返回的是子控件的文字
    （标签 + 当前值 + “关闭”按钮），并不是选项列表；所以必须先展开读
    ListItem，展开也读不到时才退回 ``texts()``。
    """
    wrappers = _combo_item_wrappers(combo)
    if wrappers:
        items = [_text_of(item).strip() for item in wrappers]
        try:
            combo.type_keys("{ESC}")     # 收起列表，避免挡住后面的控件
        except Exception:  # noqa: WPS420
            pass
        time.sleep(0.15)
        return [item for item in items if item]
    try:
        items = [str(item).strip() for item in (combo.texts() or [])]
    except Exception:  # noqa: BLE001 - 自绘下拉收起时必抛 COMError
        return []
    return [item for item in items if item]


def _select_combo_item(combo, wanted: str) -> bool:
    """按条目文本选中下拉项：展开 → 点同名 ListItem → 结束。

    「Home + 下移 N 次」这种键盘定位在这类驱动上会偏一格（实测选“纸盒1”落到
    “纸盒2”），因此优先直接点条目。
    """
    wrappers = _combo_item_wrappers(combo)
    if not wrappers:
        return False
    options = [_text_of(item).strip() for item in wrappers]
    match = _best_match(options, wanted)
    if match is None:
        try:
            combo.type_keys("{ESC}")
        except Exception:  # noqa: WPS420
            pass
        return False
    for wrapper in wrappers:
        if _text_of(wrapper).strip() == match:
            wrapper.click_input()
            time.sleep(0.3)
            return True
    return False


def _find_combo(container, label_text: str):
    """按标签找下拉框：先精确同名，再退化为包含匹配。

    精确优先很重要：纸张页上“分隔页”和“分隔页（OHP插页）纸张来源”同时存在，
    包含匹配会把前者错配到后者。
    """
    combos = [combo for combo in _descendants(container, control_type="ComboBox") if _visible_control(combo)]
    target = _norm(label_text)
    for combo in combos:
        if _norm(_label_before(combo)) == target or _norm(_text_of(combo)) == target:
            return combo
    for combo in combos:
        if _matches_labels(_label_before(combo), (label_text,)) or _matches_labels(_text_of(combo), (label_text,)):
            return combo
    return None


def _select_radio(container, words, log: Callable[[str], None] | None = None, description: str = "") -> bool:
    """按文字点单选按钮（驱动把“输出”做成 正常/分隔页/OHP插页 三个单选项）。

    已经是目标项就不再点（纸张来源为“自动选择”时整组灰化，点了也没用）；
    点完要回读选中状态，点不动就返回 False，不把“没生效”当成功。
    """
    labels = words if isinstance(words, (tuple, list)) else (words,)
    for radio in _descendants(container, control_type="RadioButton"):
        text = _text_of(radio)
        if not text.strip():
            continue
        if not _matches_labels(text, labels):
            continue
        try:
            if radio.is_selected():
                if log:
                    log(f"{description or '驱动属性'}：“{text.strip()}”已是选中项")
                return True
        except Exception:  # noqa: BLE001 - 个别控件读不到选中状态，按下面对待
            pass
        try:
            radio.click_input()
        except Exception:  # noqa: BLE001
            return False
        time.sleep(0.25)
        try:
            selected = bool(radio.is_selected())
        except Exception:  # noqa: BLE001 - 读不到状态时按已点击处理
            selected = True
        if selected and log:
            log(f"{description or '驱动属性'}：选择“{text.strip()}”")
        return selected
    return False


# 纸张页“输出”分组里的分隔页下拉；它必须先把“纸张来源”指到具体纸盒才可用
SEPARATOR_LABEL = "分隔页"
SEPARATOR_LABELS = ("分隔页", "分隔页（OHP插页）纸张来源")
SEPARATOR_SOURCE_LABELS = ("来源", "纸张来源")
# 模板里给用户看的属性名 → 驱动界面上的实际标签（真机 audit_driver_options 实测）：
# - 主纸张来源那行视觉标签是“来源”（“纸张来源”是分隔页下面那行，容易张冠李戴）；
# - “分隔页”下面那行只叫“纸张来源”，模板为了区分写成了“分隔页（OHP插页）纸张来源”。
LABEL_ALIASES = {
    "纸张来源": "来源",
    "分隔页（OHP插页）纸张来源": "纸张来源",
}


def _separator_prerequisite(container, label_text: str) -> str:
    """返回分隔页下拉不可用的原因（可用时返回空串），供报错提示复用。"""
    combo = _find_combo(container, label_text)
    if combo is None:
        return f"驱动属性页里找不到“{label_text}”下拉框"
    if _combo_enabled(combo):
        return ""
    source_combo = _find_combo(container, "来源")
    current = ""
    if source_combo is not None:
        current = _combo_current_text(source_combo)
    if current and "自动" in current:
        return (
            f"驱动把“{label_text}”灰化了：纸张来源是“{current}”时不能插分隔页。"
            "请把用例“纸盒”列改成具体纸盒（纸盒1/纸盒2），或在驱动扩展属性里加一行“纸张来源=纸盒1”。"
        )
    return (
        f"驱动把“{label_text}”灰化了（需在“输出”分组里先选中“分隔页”，"
        "并且纸张来源必须是具体纸盒）。"
    )


def _ensure_separator_enabled(container, label_text: str, log: Callable[[str], None]) -> None:
    """把“输出→分隔页”切过去，让分隔页下拉解禁；做不到就抛出可操作的错误。"""
    combo = _find_combo(container, label_text)
    if combo is not None and _combo_enabled(combo):
        return
    reason = _separator_prerequisite(container, label_text)
    if "自动" in reason:
        raise UiPrintError(reason)
    mode = "OHP插页" if "OHP" in label_text else "分隔页"
    if _select_radio(container, (mode,), log, "驱动属性"):
        time.sleep(0.4)
        combo = _find_combo(container, label_text)
        if combo is not None and _combo_enabled(combo):
            return
    raise UiPrintError(_separator_prerequisite(container, label_text) or reason)


def _set_choice_control(container, label: str, value: str, log: Callable[[str], None], description: str = "") -> bool:
    label_text = LABEL_ALIASES.get(base_label(label), base_label(label))
    if _norm(label_text) == _norm(SEPARATOR_LABEL):
        _ensure_separator_enabled(container, label_text, log)
    combo = _find_combo(container, label_text)
    if combo is None:
        return False
    options = _combo_items(combo)
    exact = next((item for item in options if _norm(item) == _norm(value)), None)
    match = exact or (_best_match(options, value) if options else None)
    if match is not None and exact is None and _norm(match) in _norm(value):
        # 用户写的是“作业之间和份数之间”，驱动只有“作业之间”：不能悄悄降级成一半
        raise UiPrintError(
            f"{label_text}：驱动下拉里没有“{value}”这一项"
            f"（可选：{'、'.join(options)}）"
        )
    if match is None:
        raise UiPrintError(
            f"{label_text}：驱动下拉里找不到“{value}”"
            f"（可选：{'、'.join(options) if options else '驱动未暴露条目'}）"
        )
    if not _combo_enabled(combo):
        raise UiPrintError(_separator_prerequisite(container, label_text) or f"{label_text}：控件当前灰化，无法选择。")
    selected = False
    try:
        combo.select(match)
        time.sleep(0.3)
        current = _combo_current_text(combo)
        # 读不到当前文本（自绘下拉）时按“已选中”处理，靠上层 DEVMODE 回读兜底
        selected = (not current) or _norm(current) == _norm(match)
    except Exception:  # noqa: BLE001 - 自绘下拉的 select 可能直接报错
        selected = False
    if not selected:
        selected = _select_combo_item(combo, match)
    if not selected:
        actual = _combo_current_text(combo) or "读取不到"
        raise UiPrintError(f"{label_text}：选择“{match}”后回读不一致（当前 {actual}）。")
    log(f"{description or '驱动属性'}：{label_text} = {match}")
    return True


def _set_text_control(container, label: str, value: str, log: Callable[[str], None], description: str = "") -> bool:
    label_text = base_label(label)
    for edit in _descendants(container, control_type="Edit"):
        if _matches_labels(_label_before(edit), (label_text,)):
            edit.set_edit_text(value)
            log(f"{description or '驱动属性'}：{label_text} = {value}")
            return True
    return False


def _select_driver_tab(driver_window, page: str, log: Callable[[str], None]) -> bool:
    keywords = page_keywords(page)
    for control_type in ("TabItem", "Button", "RadioButton"):
        for tab in _descendants(driver_window, control_type=control_type):
            if _matches_labels(_text_of(tab), keywords):
                tab.click_input()
                log(f"驱动属性：切换到“{_text_of(tab)}”页")
                return True
    return False


def apply_driver_attrs(driver_window, attrs: list[ExtendedAttr], log: Callable[[str], None]) -> list[str]:
    """在驱动属性窗口里按页面分组设置扩展属性；返回无法设置的属性列表。"""
    unresolved: list[str] = []
    grouped = group_by_page(attrs)
    for page, group in grouped.items():
        if page in {"wps_page", "print_dialog"}:
            continue
        _select_driver_tab(driver_window, page, log)
        for attr in group:
            spec = ATTRIBUTE_BY_LABEL.get(attr.label)
            kind = spec.kind if spec else ("bool" if bool_value(attr.value) is not None else "text")
            try:
                if kind == "bool":
                    desired = bool_value(attr.value)
                    if desired is None:
                        unresolved.append(f"{attr.label}={attr.value}（布尔值无法识别）")
                        continue
                    if not _set_bool_control(driver_window, attr.label, desired, log):
                        unresolved.append(f"{attr.label}={attr.value}（未找到勾选框）")
                elif kind == "choice":
                    if _norm(base_label(attr.label)) == _norm("输出模式"):
                        # 纸张页“输出”分组：正常/分隔页/OHP插页，是单选按钮不是下拉
                        if not _select_radio(driver_window, (attr.value,), log, "驱动属性"):
                            unresolved.append(f"输出模式={attr.value}（未找到该单选项）")
                    elif (_norm(base_label(attr.label)) == _norm(SEPARATOR_LABEL)
                          and _norm(attr.value) in {_norm("无"), _norm("关闭"), _norm("正常")}):
                        # 驱动没有“无”这一项：等于把“输出”切回“正常”
                        if _select_radio(driver_window, ("正常",), log, "驱动属性"):
                            log("驱动属性：分隔页 = 无（“输出”切回“正常”）")
                        else:
                            unresolved.append(f"{attr.label}={attr.value}（未找到“输出→正常”单选项）")
                    elif not _set_choice_control(driver_window, attr.label, attr.value, log):
                        unresolved.append(f"{attr.label}={attr.value}（未找到下拉框）")
                elif kind == "pair":
                    pair = split_pair(attr.value)
                    if pair is None:
                        unresolved.append(f"{attr.label}={attr.value}（应为 宽*高）")
                        continue
                    set_first = _set_text_control(driver_window, "宽", pair[0], log)
                    set_second = _set_text_control(driver_window, "长", pair[1], log)
                    if not (set_first and set_second):
                        unresolved.append(f"{attr.label}={attr.value}（未找到长/宽输入框）")
                elif kind == "combo_text":
                    parts = split_composite(attr.value)
                    applied = 0
                    for part in parts:
                        if _set_text_control(driver_window, f"{base_label(attr.label)}{applied + 1}", part, log):
                            applied += 1
                    if applied == 0 and not _set_text_control(driver_window, attr.label, attr.value, log):
                        unresolved.append(f"{attr.label}={attr.value}（未找到输入框）")
                else:  # int / text
                    if not _set_text_control(driver_window, attr.label, attr.value, log):
                        if not _set_choice_control(driver_window, attr.label, attr.value, log):
                            unresolved.append(f"{attr.label}={attr.value}（未找到输入框）")
            except UiPrintError as exc:
                unresolved.append(f"{attr.label}：{exc}")
            except Exception as exc:  # noqa: BLE001 - 自绘控件随时可能抛 COMError
                unresolved.append(f"{attr.label}={attr.value}（{type(exc).__name__}: {exc}）")
    return unresolved


def apply_print_dialog_attrs(dialog, attrs: list[ExtendedAttr], log: Callable[[str], None]) -> list[str]:
    """Ctrl+P 打印页上的扩展属性（页码范围由标准流程处理）。"""
    unresolved: list[str] = []
    for attr in attrs:
        if attr.page != "print_dialog":
            continue
        if "页码范围" in attr.label:
            log("打印对话框：页码范围已按用例“页码”列设置")
            continue
        desired = bool_value(attr.value)
        if desired is None:
            unresolved.append(f"{attr.label}={attr.value}")
            continue
        if not _set_bool_control(dialog, attr.label, desired, log, "打印对话框"):
            unresolved.append(f"{attr.label}={attr.value}（未找到勾选框）")
    return unresolved


def print_with_ui(
    printer: str,
    case: PrintCase,
    source_path: Path,
    log: Callable[[str], None],
    timeout: float = 90,
    desktop=None,
    opener: Callable[..., tuple[Callable[[], None], str]] | None = None,
    set_driver_attributes: bool = True,
    should_stop: Callable[[], bool] | None = None,
) -> tuple[str, str]:
    """通过 UI 自动化（pywinauto）下发打印。返回 (通道标识, 作业名主干)。

    ``set_driver_attributes=False`` 表示属性已由 driver_ui 在驱动“打印首选项”
    里设置并回读确认，这里只负责打开文档、选择驱动、填份数/页码、点“打印”。

    ``should_stop``：用户点“停止”时的判定函数。每个等待环节都会检查它，
    命中即抛 UiPrintInterrupted（不算失败，用例记为“已停止”），不会让界面
    卡在某个 90 秒的等待里。
    """
    source_path = Path(source_path)
    if source_path.suffix.lower() not in UI_PRINTABLE_EXTENSIONS:
        raise UiPrintError(
            f"{source_path.suffix} 没有可自动化的文档界面；请对 PDF/图片使用内置 GDI 直打模式。"
        )
    # 本线程先固定 COM 单元模式（Word COM 与 pywinauto 共用同一线程，必须一致）
    try:
        log(f"UI 线程 COM 单元模式：{ensure_sta()}")
        import_ui_engine()
    except Exception as exc:  # noqa: BLE001
        raise UiPrintError(f"UI 自动化引擎初始化失败：{exc}") from exc
    desktop = desktop or _default_desktop()
    opener = opener or _default_opener

    close_app = None
    try:
        if should_stop is not None and should_stop():
            raise UiPrintInterrupted("收到停止请求，尚未打开文档，本次不下发")
        opened = opener(source_path, log)
        close_app, title = opened[0], opened[1]
        document = opened[2] if len(opened) > 2 else None
        app = opened[3] if len(opened) > 3 else None
        word_pid = opened[4] if len(opened) > 4 else None

        # 扩展属性：WPS 页面布局（必须在调出打印对话框前设置，影响分页）
        wps_attrs = [attr for attr in (case.extended_attrs or []) if attr.page == "wps_page"]
        if wps_attrs:
            if document is None:
                raise UiPrintError("用例包含 WPS 页面布局属性，但未能取得文档对象，无法设置。")
            wps_unresolved = apply_wps_page_attrs(document, wps_attrs, log)
            if wps_unresolved:
                raise UiPrintError("WPS 页面布局属性无法设置：" + "、".join(wps_unresolved))

        window = _find_window(
            desktop, title, timeout, should_stop=should_stop,
            app=app, prefer_pid=word_pid, log=log,
        )
        log(f"文档窗口已就绪：{_text_of(window) or title}")

        container, backstage = _open_print_view(desktop, window, log, timeout, should_stop)

        log("打印界面已就绪（%s），开始选择打印机驱动" % ("后台打印视图" if backstage else "打印对话框"))
        if backstage:
            _select_printer_in_backstage(window, printer, log)
        else:
            _select_printer(container, printer, log)

        _set_copies(container, case.copies, log)
        _set_page_range(container, case.pages, log)

        unresolved = apply_print_dialog_attrs(container, case.extended_attrs or [], log)

        if set_driver_attributes:
            try:
                driver_window = _open_preferences(container, desktop, log, timeout)
            except UiPrintError as pref_error:
                if case.extended_attrs or case.borderless == "yes":
                    raise
                log(f"未找到驱动“首选项/属性”入口（{pref_error}）；"
                    "标准属性已由驱动属性页设置并回读确认，继续执行打印。")
                driver_window = None
            if driver_window is not None:
                unresolved += _apply_case_attributes(driver_window, case, log)
                unresolved += apply_driver_attrs(driver_window, case.extended_attrs or [], log)
                _confirm(driver_window, log)
                log("驱动属性已确认，返回打印对话框")

        if unresolved:
            raise UiPrintError(
                "以下显式属性无法在驱动界面设置，为避免属性失效本次不下发："
                + "、".join(unresolved)
            )

        clicked_text = None
        for attempt in range(1, PRINT_CLICK_ATTEMPTS + 1):
            clicked_text = _click_print_button(container, log, should_stop=should_stop)
            if clicked_text is None:
                if backstage and _print_view_gone(window):
                    # 按钮已经随视图消失：说明上一轮点击其实生效了
                    clicked_text = "打印"
                    log("打印视图已自行关闭（上一轮点击已生效）")
                    break
                raise UiPrintError(
                    "打印界面里没有可点击的“打印”按钮（可能仍在渲染或已被禁用），无法确认作业已提交。"
                )
            if not backstage:
                log(f"已点击“{clicked_text}”，等待打印对话框关闭")
                break
            log(f"已点击“{clicked_text}”（第 {attempt} 次），等待打印视图关闭…")
            if _left_print_view(window, PRINT_START_TIMEOUT, should_stop):
                log("打印视图已关闭，Word 已接受打印指令")
                break
            if attempt < PRINT_CLICK_ATTEMPTS:
                log(f"点击后 {PRINT_START_TIMEOUT:.0f} 秒内打印视图仍未关闭，重新点击“打印”")
        else:
            raise UiPrintError(
                f"已点击“{clicked_text}”但打印视图始终未关闭，Word 没有接受打印指令；"
                "为避免“日志显示已提交、实际没有出纸”，本次不下发，请重试。"
            )

        if backstage:
            outcome = _confirm_print_handoff(app, log, should_stop=should_stop)
        else:
            outcome = "dialog"

        if not backstage:
            _wait(lambda: not _is_visible(container), 20, "打印对话框在 20 秒内未关闭，无法确认作业已提交。")
        log(f"打印作业已提交给打印端（确认方式：{_HANDOFF_LABELS.get(outcome, outcome)}）")
        return "ui-automation", source_path.stem
    finally:
        if close_app is not None:
            try:
                close_app()
                log("已关闭文档并退出打印组件")
            except Exception:  # noqa: WPS420
                pass


def _is_visible(window) -> bool:
    try:
        return bool(window.is_visible())
    except Exception:  # noqa: WPS420
        return False


def ui_engine_available() -> tuple[bool, str]:
    """自检：在独立线程里导入 pywinauto 并创建一次 UIA 入口。

    不在此线程直接 import —— 主线程通常已被 Word/WPS 的 COM 初始化成 STA，
    comtypes 默认要 MTA，会抛 WinError -2147417850（线程模式无法更改）。
    """
    return ui_ready()
