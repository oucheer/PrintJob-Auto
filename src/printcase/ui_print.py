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


def _find_window(desktop, title: str, timeout: float):
    """按标题找文档窗口（标题包含文件名主干）。"""
    def search():
        for window in desktop.windows():
            if _norm(title) and _norm(title) in _norm(_text_of(window)):
                return window
        return None

    return _wait(search, timeout, f"未找到文档窗口：{title}（请确认文档已用 Word/WPS 打开）")


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


def _find_print_dialog_or_backstage(desktop, window, timeout: float):
    """返回 (容器, 是否后台视图)。

    Word 2013+ 的 Ctrl+P 打开的是“打印”后台视图（嵌在主窗口内），
    而不是独立对话框；WPS/旧版 Word 才是 #32770 对话框。两种都要支持。
    """
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        for candidate in desktop.windows():
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
        time.sleep(0.5)
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


def _open_print_view(desktop, window, log: Callable[[str], None], timeout: float):
    """反复尝试 Ctrl+P 直到打印界面出现；返回 (容器, 是否后台视图)。

    单次尝试只等一小段，失败就重新聚焦再发一次 Ctrl+P，避免把一次被吞掉的
    按键变成 90 秒空等 + 用例失败（批量下发时首条用例最容易踩到）。
    """
    deadline = time.monotonic() + max(1.0, timeout)
    attempts = 0
    last_error = ""
    while attempts < PRINT_VIEW_MAX_ATTEMPTS and time.monotonic() < deadline:
        attempts += 1
        focused = _ensure_document_focus(window, log)
        try:
            window.type_keys("^p")
        except Exception as exc:  # noqa: BLE001
            last_error = f"发送 Ctrl+P 失败：{type(exc).__name__}: {exc}"
        wait_for = min(PRINT_VIEW_WAIT_PER_TRY, max(1.0, deadline - time.monotonic()))
        log(f"已发送 Ctrl+P（第 {attempts} 次{f'，已聚焦正文' if focused else ''}），等待打印界面…")
        try:
            return _find_print_dialog_or_backstage(desktop, window, wait_for)
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
                        for window in desktop.windows()
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
                        wait: float = PRINT_BUTTON_READY_TIMEOUT) -> str | None:
    """点击“打印”按钮，返回实际点到的文本；找不到可点击的按钮返回 None。

    后台打印视图刚切换过来时，“打印”按钮可能还是禁用状态（预览仍在渲染），
    此时 click_input 不报错但也不生效——旧实现直接把这种情况记成“已提交”，
    于是出现“日志 3 条、实际只出 2 份”。这里先等按钮真正可点击再点。
    """
    deadline = time.monotonic() + max(0.0, wait)
    while True:
        for button, text in _print_buttons(container):
            if not _button_enabled(button):
                continue
            button.click_input()
            return text
        if time.monotonic() >= deadline:
            return None
        time.sleep(0.2)


def _print_view_gone(window) -> bool:
    """后台打印视图是否已关闭（读 UIA 失败时按“还在视图里”处理，避免误判）。"""
    try:
        return _backstage_printer_combo(window) is None
    except Exception:  # noqa: BLE001 - Word 忙时 UIA 会抛 COMError
        return False


def _left_print_view(window, timeout: float) -> bool:
    """等 Word 后台打印视图关闭（回到文档）。

    点击“打印”后视图一定会关闭；视图不动说明这次点击没有被接受，
    此时把作业记成“已提交”就会丢单。判定“已关闭”后会再确认一次，
    避免视图切换瞬间读不到控件被误判成“已关闭”。
    """
    deadline = time.monotonic() + max(0.5, float(timeout))
    while time.monotonic() < deadline:
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
                           drain_timeout: float = PRINT_DRAIN_TIMEOUT) -> str:
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
    """用 Word（优先）或 WPS 以可见方式打开文档，返回 (关闭函数, 窗口标题, 文档对象, 应用对象)。"""
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

    def close() -> None:
        try:
            document.Close(False)
        except Exception:  # noqa: WPS420
            pass
        try:
            app.Quit()
        except Exception:  # noqa: WPS420
            pass
        pythoncom.CoUninitialize()

    return close, doc_path.stem, document, app


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


def _set_choice_control(container, label: str, value: str, log: Callable[[str], None], description: str = "") -> bool:
    label_text = base_label(label)
    for combo in _descendants(container, control_type="ComboBox"):
        if not _matches_labels(_label_before(combo), (label_text,)) and not _matches_labels(_text_of(combo), (label_text,)):
            continue
        options = [str(item) for item in (getattr(combo, "texts", lambda: [])() or [])]
        match = _best_match(options, value)
        if match is None:
            raise UiPrintError(f"{label_text}：驱动下拉里找不到“{value}”（可选：{'、'.join(options[:6]) or '无'}）")
        combo.select(match)
        log(f"{description or '驱动属性'}：{label_text} = {match}")
        return True
    return False


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
                    if not _set_choice_control(driver_window, attr.label, attr.value, log):
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
) -> tuple[str, str]:
    """通过 UI 自动化（pywinauto）下发打印。返回 (通道标识, 作业名主干)。

    ``set_driver_attributes=False`` 表示属性已由 driver_ui 在驱动“打印首选项”
    里设置并回读确认，这里只负责打开文档、选择驱动、填份数/页码、点“打印”。
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
        opened = opener(source_path, log)
        close_app, title = opened[0], opened[1]
        document = opened[2] if len(opened) > 2 else None
        app = opened[3] if len(opened) > 3 else None

        # 扩展属性：WPS 页面布局（必须在调出打印对话框前设置，影响分页）
        wps_attrs = [attr for attr in (case.extended_attrs or []) if attr.page == "wps_page"]
        if wps_attrs:
            if document is None:
                raise UiPrintError("用例包含 WPS 页面布局属性，但未能取得文档对象，无法设置。")
            wps_unresolved = apply_wps_page_attrs(document, wps_attrs, log)
            if wps_unresolved:
                raise UiPrintError("WPS 页面布局属性无法设置：" + "、".join(wps_unresolved))

        window = _find_window(desktop, title, timeout)
        log(f"文档窗口已就绪：{_text_of(window) or title}")

        container, backstage = _open_print_view(desktop, window, log, timeout)
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
            clicked_text = _click_print_button(container, log)
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
            if _left_print_view(window, PRINT_START_TIMEOUT):
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
            outcome = _confirm_print_handoff(app, log)
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
