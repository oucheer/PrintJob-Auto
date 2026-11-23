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

import platform
import subprocess
import time
from dataclasses import dataclass
from typing import Any, Callable

from .models import PrintCase
from .ui_print import _best_match, _descendants, _matches_labels, _norm, _text_of, _wait
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
            actions.append(Action(TAB_PAPER, "纸张尺寸", "combo_index", paper_index, "PaperSize", expected))
        else:
            actions.append(Action(TAB_PAPER, "纸张尺寸", "combo_text", case.paper, "PaperSize", expected))

    media = str(case.media_type or "").strip()
    if media:
        # 介质代号由驱动私有编码决定，没有可预期的标准值：只校验“确实变了且非零”
        media_index = _index_in_names(media, caps.get("media_names"))
        if media_index is not None:
            actions.append(Action(TAB_PAPER, "纸张类型", "combo_index", media_index, "MediaType"))
        else:
            actions.append(Action(TAB_PAPER, "纸张类型", "combo_text", media, "MediaType"))

    bin_value = str(case.bin or "").strip()
    if bin_value:
        bin_index = _index_in_names(bin_value, caps.get("bin_names")) if not bin_value.isdigit() else None
        bin_ids = caps.get("bin_ids") or []
        expected_bin = bin_ids[bin_index] if bin_index is not None and bin_index < len(bin_ids) else (int(bin_value) if bin_value.isdigit() else None)
        if bin_index is not None:
            actions.append(Action(TAB_PAPER, "纸张来源", "combo_index", bin_index, "DefaultSource", expected_bin))
        else:
            actions.append(Action(TAB_PAPER, "纸张来源", "combo_text", bin_value, "DefaultSource", expected_bin))

    # 份数不在驱动属性页里设置：Word 打印对话框按“份数”提交，PDF/图片由内置渲染器
    # 按页重复展开；两处都设会让驱动与上层各乘一次，导致份数翻倍。

    if case.duplex and case.duplex != "simplex":
        actions.append(Action(TAB_BASIC, "双面", "radio", DUPLEX_LABELS.get(case.duplex, (case.duplex,)), "Duplex", DUPLEX_CODES.get(case.duplex)))
    elif case.duplex == "simplex":
        actions.append(Action(TAB_BASIC, "双面", "radio", DUPLEX_LABELS["simplex"], "Duplex", 1))

    if case.color:
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


def open_preferences(printer: str, timeout: float = 25, desktop=None, log=None):
    """打开驱动“打印首选项”窗口（rundll32 printui /e，等价于手工点开首选项）。"""
    desktop = desktop or _desktop()
    subprocess.Popen(
        ["rundll32.exe", "printui.dll,PrintUIEntry", "/e", "/n", printer],
        close_fds=True,
    )

    def search():
        for window in desktop.windows():
            text = _text_of(window)
            if _norm(printer) and _norm(printer) in _norm(text) and any(
                word in text for word in ("首选项", "属性", "Preferences", "Properties")
            ):
                return window
        return None

    window = _wait(search, timeout, f"驱动“打印首选项”窗口未出现（{printer}）；请确认驱动已安装。")
    if log:
        log(f"已打开驱动属性页：{_text_of(window) or printer}")
    return window


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


def _select_combo_text(combo, wanted: str) -> bool:
    options = [str(item) for item in (getattr(combo, "texts", lambda: [])() or [])]
    match = _best_match(options, wanted)
    if match is None:
        return False
    combo.select(match)
    return True


def _apply_action(window, action: Action, log=None) -> None:
    if action.kind in {"combo_index", "combo_text"}:
        combos = _visible_combos(window)
        position = PAPER_TAB_COMBOS.index(action.label) if action.label in PAPER_TAB_COMBOS else -1
        if position >= 0 and position < len(combos):
            combo = combos[position]
        else:
            combo = None
            for candidate in combos:
                if _matches_labels(_text_of(candidate), (action.label,)):
                    combo = candidate
                    break
            if combo is None:
                raise DriverUiError(f"驱动属性页里找不到“{action.label}”下拉框。")
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
        return

    if action.kind == "radio":
        words = action.value if isinstance(action.value, (tuple, list)) else (action.value,)
        for radio in _descendants(window, control_type="RadioButton"):
            if _matches_labels(_text_of(radio), words):
                radio.click_input()
                time.sleep(0.15)
                if log:
                    log(f"驱动属性页：{action.label} → {_text_of(radio)}")
                return
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
                return
        # 开关型下拉（如“跳过空白页”：关闭/开启）
        for combo in _visible_combos(window):
            if _matches_labels(_text_of(combo), words):
                if not _select_combo_text(combo, "开启" if action.value else "关闭"):
                    raise DriverUiError(f"“{action.label}”下拉里找不到开启/关闭选项。")
                if log:
                    log(f"驱动属性页：{action.describe()}")
                return
        raise DriverUiError(f"驱动属性页里找不到“{action.label}”勾选框。")

    if action.kind == "int_text":
        for edit in _descendants(window, control_type="Edit"):
            if _matches_labels(_text_of(edit), (action.label,)):
                edit.set_edit_text(str(action.value))
                if log:
                    log(f"驱动属性页：{action.describe()}")
                return
        # 驱动常把输入框的标签错位挂到前一个控件上，退化为按控件名前缀匹配
        for edit in _descendants(window, control_type="Edit"):
            if _norm(action.label) in _norm(_text_of(edit)):
                edit.set_edit_text(str(action.value))
                if log:
                    log(f"驱动属性页：{action.describe()}")
                return
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


def _close(window) -> None:
    for button in _descendants(window, control_type="Button"):
        if _norm(_text_of(button)) == _norm("取消"):
            try:
                button.click_input()
            except Exception:  # noqa: WPS420
                pass
            return
    try:
        window.close()
    except Exception:  # noqa: WPS420
        pass


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
    """只读读取驱动当前默认 DEVMODE（用于校验与还原）。"""
    w = w or _win32print()
    handle = w.OpenPrinter(printer, {"DesiredAccess": w.PRINTER_ALL_ACCESS})
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
    """把驱动默认 DEVMODE 还原成执行前快照（只写回原值，不参与属性设置）。"""
    w = w or _win32print()
    handle = w.OpenPrinter(printer, {"DesiredAccess": w.PRINTER_ALL_ACCESS})
    try:
        info = w.GetPrinter(handle, 2)
        info["pDevMode"] = before["devmode"]
        w.SetPrinter(handle, 2, info, 0)
    finally:
        w.ClosePrinter(handle)
    if log:
        log("驱动默认设置已还原为用例执行前的快照。")


def verify(printer: str, actions: list[Action], before: dict, w=None) -> list[str]:
    """回读默认 DEVMODE，核对每条可校验的操作是否真的生效。"""
    after = snapshot(printer, w=w)
    problems: list[str] = []
    for action in actions:
        if not action.expect_field:
            continue
        actual = after.get(action.expect_field)
        expected = action.expect_value
        if expected is None:
            # 无固定期望值（如驱动私有介质编码）：只要求相对快照发生变化且为正值
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
    try:
        window = open_preferences(printer, desktop=desktop, log=log)
        pages = {action.page for action in actions}
        for page in sorted(pages, key=lambda name: [TAB_BASIC, TAB_PAPER, TAB_LAYOUT].index(name) if name in {TAB_BASIC, TAB_PAPER, TAB_LAYOUT} else 9):
            if not _switch_tab(window, page, log):
                missing = [action.label for action in actions if action.page == page]
                raise DriverUiError(f"驱动属性页没有“{page}”页，无法设置：{'、'.join(missing)}")
            for action in [item for item in actions if item.page == page]:
                _apply_action(window, action, log)
        if case.extended_attrs:
            unresolved = _apply_extended(window, case.extended_attrs, log)
            if unresolved:
                raise DriverUiError("以下驱动扩展属性无法在驱动界面设置：" + "、".join(unresolved))
        _confirm(window, log)
        window = None
        time.sleep(0.6)
    except DriverUiError:
        if window is not None:
            _close(window)
        raise
    except Exception as exc:  # noqa: BLE001
        if window is not None:
            _close(window)
        raise DriverUiError(f"驱动属性页操作失败：{type(exc).__name__}: {exc}") from exc

    problems = verify(printer, actions, before, w=w)
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
