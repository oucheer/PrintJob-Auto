"""纸张页“分隔页”设置：自绘下拉 + 输出模式联动（真实故障的替身复现）。

真实驱动（Pantum M9705DN）上实测到的三个坑：
1. 下拉收起时 ``texts()`` 返回的是“标签/当前值/关闭”三个子控件文字，不是选项；
   条目只有展开后才是 ListItem。
2. 纸张来源是“自动选择”时，“输出→分隔页”单选灰化，下面的分隔页下拉也灰化；
   必须先指定具体纸盒，再点单选按钮。
3. 下拉只有“份数之间 / 作业之间”两项，没有“无”。
"""
from __future__ import annotations

import pytest

from printcase.ui_print import UiPrintError, _combo_items, _set_choice_control


class FakeListItem:
    control_type = "ListItem"

    def __init__(self, text: str, owner: "PaintedCombo"):
        self.text = text
        self.owner = owner
        self.clicks = 0

    def window_text(self):
        return self.text

    def click_input(self):
        self.clicks += 1
        self.owner.selected = self.text
        self.owner.expanded = False


class FakeRadio:
    control_type = "RadioButton"

    def __init__(self, text: str, parent, on_click=None):
        self.text = text
        self._parent = parent
        self.clicks = 0
        self._on_click = on_click
        parent.children.append(self)

    def window_text(self):
        return self.text

    def parent(self):
        return self._parent

    def click_input(self):
        self.clicks += 1
        if self._on_click:
            self._on_click()

    def descendants(self, **_kwargs):
        return []


class FakeLabel:
    control_type = "Text"

    def __init__(self, text: str, parent):
        self.text = text
        self._parent = parent
        parent.children.append(self)

    def window_text(self):
        return self.text

    def descendants(self, **_kwargs):
        return []


class PaintedCombo:
    """自绘下拉：收起时 texts() 只给出标签/当前值/关闭，展开后才有条目。"""

    control_type = "ComboBox"

    def __init__(self, label: str, options, parent, selected=None, enabled=True):
        self.text = label
        self.label = label
        self.options = list(options)
        self.selected = selected if selected is not None else self.options[0]
        self.enabled = enabled
        self.expanded = False
        self._items: list[FakeListItem] = []
        self._parent = parent
        self.children = []
        parent.children.append(self)

    # --- UIA 接口 ---
    def window_text(self):
        return self.text

    def parent(self):
        return self._parent

    def is_visible(self):
        return True

    def is_enabled(self):
        return self.enabled

    def selected_text(self):
        return self.selected

    def texts(self):
        return [self.label, self.selected, "关闭"]

    def click_input(self):
        self.expanded = not self.expanded
        if self.expanded and not self._items:
            self._items = [FakeListItem(option, self) for option in self.options]

    def type_keys(self, keys):
        if keys == "{ESC}":
            self.expanded = False

    def select(self, value):
        raise RuntimeError("自绘下拉不支持 select()")

    def descendants(self, control_type=None, **_kwargs):
        if self.expanded and control_type in (None, "ListItem"):
            return list(self._items)
        return []


class FakeWindow:
    def __init__(self):
        self.children = []
        self.text = "打印首选项"

    def window_text(self):
        return self.text

    def descendants(self, control_type=None, **_kwargs):
        found = []
        for child in self.children:
            if control_type is None or child.control_type == control_type:
                found.append(child)
        return found


def build_paper_window(source="纸盒1", separator_enabled=False, separator_value="份数之间"):
    """按真实纸张页的视觉顺序摆放控件（纸张页前三个下拉按位置取，不能乱序）。"""
    window = FakeWindow()
    FakeLabel("尺寸及纸型", window)
    PaintedCombo("尺寸及纸型", ["A4 (210 x 297mm)", "A3 (297 x 420mm)"], window)
    FakeLabel("纸张尺寸", window)
    PaintedCombo("纸张尺寸", ["普通纸(60-90gsm)", "厚纸1(106-120gsm)"], window)
    FakeLabel("来源", window)
    PaintedCombo("来源", ["自动选择", "多功能进纸器", "纸盒1", "纸盒2"], window, selected=source)

    FakeLabel("输出", window)
    separator_combo = PaintedCombo("分隔页", ["份数之间", "作业之间"], window,
                                   selected=separator_value, enabled=separator_enabled)
    FakeLabel("纸张来源", window)

    def enable_separator():
        separator_combo.enabled = True

    FakeRadio("正常", window)
    FakeRadio("分隔页", window, on_click=enable_separator)
    FakeRadio("OHP插页", window)
    return window, separator_combo


def test_combo_items_reads_listitems_instead_of_child_texts():
    window, separator = build_paper_window(separator_enabled=True)
    assert separator.texts() == ["分隔页", "份数之间", "关闭"]      # 收起时的假象
    assert _combo_items(separator) == ["份数之间", "作业之间"]
    assert separator.expanded is False                                # 读完要收起


def test_separator_enabled_after_clicking_output_radio():
    window, separator = build_paper_window(source="纸盒1", separator_enabled=False)
    logs: list[str] = []
    assert _set_choice_control(window, "分隔页", "作业之间", logs.append) is True
    assert separator.selected == "作业之间"
    assert separator.enabled is True
    assert any("分隔页" in line and "作业之间" in line for line in logs)


def test_separator_reports_auto_tray_requirement():
    """纸张来源=自动选择时驱动把分隔页灰化：必须报“要指定纸盒”，不能只抛 COMError。"""
    window, separator = build_paper_window(source="自动选择", separator_enabled=False)
    with pytest.raises(UiPrintError) as excinfo:
        _set_choice_control(window, "分隔页", "作业之间", lambda _t: None)
    message = str(excinfo.value)
    assert "纸盒" in message
    assert "自动" in message
    assert separator.selected == "份数之间"        # 没有乱改


def test_separator_rejects_option_that_driver_does_not_have():
    window, separator = build_paper_window(separator_enabled=True)
    with pytest.raises(UiPrintError) as excinfo:
        _set_choice_control(window, "分隔页", "作业之间和份数之间", lambda _t: None)
    assert "份数之间、作业之间" in str(excinfo.value)


# ---------- 纸张页下拉：序号定位不能偏一格 ----------

def test_select_combo_index_clicks_item_not_keyboard_offset():
    """实测故障：Home+下移 2 次选“纸盒1”会落到“纸盒2”，改成直接点条目。"""
    from printcase.driver_ui import _select_combo_index

    window, separator = build_paper_window(source="自动选择")
    source = [c for c in window.children if getattr(c, "label", "") == "来源"][0]
    _select_combo_index(source, 2)                 # 第 3 项 = 纸盒1
    assert source.selected == "纸盒1"
    assert [item.clicks for item in source._items] == [0, 0, 1, 0]


def test_apply_action_verifies_target_text_for_index_action():
    """按序号设置后回读不一致要报错，不能把纸盒设错还当成功。"""
    from printcase.driver_ui import Action, DriverUiError, _apply_action

    window, _separator = build_paper_window(source="自动选择")
    source = [c for c in window.children if getattr(c, "label", "") == "来源"][0]
    action = Action("纸张", "纸张来源", "combo_index", 1, "DefaultSource", 259, "纸盒9")
    with pytest.raises(DriverUiError):
        _apply_action(window, action)               # 目标项不存在 → 回读不一致
    source.selected = "自动选择"
    action = Action("纸张", "纸张来源", "combo_index", 2, "DefaultSource", 259, "纸盒1")
    assert _apply_action(window, action) is True
    assert source.selected == "纸盒1"


# ---------- 用例“分隔页”列 → 驱动“纸张→输出” ----------

def test_separator_attrs_map_to_driver_output_group(tmp_path):
    """分隔页=作业之间/份数之间 落到驱动“输出模式 + 分隔页 + 分隔页纸张来源”。"""
    from printcase.driver_ui import separator_attrs
    from printcase.models import PrintCase

    case = PrintCase(case_id="TC-SEP", enabled=True, file_path="a.docx",
                     separator_page="job", bin="纸盒1", separator_bin="纸盒2")
    attrs = [(item.label, item.value) for item in separator_attrs(case)]
    assert attrs == [
        ("输出模式", "分隔页"),
        ("分隔页", "作业之间"),
        ("分隔页（OHP插页）纸张来源", "纸盒2"),
    ]
    # 份数之间 同理；不填分隔页纸张来源时不覆盖驱动里的那一项
    case.separator_page = "copies"
    case.separator_bin = ""
    assert [(i.label, i.value) for i in separator_attrs(case)] == [
        ("输出模式", "分隔页"), ("分隔页", "份数之间"),
    ]
    # 无 → 显式把“输出”切回正常（驱动默认值可能被上一轮留在“分隔页”）
    case.separator_page = "none"
    assert [(i.label, i.value) for i in separator_attrs(case)] == [("输出模式", "正常")]


def test_separator_attrs_require_concrete_tray():
    """纸张来源=自动选择时驱动把整组灰化，必须提前报错而不是打开窗口才失败。"""
    from printcase.driver_ui import DriverUiError, separator_attrs
    from printcase.models import PrintCase

    case = PrintCase(case_id="TC-SEP", enabled=True, file_path="a.docx", separator_page="job")
    with pytest.raises(DriverUiError) as info:
        separator_attrs(case)
    assert "具体纸盒" in str(info.value)


def test_separator_source_attr_uses_driver_label():
    """模板里的“分隔页（OHP插页）纸张来源”在驱动里那行只叫“纸张来源”。"""
    from printcase.ui_print import LABEL_ALIASES

    assert LABEL_ALIASES["分隔页（OHP插页）纸张来源"] == "纸张来源"
