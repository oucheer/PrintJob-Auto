"""用例勾选与表头“全选”复选框的回归测试（不依赖 Tk 显示环境）。"""
from __future__ import annotations

from printcase.app import PrintCaseApp


class _FakeVar:
    def __init__(self, value: bool = False) -> None:
        self.value = value

    def get(self) -> bool:
        return self.value

    def set(self, value: bool) -> None:
        self.value = value


class _FakeLabel:
    def __init__(self) -> None:
        self.text = ""

    def configure(self, text: str = "") -> None:
        self.text = text

    def cget(self, key: str) -> str:
        return self.text


class _FakeApp(PrintCaseApp):
    """只装配勾选相关状态，复用真实方法，避免创建 Tk 窗口。"""

    def __init__(self, total: int) -> None:
        self.cases = [object()] * total
        self.checked: set[int] = set()
        self.select_all_var = _FakeVar()
        self.case_count_label = _FakeLabel()

    def _render_cases(self) -> None:  # noqa: D102
        self._sync_select_all()


def test_select_all_checkbox_selects_and_clears():
    app = _FakeApp(5)
    app._render_cases()
    assert app.select_all_var.get() is False

    app.select_all_var.set(True)
    app._toggle_all_from_checkbox()
    assert app.checked == {0, 1, 2, 3, 4}
    assert app.select_all_var.get() is True
    assert "已选 5 条" in app.case_count_label.cget("text")

    app.select_all_var.set(False)
    app._toggle_all_from_checkbox()
    assert app.checked == set()
    assert app.select_all_var.get() is False
    assert "已选 0 条" in app.case_count_label.cget("text")


def test_partial_selection_leaves_select_all_unchecked():
    app = _FakeApp(4)
    app.checked = {1, 3}
    app._sync_select_all()
    assert app.select_all_var.get() is False
    assert "4 条" in app.case_count_label.cget("text")
    assert "已选 2 条" in app.case_count_label.cget("text")


def test_header_click_toggles_all_and_syncs_checkbox():
    app = _FakeApp(3)
    app._toggle_all_checks()  # 点击“选”列表头 → 全选
    assert app.checked == {0, 1, 2}
    assert app.select_all_var.get() is True
    app._toggle_all_checks()  # 再点 → 全部取消
    assert app.checked == set()
    assert app.select_all_var.get() is False


def test_empty_case_list_never_marks_select_all():
    app = _FakeApp(0)
    app._sync_select_all()
    assert app.select_all_var.get() is False
