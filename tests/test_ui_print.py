"""pywinauto UI 自动化打印通道测试（用替身模拟文档窗口 / 打印对话框 / 驱动属性窗口）。"""
from pathlib import Path

import pytest

from printcase.models import PrintCase
from printcase.ui_print import UiPrintError, print_with_ui


class FakeText:
    control_type = "Text"
    children: list = []

    def __init__(self, text: str):
        self.text = text

    def window_text(self) -> str:
        return self.text


class FakeContainer:
    def __init__(self, children=None, text: str = "", control_type: str = ""):
        self.children = list(children or [])
        self.text = text
        self.control_type = control_type
        for child in self.children:
            child._parent = self

    def children_list(self):
        return list(self.children)

    def parent(self):
        return getattr(self, "_parent", None)

    def descendants(self, control_type=None, **_kwargs):
        found = []
        for child in self.children:
            if not control_type or child.control_type == control_type:
                found.append(child)
            if child.children:
                found.extend(child.descendants(control_type=control_type))
        return found


class FakeCombo(FakeContainer):
    def __init__(self, label: str, options: list[str], parent=None):
        super().__init__(text=label, control_type="ComboBox")
        self.options = list(options)
        self.selected = None
        self.edit = FakeEdit("", self)
        if parent is not None:
            parent.children.append(self)
            self._parent = parent

    def texts(self) -> list[str]:
        return list(self.options)

    def select(self, value: str) -> None:
        if value not in self.options:
            raise ValueError(value)
        self.selected = value

    def window_text(self) -> str:
        return self.selected or ""


class FakeEdit(FakeContainer):
    def __init__(self, value: str = "", parent=None):
        super().__init__(text="", control_type="Edit")
        self.value = value
        if parent is not None:
            parent.children.append(self)
            self._parent = parent

    def set_edit_text(self, text: str) -> None:
        self.value = text

    def window_text(self) -> str:
        return self.value


class FakeButton(FakeContainer):
    def __init__(self, label: str, on_click=None, parent=None):
        super().__init__(text=label, control_type="Button")
        self.on_click = on_click
        if parent is not None:
            parent.children.append(self)
            self._parent = parent

    def click_input(self) -> None:
        if self.on_click:
            self.on_click()

    def window_text(self) -> str:
        return self.text


class FakeCheck(FakeContainer):
    def __init__(self, label: str, checked: int = 0, parent=None):
        super().__init__(text=label, control_type="CheckBox")
        self.checked = checked
        if parent is not None:
            parent.children.append(self)
            self._parent = parent

    def get_toggle_state(self) -> int:
        return self.checked

    def toggle(self) -> None:
        self.checked = 1 - self.checked

    def click_input(self) -> None:
        self.toggle()

    def window_text(self) -> str:
        return self.text


class FakeRadio(FakeContainer):
    def __init__(self, label: str, parent=None):
        super().__init__(text=label, control_type="RadioButton")
        if parent is not None:
            parent.children.append(self)
            self._parent = parent

    def click_input(self) -> None:
        pass


class FakeWindow(FakeContainer):
    def __init__(self, title: str, class_name: str = "#32770", visible: bool = True):
        super().__init__(text=title)
        self.title = title
        self._class = class_name
        self.visible = visible

    def window_text(self) -> str:
        return self.title

    def class_name(self) -> str:
        return self._class

    def is_visible(self) -> bool:
        return self.visible

    def set_focus(self) -> None:
        pass

    def type_keys(self, keys: str) -> None:
        if self.on_keys:
            self.on_keys(keys)

    on_keys = None

    def descendants(self, control_type=None, **_kwargs):
        """递归查找后代控件（与 pywinauto 行为一致）。"""
        found = []
        for child in self.children:
            if not control_type or child.control_type == control_type:
                found.append(child)
            if hasattr(child, "children") and child.children:
                found.extend(child.descendants(control_type=control_type))
        return found


class FakeDesktop:
    def __init__(self, windows):
        self._windows = windows  # 可以是列表，也可以是返回列表的函数（模拟对话框动态出现）

    def windows(self):
        return list(self._windows() if callable(self._windows) else self._windows)


def make_case(**kwargs) -> PrintCase:
    base = dict(
        case_id="TC-UI", enabled=True, file_path="报告.docx", printer="Test Printer",
        paper="A4 (210 x 297mm)", media_type="普通纸", bin="纸盒1",
        duplex="duplexlong", print_quality="600", borderless="no", copies=2,
    )
    base.update(kwargs)
    return PrintCase(**base)


def build_ui(*, paper_options=None, print_clicked=None, driver_pref_button=True, confirm_close=True):
    """构造文档窗口 + 打印对话框 + 驱动属性窗口的替身。"""
    paper_options = paper_options or ["A4 (210 x 297mm)", "A3 (297 x 420mm)"]
    doc_window = FakeWindow("报告.docx - Word", class_name="OpusApp")
    dialog = FakeWindow("打印", class_name="#32770")
    driver = FakeWindow("打印首选项", class_name="#32770")

    state = {"dialog_open": False, "print_clicked": False}

    # 打印对话框：打印机下拉、份数、页码范围、首选项按钮、打印按钮
    printer_row = FakeContainer()
    FakeTextLabel = FakeText("打印机")
    printer_combo = FakeCombo("打印机", ["Pantum CM806ADN", "Microsoft Print to PDF"], parent=printer_row)
    printer_combo.edit = FakeEdit("", printer_row)
    dialog.children.append(printer_row)
    printer_row._parent = dialog

    copies_row = FakeContainer()
    copies_edit = FakeEdit("1", copies_row)
    copies_label = FakeText("份数")
    copies_row.children.insert(0, copies_label)
    copies_label._parent = copies_row
    dialog.children.append(copies_row)
    copies_row._parent = dialog

    pages_row = FakeContainer()
    pages_radio = FakeRadio("页码范围", pages_row)
    pages_edit = FakeEdit("", pages_row)
    pages_row.children.insert(0, FakeText("页码范围"))
    pages_row.children[0]._parent = pages_row
    dialog.children.append(pages_row)
    pages_row._parent = dialog

    def open_driver():
        state["dialog_open"] = True

    if driver_pref_button:
        FakeButton("首选项", on_click=open_driver, parent=dialog)

    def click_print():
        state["print_clicked"] = True
        if confirm_close:
            dialog.visible = False

    FakeButton("打印", on_click=click_print, parent=dialog)

    # 驱动属性窗口：纸张尺寸 / 纸张类型 / 纸张来源 / 双面 / 打印质量 / 无边距 / 确定
    driver_rows = {}
    for label, options, key in (
        ("纸张尺寸", paper_options, "paper"),
        ("纸张类型", ["普通纸", "厚纸2(158-209gsm)"], "media"),
        ("纸张来源", ["自动选择", "纸盒1", "多功能进纸器"], "bin"),
        ("双面", ["单面", "双面长边", "双面短边"], "duplex"),
        ("打印质量", ["600", "1200"], "quality"),
    ):
        row = FakeContainer()
        FakeCombo(label, options, parent=row)
        label_control = FakeText(label)
        row.children.insert(0, label_control)
        label_control._parent = row
        driver.children.append(row)
        row._parent = driver
        driver_rows[key] = row.children[1]

    FakeCheck("无边距", checked=0, parent=driver)
    FakeButton("确定", on_click=lambda: setattr(driver, "visible", False), parent=driver)

    def on_keys(_keys):
        if not state["dialog_open"]:
            state["dialog_open"] = True

    doc_window.on_keys = on_keys

    def window_list():
        windows = [doc_window]
        if state["dialog_open"]:
            windows.append(dialog)
            windows.append(driver)
        return windows

    desktop = FakeDesktop(window_list)
    opened = {"closed": False}

    def opener(_path, _log):
        def close():
            opened["closed"] = True
        return close, "报告"

    return desktop, opener, dialog, driver, driver_rows, state, opened


def test_ui_print_full_flow_sets_driver_attributes(tmp_path):
    docx = tmp_path / "报告.docx"
    docx.write_bytes(b"PK")
    desktop, opener, dialog, driver, rows, state, opened = build_ui()

    logs = []
    backend, stem = print_with_ui(
        "Pantum CM806ADN", make_case(pages="1-2"), docx, logs.append,
        timeout=3, desktop=desktop, opener=opener,
    )

    assert backend == "ui-automation" and stem == "报告"
    assert state["print_clicked"] is True          # 最终点击了“打印”
    assert opened["closed"] is True                # 打印后关闭文档
    assert rows["paper"].selected == "A4 (210 x 297mm)"
    assert rows["media"].selected == "普通纸"
    assert rows["bin"].selected == "纸盒1"
    assert rows["duplex"].selected == "双面长边"
    assert rows["quality"].selected == "600"
    # 份数 / 页码已填入打印对话框
    copiedits = dialog.descendants(control_type="Edit")
    assert any(getattr(e, "value", "") == "2" for e in copiedits)
    assert any(getattr(e, "value", "") == "1-2" for e in copiedits)
    # 前端可见的操作步骤
    joined = "\n".join(logs)
    assert "Ctrl+P" in joined
    assert "已选择打印机驱动 Pantum CM806ADN" in joined
    assert "纸张尺寸 设置为 A4 (210 x 297mm)" in joined
    assert "打印作业已提交" in joined


def test_ui_print_fails_when_driver_option_missing(tmp_path):
    """驱动下拉里没有用例要求的纸张 → 判定失败且不点“打印”。"""
    docx = tmp_path / "报告.docx"
    docx.write_bytes(b"PK")
    desktop, opener, dialog, driver, rows, state, opened = build_ui(paper_options=["Letter (8.5 x 11inch)"])

    with pytest.raises(UiPrintError) as excinfo:
        print_with_ui("Pantum CM806ADN", make_case(), docx, lambda _t: None,
                      timeout=3, desktop=desktop, opener=opener)
    assert "A4" in str(excinfo.value)
    assert state["print_clicked"] is False
    assert opened["closed"] is True


def test_ui_print_rejects_non_document_formats(tmp_path):
    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"%PDF")
    with pytest.raises(UiPrintError) as excinfo:
        print_with_ui("Pantum CM806ADN", make_case(), pdf, lambda _t: None, timeout=1)
    assert "GDI" in str(excinfo.value)


def test_default_channel_routes_documents_to_ui_and_other_formats_to_renderer(tmp_path, monkeypatch):
    """统一通道：Office 文档走 UI 自动化；PDF/图片由内置渲染器送入队列。"""
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from test_runner import FakeSpooler  # noqa: WPS433

    from printcase import runner as runner_module
    from printcase.audit import AuditStore
    from printcase.runner import PrintRunner

    calls = []

    def fake_ui(printer, case, source_path, log, timeout=90, **kwargs):
        calls.append(("ui", Path(source_path).suffix))
        return "ui-automation", Path(source_path).stem

    def fake_driver(printer, case, source_path, log, **kwargs):
        calls.append(("renderer", Path(source_path).suffix))
        return "gdi-pdf", Path(source_path).stem

    monkeypatch.setattr(runner_module, "print_with_ui", fake_ui)
    monkeypatch.setattr(runner_module, "print_with_driver", fake_driver)

    runner = PrintRunner(FakeSpooler(), AuditStore(tmp_path / "ui.sqlite3"))
    docx = tmp_path / "报告.docx"
    docx.write_bytes(b"PK")
    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"%PDF")

    doc_case = PrintCase(case_id="TC-DOC", enabled=True, file_path=str(docx), resolved_path=docx, printer="Test Printer")
    pdf_case = PrintCase(case_id="TC-PDF", enabled=True, file_path=str(pdf), resolved_path=pdf, printer="Test Printer")

    runner._default_ui_submit("Test Printer", doc_case, docx, lambda _t: None, tmp_path)
    runner._default_ui_submit("Test Printer", pdf_case, pdf, lambda _t: None, tmp_path)

    assert calls == [("ui", ".docx"), ("renderer", ".pdf")]
