"""驱动属性页权限与黑白设备处理的回归测试（全部用替身，不依赖真实打印机）。"""
from __future__ import annotations

from printcase.driver_ui import plan_actions, restore, snapshot
from printcase.models import PrintCase


def _case(**kwargs):
    base = dict(case_id="TC", enabled=True, file_path="a.pdf", paper="A4", duplex="simplex")
    base.update(kwargs)
    return PrintCase(**base)


def test_monochrome_device_skips_color_action():
    """黑白设备上报 color=False 时不应生成“颜色”动作，否则严格模式拒绝下发。"""
    actions = plan_actions(_case(color="color"), caps={"color": False})
    assert all(a.label != "颜色" for a in actions)

    # 显式黑白同样跳过（设备本身就是黑白，无可设置项）
    actions = plan_actions(_case(color="monochrome"), caps={"color": False})
    assert all(a.label != "颜色" for a in actions)


def test_color_device_and_unknown_caps_keep_color_action():
    """彩色设备、或驱动未提供能力数据时，仍按原逻辑到界面里设置颜色。"""
    assert any(a.label == "颜色" for a in plan_actions(_case(color="color"), caps={"color": True}))
    assert any(a.label == "颜色" for a in plan_actions(_case(color="color")))
    assert any(a.label == "颜色" for a in plan_actions(_case(color="color"), caps={"color": None}))


def test_snapshot_uses_use_access_only():
    """只读快照必须用 PRINTER_ACCESS_USE，不能用 PRINTER_ALL_ACCESS（普通用户被拒）。"""
    captured = {}

    class FakeWin:
        PRINTER_ACCESS_USE = 0x8
        PRINTER_ALL_ACCESS = 0xF003C

        def OpenPrinter(self, name, defaults):
            captured["access"] = defaults["DesiredAccess"]
            return "handle"

        def GetPrinter(self, handle, level):
            class DM:
                PaperSize = MediaType = DefaultSource = Duplex = 1
                Color = Copies = PrintQuality = Collate = 1

            return {"pDevMode": DM()}

        def ClosePrinter(self, handle):
            pass

    snapshot("Test Printer", w=FakeWin())
    assert captured["access"] == FakeWin.PRINTER_ACCESS_USE


def test_restore_falls_back_after_administer_get_failure():
    """ADMINISTER 打开成功但读取失败时，必须改试 USE，且不能被假成功误导。"""
    calls = []

    class FakeWin:
        PRINTER_ACCESS_ADMINISTER = 0x4
        PRINTER_ACCESS_USE = 0x8

        def OpenPrinter(self, name, defaults):
            calls.append(("open", defaults["DesiredAccess"]))
            return "h-admin" if defaults["DesiredAccess"] == self.PRINTER_ACCESS_ADMINISTER else "h-use"

        def GetPrinter(self, handle, level):
            calls.append(("get", handle))
            if handle == "h-admin":
                raise RuntimeError("拒绝访问")

            class DM:
                pass

            return {"pDevMode": DM()}

        def SetPrinter(self, handle, level, info, flags):
            calls.append(("set", handle))

        def ClosePrinter(self, handle):
            calls.append(("close", handle))

    ok = restore({"devmode": object()}, "Test Printer", w=FakeWin())
    assert ok is True
    assert ("set", "h-use") in calls
    assert ("close", "h-admin") in calls  # 失败的句柄必须关闭


def test_restore_degrades_gracefully_without_write_access():
    """所有权限都写不回时返回 False 并记录警告，不抛异常影响已下发的作业。"""
    logs = []

    class FakeWin:
        PRINTER_ACCESS_ADMINISTER = 0x4
        PRINTER_ACCESS_USE = 0x8

        def OpenPrinter(self, name, defaults):
            return "h"

        def GetPrinter(self, handle, level):
            class DM:
                pass

            return {"pDevMode": DM()}

        def SetPrinter(self, handle, level, info, flags):
            raise RuntimeError("拒绝访问")

        def ClosePrinter(self, handle):
            pass

    ok = restore({"devmode": object()}, "Test Printer", w=FakeWin(), log=logs.append)
    assert ok is False
    assert any("手工核对" in m for m in logs)


def _stub_snapshot(monkeypatch, value):
    from printcase import driver_ui

    monkeypatch.setattr(driver_ui, "snapshot", lambda printer, w=None: {"MediaType": value, "PaperSize": 9})


def test_verify_allows_unchanged_when_target_already_selected(monkeypatch):
    """目标纸张类型恰好等于驱动默认值（如普通纸=266）时，回读不变不能判失败。"""
    from printcase import driver_ui

    _stub_snapshot(monkeypatch, 266)
    actions = [driver_ui.Action("纸张", "纸张类型", "combo_index", 0, "MediaType")]
    before = {"MediaType": 266, "PaperSize": 9}

    # 未声明“已是目标项”时仍按原逻辑拦截（保持严格）
    assert driver_ui.verify("Test Printer", actions, before) == ["纸张类型：驱动默认值未变化（MediaType=266）"]
    # 界面里本来就是目标项：DEVMODE 不变是正确结果
    assert driver_ui.verify("Test Printer", actions, before, unchanged_labels={"纸张类型"}) == []


def test_verify_still_rejects_invalid_value_when_unchanged(monkeypatch):
    """即使界面已是目标项，回读值无效（0/空）仍需拦截。"""
    from printcase import driver_ui

    _stub_snapshot(monkeypatch, 0)
    actions = [driver_ui.Action("纸张", "纸张类型", "combo_index", 0, "MediaType")]
    problems = driver_ui.verify("Test Printer", actions, {"MediaType": 0}, unchanged_labels={"纸张类型"})
    assert problems and "回读值无效" in problems[0]


def test_combo_helpers_read_current_selection():
    """自绘/不可读的下拉不能抛异常，应安全退化为 None / 空串。"""
    from printcase import driver_ui

    class Readable:
        def selected_index(self):
            return 2

        def selected_text(self):
            return "普通纸(60-90gsm)"

    class Opaque:
        def selected_index(self):
            raise RuntimeError("not supported")

        def selected_text(self):
            raise RuntimeError("not supported")

    assert driver_ui._combo_current_index(Readable()) == 2
    assert driver_ui._combo_current_text(Readable()) == "普通纸(60-90gsm)"
    assert driver_ui._combo_current_index(Opaque()) is None
    assert driver_ui._combo_current_text(Opaque()) == ""
