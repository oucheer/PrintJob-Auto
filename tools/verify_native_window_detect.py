"""验证：UIA 全量枚举失效时，原生 EnumWindows 仍能发现并操作驱动属性页。

用户现场的症状就是“printui 明明起来了，程序却报窗口未出现”。这里直接注入
UIA 枚举失败，验证新检测方式能否救回。
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from printcase.driver_ui import (  # noqa: E402
    DriverUiError,
    _desktop,
    _text_of,
    close_stale_preferences,
    open_preferences,
)
from printcase.excel_reader import read_cases  # noqa: E402

PRINTER = "Pantum M9705DN Plus-3700 Series"
ROOT = Path(__file__).resolve().parents[1]


class BlindDesktop:
    """UIA 全量枚举直接抛异常（模拟 UIA 客户端失效），单句柄查询仍可用。"""

    def __init__(self, real) -> None:
        self._real = real
        self.scans = 0

    def windows(self):
        self.scans += 1
        raise RuntimeError("模拟 UIA 全量枚举失败")

    def window(self, *args, **kwargs):
        if "handle" not in kwargs:
            raise RuntimeError("只允许按句柄取窗口")
        return self._real.window(*args, **kwargs)


class FullyBlindDesktop(BlindDesktop):
    """连单句柄查询也失败（UIA 彻底不可用）。"""

    def window(self, *args, **kwargs):
        raise RuntimeError("模拟 UIA 单句柄查询失败")


def main() -> int:
    real = _desktop()
    close_stale_preferences(PRINTER, desktop=real)

    print("== 1. UIA 全局遍历失效，仅靠原生枚举 ==")
    blind = BlindDesktop(real)
    t0 = time.monotonic()
    try:
        window = open_preferences(PRINTER, timeout=20, desktop=blind, log=lambda t: print("   ·", t))
        print(f"   成功：{time.monotonic() - t0:.1f}s，标题={_text_of(window)!r}")
        print(f"   （UIA 全量遍历被拒绝 {blind.scans} 次，仍然找到窗口）")
        close_stale_preferences(PRINTER, desktop=real)
    except DriverUiError as exc:
        print(f"   失败：{time.monotonic() - t0:.1f}s → {exc}")
        return 1

    print("\n== 2. 在这种状态下完整跑一次驱动属性页设置（TC-010） ==")
    from printcase.driver_ui import apply_case

    cases, _w = read_cases(ROOT / "templates" / "PrintCase_Template.xlsx")
    case = next(c for c in cases if c.case_id == "TC-010")
    blind = BlindDesktop(real)
    t0 = time.monotonic()
    try:
        restore = apply_case(PRINTER, case, caps={"color": False}, desktop=blind,
                             log=lambda t: print("   ·", t))
        print(f"   设置成功：{time.monotonic() - t0:.1f}s")
        restore()
    except DriverUiError as exc:
        print(f"   设置失败：{time.monotonic() - t0:.1f}s → {exc}")
        return 1
    close_stale_preferences(PRINTER, desktop=real)

    print("\n== 3. UIA 彻底不可用时应给出准确原因（而不是“窗口未出现”） ==")
    fully = FullyBlindDesktop(real)
    t0 = time.monotonic()
    try:
        open_preferences(PRINTER, timeout=12, attempts=1, desktop=fully)
        print("   意外成功？")
    except DriverUiError as exc:
        cost = time.monotonic() - t0
        text = str(exc)
        print(f"   {cost:.1f}s 后失败：{text[:200]}")
        assert "无法通过 UI 自动化读取" in text, "失败原因不够准确"
        assert "重启程序" in text, "缺少可操作建议"
        print("   ✓ 原因准确，并给出了“重启程序”的建议")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
