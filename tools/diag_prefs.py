"""验证修复：驱动属性页的复用 / 快速失败重试 / 残留清理（真实驱动，不打印）。"""
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

PRINTER = sys.argv[1] if len(sys.argv) > 1 else "Pantum M9705DN Plus-3700 Series"


def main() -> int:
    desktop = _desktop()

    print("== 1. 正常打开（含复用） ==")
    t0 = time.monotonic()
    window = open_preferences(PRINTER, timeout=25, desktop=desktop, log=lambda t: print("   ·", t))
    print(f"   打开耗时 {time.monotonic() - t0:.1f}s，标题={_text_of(window)!r}")

    t0 = time.monotonic()
    again = open_preferences(PRINTER, timeout=25, desktop=desktop, log=lambda t: print("   ·", t))
    print(f"   复用耗时 {time.monotonic() - t0:.1f}s，同一窗口={again is window or _text_of(again) == _text_of(window)}")

    print("== 2. 清理残留（窗口 + printui 宿主） ==")
    cleared = close_stale_preferences(PRINTER, desktop=desktop, log=lambda t: print("   ·", t))
    print(f"   清理项：{cleared}")
    left = [w for w in desktop.windows() if PRINTER in _text_of(w)]
    print(f"   仍残留的窗口：{[_text_of(w) for w in left]}")

    print("== 3. printui 静默退出时的快速失败（用不存在的打印机触发） ==")
    t0 = time.monotonic()
    try:
        open_preferences("ZZ-NoSuchPrinter-9f3a", timeout=25, desktop=desktop, attempts=3,
                         log=lambda t: print("   ·", t))
        print("   意外成功？")
    except DriverUiError as exc:
        print(f"   {time.monotonic() - t0:.1f}s 后失败：{exc}")
        print("   ↑ 旧实现会在这种场景干等满 25 秒并直接把用例判失败")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
