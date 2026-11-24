"""诊断：printui /e 打不开首选项时会弹出什么窗口（用不存在的打印机触发）。"""
from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from printcase.driver_ui import _desktop, _text_of  # noqa: E402

BOGUS = "ZZ-NoSuchPrinter-9f3a"


def dump(desktop, since: float) -> None:
    for window in desktop.windows():
        try:
            pid = window.element_info.process_id
        except Exception:  # noqa: BLE001
            pid = -1
        try:
            cls = window.class_name()
        except Exception:  # noqa: BLE001
            cls = "?"
        text = _text_of(window)
        if not text and cls != "#32770":
            continue
        try:
            vis = window.is_visible()
        except Exception:  # noqa: BLE001
            vis = "?"
        print(f"  pid={pid} class={cls} visible={vis} title={text!r}")
        if cls == "#32770":
            for child in window.descendants():
                try:
                    child_text = child.window_text()
                except Exception:  # noqa: BLE001
                    child_text = ""
                if child_text:
                    print(f"      [{child.element_info.control_type}] {child_text!r}")


def main() -> int:
    desktop = _desktop()
    before = time.monotonic()
    proc = subprocess.Popen(["rundll32.exe", "printui.dll,PrintUIEntry", "/e", "/n", BOGUS], close_fds=True)
    print(f"已启动 printui（pid={proc.pid}）针对不存在的打印机 {BOGUS}")
    for i in range(1, 7):
        time.sleep(1.0)
        print(f"-- {i}s（进程存活={proc.poll() is None}）")
        dump(desktop, before)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
