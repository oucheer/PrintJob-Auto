"""启动打包后的 exe，抓窗口截图并检查进程是否稳定存活。"""
from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
EXE = ROOT / "dist" / "PrintCaseStudio.exe"
SHOT = ROOT / "dist" / "ui-smoke" / "05-打包后启动.png"

proc = subprocess.Popen([str(EXE)])
try:
    from pywinauto import Desktop

    window = None
    deadline = time.monotonic() + 40
    while time.monotonic() < deadline:
        try:
            for candidate in Desktop(backend="uia").windows():
                if "PrintCase Studio" in candidate.window_text():
                    window = candidate
                    break
        except Exception:  # noqa: BLE001
            pass
        if window is not None:
            break
        time.sleep(1.0)

    if window is None:
        print("[FAIL] 40 秒内没有出现主窗口")
        raise SystemExit(1)

    print(f"[PASS] 主窗口出现：{window.window_text()}")
    time.sleep(5)
    alive = proc.poll() is None
    print(f"[{'PASS' if alive else 'FAIL'}] 进程存活：{alive}")
    image = window.capture_as_image()
    SHOT.parent.mkdir(parents=True, exist_ok=True)
    image.save(SHOT)
    print(f"[PASS] 截图：{SHOT} ({image.width}x{image.height})")
finally:
    try:
        proc.terminate()
    except Exception:  # noqa: BLE001
        pass
    time.sleep(1)
    if proc.poll() is None:
        proc.kill()
    print("已关闭测试实例")
