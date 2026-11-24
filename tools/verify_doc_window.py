"""真实 Word 验证：文档窗口发现（UIA 全量枚举失效也不受影响）+ 停止后立刻重开。

背景：日志里出现“已启动 Word → 90 秒后 未找到文档窗口”，屏幕上 Word 其实
早就打开了。根因是 `Desktop.windows()`（UIA 桌面全量枚举）整体失效，旧实现
一路空等把用例超时耗光。这里用真实 Word 复现该场景并验证修复：

1. 冷启动打开 443 页文档 → 测量“窗口可见”与“找到窗口”的耗时差距；
2. 让 UIA 枚举彻底不可用（把 desktop.windows 替换成返回空），确认仍能找到窗口；
3. 模拟“点停止” → 立刻关闭 Word → 确认没有残留 WINWORD 进程；
4. 停止后马上重开同一文档 → 确认仍然一次成功。

用法：.venv/Scripts/python.exe tools/verify_doc_window.py
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from printcase import ui_print  # noqa: E402
from printcase.ui_print import UiPrintInterrupted, _find_window, _open_document_visible  # noqa: E402
from printcase.winnative import enum_top_windows, process_alive, window_title  # noqa: E402

DOC = ROOT / "templates" / "docs" / "A4纵向测试页2.docx"


class BlindDesktop:
    """UIA 全量枚举失效的桌面替身：windows() 返回空，只保留按句柄取值的能力。"""

    def __init__(self, real_desktop):
        self._real = real_desktop
        self.enumerations = 0

    def windows(self):
        self.enumerations += 1
        return []

    def window(self, handle):  # noqa: A003
        return self._real.window(handle=handle)


def word_windows() -> list[str]:
    return [
        text for _hwnd, text in enum_top_windows()
        if "word" in text.casefold() and "printcase" not in text.casefold()
    ]


def word_pids() -> list[int]:
    from printcase.winnative import window_pid

    pids = {window_pid(hwnd) for hwnd, text in enum_top_windows() if "word" in text.casefold()}
    return sorted(pid for pid in pids if pid)


def one_round(label: str, *, blind: bool, real_desktop) -> None:
    desktop = BlindDesktop(real_desktop) if blind else real_desktop
    log = lambda text: print(f"    · {text}")  # noqa: E731
    started = time.monotonic()
    close, title, _document, app, pid = _open_document_visible(DOC, log)
    opened = time.monotonic() - started
    try:
        found = _find_window(
            desktop, title, timeout=30, app=app, prefer_pid=pid, log=log,
        )
        spent = time.monotonic() - started
        print(f"    [{label}] 打开 {opened:.1f}s，找到窗口累计 {spent:.1f}s，标题={window_title(found.handle)!r}")
        assert found is not None
        if blind:
            print(f"    [{label}] UIA 全量枚举调用 {desktop.enumerations} 次（应为 0）")
            assert desktop.enumerations == 0
    finally:
        close()
        time.sleep(1.0)
        print(f"    [{label}] 关闭后残留 Word 进程：{word_pids() or '无'}")


def main() -> int:
    print(f"文档：{DOC.name}")
    real_desktop = ui_print._default_desktop()

    print("== 1. 常规（UIA 可用）三次冷启动 ==")
    for index in range(3):
        one_round(f"常规 {index + 1}", blind=False, real_desktop=real_desktop)

    print("== 2. UIA 全量枚举失效（真实故障复现）==")
    one_round("UIA 失效", blind=True, real_desktop=real_desktop)

    print("== 3. 模拟“点停止”：等待窗口期间收到停止请求 ==")
    log = lambda text: print(f"    · {text}")  # noqa: E731
    close, title, _document, app, pid = _open_document_visible(DOC, log)
    try:
        started = time.monotonic()
        try:
            _find_window(real_desktop, title, timeout=90, app=None, prefer_pid=None,
                         should_stop=lambda: True)
            print("    断言失败：没有按停止请求中止")
            return 1
        except UiPrintInterrupted as exc:
            print(f"    收到停止后在 {time.monotonic() - started:.2f}s 内中止：{exc}")
    finally:
        close()
        time.sleep(1.0)
    leftover = word_pids()
    print(f"    停止后残留 Word 进程：{leftover or '无'}")
    if leftover:
        print("    断言失败：停止后仍有残留 Word 进程")
        return 1

    print("== 4. 停止后立刻重新打开 ==")
    one_round("停止后重开", blind=True, real_desktop=real_desktop)
    print("验证通过：窗口发现稳定、停止立即响应、无残留进程。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
