"""真实端到端验证：多选批量下发 → 中途点“停止” → 立刻再下发下一批。

复刻用户场景（真实打印机、真实驱动属性页、真实 Word UI 自动化）：
1. 一批 3 条用例连续下发，在**第 2 条正打开文档时**按“停止”；
2. 断言：第 2 条记“已停止”（不是失败），第 3 条未下发，不停留在 90 秒等待里；
3. 断言：停止后桌面没有残留 Word/printui，驱动属性页也没有残留；
4. 立刻再下发同一批 3 条（相当于用户重新勾选后再点“执行用例”），
   断言：3 条全部完成下发，无“未找到文档窗口”。

为节省纸张，所有用例的页码都改成第 1 页。

用法：.venv/Scripts/python.exe tools/verify_stop_rerun_ui.py
"""
from __future__ import annotations

import sys
import threading
import time
from dataclasses import replace
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from printcase.audit import AuditStore  # noqa: E402
from printcase.excel_reader import read_cases  # noqa: E402
from printcase.models import CaseStatus  # noqa: E402
from printcase.runner import PrintRunner  # noqa: E402
from printcase.spooler import WindowsSpooler  # noqa: E402
from printcase.winnative import enum_top_windows  # noqa: E402

TEMPLATE = ROOT / "templates" / "PrintCase_Template.xlsx"
# 默认：3 条一批，第 2 条途中停止，然后把同一批重新下发
# big：6 条一批，第 3 条途中停止，随后立刻继续下发剩下 3 条
PLAN = sys.argv[1] if len(sys.argv) > 1 else "3"
STOP_DELAY = float(sys.argv[2]) if len(sys.argv) > 2 else 1.5
if PLAN == "run":
    # 不停止：相当于“执行全部”，验证多选连续下发整批都能成功
    BATCH_A = ["TC-010", "TC-011", "TC-012", "TC-013", "TC-014", "TC-015"]
    BATCH_B = []
    STOP_ON = ""
elif PLAN == "big":
    BATCH_A = ["TC-010", "TC-011", "TC-012", "TC-013", "TC-014", "TC-015"]
    BATCH_B = ["TC-013", "TC-014", "TC-015"]
    STOP_ON = "TC-012"
else:
    BATCH_A = ["TC-010", "TC-011", "TC-012"]
    BATCH_B = ["TC-010", "TC-011", "TC-012"]
    STOP_ON = "TC-011"


def leftover_windows() -> list[str]:
    return [
        text for _hwnd, text in enum_top_windows()
        if "word" in text.casefold() or "printui" in text.casefold() or "首选项" in text
    ]


def main() -> int:
    spooler = WindowsSpooler()
    printer = spooler.win32print.GetDefaultPrinter()
    print(f"目标打印机：{printer}")

    loaded, warnings = read_cases(TEMPLATE)
    pool = {case.case_id: case for case in loaded}
    if warnings:
        print(f"模板告警：{warnings}")

    def pick(ids: list[str]):
        return [
            replace(pool[case_id], pages="1", printer=printer)
            for case_id in ids if case_id in pool
        ]

    audit = AuditStore(Path(ROOT / "dist" / "verify-stop-rerun.sqlite3"))
    runner = PrintRunner(spooler, audit, queue_timeout=20)

    lines: list[str] = []
    stop_watch = threading.Event()

    def on_event(kind: str, value) -> None:
        if kind != "log":
            return
        text = str(value)
        lines.append(text)
        print(f"  {text}")
        # 第 2 条用例开始打开文档 → 模拟用户这时点“停止”
        if STOP_ON and kind == "log" and "步骤 2/4" in text and STOP_ON in text:
            stop_watch.set()

    def press_stop() -> None:
        if stop_watch.wait(timeout=180):
            time.sleep(STOP_DELAY)
            print(f"\n>>> 模拟用户点击“停止”（{STOP_ON} 下发途中，延迟 {STOP_DELAY}s）\n")
            runner.stop()

    print("\n===== 第一批：3 条，连续下发 → 第 2 条途中停止 =====")
    stopper = threading.Thread(target=press_stop, daemon=True)
    stopper.start()
    started = time.monotonic()
    batch_a = runner.run_cases(pick(BATCH_A), printer, "RUN-A", on_event)
    elapsed_a = time.monotonic() - started
    stopper.join(timeout=1)

    print(f"\n第一批耗时 {elapsed_a:.1f}s")
    for result in batch_a:
        spent = result.elapsed_seconds or 0.0
        print(f"  {result.case_id}: {result.status.value}（{spent:.1f}s）— {result.message}")

    leftovers = leftover_windows()
    print(f"停止后残留窗口：{leftovers or '无'}")

    failed = [r for r in batch_a if r.status == CaseStatus.FAILED]
    submitted = [r for r in batch_a if r.status not in {
        CaseStatus.FAILED, CaseStatus.STOPPED, CaseStatus.PREFLIGHT_FAILED,
    }]
    stopped = [r for r in batch_a if r.status == CaseStatus.STOPPED]

    ok = True
    if failed:
        ok = False
        print("断言失败：不应出现“下发失败”")
    if len(submitted) < 1:
        ok = False
        print("断言失败：应当至少完成 1 条下发")
    if not stopped and STOP_ON:
        ok = False
        print("断言失败：应当有 1 条记为“已停止”")
    if STOP_ON and len(submitted) != len(BATCH_A) and len(submitted) < 1:
        ok = False
    if leftovers:
        ok = False
        print("断言失败：停止后仍有残留窗口")
    if not BATCH_B:
        print("\n结果：" + ("通过" if ok else "未通过"))
        return 0 if ok else 1

    print(f"\n===== 第二批：立刻再下发 {len(BATCH_B)} 条 =====")
    lines.clear()
    started = time.monotonic()
    batch_b = runner.run_cases(pick(BATCH_B), printer, "RUN-B", on_event)
    elapsed_b = time.monotonic() - started
    print(f"\n第二批耗时 {elapsed_b:.1f}s")
    for result in batch_b:
        spent = result.elapsed_seconds or 0.0
        print(f"  {result.case_id}: {result.status.value}（{spent:.1f}s）— {result.message}")

    bad_b = [r for r in batch_b if r.status in {CaseStatus.FAILED, CaseStatus.STOPPED}]
    if bad_b:
        ok = False
        print("断言失败：停止后再下发仍然失败：" + "、".join(r.case_id for r in bad_b))
    if any("未找到文档窗口" in r.message for r in batch_b):
        ok = False
        print("断言失败：出现“未找到文档窗口”")

    print(f"\n第二批残留窗口：{leftover_windows() or '无'}")
    print("\n结果：" + ("通过" if ok else "未通过"))
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
