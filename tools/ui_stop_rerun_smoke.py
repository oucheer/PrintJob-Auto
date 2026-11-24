"""界面级黑盒：真实点“执行勾选用例” → 中途点“停止” → 立刻再执行下一批。

对应用户反馈的操作序列（真实窗口、真实 Word、真实打印机，每条用例只印 1 页）：
1. 勾选 TC-010/011/012 → 点“执行勾选用例”；
2. 第 2 条正在打开文档时点“停止”；
3. 断言界面在 30 秒内回到可操作状态，且第 2/3 条显示“已停止”，没有“下发失败”；
4. 立刻勾选 TC-013/014 → 再点“执行勾选用例”；
5. 断言两条都完成下发、没有“未找到文档窗口”，桌面无残留 Word 窗口。

用法：.venv/Scripts/python.exe tools/ui_stop_rerun_smoke.py
"""
from __future__ import annotations

import sys
import time
import tkinter as tk
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SHOT_DIR = ROOT / "dist" / "ui-stop-smoke"
sys.argv = [sys.argv[0], str(SHOT_DIR)]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tools"))

import printcase.app as app_module  # noqa: E402
from printcase.app import PrintCaseApp  # noqa: E402
from printcase.models import CaseStatus  # noqa: E402
from printcase.winnative import enum_top_windows  # noqa: E402
import ui_gui_smoke as gui  # noqa: E402  复用截图与假对话框

FIRST = ["TC-010", "TC-011", "TC-012"]
SECOND = ["TC-013", "TC-014"]
STOP_LATENCY_LIMIT = 30.0

results: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    results.append((name, bool(ok), detail))
    print(f"[{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""), flush=True)


def pump(root: tk.Tk, seconds: float, condition=None) -> bool:
    """跑 Tk 事件循环若干秒；condition 成立即提前返回。"""
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        root.update()
        if condition is not None and condition():
            return True
        time.sleep(0.05)
    return bool(condition is not None and condition())


def log_text(app: PrintCaseApp) -> str:
    return app.log_text.get("1.0", "end")


def select(app: PrintCaseApp, case_ids: list[str]) -> None:
    app.checked = {index for index, case in enumerate(app.cases) if case.case_id in set(case_ids)}
    app._sync_select_all()
    app._render_cases()


def leftover_windows() -> list[str]:
    return [
        text for _hwnd, text in enum_top_windows()
        if "word" in text.casefold() or "printui" in text.casefold()
    ]


def main() -> int:
    app_module.messagebox = gui.FakeBox()
    app_module.filedialog.askopenfilename = lambda **_k: str(ROOT / "templates" / "PrintCase_Template.xlsx")

    root = tk.Tk()
    app = PrintCaseApp(root)
    root.update()
    app._import_excel()
    root.update()
    check("导入用例", len(app.cases) == 22, f"{len(app.cases)} 条")

    print("\n===== 第一批：勾选 3 条，执行中途中止 =====")
    select(app, FIRST)
    root.update()
    check("勾选 3 条", len(app.checked) == 3, app.case_count_label.cget("text"))

    app._start_run(selected_only=True)
    root.update()
    check("开始执行且停止按钮可用", app.running and str(app.btn_stop.cget("state")) == "normal")

    target = f"[{FIRST[1]}] 步骤 2/4"
    reached = pump(root, 180, lambda: target in log_text(app))
    check("第 2 条进入打开文档阶段", reached)
    time.sleep(1.2)
    stopped_at = time.monotonic()
    app._stop_run()
    root.update()
    print(">>> 已点击“停止”")

    done = pump(root, 120, lambda: not app.running)
    latency = time.monotonic() - stopped_at
    check("停止后界面及时回到可操作状态", done, f"{latency:.1f}s")
    check("停止响应时间小于 30 秒", latency < STOP_LATENCY_LIMIT, f"{latency:.1f}s")
    check("停止后按钮状态恢复", not app.running and str(app.btn_stop.cget("state")) == "disabled")
    check("状态栏标注已停止", "已停止" in app.status_var.get(), app.status_var.get())

    first_status = {case_id: (app.case_results.get(case_id).status if app.case_results.get(case_id) else None)
                    for case_id in FIRST}
    for case_id, status in first_status.items():
        print(f"    {case_id}: {status.value if status else '无结果'}")
    check("停止后没有“下发失败”",
          all(status != CaseStatus.FAILED for status in first_status.values()),
          "、".join(f"{cid}={s.value if s else '无'}" for cid, s in first_status.items()))
    check("有 1 条记为“已停止”",
          any(status == CaseStatus.STOPPED for status in first_status.values()))
    shot = gui.shot(root, "01-停止后.png")
    check("停止后界面截图", shot is not None and shot.exists(), shot.name if shot else "失败")
    check("停止后无残留 Word 窗口", leftover_windows() == [], str(leftover_windows()))

    print("\n===== 第二批：立刻再勾选 2 条执行 =====")
    select(app, SECOND)
    root.update()
    app._start_run(selected_only=True)
    root.update()
    done = pump(root, 240, lambda: not app.running)
    check("第二批执行完成", done)
    second_status = {case_id: (app.case_results.get(case_id).status if app.case_results.get(case_id) else None)
                     for case_id in SECOND}
    for case_id, status in second_status.items():
        print(f"    {case_id}: {status.value if status else '无结果'}")
    check("第二批没有失败/停止",
          all(status not in {None, CaseStatus.FAILED, CaseStatus.STOPPED} for status in second_status.values()),
          "、".join(f"{cid}={s.value if s else '无'}" for cid, s in second_status.items()))
    messages = " ".join(
        (app.case_results.get(case_id).message if app.case_results.get(case_id) else "") for case_id in SECOND
    )
    check("第二批没有“未找到文档窗口”", "未找到文档窗口" not in messages)
    shot = gui.shot(root, "02-再次执行完成.png")
    check("第二批界面截图", shot is not None and shot.exists(), shot.name if shot else "失败")
    check("最终无残留 Word 窗口", leftover_windows() == [], str(leftover_windows()))

    root.update()
    root.destroy()

    passed = sum(1 for _name, ok, _detail in results if ok)
    print("\n" + "-" * 70)
    print(f"共 {len(results)} 项：通过 {passed}，失败 {len(results) - passed}")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
