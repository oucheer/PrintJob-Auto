"""无头验证：驱动属性页自检是否通过（不弹错误框）。"""
from __future__ import annotations

import sys
import time
import tkinter as tk
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from printcase import app as app_module  # noqa: E402

warnings_seen: list[str] = []
app_module.messagebox.showwarning = lambda title, msg, **kw: warnings_seen.append(f"{title}: {msg}")

root = tk.Tk()
root.withdraw()
application = app_module.PrintCaseApp(root)
root.update()

application._check_driver_ui()

deadline = time.time() + 40
while time.time() < deadline:
    root.update()
    time.sleep(0.1)
    if "自检中" not in application.driver_label.cget("text"):
        break

print("状态栏：", application.status_var.get())
print("驱动属性页：", application.driver_label.cget("text"))
print("警告框数量：", len(warnings_seen))
for item in warnings_seen:
    print("  警告：", item)

text = application.log_text.get("1.0", "end") if hasattr(application, "log_text") else ""
print("--- 日志尾部 ---")
print("\n".join(line for line in text.strip().splitlines()[-6:]))
root.destroy()
