"""观测真实打印作业在 Windows 队列里的可见性。

用法：python tools/diag_queue_probe.py [等待关闭_秒数]
默认在 PrintOut 后一直观察 20 秒再关闭 Word。
"""
from __future__ import annotations

import sys
import threading
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import win32print  # noqa: E402

from printcase.spooler import WindowsSpooler  # noqa: E402

PRINTER = "Pantum M9705DN Plus-3700 Series"
DOC = Path(__file__).resolve().parents[1] / "templates" / "docs" / "A4纵向测试页2.docx"
WATCH_SECONDS = float(sys.argv[1]) if len(sys.argv) > 1 else 20.0
# 打印后多久关闭 Word（-1 表示到观察结束才关）；接近 0 可以复现“提交后立刻退出”
CLOSE_DELAY = float(sys.argv[2]) if len(sys.argv) > 2 else -1.0

events: list[tuple[float, str]] = []
lock = threading.Lock()
stop = threading.Event()


def poll() -> None:
    spooler = WindowsSpooler()
    seen_prev: dict[int, str] = {}
    while not stop.is_set():
        now = time.monotonic()
        try:
            jobs = {j.job_id: f"{j.document!r} {j.status_text} {j.pages_printed}/{j.total_pages}" for j in spooler.get_jobs(PRINTER)}
        except Exception as exc:  # noqa: BLE001
            with lock:
                events.append((now, f"!! 读队列失败 {type(exc).__name__}: {exc}"))
            time.sleep(0.02)
            continue
        for job_id, text in jobs.items():
            if seen_prev.get(job_id) != text:
                with lock:
                    events.append((now, f"job#{job_id} {text}"))
        for job_id in set(seen_prev) - set(jobs):
            with lock:
                events.append((now, f"job#{job_id} 离开队列"))
        seen_prev = jobs
        time.sleep(0.02)


thread = threading.Thread(target=poll, daemon=True)
thread.start()
time.sleep(1.0)

import pythoncom  # noqa: E402
import win32com.client  # noqa: E402

pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application")
app.Visible = True
app.DisplayAlerts = 0
doc = app.Documents.Open(str(DOC), ReadOnly=True, AddToRecentFiles=False)
print(f"文档页数：{doc.ComputeStatistics(2)}")

t0 = time.monotonic()
with lock:
    events.append((t0, "===== 调用 PrintOut ====="))
doc.PrintOut(Background=True)
print(f"PrintOut 返回耗时 {time.monotonic() - t0:.2f}s")

closed = False
end = time.monotonic() + WATCH_SECONDS
while time.monotonic() < end:
    if not closed and CLOSE_DELAY >= 0 and (time.monotonic() - t0) >= CLOSE_DELAY:
        with lock:
            events.append((time.monotonic(), "===== 关闭文档并退出 Word ====="))
        try:
            doc.Close(False)
            app.Quit()
        except Exception as exc:  # noqa: BLE001
            print("关闭失败:", exc)
        closed = True
    try:
        status = app.BackgroundPrintingStatus if not closed else "已关闭"
    except Exception as exc:  # noqa: BLE001
        status = f"读取失败 {exc}"
    with lock:
        events.append((time.monotonic(), f"BackgroundPrintingStatus={status}"))
    time.sleep(1.0)

if not closed:
    with lock:
        events.append((time.monotonic(), "===== 关闭文档并退出 Word ====="))
    try:
        doc.Close(False)
        app.Quit()
    except Exception as exc:  # noqa: BLE001
        print("关闭失败:", exc)
pythoncom.CoUninitialize()
time.sleep(3.0)
stop.set()
thread.join(timeout=2)

print("\n--- 事件时间线（相对 PrintOut） ---")
for stamp, text in events:
    print(f"{stamp - t0:+7.2f}s  {text}")
