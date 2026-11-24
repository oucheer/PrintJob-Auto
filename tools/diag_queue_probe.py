"""观测真实打印作业在 Windows 队列里的可见性。

用应用自己的 UI 自动化通道下发（页码限定为第 1 页），同时高频读队列，
记录作业从出现到离开队列的时间线。

注意：**不要**用 Word COM 的 PrintOut 做这个探测——`Range=4` 是
“指定页码范围（Pages 参数）”，传 From/To 会被忽略而打印整册文档
（这份测试文档有 443 个分页符，会真的出 443 页纸）。

用法：.venv/Scripts/python.exe tools/diag_queue_probe.py [观察秒数]
"""
from __future__ import annotations

import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from printcase.models import PrintCase  # noqa: E402
from printcase.spooler import WindowsSpooler  # noqa: E402
from printcase.ui_print import print_with_ui  # noqa: E402

PRINTER = "Pantum M9705DN Plus-3700 Series"
DOC = ROOT / "templates" / "docs" / "A4纵向测试页2.docx"
WATCH_SECONDS = float(sys.argv[1]) if len(sys.argv) > 1 else 25.0

events: list[tuple[float, str]] = []
lock = threading.Lock()
stop = threading.Event()


def poll() -> None:
    spooler = WindowsSpooler()
    seen_prev: dict[int, str] = {}
    while not stop.is_set():
        now = time.monotonic()
        try:
            jobs = {
                j.job_id: f"{j.document!r} {j.status_text} {j.pages_printed}/{j.total_pages}"
                for j in spooler.get_jobs(PRINTER)
            }
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


def main() -> int:
    thread = threading.Thread(target=poll, daemon=True)
    thread.start()
    time.sleep(1.0)

    case = PrintCase(
        case_id="QUEUE-PROBE", enabled=True, file_path=str(DOC),
        printer=PRINTER, pages="1", copies=1, paper="A4 (210 x 297mm)",
        media_type="普通纸(60-90gsm)", duplex="simplex", print_quality="600",
        borderless="no", color="monochrome",
    )
    case.resolved_path = DOC

    started = time.monotonic()
    with lock:
        events.append((started, "===== 通过 UI 自动化下发第 1 页 ====="))
    backend, stem = print_with_ui(
        PRINTER, case, DOC, lambda text: print(f"  · {text}"),
        set_driver_attributes=False, timeout=90,
    )
    print(f"下发完成：{backend} / {stem}，耗时 {time.monotonic() - started:.1f}s")

    end = time.monotonic() + WATCH_SECONDS
    while time.monotonic() < end:
        time.sleep(1.0)
    stop.set()
    thread.join(timeout=2)

    print("\n--- 事件时间线（相对下发开始） ---")
    for stamp, text in events:
        print(f"{stamp - started:+7.2f}s  {text}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
