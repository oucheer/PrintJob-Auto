"""真实链路复现：3 条用例（TC-013~015）串行下发，外部独立高频监视队列。

用途：
    1. 复现“3 条用例只出来 2 份”的竞态；
    2. 对比执行器自己的队列观测与外部观测，判断问题在“没提交”还是“没看到”。

用法：
    .venv/Scripts/python.exe tools/verify_print_chain.py [TC-013 TC-014 TC-015]
"""
from __future__ import annotations

import sys
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from printcase.audit import AuditStore  # noqa: E402
from printcase.excel_reader import read_cases  # noqa: E402
from printcase.runner import PrintRunner  # noqa: E402
from printcase.spooler import WindowsSpooler  # noqa: E402

PRINTER = "Pantum M9705DN Plus-3700 Series"
wanted = [a for a in sys.argv[1:] if a.upper().startswith("TC-")] or ["TC-013", "TC-014", "TC-015"]

spooler = WindowsSpooler()
cases, _ = read_cases(ROOT / "templates" / "PrintCase_Template.xlsx")
selected = [c for c in cases if c.case_id in wanted]
missing = set(wanted) - {c.case_id for c in selected}
if missing:
    print("找不到用例：", missing)
    raise SystemExit(2)

timeline: list[tuple[float, str]] = []
stop = threading.Event()
start_tick = time.monotonic()


def watchdog() -> None:
    """独立线程：20ms 采样，记录队列里每个作业的出现/状态变化/消失。"""
    prev: dict[int, str] = {}
    while not stop.is_set():
        now = time.monotonic() - start_tick
        try:
            jobs = {j.job_id: f"{j.document!r} {j.status_text} {j.pages_printed}/{j.total_pages}" for j in spooler.get_jobs(PRINTER)}
        except Exception as exc:  # noqa: BLE001
            timeline.append((now, f"!! 读队列失败 {type(exc).__name__}: {exc}"))
            time.sleep(0.02)
            continue
        for job_id, text in jobs.items():
            if prev.get(job_id) != text:
                timeline.append((now, f"job#{job_id} {text}"))
        for job_id in set(prev) - set(jobs):
            timeline.append((now, f"job#{job_id} 离开队列"))
        prev = jobs
        time.sleep(0.02)


threading.Thread(target=watchdog, daemon=True).start()

print(f"=== 真实下发 {datetime.now():%H:%M:%S} 用例 {', '.join(wanted)} ===")
for c in selected:
    print(f"  {c.case_id}: pages={c.pages!r} q={c.print_quality} borderless={c.borderless}")

temp_dir = Path(tempfile.mkdtemp(prefix="printcase_chain_"))
audit = AuditStore(temp_dir / "chain.sqlite3")
runner = PrintRunner(spooler, audit, queue_timeout=30)


def on_event(kind, payload):
    stamp = time.monotonic() - start_tick
    if kind == "log":
        print(f"{stamp:7.1f}s [log ] {payload}", flush=True)
    elif kind == "result":
        print(f"{stamp:7.1f}s [stat] {payload.case_id} → {payload.status.value}"
              f"{' / Job#' + str(payload.job_id) if payload.job_id else ''}", flush=True)


started = time.monotonic()
results = runner.run_cases(selected, PRINTER, f"CHAIN-{int(time.time())}", on_event)
elapsed = time.monotonic() - started
stop.set()
time.sleep(0.3)

print("-" * 72)
print(f"=== 结束（{elapsed:.1f}s）===")
for r in results:
    print(f"  {r.case_id}: {r.status.value} / Job#{r.job_id or '未关联'} / {r.message[:70]}")

print("-" * 72)
print("外部监视器看到的队列作业：")
jobs_seen = sorted({line.split()[0] for _, line in timeline if line.startswith("job#")})
print(" ", jobs_seen or "无")
print("时间线（仅关键节点）：")
for stamp, text in timeline:
    if "离开队列" in text or "打印中" in text:
        print(f"{stamp:7.2f}s  {text}")
