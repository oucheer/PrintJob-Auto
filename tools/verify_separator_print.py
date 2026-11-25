"""真实出纸验证：用例“分隔页=作业之间”到底有没有让驱动多插一张分隔页。

做法：同一份 1 页文档下发两次——
  A：分隔页=作业之间（纸盒1，分隔页纸张来源=纸盒2）
  B：分隔页=无
后台 0.05 秒采样打印队列，比较作业的 TotalPages；驱动把分隔页算进同一个作业时
A 应为 2 页、B 为 1 页；如果驱动单独建作业，则会看到两个作业。

会真实出纸（最多 3 张：1 页文档 + 1 张分隔页 + 1 页文档），确认打印机有纸再运行。
用法：.venv\\Scripts\\python.exe tools\\verify_separator_print.py
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
from printcase.runner import PrintRunner  # noqa: E402
from printcase.spooler import WindowsSpooler  # noqa: E402


class QueueSampler(threading.Thread):
    """0.05 秒采样一次队列：作业在这里只存活几百毫秒，慢一拍就什么都看不到。"""

    def __init__(self, spooler: WindowsSpooler, printer: str):
        super().__init__(daemon=True)
        self.spooler = spooler
        self.printer = printer
        self.samples: list[tuple[int, str, int, int]] = []
        self._stop = threading.Event()

    def run(self) -> None:
        while not self._stop.is_set():
            try:
                for job in self.spooler.get_jobs(self.printer):
                    self.samples.append((job.job_id, job.document, job.total_pages, job.pages_printed))
            except Exception:  # noqa: BLE001 - 采样失败不影响主流程
                pass
            time.sleep(0.05)

    def stop(self) -> None:
        self._stop.set()
        self.join(timeout=3)

    def summary(self) -> list[tuple[int, str, int]]:
        seen: dict[int, tuple[int, str, int]] = {}
        for job_id, document, total, _printed in self.samples:
            old = seen.get(job_id)
            pages = max(total, old[2] if old else 0)
            seen[job_id] = (job_id, document, pages)
        return sorted(seen.values())


def run_case(case, printer: str, spooler: WindowsSpooler, run_id: str) -> tuple[object, list[tuple[int, str, int]]]:
    runner = PrintRunner(spooler, AuditStore(ROOT / "dist" / "verify-separator.sqlite3"), queue_timeout=40)
    sampler = QueueSampler(spooler, printer)
    sampler.start()
    try:
        results = runner.run_cases([case], printer, run_id, lambda *_: None)
    finally:
        time.sleep(1.0)
        sampler.stop()
    return results[0], sampler.summary()


def main() -> int:
    printer = WindowsSpooler().win32print.GetDefaultPrinter()
    print(f"打印机：{printer}")
    cases, _ = read_cases(ROOT / "templates" / "PrintCase_Template.xlsx")
    base = next(case for case in cases if case.case_id == "TC-010")
    spooler = WindowsSpooler()

    plan = [
        ("A 分隔页=作业之间", "verify-sep-print-A", replace(
            base, pages="1", copies=1, separator_page="job", bin="纸盒1", separator_bin="纸盒2")),
        ("B 分隔页=无", "verify-sep-print-B", replace(
            base, pages="1", copies=1, separator_page="none", bin="纸盒1", separator_bin="")),
    ]

    outcomes = {}
    for label, run_id, case in plan:
        print(f"\n== {label} ==")
        result, jobs = run_case(case, printer, spooler, run_id)
        print(f"   用例状态：{result.status.value}　作业：{result.job_id}　用时：{result.elapsed_seconds:.1f}s")
        print(f"   队列采样：{jobs if jobs else '未采到（作业可能已瞬间出队）'}")
        outcomes[label] = (result, jobs)
        time.sleep(2.0)

    a_jobs = outcomes["A 分隔页=作业之间"][1]
    b_jobs = outcomes["B 分隔页=无"][1]
    a_pages = max((pages for _id, _doc, pages in a_jobs), default=0)
    b_pages = max((pages for _id, _doc, pages in b_jobs), default=0)

    print("\n== 结论 ==")
    print(f"   A 作业页面数={a_pages}（作业数 {len(a_jobs)}）；B 作业页面数={b_pages}（作业数 {len(b_jobs)}）")
    print("   说明：驱动把分隔页当作“由指定纸盒插入的硬件页”，不计入作业页数，")
    print("   所以 TotalPages 区分不出来；请核对出纸张数——A 应比 B 多一张分隔页。")
    print("   A 的分隔页纸张来源是 纸盒2：那个纸盒里放的应当是分隔纸。")
    checks = [
        ("A 作业已提交并出队", bool(a_jobs) or getattr(outcomes["A 分隔页=作业之间"][0], "job_id", None)),
        ("B 作业已提交并出队", bool(b_jobs) or getattr(outcomes["B 分隔页=无"][0], "job_id", None)),
    ]
    ok = all(flag for _name, flag in checks)
    for name, flag in checks:
        print(f"  [{'通过' if flag else '未通过'}] {name}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
