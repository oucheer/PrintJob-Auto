"""队列监视与批量下发回归测试。

真实打印机上，作业在 Windows 队列里可能只存活几百毫秒；旧实现“下发后才开始
轮询”，必然错过作业，于是每条用例都报“队列未关联”并空等到超时（批量时尤其
明显：三条用例各白等 90 秒）。这里用可控替身固化修复后的行为。
"""
from __future__ import annotations

import threading
import time
from pathlib import Path

from printcase.audit import AuditStore
from printcase.models import CaseStatus, PrintCase
from printcase.queue_watch import QueueWatcher
from printcase.runner import PrintRunner
from printcase.spooler import QueueJob


class FakeWin32:
    JOB_STATUS_PRINTED = 1


class ShortLivedJobSpooler:
    """作业只在队列里存活很短时间的打印机替身（线程安全）。"""

    win32print = FakeWin32()
    supports_queue_watch = True

    def __init__(self, live_seconds: float = 0.3):
        self.live_seconds = live_seconds
        self._lock = threading.Lock()
        self._jobs: dict[int, tuple[QueueJob, float]] = {}
        self._next_id = 600
        self.assigned: list[QueueJob] = []

    def list_printers(self) -> list[str]:
        return ["Test Printer"]

    def get_jobs(self, name: str) -> list[QueueJob]:
        now = time.monotonic()
        with self._lock:
            self._jobs = {jid: item for jid, item in self._jobs.items() if item[1] > now}
            return [item[0] for item in self._jobs.values()]

    def submit_job(self, document: str) -> QueueJob:
        """投递一个很快会消失的作业，返回作业对象。"""
        with self._lock:
            self._next_id += 1
            job = QueueJob(self._next_id, document, 0, "正在打印", 0, 1)
            self._jobs[job.job_id] = (job, time.monotonic() + self.live_seconds)
            self.assigned.append(job)
            return job


def make_submitter(spooler: ShortLivedJobSpooler):
    def submit(printer, case, source_path, log, temp_dir):
        spooler.submit_job(Path(source_path).stem)
        return "word-com", Path(source_path).stem

    return submit


def make_runner(spooler, tmp_path, queue_timeout=3):
    def apply(printer, case, log=None, caps=None):
        return lambda: None

    return PrintRunner(
        spooler,
        AuditStore(tmp_path / "history.sqlite3"),
        queue_timeout=queue_timeout,
        submitter=make_submitter(spooler),
        driver_applier=apply,
    )


def test_queue_watcher_captures_short_lived_job():
    """监视器必须留档几百毫秒就消失的作业。"""
    spooler = ShortLivedJobSpooler(live_seconds=0.3)
    watcher = QueueWatcher(lambda: spooler.get_jobs("Test Printer"))
    watcher.start()
    try:
        assert watcher.wait_ready(2.0)
        known = watcher.snapshot_ids()
        job = spooler.submit_job("A4单页测试页2")
        outcome = watcher.wait_for_new(known, timeout=2.0, document="A4单页测试页2")
        assert outcome.job is not None, outcome
        assert outcome.job.job_id == job.job_id
        kind, last = watcher.wait_terminal(job.job_id, timeout=3.0)
        assert kind == "gone", (kind, last)
    finally:
        watcher.stop()


def test_batch_same_document_fast_jobs_do_not_stall(tmp_path):
    """批量下发：三条用例用同一份文档，作业瞬间完成，也必须逐条关联成功。

    修复前每条都会「队列未关联」并白等到超时（90 秒 × 3）。
    """
    spooler = ShortLivedJobSpooler(live_seconds=0.3)
    pdf = tmp_path / "A4单页测试页2.pdf"
    pdf.write_bytes(b"%PDF-1.4")
    cases = [
        PrintCase(
            case_id=f"TC-01{i}", enabled=True, file_path=str(pdf), resolved_path=pdf,
            printer="Test Printer", paper="A4", duplex="simplex", print_quality="600",
        )
        for i in range(1, 4)
    ]
    runner = make_runner(spooler, tmp_path)
    started = time.monotonic()
    results = runner.run_cases(cases, "Test Printer", "BATCH", lambda *_: None)
    elapsed = time.monotonic() - started

    assert [r.case_id for r in results] == ["TC-011", "TC-012", "TC-013"]
    # 每条用例都关联到了自己的作业，且不是「队列未关联」
    assert all(r.status == CaseStatus.QUEUE_CLEARED for r in results), [r.status.value for r in results]
    job_ids = [r.job_id for r in results]
    assert all(job_ids), job_ids
    assert len(set(job_ids)) == 3, f"作业号串号：{job_ids}"
    # 三条合计远小于旧的 90 秒/条空等
    assert elapsed < 15, f"批量下发耗时异常：{elapsed:.1f}s"


def test_batch_is_sequential_and_audited(tmp_path):
    """批量必须是串行下发（避免属性窗口互相污染），且每条都写入审计。"""
    spooler = ShortLivedJobSpooler(live_seconds=0.2)
    pdf = tmp_path / "dup.pdf"
    pdf.write_bytes(b"%PDF-1.4")
    cases = [
        PrintCase(case_id=f"TC-02{i}", enabled=True, file_path=str(pdf), resolved_path=pdf,
                  printer="Test Printer")
        for i in range(1, 4)
    ]
    overlapping: list[str] = []
    active = {"count": 0}
    lock = threading.Lock()

    def apply(printer, case, log=None, caps=None):
        with lock:
            active["count"] += 1
            if active["count"] > 1:
                overlapping.append(case.case_id)
        return lambda: _release(case.case_id)

    def _release(case_id: str) -> None:
        with lock:
            active["count"] -= 1

    audit = AuditStore(tmp_path / "audit.sqlite3")
    runner = PrintRunner(
        spooler, audit, queue_timeout=2,
        submitter=make_submitter(spooler), driver_applier=apply,
    )
    results = runner.run_cases(cases, "Test Printer", "BATCH-SEQ", lambda *_: None)
    assert len(results) == 3
    assert overlapping == [], f"出现并发下发：{overlapping}"
    rows = audit.recent(limit=10)
    assert len([r for r in rows if str(r["case_id"]).startswith("TC-02")]) == 3


def test_multiple_new_jobs_are_reported_as_ambiguous(tmp_path):
    """同时出现多个新作业且名称相同，不能瞎猜归属。"""
    spooler = ShortLivedJobSpooler(live_seconds=0.6)
    pdf = tmp_path / "same.pdf"
    pdf.write_bytes(b"%PDF-1.4")
    case = PrintCase(case_id="TC-AMB", enabled=True, file_path=str(pdf), resolved_path=pdf,
                     printer="Test Printer")

    def submit(printer, case, source_path, log, temp_dir):
        # 一次提交产生两个同名作业，模拟驱动拆分成多个作业
        spooler.submit_job(Path(source_path).stem)
        spooler.submit_job(Path(source_path).stem)
        return "word-com", Path(source_path).stem

    runner = PrintRunner(
        spooler, AuditStore(tmp_path / "a.sqlite3"), queue_timeout=2,
        submitter=submit, driver_applier=lambda p, c, log=None, caps=None: (lambda: None),
    )
    result = runner.run_cases([case], "Test Printer", "R-AMB", lambda *_: None)[0]
    assert result.status == CaseStatus.SUBMITTED_UNVERIFIED
    assert "多个新作业" in result.message or "无法安全判断" in result.message


def test_watcher_reports_sampling_failures_instead_of_silently_idle():
    """读队列失败必须留下痕迹：否则日志只能说“未捕获作业”，无法区分没提交和没看到。"""
    from printcase.queue_watch import QueueWatcher

    def boom():
        raise OSError("拒绝访问")

    watcher = QueueWatcher(boom, interval=0.01)
    watcher.start()
    watcher.wait_ready(0.3)
    time.sleep(0.1)
    watcher.stop()

    assert watcher.samples == 0
    assert watcher.failures > 0
    text = watcher.health_text()
    assert "失败" in text and "拒绝访问" in text


def test_watcher_health_text_when_healthy():
    from printcase.queue_watch import QueueWatcher
    from printcase.spooler import QueueJob

    watcher = QueueWatcher(lambda: [QueueJob(1, "a.pdf", 0, "无状态标记")], interval=0.01)
    watcher.start()
    watcher.wait_ready(1.0)
    watcher.stop()
    assert watcher.samples > 0 and watcher.failures == 0
    assert "失败" not in watcher.health_text()
