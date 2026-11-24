from pathlib import Path

from printcase.audit import AuditStore
from printcase.models import CaseStatus, PrintCase
from printcase.runner import PrintRunner
from printcase.spooler import QueueJob


class FakeWin32:
    JOB_STATUS_PRINTED = 1


class FakeSpooler:
    win32print = FakeWin32()
    def __init__(self):
        self.jobs = []
    def list_printers(self):
        return ["Test Printer"]
    def get_jobs(self, name):
        return list(self.jobs)


def fake_driver_applier(written=None):
    """替身：跳过真实驱动属性页操作，记录调用；返回空恢复函数。"""
    def apply(printer, case, log=None, caps=None):
        if written is not None:
            written.append((printer, case))
        return lambda: None
    return apply


def fake_submitter(spooler=None, record=None, backend="gdi-pdf"):
    """替身：不真正打印，仅记录并可选择向队列投放作业。"""
    def submit(printer, case, source_path, log, temp_dir):
        if record is not None:
            record.append((case.case_id, case.copies, Path(source_path)))
        if spooler is not None and spooler.jobs == []:
            spooler.jobs = [QueueJob(101, Path(source_path).name, 1, "驱动报告已打印", 2, 2)]
        return backend, Path(source_path).stem
    return submit


def make_runner(spooler, tmp_path, submitter, **kwargs):
    return PrintRunner(
        spooler, AuditStore(tmp_path / "history.sqlite3"),
        queue_timeout=1, submitter=submitter,
        driver_applier=fake_driver_applier(), **kwargs,
    )


def test_runner_correlates_spooler_job_and_reports_driver_printed(tmp_path):
    spooler = FakeSpooler()
    submit = fake_submitter(spooler=spooler)
    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"pdf")
    case = PrintCase(case_id="TC-PRINT", enabled=True, file_path=str(pdf), resolved_path=pdf, printer="Test Printer")
    events = []
    runner = make_runner(spooler, tmp_path, submit)
    results = runner.run_cases([case], "Test Printer", "RUN001", lambda k, v: events.append((k, v)))
    assert len(results) == 1
    assert results[0].status == CaseStatus.DRIVER_PRINTED
    assert results[0].job_id == 101
    assert any(kind == "log" for kind, _ in events)


class ClearingSpooler(FakeSpooler):
    def __init__(self):
        super().__init__()
        self.calls = 0
    def get_jobs(self, name):
        self.calls += 1
        if self.calls == 1:
            return []
        if self.calls == 2:
            self.jobs = [QueueJob(202, "sample.pdf", 0, "无状态标记", 0, 1)]
            return list(self.jobs)
        self.jobs = []
        return []


def test_runner_marks_queue_disappearance_as_unverified_physical_output(tmp_path):
    spooler = ClearingSpooler()
    runner = make_runner(spooler, tmp_path, fake_submitter())
    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"pdf")
    case = PrintCase(case_id="TC-CLEAR", enabled=True, file_path=str(pdf), resolved_path=pdf, printer="Test Printer")
    result = runner.run_cases([case], "Test Printer", "RUN002", lambda *_: None)[0]
    assert result.status == CaseStatus.QUEUE_CLEARED
    assert "不能证明纸张实际输出" in result.message


def test_runner_routes_pdf_to_gdi_and_logs_steps(tmp_path):
    """PDF 走内置 GDI 通道；前端日志应包含“打开文档/选择打印机/下发”步骤。"""
    spooler = FakeSpooler()
    record = []
    submit = fake_submitter(spooler=spooler, record=record)
    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"pdf")
    case = PrintCase(case_id="TC-GDI", enabled=True, file_path=str(pdf), resolved_path=pdf, printer="Test Printer")
    events = []
    runner = make_runner(spooler, tmp_path, submit)
    result = runner.run_cases([case], "Test Printer", "RUN003", lambda k, v: events.append((k, str(v))))[0]
    assert result.status == CaseStatus.DRIVER_PRINTED
    assert len(record) == 1
    assert record[0][2].suffix == ".pdf"
    assert result.settings["backend"] == "gdi-pdf"
    joined = "\n".join(text for kind, text in events if kind == "log")
    assert "步骤 1/4" in joined and "打印首选项" in joined
    assert "步骤 2/4" in joined and "打开打印文档 sample.pdf" in joined
    assert "步骤 3/4" in joined and "选择打印机 Test Printer" in joined


def test_runner_sets_driver_ui_before_submit_and_restores_after(tmp_path):
    """驱动属性页设置（步骤1）必须先于下发，且结束后调用还原。"""
    spooler = FakeSpooler()
    order = []
    written = []

    def driver_applier(printer, case, log=None, caps=None):
        order.append("apply")
        written.append((printer, case.paper))
        return lambda: order.append("restore")

    def submit(printer, case, source_path, log, temp_dir):
        order.append("submit")
        spooler.jobs = [QueueJob(305, Path(source_path).name, 1, "驱动报告已打印", 2, 2)]
        return "gdi-pdf", Path(source_path).stem

    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"pdf")
    case = PrintCase(case_id="TC-ORDER", enabled=True, file_path=str(pdf), resolved_path=pdf, printer="Test Printer", paper="A4")
    runner = PrintRunner(
        spooler, AuditStore(tmp_path / "audit.sqlite3"),
        queue_timeout=1, submitter=submit, driver_applier=driver_applier,
    )
    result = runner.run_cases([case], "Test Printer", "RUN004", lambda *_: None)[0]
    assert result.status == CaseStatus.DRIVER_PRINTED
    assert order == ["apply", "submit", "restore"]
    assert written == [("Test Printer", "A4")]


def test_runner_cleans_stale_driver_ui_after_stop(tmp_path, monkeypatch):
    """停止后必须清理驱动属性页残留，否则下一次下发会“属性页窗口未出现”。"""
    import printcase.runner as runner_module

    spooler = FakeSpooler()
    cleaned: list[str] = []

    def fake_apply(printer, case, caps=None, log=None):
        return lambda: None

    monkeypatch.setattr(runner_module, "apply_case", fake_apply)
    monkeypatch.setattr(
        runner_module, "close_stale_preferences",
        lambda printer, log=None: cleaned.append(printer) or 0,
    )

    holder = {}

    def submit(printer, case, source_path, log, temp_dir):
        holder["runner"].stop()          # 用户在下发过程中点了“停止”
        return "gdi-pdf", Path(source_path).stem

    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"pdf")
    case = PrintCase(case_id="TC-STOP", enabled=True, file_path=str(pdf), resolved_path=pdf, printer="Test Printer")
    runner = PrintRunner(
        spooler, AuditStore(tmp_path / "audit.sqlite3"),
        queue_timeout=1, submitter=submit, driver_applier=fake_apply,
    )
    holder["runner"] = runner
    results = runner.run_cases([case], "Test Printer", "RUN005", lambda *_: None)

    assert results[0].status == CaseStatus.STOPPED
    # 本轮第一次设置属性前会清一次残留，停止后再清一次
    assert cleaned and set(cleaned) == {"Test Printer"}


class BusySpooler(FakeSpooler):
    """支持高频监视的替身：上一件作业一直留在队列里。"""

    supports_queue_watch = True

    def __init__(self) -> None:
        super().__init__()
        self.busy = False

    def get_jobs(self, name):
        if self.busy:
            return [QueueJob(900, "previous.pdf", 0, "正在打印", 0, 2)]
        return []


def test_runner_waits_for_previous_job_before_opening_driver_page(tmp_path, monkeypatch):
    """上一件作业还在队列里时，先等它离开队列再打开驱动属性页。"""
    import printcase.runner as runner_module

    monkeypatch.setattr(runner_module, "QUEUE_SETTLE_TIMEOUT", 1.0)
    spooler = BusySpooler()
    order: list[str] = []

    def driver_applier(printer, case, log=None, caps=None):
        order.append("apply")
        return lambda: None

    def submit(printer, case, source_path, log, temp_dir):
        order.append("submit")
        spooler.busy = True          # 作业进了队列且迟迟不离开
        return "gdi-pdf", Path(source_path).stem

    pdf = tmp_path / "sample.pdf"
    pdf.write_bytes(b"pdf")
    cases = [
        PrintCase(case_id=f"TC-SETTLE-{i}", enabled=True, file_path=str(pdf), resolved_path=pdf, printer="Test Printer")
        for i in range(2)
    ]
    logs: list[str] = []
    runner = PrintRunner(
        spooler, AuditStore(tmp_path / "audit.sqlite3"),
        queue_timeout=1, submitter=submit, driver_applier=driver_applier,
    )
    runner.run_cases(cases, "Test Printer", "RUN006", lambda kind, value: logs.append(str(value)))

    # 第二件作业开始前出现“队列仍有作业”的等待，之后才继续设置属性
    assert any("队列仍有" in line for line in logs)
    assert order == ["apply", "submit", "apply", "submit"]
