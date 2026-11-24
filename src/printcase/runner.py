from __future__ import annotations

import shutil
import tempfile
import threading
import time
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import Callable

from .audit import AuditStore
from .direct_print import DirectPrintError, print_with_driver
from .driver_ui import DriverUiError, apply_case, close_stale_preferences
from .models import CaseResult, CaseStatus, PrintCase
from .office import DocumentConversionError
from .queue_watch import ERROR_JOB_MASKS, QueueWatcher
from .ui_print import UI_PRINTABLE_EXTENSIONS, UiPrintError, print_with_ui
from .separator import create_case_separator_pdf
from .validator import normalize_paper, validate_case


# 关联队列作业的等待上限：作业在下发后通常 1-3 秒内进入队列，没必要按用例
# 超时（默认 90s）整段空等——那会让“未关联”的用例白等一整轮。
CORRELATE_TIMEOUT_CAP = 20.0

# 打开驱动属性页前，等上一轮作业离开队列的最长时间。停止/中断后立刻重开属性页，
# 驱动常常还在收尾：printui 会静默退出，一个窗口都不建。
QUEUE_SETTLE_TIMEOUT = 6.0


class PrintRunner:
    """按“驱动界面设属性 → 打开文档 → 下发打印 → 还原驱动默认”的流程执行用例。

    只有一条通道：属性一律在驱动自己的“打印首选项”里设置并回读确认
    （见 driver_ui），不再直接改写驱动 DEVMODE；文档提交时：

    - Word/WPS 能打开的文档走 UI 自动化（Ctrl+P → 选驱动 → 打印）；
    - PDF/图片等没有文档界面的格式由内置渲染器送入打印队列。

    两种提交方式都继承刚刚设置并校验过的驱动默认设置。每个步骤都通过
    on_event("log", ...) 推送到前端日志页，用户可以完整看到操作过程。
    """

    def __init__(
        self,
        spooler,
        audit: AuditStore,
        queue_timeout: int = 90,
        submitter: Callable[..., tuple[str, str]] | None = None,
        document_converter: Callable[..., Path] | None = None,
        driver_applier: Callable[..., Callable[[], None]] | None = None,
        ui_submitter: Callable[..., tuple[str, str]] | None = None,
        queue_watcher_factory: Callable[..., QueueWatcher] | None = None,
    ) -> None:
        self.spooler = spooler
        self.audit = audit
        self.queue_timeout = queue_timeout
        self.document_converter = document_converter or None
        self.submit = submitter or self._default_submit
        self.driver_applier = driver_applier or apply_case
        self.ui_submitter = ui_submitter or self._default_ui_submit
        self.queue_watcher_factory = queue_watcher_factory or (lambda reader: QueueWatcher(reader))
        self.stop_event = threading.Event()
        # 本轮是否已经清理过“上一轮中断留下的驱动属性页”
        self._driver_ui_reset = False
        # 上一件事是否真的向队列提交过作业（决定下一件事前要不要等队列收尾）
        self._submitted_since_idle = False

    def _make_watcher(self, printer: str, log: Callable[[str], None] | None = None) -> QueueWatcher | None:
        """后台队列监视器；不支持高频监视的替身（测试用）返回 None，走同步轮询。"""
        if not getattr(self.spooler, "supports_queue_watch", False):
            return None
        try:
            watcher = self.queue_watcher_factory(lambda: self.spooler.get_jobs(printer))
            printed_mask = getattr(getattr(self.spooler, "win32print", None), "JOB_STATUS_PRINTED", 0)
            if printed_mask:
                watcher._printed_mask = printed_mask  # noqa: WPS437 - 注入驱动“已打印”状态位
            watcher.start()
            if not watcher.wait_ready(1.5) and log:
                log("队列监视器 1.5 秒内没有完成首次采样，队列关联可能不准")
            return watcher
        except Exception as exc:  # noqa: BLE001 - 监视器不可用时退回同步轮询
            if log:
                log(f"队列监视器启动失败，改用同步轮询：{type(exc).__name__}: {exc}")
            return None

    def _default_submit(self, printer: str, job_case: PrintCase, source_path: Path, log: Callable[[str], None], temp_dir: Path) -> tuple[str, str]:
        return print_with_driver(
            printer, job_case, source_path, log,
            stop_event=self.stop_event,
            document_converter=self.document_converter,
            temp_dir=temp_dir,
        )

    def _default_ui_submit(self, printer: str, job_case: PrintCase, source_path: Path, log: Callable[[str], None], temp_dir: Path) -> tuple[str, str]:
        """文档界面自动化下发；PDF/图片等没有文档界面的格式交由内置渲染器送入队列。"""
        if Path(source_path).suffix.lower() in UI_PRINTABLE_EXTENSIONS:
            # 属性已由驱动属性页设置并回读确认，这里只做“打开文档→选驱动→打印”
            return print_with_ui(
                printer, job_case, source_path, log,
                timeout=job_case.timeout_seconds,
                set_driver_attributes=False,
            )
        log(f"{Path(source_path).suffix} 没有可自动化的文档界面，本次由内置渲染器送入打印队列")
        return self.submit(printer, job_case, source_path, log, temp_dir)

    def _submit_case(self, printer: str, job_case: PrintCase, source_path: Path, log: Callable[[str], None], temp_dir: Path):
        return self.ui_submitter(printer, job_case, source_path, log, temp_dir)

    def stop(self) -> None:
        self.stop_event.set()

    # ---------- 驱动属性页的残留清理与队列收尾 ----------
    def _uses_real_driver_ui(self) -> bool:
        """只有真实驱动属性页通道才需要清理桌面残留（替身注入时跳过）。"""
        return self.driver_applier is apply_case

    def _reset_driver_ui(self, printer: str, emit: Callable[[str], None]) -> None:
        """每轮第一次设置属性前，先清掉上一轮停止/失败留下的属性页窗口与进程。"""
        if not self._uses_real_driver_ui():
            return
        try:
            cleared = close_stale_preferences(printer, log=emit)
        except Exception as exc:  # noqa: BLE001
            emit(f"清理上一轮残留的驱动属性页失败：{type(exc).__name__}: {exc}")
            return
        if cleared:
            emit(f"已清理上一轮残留的驱动属性页窗口/进程（{cleared} 项）")

    def _wait_queue_settled(self, printer: str, emit: Callable[[str], None]) -> None:
        """等上一个作业离开队列再打开驱动属性页。

        作业还在队列里时驱动仍在收尾，此时 printui 常常直接静默退出，
        表现为“属性页窗口未出现”。最多等 QUEUE_SETTLE_TIMEOUT 秒，等不到也照常继续。
        """
        if not self._submitted_since_idle or not getattr(self.spooler, "supports_queue_watch", False):
            return
        self._submitted_since_idle = False
        started = time.monotonic()
        announced = False
        while time.monotonic() - started < QUEUE_SETTLE_TIMEOUT:
            try:
                jobs = self.spooler.get_jobs(printer)
            except Exception:  # noqa: BLE001
                return
            if not jobs:
                elapsed = time.monotonic() - started
                if announced:
                    emit(f"打印队列已清空（等待 {elapsed:.1f} 秒），继续设置驱动属性")
                return
            if not announced:
                names = "、".join(f"#{job.job_id}" for job in jobs[:3])
                emit(f"打印机队列仍有 {len(jobs)} 个作业（{names}）在处理，等其离开队列后再打开驱动属性页")
                announced = True
            if self.stop_event.wait(0.4):
                return
        emit(f"打印队列在 {QUEUE_SETTLE_TIMEOUT:.0f} 秒内未清空，仍继续设置驱动属性（属性页可能被驱动暂时拒绝打开）")

    def run_cases(
        self,
        cases: list[PrintCase],
        fallback_printer: str,
        run_id: str,
        on_event: Callable[[str, object], None],
    ) -> list[CaseResult]:
        self.stop_event.clear()
        self._driver_ui_reset = False
        self._submitted_since_idle = False
        results: list[CaseResult] = []
        printers = set(self.spooler.list_printers())
        for case in cases:
            if self.stop_event.is_set():
                result = self._base_result(case, CaseStatus.STOPPED, "用户请求停止；未下发此用例。", case.printer or fallback_printer)
                results.append(result)
                self.audit.save(run_id, result)
                on_event("result", result)
                continue
            if not case.enabled:
                result = self._base_result(case, CaseStatus.STOPPED, "用例未启用，已跳过。", case.printer or fallback_printer)
                results.append(result)
                self.audit.save(run_id, result)
                on_event("result", result)
                continue
            errors = validate_case(case, printers, fallback_printer)
            if errors:
                result = self._base_result(case, CaseStatus.PREFLIGHT_FAILED, "；".join(errors), case.printer or fallback_printer)
                results.append(result)
                self.audit.save(run_id, result)
                on_event("result", result)
                continue

            printer = case.printer.strip() or fallback_printer.strip()
            result = self._execute_one(case, printer, on_event)
            results.append(result)
            self.audit.save(run_id, result)
            on_event("result", result)
        return results

    def _base_result(self, case: PrintCase, status: CaseStatus, message: str, printer: str) -> CaseResult:
        return CaseResult(case.case_id, status, message, printer=printer, source_file=case.file_path)

    def _execute_one(self, case: PrintCase, printer: str, on_event: Callable[[str, object], None]) -> CaseResult:
        started = datetime.now()
        started_tick = time.monotonic()
        emit = lambda text: on_event("log", f"[{case.case_id}] {text}")  # noqa: WPS430
        result = CaseResult(
            case_id=case.case_id,
            status=CaseStatus.SUBMITTING,
            printer=printer,
            started_at=started,
            settings={
                "pages": case.pages or "all", "copies": case.copies,
                "paper": case.paper, "orientation": case.orientation,
                "color": case.color, "duplex": case.duplex,
                "bin": case.bin,
                "collate": case.collate or "driver_default",
                "print_quality": case.print_quality,
                "media_type": case.media_type or "driver_default",
                "mixed_sizes": case.mixed_sizes,
                "skip_blank": case.skip_blank,
                "borderless": case.borderless,
                "separator_page": case.separator_page,
                "separator_bin": case.separator_bin,
                "binding_location": case.binding_location,
                "binding_method": case.binding_method,
            },
            source_file=str(case.resolved_path or case.file_path),
        )
        emit(f"用例开始，目标打印机：{printer}")
        if case.borderless == "yes":
            emit("无边距将在驱动属性页里勾选；驱动界面没有该选项时用例判定失败且不下发。")
        on_event("result", result)

        temp_dir: Path | None = None
        restore_driver = None
        try:
            try:
                _, printer_status = self.spooler.get_printer_status(printer)
                emit(f"打印机状态：{printer_status}")
            except Exception as status_error:
                emit(f"无法读取打印机状态：{status_error}")

            # 步骤 1/4：在驱动属性页里设置用例属性（设置后只读回读默认 DEVMODE 校验）
            if not self._driver_ui_reset:
                self._driver_ui_reset = True
                self._reset_driver_ui(printer, emit)
            self._wait_queue_settled(printer, emit)
            emit("步骤 1/4：打开驱动“打印首选项”，按用例设置纸张尺寸/纸张类型/来源/双面/颜色/份数")
            caps = None
            try:
                caps = self.spooler.get_capabilities(printer)
            except Exception as caps_error:
                emit(f"无法读取驱动能力，属性将按名称在驱动界面里匹配：{caps_error}")
            try:
                restore_driver = self.driver_applier(printer, case, caps=caps, log=emit)
            except DriverUiError as du_error:
                result.status = CaseStatus.FAILED
                result.message = f"为保证属性准确，本次不下发：{du_error}"
                emit(result.message)
                return result
            except Exception as du_error:  # noqa: WPS440
                result.status = CaseStatus.FAILED
                result.message = f"驱动属性页操作异常，为保证属性准确本次不下发：{type(du_error).__name__}: {du_error}"
                emit(result.message)
                return result

            source_path = Path(case.resolved_path or case.file_path)
            temp_dir = Path(tempfile.mkdtemp(prefix="printcase_job_"))

            each_copy_mode = case.separator_page == "each_copy" and case.copies > 1
            if not each_copy_mode and case.separator_page in {"before", "both"}:
                self._print_separator_page(case, printer, on_event, position="before", temp_dir=temp_dir, caps=caps, emit=emit)

            # 步骤 2-3/4：打开文档 → 选择驱动 → 按用例属性下发（在 submit 内逐步汇报）
            job_case = replace(case, copies=1) if each_copy_mode else case
            if each_copy_mode:
                emit(f"分隔页=份数之间：将按 {case.copies} 份逐份下发，份与份之间插入分隔页。")
            emit(f"步骤 2/4：打开打印文档 {source_path.name}")
            emit(f"步骤 3/4：选择打印机 {printer} 并按用例属性下发")
            baseline = {job.job_id for job in self.spooler.get_jobs(printer)}
            # 队列监视器必须在提交之前启动：作业在队列里可能只存活几百毫秒
            watcher = self._make_watcher(printer, emit)
            known_ids = set(baseline)
            if watcher is not None:
                known_ids |= watcher.snapshot_ids()
            submitted = False
            try:
                submit_result = self._submit_case(printer, job_case, source_path, emit, temp_dir)
                self._submitted_since_idle = True
                submitted = True
                backend, doc_stem = submit_result[0], submit_result[1]
                job_hint = submit_result[2] if len(submit_result) > 2 else None
                if job_hint:
                    result.job_id = job_hint
            except UiPrintError as submit_error:
                result.status = CaseStatus.FAILED
                result.message = f"UI 自动化打印失败：{submit_error}"
                emit(result.message)
                return result
            except DirectPrintError as submit_error:
                result.status = CaseStatus.FAILED
                result.message = f"打印下发失败：{submit_error}"
                emit(result.message)
                return result
            except FileNotFoundError as submit_error:
                result.status = CaseStatus.FAILED
                result.message = f"找不到打印文档或组件：{submit_error}"
                emit(result.message)
                return result
            except DocumentConversionError as submit_error:
                result.status = CaseStatus.FAILED
                result.message = f"Office 文档转 PDF 失败：{submit_error}"
                emit(result.message)
                return result
            finally:
                # 下发阶段就失败时不会走到下面的 await，必须先停掉监视线程，
                # 否则它每 0.08 秒 OpenPrinter 一次，跨用例累积成后台泄漏。
                if not submitted and watcher is not None:
                    watcher.stop()
            result.settings["backend"] = backend
            emit(f"步骤 3/4 完成：作业已通过 {backend} 通道提交，开始监视打印队列")

            try:
                self._await_queue_result(
                    printer, baseline, doc_stem, case, result, job_hint=job_hint,
                    watcher=watcher, known_ids=known_ids,
                    notify=lambda: on_event("result", result),
                )
            finally:
                if watcher is not None:
                    watcher.stop()

            if each_copy_mode and result.status != CaseStatus.STOPPED:
                self._print_remaining_copies(case, printer, on_event, temp_dir, source_path, caps, emit)

            if case.separator_page in {"after", "both"} and result.status != CaseStatus.STOPPED:
                self._print_separator_page(case, printer, on_event, position="after", temp_dir=temp_dir, caps=caps, emit=emit)

            emit("步骤 4/4：还原驱动默认设置（在 finally 中执行）")

        except Exception as exc:
            result.status = CaseStatus.FAILED
            result.message = f"执行异常：{type(exc).__name__}: {exc}"
        finally:
            if restore_driver is not None:
                try:
                    restore_driver()
                except Exception as restore_error:
                    emit(f"驱动默认设置还原失败，请在驱动首选项中手工核对：{restore_error}")
            if self.stop_event.is_set() or result.status in {CaseStatus.FAILED, CaseStatus.STOPPED}:
                # 停止/失败后驱动属性页可能还挂在桌面上（甚至已经卡死），
                # 不清掉的话下一次下发会“属性页窗口未出现”。
                self._reset_driver_ui(printer, emit)
            if temp_dir is not None:
                shutil.rmtree(temp_dir, ignore_errors=True)
            result.ended_at = datetime.now()
            result.elapsed_seconds = time.monotonic() - started_tick
            on_event("log", f"[{case.case_id}] 结果：{result.status.value} — {result.message}")
        return result

    def _await_queue_result(
        self, printer: str, baseline: set, doc_stem: str, case: PrintCase, result: CaseResult,
        job_hint: int | None = None, watcher=None, known_ids: set | None = None, notify=None,
    ) -> None:
        """在 Windows 打印队列里关联并跟踪本次作业，直到到达可观察的终态。

        有后台监视器时以它的留档为准（不会漏掉几百毫秒的短命作业）；
        没有时退回同步轮询（测试替身/监视器不可用），并给关联等待设上限。
        """
        if watcher is not None:
            self._await_with_watcher(doc_stem, case, result, job_hint, watcher, known_ids or set(), notify)
            return

        # ---------- 同步轮询（无后台监视器） ----------
        seen_job = None
        if job_hint:
            probe_deadline = time.monotonic() + 3
            while time.monotonic() < probe_deadline:
                try:
                    jobs = self.spooler.get_jobs(printer)
                except Exception:  # noqa: WPS420
                    jobs = []
                hit = next((job for job in jobs if job.job_id == job_hint), None)
                if hit is not None:
                    seen_job = hit
                    break
                time.sleep(0.2)
            if seen_job is None:
                result.status = CaseStatus.SUBMITTED_UNVERIFIED
                result.message = (
                    f"Windows 打印队列已分配作业号 #{job_hint}，但作业在监视窗口内已完成或已清除；"
                    "这不能证明纸张实际输出。"
                )
                return

        # 阶段一：按作业名关联（驱动未返回作业号时使用）
        correlate_window = max(3.0, min(float(case.timeout_seconds or 90), CORRELATE_TIMEOUT_CAP))
        correlate_deadline = time.monotonic() + correlate_window
        while time.monotonic() < correlate_deadline:
            if self.stop_event.is_set():
                result.status = CaseStatus.STOPPED
                result.message = "下发过程中收到停止请求；请检查打印队列，不能假定作业已撤回。"
                return
            try:
                current_jobs = self.spooler.get_jobs(printer)
            except Exception as spool_error:
                current_jobs = []
                result.message = f"读取打印队列失败：{spool_error}"

            new_jobs = [j for j in current_jobs if j.job_id not in baseline]
            matched = [
                j for j in new_jobs
                if (j.document or "").strip()
                and (
                    doc_stem.casefold() in j.document.casefold()
                    or j.document.casefold() in doc_stem.casefold()
                )
            ]
            if len(matched) == 1:
                seen_job = matched[0]
                break
            if len(matched) > 1:
                result.status = CaseStatus.SUBMITTED_UNVERIFIED
                result.message = "发现多个名称匹配的队列作业，无法安全判断归属；请人工核对，不建议自动重试。"
                return
            if len(new_jobs) == 1:
                seen_job = new_jobs[0]
                break
            time.sleep(0.12)

        if seen_job is None:
            result.status = CaseStatus.SUBMITTED_UNVERIFIED
            result.message = (
                f"作业已提交，但在 {correlate_window:.0f} 秒监视窗口内未捕获到队列作业记录"
                "（驱动可能直接送打印机，或作业瞬间完成）；这不能证明纸张实际输出。"
            )
            return

        result.job_id = seen_job.job_id
        result.status = CaseStatus.QUEUED

        # 阶段二：跟踪到终态
        terminal_deadline = time.monotonic() + self.queue_timeout
        while True:
            result.message = f"已关联 Windows 作业 #{seen_job.job_id}：{seen_job.status_text}"
            if any(flag in seen_job.status_text for flag in ERROR_JOB_MASKS):
                result.status = CaseStatus.FAILED
                result.message = f"打印队列异常：{seen_job.status_text}"
                return
            if seen_job.status & getattr(self.spooler.win32print, "JOB_STATUS_PRINTED", 0):
                result.status = CaseStatus.DRIVER_PRINTED
                result.message = "Windows 驱动报告作业已打印；不等同于人工确认纸张实际出纸。"
                return
            if time.monotonic() >= terminal_deadline:
                result.status = CaseStatus.TIMEOUT_QUEUED
                result.message = f"作业 #{seen_job.job_id} 仍在队列中，等待超过监控窗口；请检查打印机面板及队列。"
                return
            time.sleep(0.25)
            try:
                jobs = self.spooler.get_jobs(printer)
            except Exception as spool_error:
                result.message = f"读取打印队列失败：{spool_error}"
                continue
            current = next((j for j in jobs if j.job_id == seen_job.job_id), None)
            if current is None:
                result.status = CaseStatus.QUEUE_CLEARED
                result.message = f"Windows 队列中已找不到作业 #{seen_job.job_id}；这只能证明队列已清除，不能证明纸张实际输出。"
                return
            seen_job = current
            result.message = f"作业 #{current.job_id}：{current.status_text}（已打印页 {current.pages_printed}/{current.total_pages or '?'}）"

    # ---------- 有后台监视器时的关联与跟踪 ----------
    def _await_with_watcher(self, doc_stem: str, case: PrintCase, result: CaseResult,
                            job_hint: int | None, watcher, known_ids: set, notify) -> None:
        def push() -> None:
            if notify is not None:
                try:
                    notify()
                except Exception:  # noqa: WPS420
                    pass

        job = None
        if job_hint:
            deadline = time.monotonic() + 3
            while time.monotonic() < deadline:
                observation = watcher.observation(int(job_hint))
                if observation is not None:
                    job = observation.job
                    break
                if self.stop_event.wait(0.05):
                    break
            if job is None:
                result.status = CaseStatus.SUBMITTED_UNVERIFIED
                result.message = (
                    f"Windows 打印队列已分配作业号 #{job_hint}，但作业在监视窗口内已完成或已清除；"
                    "这不能证明纸张实际输出。"
                )
                push()
                return
        else:
            result.message = "作业已提交，正在关联打印队列…"
            push()
            window = max(3.0, min(float(case.timeout_seconds or 90), CORRELATE_TIMEOUT_CAP))
            outcome = watcher.wait_for_new(
                known_ids, window, doc_stem, should_stop=self.stop_event.is_set,
            )
            if outcome.note == "stopped":
                result.status = CaseStatus.STOPPED
                result.message = "下发过程中收到停止请求；请检查打印队列，不能假定作业已撤回。"
                push()
                return
            if outcome.job is None:
                result.status = CaseStatus.SUBMITTED_UNVERIFIED
                health = watcher.health_text()
                if outcome.ambiguous:
                    result.message = "打印队列中同时出现多个新作业，无法安全判断归属；请人工核对，不建议自动重试。"
                else:
                    result.message = (
                        f"作业已提交，但在 {window:.0f} 秒监视窗口内未捕获到队列作业记录"
                        f"（驱动可能直接送打印机，或作业瞬间完成）；队列{health}。这不能证明纸张实际输出。"
                    )
                push()
                return
            job = outcome.job

        result.job_id = job.job_id
        result.status = CaseStatus.QUEUED
        result.message = f"已关联 Windows 作业 #{job.job_id}：{job.status_text}"
        push()

        kind, last = watcher.wait_terminal(
            job.job_id, self.queue_timeout, should_stop=self.stop_event.is_set,
        )
        if kind == "printed":
            result.status = CaseStatus.DRIVER_PRINTED
            result.message = "Windows 驱动报告作业已打印；不等同于人工确认纸张实际出纸。"
        elif kind == "error":
            result.status = CaseStatus.FAILED
            result.message = f"打印队列异常：{(last.status_text if last else '未知错误')}"
        elif kind == "gone":
            result.status = CaseStatus.QUEUE_CLEARED
            result.message = f"Windows 队列中已找不到作业 #{job.job_id}；这只能证明队列已清除，不能证明纸张实际输出。"
        elif kind == "stopped":
            result.status = CaseStatus.STOPPED
            result.message = "下发过程中收到停止请求；请检查打印队列，不能假定作业已撤回。"
        else:
            result.status = CaseStatus.TIMEOUT_QUEUED
            result.message = f"作业 #{job.job_id} 仍在队列中，等待超过监控窗口；请检查打印机面板及队列。"
        push()

    def _print_separator_page(self, case: PrintCase, printer: str, on_event: Callable[[str, object], None], position: str, temp_dir: Path, caps=None, emit=None) -> None:
        """在主作业之前/之后插打一张分隔页；分隔页失败只记录，不影响主用例状态。"""
        try:
            sep_pdf = create_case_separator_pdf(case, printer, temp_dir)
            sep_case = PrintCase(
                case_id=f"{case.case_id}-{position}",
                enabled=True,
                file_path=str(sep_pdf),
                paper=case.paper if case.paper.lower() not in {"auto", "mixed"} else "A4",
                color="monochrome",
                duplex="simplex",
                scaling="fit",
                bin=case.separator_bin,
            )
            sep_case.resolved_path = sep_pdf
            label = {"before": "前置", "after": "后置"}.get(position, position)
            emit(f"正在打印{label}分隔页：{sep_pdf.name}")
            # 分隔页要求单面黑白：为分隔页单独设置驱动属性（结束后恢复到主用例属性）
            sep_restore = None
            try:
                sep_restore = self.driver_applier(printer, sep_case, caps=caps, log=emit)
            except DriverUiError as sep_ds_error:
                emit(f"分隔页驱动属性设置失败，按当前驱动默认设置打印分隔页：{sep_ds_error}")
            try:
                self._submit_case(printer, sep_case, sep_pdf, emit, temp_dir)
                emit(f"{label}分隔页已提交打印。")
            finally:
                if sep_restore is not None:
                    try:
                        sep_restore()
                    except Exception as restore_error:  # noqa: WPS420
                        emit(f"分隔页驱动设置恢复失败：{restore_error}")
        except Exception as sep_error:
            emit(f"分隔页打印失败（不影响主用例）：{sep_error}")

    def _print_remaining_copies(self, case: PrintCase, printer: str, on_event: Callable[[str, object], None], temp_dir: Path, source_path: Path, caps=None, emit=None) -> None:
        """分隔页=份数之间：逐份下发剩余份数，份与份之间插入分隔页。单份失败只记录，不中断。"""
        for index in range(2, case.copies + 1):
            if self.stop_event.is_set():
                emit("收到停止请求，剩余份数不再下发。")
                break
            try:
                self._print_separator_page(case, printer, on_event, position=f"第{index}份前", temp_dir=temp_dir, caps=caps, emit=emit)
                one = replace(case, copies=1)
                emit(f"正在下发第 {index}/{case.copies} 份")
                self._submit_case(printer, one, source_path, emit, temp_dir)
                emit(f"第 {index}/{case.copies} 份已提交打印。")
            except Exception as copy_error:
                emit(f"第 {index} 份下发失败（不影响已完成部分）：{copy_error}")
